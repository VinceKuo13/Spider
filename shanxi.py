# -*- coding: utf-8 -*-
"""
山西省政府采购网爬虫：山西政府采购网全站搜索 / 采购公告 / 结果公告 / 采购意向公开

适用网站：
    http://www.ccgp-shanxi.gov.cn/site/search?k=

已确认接口：
    列表接口：http://www.ccgp-shanxi.gov.cn/portal/all
    详情接口：http://www.ccgp-shanxi.gov.cn/portal/detail

默认功能：
1. 关键词默认：营商环境，可在同目录 config.yaml 中修改。
2. 时间默认：2025-01-01 至今天。
3. 山西接口单次日期范围不宜超过 365 天，代码会自动按 365 天拆分。
4. 默认抓取三类：采购公告、结果公告、采购意向公开。
5. 接口本身不是严格标题搜索，因此代码会本地过滤“标题包含关键词”的记录。
6. 每条记录保存 JSONL，最后导出 Excel。

运行：
    pip install requests beautifulsoup4 pandas openpyxl pyyaml
    python spider_shanxi.py

输出：
    outputs_shanxi/shanxi_records.jsonl
    outputs_shanxi/shanxi_visited.txt
    outputs_shanxi/山西政府采购网_标讯明细.xlsx
"""

from __future__ import annotations

import argparse
import html
import json
import logging
import math
import os
import random
import re
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup

try:
    import yaml  # type: ignore
except Exception:  # pragma: no cover
    yaml = None

try:
    import pandas as pd  # type: ignore
except Exception:  # pragma: no cover
    pd = None


BASE_URL = "http://www.ccgp-shanxi.gov.cn"
LIST_API = f"{BASE_URL}/portal/all"
DETAIL_API = f"{BASE_URL}/portal/detail"

# 中国标准时间。山西站日期参数是毫秒时间戳，浏览器 DatePicker 传的是本地 00:00 的时间戳。
CN_TZ = timezone(timedelta(hours=8))
BASE_DIR = Path(__file__).resolve().parent

DEFAULT_CONFIG: Dict[str, Any] = {
    "province": "山西",
    "keywords": ["营商环境"],
    "start_date": "2025-01-01",
    "end_date": None,  # None 表示今天
    "output_dir": "outputs_shanxi",
    "page_size": 15,
    "max_days_per_query": 365,
    "title_filter": True,  # 接口会搜全文，本地再过滤标题包含关键词
    "fetch_detail": True,
    "request_timeout": 25,
    "max_retries": 4,
    "sleep_min": 0.8,
    "sleep_max": 2.0,
    "detail_sleep_min": 0.5,
    "detail_sleep_max": 1.5,
    "announcement_types": {
        "采购公告": "ZcyAnnouncement1",
        "结果公告": "ZcyAnnouncement2",
        "采购意向公开": "ZcyAnnouncement6",
    },
}

UNIFIED_COLUMNS = [
    "省份",
    "检索关键词",
    "公告类型",
    "栏目路径",
    "标题",
    "发布时间",
    "地区",
    "采购人/招标人",
    "采购代理机构",
    "项目名称",
    "项目编号",
    "采购方式",
    "预算金额",
    "最高限价",
    "中标单位/成交供应商",
    "中标金额/成交金额",
    "供应商地址",
    "投标/响应截止时间",
    "信息来源",
    "详情链接",
    "附件",
    "正文摘要",
    "articleId",
    "annId",
    "解析状态",
    "错误信息",
]


@dataclass
class DateRange:
    start: date
    end: date  # inclusive


# ------------------------- 基础工具 -------------------------

def setup_logging(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log_file = output_dir / "shanxi_spider.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, encoding="utf-8"),
        ],
    )


def load_config(config_path: Optional[str]) -> Dict[str, Any]:
    cfg = dict(DEFAULT_CONFIG)
    if not config_path:
        default_file = BASE_DIR / "config.yaml"
        config_path = str(default_file) if default_file.exists() else None

    if config_path:
        if yaml is None:
            raise RuntimeError("检测到配置文件，但未安装 pyyaml，请执行：pip install pyyaml")
        p = Path(config_path)
        if not p.exists():
            raise FileNotFoundError(f"配置文件不存在：{p}")
        with p.open("r", encoding="utf-8") as f:
            user_cfg = yaml.safe_load(f) or {}
        deep_update(cfg, user_cfg)
    if not cfg.get("end_date"):
        cfg["end_date"] = date.today().isoformat()
    return cfg


def deep_update(base: Dict[str, Any], extra: Dict[str, Any]) -> None:
    for k, v in extra.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            base[k].update(v)
        else:
            base[k] = v


def parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def to_cn_midnight_ms(d: date) -> int:
    dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=CN_TZ)
    return int(dt.timestamp() * 1000)


def end_exclusive_ms(d: date) -> int:
    """接口 end 参数按次日 00:00 传，保证包含 d 当天。"""
    return to_cn_midnight_ms(d + timedelta(days=1))


def split_date_ranges(start: date, end: date, max_days: int = 365) -> List[DateRange]:
    if start > end:
        raise ValueError(f"开始日期不能晚于结束日期：{start} > {end}")
    ranges: List[DateRange] = []
    cur = start
    while cur <= end:
        seg_end = min(cur + timedelta(days=max_days - 1), end)
        ranges.append(DateRange(cur, seg_end))
        cur = seg_end + timedelta(days=1)
    return ranges


def clean_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s)
    s = html.unescape(s)
    s = re.sub(r"<\s*/?\s*em\s*>", "", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    s = s.replace("\xa0", " ").replace("&nbsp;", " ")
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n\s*\n+", "\n", s)
    return s.strip()


def normalize_for_match(s: str) -> str:
    s = clean_text(s)
    return re.sub(r"\s+", "", s)


def ms_to_date_str(ms: Any) -> str:
    if ms in (None, ""):
        return ""
    try:
        v = int(ms)
        return datetime.fromtimestamp(v / 1000, tz=CN_TZ).strftime("%Y-%m-%d")
    except Exception:
        return str(ms)


def amount_from_text(s: str) -> str:
    """从“投标报价（小写）：4369800（元）”等文本中提取金额。"""
    if not s:
        return ""
    s = clean_text(s)
    m = re.search(r"([0-9]+(?:[,，][0-9]{3})*(?:\.\d+)?)\s*(?:元|万元|万)?", s)
    if not m:
        return s
    amount = m.group(1).replace(",", "").replace("，", "")
    unit_match = re.search(r"(万元|万|元)", s)
    unit = unit_match.group(1) if unit_match else ""
    return amount + (unit if unit else "")


def short_summary(text: str, limit: int = 400) -> str:
    text = clean_text(text)
    text = re.sub(r"\s+", " ", text)
    return text[:limit]


def unique_keep_order(items: Iterable[str]) -> List[str]:
    seen = set()
    out: List[str] = []
    for x in items:
        x = clean_text(x)
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


# ------------------------- HTTP -------------------------

def make_session() -> requests.Session:
    sess = requests.Session()
    sess.trust_env = False
    sess.headers.update(
        {
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Content-Type": "application/json;charset=UTF-8",
            "Origin": BASE_URL,
            "Referer": f"{BASE_URL}/site/search?k=",
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/147.0.0.0 Safari/537.36"
            ),
            "X-Requested-With": "XMLHttpRequest",
        }
    )
    # 与浏览器一致，给一个客户端 uuid；不是鉴权，只是站点日志标识。
    sess.cookies.set("_zcy_log_client_uuid", "67128f90-4871-11f1-99f6-4de7ddd4cc50", domain="www.ccgp-shanxi.gov.cn")
    return sess


def request_json_with_retry(
    sess: requests.Session,
    method: str,
    url: str,
    *,
    cfg: Dict[str, Any],
    **kwargs: Any,
) -> Dict[str, Any]:
    max_retries = int(cfg.get("max_retries", 4))
    timeout = int(cfg.get("request_timeout", 25))
    last_error: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = sess.request(method, url, timeout=timeout, **kwargs)
            resp.raise_for_status()
            if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
                resp.encoding = "utf-8"
            return resp.json()
        except Exception as e:
            last_error = e
            wait = min(2 ** attempt, 20)
            logging.warning("请求 JSON 失败，第 %s/%s 次，等待 %s 秒：%s", attempt, max_retries, wait, e)
            time.sleep(wait)
    raise RuntimeError(f"请求失败：{url}; last_error={last_error}")


def request_text_with_retry(
    sess: requests.Session,
    method: str,
    url: str,
    *,
    cfg: Dict[str, Any],
    **kwargs: Any,
) -> str:
    max_retries = int(cfg.get("max_retries", 4))
    timeout = int(cfg.get("request_timeout", 25))
    last_error: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = sess.request(method, url, timeout=timeout, **kwargs)
            resp.raise_for_status()
            if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
                resp.encoding = "utf-8"
            return resp.text
        except Exception as e:
            last_error = e
            wait = min(2 ** attempt, 20)
            logging.warning("请求文本失败，第 %s/%s 次，等待 %s 秒：%s", attempt, max_retries, wait, e)
            time.sleep(wait)
    raise RuntimeError(f"请求失败：{url}; last_error={last_error}")


# ------------------------- 接口解析 -------------------------

def build_list_payload(keyword: str, second_code: str, dr: DateRange, page_no: int, page_size: int) -> Dict[str, Any]:
    return {
        "keyword": keyword,
        "firstCode": "ZcyAnnouncement",
        "secondCode": second_code,
        "districtCode": [],
        "publishDateBegin": to_cn_midnight_ms(dr.start),
        "publishDateEnd": end_exclusive_ms(dr.end),
        "pageNo": page_no,
        "pageSize": page_size,
        "isTitleSearch": None,
        "order": "desc",
        "leaf": "0",
    }


def fetch_list_page(
    sess: requests.Session,
    cfg: Dict[str, Any],
    keyword: str,
    second_code: str,
    dr: DateRange,
    page_no: int,
) -> Tuple[int, List[Dict[str, Any]]]:
    payload = build_list_payload(keyword, second_code, dr, page_no, int(cfg.get("page_size", 15)))
    data = request_json_with_retry(sess, "POST", LIST_API, cfg=cfg, json=payload)
    if not data.get("success"):
        raise RuntimeError(f"列表接口返回失败：{data}")
    result_data = (((data.get("result") or {}).get("data")) or {})
    total = int(result_data.get("total") or 0)
    rows = result_data.get("data") or []
    if not isinstance(rows, list):
        rows = []
    return total, rows


def build_detail_page_url(article_id: str, parent_id: Any) -> str:
    query = urlencode(
        {
            "categoryCode": "ZcyAnnouncement",
            "parentId": str(parent_id or "138010"),
            "articleId": article_id,
        }
    )
    return f"{BASE_URL}/site/detail?{query}"


def extract_html_from_detail_json(obj: Any) -> str:
    """portal/detail 有时返回 JSON，正文 HTML 可能藏在 result/content 等字段里。递归寻找 HTML。"""
    if obj is None:
        return ""
    if isinstance(obj, str):
        if "<" in obj and ("content-head" in obj or "项目编号" in obj or "<p" in obj or "<table" in obj):
            return obj
        return ""
    if isinstance(obj, dict):
        preferred_keys = [
            "content",
            "html",
            "body",
            "detailContent",
            "articleContent",
            "noticeContent",
            "data",
            "result",
        ]
        for k in preferred_keys:
            if k in obj:
                html_text = extract_html_from_detail_json(obj.get(k))
                if html_text:
                    return html_text
        for v in obj.values():
            html_text = extract_html_from_detail_json(v)
            if html_text:
                return html_text
    if isinstance(obj, list):
        for v in obj:
            html_text = extract_html_from_detail_json(v)
            if html_text:
                return html_text
    return ""


def fetch_detail_html(sess: requests.Session, cfg: Dict[str, Any], article_id: str, parent_id: Any) -> str:
    params = {
        "articleId": article_id,
        "parentId": str(parent_id or "138010"),
        "timestamp": str(int(time.time())),
    }
    text = request_text_with_retry(sess, "GET", DETAIL_API, cfg=cfg, params=params)
    # portal/detail 可能返回 JSON，也可能直接返回 HTML 片段，做双保险。
    try:
        obj = json.loads(text)
        html_text = extract_html_from_detail_json(obj)
        if html_text:
            return html_text
        # 如果没有找到 HTML，也保留 JSON 字符串，便于调试。
        return text
    except Exception:
        return text


# ------------------------- 详情字段解析 -------------------------

def class_contains(fragment: str):
    def _predicate(c: Any) -> bool:
        if not c:
            return False
        if isinstance(c, list):
            return any(fragment in str(x) for x in c)
        return fragment in str(c)
    return _predicate


def select_text_by_class(soup: BeautifulSoup, fragment: str, all_values: bool = False) -> str:
    nodes = soup.find_all(class_=class_contains(fragment))
    values = unique_keep_order(n.get_text(" ", strip=True) for n in nodes)
    if not values:
        return ""
    return "；".join(values) if all_values else values[0]


def soup_plain_text(soup: BeautifulSoup) -> str:
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = soup.get_text(" ", strip=True)
    text = html.unescape(text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def regex_value(text: str, label_patterns: List[str], stop_patterns: Optional[List[str]] = None) -> str:
    if stop_patterns is None:
        stop_patterns = [
            "项目名称", "采购方式", "预算金额", "最高限价", "采购需求", "合同履约期限", "本项目",
            "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、",
            "名 称", "地 址", "联系方式", "联系电话", "电 话", "地点", "时间", "方式", "售价", "截止时间",
        ]
    label = "|".join(label_patterns)
    stop = "|".join(re.escape(x) for x in stop_patterns)
    pattern = rf"(?:{label})\s*[:：]\s*(.+?)(?=\s*(?:{stop})\s*[:：]?|$)"
    m = re.search(pattern, text)
    if not m:
        return ""
    return clean_text(m.group(1))


def section_name(text: str, section_marker: str) -> str:
    # 例：1.采购人信息 名 称：介休市营商环境局 地 址：...
    pat = rf"{re.escape(section_marker)}.{0,120}?名\s*称\s*[:：]\s*(.+?)(?=\s*(?:地\s*址|联系方式|联系电话|电\s*话|2\.|3\.|$))"
    m = re.search(pat, text)
    return clean_text(m.group(1)) if m else ""


def parse_attachments(soup: BeautifulSoup) -> str:
    items: List[str] = []
    for a in soup.find_all("a", href=True):
        name = clean_text(a.get_text(" ", strip=True))
        href = a.get("href", "")
        if not name and not href:
            continue
        if any(ext in name.lower() or ext in href.lower() for ext in [".doc", ".docx", ".pdf", ".xls", ".xlsx", ".jpg", ".png", ".zip", ".rar"]):
            items.append(f"{name}：{href}" if href else name)
    return "；".join(unique_keep_order(items))


def parse_detail_fields(html_text: str, announcement_type: str) -> Dict[str, str]:
    soup = BeautifulSoup(html_text or "", "html.parser")
    text = soup_plain_text(soup)

    title = ""
    title_node = soup.select_one(".content-head-title")
    if title_node:
        title = clean_text(title_node.get_text(" ", strip=True))

    source = ""
    source_node = soup.select_one(".content-head-info-source")
    if source_node:
        source = clean_text(source_node.get_text(" ", strip=True)).replace("来源：", "").strip()

    publish_date = ""
    release_node = soup.select_one(".content-head-info-releaseTime")
    if release_node:
        publish_date = clean_text(release_node.get_text(" ", strip=True)).replace("发布时间：", "").strip()

    # 政采云模板中常用 code class。
    project_code = select_text_by_class(soup, "code-00004") or regex_value(text, ["项目编号"])
    project_name = select_text_by_class(soup, "code-00003") or regex_value(text, ["项目名称"])
    purchaser = select_text_by_class(soup, "code-00014") or section_name(text, "1.采购人信息")
    agency = select_text_by_class(soup, "code-00009") or section_name(text, "2.采购代理机构信息")

    procurement_method = regex_value(text, ["采购方式"], ["预算金额", "最高限价", "采购需求", "合同履约期限", "二、", "三、"])
    budget_amount = select_text_by_class(soup, "code-AM01400034") or regex_value(text, [r"预算金额（元）", r"预算金额"])
    price_ceiling = select_text_by_class(soup, "code-AM014priceCeiling") or regex_value(text, [r"最高限价（元）", r"最高限价"])

    deadline = select_text_by_class(soup, "code-25011") or regex_value(text, ["截止时间", "提交响应文件截止时间", "投标截止时间"])

    # 结果公告常用字段。
    winning_supplier = select_text_by_class(soup, "code-winningSupplierName", all_values=True)
    supplier_address = select_text_by_class(soup, "code-winningSupplierAddr", all_values=True)
    winning_amount_raw = select_text_by_class(soup, "code-summaryPrice", all_values=True)
    winning_amount = "；".join(amount_from_text(x) for x in winning_amount_raw.split("；") if clean_text(x))

    if not winning_supplier and "中标" in announcement_type or "结果" in announcement_type:
        # 表格或纯文本兜底。
        m = re.search(r"供应商名称\s+供应商地址\s+中标（成交）金额.*?\s+([^\s]+(?:（有限公司）|有限公司|公司|中心|厂|院|所|合作社)?)", text)
        if m:
            winning_supplier = clean_text(m.group(1))

    attachments = parse_attachments(soup)

    return {
        "标题": title,
        "信息来源": source,
        "发布时间": publish_date,
        "项目编号": project_code,
        "项目名称": project_name,
        "采购人/招标人": purchaser,
        "采购代理机构": agency,
        "采购方式": procurement_method,
        "预算金额": budget_amount,
        "最高限价": price_ceiling,
        "投标/响应截止时间": deadline,
        "中标单位/成交供应商": winning_supplier,
        "中标金额/成交金额": winning_amount,
        "供应商地址": supplier_address,
        "附件": attachments,
        "正文摘要": short_summary(text, 500),
    }


# ------------------------- 记录构造与保存 -------------------------

def base_record_from_list(
    province: str,
    keyword: str,
    announcement_type: str,
    row: Dict[str, Any],
) -> Dict[str, Any]:
    article_id = str(row.get("articleId") or "")
    parent_id = row.get("parentId") or 138010
    return {
        "省份": province,
        "检索关键词": keyword,
        "公告类型": announcement_type,
        "栏目路径": clean_text(row.get("pathName")),
        "标题": clean_text(row.get("title")),
        "发布时间": ms_to_date_str(row.get("publishDate")),
        "地区": clean_text(row.get("districtName")),
        "采购人/招标人": clean_text(row.get("purchaseName")),
        "采购代理机构": "",
        "项目名称": clean_text(row.get("projectName")),
        "项目编号": clean_text(row.get("projectCode")),
        "采购方式": clean_text(row.get("procurementMethod")),
        "预算金额": clean_text(row.get("budgetPrice")),
        "最高限价": "",
        "中标单位/成交供应商": clean_text(row.get("supplierName")),
        "中标金额/成交金额": clean_text(row.get("totalContractAmount")),
        "供应商地址": "",
        "投标/响应截止时间": ms_to_date_str(row.get("bidOpeningTime")),
        "信息来源": clean_text(row.get("author")),
        "详情链接": build_detail_page_url(article_id, parent_id),
        "附件": "",
        "正文摘要": short_summary(clean_text(row.get("content")), 500),
        "articleId": article_id,
        "annId": str(row.get("annId") or ""),
        "解析状态": "列表成功",
        "错误信息": "",
        "_parentId": str(parent_id),
        "_raw": row,
    }


def merge_detail_fields(rec: Dict[str, Any], detail_fields: Dict[str, str]) -> None:
    # 详情页字段优先，但如果详情为空则保留列表字段。
    for k, v in detail_fields.items():
        v = clean_text(v)
        if v:
            rec[k] = v


def append_jsonl(path: Path, obj: Dict[str, Any]) -> None:
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")


def load_visited(path: Path) -> set:
    if not path.exists():
        return set()
    return set(x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip())


def add_visited(path: Path, key: str) -> None:
    with path.open("a", encoding="utf-8") as f:
        f.write(key + "\n")


def load_jsonl(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    records: List[Dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                records.append(json.loads(line))
            except Exception:
                continue
    return records


def export_excel(records: List[Dict[str, Any]], output_path: Path) -> None:
    if not records:
        logging.warning("没有可导出的记录。")
        return
    if pd is None:
        logging.warning("未安装 pandas，跳过 Excel 导出。可执行：pip install pandas openpyxl")
        return
    df = pd.DataFrame(records)
    for col in UNIFIED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[UNIFIED_COLUMNS]
    # 简单去重：同一 articleId 只保留最后一次。
    df = df.drop_duplicates(subset=["articleId"], keep="last")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="山西标讯明细")
        ws = writer.book["山西标讯明细"]
        width_map = {
            "A": 8, "B": 14, "C": 14, "D": 28, "E": 48, "F": 14, "G": 14,
            "H": 28, "I": 30, "J": 36, "K": 22, "L": 16, "M": 16, "N": 16,
            "O": 34, "P": 18, "Q": 34, "R": 20, "S": 24, "T": 58, "U": 46,
            "V": 60, "W": 26, "X": 18, "Y": 14, "Z": 30,
        }
        for col, width in width_map.items():
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    logging.info("Excel 已导出：%s", output_path)


# ------------------------- 主爬取逻辑 -------------------------

def crawl_shanxi(config: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    cfg = dict(DEFAULT_CONFIG)
    if config:
        deep_update(cfg, config)
    if not cfg.get("end_date"):
        cfg["end_date"] = date.today().isoformat()

    output_dir = Path(cfg["output_dir"])
    setup_logging(output_dir)

    jsonl_path = output_dir / "shanxi_records.jsonl"
    visited_path = output_dir / "shanxi_visited.txt"
    excel_path = output_dir / "山西政府采购网_标讯明细.xlsx"

    province = str(cfg.get("province", "山西"))
    keywords = list(cfg.get("keywords") or [])
    ann_types = dict(cfg.get("announcement_types") or {})
    start = parse_date(str(cfg["start_date"]))
    end = parse_date(str(cfg["end_date"]))
    date_ranges = split_date_ranges(start, end, int(cfg.get("max_days_per_query", 365)))

    logging.info("开始爬取：山西政府采购网")
    logging.info("关键词：%s", "、".join(keywords))
    logging.info("时间范围：%s 至 %s；自动拆分为 %s 段", start, end, len(date_ranges))
    logging.info("公告类型：%s", "、".join(ann_types.keys()))
    logging.info("标题过滤：%s", bool(cfg.get("title_filter", True)))

    sess = make_session()
    visited = load_visited(visited_path)
    total_saved = 0

    try:
        for keyword in keywords:
            for ann_name, second_code in ann_types.items():
                for dr in date_ranges:
                    page_size = int(cfg.get("page_size", 15))
                    logging.info(
                        "任务：关键词=%s，公告类型=%s，日期=%s 至 %s，pageSize=%s",
                        keyword, ann_name, dr.start, dr.end, page_size,
                    )
                    try:
                        total, rows = fetch_list_page(sess, cfg, keyword, second_code, dr, 1)
                    except Exception as e:
                        logging.error("列表第一页失败：关键词=%s 类型=%s 日期=%s-%s，错误：%s", keyword, ann_name, dr.start, dr.end, e)
                        continue

                    total_pages = max(1, math.ceil(total / page_size)) if total else 0
                    logging.info("列表结果：total=%s，pages=%s", total, total_pages)

                    for page_no in range(1, total_pages + 1):
                        if page_no == 1:
                            page_rows = rows
                        else:
                            time.sleep(random.uniform(float(cfg["sleep_min"]), float(cfg["sleep_max"])))
                            try:
                                _, page_rows = fetch_list_page(sess, cfg, keyword, second_code, dr, page_no)
                            except Exception as e:
                                logging.error("列表第 %s 页失败：%s", page_no, e)
                                continue

                        logging.info("处理第 %s/%s 页，记录数=%s", page_no, total_pages, len(page_rows))

                        for row in page_rows:
                            title = clean_text(row.get("title"))
                            if cfg.get("title_filter", True):
                                if normalize_for_match(keyword) not in normalize_for_match(title):
                                    continue

                            rec = base_record_from_list(province, keyword, ann_name, row)
                            article_id = rec.get("articleId", "")
                            unique_key = f"shanxi::{article_id}"
                            if not article_id:
                                unique_key = f"shanxi::{rec.get('annId')}::{rec.get('标题')}"
                            if unique_key in visited:
                                continue

                            if cfg.get("fetch_detail", True) and article_id:
                                time.sleep(random.uniform(float(cfg["detail_sleep_min"]), float(cfg["detail_sleep_max"])))
                                try:
                                    detail_html = fetch_detail_html(sess, cfg, article_id, rec.get("_parentId"))
                                    detail_fields = parse_detail_fields(detail_html, ann_name)
                                    merge_detail_fields(rec, detail_fields)
                                    rec["解析状态"] = "详情成功"
                                except Exception as e:
                                    rec["解析状态"] = "详情失败"
                                    rec["错误信息"] = str(e)
                                    logging.warning("详情解析失败：%s | %s", rec.get("标题"), e)

                            # 删除内部字段再保存。
                            rec.pop("_raw", None)
                            rec.pop("_parentId", None)
                            append_jsonl(jsonl_path, rec)
                            add_visited(visited_path, unique_key)
                            visited.add(unique_key)
                            total_saved += 1

                    time.sleep(random.uniform(float(cfg["sleep_min"]), float(cfg["sleep_max"])))

    except KeyboardInterrupt:
        logging.warning("用户中断，正在导出已抓取数据...")

    records = load_jsonl(jsonl_path)
    export_excel(records, excel_path)
    logging.info("山西爬取结束。本次新增 %s 条；累计 JSONL %s 条。", total_saved, len(records))
    return records


def main() -> None:
    parser = argparse.ArgumentParser(description="山西政府采购网标讯爬虫")
    parser.add_argument("--config", default=None, help="配置文件路径，可选，默认读取同目录 config.yaml")
    parser.add_argument("--keyword", action="append", help="临时指定关键词，可多次传入，例如 --keyword 营商环境 --keyword 水环境")
    parser.add_argument("--start-date", default=None, help="开始日期 YYYY-MM-DD")
    parser.add_argument("--end-date", default=None, help="结束日期 YYYY-MM-DD")
    parser.add_argument("--output-dir", default=None, help="输出目录")
    parser.add_argument("--no-detail", action="store_true", help="只抓列表，不抓详情")
    parser.add_argument("--no-title-filter", action="store_true", help="不做标题包含关键词过滤")
    args = parser.parse_args()

    cfg = load_config(args.config)
    if args.keyword:
        cfg["keywords"] = args.keyword
    if args.start_date:
        cfg["start_date"] = args.start_date
    if args.end_date:
        cfg["end_date"] = args.end_date
    if args.output_dir:
        cfg["output_dir"] = args.output_dir
    if args.no_detail:
        cfg["fetch_detail"] = False
    if args.no_title_filter:
        cfg["title_filter"] = False

    crawl_shanxi(cfg)


if __name__ == "__main__":
    main()
