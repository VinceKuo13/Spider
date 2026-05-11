# -*- coding: utf-8 -*-
"""
河北政府采购网爬虫：招标公告 + 中标公告

适用网站：
1) 招标公告：https://www.ccgp-hebei.gov.cn/province/cggg/zbgg/
2) 中标公告：https://www.ccgp-hebei.gov.cn/province/cggg/zhbgg/

核心特点：
1. 不使用 Selenium，直接请求 was5/web/search 接口。
2. 搜索范围固定为标题，即接口参数 doctitle=关键词。
3. 支持 2025-01-01 至今天的数据，自动按不超过 365 天切分时间段。
4. 同时抓取招标公告 zbgg 和中标公告 zhbgg。
5. 自动进入详情页解析采购人、代理机构、预算金额、最高限价、中标/成交供应商、中标/成交金额等字段。
6. 保存河北省单独 JSON，方便核对；同时导出 Excel。

运行：
    python spider_hebei.py

依赖：
    pip install requests beautifulsoup4 pandas openpyxl
"""

from __future__ import annotations

import json
import logging
import math
import os
import random
import re
import time
from copy import deepcopy
from datetime import date, datetime, timedelta
from html import unescape
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import quote, urlencode, urljoin

import requests
import urllib3
from bs4 import BeautifulSoup

try:
    import pandas as pd
except Exception:  # pragma: no cover
    pd = None

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# =========================
# 1. 可修改配置
# =========================

PROVINCE = "河北"
SITE_NAME = "河北政府采购网"
BASE_URL = "https://www.ccgp-hebei.gov.cn"
SEARCH_API = "https://www.ccgp-hebei.gov.cn/was5/web/search"
CHANNEL_ID = "217003"

# 关键词后续可以继续加，例如：KEYWORDS = ["营商环境", "数字政府", "政务服务"]
KEYWORDS = ["营商环境"]

# 默认 2025-01-01 至今天。END_DATE 留空表示运行当天。
START_DATE = "2025-01-01"
END_DATE = ""  # 例如 "2026-05-05"；空字符串表示今天

# 河北接口的查询时间跨度不要超过一年，这里按 365 天切分。
MAX_WINDOW_DAYS = 365

# 每页数量。你提供的接口 perpage=50，这里沿用。
PER_PAGE = 50
MAX_PAGES_PER_WINDOW = 200

# 请求间隔，太快容易被拦。可按实际情况调大。
LIST_DELAY_RANGE = (0.8, 1.8)
DETAIL_DELAY_RANGE = (0.6, 1.5)
RETRY_TIMES = 4
TIMEOUT = 25

# 是否校验 SSL。部分政采网站证书链不稳定，建议 False。
VERIFY_SSL = False

# 输出目录
OUTPUT_DIR = "outputs_hebei"
RAW_DIR = os.path.join(OUTPUT_DIR, "raw")
JSON_PATH = os.path.join(OUTPUT_DIR, "河北政府采购网_标讯明细.json")
EXCEL_PATH = os.path.join(OUTPUT_DIR, "河北政府采购网_标讯明细.xlsx")
PROGRESS_PATH = os.path.join(OUTPUT_DIR, "progress_state.json")
LOG_PATH = os.path.join(OUTPUT_DIR, "spider_hebei.log")

# 公告类型配置
ANN_TYPES = {
    "招标公告": {
        "lanmu": "zbgg",
        "referer": "https://www.ccgp-hebei.gov.cn/province/cggg/zbgg/",
    },
    "中标公告": {
        "lanmu": "zhbgg",
        "referer": "https://www.ccgp-hebei.gov.cn/province/cggg/zhbgg/",
    },
}


# =========================
# 2. 日志与基础工具
# =========================


# 读取同目录 config.yaml 中的 keywords/start_date/end_date。
from common_config import apply_common_config
apply_common_config(globals())

def ensure_dirs() -> None:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(RAW_DIR, exist_ok=True)


def setup_logging() -> None:
    ensure_dirs()
    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
        ],
    )


def today_str() -> str:
    return date.today().strftime("%Y-%m-%d")


def parse_date_obj(s: str) -> date:
    s = s.strip().replace("/", "-").replace("年", "-").replace("月", "-").replace("日", "")
    return datetime.strptime(s, "%Y-%m-%d").date()


def date_windows(start: str, end: str, max_days: int = 365) -> List[Tuple[str, str]]:
    s = parse_date_obj(start)
    e = parse_date_obj(end)
    if s > e:
        raise ValueError(f"开始日期不能晚于结束日期：{start} > {end}")

    windows: List[Tuple[str, str]] = []
    cur = s
    while cur <= e:
        win_end = min(cur + timedelta(days=max_days), e)
        windows.append((cur.strftime("%Y-%m-%d"), win_end.strftime("%Y-%m-%d")))
        cur = win_end + timedelta(days=1)
    return windows


def sleep_random(rng: Tuple[float, float]) -> None:
    time.sleep(random.uniform(*rng))


def clean_text(text: Any) -> str:
    if text is None:
        return ""
    text = str(text)
    text = unescape(text)
    text = re.sub(r"<\s*em\s*>", "", text, flags=re.I)
    text = re.sub(r"<\s*/\s*em\s*>", "", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("\xa0", " ").replace("&nbsp;", " ")
    text = text.replace("　", " ")
    text = re.sub(r"[\t\r\f\v]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    text = re.sub(r"[ ]{2,}", " ", text)
    return text.strip()


def soup_text(soup: BeautifulSoup) -> str:
    # 去掉脚本、样式、浏览器插件残留
    for tag in soup(["script", "style", "noscript", "iframe"]):
        tag.decompose()
    for tag in soup.select("#immersive-translate-popup"):
        tag.decompose()
    text = soup.get_text("\n")
    lines = [clean_text(x) for x in text.splitlines()]
    lines = [x for x in lines if x]
    return "\n".join(lines)


def normalize_date_text(s: str) -> str:
    s = clean_text(s)
    m = re.search(r"(20\d{2})[年/\-\.](\d{1,2})[月/\-\.](\d{1,2})", s)
    if not m:
        return ""
    y, mo, d = m.groups()
    return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"


def unique_join(items: Iterable[Any], sep: str = "; ") -> str:
    seen = set()
    out = []
    for x in items:
        x = clean_text(x)
        if not x or x.lower() == "null" or x in seen:
            continue
        seen.add(x)
        out.append(x)
    return sep.join(out)


def safe_filename(s: str, max_len: int = 80) -> str:
    s = clean_text(s)
    s = re.sub(r"[\\/:*?\"<>|\s]+", "_", s)
    return s[:max_len] or "empty"


# =========================
# 3. 请求函数
# =========================

def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/147.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Connection": "keep-alive",
        }
    )
    # 河北网站有时需要 cityurl cookie，不固定 JSESSIONID 也能请求。
    session.cookies.set("cityurl", "..%2Fprovince%2F", domain="www.ccgp-hebei.gov.cn")
    return session


def request_text(
    session: requests.Session,
    method: str,
    url: str,
    *,
    headers: Optional[Dict[str, str]] = None,
    params: Optional[Dict[str, Any]] = None,
    data: Optional[Any] = None,
    retry_times: int = RETRY_TIMES,
) -> str:
    last_err: Optional[Exception] = None
    for i in range(1, retry_times + 1):
        try:
            resp = session.request(
                method=method.upper(),
                url=url,
                headers=headers,
                params=params,
                data=data,
                timeout=TIMEOUT,
                verify=VERIFY_SSL,
            )
            resp.raise_for_status()
            # 优先使用响应或页面标注的编码，不准确时用 apparent_encoding 兜底
            if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
                resp.encoding = resp.apparent_encoding or "utf-8"
            text = resp.text
            # 常见乱码兜底
            if "锟" in text[:500] or "Ã" in text[:500]:
                try:
                    text = resp.content.decode("utf-8", errors="ignore")
                except Exception:
                    pass
            return text
        except Exception as e:  # noqa: BLE001
            last_err = e
            wait_s = min(2 ** i, 12) + random.uniform(0, 1)
            logging.warning("请求失败 %s/%s：%s，等待 %.1f 秒；错误：%s", i, retry_times, url, wait_s, e)
            time.sleep(wait_s)
    raise RuntimeError(f"请求失败：{url}；最后错误：{last_err}")


def init_referer(session: requests.Session, referer: str) -> None:
    try:
        request_text(session, "GET", referer, headers={"Referer": BASE_URL + "/"}, retry_times=2)
    except Exception as e:  # noqa: BLE001
        logging.warning("初始化栏目页失败，可继续尝试接口：%s", e)


# =========================
# 4. 列表页解析
# =========================

def build_search_params(keyword: str, lanmu: str, start_date: str, end_date: str, page: int) -> Dict[str, Any]:
    # 注意：搜索范围是标题，对应 doctitle=关键词。
    return {
        "channelid": CHANNEL_ID,
        "lanmu": lanmu,
        "admindivcode": "",
        "purchaseWay": "",
        "procurementcode": "",
        "agencyfullname": "",
        "PurchaserName": "",
        "doctitle": keyword,
        "fstarttime": start_date,
        "fendtime": end_date,
        "page": page,
        "perpage": PER_PAGE,
    }


def search_url_for_log(params: Dict[str, Any]) -> str:
    return SEARCH_API + "?" + urlencode(params, doseq=True)


def extract_total_count(text: str) -> Optional[int]:
    text2 = clean_text(text)
    patterns = [
        r"搜索结果[^0-9]{0,20}(\d+)\s*条",
        r"共\s*(\d+)\s*条",
        r"total[^0-9]{0,10}(\d+)",
        r"recordcount[^0-9]{0,10}(\d+)",
    ]
    for pat in patterns:
        m = re.search(pat, text2, flags=re.I)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                pass
    return None


def recursive_find_records(obj: Any) -> List[Dict[str, Any]]:
    """有些接口可能返回 JSON，这里递归找包含 title/url/href 的对象。"""
    records: List[Dict[str, Any]] = []
    if isinstance(obj, dict):
        keys = {str(k).lower() for k in obj.keys()}
        if ("url" in keys or "href" in keys or "link" in keys or "docurl" in keys) and (
            "title" in keys or "doctitle" in keys or "name" in keys
        ):
            records.append(obj)
        for v in obj.values():
            records.extend(recursive_find_records(v))
    elif isinstance(obj, list):
        for it in obj:
            records.extend(recursive_find_records(it))
    return records


def item_from_json_record(d: Dict[str, Any], ann_type: str, keyword: str) -> Optional[Dict[str, Any]]:
    def first_key(*names: str) -> str:
        low_map = {str(k).lower(): k for k in d.keys()}
        for name in names:
            k = low_map.get(name.lower())
            if k is not None:
                return clean_text(d.get(k))
        return ""

    title = first_key("title", "doctitle", "name")
    href = first_key("url", "href", "link", "docurl")
    if not href or ".html" not in href:
        return None
    url = urljoin(BASE_URL + "/", href)
    pub = first_key("date", "pubdate", "publishdate", "发布时间", "fbrq")
    pub = normalize_date_text(pub) or pub
    return {
        "province": PROVINCE,
        "site": SITE_NAME,
        "keyword": keyword,
        "announcement_type": ann_type,
        "title": clean_text(title),
        "publish_date": pub,
        "detail_url": url,
        "list_region": first_key("district", "region", "area", "地区"),
        "list_purchaser": first_key("purchaser", "purchasername", "采购人"),
        "list_raw_text": "",
    }


def parse_list_items(response_text: str, ann_type: str, keyword: str) -> Tuple[List[Dict[str, Any]], Optional[int]]:
    """解析 was5/web/search 的返回。兼容 JSON、HTML、HTML 片段。"""
    total = extract_total_count(response_text)
    items: List[Dict[str, Any]] = []

    # 1) JSON 兜底
    try:
        obj = json.loads(response_text)
        for rec in recursive_find_records(obj):
            item = item_from_json_record(rec, ann_type, keyword)
            if item:
                items.append(item)
        if items:
            return dedup_list_items(items), total
    except Exception:
        pass

    # 2) 有些 WAS5 返回 JS 字符串或 document.write，先反转义
    html = response_text
    html = html.replace("\\/", "/")
    html = html.replace("\\\"", "\"").replace("\\'", "'")
    html = unescape(html)

    soup = BeautifulSoup(html, "html.parser")

    # 找所有详情链接
    candidates = []
    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        title = clean_text(a.get_text(" ")) or clean_text(a.get("title", ""))
        if not href or ".html" not in href:
            continue
        if "/cggg/" not in href and "zbgg" not in href and "zhbgg" not in href:
            continue
        if not title or title in {"首页", "上一页", "下一页", "末页", "关闭窗口"}:
            continue
        candidates.append((a, href, title))

    # 正则兜底：如果 a 标签没有被 BeautifulSoup 正确解析
    if not candidates:
        url_patterns = [
            r"https?://www\.ccgp-hebei\.gov\.cn/[^\"'<>\s]+?\.html",
            r"(?:\.\./)+[^\"'<>\s]+?\.html",
            r"/[a-zA-Z0-9_\-/]+?\.html",
        ]
        seen_urls = set()
        for pat in url_patterns:
            for m in re.finditer(pat, html):
                href = m.group(0)
                url = urljoin(BASE_URL + "/province/cggg/", href)
                if url in seen_urls or "/cggg/" not in url:
                    continue
                seen_urls.add(url)
                # 在链接附近取标题
                start = max(0, m.start() - 300)
                end = min(len(html), m.end() + 300)
                near = clean_text(html[start:end])
                title = ""
                # 尝试取链接后或 title 属性
                mt = re.search(r"title=[\"']([^\"']+)[\"']", html[start:end], flags=re.I)
                if mt:
                    title = clean_text(mt.group(1))
                if not title:
                    title = clean_text(re.sub(r"https?://\S+|(?:\.\./)+\S+?\.html", " ", near))[:120]
                if title:
                    items.append(
                        {
                            "province": PROVINCE,
                            "site": SITE_NAME,
                            "keyword": keyword,
                            "announcement_type": ann_type,
                            "title": title,
                            "publish_date": normalize_date_text(near),
                            "detail_url": url,
                            "list_region": extract_simple_field(near, ["地区"], ["采购人", "发布时间"]),
                            "list_purchaser": extract_simple_field(near, ["采购人"], ["代理机构", "发布时间", "地区"]),
                            "list_raw_text": near,
                        }
                    )
        return dedup_list_items(items), total

    for a, href, title in candidates:
        url = urljoin(BASE_URL + "/", href)
        # 取附近容器文本
        parent = a
        for _ in range(4):
            if parent.parent is None:
                break
            parent = parent.parent
            ptxt = clean_text(parent.get_text(" "))
            # 一条记录通常会包含发布时间、地区、采购人等；太长可能是整个列表
            if ("发布时间" in ptxt or "采购人" in ptxt or "地区" in ptxt) and len(ptxt) < 1500:
                break
        raw = clean_text(parent.get_text(" ")) if parent else title
        if len(raw) > 1500:
            raw = title
        items.append(
            {
                "province": PROVINCE,
                "site": SITE_NAME,
                "keyword": keyword,
                "announcement_type": ann_type,
                "title": title,
                "publish_date": normalize_date_text(raw),
                "detail_url": url,
                "list_region": extract_simple_field(raw, ["地区"], ["采购人", "发布时间"]),
                "list_purchaser": extract_simple_field(raw, ["采购人"], ["代理机构", "发布时间", "地区"]),
                "list_raw_text": raw,
            }
        )

    return dedup_list_items(items), total


def dedup_list_items(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen = set()
    out: List[Dict[str, Any]] = []
    for item in items:
        key = item.get("detail_url") or (item.get("title"), item.get("publish_date"))
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out


# =========================
# 5. 详情页解析
# =========================

def extract_simple_field(text: str, labels: List[str], stop_labels: List[str], max_len: int = 200) -> str:
    text = clean_text(text)
    flat = re.sub(r"\s+", " ", text)
    for label in labels:
        # 支持 “名 称” 这种中间有空格的标签
        lab_pat = "\\s*".join(map(re.escape, label))
        stop_pat = "|".join("\\s*".join(map(re.escape, s)) for s in stop_labels) if stop_labels else r"$"
        pat = rf"{lab_pat}\s*[:：]?\s*(.*?)(?={stop_pat}\s*[:：]?|$)"
        m = re.search(pat, flat)
        if m:
            val = clean_text(m.group(1))
            val = re.sub(r"^[：:]+", "", val).strip()
            return val[:max_len]
    return ""


def block_between(text: str, start_labels: List[str], end_labels: List[str]) -> str:
    flat = re.sub(r"\s+", " ", clean_text(text))
    start_pos = -1
    for lab in start_labels:
        m = re.search(re.escape(lab), flat)
        if m:
            start_pos = m.end()
            break
    if start_pos < 0:
        return ""
    end_pos = len(flat)
    for lab in end_labels:
        m = re.search(re.escape(lab), flat[start_pos:])
        if m:
            end_pos = min(end_pos, start_pos + m.start())
    return flat[start_pos:end_pos].strip()


def parse_table_by_id(soup: BeautifulSoup, table_id: str) -> List[Dict[str, str]]:
    table = soup.find("table", id=table_id)
    if table is None:
        return []
    rows = []
    tr_list = table.find_all("tr")
    if not tr_list:
        return []

    headers: List[str] = []
    for cell in tr_list[0].find_all(["th", "td"]):
        headers.append(clean_text(cell.get_text(" ")))
    if not headers:
        return []

    for tr in tr_list[1:]:
        cells = [clean_text(td.get_text(" ")) for td in tr.find_all(["td", "th"])]
        if not any(cells):
            continue
        row: Dict[str, str] = {}
        for idx, val in enumerate(cells):
            key = headers[idx] if idx < len(headers) else f"列{idx + 1}"
            row[key] = val
        rows.append(row)
    return rows


def parse_all_tables(soup: BeautifulSoup) -> List[Dict[str, str]]:
    parsed: List[Dict[str, str]] = []
    for table in soup.find_all("table"):
        tr_list = table.find_all("tr")
        if len(tr_list) < 2:
            continue
        headers = [clean_text(c.get_text(" ")) for c in tr_list[0].find_all(["th", "td"])]
        if not headers:
            continue
        for tr in tr_list[1:]:
            cells = [clean_text(td.get_text(" ")) for td in tr.find_all(["td", "th"])]
            if not any(cells):
                continue
            row = {headers[i] if i < len(headers) else f"列{i + 1}": v for i, v in enumerate(cells)}
            parsed.append(row)
    return parsed


def row_get(row: Dict[str, str], contains: List[str]) -> str:
    for k, v in row.items():
        if all(x in k for x in contains):
            return clean_text(v)
    return ""


def parse_supplier_and_amount(soup: BeautifulSoup, text: str) -> Tuple[str, str, str]:
    """返回：供应商、金额、标的信息简表。"""
    supplier_names: List[str] = []
    amounts: List[str] = []
    bid_rows_summary: List[str] = []

    # 1) 三、中标（成交）信息：SupplierInfos
    for row in parse_table_by_id(soup, "SupplierInfos"):
        supplier = row_get(row, ["供应商", "名称"]) or row_get(row, ["供应商名称"])
        if supplier:
            supplier_names.append(supplier)

    # 2) 四、主要标的信息：Goods/Engineering/Service 表格
    for tid in ["GoodsSupplierInfo", "EngineeringSupplierInfo", "ServiceSupplierInfo", "detail"]:
        for row in parse_table_by_id(soup, tid):
            supplier = row_get(row, ["供应商", "名称"]) or row_get(row, ["供应商名单"])
            amount = ""
            for k, v in row.items():
                if any(x in k for x in ["中标金额", "成交金额", "中标价", "成交价"]):
                    amount = clean_text(v)
                    break
            if supplier:
                supplier_names.append(supplier)
            if amount:
                amounts.append(amount)
            if supplier or amount:
                bid_rows_summary.append(f"供应商={supplier}; 金额={amount}; " + json.dumps(row, ensure_ascii=False))

    # 3) 任意表格兜底
    if not supplier_names or not amounts:
        for row in parse_all_tables(soup):
            supplier = row_get(row, ["供应商", "名称"])
            amount = ""
            for k, v in row.items():
                if any(x in k for x in ["中标金额", "成交金额", "中标价", "成交价"]):
                    amount = clean_text(v)
                    break
            if supplier:
                supplier_names.append(supplier)
            if amount:
                amounts.append(amount)

    # 4) 文本兜底
    flat = re.sub(r"\s+", " ", clean_text(text))
    if not supplier_names:
        for pat in [
            r"供应商名称\s*[:：]?\s*([^\s，,；;。]+)",
            r"中标供应商\s*[:：]?\s*([^\s，,；;。]+)",
            r"成交供应商\s*[:：]?\s*([^\s，,；;。]+)",
            r"供应商名单\s*[:：]?\s*([^\s，,；;。]+)",
        ]:
            supplier_names.extend(re.findall(pat, flat))
    if not amounts:
        for pat in [
            r"中标金额\s*[:：]?\s*([0-9][0-9,\.]*\s*(?:元|万元|亿元|人民币)?)",
            r"成交金额\s*[:：]?\s*([0-9][0-9,\.]*\s*(?:元|万元|亿元|人民币)?)",
            r"中标价\s*[:：]?\s*([0-9][0-9,\.]*\s*(?:元|万元|亿元|人民币)?)",
            r"成交价\s*[:：]?\s*([0-9][0-9,\.]*\s*(?:元|万元|亿元|人民币)?)",
        ]:
            amounts.extend(re.findall(pat, flat))

    return unique_join(supplier_names), unique_join(amounts), unique_join(bid_rows_summary, sep=" || ")


def parse_raw_pack_detail_if_needed(soup: BeautifulSoup, parsed: Dict[str, Any]) -> None:
    """
    河北详情页有些字段由 JS content() 从隐藏的 con 字段中拆出来。
    如果 requests 拿到的是未执行 JS 的原始页，这里尽量从 #con 原始字符串里提取供应商/金额。
    """
    con = soup.find(id="con")
    if not con:
        return
    raw = clean_text(con.get_text(" "))
    if not raw or "#detail#" not in raw:
        return

    # 招标公告：#detail# 前通常是采购需求
    before = raw.split("#detail#", 1)[0]
    if before and not parsed.get("采购需求"):
        parsed["采购需求"] = before

    # 中标公告：#detail# 后可能包含 packdetail#filename#附件
    if "#filename#" in raw and not parsed.get("中标/成交金额"):
        try:
            temp = raw.split("#detail#", 1)[1]
            packdetail = temp.split("#filename#", 1)[0]
            pack = packdetail.split("#_@_@")
            # 老版常见：pack[1] 供应商，pack[2] 金额；新版 V2020：pack[3] 供应商，pack[10] 金额
            suppliers: List[str] = []
            amounts: List[str] = []
            if len(pack) > 3:
                suppliers.extend(pack[3].split("#_#"))
            if len(pack) > 10:
                amounts.extend(pack[10].split("#_#"))
            if len(pack) > 2:
                suppliers.extend(pack[1].split("#_#"))
                amounts.extend(pack[2].split("#_#"))
            if suppliers and not parsed.get("中标/成交供应商"):
                parsed["中标/成交供应商"] = unique_join(suppliers)
            if amounts and not parsed.get("中标/成交金额"):
                parsed["中标/成交金额"] = unique_join(amounts)
        except Exception:
            pass


def parse_detail_page(html: str, url: str, ann_type: str) -> Dict[str, Any]:
    soup = BeautifulSoup(html, "html.parser")

    meta_title = soup.find("meta", attrs={"name": "ArticleTitle"})
    meta_date = soup.find("meta", attrs={"name": "PubDate"})
    title = clean_text(meta_title.get("content")) if meta_title else ""
    pub_date = clean_text(meta_date.get("content")) if meta_date else ""

    if not title:
        h = soup.select_one(".txt2") or soup.find("title")
        title = clean_text(h.get_text(" ")) if h else ""
    if title in {"招标公告详细页面", "中标公告详细页面"}:
        h = soup.select_one(".txt2")
        title = clean_text(h.get_text(" ")) if h else title

    text = soup_text(deepcopy(soup))
    if not pub_date:
        pub_date = normalize_date_text(extract_simple_field(text, ["发布时间"], ["一、", "一、项目", "项目编号"]))
    pub_date = normalize_date_text(pub_date) or pub_date

    purchaser_block = block_between(
        text,
        ["1.采购人信息", "1、采购人信息", "采购人信息"],
        ["2.采购代理机构信息", "2、采购代理机构信息", "采购代理机构信息", "3.项目联系方式"],
    )
    agent_block = block_between(
        text,
        ["2.采购代理机构信息", "2、采购代理机构信息", "采购代理机构信息"],
        ["3.项目联系方式", "项目联系方式", "十、附件", "九、凡对本次公告内容提出询问"],
    )

    parsed: Dict[str, Any] = {
        "标题": title,
        "发布时间": pub_date,
        "项目编号": extract_simple_field(text, ["项目编号"], ["项目名称", "二、", "采购项目名称"]),
        "项目名称": extract_simple_field(text, ["项目名称"], ["预算金额", "三、", "采购需求", "中标", "成交"]),
        "采购人": extract_simple_field(purchaser_block, ["名 称", "名称"], ["地址", "联系方式", "联系人"]),
        "采购代理机构": extract_simple_field(agent_block, ["名 称", "名称"], ["地址", "地 址", "地　址", "联系方式", "项目联系方式"]),
        "预算金额": extract_simple_field(text, ["预算金额"], ["最高限价", "采购需求", "合同履行期限"]),
        "最高限价": extract_simple_field(text, ["最高限价"], ["采购需求", "合同履行期限", "二、"]),
        "采购需求": extract_simple_field(text, ["采购需求"], ["合同履行期限", "本项目", "二、申请人"]),
        "采购方式": extract_simple_field(text, ["采购方式"], ["采购单位", "采购人", "代理机构", "预算金额"]),
        "中标/成交供应商": "",
        "中标/成交金额": "",
        "标的信息": "",
        "详情链接": url,
        "正文": text[:20000],
    }

    if ann_type == "中标公告":
        supplier, amount, bid_info = parse_supplier_and_amount(soup, text)
        parsed["中标/成交供应商"] = supplier
        parsed["中标/成交金额"] = amount
        parsed["标的信息"] = bid_info

    parse_raw_pack_detail_if_needed(soup, parsed)

    # 清理明显过长字段
    for k, v in list(parsed.items()):
        if isinstance(v, str):
            parsed[k] = clean_text(v)
    return parsed


# =========================
# 6. 状态保存与导出
# =========================

def load_json_list(path: str) -> List[Dict[str, Any]]:
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def save_json_list(path: str, records: List[Dict[str, Any]]) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def load_progress() -> Dict[str, Any]:
    if not os.path.exists(PROGRESS_PATH):
        return {"visited_urls": []}
    try:
        with open(PROGRESS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {"visited_urls": []}
        data.setdefault("visited_urls", [])
        return data
    except Exception:
        return {"visited_urls": []}


def save_progress(visited_urls: Iterable[str]) -> None:
    tmp = PROGRESS_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"visited_urls": sorted(set(visited_urls))}, f, ensure_ascii=False, indent=2)
    os.replace(tmp, PROGRESS_PATH)


def normalize_records_for_export(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    # 去重：优先按详情链接；无链接时按 标题+发布时间+公告类型
    seen = set()
    out = []
    for r in records:
        key = r.get("详情链接") or r.get("detail_url") or (
            r.get("标题") or r.get("title"),
            r.get("发布时间") or r.get("publish_date"),
            r.get("公告类型") or r.get("announcement_type"),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def export_excel(records: List[Dict[str, Any]]) -> None:
    records = normalize_records_for_export(records)
    if not records:
        logging.warning("没有可导出的记录。")
        return

    columns = [
        "省份",
        "网站",
        "关键词",
        "公告类型",
        "标题",
        "发布时间",
        "地区",
        "采购人",
        "采购代理机构",
        "项目编号",
        "项目名称",
        "采购方式",
        "预算金额",
        "最高限价",
        "中标/成交供应商",
        "中标/成交金额",
        "详情链接",
        "标的信息",
        "正文",
    ]

    if pd is None:
        # pandas 不可用时至少保存 JSON
        logging.warning("未安装 pandas，已保存 JSON，无法导出 Excel。请执行 pip install pandas openpyxl")
        return

    df = pd.DataFrame(records)
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    df = df[columns]
    df.to_excel(EXCEL_PATH, index=False)
    logging.info("Excel 已导出：%s，共 %s 条", EXCEL_PATH, len(df))


def append_record(records: List[Dict[str, Any]], record: Dict[str, Any]) -> None:
    records.append(record)
    save_json_list(JSON_PATH, normalize_records_for_export(records))


# =========================
# 7. 主爬取逻辑
# =========================

def crawl_one(
    session: requests.Session,
    keyword: str,
    ann_type: str,
    lanmu: str,
    referer: str,
    start_date: str,
    end_date: str,
    records: List[Dict[str, Any]],
    visited_urls: set,
) -> None:
    logging.info("开始任务：关键词=%s，公告类型=%s，时间=%s 至 %s", keyword, ann_type, start_date, end_date)
    init_referer(session, referer)

    previous_page_urls: set = set()
    total_pages_hint: Optional[int] = None

    for page in range(1, MAX_PAGES_PER_WINDOW + 1):
        params = build_search_params(keyword, lanmu, start_date, end_date, page)
        headers = {
            "Accept": "*/*",
            "Referer": referer,
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "X-Requested-With": "XMLHttpRequest",
        }
        url_for_log = search_url_for_log(params)
        logging.info("列表页：%s", url_for_log)

        try:
            list_text = request_text(session, "GET", SEARCH_API, headers=headers, params=params)
        except Exception as e:  # noqa: BLE001
            logging.error("列表页失败，跳过该页：%s", e)
            break

        # 保存少量原始列表响应，便于排错
        raw_name = f"{safe_filename(keyword)}_{ann_type}_{start_date}_{end_date}_p{page}.html"
        try:
            with open(os.path.join(RAW_DIR, raw_name), "w", encoding="utf-8") as f:
                f.write(list_text)
        except Exception:
            pass

        items, total = parse_list_items(list_text, ann_type, keyword)
        if total is not None and total_pages_hint is None:
            total_pages_hint = max(1, math.ceil(total / PER_PAGE))
            logging.info("接口显示总数约 %s 条，总页数约 %s", total, total_pages_hint)

        logging.info("第 %s 页解析到 %s 条列表记录", page, len(items))
        if not items:
            break

        page_urls = {x.get("detail_url", "") for x in items if x.get("detail_url")}
        if page_urls and page_urls == previous_page_urls:
            logging.warning("第 %s 页与上一页链接完全相同，停止翻页，避免死循环。", page)
            break
        previous_page_urls = page_urls

        new_count = 0
        for item in items:
            detail_url = item.get("detail_url", "")
            title = clean_text(item.get("title", ""))
            if not detail_url:
                continue
            if detail_url in visited_urls:
                continue

            logging.info("详情：%s", title)
            detail_headers = {
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Referer": referer,
            }
            try:
                detail_html = request_text(session, "GET", detail_url, headers=detail_headers)
                detail = parse_detail_page(detail_html, detail_url, ann_type)
            except Exception as e:  # noqa: BLE001
                logging.warning("详情页解析失败，保留列表字段：%s；错误：%s", detail_url, e)
                detail = {
                    "标题": title,
                    "发布时间": item.get("publish_date", ""),
                    "详情链接": detail_url,
                    "正文": "",
                }

            record = {
                "省份": PROVINCE,
                "网站": SITE_NAME,
                "关键词": keyword,
                "公告类型": ann_type,
                "标题": detail.get("标题") or title,
                "发布时间": detail.get("发布时间") or item.get("publish_date", ""),
                "地区": item.get("list_region", ""),
                "采购人": detail.get("采购人") or item.get("list_purchaser", ""),
                "采购代理机构": detail.get("采购代理机构", ""),
                "项目编号": detail.get("项目编号", ""),
                "项目名称": detail.get("项目名称", ""),
                "采购方式": detail.get("采购方式", ""),
                "预算金额": detail.get("预算金额", ""),
                "最高限价": detail.get("最高限价", ""),
                "中标/成交供应商": detail.get("中标/成交供应商", ""),
                "中标/成交金额": detail.get("中标/成交金额", ""),
                "详情链接": detail_url,
                "标的信息": detail.get("标的信息", ""),
                "正文": detail.get("正文", ""),
            }
            append_record(records, record)
            visited_urls.add(detail_url)
            save_progress(visited_urls)
            new_count += 1
            sleep_random(DETAIL_DELAY_RANGE)

        logging.info("第 %s 页新增 %s 条；当前累计 %s 条", page, new_count, len(normalize_records_for_export(records)))

        # 停止条件
        if len(items) < PER_PAGE:
            break
        if total_pages_hint is not None and page >= total_pages_hint:
            break

        sleep_random(LIST_DELAY_RANGE)


def main() -> None:
    ensure_dirs()
    setup_logging()

    end_date = END_DATE.strip() or today_str()
    windows = date_windows(START_DATE, end_date, MAX_WINDOW_DAYS)
    logging.info("开始爬取：%s；搜索范围=标题 doctitle；关键词=%s", SITE_NAME, KEYWORDS)
    logging.info("总时间范围：%s 至 %s；自动切分：%s", START_DATE, end_date, windows)

    records = load_json_list(JSON_PATH)
    progress = load_progress()
    visited_urls = set(progress.get("visited_urls", []))
    # 如果已有 JSON，但 progress 丢失，也从 JSON 里恢复
    for r in records:
        if r.get("详情链接"):
            visited_urls.add(r["详情链接"])

    session = make_session()

    try:
        for keyword in KEYWORDS:
            for ann_type, cfg in ANN_TYPES.items():
                for s, e in windows:
                    crawl_one(
                        session=session,
                        keyword=keyword,
                        ann_type=ann_type,
                        lanmu=cfg["lanmu"],
                        referer=cfg["referer"],
                        start_date=s,
                        end_date=e,
                        records=records,
                        visited_urls=visited_urls,
                    )
        records = normalize_records_for_export(records)
        save_json_list(JSON_PATH, records)
        export_excel(records)
        logging.info("全部完成。JSON：%s", JSON_PATH)
    except KeyboardInterrupt:
        logging.warning("用户中断，正在导出已抓取数据...")
        records = normalize_records_for_export(records)
        save_json_list(JSON_PATH, records)
        save_progress(visited_urls)
        export_excel(records)
    except Exception:
        logging.exception("运行失败，正在导出已抓取数据...")
        records = normalize_records_for_export(records)
        save_json_list(JSON_PATH, records)
        save_progress(visited_urls)
        export_excel(records)
        raise


if __name__ == "__main__":
    main()
