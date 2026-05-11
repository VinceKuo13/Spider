# -*- coding: utf-8 -*-
"""
北京市公共资源交易服务平台 - 政府采购公告/成交结果公告爬虫

适用站点：
https://ggzyfw.beijing.gov.cn/elasticsearch/index.jsp?c1=jyxx&c2=jyxxzfcg&...

核心功能：
1. 不需要登录，不使用 Selenium。
2. 直接请求搜索接口：https://ggzyfw.beijing.gov.cn/elasticsearch/search
3. 搜索范围固定为标题：scope=title。
4. 默认抓取 2025-01-01 至当天。
5. 同时抓取：采购公告 jyxxcggg、成交结果公告 jyxxzbjggg。
6. 解析详情页字段：项目名称、项目编号、交易项目编号、采购人、代理机构、采购方式、预算金额、最高限价、中标/成交供应商、中标/成交金额等。
7. 保存 JSONL 断点、visited_urls 去重、Excel 明细表和项目合并表。

运行：
    python beijing_ggzy_zfcg_spider.py

首次运行会自动读取同目录 config.yaml 中的关键词和日期；其他参数使用 DEFAULT_CONFIG。
"""

from __future__ import annotations

import json
import logging
import math
import random
import re
import sys
import time
import ssl
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests import Session
from requests.adapters import HTTPAdapter

try:
    import yaml
except ImportError:  # 允许没有 pyyaml 时直接用默认配置运行
    yaml = None


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.yaml"
BASE_URL = "https://ggzyfw.beijing.gov.cn"
SEARCH_URL = f"{BASE_URL}/elasticsearch/search"

DEFAULT_CONFIG: Dict[str, Any] = {
    "keywords": ["营商环境"],
    "start_date": "2025-01-01",
    "end_date": "",  # 留空表示当天
    "scope": "title",  # 标题搜索，不是全文
    "notice_types": {
        "采购公告": "jyxxcggg",
        "成交结果公告": "jyxxzbjggg",
    },
    "max_pages_per_query": 9999,
    "request_delay": 1.0,
    "detail_delay": 0.8,
    "timeout": 30,
    "retry_times": 4,
    "http_backend": "auto",  # auto / curl_cffi / requests
    "curl_impersonate": "chrome",
    "auto_export_every_records": 20,
    "output_dir": "outputs_beijing_ggzy",
    "checkpoint_jsonl": "outputs_beijing_ggzy/beijing_ggzy_zfcg_records.jsonl",
    "visited_urls_file": "outputs_beijing_ggzy/visited_urls.txt",
    "progress_state_file": "outputs_beijing_ggzy/progress_state.json",
    "output_excel_detail": "outputs_beijing_ggzy/北京公共资源_政府采购_公告明细.xlsx",
    "output_excel_merged": "outputs_beijing_ggzy/北京公共资源_政府采购_项目合并.xlsx",
    "log_file": "outputs_beijing_ggzy/beijing_ggzy_zfcg_spider.log",
}


@dataclass
class NoticeRecord:
    search_keyword: str = ""
    notice_type: str = ""
    channel_third: str = ""

    project_name: str = ""
    transaction_project_code: str = ""  # 交易项目编号，如 S110000C005089650001
    project_no: str = ""  # 政府采购项目编号，如 11010525210200021060-XM001

    tender_unit: str = ""  # 采购人/招标人
    agency: str = ""  # 采购代理机构
    purchase_type: str = ""  # 公开招标/竞争性磋商等
    budget_amount: str = ""
    max_price: str = ""

    winner_unit: str = ""  # 中标/成交供应商
    winner_amount: str = ""  # 中标/成交金额，供应商行金额
    total_winner_amount: str = ""  # 总中标成交金额
    supplier_address: str = ""
    unified_credit_code: str = ""
    review_score_or_remark: str = ""

    publish_time: str = ""
    notice_end_time: str = ""
    source_region: str = ""
    source: str = ""
    detail_url: str = ""
    raw_title: str = ""
    raw_text_summary: str = ""

    parse_status: str = ""
    error: str = ""


FIELD_ALIASES = {
    "project_no": ["项目编号", "项目号", "招标编号", "采购编号", "项目编码", "采购代理机构项目编号"],
    "project_name": ["项目名称", "采购项目名称", "标项名称", "标的名称"],
    "tender_unit": ["采购人", "采购单位", "采购人信息", "招标人", "招标单位", "建设单位", "业主单位"],
    "agency": ["采购代理机构", "代理机构", "招标代理", "采购代理", "代理单位"],
    "purchase_type": ["采购方式", "招标方式", "采购形式", "招标形式"],
    "budget_amount": ["预算金额", "采购预算"],
    "max_price": ["最高限价"],
    "winner_unit": [
        "中标成交供应商名称", "中标（成交）供应商名称", "中标(成交)供应商名称",
        "成交供应商名称", "中标供应商名称", "供应商名称", "中标人", "成交人",
    ],
    "winner_amount": ["中标金额", "成交金额", "中标成交金额", "中标（成交）金额", "中标(成交)金额"],
    "total_winner_amount": ["总中标成交金额", "总成交金额", "总中标金额"],
}

PURCHASE_TYPE_KEYWORDS = [
    "公开招标", "邀请招标", "竞争性磋商", "竞争性谈判", "询价", "比选", "单一来源", "框架协议",
    "公开询价", "询比价", "谈判", "磋商", "竞价", "遴选", "比价",
]


# ---------- 基础工具 ----------

def load_config(path: Path = CONFIG_PATH) -> Dict[str, Any]:
    cfg = dict(DEFAULT_CONFIG)
    if path.exists():
        if yaml is None:
            raise RuntimeError("检测到 config.yaml，但当前环境没有 pyyaml，请执行：pip install pyyaml")
        with path.open("r", encoding="utf-8") as f:
            user_cfg = yaml.safe_load(f) or {}
        cfg.update(user_cfg)
        if "notice_types" in user_cfg:
            cfg["notice_types"] = user_cfg["notice_types"]
    return cfg


def setup_logging(cfg: Dict[str, Any]) -> None:
    log_path = BASE_DIR / cfg.get("log_file", DEFAULT_CONFIG["log_file"])
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def parse_date(s: str) -> date:
    return datetime.strptime(s[:10], "%Y-%m-%d").date()


def today_str() -> str:
    return date.today().strftime("%Y-%m-%d")


def normalize_space(text: str) -> str:
    if text is None:
        return ""
    text = str(text).replace("\xa0", " ").replace("\u3000", " ").replace("\ufeff", "")
    text = re.sub(r"[\t\r\f\v]+", " ", text)
    text = re.sub(r"[ ]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def clean_value(value: str) -> str:
    value = strip_html(fix_mojibake(value))
    value = normalize_space(value)
    value = re.sub(r"^[：:\-—\s]+", "", value)
    value = re.sub(r"[。；;，,\s]+$", "", value)
    return value.strip()


def fix_mojibake(text: Any) -> str:
    """修复常见 UTF-8 被按 latin1 显示造成的乱码，如 é¡¹ç›® -> 项目。"""
    if text is None:
        return ""
    s = str(text)
    if not s:
        return ""
    # 常见中文 UTF-8 误解码后的字符特征
    suspicious = ("é" in s or "å" in s or "è" in s or "ç" in s or "ã" in s or "ï" in s)
    if suspicious:
        try:
            return s.encode("latin1").decode("utf-8")
        except Exception:
            return s
    return s


def strip_html(html_text: Any) -> str:
    s = "" if html_text is None else str(html_text)
    if "<" in s and ">" in s:
        return BeautifulSoup(s, "lxml").get_text(" ", strip=True)
    return s


def unique_keep_order(items: Iterable[str]) -> List[str]:
    seen = set()
    out: List[str] = []
    for item in items:
        item = clean_value(item)
        if item and item not in seen:
            seen.add(item)
            out.append(item)
    return out


def sleep_jitter(seconds: float) -> None:
    if seconds <= 0:
        return
    time.sleep(seconds + random.uniform(0, min(0.8, seconds * 0.5)))


def build_headers() -> Dict[str, str]:
    return {
        "Accept": "*/*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Origin": BASE_URL,
        "Referer": f"{BASE_URL}/elasticsearch/index.jsp?c1=jyxx&c2=jyxxzfcg&c3=jyxxcggg&c4=&e=&ext8=&inDates=&channelId=126&q=",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"
        ),
        "X-Requested-With": "XMLHttpRequest",
    }


class LegacyTLSAdapter(HTTPAdapter):
    """给少数老旧/特殊政务站点使用的 TLS 兼容适配器。"""

    def _make_context(self) -> ssl.SSLContext:
        ctx = ssl.create_default_context()
        # 部分站点在 Python/conda 的 OpenSSL 握手时会触发 BAD_ECPOINT。
        # 降低 OpenSSL 安全等级通常比 verify=False 更有效；verify=False 对这个错误一般没用。
        try:
            ctx.set_ciphers("DEFAULT@SECLEVEL=1")
        except Exception:
            pass
        return ctx

    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        pool_kwargs["ssl_context"] = self._make_context()
        return super().init_poolmanager(connections, maxsize, block=block, **pool_kwargs)

    def proxy_manager_for(self, proxy, **proxy_kwargs):
        proxy_kwargs["ssl_context"] = self._make_context()
        return super().proxy_manager_for(proxy, **proxy_kwargs)


class CompatSession:
    """requests / curl_cffi 兼容会话。

    http_backend 可选：
    - auto：优先使用 curl_cffi；没安装时退回 requests
    - curl_cffi：强制使用 curl_cffi，模拟 Chrome TLS 指纹
    - requests：使用 requests，并挂载 LegacyTLSAdapter
    """

    def __init__(self, cfg: Dict[str, Any]):
        self.headers = build_headers()
        backend = str(cfg.get("http_backend", "auto") or "auto").lower()
        self.impl = "requests"
        self._curl_impersonate = str(cfg.get("curl_impersonate", "chrome") or "chrome")

        if backend in {"auto", "curl_cffi", "curl"}:
            try:
                from curl_cffi import requests as curl_requests  # type: ignore
                self.session = curl_requests.Session()
                self.impl = "curl_cffi"
            except Exception as e:
                if backend in {"curl_cffi", "curl"}:
                    raise RuntimeError(
                        "配置要求使用 curl_cffi，但当前环境未安装。请执行：pip install curl_cffi"
                    ) from e
                self.session = requests.Session()
        else:
            self.session = requests.Session()

        try:
            self.session.headers.update(self.headers)
        except Exception:
            pass

        if self.impl == "requests":
            try:
                adapter = LegacyTLSAdapter(max_retries=0)
                self.session.mount("https://", adapter)
            except Exception:
                pass

        logging.info("HTTP backend：%s", self.impl)

    def post(self, url: str, **kwargs):
        if self.impl == "curl_cffi":
            kwargs.setdefault("impersonate", self._curl_impersonate)
        return self.session.post(url, **kwargs)

    def get(self, url: str, **kwargs):
        if self.impl == "curl_cffi":
            kwargs.setdefault("impersonate", self._curl_impersonate)
        return self.session.get(url, **kwargs)


def make_session(cfg: Optional[Dict[str, Any]] = None):
    return CompatSession(cfg or {})


# ---------- 断点与输出 ----------

def load_visited(path: Path) -> set:
    if not path.exists():
        return set()
    with path.open("r", encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def save_visited(path: Path, url: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(url + "\n")


def append_jsonl(path: Path, record: NoticeRecord) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(asdict(record), ensure_ascii=False) + "\n")


def read_jsonl_records(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    rows = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except Exception:
                continue
    return rows


def load_progress(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {"completed_tasks": [], "current": None, "updated_at": ""}
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("progress_state 不是 JSON object")
        data.setdefault("completed_tasks", [])
        data.setdefault("current", None)
        data.setdefault("updated_at", "")
        return data
    except Exception as e:
        backup = path.with_suffix(path.suffix + f".broken_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        try:
            path.rename(backup)
            logging.warning("断点文件损坏，已备份为：%s，错误：%s", backup, e)
        except Exception:
            logging.warning("断点文件读取失败：%s", e)
        return {"completed_tasks": [], "current": None, "updated_at": ""}


def save_progress(path: Path, progress: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    progress["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    content = json.dumps(progress, ensure_ascii=False, indent=2)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        f.write(content)

    last_error = None
    for i in range(10):
        try:
            tmp_path.replace(path)
            return
        except PermissionError as e:
            last_error = e
            time.sleep(0.2 + i * 0.3)
        except Exception as e:
            last_error = e
            time.sleep(0.2 + i * 0.3)

    try:
        with path.open("w", encoding="utf-8") as f:
            f.write(content)
        if tmp_path.exists():
            tmp_path.unlink()
    except Exception:
        recovery = path.with_name(f"{path.stem}_recovery_{datetime.now().strftime('%Y%m%d_%H%M%S')}{path.suffix}")
        with recovery.open("w", encoding="utf-8") as f:
            f.write(content)
        logging.error("断点文件保存失败，已另存为恢复文件：%s，原错误：%s", recovery, last_error)


def task_key(keyword: str, notice_type: str) -> str:
    return f"{keyword}||{notice_type}"


def get_resume_page(progress: Dict[str, Any], key: str) -> int:
    cur = progress.get("current") or {}
    if cur.get("task_key") == key:
        try:
            return max(1, int(cur.get("next_page", 1)))
        except Exception:
            return 1
    return 1


def set_current_progress(path: Path, progress: Dict[str, Any], key: str, next_page: int, meta: Dict[str, Any]) -> None:
    progress["current"] = {"task_key": key, "next_page": int(next_page), **meta}
    save_progress(path, progress)


def mark_task_completed(path: Path, progress: Dict[str, Any], key: str) -> None:
    completed = list(progress.get("completed_tasks", []))
    if key not in completed:
        completed.append(key)
    progress["completed_tasks"] = completed
    cur = progress.get("current") or {}
    if cur.get("task_key") == key:
        progress["current"] = None
    save_progress(path, progress)


# ---------- 搜索接口 ----------

def request_search_page(
    session: Session,
    cfg: Dict[str, Any],
    keyword: str,
    channel_third: str,
    page: int,
) -> Tuple[List[Dict[str, Any]], int, int]:
    start_time = cfg.get("start_date") or "2025-01-01"
    end_time = cfg.get("end_date") or today_str()
    payload = {
        "searchword": keyword,
        "scope": cfg.get("scope", "title"),
        "channel_first": "jyxx",
        "channel_second": "jyxxzfcg",
        "channel_third": channel_third,
        "channel_fourth": "",
        "legislationType": "",
        "ext": "all",
        "ext8": "",
        "starttime": start_time,
        "endtime": end_time,
        "sort": "",
        "page": str(page),
        "size": str(cfg.get("page_size", "") or ""),
    }

    retry_times = int(cfg.get("retry_times", 4) or 4)
    timeout = int(cfg.get("timeout", 30) or 30)
    last_error: Optional[Exception] = None

    for attempt in range(1, retry_times + 1):
        try:
            resp = session.post(SEARCH_URL, data=payload, timeout=timeout)
            resp.encoding = "utf-8"
            resp.raise_for_status()
            outer = resp.json()
            total = int(outer.get("total") or 0)
            size = int(outer.get("size") or 0)
            raw_result = outer.get("result") or []
            if isinstance(raw_result, str):
                rows = json.loads(raw_result) if raw_result.strip() else []
            elif isinstance(raw_result, list):
                rows = raw_result
            else:
                rows = []
            if not size:
                size = len(rows) or 10
            return rows, total, size
        except Exception as e:
            last_error = e
            wait = min(60, 2 ** attempt)
            logging.warning("列表接口请求失败，第 %s/%s 次，等待 %s 秒：%s", attempt, retry_times, wait, e)
            time.sleep(wait)

    raise RuntimeError(f"列表接口请求失败：{last_error}")


def normalize_search_item(item: Dict[str, Any]) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for k, v in item.items():
        if isinstance(v, str):
            out[k] = clean_value(v)
        else:
            out[k] = v
    return out


# ---------- 详情页解析 ----------

def get_soup(session: Session, url: str, cfg: Dict[str, Any]) -> BeautifulSoup:
    retry_times = int(cfg.get("retry_times", 4) or 4)
    timeout = int(cfg.get("timeout", 30) or 30)
    headers = dict(build_headers())
    headers["Referer"] = f"{BASE_URL}/elasticsearch/index.jsp?c1=jyxx&c2=jyxxzfcg&c3=jyxxcggg&c4=&e=&ext8=&inDates=&channelId=126&q="
    last_error: Optional[Exception] = None

    for attempt in range(1, retry_times + 1):
        try:
            resp = session.get(url, headers=headers, timeout=timeout)
            resp.encoding = "utf-8"
            if resp.status_code in {403, 408, 429, 500, 502, 503, 520, 521, 522, 523, 524}:
                raise RuntimeError(f"HTTP {resp.status_code}")
            resp.raise_for_status()
            html = resp.text or ""
            suspicious = ["验证码", "请登录", "访问过于频繁", "安全验证"]
            if any(w in html for w in suspicious) and "newsCon" not in html:
                raise RuntimeError("详情页疑似风控/异常页面")
            return BeautifulSoup(html, "lxml")
        except Exception as e:
            last_error = e
            wait = min(60, 2 ** attempt)
            logging.warning("详情页请求失败，第 %s/%s 次，等待 %s 秒：%s | %s", attempt, retry_times, wait, url, e)
            time.sleep(wait)

    raise RuntimeError(f"详情页请求失败：{url} | {last_error}")


def meta_content(soup: BeautifulSoup, name: str) -> str:
    el = soup.find("meta", attrs={"http-equiv": name}) or soup.find("meta", attrs={"name": name})
    return clean_value(el.get("content", "")) if el else ""


def extract_detail_text(soup: BeautifulSoup) -> str:
    container = soup.select_one(".newsCon") or soup.select_one(".vT_detail_content") or soup.select_one(".div-article2") or soup.body
    if not container:
        return normalize_space(soup.get_text("\n", strip=True))
    return normalize_space(container.get_text("\n", strip=True))


def extract_title(soup: BeautifulSoup, fallback: str = "") -> str:
    title = meta_content(soup, "ArticleTitle")
    if title:
        return title
    div_title = soup.select_one(".div-title")
    if div_title:
        # 移除 p 里的交易项目编号，只保留标题主文本
        clone = BeautifulSoup(str(div_title), "lxml")
        for p in clone.select("p"):
            p.decompose()
        txt = clean_value(clone.get_text(" ", strip=True))
        if txt:
            return txt
    return clean_value(fallback)


def extract_publish_time(soup: BeautifulSoup, fallback: str = "") -> str:
    pub = meta_content(soup, "PubDate")
    if pub:
        return pub[:10]
    title2 = soup.select_one(".div-title2")
    if title2:
        m = re.search(r"发布时间\s*[：:]\s*(\d{4}-\d{2}-\d{2})", title2.get_text(" ", strip=True))
        if m:
            return m.group(1)
    return clean_value(fallback)


def extract_transaction_project_code(soup: BeautifulSoup, fallback: str = "") -> str:
    div_title = soup.select_one(".div-title")
    if div_title:
        m = re.search(r"交易项目编号\s*[：:]\s*([A-Za-z0-9_\-]+)", div_title.get_text(" ", strip=True))
        if m:
            return clean_value(m.group(1))
    text = soup.get_text("\n", strip=True)
    m = re.search(r"交易项目编号\s*[：:]\s*([A-Za-z0-9_\-]+)", text)
    if m:
        return clean_value(m.group(1))
    return clean_value(fallback)


def regex_label_extract(text: str, aliases: List[str], max_len: int = 180) -> str:
    # 行式优先，避免跨太多内容
    lines = [normalize_space(x) for x in text.split("\n") if normalize_space(x)]
    for line in lines:
        for alias in aliases:
            # 兼容：一、项目编号：xxx；<strong>一、项目编号：xxx</strong>
            pattern = r"(?:^[一二三四五六七八九十]+、)?" + re.escape(alias).replace(r"\ ", r"\s*") + r"\s*[：:]\s*(.+)$"
            m = re.search(pattern, line)
            if m:
                val = clean_value(m.group(1))
                if val and len(val) <= max_len:
                    # 排除“供应商名称、地址及金额”这种说明行
                    if "、地址" in val and "金额" in val:
                        continue
                    return val

    # 全文兜底
    for alias in aliases:
        pattern = r"(?:[一二三四五六七八九十]+、)?" + re.escape(alias).replace(r"\ ", r"\s*") + r"\s*[：:]\s*([^\n。；;]{1," + str(max_len) + r"})"
        m = re.search(pattern, text)
        if m:
            val = clean_value(m.group(1))
            if val and not ("、地址" in val and "金额" in val):
                return val
    return ""


def section_name_extract(text: str, section_keyword: str) -> str:
    lines = [normalize_space(x) for x in text.split("\n") if normalize_space(x)]
    for i, line in enumerate(lines):
        if section_keyword in line:
            for nxt in lines[i + 1:i + 8]:
                m = re.search(r"名\s*称\s*[：:]\s*(.+)$", nxt)
                if m:
                    return clean_value(m.group(1))
    return ""


def section_value_extract(text: str, section_keyword: str, label_regex: str) -> str:
    lines = [normalize_space(x) for x in text.split("\n") if normalize_space(x)]
    for i, line in enumerate(lines):
        if section_keyword in line:
            for nxt in lines[i + 1:i + 10]:
                m = re.search(label_regex + r"\s*[：:]\s*(.+)$", nxt)
                if m:
                    return clean_value(m.group(1))
    return ""


def extract_purchase_type(text: str, title: str) -> str:
    val = regex_label_extract(text, FIELD_ALIASES["purchase_type"], max_len=80)
    if val:
        return val
    # 兜底只看标题，避免成交结果公告中的“单一来源采购人员名单”等固定表述被误判为采购方式。
    found = [kw for kw in PURCHASE_TYPE_KEYWORDS if kw in (title or "")]
    return "；".join(unique_keep_order(found[:3]))


def extract_deadline(text: str) -> str:
    # 采购公告里常见：截止时间：2025-09-22 09:30（北京时间）
    m = re.search(r"截止时间\s*[：:]\s*([^\n]{1,80})", text)
    if m:
        return clean_value(m.group(1))
    m = re.search(r"(?:提交投标文件|提交响应文件).*?(\d{4}-\d{2}-\d{2}\s*\d{1,2}:\d{2})", text)
    if m:
        return clean_value(m.group(1))
    return ""


def extract_from_tables(soup: BeautifulSoup) -> Dict[str, str]:
    """从成交结果公告表格里提取供应商、地址、信用代码、中标金额、备注等。"""
    container = soup.select_one(".newsCon") or soup
    result = {
        "winner_unit": "",
        "winner_amount": "",
        "supplier_address": "",
        "unified_credit_code": "",
        "review_score_or_remark": "",
    }
    for table in container.select("table"):
        rows: List[List[str]] = []
        for tr in table.select("tr"):
            cells = [clean_value(td.get_text(" ", strip=True)) for td in tr.find_all(["th", "td"])]
            cells = [c for c in cells if c != ""]
            if cells:
                rows.append(cells)
        if len(rows) < 2:
            continue
        header = rows[0]
        data_rows = rows[1:]
        header_compact = [h.replace(" ", "") for h in header]

        def find_idx(names: List[str]) -> Optional[int]:
            for i, h in enumerate(header_compact):
                if any(name.replace(" ", "") in h for name in names):
                    return i
            return None

        # 优先第一张“供应商名称/中标金额”表，不要误用主要标的信息表
        idx_supplier = find_idx(["供应商名称", "中标成交供应商名称", "中标供应商", "成交供应商"])
        idx_amount = find_idx(["中标金额", "成交金额", "中标成交金额", "中标（成交）金额"])
        if idx_supplier is None and idx_amount is None:
            continue

        idx_addr = find_idx(["供应商地址", "地址"])
        idx_code = find_idx(["统一信用代码", "统一社会信用代码"])
        idx_remark = find_idx(["备注", "中标成交备注信息", "评审总得分"])

        suppliers, amounts, addrs, codes, remarks = [], [], [], [], []
        for row in data_rows:
            if idx_supplier is not None and idx_supplier < len(row):
                suppliers.append(row[idx_supplier])
            if idx_amount is not None and idx_amount < len(row):
                amounts.append(row[idx_amount])
            if idx_addr is not None and idx_addr < len(row):
                addrs.append(row[idx_addr])
            if idx_code is not None and idx_code < len(row):
                codes.append(row[idx_code])
            if idx_remark is not None and idx_remark < len(row):
                remarks.append(row[idx_remark])
        if suppliers:
            result["winner_unit"] = "；".join(unique_keep_order(suppliers))
        if amounts:
            result["winner_amount"] = "；".join(unique_keep_order(amounts))
        if addrs:
            result["supplier_address"] = "；".join(unique_keep_order(addrs))
        if codes:
            result["unified_credit_code"] = "；".join(unique_keep_order(codes))
        if remarks:
            result["review_score_or_remark"] = "；".join(unique_keep_order(remarks))
        if result["winner_unit"] or result["winner_amount"]:
            return result
    return result


def extract_summary(text: str, max_len: int = 500) -> str:
    return normalize_space(text).replace("\n", " ")[:max_len]


def parse_detail_page(session: Session, cfg: Dict[str, Any], url: str, base: NoticeRecord) -> NoticeRecord:
    soup = get_soup(session, url, cfg)
    text = extract_detail_text(soup)
    title = extract_title(soup, base.raw_title)

    base.raw_title = title or base.raw_title
    base.publish_time = extract_publish_time(soup, base.publish_time)
    base.transaction_project_code = extract_transaction_project_code(soup, base.transaction_project_code)
    base.notice_type = meta_content(soup, "ColumnName") or meta_content(soup, "ColumnType") or base.notice_type

    base.project_no = regex_label_extract(text, FIELD_ALIASES["project_no"], max_len=120) or base.project_no
    base.project_name = regex_label_extract(text, FIELD_ALIASES["project_name"], max_len=200) or base.project_name or title
    base.purchase_type = extract_purchase_type(text, title) or base.purchase_type
    base.budget_amount = regex_label_extract(text, FIELD_ALIASES["budget_amount"], max_len=100) or base.budget_amount
    base.max_price = regex_label_extract(text, FIELD_ALIASES["max_price"], max_len=100) or base.max_price

    base.tender_unit = section_name_extract(text, "采购人信息") or regex_label_extract(text, FIELD_ALIASES["tender_unit"], max_len=180) or base.tender_unit
    base.agency = section_name_extract(text, "采购代理机构信息") or regex_label_extract(text, FIELD_ALIASES["agency"], max_len=180) or base.agency

    # 成交结果字段
    base.total_winner_amount = regex_label_extract(text, FIELD_ALIASES["total_winner_amount"], max_len=120) or base.total_winner_amount
    base.winner_unit = regex_label_extract(text, FIELD_ALIASES["winner_unit"], max_len=200) or base.winner_unit
    base.winner_amount = regex_label_extract(text, FIELD_ALIASES["winner_amount"], max_len=120) or base.winner_amount
    table_vals = extract_from_tables(soup)
    if table_vals.get("winner_unit"):
        base.winner_unit = table_vals["winner_unit"]
    if table_vals.get("winner_amount"):
        base.winner_amount = table_vals["winner_amount"]
    base.supplier_address = table_vals.get("supplier_address") or section_value_extract(text, "中标（成交）信息", r"(?:中标成交供应商地址|供应商地址|地址)") or base.supplier_address
    base.unified_credit_code = table_vals.get("unified_credit_code") or base.unified_credit_code
    base.review_score_or_remark = table_vals.get("review_score_or_remark") or base.review_score_or_remark

    # 采购公告截止时间优先用搜索列表 noticeEndTime，详情兜底
    if not base.notice_end_time:
        base.notice_end_time = extract_deadline(text)

    base.raw_text_summary = extract_summary(text)
    base.parse_status = "ok"
    return base


# ---------- Excel 导出与项目合并 ----------

def records_to_dataframe(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    col_order = [
        "search_keyword", "notice_type", "channel_third",
        "project_name", "transaction_project_code", "project_no",
        "tender_unit", "agency", "purchase_type", "budget_amount", "max_price",
        "winner_unit", "winner_amount", "total_winner_amount", "supplier_address", "unified_credit_code", "review_score_or_remark",
        "publish_time", "notice_end_time", "source_region", "source", "detail_url", "raw_title", "raw_text_summary",
        "parse_status", "error",
    ]
    for c in col_order:
        if c not in df.columns:
            df[c] = ""
    return df[col_order]


def chinese_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={
        "search_keyword": "检索关键词",
        "notice_type": "公告类型",
        "channel_third": "栏目代码",
        "project_name": "项目名称",
        "transaction_project_code": "交易项目编号",
        "project_no": "项目编号",
        "tender_unit": "采购人/招标单位",
        "agency": "采购代理机构",
        "purchase_type": "采购方式",
        "budget_amount": "预算金额",
        "max_price": "最高限价",
        "winner_unit": "中标/成交供应商",
        "winner_amount": "中标/成交金额",
        "total_winner_amount": "总中标成交金额",
        "supplier_address": "供应商地址",
        "unified_credit_code": "统一信用代码",
        "review_score_or_remark": "评审得分/备注",
        "publish_time": "发布时间",
        "notice_end_time": "投标/响应截止时间",
        "source_region": "来源地区",
        "source": "信息来源",
        "detail_url": "详情链接",
        "raw_title": "原始标题",
        "raw_text_summary": "正文摘要",
        "parse_status": "解析状态",
        "error": "错误信息",
    })


def first_nonempty(values: Iterable[Any]) -> str:
    for v in values:
        s = clean_value(v)
        if s:
            return s
    return ""


def join_unique(values: Iterable[Any]) -> str:
    return "；".join(unique_keep_order([str(v) for v in values if clean_value(v)]))


def make_merged_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    df = df.copy()
    df["merge_key"] = df["transaction_project_code"].where(df["transaction_project_code"].astype(str).str.len() > 0, df["project_name"])

    merged_rows: List[Dict[str, Any]] = []
    for _, g in df.groupby("merge_key", dropna=False):
        # 采购公告优先提供采购信息；成交结果公告优先提供成交信息
        purch = g[g["notice_type"].astype(str).str.contains("采购公告|招标公告", regex=True, na=False)]
        result = g[g["notice_type"].astype(str).str.contains("成交|中标|结果", regex=True, na=False)]

        def pick(col: str, prefer: Optional[pd.DataFrame] = None) -> str:
            if prefer is not None and not prefer.empty:
                v = first_nonempty(prefer[col].tolist())
                if v:
                    return v
            return first_nonempty(g[col].tolist())

        row = {
            "search_keyword": join_unique(g["search_keyword"].tolist()),
            "notice_type": join_unique(g["notice_type"].tolist()),
            "project_name": pick("project_name", purch if not purch.empty else result),
            "transaction_project_code": pick("transaction_project_code"),
            "project_no": pick("project_no", purch if not purch.empty else result),
            "tender_unit": pick("tender_unit", purch if not purch.empty else result),
            "agency": pick("agency", purch if not purch.empty else result),
            "purchase_type": pick("purchase_type", purch if not purch.empty else result),
            "budget_amount": pick("budget_amount", purch),
            "max_price": pick("max_price", purch),
            "winner_unit": pick("winner_unit", result),
            "winner_amount": pick("winner_amount", result),
            "total_winner_amount": pick("total_winner_amount", result),
            "supplier_address": pick("supplier_address", result),
            "unified_credit_code": pick("unified_credit_code", result),
            "review_score_or_remark": pick("review_score_or_remark", result),
            "publish_time": join_unique(g["publish_time"].tolist()),
            "notice_end_time": pick("notice_end_time", purch),
            "source_region": join_unique(g["source_region"].tolist()),
            "source": join_unique(g["source"].tolist()),
            "detail_url": join_unique(g["detail_url"].tolist()),
            "raw_title": join_unique(g["raw_title"].tolist()),
            "parse_status": join_unique(g["parse_status"].tolist()),
            "error": join_unique(g["error"].tolist()),
        }
        merged_rows.append(row)
    return pd.DataFrame(merged_rows)


def export_excel(cfg: Dict[str, Any]) -> None:
    jsonl_path = BASE_DIR / cfg.get("checkpoint_jsonl", DEFAULT_CONFIG["checkpoint_jsonl"])
    detail_path = BASE_DIR / cfg.get("output_excel_detail", DEFAULT_CONFIG["output_excel_detail"])
    merged_path = BASE_DIR / cfg.get("output_excel_merged", DEFAULT_CONFIG["output_excel_merged"])
    rows = read_jsonl_records(jsonl_path)
    if not rows:
        logging.warning("没有可导出的记录。")
        return

    # 按 URL 去重，保留最后一次解析结果
    by_url: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        key = row.get("detail_url") or f"row_{len(by_url)}"
        by_url[key] = row

    df = records_to_dataframe(list(by_url.values()))
    detail_path.parent.mkdir(parents=True, exist_ok=True)
    chinese_columns(df).to_excel(detail_path, index=False)

    merged = make_merged_dataframe(df)
    # merged 缺少部分列时补齐
    for c in [c for c in records_to_dataframe([]).columns if c not in merged.columns and c != "channel_third"]:
        merged[c] = ""
    merged_cols = [
        "search_keyword", "notice_type", "project_name", "transaction_project_code", "project_no",
        "tender_unit", "agency", "purchase_type", "budget_amount", "max_price",
        "winner_unit", "winner_amount", "total_winner_amount", "supplier_address", "unified_credit_code", "review_score_or_remark",
        "publish_time", "notice_end_time", "source_region", "source", "detail_url", "raw_title", "parse_status", "error",
    ]
    chinese_columns(merged[merged_cols]).to_excel(merged_path, index=False)

    logging.info("Excel 明细已导出：%s，共 %d 条。", detail_path, len(df))
    logging.info("Excel 合并表已导出：%s，共 %d 条。", merged_path, len(merged))


# ---------- 主流程 ----------

def build_record_from_search_item(keyword: str, notice_name: str, channel_third: str, item: Dict[str, Any]) -> NoticeRecord:
    item = normalize_search_item(item)
    link = item.get("link", "")
    detail_url = urljoin(BASE_URL, link)
    return NoticeRecord(
        search_keyword=keyword,
        notice_type=notice_name,
        channel_third=channel_third,
        project_name=item.get("title", ""),
        transaction_project_code=item.get("projectCode", ""),
        publish_time=item.get("releaseDate", ""),
        notice_end_time=item.get("noticeEndTime", ""),
        source_region=item.get("region", ""),
        source=item.get("source", ""),
        detail_url=detail_url,
        raw_title=item.get("title", ""),
        raw_text_summary=extract_summary(item.get("content", "")),
        parse_status="list_only",
    )


def crawl(cfg: Dict[str, Any]) -> None:
    setup_logging(cfg)
    logging.info("开始爬取，北京公共资源交易服务平台，搜索范围：%s", cfg.get("scope", "title"))
    logging.info("时间范围：%s 至 %s", cfg.get("start_date"), cfg.get("end_date") or today_str())

    jsonl_path = BASE_DIR / cfg.get("checkpoint_jsonl", DEFAULT_CONFIG["checkpoint_jsonl"])
    visited_path = BASE_DIR / cfg.get("visited_urls_file", DEFAULT_CONFIG["visited_urls_file"])
    progress_path = BASE_DIR / cfg.get("progress_state_file", DEFAULT_CONFIG["progress_state_file"])

    visited = load_visited(visited_path)
    progress = load_progress(progress_path)
    session = make_session(cfg)

    keywords = cfg.get("keywords") or ["营商环境"]
    notice_types: Dict[str, str] = cfg.get("notice_types") or DEFAULT_CONFIG["notice_types"]
    max_pages = int(cfg.get("max_pages_per_query", 9999) or 9999)
    request_delay = float(cfg.get("request_delay", 1.0) or 0)
    detail_delay = float(cfg.get("detail_delay", 0.8) or 0)
    auto_export_every = int(cfg.get("auto_export_every_records", 20) or 0)

    saved_since_export = 0
    total_tasks = len(keywords) * len(notice_types)
    task_idx = 0

    for keyword in keywords:
        for notice_name, channel_third in notice_types.items():
            task_idx += 1
            key = task_key(keyword, notice_name)
            if key in set(progress.get("completed_tasks", [])):
                logging.info("[%d/%d] 跳过已完成任务：%s", task_idx, total_tasks, key)
                continue

            page = get_resume_page(progress, key)
            logging.info("[%d/%d] 开始任务：关键词=%s，公告类型=%s，起始页=%s", task_idx, total_tasks, keyword, notice_name, page)

            total_pages = max_pages
            while page <= total_pages and page <= max_pages:
                set_current_progress(progress_path, progress, key, page, {
                    "keyword": keyword,
                    "notice_type": notice_name,
                    "channel_third": channel_third,
                })

                rows, total_count, page_size = request_search_page(session, cfg, keyword, channel_third, page)
                if total_count == 0 or not rows:
                    logging.info("无结果或当前页为空：关键词=%s，公告类型=%s，page=%s", keyword, notice_name, page)
                    break
                total_pages = min(max_pages, max(1, math.ceil(total_count / max(1, page_size))))
                logging.info("列表页：关键词=%s 类型=%s 第 %s/%s 页，本页 %d 条，总计 %d 条", keyword, notice_name, page, total_pages, len(rows), total_count)

                for item in rows:
                    record = build_record_from_search_item(keyword, notice_name, channel_third, item)
                    if not record.detail_url:
                        continue
                    if record.detail_url in visited:
                        continue
                    try:
                        record = parse_detail_page(session, cfg, record.detail_url, record)
                    except Exception as e:
                        record.parse_status = "detail_failed"
                        record.error = str(e)
                        logging.warning("详情页解析失败，保存列表基础字段后继续：%s | %s", record.detail_url, e)

                    append_jsonl(jsonl_path, record)
                    save_visited(visited_path, record.detail_url)
                    visited.add(record.detail_url)
                    saved_since_export += 1
                    logging.info("保存：%s | %s | %s", record.notice_type, record.publish_time, record.raw_title[:80])
                    sleep_jitter(detail_delay)

                    if auto_export_every > 0 and saved_since_export >= auto_export_every:
                        export_excel(cfg)
                        saved_since_export = 0

                page += 1
                set_current_progress(progress_path, progress, key, page, {
                    "keyword": keyword,
                    "notice_type": notice_name,
                    "channel_third": channel_third,
                })
                sleep_jitter(request_delay)

            mark_task_completed(progress_path, progress, key)
            logging.info("任务完成：%s", key)

    export_excel(cfg)
    logging.info("全部完成。")


def main() -> None:
    cfg = load_config()
    try:
        crawl(cfg)
    except KeyboardInterrupt:
        print("\n用户中断，正在导出已抓取数据...")
        try:
            export_excel(cfg)
        except Exception:
            pass
    except Exception as e:
        logging.exception("运行失败：%s", e)
        print("\n程序异常退出，尝试导出已抓取数据...")
        try:
            export_excel(cfg)
        except Exception:
            pass
        raise


if __name__ == "__main__":
    main()
