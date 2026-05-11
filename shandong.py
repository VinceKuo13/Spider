# -*- coding: utf-8 -*-
"""
山东省公共资源交易中心 / 山东省公共资源交易平台 标讯爬虫 v10

修复点：增强山东详情页字段解析，兼容“项目编号 ：”“采购人 信息 名 称：”“名 称：”等带空格标签。

适用站点：
https://ggzyjyzx.shandong.gov.cn/jsearchfront/search.do

已按用户提供的接口重构：
1. 列表接口：/jsearchfront/interfaces/cateSearch.do
2. 分类：cateid=16730，交易信息
3. 日期参数：begin=YYYYMMDD, end=YYYYMMDD
4. 分页参数：p=页码，pg=每页条数
5. 详情页：/art/YYYY/M/D/art_xxx_xxx.html
6. 使用高级检索 pq=完整关键词，避免“营商环境”被拆成“环境”；默认不跳过 visited，避免调试时空导出。

运行：
    pip install requests beautifulsoup4 pandas openpyxl
    python spider_shandong.py

输出：
    outputs_shandong/
      ├─ shandong_records.jsonl
      ├─ 山东公共资源交易中心_标讯明细.xlsx
      ├─ shandong_visited.txt
      ├─ progress_state.json
      └─ raw/
"""

from __future__ import annotations

import json
import logging
import math
import random
import re
import time
from dataclasses import dataclass, asdict
from datetime import date, datetime
from html import unescape
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin, urlencode, quote_plus

import pandas as pd
import requests
from bs4 import BeautifulSoup


# =========================
# 可修改配置
# =========================
KEYWORDS = ["营商环境"]
START_DATE = "2025-01-01"
END_DATE = date.today().strftime("%Y-%m-%d")

# 高级检索成功 URL 中 cateid=15503。
# 如果你只想限定“交易信息”，可改回 16730；但当前默认与手动高级检索页面保持一致。
CATE_ID = "16730"

# 山东站普通检索会把“营商环境”拆成“营商”和“环境”。
# 高级检索“包含以下完整关键词”对应 pq 参数，因此接口层面使用 pq=关键词。
SEARCH_POS = "title,content,_default_search"
USE_PQ_PHRASE = True
USE_EQ_PHRASE = False
# 接口返回后仍做本地标题严格过滤：只保留标题中连续出现完整关键词的记录。
LOCAL_TITLE_FILTER = False
STRICT_TITLE_PHRASE_FILTER = False

# 调试阶段建议关闭 visited 跳过：此前失败/中断过的详情链接可能已经写入 visited，
# 会导致“列表解析到 15 条，但没有可导出记录”。稳定运行后可改为 True。
SKIP_VISITED = False

PAGE_SIZE = 20
MAX_PAGES = 0  # 0 表示按接口 total 自动爬完；测试时可改成 1/2
REQUEST_TIMEOUT = 60
RETRIES = 5
# 详情页经常很慢。为了避免第一条详情卡住几十秒导致没有任何输出，
# 详情页单独使用较短超时；失败会保留列表信息并继续后续记录。
DETAIL_TIMEOUT = 20
DETAIL_RETRIES = 2
SLEEP_LIST = (1.0, 2.0)
SLEEP_DETAIL = (1.0, 2.0)

BASE_URL = "https://ggzyjyzx.shandong.gov.cn"

# 读取同目录 config.yaml 中的 keywords/start_date/end_date。
from common_config import apply_common_config
apply_common_config(globals())


SEARCH_PAGE = BASE_URL + "/jsearchfront/search.do"
CATE_SEARCH_URL = BASE_URL + "/jsearchfront/interfaces/cateSearch.do"
WEBSITE_ID = "370000000000110"
TPL = "1164"

OUT_DIR = Path("outputs_shandong")
RAW_DIR = OUT_DIR / "raw"
JSONL_PATH = OUT_DIR / "shandong_records.jsonl"
EXCEL_PATH = OUT_DIR / "山东公共资源交易中心_标讯明细.xlsx"
VISITED_PATH = OUT_DIR / "shandong_visited.txt"
PROGRESS_PATH = OUT_DIR / "progress_state.json"
LOG_PATH = OUT_DIR / "shandong_spider.log"

# 本次运行内存记录。即使 jsonl 写入或读取路径出现问题，也能直接导出 Excel。
RUN_RECORDS: List[Dict[str, Any]] = []


# =========================
# 日志与工具函数
# =========================
def setup_logging() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def clean_html_text(value: Any) -> str:
    if value is None:
        return ""
    s = str(value)
    s = unescape(s)
    s = re.sub(r"<\s*br\s*/?\s*>", "\n", s, flags=re.I)
    s = re.sub(r"</\s*p\s*>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    s = s.replace("\xa0", " ").replace("&nbsp;", " ")
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s.strip()


def normalize_space(s: Any) -> str:
    s = clean_html_text(s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def compact_for_match(s: Any) -> str:
    """用于标题短语匹配：去掉 HTML、空白、标点和高亮标签影响。"""
    s = normalize_space(s)
    s = re.sub(r"[\s\u3000]+", "", s)
    s = re.sub(r'[，。；;：:、,.!?！？（）()【】\[\]《》<>"“”‘’/\\|_-]+', "", s)
    return s.lower()


def is_title_phrase_match(title: str, keyword: str) -> bool:
    """严格判断标题是否包含完整关键词。

    例如关键词为“营商环境”时：
    - “优化营商环境项目” => 保留
    - “生态环境监测项目” => 过滤
    - “营商 服务 环境” => 过滤，因为不是连续短语
    """
    if not keyword:
        return True
    if not STRICT_TITLE_PHRASE_FILTER:
        return keyword in normalize_space(title)
    return compact_for_match(keyword) in compact_for_match(title)


def ymd_to_compact(s: str) -> str:
    return s.replace("-", "")


def safe_filename(s: str, max_len: int = 80) -> str:
    s = normalize_space(s)
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    return s[:max_len] or "untitled"


def now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_date_any(s: str) -> str:
    if not s:
        return ""
    s = normalize_space(s)
    m = re.search(r"(20\d{2})[-年/.](\d{1,2})[-月/.](\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r"(20\d{6})", s)
    if m:
        x = m.group(1)
        return f"{x[:4]}-{x[4:6]}-{x[6:]}"
    return s


def load_visited() -> set[str]:
    if not VISITED_PATH.exists():
        return set()
    return {line.strip() for line in VISITED_PATH.read_text(encoding="utf-8").splitlines() if line.strip()}


def add_visited(url: str) -> None:
    with VISITED_PATH.open("a", encoding="utf-8") as f:
        f.write(url + "\n")


def append_jsonl(record: Dict[str, Any]) -> None:
    """同时写 JSONL 和本次运行内存，保证最后 Excel 不会空。"""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_RECORDS.append(dict(record))
    with JSONL_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")
        f.flush()


def save_progress(state: Dict[str, Any]) -> None:
    PROGRESS_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def load_records() -> List[Dict[str, Any]]:
    if not JSONL_PATH.exists():
        return []
    records = []
    for line in JSONL_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            records.append(json.loads(line))
        except Exception:
            continue
    return records


def export_excel() -> None:
    records = load_records()
    # v9：优先合并本次运行内存记录，避免 JSONL 路径/读取问题造成“没有可导出记录”。
    if RUN_RECORDS:
        seen = set()
        merged = []
        for r in records + RUN_RECORDS:
            key = r.get("详情链接") or json.dumps(r, ensure_ascii=False, sort_keys=True)
            if key in seen:
                continue
            seen.add(key)
            merged.append(r)
        records = merged
    if not records:
        logging.warning("没有可导出的记录。请检查 outputs_shandong/raw 中的列表响应，或把 raw 文件发我。")
        return
    df = pd.DataFrame(records)
    preferred_cols = [
        "省份", "检索关键词", "公告类型", "标题", "项目名称", "项目编号", "采购方式",
        "采购人/采购单位", "采购代理机构", "预算金额", "最高限价",
        "中标单位/成交供应商", "供应商地址", "中标金额/成交金额",
        "发布时间", "响应/投标截止时间", "详情链接", "信息来源", "来源栏目",
        "正文摘要", "解析状态", "错误信息", "抓取时间",
    ]
    cols = [c for c in preferred_cols if c in df.columns] + [c for c in df.columns if c not in preferred_cols]
    df = df[cols]
    df.to_excel(EXCEL_PATH, index=False)
    logging.info("Excel 已导出：%s，共 %s 条", EXCEL_PATH, len(df))


# =========================
# HTTP 请求
# =========================
def make_session() -> requests.Session:
    session = requests.Session()
    # 避免本机系统代理/浏览器代理导致 requests 走坏代理，出现 ProxyError/SSLEOFError。
    # 如果你必须走代理，可改回 True。
    session.trust_env = False
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/147.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "keep-alive",
    })
    return session


def request_text(session: requests.Session, method: str, url: str, *, timeout: Optional[int] = None, retries: Optional[int] = None, **kwargs) -> str:
    last_err: Optional[Exception] = None
    use_timeout = timeout or REQUEST_TIMEOUT
    use_retries = retries or RETRIES
    for i in range(1, use_retries + 1):
        try:
            resp = session.request(method, url, timeout=use_timeout, **kwargs)
            resp.raise_for_status()
            # 山东站一般 UTF-8，但显式设一下更稳。
            if not resp.encoding or resp.encoding.lower() in {"iso-8859-1", "ascii"}:
                resp.encoding = resp.apparent_encoding or "utf-8"
            return resp.text
        except Exception as e:
            last_err = e
            wait = 1.5 * i
            logging.warning("请求失败 %s/%s：%s，等待 %.1f 秒；错误：%s", i, use_retries, url, wait, e)
            time.sleep(wait)
    raise RuntimeError(f"请求失败：{url}；最后错误：{last_err}")


def warmup(session: requests.Session) -> None:
    params = {
        "pagemode": "result",
        "appid": "all",
        "style": "1",
        "ck": "0",
        "pos": SEARCH_POS,
        "od": "0",
        "webid": "428",
        "tpl": TPL,
        "websiteid": WEBSITE_ID,
        "q": KEYWORDS[0] if KEYWORDS else "",
        "submit": "",
    }
    try:
        request_text(session, "GET", SEARCH_PAGE, params=params)
    except Exception as e:
        logging.warning("预热搜索页失败，可忽略：%s", e)


def warmup_search_context(session: requests.Session, keyword: str, page: int, begin: str, end: str) -> str:
    """按浏览器真实结果页 URL 预热一次，拿到 searchsign/JSESSIONID 等 cookie。

    山东站 cateSearch.do 对 cookie/searchsign 有时比较敏感。浏览器流程是先打开
    search.do?...cateid=16730...begin=...end=...，再由页面 JS 调 cateSearch.do。
    所以每次请求列表前先访问同条件的 search.do，稳定性更高。
    """
    referer = build_referer(keyword, page, begin, end)
    try:
        request_text(session, "GET", referer, headers={"Referer": BASE_URL + "/"})
    except Exception as e:
        logging.warning("预热结果页失败，继续尝试接口：%s", e)
    return referer


# =========================
# 列表解析
# =========================
def build_referer(keyword: str, page: int, begin: str, end: str) -> str:
    params = {
        "websiteid": WEBSITE_ID,
        "searchid": "5402",
        "pg": "",
        "p": str(page),
        "tpl": TPL,
        "cateid": CATE_ID,
        "total": "",
        "q": " " + keyword,
        "pq": keyword if USE_PQ_PHRASE else "",
        "oq": "",
        "eq": keyword if USE_EQ_PHRASE else "",
        "pos": SEARCH_POS,
        "begin": begin,
        "end": end,
    }
    return SEARCH_PAGE + "?" + urlencode(params)


def post_cate_search(session: requests.Session, keyword: str, page: int, begin: str, end: str) -> Tuple[str, Dict[str, Any]]:
    referer = warmup_search_context(session, keyword, page, begin, end)
    data = {
        "websiteid": WEBSITE_ID,
        "q": " " + keyword,
        "p": str(page),
        "pg": str(PAGE_SIZE),
        "cateid": CATE_ID,
        "pos": SEARCH_POS,
        "pq": keyword if USE_PQ_PHRASE else "",
        "oq": "",
        "eq": keyword if USE_EQ_PHRASE else "",
        "begin": begin,
        "end": end,
        "tpl": TPL,
    }
    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Origin": BASE_URL,
        "Referer": referer,
        "X-Requested-With": "XMLHttpRequest",
    }
    text = request_text(session, "POST", CATE_SEARCH_URL, headers=headers, data=data)
    return text, data


def find_record_lists(obj: Any) -> List[List[Dict[str, Any]]]:
    """递归寻找 JSON 中像搜索结果列表的 list[dict]。"""
    found: List[List[Dict[str, Any]]] = []
    if isinstance(obj, list):
        if obj and all(isinstance(x, dict) for x in obj):
            keys = set().union(*(x.keys() for x in obj[:5]))
            key_str = " ".join(keys).lower()
            if any(k in key_str for k in ["title", "url", "link", "content", "doc", "time", "date"]):
                found.append(obj)  # type: ignore[arg-type]
        for x in obj:
            found.extend(find_record_lists(x))
    elif isinstance(obj, dict):
        for v in obj.values():
            found.extend(find_record_lists(v))
    return found


def find_total(obj: Any) -> Optional[int]:
    if isinstance(obj, dict):
        for k, v in obj.items():
            if str(k).lower() in {"total", "totalcount", "count", "recordcount", "allcount"}:
                try:
                    return int(str(v).replace(",", ""))
                except Exception:
                    pass
        for v in obj.values():
            r = find_total(v)
            if r is not None:
                return r
    elif isinstance(obj, list):
        for v in obj:
            r = find_total(v)
            if r is not None:
                return r
    return None


def first_value(d: Dict[str, Any], names: Iterable[str]) -> str:
    lower_map = {str(k).lower(): k for k in d.keys()}
    for name in names:
        k = lower_map.get(name.lower())
        if k is not None:
            return str(d.get(k) or "")
    return ""


def normalize_url(url: str) -> str:
    url = clean_html_text(url)
    if not url:
        return ""
    # 有些 JSON 字段可能把 href 写在 HTML 里。
    m = re.search(r"href=[\"']([^\"']+)[\"']", url, flags=re.I)
    if m:
        url = m.group(1)
    return urljoin(BASE_URL + "/", url)


def iter_strings(obj: Any) -> Iterable[str]:
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from iter_strings(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from iter_strings(v)


def parse_list_from_embedded_html(obj: Any) -> Tuple[List[Dict[str, Any]], Optional[int]]:
    """有些 jsearch 接口会把结果列表作为 HTML 字符串放在 JSON 字段里。"""
    best_rows: List[Dict[str, Any]] = []
    best_total: Optional[int] = None
    for s in iter_strings(obj):
        if "/art/" not in s:
            continue
        candidate = unescape(s)
        rows, total = parse_list_from_html(candidate)
        if len(rows) > len(best_rows):
            best_rows, best_total = rows, total
    return best_rows, best_total


def try_load_json_loose(text: str) -> Optional[Any]:
    """兼容 JSON / JSONP / 前后带杂字符的响应。"""
    raw = text.strip()
    if not raw:
        return None
    for candidate in [raw, unescape(raw)]:
        try:
            return json.loads(candidate)
        except Exception:
            pass
    # JSONP: callback({...})
    m = re.search(r"^[\w$.]+\((.*)\)\s*;?\s*$", raw, flags=re.S)
    if m:
        try:
            return json.loads(m.group(1))
        except Exception:
            pass
    # 从响应里截取第一个 { 到最后一个 }
    i, j = raw.find("{"), raw.rfind("}")
    if 0 <= i < j:
        try:
            return json.loads(raw[i:j + 1])
        except Exception:
            pass
    # 从响应里截取第一个 [ 到最后一个 ]
    i, j = raw.find("["), raw.rfind("]")
    if 0 <= i < j:
        try:
            return json.loads(raw[i:j + 1])
        except Exception:
            pass
    return None


def parse_list_from_regex(text: str) -> Tuple[List[Dict[str, Any]], Optional[int]]:
    """最后兜底：直接从原始响应中用正则抓 /art/...html 链接及附近标题。"""
    raw = unescape(text)
    # 处理 JSON 字符串里的转义斜杠和 unicode 转义
    try:
        raw = raw.encode("utf-8").decode("unicode_escape")
    except Exception:
        pass
    raw = raw.replace("\/", "/")
    rows: List[Dict[str, Any]] = []
    seen = set()
    pattern = re.compile(r"(?P<url>(?:https?://ggzyjyzx\.shandong\.gov\.cn)?/art/\d{4}/\d{1,2}/\d{1,2}/art_\d+_\d+\.html)")
    for m in pattern.finditer(raw):
        url = normalize_url(m.group("url"))
        if url in seen:
            continue
        seen.add(url)
        start = max(0, m.start() - 800)
        end = min(len(raw), m.end() + 800)
        ctx = raw[start:end]
        title = ""
        # 常见：<a href="...">标题</a> 或 title字段
        esc_url = re.escape(m.group("url"))
        ma = re.search(r"<a[^>]+href=[\"']?" + esc_url + r"[\"']?[^>]*>(.*?)</a>", ctx, flags=re.I | re.S)
        if ma:
            title = clean_html_text(ma.group(1))
        if not title:
            mt = re.search(r"[\"'](?:title|doctitle|docTitle|ArticleTitle)[\"']\s*[:=]\s*[\"'](.{1,200}?)[\"']", ctx, flags=re.I | re.S)
            if mt:
                title = clean_html_text(mt.group(1))
        if not title:
            # 列表页常见：红色公告类型后跟蓝色标题，绿色链接在标题/摘要之后。
            # cateSearch.do 的返回有时不是标准 JSON，而是 HTML 片段/转义字符串，
            # 所以从链接前的上下文中取“最近一个以公告结尾的短句”作为标题。
            before = clean_html_text(ctx[:m.start() - start])
            candidates = re.findall(
                r"([\u4e00-\u9fa5A-Za-z0-9（）()【】《》、，,·.．_\-—\s]{4,180}?公告)",
                before,
            )
            candidates = [normalize_space(x) for x in candidates if normalize_space(x)]
            if candidates:
                # 取最后一个，一般最靠近链接，且更可能是标题。
                title = candidates[-1]
        if not title:
            before = clean_html_text(ctx[:m.start() - start])
            chunks = [x.strip() for x in re.split(r"\s{2,}|\n|https?://", before) if x.strip()]
            for ch in reversed(chunks):
                if 4 <= len(ch) <= 180 and any(k in ch for k in ["公告", "项目", "采购", "成交", "中标", "营商"]):
                    title = ch
                    break
        block = clean_html_text(ctx)[:500]
        pub = parse_date_any(block)
        col = ""
        mc = re.search(r"(采购公告|中标公告|成交公告|结果公告|废标公告|终止公告|更正公告)", block)
        if mc:
            col = mc.group(1)
        rows.append({"title": normalize_space(title), "url": url, "content": block, "pub_time": pub, "column": col, "raw": {}})
    mt = re.search(r"(?:找到|共|total[\"']?\s*[:=])\s*(\d+)\s*(?:条|,)?", clean_html_text(raw), flags=re.I)
    total = int(mt.group(1)) if mt else None
    return rows, total


def parse_list_from_json(text: str) -> Tuple[List[Dict[str, Any]], Optional[int]]:
    obj = try_load_json_loose(text)
    if obj is None:
        return [], None
    total = find_total(obj)
    lists = find_record_lists(obj)
    if not lists:
        rows_from_html, total_from_html = parse_list_from_embedded_html(obj)
        return rows_from_html, total or total_from_html
    # 选最长的结果列表。
    rows = max(lists, key=len)
    parsed: List[Dict[str, Any]] = []
    for item in rows:
        title = first_value(item, ["title", "doctitle", "name", "docTitle", "articleTitle", "shortTitle"])
        url = first_value(item, ["url", "link", "href", "docpuburl", "docpubUrl", "docUrl", "path", "linkurl", "linkUrl"])
        content = first_value(item, ["content", "summary", "digest", "description", "abs", "abstract", "subTitle"])
        pub_time = first_value(item, ["pubtime", "pubTime", "time", "date", "publishdate", "publishDate", "docreltime", "docRelTime", "pubDate"])
        column = first_value(item, ["category", "column", "columnName", "channel", "chnlname", "typename", "type", "className"])
        # 如果 url 没直接给，从字段里扫 art 链接。
        if not url:
            for v in item.values():
                m = re.search(r"https?://[^\s\"'<>]+/art/[^\s\"'<>]+\.html", str(v))
                if m:
                    url = m.group(0)
                    break
                m = re.search(r"/art/\d{4}/\d{1,2}/\d{1,2}/art_\d+_\d+\.html", str(v))
                if m:
                    url = m.group(0)
                    break
        # 如果标题字段是 HTML 链接，也从里面取干净标题。
        if title and "<" in title:
            title = clean_html_text(title)
        parsed.append({
            "title": normalize_space(title),
            "url": normalize_url(url),
            "content": normalize_space(content),
            "pub_time": parse_date_any(pub_time),
            "column": normalize_space(column),
            "raw": item,
        })
    return parsed, total


def parse_list_from_html(text: str) -> Tuple[List[Dict[str, Any]], Optional[int]]:
    soup = BeautifulSoup(text, "html.parser")
    parsed: List[Dict[str, Any]] = []

    total = None
    plain = normalize_space(soup.get_text(" "))
    m = re.search(r"约?找到\s*(\d+)\s*条", plain)
    if not m:
        m = re.search(r"共\s*(\d+)\s*条", plain)
    if m:
        total = int(m.group(1))

    # 优先解析搜索结果里带 /art/ 的链接。
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a.get("href") or ""
        if "/art/" not in href or not href.endswith(".html"):
            continue
        url = normalize_url(href)
        if url in seen:
            continue
        seen.add(url)
        title = normalize_space(a.get_text(" "))
        # 尝试拿附近摘要和日期。
        parent = a.find_parent(["li", "div", "p"]) or a.parent
        block_text = normalize_space(parent.get_text(" ") if parent else a.get_text(" "))
        pub = parse_date_any(block_text)
        column = ""
        mcol = re.search(r"(采购公告|中标公告|成交公告|结果公告|废标公告|终止公告|更正公告)", block_text)
        if mcol:
            column = mcol.group(1)
        parsed.append({
            "title": title,
            "url": url,
            "content": block_text,
            "pub_time": pub,
            "column": column,
            "raw": {},
        })
    return parsed, total


def parse_list_response(text: str) -> Tuple[List[Dict[str, Any]], Optional[int], str]:
    rows, total = parse_list_from_json(text)
    if rows:
        return rows, total, "json"
    rows, total = parse_list_from_html(text)
    if rows:
        return rows, total, "html"
    rows, total = parse_list_from_regex(text)
    if rows:
        return rows, total, "regex"
    return [], total, "none"


# =========================
# 详情解析
# =========================
def get_meta(soup: BeautifulSoup, name: str) -> str:
    tag = soup.find("meta", attrs={"name": name})
    if tag and tag.get("content") is not None:
        return normalize_space(tag.get("content"))
    return ""


def regex_first(text: str, patterns: Iterable[str]) -> str:
    for p in patterns:
        m = re.search(p, text, flags=re.I)
        if m:
            val = m.group(1).strip()
            val = re.split(r"\s{2,}|[；;]\s*", val)[0].strip()
            return val
    return ""


def text_until_next_heading(text: str, start_pattern: str) -> str:
    m = re.search(start_pattern, text)
    if not m:
        return ""
    sub = text[m.end():]
    # 到下一个中文序号标题为止。
    n = re.search(r"(?:\n|\s)(?:[一二三四五六七八九十]+、|\d+[、.．])", sub)
    return sub[: n.start()].strip() if n else sub[:500].strip()


def table_matrix(table) -> List[List[str]]:
    rows: List[List[str]] = []
    for tr in table.find_all("tr"):
        cells = [normalize_space(td.get_text(" ")) for td in tr.find_all(["th", "td"])]
        if any(cells):
            rows.append(cells)
    return rows


def parse_tables_for_money_and_supplier(soup: BeautifulSoup) -> Dict[str, str]:
    result = {
        "预算金额": "",
        "最高限价": "",
        "中标单位/成交供应商": "",
        "供应商地址": "",
        "中标金额/成交金额": "",
    }
    budget_parts: List[str] = []
    limit_parts: List[str] = []
    supplier_parts: List[str] = []
    addr_parts: List[str] = []
    amount_parts: List[str] = []

    for table in soup.find_all("table"):
        rows = table_matrix(table)
        if len(rows) < 2:
            continue
        header = rows[0]
        header_join = "|".join(header)
        for r in rows[1:]:
            if not r:
                continue
            # 对齐长度，不够时补空。
            rr = r + [""] * max(0, len(header) - len(r))
            mapping = {header[i]: rr[i] for i in range(min(len(header), len(rr)))}
            if "预算金额" in header_join:
                for h, v in mapping.items():
                    if "预算金额" in h and v:
                        budget_parts.append(v)
            if "最高限价" in header_join:
                for h, v in mapping.items():
                    if "最高限价" in h and v:
                        limit_parts.append(v)
            if any(k in header_join for k in ["中标供应商", "成交供应商", "供应商名称"]):
                for h, v in mapping.items():
                    if any(k in h for k in ["中标供应商名称", "成交供应商", "供应商名称", "中标供应商"]):
                        if v:
                            supplier_parts.append(v)
                    if "地址" in h and "供应商" in h:
                        if v:
                            addr_parts.append(v)
                    if any(k in h for k in ["中标金额", "成交金额"]):
                        if v:
                            amount_parts.append(v)
            if any(k in header_join for k in ["供应商报价", "小计", "单价"]):
                for h, v in mapping.items():
                    if any(k in h for k in ["供应商报价", "小计"]):
                        if v:
                            amount_parts.append(v)

    if budget_parts:
        result["预算金额"] = "; ".join(dict.fromkeys(budget_parts))
    if limit_parts:
        result["最高限价"] = "; ".join(dict.fromkeys(limit_parts))
    if supplier_parts:
        result["中标单位/成交供应商"] = "; ".join(dict.fromkeys(supplier_parts))
    if addr_parts:
        result["供应商地址"] = "; ".join(dict.fromkeys(addr_parts))
    if amount_parts:
        result["中标金额/成交金额"] = "; ".join(dict.fromkeys(amount_parts))
    return result


def parse_detail(html: str, url: str) -> Dict[str, str]:
    soup = BeautifulSoup(html, "html.parser")
    title = get_meta(soup, "ArticleTitle") or normalize_space((soup.select_one(".article-title") or soup.find("title") or soup).get_text(" "))
    column = get_meta(soup, "ColumnName") or get_meta(soup, "ColumnType")
    pubdate = get_meta(soup, "pubdate") or regex_first(normalize_space(soup.get_text(" ")), [r"发布时间[:：]\s*([0-9\-: ]{10,20})"])
    source = get_meta(soup, "contentSource") or get_meta(soup, "ContentSource")

    content_node = soup.select_one(".zhengwen") or soup.select_one(".article") or soup.body or soup
    text = normalize_space(content_node.get_text(" "))

    table_info = parse_tables_for_money_and_supplier(content_node if hasattr(content_node, 'find_all') else soup)

    project_code = regex_first(text, [
        r"项目编号\s*[:：]\s*([^\s，。；;]+)",
        r"采购项目编号\s*[:：]\s*([^\s，。；;]+)",
        r"订单\s*[:：]\s*([^，。；;\s]+)",
    ])
    project_name = regex_first(text, [
        r"项目名称\s*[:：]\s*(.+?)(?:\s+包号\s+预算金额|\s+采购需求\s*[:：]|\s+合同履行期限|\s+采购方式|\s+三[、.．]|\s+二[、.．]|$)",
        r"采购项目名称\s*[:：]\s*(.+?)(?:\s+采购方式|\s+三[、.．]|\s+二[、.．]|$)",
    ])
    purchase_method = regex_first(text, [
        r"采购方式\s*[:：]\s*([^\s，。；;]+)",
        r"项目联系方式.*?采购方式\s*[:：]\s*([^\s，。；;]+)",
    ])

    budget = regex_first(text, [
        r"预算金额(?:\([^)]*\))?[:：]\s*([\d,\.]+\s*(?:万元|元)?)",
        r"预算金额\s*([\d,\.]+\s*(?:万元|元)?)",
    ]) or table_info.get("预算金额", "")
    limit = regex_first(text, [
        r"最高限价(?:\([^)]*\))?[:：]\s*([\d,\.]+\s*(?:万元|元)?)",
        r"最高限价\s*([\d,\.]+\s*(?:万元|元)?)",
    ]) or table_info.get("最高限价", "")

    supplier = regex_first(text, [
        r"中标供应商名称\s*([^\n\s]+(?:公司|学院|中心|厂|店|社|集团|研究院)?)",
        r"成交供应商[:：]\s*([^\n，。；;]+)",
        r"供应商[:：]\s*([^\n，。；;]+)",
        r"中标供应商[:：]\s*([^\n，。；;]+)",
    ]) or table_info.get("中标单位/成交供应商", "")
    supplier_addr = regex_first(text, [
        r"供应商地址[:：]\s*([^\n；;]+)",
        r"中标供应商地址\s*([^\n；;]+)",
    ]) or table_info.get("供应商地址", "")
    amount = regex_first(text, [
        r"成交金额[:：]\s*([￥¥]?[\d,\.]+\s*(?:万元|元)?)",
        r"中标金额(?:\([^)]*\))?[:：]?\s*([￥¥]?[\d,\.]+\s*(?:万元|元)?)",
        r"供应商报价\(元\)\s*是否中标.*?([\d,\.]+)\s*是",
    ]) or table_info.get("中标金额/成交金额", "")

    purchaser = regex_first(text, [
        # 山东详情页常见格式：1 .采购人 信息 名 称： XXX 地址： XXX 联系方式：XXX
        r"采购人\s*信息.*?名\s*称\s*[:：]\s*(.+?)(?:\s+地址\s*[:：]|\s+地\s*址\s*[:：]|\s+联系方式\s*[:：]|\s+2[.．、]\s*采购代理机构|$)",
        r"采购单位\s*[:：]\s*(.+?)(?:\s+联系方式\s*[:：]|\s+20\d{2}年|\s+20\d{2}-\d{1,2}-\d{1,2}|$)",
        r"采购人\s*[:：]\s*(.+?)(?:\s+联系方式\s*[:：]|\s+地址\s*[:：]|\s+20\d{2}年|$)",
        r"采购单位名称\s*[:：]\s*([^\n，。；;]+)",
    ])
    agency = regex_first(text, [
        # 山东详情页常见格式：2.采购代理机构信息 名 称： XXX 地 址： XXX 联系方式：XXX
        r"采购代理机构\s*信息.*?名\s*称\s*[:：]\s*(.+?)(?:\s+地\s*址\s*[:：]|\s+地址\s*[:：]|\s+联系方式\s*[:：]|\s+3[.．、]\s*项目联系方式|$)",
        r"采购代理机构\s*[:：]\s*(.+?)(?:\s+联系方式\s*[:：]|\s+地\s*址\s*[:：]|\s+地址\s*[:：]|$)",
        r"代理机构\s*[:：]\s*(.+?)(?:\s+联系方式\s*[:：]|\s+地\s*址\s*[:：]|\s+地址\s*[:：]|$)",
    ])
    deadline = regex_first(text, [
        r"(?:提交响应文件|投标文件|响应文件提交|截止时间).*?(20\d{2}年\d{1,2}月\d{1,2}日\d{1,2}时\d{1,2}分)",
        r"(?:提交响应文件|投标文件|响应文件提交|截止时间).*?(20\d{2}[-/]\d{1,2}[-/]\d{1,2}\s*\d{1,2}:\d{1,2})",
    ])

    return {
        "公告类型": column,
        "标题": title,
        "项目名称": project_name,
        "项目编号": project_code,
        "采购方式": purchase_method,
        "采购人/采购单位": purchaser,
        "采购代理机构": agency,
        "预算金额": budget,
        "最高限价": limit,
        "中标单位/成交供应商": supplier,
        "供应商地址": supplier_addr,
        "中标金额/成交金额": amount,
        "发布时间": pubdate,
        "响应/投标截止时间": deadline,
        "详情链接": url,
        "信息来源": source,
        "来源栏目": column,
        "正文摘要": text[:500],
    }


# =========================
# 主流程
# =========================
def crawl_detail(session: requests.Session, url: str, referer: str = BASE_URL + "/") -> Tuple[Dict[str, str], str]:
    html = request_text(
        session,
        "GET",
        url,
        headers={"Referer": referer},
        timeout=DETAIL_TIMEOUT,
        retries=DETAIL_RETRIES,
    )
    info = parse_detail(html, url)
    return info, html


def crawl_one_keyword(session: requests.Session, keyword: str, visited: set[str]) -> None:
    begin = ymd_to_compact(START_DATE)
    end = ymd_to_compact(END_DATE)

    logging.info("开始关键词：%s，日期：%s 至 %s，分类 cateid=%s", keyword, START_DATE, END_DATE, CATE_ID)

    total_pages: Optional[int] = None
    page = 1
    while True:
        if MAX_PAGES and page > MAX_PAGES:
            break
        if total_pages is not None and page > total_pages:
            break

        logging.info("请求列表：关键词=%s，第 %s 页", keyword, page)
        text, payload = post_cate_search(session, keyword, page, begin, end)
        rows, total, parser_type = parse_list_response(text)
        raw_path = RAW_DIR / f"shandong_{safe_filename(keyword)}_p{page}_{parser_type}.txt"
        raw_path.write_text(text, encoding="utf-8", errors="ignore")

        # 如果接口响应没解析出列表，直接抓同条件搜索结果页再解析一次。
        if not rows:
            page_url = build_referer(keyword, page, begin, end)
            try:
                page_html = request_text(session, "GET", page_url, headers={"Referer": BASE_URL + "/"})
                (RAW_DIR / f"shandong_{safe_filename(keyword)}_p{page}_page.html").write_text(page_html, encoding="utf-8", errors="ignore")
                rows2, total2, parser2 = parse_list_response(page_html)
                if rows2:
                    rows, total, parser_type = rows2, total or total2, "page_" + parser2
            except Exception as e:
                logging.warning("列表接口无结果，搜索页兜底也失败：%s", e)

        if not rows:
            logging.warning("第 %s 页未解析到记录。原始响应前 300 字：%s", page, normalize_space(text[:300]))

        if total is not None and total_pages is None:
            total_pages = max(1, math.ceil(total / PAGE_SIZE))
            logging.info("接口总数=%s，预计总页数=%s", total, total_pages)

        logging.info("第 %s 页解析到 %s 条列表记录，解析方式=%s", page, len(rows), parser_type)
        if not rows:
            break

        kept = 0
        skipped_no_url = 0
        skipped_filter = 0
        skipped_visited = 0
        for row in rows:
            title = normalize_space(row.get("title", ""))
            detail_url = normalize_url(row.get("url", ""))
            if not detail_url:
                skipped_no_url += 1
                continue
            if LOCAL_TITLE_FILTER and not is_title_phrase_match(title, keyword):
                skipped_filter += 1
                logging.debug("标题不含完整关键词，过滤：keyword=%s, title=%s", keyword, title)
                continue
            if SKIP_VISITED and detail_url in visited:
                skipped_visited += 1
                continue

            kept += 1
            logging.info("详情：%s", title[:80] if title else detail_url)
            record: Dict[str, Any] = {
                "省份": "山东",
                "检索关键词": keyword,
                "列表标题": title,
                "列表公告类型": row.get("column", ""),
                "列表发布时间": row.get("pub_time", ""),
                "列表摘要": row.get("content", ""),
                "详情链接": detail_url,
                "抓取时间": now_ts(),
            }
            try:
                detail, html = crawl_detail(session, detail_url, referer=build_referer(keyword, page, begin, end))
                record.update(detail)
                record["解析状态"] = "成功"
                record["错误信息"] = ""
                detail_raw = RAW_DIR / f"detail_{safe_filename(title)}.html"
                detail_raw.write_text(html, encoding="utf-8", errors="ignore")
            except Exception as e:
                logging.exception("详情解析失败：%s", detail_url)
                record["公告类型"] = row.get("column", "")
                record["标题"] = title
                record["发布时间"] = row.get("pub_time", "")
                record["正文摘要"] = row.get("content", "")
                record["解析状态"] = "失败"
                record["错误信息"] = str(e)

            append_jsonl(record)
            logging.info("已写入记录：%s | 状态=%s | 当前本次记录数=%s", (record.get("标题") or record.get("列表标题") or detail_url)[:80], record.get("解析状态", ""), len(RUN_RECORDS))
            visited.add(detail_url)
            add_visited(detail_url)
            time.sleep(random.uniform(*SLEEP_DETAIL))

        logging.info(
            "第 %s 页处理完成：进入详情=%s，跳过无URL=%s，跳过标题过滤=%s，跳过已访问=%s",
            page, kept, skipped_no_url, skipped_filter, skipped_visited
        )

        save_progress({
            "keyword": keyword,
            "page": page,
            "total_pages": total_pages,
            "kept_on_page": kept,
            "updated_at": now_ts(),
        })
        page += 1
        time.sleep(random.uniform(*SLEEP_LIST))


def main() -> None:
    setup_logging()
    logging.info("开始爬取山东省公共资源交易中心（v9 保底落盘版）")
    logging.info("关键词：%s；日期：%s 至 %s；pq完整关键词：%s；本地标题过滤：%s；cateid=%s；跳过已访问=%s", KEYWORDS, START_DATE, END_DATE, USE_PQ_PHRASE, LOCAL_TITLE_FILTER, CATE_ID, SKIP_VISITED)
    session = make_session()
    warmup(session)
    visited = load_visited()

    try:
        for keyword in KEYWORDS:
            crawl_one_keyword(session, keyword, visited)
    except KeyboardInterrupt:
        logging.warning("用户中断，正在导出已抓取数据...")
    except Exception:
        logging.exception("运行失败，正在导出已抓取数据...")
        raise
    finally:
        export_excel()


if __name__ == "__main__":
    main()
