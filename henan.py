# -*- coding: utf-8 -*-
"""
河南省政府采购网自动爬虫（人工验证码版）

适用网站：
    https://zfcg.henan.gov.cn/henan/ggcx

特点：
1. 从 0 开始访问查询页，不依赖预先保存的 HTML。
2. 自动下载验证码图片，人工查看后在终端输入验证码。
3. 使用同一个 requests.Session 保持 cookie，提交查询表单后自动翻页。
4. 支持公告类型：采购公告、结果公告，也可扩展其他类型。
5. 详情页自动解析项目编号、项目名称、采购方式、预算金额、采购人、代理机构、供应商、中标/成交金额等字段。
6. 输出 JSONL、Excel、日志和 visited 去重文件。

运行：
    pip install requests beautifulsoup4 pandas openpyxl lxml
    python spider_henan_auto.py

如果验证码看不清，直接回车会重新拉取验证码；输入 q 可跳过当前公告类型。
"""

from __future__ import annotations

import json
import logging
import os
import random
import re
import sys
import time
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin, urlparse, parse_qs, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup


# ======================== 配置区 ========================

BASE_URL = "https://zfcg.henan.gov.cn"
HENAN_BASE_URL = f"{BASE_URL}/henan"
SEARCH_PAGE = f"{HENAN_BASE_URL}/ggcx"

KEYWORDS = ["营商环境"]
START_DATE = "2025-01-01"
END_DATE = date.today().strftime("%Y-%m-%d")

# 河南页面中公告类型按钮的 index：
# 所有类型=0，采购公告=1，变更公告=2，结果公告=3，废标公告=4，合同公告=5，验收结果公告=6，单一来源公示=7，非政府采购=8，其他=9，采购意向=10
ANNOUNCEMENT_TYPES = {
    "采购公告": "1",
    "结果公告": "3",
    # 如需要可取消注释：
    # "采购意向": "10",
    # "合同公告": "5",
}

PAGE_SIZE = 15
OUTPUT_DIR = Path("outputs_henan_auto")
CAPTCHA_DIR = OUTPUT_DIR / "captcha"
JSONL_PATH = OUTPUT_DIR / "henan_records.jsonl"
EXCEL_PATH = OUTPUT_DIR / "河南政府采购网_标讯明细.xlsx"
VISITED_PATH = OUTPUT_DIR / "henan_visited.txt"
LOG_PATH = OUTPUT_DIR / "henan_spider.log"

REQUEST_TIMEOUT = 25
DETAIL_DELAY_RANGE = (0.6, 1.5)
PAGE_DELAY_RANGE = (0.8, 1.8)
MAX_RETRY = 3

# 只保留标题包含关键词的记录；因为你要求搜索范围是“题目/标题”。
TITLE_FILTER = True


# ======================== 基础工具 ========================


# 读取同目录 config.yaml 中的 keywords/start_date/end_date。
from common_config import apply_common_config
apply_common_config(globals())



def ensure_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)


def setup_logging() -> None:
    ensure_dirs()
    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[logging.StreamHandler(sys.stdout), logging.FileHandler(LOG_PATH, encoding="utf-8")],
    )


def clean_text(s: Optional[str]) -> str:
    if not s:
        return ""
    s = BeautifulSoup(str(s), "html.parser").get_text(" ", strip=True)
    s = s.replace("\xa0", " ").replace("\u3000", " ")
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s.strip()


def text_one_line(s: Optional[str]) -> str:
    return re.sub(r"\s+", " ", clean_text(s)).strip()


def make_henan_abs_url(href: str, base: str = SEARCH_PAGE) -> str:
    """把河南站相对链接修正成完整 URL。

    河南站分页有时返回 /ggcx?...，但真实路径应为 /henan/ggcx?...；
    详情页有时也可能返回 /content?...，真实路径应为 /henan/content?...。
    直接用 urljoin(BASE_URL, href) 会丢掉 /henan，导致 404。
    """
    if not href:
        return ""
    url = urljoin(base, href)
    parsed = urlparse(url)
    if parsed.netloc == urlparse(BASE_URL).netloc:
        path = parsed.path or ""
        if path in {"/ggcx", "/content", "/search", "/getImage"} or path.startswith("/ggcx") or path.startswith("/content") or path.startswith("/getImage"):
            parsed = parsed._replace(path="/henan" + path)
            url = urlunparse(parsed)
    return url


def filename_safe(s: str, max_len: int = 50) -> str:
    s = re.sub(r"[\\/:*?\"<>|\s]+", "_", s.strip())
    return s[:max_len] or "captcha"


def msleep(a_b: Tuple[float, float]) -> None:
    time.sleep(random.uniform(*a_b))


def load_visited() -> set[str]:
    if not VISITED_PATH.exists():
        return set()
    return {line.strip() for line in VISITED_PATH.read_text(encoding="utf-8", errors="ignore").splitlines() if line.strip()}


def append_visited(url: str) -> None:
    with VISITED_PATH.open("a", encoding="utf-8") as f:
        f.write(url + "\n")


def append_jsonl(record: Dict) -> None:
    with JSONL_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def read_jsonl(path: Path) -> List[Dict]:
    if not path.exists():
        return []
    rows: List[Dict] = []
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            rows.append(json.loads(line))
        except Exception:
            pass
    return rows


def export_excel() -> None:
    rows = read_jsonl(JSONL_PATH)
    if not rows:
        logging.warning("没有可导出的记录。")
        return

    # 用详情链接去重，后出现的覆盖前面的空字段
    merged: Dict[str, Dict] = {}
    for r in rows:
        key = r.get("详情链接") or f"{r.get('标题')}|{r.get('发布时间')}"
        if key not in merged:
            merged[key] = r
        else:
            old = merged[key]
            for k, v in r.items():
                if v not in (None, "", [], {}):
                    old[k] = v

    df = pd.DataFrame(list(merged.values()))
    preferred_cols = [
        "省份", "检索关键词", "公告类型", "标题", "项目名称", "项目编号", "采购方式",
        "预算金额", "最高限价", "中标单位/成交供应商", "中标金额/成交金额", "供应商地址",
        "采购人/招标人", "代理机构", "区域", "发布时间", "投标/响应截止时间", "评审日期",
        "详情链接", "附件", "正文摘要", "解析状态", "错误信息",
    ]
    cols = [c for c in preferred_cols if c in df.columns] + [c for c in df.columns if c not in preferred_cols]
    df = df[cols]
    df.to_excel(EXCEL_PATH, index=False)
    logging.info("Excel 已导出：%s，共 %d 条", EXCEL_PATH, len(df))


# ======================== HTTP 客户端 ========================


def create_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Referer": SEARCH_PAGE,
    })
    return s


def request_text(session: requests.Session, method: str, url: str, **kwargs) -> str:
    last_err: Optional[Exception] = None
    for i in range(1, MAX_RETRY + 1):
        try:
            resp = session.request(method, url, timeout=REQUEST_TIMEOUT, **kwargs)
            resp.raise_for_status()
            # 河南站正常是 utf-8，保险处理。
            if not resp.encoding or resp.encoding.lower() in {"iso-8859-1", "ascii"}:
                resp.encoding = resp.apparent_encoding or "utf-8"
            return resp.text
        except Exception as e:
            last_err = e
            logging.warning("请求失败 %s/%s：%s，等待 %.1f 秒", i, MAX_RETRY, url, 1.5 * i)
            time.sleep(1.5 * i)
    raise RuntimeError(f"请求失败：{url}；最后错误：{last_err}")


def request_bytes(session: requests.Session, url: str, **kwargs) -> bytes:
    last_err: Optional[Exception] = None
    for i in range(1, MAX_RETRY + 1):
        try:
            resp = session.get(url, timeout=REQUEST_TIMEOUT, **kwargs)
            resp.raise_for_status()
            return resp.content
        except Exception as e:
            last_err = e
            logging.warning("下载失败 %s/%s：%s", i, MAX_RETRY, url)
            time.sleep(1.5 * i)
    raise RuntimeError(f"下载失败：{url}；最后错误：{last_err}")


# ======================== 验证码与查询提交 ========================


@dataclass
class SearchContext:
    so_code: str
    action_url: str
    captcha_url: str


def parse_search_context(html: str, current_url: str = SEARCH_PAGE) -> SearchContext:
    soup = BeautifulSoup(html, "lxml")

    action = ""
    form = soup.select_one("form#cgxxForm") or soup.find("form", attrs={"name": "cgxxForm"})
    if form and form.get("action"):
        action = make_henan_abs_url(form.get("action"), current_url)

    so_code = ""
    for candidate in [action, current_url]:
        qs = parse_qs(urlparse(candidate).query)
        if qs.get("soCode"):
            so_code = qs["soCode"][0]
            break

    img = soup.select_one("img#recode")
    captcha_url = ""
    if img and img.get("src"):
        captcha_url = make_henan_abs_url(img.get("src"), current_url)
        # 如果 src 是保存文件的本地相对路径，改用 soCode 拼接口。
        if not captcha_url.startswith("http") or "_files" in captcha_url:
            captcha_url = ""

    if not so_code:
        # 从 getImage/xxxx 中兜底解析
        m = re.search(r"/henan/getImage/([0-9a-fA-F]+)", html)
        if m:
            so_code = m.group(1)

    if not action and so_code:
        action = f"{BASE_URL}/henan/ggcx?soCode={so_code}"
    if not captcha_url and so_code:
        captcha_url = f"{BASE_URL}/henan/getImage/{so_code}?a={random.random()}"

    if not so_code or not action or not captcha_url:
        raise RuntimeError("未能从查询页解析 soCode/action/captcha_url，请检查页面结构或是否被拦截。")

    return SearchContext(so_code=so_code, action_url=action, captcha_url=captcha_url)


def open_file_for_user(path: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f"open '{path}'")
        else:
            os.system(f"xdg-open '{path}' >/dev/null 2>&1 &")
    except Exception:
        pass


def download_and_show_captcha(session: requests.Session, ctx: SearchContext, keyword: str, ann_type: str) -> Path:
    # 每次加随机数刷新验证码
    captcha_url = re.sub(r"(\?a=).*", rf"\g<1>{random.random()}", ctx.captcha_url)
    if "?" not in captcha_url:
        captcha_url += f"?a={random.random()}"
    content = request_bytes(session, captcha_url)
    path = CAPTCHA_DIR / f"{filename_safe(keyword)}_{filename_safe(ann_type)}_{int(time.time())}.jpg"
    path.write_bytes(content)
    logging.info("验证码图片已保存：%s", path)
    open_file_for_user(path)
    return path


def submit_search_with_manual_captcha(
    session: requests.Session,
    keyword: str,
    ann_type: str,
    ann_type_value: str,
    start_date: str,
    end_date: str,
) -> Optional[str]:
    """打开查询页 -> 下载验证码 -> 人工输入 -> POST 查询，成功返回搜索结果 HTML。"""
    for attempt in range(1, 8):
        logging.info("打开河南查询页，准备提交：关键词=%s，公告类型=%s，第 %d 次", keyword, ann_type, attempt)
        html = request_text(session, "GET", SEARCH_PAGE)
        ctx = parse_search_context(html, SEARCH_PAGE)
        download_and_show_captcha(session, ctx, keyword, ann_type)

        code = input(f"请输入验证码（{ann_type}/{keyword}，看不清直接回车刷新，输入 q 跳过）：").strip()
        if code.lower() == "q":
            logging.warning("用户跳过：%s/%s", ann_type, keyword)
            return None
        if not code:
            continue

        # timeType=5 表示指定时间；bidType 使用页面按钮 index。
        data = {
            "bidType": ann_type_value,
            "timeType": "5",
            "fromtime": start_date,
            "endtime": end_date,
            "title": keyword,
            "croporgan_name": "",
            "project_no": "",
            "gpmethod": "",
            "agency_name": "",
            "code": code,
        }
        headers = {
            "Referer": ctx.action_url,
            "Origin": BASE_URL,
            "Content-Type": "application/x-www-form-urlencoded",
        }
        result_html = request_text(session, "POST", ctx.action_url, data=data, headers=headers, allow_redirects=True)

        if is_search_success(result_html):
            logging.info("查询成功：关键词=%s，公告类型=%s", keyword, ann_type)
            return result_html

        msg = detect_error_message(result_html)
        logging.warning("查询可能失败：%s", msg or "未出现搜索结果，可能验证码错误或会话失效")
    raise RuntimeError(f"多次输入验证码仍未查询成功：{ann_type}/{keyword}")


def is_search_success(html: str) -> bool:
    text = text_one_line(html)
    return ("搜索结果" in text and "List2" in html) or ("您共搜到" in text)


def detect_error_message(html: str) -> str:
    text = text_one_line(html)
    for pat in [r"验证码.{0,20}(错误|不正确|失效)", r"请输入验证码", r"查询条件.{0,20}错误"]:
        m = re.search(pat, text)
        if m:
            return m.group(0)
    return ""


# ======================== 列表页解析 ========================


def parse_total_count(html: str) -> Optional[int]:
    text = text_one_line(html)
    m = re.search(r"共搜到\s*(\d+)\s*条", text)
    if m:
        return int(m.group(1))
    return None


def parse_total_pages(html: str) -> Optional[int]:
    text = text_one_line(html)
    m = re.search(r"共\s*(\d+)\s*页", text)
    if m:
        return int(m.group(1))
    return None


def parse_selected_announcement_type(html: str) -> str:
    soup = BeautifulSoup(html, "lxml")
    selected = soup.select_one("#searchbidTypeSel .item.select .item-title")
    return text_one_line(selected.get_text()) if selected else ""


def parse_list_items(html: str, keyword: str) -> List[Dict]:
    soup = BeautifulSoup(html, "lxml")
    rows: List[Dict] = []
    for li in soup.select(".List2 ul li"):
        a = li.find("a", href=True)
        if not a:
            continue
        title = text_one_line(a.get_text(" ", strip=True))
        detail_url = make_henan_abs_url(a.get("href"), SEARCH_PAGE)
        p_text = text_one_line(li.find("p").get_text(" ", strip=True) if li.find("p") else "")

        # 这些字段格式较稳定：公告类型：结果公告 区域：xxx 采购人：xxx 发布时间：yyyy-mm-dd
        record = {
            "省份": "河南",
            "检索关键词": keyword,
            "标题": title,
            "详情链接": detail_url,
            "公告类型": extract_inline_field(p_text, "公告类型", ["区域", "采购人", "发布时间"]),
            "区域": extract_inline_field(p_text, "区域", ["采购人", "发布时间"]),
            "采购人/招标人": extract_inline_field(p_text, "采购人", ["发布时间"]),
            "发布时间": extract_inline_field(p_text, "发布时间", []),
            "列表摘要": p_text,
        }
        rows.append(record)
    return rows


def extract_inline_field(text: str, field: str, next_fields: List[str]) -> str:
    # 兼容 “字段：值 字段2：值2”
    pattern = re.escape(field) + r"\s*[:：]\s*(.*?)"
    if next_fields:
        pattern += r"(?=" + "|".join(re.escape(f) + r"\s*[:：]" for f in next_fields) + r"|$)"
    else:
        pattern += r"$"
    m = re.search(pattern, text)
    return m.group(1).strip() if m else ""


def build_page_url_from_current(html: str, page_no: int) -> Optional[str]:
    soup = BeautifulSoup(html, "lxml")
    # 直接找目标页链接。
    for a in soup.select(".pager a[href]"):
        if text_one_line(a.get_text()) == str(page_no):
            return make_henan_abs_url(a.get("href"), SEARCH_PAGE)

    # 从任一分页链接提取 soCode 构造。
    any_a = soup.select_one(".pager a[href*='soCode=']")
    if any_a:
        href = make_henan_abs_url(any_a.get("href"), SEARCH_PAGE)
        qs = parse_qs(urlparse(href).query)
        so = qs.get("soCode", [""])[0]
        if so:
            return f"{BASE_URL}/henan/ggcx?appCode=H60&pageSize={PAGE_SIZE}&soCode={so}&pageNo={page_no}"
    return None


# ======================== 详情页解析 ========================


def fetch_detail_html(session: requests.Session, detail_url: str) -> str:
    html = request_text(session, "GET", detail_url, headers={"Referer": SEARCH_PAGE})
    # 有些详情正文会通过 $.get('/cmsweb.../webinfo/...htm') 再加载，若正文不完整则抓取该片段。
    if "项目基本情况" not in html and "中标情况" not in html and "成交情况" not in html:
        m = re.search(r"\$\.get\(\s*[\"']([^\"']+)[\"']", html)
        if m:
            inner_url = urljoin(BASE_URL, m.group(1))
            try:
                inner = request_text(session, "GET", inner_url, headers={"Referer": detail_url})
                html += "\n" + inner
            except Exception as e:
                logging.warning("详情内嵌正文抓取失败：%s，%s", inner_url, e)
    else:
        # 即使已有正文，也尝试拿最新 cmsweb 内容拼接，便于解析附件或动态正文。
        m = re.search(r"\$\.get\(\s*[\"']([^\"']+)[\"']", html)
        if m:
            inner_url = urljoin(BASE_URL, m.group(1))
            try:
                inner = request_text(session, "GET", inner_url, headers={"Referer": detail_url})
                if len(inner) > 100 and inner not in html:
                    html += "\n" + inner
            except Exception:
                pass
    return html


def parse_detail(html: str, base_record: Dict) -> Dict:
    soup = BeautifulSoup(html, "lxml")
    content = soup.select_one("#print-content") or soup.select_one("#content") or soup.body or soup
    text = content.get_text("\n", strip=True)
    flat = text_one_line(text)

    record = dict(base_record)
    record["解析状态"] = "success"
    record["错误信息"] = ""

    # 标题、发布机构、发布日期
    h1 = soup.select_one("#print-content h1") or soup.find("h1")
    if h1:
        record["标题"] = text_one_line(h1.get_text()) or record.get("标题", "")
    pub_org = regex_first(flat, r"发布机构[:：]\s*(.*?)\s*发布日期[:：]")
    pub_date = regex_first(flat, r"发布日期[:：]\s*([0-9]{4}-[0-9]{2}-[0-9]{2}(?:\s+[0-9]{2}:[0-9]{2})?)")
    if pub_org and not record.get("采购人/招标人"):
        record["采购人/招标人"] = pub_org
    if pub_date:
        record["发布时间"] = pub_date

    # 常规字段
    patterns = {
        "项目编号": [r"采购项目编号[:：]\s*([^\n\r]+?)(?=\s*\d+[、.]|\s*[一二三四五六七八九十][、.]|$)", r"项目编号[:：]\s*([^\n\r]+?)(?=\s*\d+[、.]|\s*[一二三四五六七八九十][、.]|$)"],
        "项目名称": [r"采购项目名称[:：]\s*([^\n\r]+?)(?=\s*\d+[、.]|\s*[一二三四五六七八九十][、.]|$)", r"项目名称[:：]\s*([^\n\r]+?)(?=\s*\d+[、.]|\s*[一二三四五六七八九十][、.]|$)"],
        "采购方式": [r"采购方式[:：]\s*([^\n\r]+?)(?=\s*\d+[、.]|\s*[一二三四五六七八九十][、.]|$)"],
        "预算金额": [r"预算金额[:：]\s*([^\n\r]+?)(?=\s*最高限价|\s*采购需求|\s*\d+[、.]|\s*[一二三四五六七八九十][、.]|$)"],
        "最高限价": [r"最高限价[:：]\s*([^\n\r]+?)(?=\s*采购需求|\s*合同履行|\s*\d+[、.]|\s*[一二三四五六七八九十][、.]|$)"],
        "投标/响应截止时间": [r"响应文件提交截止时间[:：]\s*([^\n\r]+?)(?=\s*地点|\s*[一二三四五六七八九十][、.]|$)", r"投标截止时间[:：]\s*([^\n\r]+?)(?=\s*地点|\s*[一二三四五六七八九十][、.]|$)", r"开标时间[:：]\s*([^\n\r]+?)(?=\s*地点|\s*[一二三四五六七八九十][、.]|$)"],
        "评审日期": [r"评审日期[:：]\s*([^\n\r]+?)(?=\s*[一二三四五六七八九十][、.]|$)"],
    }
    for field, pats in patterns.items():
        val = first_match(flat, pats)
        if val:
            record[field] = normalize_value(val)

    purchaser, agency = parse_contact_info(flat)
    if purchaser:
        record["采购人/招标人"] = purchaser
    if agency:
        record["代理机构"] = agency

    supplier, amount, supplier_addr = parse_supplier_and_amount(soup, flat)
    if supplier:
        record["中标单位/成交供应商"] = supplier
    if amount:
        record["中标金额/成交金额"] = amount
    if supplier_addr:
        record["供应商地址"] = supplier_addr

    attachments = []
    for a in soup.select(".List1 a[href], a[href$='.pdf'], a[href$='.doc'], a[href$='.docx'], a[href$='.xls'], a[href$='.xlsx']"):
        name = text_one_line(a.get_text())
        href = urljoin(BASE_URL, a.get("href"))
        if href and name:
            attachments.append(f"{name} {href}")
    record["附件"] = "；".join(dict.fromkeys(attachments))

    record["正文摘要"] = flat[:800]
    return record


def regex_first(text: str, pat: str) -> str:
    m = re.search(pat, text, flags=re.S)
    return normalize_value(m.group(1)) if m else ""


def first_match(text: str, patterns: Iterable[str]) -> str:
    for pat in patterns:
        val = regex_first(text, pat)
        if val:
            return val
    return ""


def normalize_value(v: str) -> str:
    v = text_one_line(v)
    v = re.sub(r"^[：:、，,\s]+", "", v)
    v = re.sub(r"\s+$", "", v)
    return v


def parse_contact_info(flat: str) -> Tuple[str, str]:
    purchaser = ""
    agency = ""

    # 常见格式：1. 采购人信息 名称：xxx 地址：xxx 联系人...
    m = re.search(r"1\.\s*采购人信息\s*名称[:：]\s*(.*?)(?=\s*地址[:：]|\s*联系人[:：]|\s*联系方式[:：]|\s*2\.\s*采购代理机构|$)", flat)
    if m:
        purchaser = normalize_value(m.group(1))
    else:
        m = re.search(r"采购人[:：]\s*(.*?)(?=\s*地址[:：]|\s*联系人[:：]|\s*联系方式[:：]|\s*代理机构|$)", flat)
        if m:
            purchaser = normalize_value(m.group(1))

    m = re.search(r"2\.\s*采购代理机构信息(?:（如有）)?\s*名称[:：]\s*(.*?)(?=\s*地址[:：]|\s*联系人[:：]|\s*联系方式[:：]|\s*3\.\s*项目联系方式|$)", flat)
    if m:
        agency = normalize_value(m.group(1))
    else:
        m = re.search(r"代理机构[:：]\s*(.*?)(?=\s*地址[:：]|\s*联系人[:：]|\s*联系方式[:：]|$)", flat)
        if m:
            agency = normalize_value(m.group(1))

    return purchaser, agency


def parse_supplier_and_amount(soup: BeautifulSoup, flat: str) -> Tuple[str, str, str]:
    supplier = ""
    amount = ""
    addr = ""

    # 表格解析：找包含“供应商名称”和“中标金额/成交金额”的表头，然后读下一行。
    for table in soup.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            cells = [text_one_line(c.get_text(" ", strip=True)) for c in tr.find_all(["td", "th"])]
            cells = [c for c in cells if c]
            if cells:
                rows.append(cells)
        for i, cells in enumerate(rows):
            header = "|".join(cells)
            if "供应商名称" in header and ("中标金额" in header or "成交金额" in header or "中标（成交）金额" in header):
                # 建列索引
                idx_supplier = find_cell_index(cells, ["供应商名称", "中标供应商", "成交供应商"])
                idx_addr = find_cell_index(cells, ["地址", "地 址"])
                idx_amount = find_cell_index(cells, ["中标金额", "成交金额", "中标（成交）金额"])
                if i + 1 < len(rows):
                    data = rows[i + 1]
                    if idx_supplier is not None and idx_supplier < len(data):
                        supplier = data[idx_supplier]
                    if idx_addr is not None and idx_addr < len(data):
                        addr = data[idx_addr]
                    if idx_amount is not None and idx_amount < len(data):
                        amount = data[idx_amount]
                        # 单位可能在后一列。
                        if idx_amount + 1 < len(data) and data[idx_amount + 1] in {"元", "万元", "%"}:
                            amount += data[idx_amount + 1]
                    return supplier, amount, addr

    # 文本兜底。
    supplier = first_match(flat, [
        r"供应商名称[:：]\s*(.*?)(?=\s*供应商地址|\s*地址[:：]|\s*中标|\s*成交|$)",
        r"中标供应商[:：]\s*(.*?)(?=\s*供应商地址|\s*地址[:：]|\s*中标|\s*成交|$)",
        r"成交供应商[:：]\s*(.*?)(?=\s*供应商地址|\s*地址[:：]|\s*中标|\s*成交|$)",
    ])
    addr = first_match(flat, [
        r"供应商地址[:：]\s*(.*?)(?=\s*中标|\s*成交|\s*四、|$)",
        r"地\s*址[:：]\s*(.*?)(?=\s*中标|\s*成交|\s*四、|$)",
    ])
    amount = first_match(flat, [
        r"中标（成交）金额[:：]\s*([0-9,.]+\s*(?:元|万元)?)",
        r"中标金额[:：]\s*([0-9,.]+\s*(?:元|万元)?)",
        r"成交金额[:：]\s*([0-9,.]+\s*(?:元|万元)?)",
    ])
    return supplier, amount, addr


def find_cell_index(cells: List[str], names: List[str]) -> Optional[int]:
    for idx, c in enumerate(cells):
        for name in names:
            if name in c:
                return idx
    return None


# ======================== 主流程 ========================


def crawl_one_search(session: requests.Session, keyword: str, ann_type: str, ann_value: str, visited: set[str]) -> None:
    html = submit_search_with_manual_captcha(session, keyword, ann_type, ann_value, START_DATE, END_DATE)
    if not html:
        return

    selected = parse_selected_announcement_type(html)
    total_count = parse_total_count(html)
    total_pages = parse_total_pages(html) or 1
    logging.info("搜索完成：关键词=%s，公告类型=%s，页面选中=%s，总数=%s，总页数=%s", keyword, ann_type, selected, total_count, total_pages)

    current_html = html
    for page_no in range(1, total_pages + 1):
        if page_no > 1:
            page_url = build_page_url_from_current(html, page_no)
            if not page_url:
                logging.warning("未能构造第 %s 页 URL，停止翻页。", page_no)
                break
            logging.info("抓取列表页：%s", page_url)
            try:
                current_html = request_text(session, "GET", page_url, headers={"Referer": SEARCH_PAGE})
            except Exception as e:
                logging.error("列表页抓取失败，跳过后续分页：%s", e)
                break
            msleep(PAGE_DELAY_RANGE)

        items = parse_list_items(current_html, keyword)
        logging.info("第 %s/%s 页解析到 %s 条列表记录", page_no, total_pages, len(items))
        for item in items:
            title = item.get("标题", "")
            if TITLE_FILTER and keyword not in title:
                logging.info("标题不含关键词，跳过：%s", title)
                continue

            url = item.get("详情链接", "")
            if not url:
                continue
            if url in visited:
                logging.info("已抓取，跳过：%s", url)
                continue

            try:
                logging.info("详情：%s", title)
                detail_html = fetch_detail_html(session, url)
                record = parse_detail(detail_html, item)
            except Exception as e:
                logging.exception("详情解析失败：%s", url)
                record = dict(item)
                record["解析状态"] = "failed"
                record["错误信息"] = str(e)

            append_jsonl(record)
            append_visited(url)
            visited.add(url)
            msleep(DETAIL_DELAY_RANGE)


def main() -> None:
    ensure_dirs()
    setup_logging()
    logging.info("开始爬取河南省政府采购网，人工验证码模式")
    logging.info("时间范围：%s 至 %s；关键词：%s；公告类型：%s", START_DATE, END_DATE, KEYWORDS, list(ANNOUNCEMENT_TYPES.keys()))

    session = create_session()
    visited = load_visited()

    try:
        for keyword in KEYWORDS:
            for ann_type, ann_value in ANNOUNCEMENT_TYPES.items():
                crawl_one_search(session, keyword, ann_type, ann_value, visited)
        export_excel()
    except KeyboardInterrupt:
        logging.warning("用户中断，正在导出已抓取数据...")
        export_excel()
    except Exception:
        logging.exception("运行失败，正在导出已抓取数据...")
        export_excel()
        raise


if __name__ == "__main__":
    main()
