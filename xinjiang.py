# -*- coding: utf-8 -*-
"""
新疆政府采购网爬虫（Playwright 浏览器后台版）

适用网站：
    http://www.ccgp-xinjiang.gov.cn/site/category?parentId=3661&childrenCode=ZcyAnnouncement

特点：
1. 不需要手工复制 Cookie；
2. 使用 Playwright 后台 Chromium 自动执行网页 JS，通过新疆站点 WAF 校验；
3. 在浏览器上下文中 fetch /portal/category，自动携带浏览器 Cookie；
4. 支持采购意向、采购公告、采购结果公告三类；
5. 输出 JSON、JSONL、Excel；
6. 支持断点去重、失败继续、原始 HTML/响应保存。

首次运行前安装：
    pip install playwright pandas openpyxl beautifulsoup4
    python -m playwright install chromium

运行：
    python xinjiang.py
"""

from __future__ import annotations

import json
import logging
import math
import os
import random
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import quote, urljoin

import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


# =========================
# 1. 用户配置区
# =========================

KEYWORDS = ["营商环境"]

START_DATE = "2025-01-01"
END_DATE = "2026-05-08"

# 新疆政府采购网：采购意向、采购公告、采购结果公告
CATEGORY_TASKS = [
    {"name": "采购意向", "categoryCode": "ZcyAnnouncement11"},
    {"name": "采购公告", "categoryCode": "ZcyAnnouncement2"},
    {"name": "采购结果公告", "categoryCode": "ZcyAnnouncement4"},
]

PAGE_SIZE = 15

# Codespaces / 服务器后台建议 True；本地调试想看浏览器可改 False
HEADLESS = True

# 如果新疆站点偶发 WAF 失败，可以把 WAIT_AFTER_OPEN 调大到 8~12
WAIT_AFTER_OPEN = 5.0

# 详情页最多等待秒数
DETAIL_TIMEOUT_MS = 45000

# 列表接口 fetch 超时时间。避免新疆站接口/WAF 卡住后长期无响应。
API_FETCH_TIMEOUT_MS = 15000

# 列表接口最多重试次数
LIST_RETRIES = 4

# 详情页最多重试次数
DETAIL_RETRIES = 2

# 是否跳过已经爬过的详情 URL
SKIP_VISITED = True

# 是否只保留标题中连续包含关键词的记录
# 新疆接口本身 keyword 已经比较准，一般 False 即可；如果混入很多无关结果，可改 True。
STRICT_TITLE_FILTER = False

# 保存浏览器缓存目录。保留它可以复用 WAF Cookie，减少每次重新校验。
BROWSER_USER_DATA_DIR = "browser_cache/xinjiang"

OUTPUT_DIR = "outputs_xinjiang"

BASE_URL = "http://www.ccgp-xinjiang.gov.cn"
CATEGORY_PAGE = (
    BASE_URL
    + "/site/category?parentId=3661&childrenCode=ZcyAnnouncement"
    + "&utm=site.site-PC-42166.1718-block_comp_1709189989843012.4"
)
CATEGORY_API = BASE_URL + "/portal/category"
DETAIL_URL_TEMPLATE = BASE_URL + "/site/detail?parentId=3661&articleId={article_id}"


# =========================
# 2. 日志与工具函数
# =========================


# 读取同目录 config.yaml 中的 keywords/start_date/end_date。
from common_config import apply_common_config
apply_common_config(globals())


OUT = Path(OUTPUT_DIR)
RAW_DIR = OUT / "raw"
CAPTURE_DIR = OUT / "screenshots"
JSONL_PATH = OUT / "xinjiang_records.jsonl"
JSON_PATH = OUT / "xinjiang_records.json"
XLSX_PATH = OUT / "新疆政府采购网_标讯明细.xlsx"
VISITED_PATH = OUT / "xinjiang_visited.txt"
LOG_PATH = OUT / "xinjiang_spider.log"


def ensure_dirs() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    CAPTURE_DIR.mkdir(parents=True, exist_ok=True)
    Path(BROWSER_USER_DATA_DIR).mkdir(parents=True, exist_ok=True)


def setup_logging() -> None:
    ensure_dirs()
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)


def now_ms() -> int:
    return int(time.time() * 1000)


def sleep_jitter(a: float = 0.8, b: float = 1.8) -> None:
    time.sleep(random.uniform(a, b))


def clean_html_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"<script[\s\S]*?</script>", " ", s, flags=re.I)
    s = re.sub(r"<style[\s\S]*?</style>", " ", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = s.replace("&nbsp;", " ").replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def normalize_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ")
    # 保留换行，方便正则按段落截取
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n\s+", "\n", s)
    s = re.sub(r"\s+\n", "\n", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def compact_text(s: Any) -> str:
    return re.sub(r"\s+", "", str(s or ""))


def ms_to_date(ms: Any) -> str:
    if ms in (None, "", "null"):
        return ""
    try:
        ms_int = int(float(ms))
        return datetime.fromtimestamp(ms_int / 1000).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return str(ms)


def amount_to_yuan(value: Any) -> str:
    """
    接口里的 budgetPrice 多数是元字符串，如 "1980000"；
    正文里可能是 "198.00万元" 或 "￥24000.00"。
    这里不强行换算所有文本，只清洗接口数值。
    """
    if value in (None, "", "null"):
        return ""
    s = str(value).strip()
    if not s or s == "0":
        return ""
    try:
        n = float(s)
        if n.is_integer():
            return str(int(n))
        return f"{n:.2f}".rstrip("0").rstrip(".")
    except Exception:
        return s


def write_jsonl(record: Dict[str, Any]) -> None:
    with JSONL_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def read_jsonl(path: Path) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if not path.exists():
        return rows
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except Exception:
                pass
    return rows


def load_visited() -> set[str]:
    if not VISITED_PATH.exists():
        return set()
    return set(x.strip() for x in VISITED_PATH.read_text(encoding="utf-8").splitlines() if x.strip())


def save_visited(url: str) -> None:
    with VISITED_PATH.open("a", encoding="utf-8") as f:
        f.write(url + "\n")


def export_outputs(records: List[Dict[str, Any]]) -> None:
    # 合并历史 jsonl 和本次记录，并按 detail_url 去重
    all_records = read_jsonl(JSONL_PATH)
    if records:
        all_records.extend(records)

    dedup: Dict[str, Dict[str, Any]] = {}
    for r in all_records:
        key = r.get("detail_url") or (r.get("province", "") + r.get("title", "") + r.get("publish_time", ""))
        if key:
            dedup[key] = r

    final_records = list(dedup.values())
    final_records.sort(key=lambda x: (x.get("publish_time", ""), x.get("title", "")), reverse=True)

    JSON_PATH.write_text(json.dumps(final_records, ensure_ascii=False, indent=2), encoding="utf-8")

    if not final_records:
        logging.warning("没有可导出的记录。")
        return

    columns = [
        "province",
        "keyword",
        "category",
        "announcement_type",
        "title",
        "project_name",
        "project_code",
        "purchase_name",
        "agency_name",
        "supplier_name",
        "budget_amount",
        "win_amount",
        "procurement_method",
        "gp_catalog_name",
        "district_name",
        "publish_time",
        "bid_opening_time",
        "detail_url",
        "crawl_status",
        "error",
        "summary",
    ]

    df = pd.DataFrame(final_records)
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    df = df[columns]

    with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="新疆标讯明细")

    logging.info("JSON 已导出：%s，共 %d 条", JSON_PATH, len(final_records))
    logging.info("Excel 已导出：%s，共 %d 条", XLSX_PATH, len(final_records))


# =========================
# 3. 字段解析
# =========================

def first_match(text: str, patterns: Iterable[str], flags: int = re.S) -> str:
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            val = m.group(1) if m.groups() else m.group(0)
            val = clean_field(val)
            if val:
                return val
    return ""


def clean_field(s: Any) -> str:
    s = str(s or "")
    s = s.replace("\xa0", " ").replace("&nbsp;", " ")
    s = re.sub(r"\s+", " ", s).strip(" \t\r\n:：，,。;；")
    # 截断明显串到下一个字段的情况
    s = re.split(
        r"(?:\s{2,}|项目编号|项目名称|采购人信息|采购代理机构信息|供应商名称|供应商地址|中标金额|成交金额|预算金额|最高限价|联系方式|四、|五、|六、|七、)",
        s,
        maxsplit=1,
    )[0].strip(" \t\r\n:：，,。;；")
    return s


def parse_amount_from_text(text: str) -> str:
    patterns = [
        r"(?:中标|成交|中标成交|中标\(成交\)|成交总|中标总)\s*(?:金额|价|报价)\s*[:：]?\s*[￥¥]?\s*([0-9][0-9,]*(?:\.\d+)?)\s*(万元|万|元)?",
        r"(?:总价|报价|投标报价)\s*[:：]?\s*[￥¥]?\s*([0-9][0-9,]*(?:\.\d+)?)\s*(万元|万|元)?",
        r"(?:小计|合计)\s*(?:（元）|\(元\)|元)?\s*[:：]?\s*[￥¥]?\s*([0-9][0-9,]*(?:\.\d+)?)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.S)
        if m:
            num = m.group(1).replace(",", "")
            unit = m.group(2) if len(m.groups()) >= 2 else ""
            if unit in ("万元", "万"):
                try:
                    return str(float(num) * 10000).rstrip("0").rstrip(".")
                except Exception:
                    return num + unit
            return num
    return ""


def parse_budget_from_text(text: str) -> str:
    patterns = [
        r"(?:预算金额|项目预算|预算价|采购预算)\s*[:：]?\s*[￥¥]?\s*([0-9][0-9,]*(?:\.\d+)?)\s*(万元|万|元)?",
        r"(?:最高限价)\s*[:：]?\s*[￥¥]?\s*([0-9][0-9,]*(?:\.\d+)?)\s*(万元|万|元)?",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.S)
        if m:
            num = m.group(1).replace(",", "")
            unit = m.group(2) if len(m.groups()) >= 2 else ""
            if unit in ("万元", "万"):
                try:
                    return str(float(num) * 10000).rstrip("0").rstrip(".")
                except Exception:
                    return num + unit
            return num
    return ""


def parse_supplier_from_text(text: str) -> str:
    patterns = [
        r"(?:供应商名称|成交供应商|中标供应商|中标人|成交人)\s*[:：]?\s*([^\n\r；;，,。]+)",
        r"(?:供应商)\s*[:：]\s*([^\n\r；;，,。]+)",
        # 表格纯文本兜底：供应商名称 供应商地址 中标金额...
        r"供应商名称\s+供应商地址[\s\S]{0,80}?\n?\s*([^\n\r]+?)\s+(?:新疆|北京市|上海市|天津市|重庆市|山东|河南|河北|山西|广东|四川|陕西|甘肃|宁夏|青海|内蒙古)",
    ]
    return first_match(text, patterns)


def parse_purchase_name_from_text(text: str) -> str:
    patterns = [
        r"采购人信息[\s\S]{0,80}?名\s*称\s*[:：]\s*([^\n\r]+)",
        r"采购单位信息[\s\S]{0,80}?名\s*称\s*[:：]\s*([^\n\r]+)",
        r"采购人\s*[:：]\s*([^\n\r；;，,。]+)",
        r"采购单位\s*[:：]\s*([^\n\r；;，,。]+)",
        r"名\s*称\s*[:：]\s*([^\n\r]+)",
    ]
    return first_match(text, patterns)


def parse_agency_from_text(text: str) -> str:
    patterns = [
        r"采购代理机构信息[\s\S]{0,100}?名\s*称\s*[:：]\s*([^\n\r]+)",
        r"代理机构信息[\s\S]{0,100}?名\s*称\s*[:：]\s*([^\n\r]+)",
        r"采购代理机构\s*[:：]\s*([^\n\r；;，,。]+)",
        r"代理机构\s*[:：]\s*([^\n\r；;，,。]+)",
    ]
    return first_match(text, patterns)


def parse_detail_html(html: str, body_text: str = "") -> Dict[str, str]:
    soup = BeautifulSoup(html or "", "html.parser")

    title = ""
    if soup.title and soup.title.get_text(strip=True):
        title = soup.title.get_text(strip=True)

    # 常见标题位置兜底
    for sel in ["h1", ".article-title", ".title", ".detail-title"]:
        node = soup.select_one(sel)
        if node and node.get_text(strip=True):
            t = node.get_text(" ", strip=True)
            if len(t) >= 6:
                title = t
                break

    html_text = normalize_text(soup.get_text("\n", strip=True))
    text = normalize_text((body_text or "") + "\n" + html_text)

    # 兼容政采云详情页中 window.__INITIAL_STATE__ / renderData 文本
    if len(text) < 200:
        text = normalize_text(clean_html_text(html))

    project_code = first_match(text, [
        r"项目编号\s*[:：]\s*([^\n\r]+)",
        r"采购项目编号\s*[:：]\s*([^\n\r]+)",
    ])

    project_name = first_match(text, [
        r"项目名称\s*[:：]\s*([^\n\r]+)",
        r"采购项目名称\s*[:：]\s*([^\n\r]+)",
    ])

    purchase_name = parse_purchase_name_from_text(text)
    agency_name = parse_agency_from_text(text)
    supplier_name = parse_supplier_from_text(text)
    budget_amount = parse_budget_from_text(text)
    win_amount = parse_amount_from_text(text)

    # 发布时间
    publish_time = first_match(text, [
        r"发布时间\s*[:：]\s*([0-9]{4}[-年/][0-9]{1,2}[-月/][0-9]{1,2}(?:\s+[0-9]{1,2}[:：][0-9]{1,2}(?::[0-9]{1,2})?)?)",
        r"发布日期\s*[:：]\s*([0-9]{4}[-年/][0-9]{1,2}[-月/][0-9]{1,2}(?:\s+[0-9]{1,2}[:：][0-9]{1,2}(?::[0-9]{1,2})?)?)",
    ])

    summary = clean_field(text[:500])

    return {
        "detail_title": title,
        "project_code_detail": project_code,
        "project_name_detail": project_name,
        "purchase_name_detail": purchase_name,
        "agency_name_detail": agency_name,
        "supplier_name_detail": supplier_name,
        "budget_amount_detail": budget_amount,
        "win_amount_detail": win_amount,
        "publish_time_detail": publish_time,
        "summary": summary,
    }


# =========================
# 4. Playwright 爬虫核心
# =========================

@dataclass
class CategoryResult:
    total: int
    rows: List[Dict[str, Any]]
    raw: Dict[str, Any]


class XinjiangBrowserSpider:
    def __init__(self) -> None:
        ensure_dirs()
        self.records_this_run: List[Dict[str, Any]] = []
        self.visited = load_visited()

    def _launch_context(self, p):
        """
        使用 persistent_context 保存浏览器缓存。
        注意：persistent_context 本身就是 context，不需要 browser.new_context。
        """
        args = [
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
            "--disable-infobars",
        ]

        context = p.chromium.launch_persistent_context(
            user_data_dir=BROWSER_USER_DATA_DIR,
            headless=HEADLESS,
            args=args,
            viewport={"width": 1366, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/147.0.0.0 Safari/537.36"
            ),
            locale="zh-CN",
            timezone_id="Asia/Shanghai",
            ignore_https_errors=True,
        )
        context.set_default_timeout(60000)
        return context

    def open_home_and_pass_waf(self, page) -> None:
        logging.info("访问首页获取 Cookie：%s", CATEGORY_PAGE)
        try:
            page.goto(CATEGORY_PAGE, wait_until="domcontentloaded", timeout=90000)
        except PlaywrightTimeoutError:
            logging.warning("首页 domcontentloaded 超时，继续等待页面脚本。")

        page.wait_for_timeout(int(WAIT_AFTER_OPEN * 1000))

        # 若是 WAF 空白/校验页，刷新一次再等
        html = ""
        try:
            html = page.content()
        except Exception:
            pass

        if "aliyun_waf" in html or "renderData" in html or "window._waf" in html:
            logging.warning("检测到 WAF 页面，等待并刷新一次。")
            page.wait_for_timeout(5000)
            try:
                page.reload(wait_until="domcontentloaded", timeout=90000)
            except Exception:
                pass
            page.wait_for_timeout(7000)

    def fetch_category(
        self,
        page,
        keyword: str,
        category_code: str,
        category_name: str,
        page_no: int,
    ) -> CategoryResult:
        payload = {
            "pageNo": page_no,
            "pageSize": PAGE_SIZE,
            "categoryCode": category_code,
            "keyword": keyword,
            "publishDateBegin": START_DATE,
            "publishDateEnd": END_DATE,
            "_t": now_ms(),
        }

        last_text = ""
        for attempt in range(1, LIST_RETRIES + 1):
            logging.info("请求列表：公告类型=%s，关键词=%s，第 %d 页，第 %d/%d 次",
                         category_name, keyword, page_no, attempt, LIST_RETRIES)
            try:
                data = page.evaluate(
                    """async ({payload, timeoutMs}) => {
                        const controller = new AbortController();
                        const timer = setTimeout(() => controller.abort(), timeoutMs);
                        try {
                            const res = await fetch('/portal/category', {
                                method: 'POST',
                                headers: {
                                    'Accept': 'application/json, text/plain, */*',
                                    'Content-Type': 'application/json;charset=UTF-8',
                                    'X-Requested-With': 'XMLHttpRequest'
                                },
                                body: JSON.stringify(payload),
                                signal: controller.signal
                            });
                            const text = await res.text();
                            clearTimeout(timer);
                            try {
                                return {
                                    ok: true,
                                    status: res.status,
                                    json: JSON.parse(text),
                                    text: text.slice(0, 1000),
                                    aborted: false
                                };
                            } catch (e) {
                                return {
                                    ok: false,
                                    status: res.status,
                                    json: null,
                                    text: text.slice(0, 1200),
                                    aborted: false,
                                    parseError: String(e)
                                };
                            }
                        } catch (e) {
                            clearTimeout(timer);
                            return {
                                ok: false,
                                status: 0,
                                json: null,
                                text: "",
                                aborted: String(e).includes("AbortError"),
                                error: String(e)
                            };
                        }
                    }""",
                    {"payload": payload, "timeoutMs": API_FETCH_TIMEOUT_MS},
                )

                last_text = data.get("text") or ""
                if data.get("ok") and isinstance(data.get("json"), dict):
                    js = data["json"]
                    if js.get("success") is True:
                        block = (((js.get("result") or {}).get("data") or {}))
                        total = int(block.get("total") or 0)
                        rows = block.get("data") or []
                        if not isinstance(rows, list):
                            rows = []
                        return CategoryResult(total=total, rows=rows, raw=js)

                    logging.warning("接口 JSON success 非 true：%s", str(js)[:300])
                else:
                    if data.get("aborted"):
                        logging.warning("列表 fetch 超时中断，status=%s，错误=%s",
                                        data.get("status"), data.get("error"))
                    else:
                        logging.warning("列表返回不是 JSON，status=%s，前 300 字：%s",
                                        data.get("status"), last_text[:300].replace("\n", " "))

                if "aliyun_waf" in last_text or "renderData" in last_text or "window._waf" in last_text:
                    logging.warning("列表接口触发 WAF，重新打开栏目页。")
                    self.open_home_and_pass_waf(page)

            except Exception as e:
                logging.warning("列表请求异常：%s", e)

            page.wait_for_timeout(int((1500 + attempt * 1500) + random.randint(0, 1000)))

        raw_path = RAW_DIR / f"category_fail_{category_code}_{keyword}_p{page_no}.txt"
        raw_path.write_text(last_text, encoding="utf-8", errors="ignore")
        try:
            shot_path = CAPTURE_DIR / f"category_fail_{category_code}_{keyword}_p{page_no}.png"
            page.screenshot(path=str(shot_path), full_page=True)
            logging.error("失败截图已保存：%s", shot_path)
        except Exception:
            pass
        raise RuntimeError(f"列表请求失败：{category_name}/{keyword}/page={page_no}，已保存 {raw_path}")

    def goto_detail_and_parse(self, page, detail_url: str, row: Dict[str, Any]) -> Tuple[Dict[str, str], str]:
        """
        返回 detail_fields, status
        即使详情失败，也返回空字段，主流程仍保留列表字段。
        """
        html = ""
        body_text = ""

        for attempt in range(1, DETAIL_RETRIES + 1):
            try:
                page.goto(detail_url, wait_until="domcontentloaded", timeout=DETAIL_TIMEOUT_MS)
                page.wait_for_timeout(1800)

                # 尽量等待政采云组件渲染
                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                except Exception:
                    pass

                html = page.content()
                try:
                    body_text = page.locator("body").inner_text(timeout=8000)
                except Exception:
                    body_text = ""

                if "aliyun_waf" in html or "renderData" in html and len(body_text) < 200:
                    logging.warning("详情页疑似 WAF/未渲染，第 %d/%d 次：%s", attempt, DETAIL_RETRIES, detail_url)
                    page.wait_for_timeout(3000)
                    continue

                # 内容过短时仍然保存 raw，但不视为致命错误
                fields = parse_detail_html(html, body_text)
                return fields, "成功"

            except Exception as e:
                logging.warning("详情请求失败 %d/%d：%s；错误：%s", attempt, DETAIL_RETRIES, detail_url, e)
                page.wait_for_timeout(int(1200 + attempt * 1200))

        safe_name = re.sub(r"[^0-9A-Za-z_\-\u4e00-\u9fa5]+", "_", (row.get("articleId") or "detail"))[:80]
        raw_path = RAW_DIR / f"detail_fail_{safe_name}.html"
        try:
            raw_path.write_text(html or body_text or "", encoding="utf-8", errors="ignore")
        except Exception:
            pass
        return {}, f"详情失败，raw={raw_path.name}"

    def build_detail_url(self, article_id: str) -> str:
        return DETAIL_URL_TEMPLATE.format(article_id=quote(article_id or "", safe=""))

    def row_to_record_base(
        self,
        row: Dict[str, Any],
        keyword: str,
        category_name: str,
        category_code: str,
    ) -> Dict[str, Any]:
        article_id = row.get("articleId") or ""
        detail_url = self.build_detail_url(article_id)

        return {
            "province": "新疆",
            "keyword": keyword,
            "category": category_name,
            "category_code": category_code,
            "announcement_type": row.get("pathName") or category_name,
            "title": clean_html_text(row.get("title") or ""),
            "project_name": row.get("projectName") or "",
            "project_code": row.get("projectCode") or "",
            "purchase_name": row.get("purchaseName") or "",
            "agency_name": row.get("author") or "",
            "supplier_name": row.get("supplierName") or "",
            "budget_amount": amount_to_yuan(row.get("budgetPrice")),
            "win_amount": amount_to_yuan(row.get("totalContractAmount")),
            "procurement_method": row.get("procurementMethod") or "",
            "gp_catalog_name": row.get("gpCatalogName") or "",
            "district_name": row.get("districtName") or "",
            "publish_time": ms_to_date(row.get("publishDate")),
            "bid_opening_time": ms_to_date(row.get("bidOpeningTime")),
            "article_id": article_id,
            "detail_url": detail_url,
            "crawl_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "crawl_status": "列表成功",
            "error": "",
            "summary": clean_html_text(row.get("content") or ""),
        }

    def merge_detail_fields(self, base: Dict[str, Any], detail: Dict[str, str]) -> Dict[str, Any]:
        if detail.get("detail_title"):
            base["title"] = detail["detail_title"]

        if detail.get("project_code_detail"):
            base["project_code"] = detail["project_code_detail"]

        if detail.get("project_name_detail"):
            base["project_name"] = detail["project_name_detail"]

        if detail.get("purchase_name_detail"):
            base["purchase_name"] = detail["purchase_name_detail"]

        if detail.get("agency_name_detail"):
            base["agency_name"] = detail["agency_name_detail"]

        if detail.get("supplier_name_detail"):
            base["supplier_name"] = detail["supplier_name_detail"]

        # 预算金额：优先接口，接口没有再用详情
        if not base.get("budget_amount") and detail.get("budget_amount_detail"):
            base["budget_amount"] = detail["budget_amount_detail"]

        # 中标/成交金额：优先详情，因为新疆结果公告接口 totalContractAmount 常为空
        if detail.get("win_amount_detail"):
            base["win_amount"] = detail["win_amount_detail"]

        if detail.get("publish_time_detail") and not base.get("publish_time"):
            base["publish_time"] = detail["publish_time_detail"]

        if detail.get("summary"):
            base["summary"] = detail["summary"]

        return base

    def crawl(self) -> List[Dict[str, Any]]:
        ensure_dirs()

        with sync_playwright() as p:
            context = self._launch_context(p)
            page = context.new_page()
            self.open_home_and_pass_waf(page)

            try:
                for keyword in KEYWORDS:
                    for task in CATEGORY_TASKS:
                        category_name = task["name"]
                        category_code = task["categoryCode"]

                        # 第一页拿 total
                        result = self.fetch_category(page, keyword, category_code, category_name, 1)
                        total_pages = max(1, math.ceil(result.total / PAGE_SIZE)) if result.total else 1
                        logging.info("公告类型=%s，关键词=%s，总数=%s，预计页数=%s",
                                     category_name, keyword, result.total, total_pages)

                        for page_no in range(1, total_pages + 1):
                            if page_no == 1:
                                rows = result.rows
                            else:
                                rows = self.fetch_category(page, keyword, category_code, category_name, page_no).rows

                            logging.info("处理列表：%s / %s / 第 %d 页，%d 条",
                                         category_name, keyword, page_no, len(rows))

                            for idx, row in enumerate(rows, 1):
                                title_for_log = clean_html_text(row.get("title") or "")
                                if STRICT_TITLE_FILTER and keyword not in title_for_log:
                                    logging.info("跳过标题不含关键词：%s", title_for_log)
                                    continue

                                base = self.row_to_record_base(row, keyword, category_name, category_code)
                                detail_url = base["detail_url"]

                                if SKIP_VISITED and detail_url in self.visited:
                                    logging.info("跳过已访问：%s", detail_url)
                                    continue

                                logging.info("详情 [%s 第%d页 %d/%d]：%s",
                                             category_name, page_no, idx, len(rows), title_for_log or detail_url)

                                detail_fields, status = self.goto_detail_and_parse(page, detail_url, row)
                                record = self.merge_detail_fields(base, detail_fields)
                                record["crawl_status"] = status
                                if status != "成功":
                                    record["error"] = status

                                write_jsonl(record)
                                self.records_this_run.append(record)

                                if detail_url:
                                    save_visited(detail_url)
                                    self.visited.add(detail_url)

                                sleep_jitter(0.7, 1.5)

                            sleep_jitter(0.8, 1.8)

            finally:
                try:
                    context.close()
                except Exception:
                    pass

        return self.records_this_run


# =========================
# 5. 主函数
# =========================

def main() -> None:
    setup_logging()
    logging.info("开始爬取新疆政府采购网（Playwright 后台浏览器版）")
    logging.info("关键词：%s；时间：%s 至 %s；公告类型：%s",
                 KEYWORDS, START_DATE, END_DATE, [x["name"] for x in CATEGORY_TASKS])
    logging.info("HEADLESS=%s；浏览器缓存目录=%s；列表接口超时=%sms", HEADLESS, BROWSER_USER_DATA_DIR, API_FETCH_TIMEOUT_MS)

    spider = XinjiangBrowserSpider()

    try:
        records = spider.crawl()
        logging.info("本次运行新增记录：%d 条", len(records))
        export_outputs(records)
    except KeyboardInterrupt:
        logging.warning("用户中断，正在导出已抓取数据。")
        export_outputs(spider.records_this_run)
    except Exception as e:
        logging.exception("运行失败，正在导出已抓取数据：%s", e)
        export_outputs(spider.records_this_run)
        raise


if __name__ == "__main__":
    main()
