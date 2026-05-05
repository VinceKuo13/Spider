# -*- coding: utf-8 -*-
"""
招标网历史标讯爬虫

功能：
1. Selenium 打开浏览器，由用户手动完成图形验证码登录。
2. 自动读取登录后的 Cookie。
3. 直接 POST 请求 https://center.zhaobiao.cn/www/hallIndex/hisAjax 获取历史标讯列表。
4. 按 年/月 × 地区 × 关键词 × 公告类型 × 页码 批量爬取。
5. 进入详情页解析：项目名称、招标单位、中标单位、采购形式、代理机构、时间、金额、链接等。
6. 保存 JSONL checkpoint 和 Excel。

说明：
- 第一版采用公告级别输出，每条详情页一行。
- 后续可基于 项目编号 + 项目名称相似度 + 招标进度链接 做“项目级合并”。
"""

from __future__ import annotations

import json
import logging
import os
import random
import re
import sys
import time
from dataclasses import dataclass, asdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests
import yaml
from bs4 import BeautifulSoup
from requests import Session


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.yaml"


@dataclass
class NoticeRecord:
    search_keyword: str = ""
    search_province: str = ""
    search_province_code: str = ""
    search_year: str = ""
    search_month: str = ""
    notice_type: str = ""
    notice_type_code: str = ""

    project_name: str = ""
    project_no: str = ""
    tender_unit: str = ""
    winner_unit: str = ""
    purchase_type: str = ""
    agency: str = ""
    publish_time: str = ""
    amount: str = ""
    detail_url: str = ""

    source_region: str = ""
    raw_title: str = ""
    notice_id: str = ""
    enc_id: str = ""
    raw_text_summary: str = ""
    parse_status: str = ""
    error: str = ""


FIELD_ALIASES = {
    "project_no": ["项目编号", "招标编号", "采购编号", "项目编码", "编号"],
    "project_name": ["项目名称", "采购项目名称", "标项名称", "标的名称"],
    "tender_unit": ["招标单位", "招标人", "采购人", "采购单位", "采购人信息", "采购单位名称", "建设单位", "业主单位"],
    "winner_unit": ["中标单位", "中标人", "成交供应商", "成交单位", "供应商名称", "中标候选", "中标供应商", "成交人"],
    "agency": ["代理机构", "采购代理机构", "代理机构名称", "招标代理", "采购代理", "代理单位"],
    "amount": ["预算金额", "采购预算", "最高限价", "中标金额", "成交金额", "中标（成交）金额", "中标(成交)金额", "总价", "报价", "金额"],
    "purchase_type": ["采购方式", "招标方式", "采购形式", "招标形式"],
}

PURCHASE_TYPE_KEYWORDS = [
    "公开招标", "邀请招标", "竞争性磋商", "竞争性谈判", "询价", "比选", "单一来源", "框架协议",
    "公开询价", "询比价", "谈判", "磋商", "竞价", "遴选", "比价",
]

NOTICE_SUFFIX_PATTERNS = [
    r"公开招标公告$", r"竞争性磋商公告$", r"竞争性谈判公告$", r"询价公告$", r"比选公告$",
    r"采购公告$", r"招标公告$", r"中标公告$", r"成交公告$", r"结果公告$", r"谈判结果公告$",
    r"中标\（成交\）公告$", r"中标\(成交\)公告$", r"的公告$", r"公告$",
]


def load_config(path: Path = CONFIG_PATH) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"配置文件不存在：{path}")
    with path.open("r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    return cfg


def setup_logging(cfg: Dict[str, Any]) -> None:
    log_file = BASE_DIR / cfg.get("log_file", "logs/zhaobiao_history.log")
    log_file.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def month_range(start: date, end: date) -> List[Tuple[str, str]]:
    months: List[Tuple[str, str]] = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        months.append((str(y), f"{m:02d}"))
        if m == 12:
            y += 1
            m = 1
        else:
            m += 1
    return months


def normalize_space(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[\t\r\f\v]+", " ", text)
    text = re.sub(r"[ ]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def safe_filename_part(s: str, max_len: int = 80) -> str:
    s = re.sub(r"[\\/:*?\"<>|\s]+", "_", s or "")
    return s[:max_len].strip("_")


def build_headers() -> Dict[str, str]:
    return {
        "accept": "application/json, text/javascript, */*; q=0.01",
        "accept-language": "zh-CN,zh;q=0.9",
        "content-type": "application/json",
        "origin": "https://center.zhaobiao.cn",
        "referer": "https://center.zhaobiao.cn/www/hallIndex/historyOfBidding",
        "user-agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"
        ),
        "x-requested-with": "XMLHttpRequest",
    }


class LoginManager:
    def __init__(self, cfg: Dict[str, Any]):
        self.cfg = cfg
        self.driver: Optional[webdriver.Chrome] = None

    def start_driver(self) -> webdriver.Chrome:
        if self.driver is not None:
            return self.driver
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.options import Options
        except ImportError as e:
            raise RuntimeError("缺少 selenium，请先执行：pip install -r requirements.txt") from e

        options = Options()
        if self.cfg.get("headless", False):
            options.add_argument("--headless=new")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--start-maximized")
        self.driver = webdriver.Chrome(options=options)
        return self.driver

    def manual_login_and_get_session(self) -> Session:
        driver = self.start_driver()
        login_url = self.cfg.get("login_url") or "https://user.zhaobiao.cn/homePageUc.do"
        history_url = self.cfg.get("history_url") or "https://center.zhaobiao.cn/www/hallIndex/historyOfBidding"

        print("\n即将打开招标网登录/会员中心页面。")
        print("请在浏览器里手动完成登录和图形验证码。")
        print("登录成功后，建议进入一次『招标类历史数据查询』页面，然后回到终端按回车。\n")
        driver.get(login_url)
        input("完成登录后请按回车继续... ")

        # 主动访问历史查询页，让 center.zhaobiao.cn 域名下 Cookie 完整产生。
        try:
            driver.get(history_url)
            time.sleep(2)
        except Exception as e:
            logging.warning("访问历史查询页失败，但仍尝试读取 Cookie：%s", e)

        session = requests.Session()
        session.headers.update(build_headers())
        self.copy_cookies_to_session(session)
        return session

    def relogin(self, session: Session) -> None:
        driver = self.start_driver()
        history_url = self.cfg.get("history_url") or "https://center.zhaobiao.cn/www/hallIndex/historyOfBidding"
        print("\n检测到请求失败、Cookie 失效或需要人工验证。")
        print("请在已打开的浏览器中重新登录/完成验证，并进入历史数据查询页面。")
        try:
            driver.get(history_url)
        except Exception:
            pass
        input("处理完成后按回车继续... ")
        session.cookies.clear()
        self.copy_cookies_to_session(session)

    def copy_cookies_to_session(self, session: Session) -> None:
        if self.driver is None:
            return
        # 尽量访问几个相关域名，让 Selenium 能读到对应域的 cookie。
        for url in [
            "https://user.zhaobiao.cn/homePageUc.do",
            "https://center.zhaobiao.cn/www/hallIndex/historyOfBidding",
            "https://zb.zhaobiao.cn/",
        ]:
            try:
                self.driver.get(url)
                time.sleep(0.8)
                for c in self.driver.get_cookies():
                    name = c.get("name")
                    value = c.get("value")
                    domain = c.get("domain")
                    path = c.get("path", "/")
                    if name and value is not None:
                        try:
                            session.cookies.set(name, value, domain=domain, path=path)
                        except Exception:
                            session.cookies.set(name, value)
            except Exception as e:
                logging.debug("读取 Cookie 时访问 %s 失败：%s", url, e)

    def close(self) -> None:
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None


def request_list_page(
    session: Session,
    cfg: Dict[str, Any],
    keyword: str,
    notice_type_code: str,
    prov_code: str,
    year: str,
    month: str,
    page: int,
) -> Dict[str, Any]:
    payload = {
        "queryWord": keyword,
        "newType": notice_type_code,
        "provCode": prov_code,
        "hisYear": year,
        "hisMonth": month,
        "field": cfg.get("field", "all"),
        "fileSearch": cfg.get("file_search", "all"),
        "currpage": str(page),
    }
    api_url = cfg.get("api_url") or "https://center.zhaobiao.cn/www/hallIndex/hisAjax"
    resp = session.post(api_url, json=payload, timeout=30)
    resp.raise_for_status()
    try:
        data = resp.json()
    except Exception as e:
        raise RuntimeError(f"列表接口返回不是 JSON，可能登录失效。响应前200字：{resp.text[:200]}") from e
    return data


def is_login_or_auth_failed(data: Dict[str, Any]) -> bool:
    if not isinstance(data, dict):
        return True
    msg = str(data.get("msg", ""))
    status = data.get("status")
    if status in (401, 403):
        return True
    bad_words = ["登录", "验证码", "无权限", "未授权", "请升级", "重新登录"]
    return any(w in msg for w in bad_words)


def get_soup(session: Session, url: str, timeout: int = 30) -> BeautifulSoup:
    """请求详情页并返回 BeautifulSoup。

    关键改动：遇到 403/429/500/502/503/520/521/522 等疑似反爬或服务异常时，
    先重试；重试仍失败则抛出异常。上层捕获后不会保存数据，也不会写入 visited_urls.txt。
    """
    headers = {
        "user-agent": build_headers()["user-agent"],
        "referer": "https://center.zhaobiao.cn/www/hallIndex/historyOfBidding",
    }

    max_retries = 4
    retry_status_codes = {403, 408, 429, 500, 502, 503, 520, 521, 522, 523, 524}
    wait_seconds = [3, 8, 20, 60]
    last_error: Optional[Exception] = None

    for attempt in range(1, max_retries + 1):
        try:
            resp = session.get(url, headers=headers, timeout=timeout)

            # 这些状态码经常代表限流、反爬、Cloudflare/服务器临时异常。
            if resp.status_code in retry_status_codes:
                last_error = RuntimeError(f"{resp.status_code} Server Error")
                if attempt < max_retries:
                    sleep_time = wait_seconds[min(attempt - 1, len(wait_seconds) - 1)]
                    logging.warning(
                        "详情页异常 %s，第 %s/%s 次重试，等待 %s 秒：%s",
                        resp.status_code,
                        attempt,
                        max_retries,
                        sleep_time,
                        url,
                    )
                    time.sleep(sleep_time)
                    continue
                raise last_error

            resp.raise_for_status()

            # requests 通常能识别 UTF-8；这里再兜底。
            if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
                resp.encoding = resp.apparent_encoding or "utf-8"

            html = resp.text or ""
            # 登录失效/验证码/权限页通常不是正常详情页。这里做一个轻量检查。
            suspicious_words = ["验证码", "请登录", "重新登录", "无权限", "访问过于频繁", "安全验证"]
            if any(w in html for w in suspicious_words) and "#infotitle" not in html and "w-detailTit" not in html:
                last_error = RuntimeError("详情页疑似登录失效、验证码或风控页面")
                if attempt < max_retries:
                    sleep_time = wait_seconds[min(attempt - 1, len(wait_seconds) - 1)]
                    logging.warning(
                        "详情页疑似风控/登录页，第 %s/%s 次重试，等待 %s 秒：%s",
                        attempt,
                        max_retries,
                        sleep_time,
                        url,
                    )
                    time.sleep(sleep_time)
                    continue
                raise last_error

            return BeautifulSoup(html, "lxml")

        except Exception as e:
            last_error = e
            if attempt < max_retries:
                sleep_time = wait_seconds[min(attempt - 1, len(wait_seconds) - 1)]
                logging.warning(
                    "详情页请求失败，第 %s/%s 次重试，等待 %s 秒：%s | %s",
                    attempt,
                    max_retries,
                    sleep_time,
                    url,
                    e,
                )
                time.sleep(sleep_time)
                continue
            raise last_error

    raise last_error or RuntimeError(f"详情页请求失败：{url}")


def collect_overview_fields(soup: BeautifulSoup) -> Dict[str, List[str]]:
    result: Dict[str, List[str]] = {}
    container = soup.select_one("#overview .new-noticeInfo") or soup.select_one(".new-noticeInfo")
    if not container:
        return result
    for label_div in container.select(".label"):
        label = normalize_space(label_div.get_text(" ", strip=True)).strip("：:")
        if not label:
            continue
        value_div = label_div.find_next_sibling(class_="value")
        if not value_div:
            continue
        value = normalize_space(value_div.get("title") or value_div.get_text(" ", strip=True))
        if value:
            result.setdefault(label, []).append(value)
    return result


def overview_get(fields: Dict[str, List[str]], aliases: List[str]) -> str:
    vals: List[str] = []
    for k, vlist in fields.items():
        kk = k.replace(" ", "").strip("：:")
        if any(alias.replace(" ", "") in kk for alias in aliases):
            vals.extend(vlist)
    return "；".join(unique_keep_order(vals))


def unique_keep_order(items: Iterable[str]) -> List[str]:
    seen = set()
    out = []
    for item in items:
        item = normalize_space(str(item))
        if item and item not in seen:
            seen.add(item)
            out.append(item)
    return out


def extract_text(soup: BeautifulSoup) -> str:
    cont = soup.select_one("#cont") or soup.select_one(".bid_details_zw") or soup.body
    text = cont.get_text("\n", strip=True) if cont else soup.get_text("\n", strip=True)
    return normalize_space(text)


def extract_summary(text: str, max_len: int = 500) -> str:
    return normalize_space(text).replace("\n", " ")[:max_len]


def clean_value(value: str) -> str:
    value = normalize_space(value)
    value = re.sub(r"^[：:\-—\s]+", "", value)
    value = re.sub(r"[。；;，,\s]+$", "", value)
    return value.strip()


def line_based_extract(text: str, aliases: List[str], max_len: int = 120) -> str:
    lines = [normalize_space(x) for x in text.split("\n") if normalize_space(x)]
    # 1. 同行：采购人：xxx
    for line in lines:
        compact = line.replace(" ", "")
        for alias in aliases:
            a = alias.replace(" ", "")
            if a in compact:
                # 排除章节标题，如“九、对本次公告内容提出询问，请按以下方式联系”
                m = re.search(re.escape(alias) + r"\s*[：:]\s*(.+)$", line)
                if m:
                    val = clean_value(m.group(1))
                    if 0 < len(val) <= max_len:
                        return val
                # “采购人信息”这类章节，下一行常是“名 称：xxx”
                idx = lines.index(line)
                for nxt in lines[idx + 1: idx + 5]:
                    m2 = re.search(r"(?:名称|名\s*称|单位名称)\s*[：:]\s*(.+)$", nxt)
                    if m2:
                        val = clean_value(m2.group(1))
                        if 0 < len(val) <= max_len:
                            return val
    return ""


def regex_extract_first(text: str, aliases: List[str], max_len: int = 160) -> str:
    for alias in aliases:
        pattern = re.escape(alias) + r"\s*[：:]\s*([^\n。；;]{1," + str(max_len) + r"})"
        m = re.search(pattern, text)
        if m:
            return clean_value(m.group(1))
    return ""


def extract_project_name(text: str, title: str, overview_fields: Dict[str, List[str]]) -> str:
    val = overview_get(overview_fields, FIELD_ALIASES["project_name"])
    if val:
        return val
    val = regex_extract_first(text, FIELD_ALIASES["project_name"], max_len=160)
    if val:
        return val
    val = title or ""
    for pat in NOTICE_SUFFIX_PATTERNS:
        val = re.sub(pat, "", val)
    return clean_value(val) or title


def extract_project_no(text: str, overview_fields: Dict[str, List[str]]) -> str:
    val = overview_get(overview_fields, FIELD_ALIASES["project_no"])
    if val:
        return val
    return regex_extract_first(text, FIELD_ALIASES["project_no"], max_len=80)


def extract_purchase_type(text: str, title: str, overview_fields: Dict[str, List[str]]) -> str:
    val = overview_get(overview_fields, FIELD_ALIASES["purchase_type"])
    if val:
        return val
    val = regex_extract_first(text, FIELD_ALIASES["purchase_type"], max_len=60)
    if val:
        return val
    hay = f"{title}\n{text[:1000]}"
    found = [kw for kw in PURCHASE_TYPE_KEYWORDS if kw in hay]
    return "；".join(unique_keep_order(found[:3]))


def table_extract_by_header(soup: BeautifulSoup, header_aliases: List[str]) -> str:
    vals: List[str] = []
    for table in soup.select("#cont table, .bid_details_zw table, table"):
        rows = []
        for tr in table.select("tr"):
            cells = [normalize_space(td.get_text(" ", strip=True)) for td in tr.find_all(["th", "td"])]
            cells = [c for c in cells if c]
            if cells:
                rows.append(cells)
        if len(rows) < 2:
            continue
        header = rows[0]
        target_indices = []
        for i, h in enumerate(header):
            h_compact = h.replace(" ", "")
            if any(alias.replace(" ", "") in h_compact for alias in header_aliases):
                target_indices.append(i)
        if not target_indices:
            continue
        for row in rows[1:4]:
            for idx in target_indices:
                if idx < len(row) and row[idx]:
                    vals.append(row[idx])
    return "；".join(unique_keep_order(vals))


def extract_winner_unit(soup: BeautifulSoup, text: str, overview_fields: Dict[str, List[str]]) -> str:
    val = overview_get(overview_fields, FIELD_ALIASES["winner_unit"])
    if val:
        return val
    val = table_extract_by_header(soup, FIELD_ALIASES["winner_unit"])
    if val:
        return val
    val = regex_extract_first(text, FIELD_ALIASES["winner_unit"], max_len=160)
    if val:
        return val
    return line_based_extract(text, FIELD_ALIASES["winner_unit"], max_len=160)


def extract_amount(soup: BeautifulSoup, text: str, overview_fields: Dict[str, List[str]]) -> str:
    val = overview_get(overview_fields, FIELD_ALIASES["amount"])
    if val:
        return val
    val = table_extract_by_header(soup, FIELD_ALIASES["amount"])
    if val:
        return val
    # 优先明确字段
    for alias in FIELD_ALIASES["amount"]:
        pattern = re.escape(alias) + r"\s*[：:]\s*([^\n]{1,80}(?:元|万元|人民币|%|折扣|人餐|/人餐)?)"
        m = re.search(pattern, text)
        if m:
            return clean_value(m.group(1))
    # 兜底金额
    m = re.search(r"(?:总价|报价)\s*[：:]\s*([^\n]{1,80})", text)
    if m:
        return clean_value(m.group(1))
    return ""


def parse_detail_page(session: Session, url: str, base_record: NoticeRecord) -> NoticeRecord:
    try:
        soup = get_soup(session, url)
        title = ""
        title_el = soup.select_one("#infotitle")
        if title_el:
            title = normalize_space(title_el.get_text(" ", strip=True))
        if not title:
            hidden_title = soup.select_one("input#title")
            title = hidden_title.get("value", "") if hidden_title else ""
        if not title:
            title = base_record.raw_title

        publish_time = base_record.publish_time
        hidden_date = soup.select_one("input#publishDate")
        if hidden_date and hidden_date.get("value"):
            publish_time = hidden_date.get("value")

        overview_fields = collect_overview_fields(soup)
        text = extract_text(soup)

        project_name = extract_project_name(text, title, overview_fields)
        project_no = extract_project_no(text, overview_fields)
        tender_unit = overview_get(overview_fields, FIELD_ALIASES["tender_unit"])
        if not tender_unit:
            tender_unit = line_based_extract(text, FIELD_ALIASES["tender_unit"], max_len=160)
        agency = overview_get(overview_fields, FIELD_ALIASES["agency"])
        if not agency:
            agency = line_based_extract(text, FIELD_ALIASES["agency"], max_len=160)
        winner_unit = extract_winner_unit(soup, text, overview_fields)
        amount = extract_amount(soup, text, overview_fields)
        purchase_type = extract_purchase_type(text, title, overview_fields)

        base_record.project_name = project_name
        base_record.project_no = project_no
        base_record.tender_unit = tender_unit
        base_record.winner_unit = winner_unit
        base_record.purchase_type = purchase_type
        base_record.agency = agency
        base_record.publish_time = publish_time
        base_record.raw_title = title or base_record.raw_title
        base_record.raw_text_summary = extract_summary(text)
        base_record.parse_status = "ok"
        return base_record
    except Exception as e:
        # 关键改动：详情页失败时直接抛出异常。
        # 主循环会捕获该异常，并且不会保存 jsonl、不会写入 visited_urls.txt、不会推进页码。
        logging.warning("详情页解析失败：%s | %s", url, e)
        raise


def date_in_range(date_str: str, start: date, end: date) -> bool:
    try:
        d = parse_date(date_str)
        return start <= d <= end
    except Exception:
        return True


def province_match(source_region: str, search_province: str) -> bool:
    if not source_region:
        return True
    return source_region.startswith(search_province)


def append_jsonl(path: Path, record: NoticeRecord) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(asdict(record), ensure_ascii=False) + "\n")


def load_visited(path: Path) -> set:
    if not path.exists():
        return set()
    with path.open("r", encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def save_visited(path: Path, url: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(url + "\n")


def task_key(keyword: str, year: str, month: str, province: str, notice_type: str) -> str:
    """生成断点任务键。

    第一版的实际运行顺序是：关键词 -> 年月 -> 省份 -> 公告类型。
    因为一个“关键词+年月+省份”下面需要同时爬招标公告和中标公告，
    所以最小可恢复任务粒度设置为：关键词+年月+省份+公告类型。
    """
    return "||".join([keyword, year, month, province, notice_type])


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


# def save_progress(path: Path, progress: Dict[str, Any]) -> None:
#     path.parent.mkdir(parents=True, exist_ok=True)
#     progress["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#     tmp_path = path.with_suffix(path.suffix + ".tmp")
#     with tmp_path.open("w", encoding="utf-8") as f:
#         json.dump(progress, f, ensure_ascii=False, indent=2)
#     tmp_path.replace(path)

def save_progress(path: Path, progress: Dict[str, Any]) -> None:
    """
    保存断点文件。

    Windows 下 os.replace / Path.replace 偶尔会因为文件被编辑器、杀毒软件、
    同步软件短暂占用而报 WinError 5。这里增加重试和兜底备份，避免程序直接崩溃。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    progress["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    content = json.dumps(progress, ensure_ascii=False, indent=2)

    tmp_path = path.with_suffix(path.suffix + ".tmp")

    # 先写临时文件
    with tmp_path.open("w", encoding="utf-8") as f:
        f.write(content)

    # 多次尝试原子替换
    last_error = None
    for i in range(10):
        try:
            tmp_path.replace(path)
            return
        except PermissionError as e:
            last_error = e
            wait = 0.2 + i * 0.3
            logging.warning(
                "断点文件被占用，等待 %.1f 秒后重试保存：%s",
                wait,
                path,
            )
            time.sleep(wait)
        except Exception as e:
            last_error = e
            wait = 0.2 + i * 0.3
            logging.warning(
                "断点文件保存失败，等待 %.1f 秒后重试：%s | %s",
                wait,
                path,
                e,
            )
            time.sleep(wait)

    # 如果替换一直失败，尝试直接写入正式文件
    try:
        with path.open("w", encoding="utf-8") as f:
            f.write(content)
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except Exception:
            pass
        logging.warning("断点文件通过直接写入方式保存成功：%s", path)
        return
    except Exception as e:
        last_error = e

    # 最后兜底：保存一个 recovery 文件，避免断点信息完全丢失
    recovery_path = path.with_name(
        f"{path.stem}_recovery_{datetime.now().strftime('%Y%m%d_%H%M%S')}{path.suffix}"
    )
    with recovery_path.open("w", encoding="utf-8") as f:
        f.write(content)

    logging.error(
        "正式断点文件保存失败，已另存为恢复文件：%s。原错误：%s",
        recovery_path,
        last_error,
    )


def is_task_completed(progress: Dict[str, Any], key: str) -> bool:
    return key in set(progress.get("completed_tasks", []))


def get_resume_page(progress: Dict[str, Any], key: str) -> int:
    current = progress.get("current") or {}
    if current.get("task_key") == key:
        try:
            return max(1, int(current.get("next_page", 1)))
        except Exception:
            return 1
    return 1


def set_current_progress(
    path: Path,
    progress: Dict[str, Any],
    key: str,
    next_page: int,
    meta: Dict[str, Any],
) -> None:
    progress["current"] = {
        "task_key": key,
        "next_page": int(next_page),
        **meta,
    }
    save_progress(path, progress)


def mark_task_completed(path: Path, progress: Dict[str, Any], key: str) -> None:
    completed = list(progress.get("completed_tasks", []))
    if key not in completed:
        completed.append(key)
    progress["completed_tasks"] = completed
    current = progress.get("current") or {}
    if current.get("task_key") == key:
        progress["current"] = None
    save_progress(path, progress)


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


def export_excel(cfg: Dict[str, Any]) -> None:
    jsonl_path = BASE_DIR / cfg.get("checkpoint_jsonl", "outputs/招标网_历史标讯_明细.jsonl")
    excel_path = BASE_DIR / cfg.get("output_excel", "outputs/招标网_历史标讯_公告级别.xlsx")
    rows = read_jsonl_records(jsonl_path)
    if not rows:
        logging.warning("没有可导出的记录。")
        return
    # URL 去重，保留最后一次解析结果。
    by_url: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        url = row.get("detail_url") or row.get("enc_id") or str(len(by_url))
        by_url[url] = row
    df = pd.DataFrame(list(by_url.values()))
    column_map = {
        "search_keyword": "检索关键词",
        "search_province": "检索省份",
        "search_province_code": "检索省份代码",
        "search_year": "检索年份",
        "search_month": "检索月份",
        "notice_type": "公告类型",
        "notice_type_code": "公告类型代码",
        "project_name": "项目名称",
        "project_no": "项目编号",
        "tender_unit": "招标单位/采购人",
        "winner_unit": "中标单位/成交供应商",
        "purchase_type": "采购形式/采购方式",
        "agency": "代理机构",
        "publish_time": "时间/发布时间",
        "amount": "金额",
        "detail_url": "链接",
        "source_region": "来源地区",
        "raw_title": "原始标题",
        "notice_id": "公告ID",
        "enc_id": "公告encId",
        "raw_text_summary": "正文摘要",
        "parse_status": "解析状态",
        "error": "错误信息",
    }
    desired_cols = list(column_map.keys())
    for c in desired_cols:
        if c not in df.columns:
            df[c] = ""
    df = df[desired_cols].rename(columns=column_map)
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(excel_path, index=False)
    logging.info("Excel 已导出：%s，共 %d 条。", excel_path, len(df))


def sleep_jitter(seconds: float) -> None:
    if seconds <= 0:
        return
    time.sleep(seconds + random.uniform(0, min(0.8, seconds * 0.5)))


def crawl(cfg: Dict[str, Any]) -> None:
    setup_logging(cfg)
    start = parse_date(cfg.get("start_date", "2025-01-01"))
    end_str = cfg.get("end_date")
    end = parse_date(end_str) if end_str else date.today()
    months = month_range(start, end)

    jsonl_path = BASE_DIR / cfg.get("checkpoint_jsonl", "outputs/招标网_历史标讯_明细.jsonl")
    visited_path = BASE_DIR / cfg.get("visited_urls_file", "outputs/visited_urls.txt")
    progress_path = BASE_DIR / cfg.get("progress_state_file", "outputs/progress_state.json")
    visited = load_visited(visited_path)
    progress = load_progress(progress_path)

    logging.info("断点文件：%s", progress_path)
    if progress.get("current"):
        logging.info("检测到未完成任务，将尝试从断点继续：%s", progress.get("current"))
    logging.info("已完成任务数：%d", len(progress.get("completed_tasks", [])))

    login_mgr = LoginManager(cfg)
    session = login_mgr.manual_login_and_get_session()

    provinces: Dict[str, str] = cfg.get("provinces", {})
    keywords: List[str] = cfg.get("keywords", [])
    notice_types: Dict[str, str] = cfg.get("notice_types", {})
    max_pages = int(cfg.get("max_pages_per_query", 100) or 100)
    list_delay = float(cfg.get("request_delay", 1.2) or 0)
    detail_delay = float(cfg.get("detail_delay", 0.8) or 0)
    manual_relogin = bool(cfg.get("manual_relogin_on_fail", True))
    auto_export_every_tasks = int(cfg.get("auto_export_every_tasks", 10) or 0)
    completed_since_export = 0

    total_combos = len(keywords) * len(months) * len(provinces) * len(notice_types)
    combo_idx = 0

    try:
        # 指定顺序：关键词 -> 年月 -> 省份 -> 公告类型。
        # 这样一个关键词会先从 2025 年 1 月各省开始，逐月跑到当前月份，再换下一个关键词。
        for keyword in keywords:
            for year, month in months:
                for prov_name, prov_code in provinces.items():
                    for notice_name, notice_code in notice_types.items():
                        combo_idx += 1
                        key = task_key(keyword, year, month, prov_name, notice_name)
                        meta = {
                            "keyword": keyword,
                            "year": year,
                            "month": month,
                            "province": prov_name,
                            "province_code": prov_code,
                            "notice_type": notice_name,
                            "notice_type_code": notice_code,
                        }
                        if is_task_completed(progress, key):
                            logging.info(
                                "[%d/%d] 跳过已完成：关键词=%s 年月=%s-%s 省份=%s 类型=%s",
                                combo_idx, total_combos, keyword, year, month, prov_name, notice_name,
                            )
                            continue

                        logging.info(
                            "[%d/%d] 开始：关键词=%s 年月=%s-%s 省份=%s 类型=%s",
                            combo_idx, total_combos, keyword, year, month, prov_name, notice_name,
                        )
                        page = get_resume_page(progress, key)
                        if page > 1:
                            logging.info("从断点页继续：%s 第 %s 页", key, page)
                        total_pages = max_pages
                        task_finished = False

                        while page <= total_pages and page <= max_pages:
                            set_current_progress(progress_path, progress, key, page, meta)
                            try:
                                data = request_list_page(
                                    session, cfg, keyword, notice_code, prov_code, year, month, page
                                )
                                if is_login_or_auth_failed(data):
                                    raise RuntimeError(f"疑似登录/权限失败：{data.get('msg')}")
                            except Exception as e:
                                logging.warning("列表请求失败：%s", e)
                                if manual_relogin:
                                    login_mgr.relogin(session)
                                    data = request_list_page(
                                        session, cfg, keyword, notice_code, prov_code, year, month, page
                                    )
                                    if is_login_or_auth_failed(data):
                                        logging.warning("重新登录后仍疑似失败，暂停当前任务：%s", data.get("msg"))
                                        break
                                else:
                                    break

                            if int(data.get("status", 0)) != 200:
                                logging.warning("接口状态异常：%s", data)
                                break
                            page_info = data.get("pageInfo") or {}
                            total_pages = int(page_info.get("totalPages") or 1)
                            total_count = int(page_info.get("totalCount") or 0)
                            is_last = bool(page_info.get("isLastPage"))
                            rows = data.get("data") or []
                            if page == 1 and total_pages >= max_pages and total_count > max_pages * 25:
                                logging.warning(
                                    "组合结果可能超过上限：关键词=%s 年月=%s-%s 省份=%s 类型=%s，totalCount=%s, totalPages=%s。必要时需进一步按天拆分。",
                                    keyword, year, month, prov_name, notice_name, total_count, total_pages,
                                )
                            if not rows:
                                task_finished = True
                                break

                            page_has_detail_failure = False

                            for item in rows:
                                url = item.get("url") or ""
                                if not url:
                                    continue
                                pub_date = item.get("pubDateStr", "")
                                if not date_in_range(pub_date, start, end):
                                    continue
                                source_region = item.get("provName", "")
                                if not province_match(source_region, prov_name):
                                    continue
                                if url in visited:
                                    continue

                                record = NoticeRecord(
                                    search_keyword=keyword,
                                    search_province=prov_name,
                                    search_province_code=prov_code,
                                    search_year=year,
                                    search_month=month,
                                    notice_type=item.get("newTypeName", notice_name),
                                    notice_type_code=item.get("newType", notice_code),
                                    project_name=item.get("title", ""),
                                    publish_time=pub_date,
                                    detail_url=url,
                                    source_region=source_region,
                                    raw_title=item.get("title", ""),
                                    notice_id=str(item.get("id", "")),
                                    enc_id=str(item.get("encId", "")),
                                )
                                try:
                                    record = parse_detail_page(session, url, record)
                                except Exception as e:
                                    # 关键改动：被反爬/521/详情页异常时，不保存数据、不写 visited。
                                    # 断点保留在当前页；下次运行会重新请求当前页，已成功保存过的 URL 会自动跳过。
                                    logging.warning("详情页失败，暂不保存，后续重试：%s | %s", url, e)
                                    page_has_detail_failure = True
                                    break

                                append_jsonl(jsonl_path, record)
                                save_visited(visited_path, url)
                                visited.add(url)
                                logging.info("保存：%s | %s", record.notice_type, record.raw_title[:60])
                                sleep_jitter(detail_delay)

                            if page_has_detail_failure:
                                set_current_progress(progress_path, progress, key, page, meta)
                                logging.warning(
                                    "当前页存在详情页失败，未保存失败记录，断点保留在第 %s 页，下次将从该页继续。",
                                    page,
                                )
                                export_excel(cfg)
                                return

                            # 当前页已经完整处理，断点推进到下一页。
                            set_current_progress(progress_path, progress, key, page + 1, meta)

                            if is_last:
                                task_finished = True
                                break
                            page += 1
                            sleep_jitter(list_delay)
                        else:
                            # while 正常退出，说明页码超过 total_pages 或 max_pages。
                            task_finished = True

                        if task_finished:
                            mark_task_completed(progress_path, progress, key)
                            completed_since_export += 1
                            logging.info("任务完成并写入断点：关键词=%s 年月=%s-%s 省份=%s 类型=%s",
                                         keyword, year, month, prov_name, notice_name)
                            if auto_export_every_tasks > 0 and completed_since_export >= auto_export_every_tasks:
                                export_excel(cfg)
                                completed_since_export = 0
                        else:
                            logging.warning("当前任务未完成，已保留断点，下次将从该任务继续：%s", key)
                            export_excel(cfg)
                            return

        export_excel(cfg)
    finally:
        # 默认不强制关闭浏览器，方便用户排查；如需关闭，取消下一行注释。
        # login_mgr.close()
        pass


if __name__ == "__main__":
    cfg = load_config()
    try:
        crawl(cfg)
    except KeyboardInterrupt:
        print("\n用户中断，正在导出已抓取数据...")
        export_excel(cfg)
    except Exception as e:
        logging.exception("运行失败：%s", e)
        print("\n程序异常退出，尝试导出已抓取数据...")
        try:
            export_excel(cfg)
        except Exception:
            pass
        raise
