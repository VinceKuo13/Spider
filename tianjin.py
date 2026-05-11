# -*- coding: utf-8 -*-
"""
天津市政府采购网爬虫 v2

适用站点：http://www.ccgp-tianjin.gov.cn/
特点：
1. 不需要浏览器、不需要 Playwright、不需要复制 Cookie；
2. 支持采购公告、采购结果公告、采购意向公开；
3. 支持市级、区级；
4. 支持标题关键词和日期范围；
5. 自动翻页、自动抓详情、导出 Excel/CSV/JSON；
6. 修复天津站隐藏字段名 tokken 导致查不到结果的问题。

安装依赖：
    pip install requests beautifulsoup4 lxml pandas openpyxl

运行：
    python tianjin_spider_v2.py
"""

import csv
import json
import logging
import os
import random
import re
import time
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from urllib.parse import parse_qs, urlencode, urljoin, urlparse

import requests
from bs4 import BeautifulSoup

try:
    import pandas as pd
except Exception:
    pd = None


# =========================
# 1. 用户配置区
# =========================

BASE_URL = "http://www.ccgp-tianjin.gov.cn"

# 标题关键词。对应网页高级检索里的“标题”。
KEYWORDS = ["营商环境"]

# 日期范围。
BEGIN_DATE = "2025-01-01"
END_DATE = "2026-05-09"

# 每个分类最多爬多少页；None 表示按网页显示页数全部爬。
MAX_PAGES_PER_CATEGORY = None

# 是否抓详情页。False 则只导出列表。
FETCH_DETAIL = True

# 访问间隔。
SLEEP_RANGE = (0.4, 1.0)

# 输出目录。
OUTPUT_DIR = "outputs_tianjin"

# 分类配置。
CATEGORIES = [
    {"name": "采购公告-市级", "id": "1665", "view": "Infor", "st": "1"},
    {"name": "采购公告-区级", "id": "1664", "view": "Infor", "st": ""},
    {"name": "采购结果公告-市级", "id": "2014", "view": "Infor", "st": "1"},
    {"name": "采购结果公告-区级", "id": "2013", "view": "Infor", "st": ""},
    {"name": "采购意向公开-市级", "id": "2021", "view": "intention", "st": "1"},
    {"name": "采购意向公开-区级", "id": "2022", "view": "intention", "st": ""},
]

# 如果只爬某几类，可以这样改：
# CATEGORIES = [CATEGORIES[0], CATEGORIES[2], CATEGORIES[4]]


# =========================
# 2. 通用工具
# =========================


# 读取同目录 config.yaml 中的 keywords/start_date/end_date。
from common_config import apply_common_config
apply_common_config(globals())


os.makedirs(OUTPUT_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(os.path.join(OUTPUT_DIR, "tianjin_spider.log"), encoding="utf-8"),
    ],
)


def sleep_random():
    time.sleep(random.uniform(*SLEEP_RANGE))


def clean_text(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\xa0", " ").replace("&nbsp;", " ")
    s = re.sub(r"[\u3000\t\r]+", " ", s)
    s = re.sub(r"[ ]{2,}", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()


def soup_text(soup: BeautifulSoup) -> str:
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    return clean_text(soup.get_text("\n"))


def meta_content(soup: BeautifulSoup, name: str) -> str:
    tag = soup.find("meta", attrs={"name": name})
    if tag and tag.get("content"):
        return clean_text(tag.get("content"))
    return ""


def first_regex(text: str, patterns: List[str], flags=re.S) -> str:
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            return clean_text(m.group(1))
    return ""


def normalize_amount(amount: str) -> str:
    amount = clean_text(amount)
    if not amount:
        return ""
    m = re.search(r"([0-9]+(?:,[0-9]{3})*(?:\.\d+)?|[0-9]+(?:\.\d+)?)", amount)
    return m.group(1).replace(",", "") if m else amount


def parse_pubdate_from_meta(pub: str) -> str:
    pub = clean_text(pub)
    if not pub:
        return ""
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", pub)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    month_map = {
        "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04", "May": "05", "Jun": "06",
        "Jul": "07", "Aug": "08", "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12",
    }
    m = re.search(r"\b([A-Z][a-z]{2})\s+(\d{1,2}).*?(\d{4})", pub)
    if m and m.group(1) in month_map:
        return f"{m.group(3)}-{month_map[m.group(1)]}-{int(m.group(2)):02d}"
    return pub


def extract_between(text: str, start_pat: str, end_pat: str) -> str:
    m = re.search(start_pat + r"([\s\S]*?)" + end_pat, text)
    return clean_text(m.group(1)) if m else ""


def extract_label_in_section(section: str, label: str) -> str:
    if not section:
        return ""
    lab = re.escape(label)
    return first_regex(section, [rf"{lab}\s*[:：]\s*([^\n]+)"])


def absolute_url(href: str) -> str:
    href = clean_text(href)
    if not href:
        return ""
    href = href.replace("&amp;", "&")
    return urljoin(BASE_URL, href)


def viewer_to_document_url(url: str) -> str:
    """列表页通常是 /viewer.do?id=xxx&ver=2，详情页用 documentView.do 更稳定。"""
    url = absolute_url(url)
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)
    if parsed.path.endswith("/viewer.do") and qs.get("id"):
        doc_qs = {"method": "view", "id": qs["id"][0], "ver": qs.get("ver", ["2"])[0]}
        return f"{BASE_URL}/portal/documentView.do?{urlencode(doc_qs)}"
    return url


def get_id_from_url(url: str) -> str:
    return parse_qs(urlparse(url).query).get("id", [""])[0]


# =========================
# 3. 爬虫主体
# =========================

class TianjinSpider:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Connection": "keep-alive",
        })
        self.seen_detail_ids = set()

    def request(self, method: str, url: str, **kwargs) -> Optional[requests.Response]:
        for i in range(1, 4):
            try:
                resp = self.session.request(method, url, timeout=30, **kwargs)
                if resp.encoding is None or resp.encoding.lower() in ["iso-8859-1", "latin-1"]:
                    resp.encoding = resp.apparent_encoding or "utf-8"
                if resp.status_code == 200:
                    return resp
                logging.warning("请求状态异常 %s/3：%s，status=%s", i, url, resp.status_code)
            except Exception as e:
                logging.warning("请求失败 %s/3：%s，错误：%s", i, url, e)
            time.sleep(1.5 * i)
        return None

    def get_category_url(self, cat: Dict) -> str:
        params = {
            "method": "view",
            "view": cat["view"],
            "id": cat["id"],
            "ver": "2",
        }
        if cat.get("st"):
            params["st"] = cat["st"]
        params["stmp"] = str(int(time.time() * 1000))
        return f"{BASE_URL}/portal/topicView.do?{urlencode(params)}"

    def extract_token(self, html: str) -> str:
        """
        天津站分页隐藏字段叫 tokken，不是 token/tokens。
        下一页接口实际提交参数叫 tokens，所以这里先提取 tokken 的值，再放进 tokens。
        """
        if not html:
            return ""
        soup = BeautifulSoup(html, "lxml")
        for key in ["tokken", "tokens", "token"]:
            inp = soup.find("input", attrs={"name": key}) or soup.find("input", attrs={"id": key})
            if inp and inp.get("value"):
                return clean_text(inp.get("value"))
        patterns = [
            r'name=["\']tokken["\'][^>]*value=["\']([^"\']+)["\']',
            r'id=["\']tokken["\'][^>]*value=["\']([^"\']+)["\']',
            r'name=["\']tokens["\'][^>]*value=["\']([^"\']+)["\']',
            r'id=["\']tokens["\'][^>]*value=["\']([^"\']+)["\']',
            r'tokens\s*[:=]\s*["\']([0-9a-fA-F]{16,})["\']',
            r'tokken\s*[:=]\s*["\']([0-9a-fA-F]{16,})["\']',
        ]
        return first_regex(html, patterns)

    def open_category(self, cat: Dict) -> Tuple[str, str]:
        url = self.get_category_url(cat)
        logging.info("访问分类页：%s | %s", cat["name"], url)
        resp = self.request("GET", url, headers={"Referer": BASE_URL + "/"})
        if not resp:
            return "", url
        token = self.extract_token(resp.text)
        if token:
            logging.info("获取 tokken 成功：%s...", token[:8])
        else:
            logging.warning("未解析到 tokken；仍会尝试 POST 查询。")
        return token, url

    def build_search_data(self, cat: Dict, keyword: str, page: int, token: str) -> Dict[str, str]:
        return {
            "method": "find",
            "tokens": token or "",
            "id": cat["id"],
            "page": str(page),
            "name": keyword,
            "st": cat.get("st", ""),
            "view": cat["view"],
            "ldateQGE": BEGIN_DATE,
            "ldateQLE": END_DATE,
            "siteLists": "",
            "buyUnitName": "",
            "projectCode": "",
            "projectName": "",
        }

    def search_page(self, cat: Dict, keyword: str, page: int, token: str, referer: str) -> Tuple[str, str]:
        # 注意：真实接口是 POST /portal/topicView.do，表单参数 method=find。
        url = f"{BASE_URL}/portal/topicView.do"
        headers = {
            "Accept": "*/*",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Origin": BASE_URL,
            "Referer": referer,
            "X-Requested-With": "XMLHttpRequest",
        }
        data = self.build_search_data(cat, keyword, page, token)
        logging.info("请求列表：分类=%s，关键词=%s，第 %s 页", cat["name"], keyword, page)
        resp = self.request("POST", url, headers=headers, data=data)
        if not resp:
            return "", token

        html = resp.text
        new_token = self.extract_token(html) or token

        # 某些情况下 XHR 返回不完整，兜底再走一次普通表单 POST URL。
        if "dataList" not in html and "viewer.do" not in html and page == 1:
            alt_url = f"{BASE_URL}/portal/topicView.do?method=find"
            resp2 = self.request("POST", alt_url, headers=headers, data=data)
            if resp2 and ("dataList" in resp2.text or "viewer.do" in resp2.text):
                html = resp2.text
                new_token = self.extract_token(html) or new_token

        return html, new_token

    def parse_total_pages(self, html: str) -> int:
        if not html:
            return 1
        m = re.search(r"共\s*<b[^>]*>\s*(\d+)\s*</b>\s*页", html)
        if m:
            return int(m.group(1))
        text = clean_text(BeautifulSoup(html, "lxml").get_text(" "))
        m = re.search(r"共\s*(\d+)\s*页", text)
        if m:
            return int(m.group(1))
        return 1

    def parse_list_items(self, html: str, cat: Dict, keyword: str) -> List[Dict[str, str]]:
        soup = BeautifulSoup(html or "", "lxml")
        items = []

        # 主解析：ul.dataList li
        li_nodes = soup.select("ul.dataList li")
        # 兜底：全页面 viewer/documentView 链接
        if not li_nodes:
            li_nodes = []
            for a in soup.select("a[href]"):
                href = a.get("href", "")
                if "viewer.do" in href or "documentView.do" in href:
                    li_nodes.append(a.parent or a)

        for node in li_nodes:
            a = node.find("a", href=True) if hasattr(node, "find") else None
            if not a:
                continue
            href_raw = a.get("href", "")
            if "viewer.do" not in href_raw and "documentView.do" not in href_raw:
                continue
            title = clean_text(a.get("title") or a.get_text(" "))
            if not title:
                continue
            time_tag = node.select_one("span.time") if hasattr(node, "select_one") else None
            pub_date = clean_text(time_tag.get_text(" ")) if time_tag else ""
            detail_url = viewer_to_document_url(href_raw)
            items.append({
                "关键词": keyword,
                "公告分类": cat["name"],
                "分类ID": cat["id"],
                "列表标题": title,
                "列表发布日期": pub_date,
                "详情URL": detail_url,
                "详情ID": get_id_from_url(detail_url),
            })

        uniq = []
        seen = set()
        for it in items:
            key = it.get("详情URL") or it.get("列表标题")
            if key in seen:
                continue
            seen.add(key)
            uniq.append(it)
        return uniq

    def fetch_detail(self, url: str) -> str:
        headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Referer": BASE_URL + "/",
        }
        resp = self.request("GET", url, headers=headers)
        return resp.text if resp else ""

    def parse_tables(self, soup: BeautifulSoup) -> List[Tuple[List[str], List[List[str]]]]:
        tables = []
        for table in soup.find_all("table"):
            rows = []
            for tr in table.find_all("tr"):
                cells = [clean_text(td.get_text(" ")) for td in tr.find_all(["td", "th"])]
                cells = [c for c in cells if c]
                if cells:
                    rows.append(cells)
            if len(rows) >= 2:
                tables.append((rows[0], rows[1:]))
        return tables

    def parse_supplier_rows(self, soup: BeautifulSoup) -> List[Dict[str, str]]:
        result = []
        for headers, data_rows in self.parse_tables(soup):
            header_text = "|".join(headers)
            if "供应商名称" not in header_text and "中标" not in header_text and "成交" not in header_text:
                continue
            for row in data_rows:
                row = row + [""] * max(0, len(headers) - len(row))
                d = {headers[i]: row[i] for i in range(min(len(headers), len(row)))}
                supplier = d.get("供应商名称", "") or d.get("中标供应商名称", "") or d.get("成交供应商名称", "")
                if not supplier or supplier in ["供应商名称", "第1包", "第2包"]:
                    continue
                amount = ""
                for k, v in d.items():
                    if "中标金额" in k or "成交金额" in k:
                        amount = v
                        break
                result.append({
                    "中标供应商": supplier,
                    "供应商地址": d.get("供应商地址", ""),
                    "统一社会信用代码": d.get("统一社会信用代码", "") or d.get("统一信用代码", ""),
                    "中标金额(万元)": normalize_amount(amount),
                    "评审得分": d.get("评审得分", ""),
                })
        return result

    def parse_intention_rows(self, soup: BeautifulSoup) -> List[Dict[str, str]]:
        result = []
        for headers, data_rows in self.parse_tables(soup):
            header_text = "|".join(headers)
            if "采购项目名称" not in header_text or "预算金额" not in header_text:
                continue
            for row in data_rows:
                row = row + [""] * max(0, len(headers) - len(row))
                d = {headers[i]: row[i] for i in range(min(len(headers), len(row)))}
                project_name = d.get("采购项目名称", "")
                if not project_name or project_name == "采购项目名称":
                    continue
                budget = ""
                for k, v in d.items():
                    if "预算金额" in k:
                        budget = v
                        break
                result.append({
                    "意向采购项目名称": project_name,
                    "采购需求概况": d.get("采购需求概况", ""),
                    "意向预算金额(万元)": normalize_amount(budget),
                    "预计采购时间": d.get("预计采购时间", "").replace("\n", " "),
                    "政府采购政策": d.get("执行的政府采购政策", ""),
                    "备注": d.get("备注", ""),
                })
        return result

    def parse_detail(self, list_item: Dict[str, str], html: str) -> List[Dict[str, str]]:
        base = deepcopy(list_item)
        if not html:
            base["详情解析状态"] = "详情页获取失败"
            return [base]

        soup = BeautifulSoup(html, "lxml")
        text = soup_text(BeautifulSoup(html, "lxml"))

        base.update({
            "详情解析状态": "成功",
            "标题": meta_content(soup, "ArticleTitle") or base.get("列表标题", ""),
            "网页栏目": meta_content(soup, "ColumnName") or base.get("公告分类", ""),
            "详情发布日期": parse_pubdate_from_meta(meta_content(soup, "PubDate")) or base.get("列表发布日期", ""),
            "内容来源": meta_content(soup, "ContentSource"),
            "正文文本": text[:3000],
        })

        base["项目编号"] = first_regex(text, [
            r"项目编号\s*[:：]\s*([^\n]+)",
            r"一、项目编号\s*[:：]\s*([^\n]+)",
            r"项目编号\s*[:：]?\s*([A-Za-z0-9_\-（）()\.]+)",
        ])
        base["项目编号"] = re.sub(r"^.*?项目编号\s*[:：]", "", base["项目编号"]).strip()

        base["项目名称"] = first_regex(text, [
            r"项目名称\s*[:：]\s*([^\n]+)",
            r"二、项目名称\s*[:：]\s*([^\n]+)",
        ])
        base["采购方式"] = first_regex(text, [r"采购方式\s*[:：]\s*([^\n]+)"])
        base["预算金额(万元)"] = normalize_amount(first_regex(text, [
            r"预算金额\s*[:：]\s*([^\n]+)",
            r"预算金额（万元）\s*([^\n]+)",
        ]))
        base["最高限价(万元)"] = normalize_amount(first_regex(text, [r"最高限价\s*[:：]\s*([^\n]+)"]))

        purchaser_section = extract_between(text, r"1\.?\s*采购人信息", r"2\.?\s*采购代理机构信息")
        agency_section = extract_between(text, r"2\.?\s*采购代理机构信息", r"3\.?\s*项目联系方式")
        contact_section = extract_between(text, r"3\.?\s*项目联系方式", r"(?:十、附件|附件|$)")

        base["采购人"] = extract_label_in_section(purchaser_section, "名称")
        base["采购人地址"] = extract_label_in_section(purchaser_section, "地址")
        base["采购人联系方式"] = extract_label_in_section(purchaser_section, "联系方式")
        base["代理机构"] = extract_label_in_section(agency_section, "名称")
        base["代理机构地址"] = extract_label_in_section(agency_section, "地址")
        base["代理机构联系方式"] = extract_label_in_section(agency_section, "联系方式")
        base["项目联系人"] = extract_label_in_section(contact_section, "项目联系人")
        base["项目联系电话"] = first_regex(contact_section, [r"电\s*话\s*[:：]\s*([^\n]+)", r"电话\s*[:：]\s*([^\n]+)"])

        base["总中标成交金额(万元)"] = normalize_amount(first_regex(text, [
            r"总中标成交金额\s*[:：]\s*([^\n]+)",
            r"总成交金额\s*[:：]\s*([^\n]+)",
            r"中标金额\s*[:：]\s*([^\n]+)",
            r"成交金额\s*[:：]\s*([^\n]+)",
        ]))

        supplier_rows = self.parse_supplier_rows(soup)
        if supplier_rows:
            records = []
            for r in supplier_rows:
                item = deepcopy(base)
                item.update(r)
                if not item.get("总中标成交金额(万元)"):
                    item["总中标成交金额(万元)"] = item.get("中标金额(万元)", "")
                records.append(item)
            return records

        intention_rows = self.parse_intention_rows(soup)
        if intention_rows:
            records = []
            for r in intention_rows:
                item = deepcopy(base)
                item.update(r)
                if not item.get("项目名称"):
                    item["项目名称"] = item.get("意向采购项目名称", "")
                if not item.get("预算金额(万元)"):
                    item["预算金额(万元)"] = item.get("意向预算金额(万元)", "")
                records.append(item)
            return records

        return [base]

    def crawl(self) -> List[Dict[str, str]]:
        all_records = []
        logging.info("开始爬取天津市政府采购网")
        logging.info("关键词：%s；时间：%s 至 %s；分类数：%s", KEYWORDS, BEGIN_DATE, END_DATE, len(CATEGORIES))

        for keyword in KEYWORDS:
            for cat in CATEGORIES:
                token, referer = self.open_category(cat)
                html, token = self.search_page(cat, keyword, 1, token, referer)
                if not html:
                    continue

                total_pages = self.parse_total_pages(html)
                if MAX_PAGES_PER_CATEGORY is not None:
                    total_pages = min(total_pages, int(MAX_PAGES_PER_CATEGORY))
                logging.info("分类=%s，关键词=%s，预计页数=%s", cat["name"], keyword, total_pages)

                for page in range(1, total_pages + 1):
                    if page == 1:
                        page_html = html
                    else:
                        sleep_random()
                        page_html, token = self.search_page(cat, keyword, page, token, referer)
                        if not page_html:
                            continue

                    items = self.parse_list_items(page_html, cat, keyword)
                    logging.info("第 %s 页解析到 %s 条列表记录", page, len(items))
                    if not items:
                        logging.warning("本页未解析到列表：分类=%s，关键词=%s，第%s页", cat["name"], keyword, page)
                        if page != 1:
                            break
                        continue

                    for it in items:
                        detail_id = it.get("详情ID") or it.get("详情URL")
                        if detail_id in self.seen_detail_ids:
                            continue
                        self.seen_detail_ids.add(detail_id)
                        if FETCH_DETAIL:
                            logging.info("详情：%s", it.get("详情URL"))
                            sleep_random()
                            detail_html = self.fetch_detail(it.get("详情URL", ""))
                            all_records.extend(self.parse_detail(it, detail_html))
                        else:
                            all_records.append(it)

        logging.info("爬取完成，共得到 %s 条导出记录", len(all_records))
        return all_records


# =========================
# 4. 导出
# =========================

def export_records(records: List[Dict[str, str]]):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_path = os.path.join(OUTPUT_DIR, f"tianjin_records_{ts}.json")
    csv_path = os.path.join(OUTPUT_DIR, f"tianjin_records_{ts}.csv")
    xlsx_path = os.path.join(OUTPUT_DIR, f"tianjin_records_{ts}.xlsx")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    preferred_cols = [
        "关键词", "公告分类", "网页栏目", "列表发布日期", "详情发布日期",
        "标题", "列表标题", "项目编号", "项目名称", "采购方式",
        "采购人", "采购人地址", "采购人联系方式",
        "代理机构", "代理机构地址", "代理机构联系方式",
        "项目联系人", "项目联系电话",
        "预算金额(万元)", "最高限价(万元)",
        "中标供应商", "供应商地址", "统一社会信用代码", "中标金额(万元)", "总中标成交金额(万元)", "评审得分",
        "意向采购项目名称", "采购需求概况", "意向预算金额(万元)", "预计采购时间", "政府采购政策", "备注",
        "详情URL", "详情ID", "详情解析状态", "正文文本",
    ]
    all_keys = []
    for r in records:
        for k in r.keys():
            if k not in all_keys:
                all_keys.append(k)
    cols = [c for c in preferred_cols if c in all_keys] + [c for c in all_keys if c not in preferred_cols]

    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=cols)
        writer.writeheader()
        for r in records:
            writer.writerow({k: r.get(k, "") for k in cols})

    if pd is not None:
        df = pd.DataFrame(records)
        if not df.empty:
            df = df.reindex(columns=cols)
        else:
            df = pd.DataFrame(columns=cols)
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="天津政府采购")
            ws = writer.book["天津政府采购"]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = 10
                col_letter = col[0].column_letter
                for cell in col[:200]:
                    max_len = max(max_len, min(50, len(str(cell.value or ""))))
                ws.column_dimensions[col_letter].width = max_len + 2
        logging.info("Excel 已导出：%s", xlsx_path)
    else:
        logging.warning("未安装 pandas/openpyxl，已跳过 Excel。")

    logging.info("CSV 已导出：%s", csv_path)
    logging.info("JSON 已导出：%s", json_path)


if __name__ == "__main__":
    try:
        spider = TianjinSpider()
        records = spider.crawl()
        export_records(records)
        if not records:
            logging.warning("没有导出记录。请检查关键词、日期或网站返回页面。")
    except KeyboardInterrupt:
        logging.warning("用户中断。")
    except Exception as e:
        logging.exception("运行失败：%s", e)
