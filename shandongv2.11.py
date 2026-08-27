# -*- coding: utf-8 -*-
"""
山东省政府采购信息公开平台爬虫
运行方式：
    python shandong.py

配置方式：
    修改同目录下的 config.yaml：
        keywords:
          - 环境
          - 营商环境
        start_date: "2025-01-01"
        end_date: ""

说明：
1. 数据源改为 http://www.ccgp-shandong.gov.cn/xxgk
2. 列表接口为 https://www.ccgp-shandong.gov.cn:8087/api/website/site/getListByCode
3. 采购意向 colCode=2500
4. 采购公告-省级 colCode=0301
5. 结果公告-省级 colCode=0302
6. 采购意向-市区县 colCode=2504
7. 采购公告-市区县 colCode=0303
8. 结果公告-市区县 colCode=0304
9. 本站搜索需要验证码。程序优先调用真实验证码接口 /api/website/captcha，正常只需输入图片验证码一次；
   只有接口解析失败时，才需要从 F12 的 getListByCode 请求中复制 captchaUuid。
"""

from __future__ import annotations

import base64
import json
import logging
import math
import mimetypes
import os
import random
import re
import sys
import time
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import quote_plus, urlencode, urlparse, parse_qs, unquote

import pandas as pd
import requests
import yaml
from bs4 import BeautifulSoup


# ========================
# 1. 基础配置
# ========================

BASE_URL = "http://www.ccgp-shandong.gov.cn"
HOME_URL = "http://www.ccgp-shandong.gov.cn/xxgk"
LIST_API = "https://www.ccgp-shandong.gov.cn:8087/api/website/site/getListByCode"

CONFIG_PATH = Path("config.yaml")

DEFAULT_KEYWORDS = ["环境"]
DEFAULT_START_DATE = "2025-01-01"
DEFAULT_END_DATE = ""  # 留空代表今天

OUTPUT_DIR = Path("outputs_shandong")
CAPTCHA_DIR = OUTPUT_DIR / "captcha"
RAW_DIR = OUTPUT_DIR / "raw"

JSONL_PATH = OUTPUT_DIR / "山东政府采购_记录.jsonl"
JSON_PATH = OUTPUT_DIR / "山东政府采购_记录.json"
EXCEL_PATH = OUTPUT_DIR / "山东政府采购_标讯明细.xlsx"
VISITED_PATH = OUTPUT_DIR / "visited_urls.txt"
LOG_PATH = OUTPUT_DIR / "shandong_spider.log"

PAGE_SIZE = 10
MAX_PAGES_PER_QUERY = 9999
REQUEST_DELAY = 0.8
DETAIL_DELAY = 0.6
TIMEOUT = 30
RETRY_TIMES = 3

# 是否抓取市区县数据
INCLUDE_CITY_COUNTY = True

# 详情页如果 requests 抓不到正文，是否尝试 Playwright 渲染详情页。
# 如果没有安装 Playwright，也不会影响列表爬取，只是部分详情字段可能为空。
USE_PLAYWRIGHT_DETAIL_FALLBACK = True

# 每个关键词是否重新获取一次验证码。
# False：本次运行尽量复用一次验证码；如果验证码失效再重新输入。
CAPTCHA_EVERY_KEYWORD = False


# 任务配置
TASKS = [
    # 采购意向
    # 山东新站“意向公开-省级”和“意向公开-市区县”不是同一个 colCode：
    # 省级：colCode=2500, area=370000
    # 市区县：colCode=2504, area=""
    {"category": "采购意向", "level": "省级", "colCode": "2500", "area": "370000"},
    {"category": "采购意向", "level": "市区县", "colCode": "2504", "area": "", "colCode_candidates": ["2504"]},

    # 采购公告
    {"category": "采购公告", "level": "省级", "colCode": "0301", "area": "370000"},
    {"category": "采购公告", "level": "市区县", "colCode": "0303", "area": ""},

    # 采购结果
    {"category": "采购结果", "level": "省级", "colCode": "0302", "area": "370000"},
    {"category": "采购结果", "level": "市区县", "colCode": "0304", "area": ""},
]


# ========================
# 2. 通用工具
# ========================

def ensure_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)


def setup_logging() -> None:
    ensure_dirs()
    logger = logging.getLogger()
    logger.handlers.clear()
    logger.setLevel(logging.INFO)

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)


def load_config() -> Dict[str, Any]:
    cfg: Dict[str, Any] = {}
    if CONFIG_PATH.exists():
        try:
            cfg = yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8")) or {}
        except Exception as e:
            logging.warning("读取 config.yaml 失败，将使用默认配置：%s", e)

    shandong_cfg = cfg.get("shandong") or {}

    global OUTPUT_DIR, CAPTCHA_DIR, RAW_DIR, JSONL_PATH, JSON_PATH, EXCEL_PATH, VISITED_PATH, LOG_PATH
    global PAGE_SIZE, MAX_PAGES_PER_QUERY, REQUEST_DELAY, DETAIL_DELAY, TIMEOUT, RETRY_TIMES
    global INCLUDE_CITY_COUNTY, USE_PLAYWRIGHT_DETAIL_FALLBACK, CAPTCHA_EVERY_KEYWORD

    OUTPUT_DIR = Path(str(shandong_cfg.get("output_dir") or cfg.get("output_dir") or OUTPUT_DIR))
    CAPTCHA_DIR = OUTPUT_DIR / "captcha"
    RAW_DIR = OUTPUT_DIR / "raw"
    JSONL_PATH = OUTPUT_DIR / "山东政府采购_记录.jsonl"
    JSON_PATH = OUTPUT_DIR / "山东政府采购_记录.json"
    EXCEL_PATH = OUTPUT_DIR / "山东政府采购_标讯明细.xlsx"
    VISITED_PATH = OUTPUT_DIR / "visited_urls.txt"
    LOG_PATH = OUTPUT_DIR / "shandong_spider.log"

    PAGE_SIZE = int(shandong_cfg.get("page_size") or cfg.get("page_size") or PAGE_SIZE)
    MAX_PAGES_PER_QUERY = int(shandong_cfg.get("max_pages_per_query") or cfg.get("max_pages_per_query") or MAX_PAGES_PER_QUERY)
    REQUEST_DELAY = float(shandong_cfg.get("request_delay") or cfg.get("request_delay") or REQUEST_DELAY)
    DETAIL_DELAY = float(shandong_cfg.get("detail_delay") or cfg.get("detail_delay") or DETAIL_DELAY)
    TIMEOUT = int(shandong_cfg.get("timeout") or cfg.get("timeout") or TIMEOUT)
    RETRY_TIMES = int(shandong_cfg.get("retry_times") or cfg.get("retry_times") or RETRY_TIMES)

    INCLUDE_CITY_COUNTY = bool(shandong_cfg.get("include_city_county", INCLUDE_CITY_COUNTY))
    USE_PLAYWRIGHT_DETAIL_FALLBACK = bool(shandong_cfg.get("use_playwright_detail_fallback", USE_PLAYWRIGHT_DETAIL_FALLBACK))
    CAPTCHA_EVERY_KEYWORD = bool(shandong_cfg.get("captcha_every_keyword", CAPTCHA_EVERY_KEYWORD))

    return cfg


def get_keywords(cfg: Dict[str, Any]) -> List[str]:
    kws = cfg.get("keywords") or DEFAULT_KEYWORDS
    if isinstance(kws, str):
        kws = [kws]
    kws = [str(x).strip() for x in kws if str(x).strip()]
    return kws or DEFAULT_KEYWORDS


def get_date_range(cfg: Dict[str, Any]) -> Tuple[str, str]:
    start = str(cfg.get("start_date") or DEFAULT_START_DATE).strip()
    end = str(cfg.get("end_date") or DEFAULT_END_DATE).strip()
    if not end:
        end = datetime.now().strftime("%Y-%m-%d")
    return start, end


def clean_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\xa0", " ").replace("&nbsp;", " ")
    s = re.sub(r"<script[\s\S]*?</script>", " ", s, flags=re.I)
    s = re.sub(r"<style[\s\S]*?</style>", " ", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def normalize_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ").replace("&nbsp;", " ")
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n\s+", "\n", s)
    s = re.sub(r"\s+\n", "\n", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def safe_filename(s: str, max_len: int = 80) -> str:
    s = clean_text(s)
    s = re.sub(r'[\\/:*?"<>|\s]+', "_", s)
    return s[:max_len] or "file"


def open_file_for_user(path: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f"open '{path}' >/dev/null 2>&1 &")
        else:
            os.system(f"xdg-open '{path}' >/dev/null 2>&1 &")
    except Exception:
        pass


def sleep_delay(base: float) -> None:
    time.sleep(base + random.uniform(0.1, 0.6))


def first_nonempty(*vals: Any) -> str:
    for v in vals:
        if v is None:
            continue
        s = str(v).strip()
        if s and s.lower() not in {"none", "null"}:
            return s
    return ""


def get_by_keys(obj: Dict[str, Any], keys: Iterable[str]) -> str:
    for k in keys:
        if k in obj and obj[k] not in (None, "", "null"):
            return str(obj[k]).strip()
    return ""


def make_session() -> requests.Session:
    s = requests.Session()

    # 不继承 PowerShell / Conda / Clash 等环境中的 HTTP_PROXY、HTTPS_PROXY。
    # 山东政府采购平台在本脚本中按直连访问，避免请求被错误转发到
    # 127.0.0.1:7897 等本地代理后超时。
    s.trust_env = False

    s.headers.update({
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": BASE_URL,
        "Referer": BASE_URL + "/",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/147.0.0.0 Safari/537.36"
        ),
    })
    return s


def request_json(session: requests.Session, method: str, url: str, **kwargs) -> Dict[str, Any]:
    last_err = None
    for i in range(1, RETRY_TIMES + 1):
        try:
            resp = session.request(method, url, timeout=TIMEOUT, **kwargs)
            resp.encoding = resp.apparent_encoding or "utf-8"
            if resp.status_code >= 400:
                raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:300]}")
            return resp.json()
        except Exception as e:
            last_err = e
            logging.warning("请求失败 %d/%d：%s；错误：%s", i, RETRY_TIMES, url, e)
            time.sleep(1.0 + i)
    raise RuntimeError(f"请求多次失败：{url}；最后错误：{last_err}")


def request_text(session: requests.Session, url: str) -> str:
    last_err = None
    for i in range(1, RETRY_TIMES + 1):
        try:
            resp = session.get(url, timeout=TIMEOUT)
            resp.encoding = resp.apparent_encoding or "utf-8"
            if resp.status_code >= 400:
                raise RuntimeError(f"HTTP {resp.status_code}")
            return resp.text
        except Exception as e:
            last_err = e
            logging.warning("详情请求失败 %d/%d：%s；错误：%s", i, RETRY_TIMES, url, e)
            time.sleep(1.0 + i)
    raise RuntimeError(f"详情请求失败：{url}；最后错误：{last_err}")


# ========================
# 3. 验证码处理
# ========================

def find_in_dict(d: Any, keys: Iterable[str]) -> str:
    """在嵌套 dict/list 中递归找 key。"""
    if isinstance(d, dict):
        for k in keys:
            if k in d and d[k] not in (None, ""):
                return str(d[k])
        for v in d.values():
            r = find_in_dict(v, keys)
            if r:
                return r
    elif isinstance(d, list):
        for v in d:
            r = find_in_dict(v, keys)
            if r:
                return r
    return ""


def save_base64_image(img_text: str, prefix: str = "captcha") -> Optional[Path]:
    if not img_text:
        return None

    img_text = img_text.strip()
    ext = ".png"

    if img_text.startswith("data:image"):
        m = re.match(r"data:image/([a-zA-Z0-9+]+);base64,(.*)", img_text, re.S)
        if not m:
            return None
        ext = "." + m.group(1).replace("jpeg", "jpg")
        b64 = m.group(2)
    else:
        b64 = img_text

    try:
        content = base64.b64decode(b64)
        path = CAPTCHA_DIR / f"{prefix}_{int(time.time())}{ext}"
        path.write_bytes(content)
        return path
    except Exception:
        return None



def save_captcha_from_value(value: str, session: Optional[requests.Session] = None) -> Optional[Path]:
    """保存验证码图片。支持 base64、data:image、http URL、相对 URL。"""
    if not value:
        return None
    value = str(value).strip()

    # base64 / data:image
    if value.startswith("data:image") or re.match(r"^[A-Za-z0-9+/=\s]{80,}$", value):
        return save_base64_image(value, "shandong_captcha")

    # 图片 URL
    if value.startswith("//"):
        value = "http:" + value
    elif value.startswith("/"):
        value = BASE_URL + value

    if value.startswith("http://") or value.startswith("https://"):
        if session is None:
            session = make_session()
        try:
            resp = session.get(value, timeout=TIMEOUT, headers={
                "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
                "Referer": BASE_URL + "/",
                "User-Agent": session.headers.get("User-Agent", ""),
            })
            ctype = resp.headers.get("Content-Type", "").lower()
            if resp.status_code < 400 and ("image/" in ctype or resp.content[:8].startswith(b"\x89PNG") or resp.content[:3] == b"\xff\xd8\xff"):
                ext = mimetypes.guess_extension(ctype.split(";")[0]) or ".png"
                path = CAPTCHA_DIR / f"shandong_captcha_{int(time.time())}{ext}"
                path.write_bytes(resp.content)
                return path
        except Exception:
            return None

    return None


def find_captcha_fields(obj: Any) -> Tuple[str, str]:
    """
    递归查找山东验证码返回中的 uuid 和图片字段。
    兼容常见字段名。
    """
    uuid_keys = {
        "captchaUuid", "captchaUUID", "uuid", "captchaId", "captchaID",
        "captchaKey", "key", "id", "token", "captchaToken"
    }
    image_keys = {
        "captchaImg", "captchaImage", "captchaBase64", "base64",
        "image", "img", "pic", "picPath", "imgUrl", "imageUrl",
        "captchaUrl", "url", "codeImg", "verifyImg", "verifyImage"
    }

    found_uuid = ""
    found_img = ""

    def walk(x: Any) -> None:
        nonlocal found_uuid, found_img
        if isinstance(x, dict):
            for k, v in x.items():
                kl = str(k)
                if not found_uuid and kl in uuid_keys and v not in (None, ""):
                    found_uuid = str(v)
                if not found_img and kl in image_keys and v not in (None, ""):
                    found_img = str(v)
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)

    walk(obj)
    return found_uuid, found_img

def extract_captcha_uuid_from_response(resp: requests.Response, session: Optional[requests.Session] = None) -> str:
    """
    山东验证码接口有时直接返回图片，captchaUuid 可能放在响应头或 Cookie 中。
    这里尽量从 headers / cookies / session.cookies 中提取。
    """
    candidates = []

    # 1. 响应头
    for k, v in resp.headers.items():
        name = str(k).lower()
        val = str(v).strip()
        if any(x in name for x in ["captcha", "uuid", "token", "key", "code"]):
            candidates.append(val)

    # 2. 当前响应 cookie
    for k, v in resp.cookies.items():
        name = str(k).lower()
        val = str(v).strip()
        if any(x in name for x in ["captcha", "uuid", "token", "key", "code"]):
            candidates.append(val)

    # 3. session cookie
    if session is not None:
        for k, v in session.cookies.items():
            name = str(k).lower()
            val = str(v).strip()
            if any(x in name for x in ["captcha", "uuid", "token", "key", "code"]):
                candidates.append(val)

    # 从候选字符串里提取像 uuid/md5 的长串
    for val in candidates:
        if not val:
            continue
        m = re.search(r"([A-Fa-f0-9]{32})", val)
        if m:
            return m.group(1)
        m = re.search(r"([A-Za-z0-9_-]{16,})", val)
        if m:
            return m.group(1)

    return ""



def try_auto_get_captcha(session: requests.Session) -> Tuple[str, Optional[Path]]:
    """
    通过山东站真实验证码接口获取验证码。
    用户提供的真实接口：
        https://www.ccgp-shandong.gov.cn:8087/api/website/captcha

    正常返回 JSON，里面包含 captchaUuid 和验证码图片/base64/图片URL。
    """
    # BASE_URL 当前就是 http://www.ccgp-shandong.gov.cn，
    # 原来的两个候选地址实际完全相同；去重后避免失败时重复等待一次 TIMEOUT。
    candidates = list(dict.fromkeys([
        "https://www.ccgp-shandong.gov.cn:8087/api/website/captcha",
        "https://www.ccgp-shandong.gov.cn:8087/api/website/captcha",
    ]))

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Origin": BASE_URL,
        "Referer": BASE_URL + "/",
        "User-Agent": session.headers.get("User-Agent", ""),
    }

    last_text = ""

    for url in candidates:
        try:
            # 先按用户抓到的 cURL 原样请求，不追加随机参数。
            resp = session.get(url, timeout=TIMEOUT, headers=headers)
            resp.encoding = resp.apparent_encoding or "utf-8"
            ctype = resp.headers.get("Content-Type", "").lower()
            last_text = resp.text[:1500] if resp.text else ""

            # JSON 返回
            if "json" in ctype or resp.text.strip().startswith("{"):
                js = resp.json()
                uuid, img_value = find_captcha_fields(js)

                # 保存原始 JSON，便于后续排查字段结构
                raw_path = RAW_DIR / f"captcha_response_{int(time.time())}.json"
                raw_path.write_text(json.dumps(js, ensure_ascii=False, indent=2), encoding="utf-8")

                path = save_captcha_from_value(img_value, session=session) if img_value else None

                if uuid and path:
                    logging.info("自动获取验证码成功：uuid=%s，图片=%s", uuid, path)
                    return uuid, path

                if path:
                    logging.info("自动获取验证码图片成功，但未解析到 captchaUuid：%s", path)
                    return "", path

                if uuid:
                    logging.info("自动获取到 captchaUuid=%s，但未解析到验证码图片。原始响应已保存：%s", uuid, raw_path)
                    return uuid, None

                logging.warning("验证码接口返回 JSON，但未识别到 uuid/图片字段，原始响应已保存：%s", raw_path)

            # 少数情况直接返回图片。此时 captchaUuid 可能在响应头或 Cookie 中。
            elif "image/" in ctype:
                ext = mimetypes.guess_extension(ctype.split(";")[0]) or ".png"
                path = CAPTCHA_DIR / f"shandong_captcha_{int(time.time())}{ext}"
                path.write_bytes(resp.content)

                uuid = extract_captcha_uuid_from_response(resp, session=session)

                # 保存响应头，便于排查 captchaUuid 到底在哪里。
                try:
                    header_path = RAW_DIR / f"captcha_image_headers_{int(time.time())}.json"
                    header_path.write_text(json.dumps({
                        "url": url,
                        "headers": dict(resp.headers),
                        "cookies": requests.utils.dict_from_cookiejar(resp.cookies),
                        "session_cookies": requests.utils.dict_from_cookiejar(session.cookies),
                        "parsed_uuid": uuid,
                    }, ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception:
                    pass

                if uuid:
                    logging.info("验证码接口直接返回图片，并解析到 captchaUuid=%s：%s", uuid, path)
                else:
                    logging.info("验证码接口直接返回图片，但未从响应头/Cookie解析到 captchaUuid：%s", path)

                return uuid, path

        except Exception as e:
            logging.warning("验证码接口请求失败：%s；错误：%s", url, e)
            continue

    if last_text:
        raw_path = RAW_DIR / f"captcha_unknown_{int(time.time())}.txt"
        raw_path.write_text(last_text, encoding="utf-8", errors="ignore")
        logging.warning("验证码接口未能解析，原始响应已保存：%s", raw_path)

    return "", None

def extract_uuid_from_url(url: str) -> str:
    """从验证码图片 URL 中尽量提取 captchaUuid。"""
    try:
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)
        for key in ["captchaUuid", "uuid", "captchaId", "captchaKey", "key"]:
            val = qs.get(key)
            if val and val[0]:
                return str(val[0])
        m = re.search(r"([A-Za-z0-9_-]{16,})", unquote(parsed.path))
        if m:
            return m.group(1)
    except Exception:
        pass
    return ""


def try_playwright_get_captcha_image() -> Tuple[str, Optional[Path]]:
    """
    用独立子进程打开山东页面并截取验证码图片。

    不能在主进程里再次调用 sync_playwright()，否则在某些环境会报：
    It looks like you are using Playwright Sync API inside the asyncio loop.
    因此这里把验证码截图动作放到一个临时 Python 子进程中执行。
    """
    child_code = r"""
import json
import time
import re
import mimetypes
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote

HOME_URL = {home_url!r}
CAPTCHA_DIR = Path({captcha_dir!r})
CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)

def extract_uuid_from_url(url):
    try:
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)
        for key in ["captchaUuid", "uuid", "captchaId", "captchaKey", "key"]:
            val = qs.get(key)
            if val and val[0]:
                return str(val[0])
        m = re.search(r"([A-Za-z0-9_-]{{16,}})", unquote(parsed.path))
        if m:
            return m.group(1)
    except Exception:
        pass
    return ""

def save_bytes(content, suffix=".png"):
    if not content:
        return ""
    path = CAPTCHA_DIR / ("shandong_captcha_%d%s" % (int(time.time()), suffix))
    path.write_bytes(content)
    return str(path)

result = {{"uuid": "", "path": "", "error": ""}}

try:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage", "--no-proxy-server"],
        )
        page = browser.new_page(
            viewport={{"width": 1366, "height": 900}},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/147.0.0.0 Safari/537.36"
            ),
        )

        def on_response(resp):
            try:
                url = resp.url
                lower_url = url.lower()
                ctype = (resp.headers.get("content-type") or "").lower()

                if "image/" in ctype and any(x in lower_url for x in ["captcha", "verify", "valid", "code"]):
                    if not result["uuid"]:
                        result["uuid"] = extract_uuid_from_url(url)
                    if not result["path"]:
                        suffix = mimetypes.guess_extension(ctype.split(";")[0]) or ".png"
                        result["path"] = save_bytes(resp.body(), suffix)

                if "json" in ctype and any(x in lower_url for x in ["captcha", "verify", "valid", "code"]):
                    try:
                        js = resp.json()
                    except Exception:
                        return

                    def find_in(obj, keys):
                        if isinstance(obj, dict):
                            for k in keys:
                                if k in obj and obj[k]:
                                    return str(obj[k])
                            for v in obj.values():
                                r = find_in(v, keys)
                                if r:
                                    return r
                        elif isinstance(obj, list):
                            for v in obj:
                                r = find_in(v, keys)
                                if r:
                                    return r
                        return ""

                    uuid = find_in(js, ["captchaUuid", "uuid", "captchaId", "captchaKey", "key"])
                    img = find_in(js, ["captchaImg", "captchaImage", "image", "img", "base64", "captchaBase64"])
                    if uuid and not result["uuid"]:
                        result["uuid"] = uuid
                    if img and not result["path"]:
                        import base64
                        if img.startswith("data:image"):
                            img = img.split(",", 1)[1]
                        content = base64.b64decode(img)
                        result["path"] = save_bytes(content, ".png")
            except Exception:
                pass

        page.on("response", on_response)
        page.goto(HOME_URL, wait_until="domcontentloaded", timeout=45000)
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        page.wait_for_timeout(3000)

        # DOM 兜底：找页面上的验证码图片。
        if not result["path"]:
            imgs = page.query_selector_all("img")
            candidates = []
            for i, img in enumerate(imgs):
                try:
                    src = img.get_attribute("src") or ""
                    alt = img.get_attribute("alt") or ""
                    cls = img.get_attribute("class") or ""
                    box = img.bounding_box()
                    if not box:
                        continue
                    mark = (src + " " + alt + " " + cls).lower()

                    if any(skip in mark for skip in ["guohui", "logo", "baidu", "hm.js"]):
                        continue

                    likely_by_name = any(x in mark for x in ["captcha", "verify", "code", "验证码"])
                    likely_by_size = 45 <= box.get("width", 0) <= 220 and 15 <= box.get("height", 0) <= 90

                    if likely_by_name or likely_by_size:
                        score = 0
                        if likely_by_name:
                            score += 10
                        if likely_by_size:
                            score += 3
                        # 验证码一般在页面中部偏上，而不是顶部 logo 或底部页脚
                        top = box.get("y", 0)
                        if 150 <= top <= 700:
                            score += 2
                        candidates.append((score, i, src))
                except Exception:
                    pass

            candidates.sort(reverse=True)
            if candidates:
                _, i, src = candidates[0]
                if src and src.startswith("data:image"):
                    import base64
                    b64 = src.split(",", 1)[1]
                    result["path"] = save_bytes(base64.b64decode(b64), ".png")
                else:
                    if src and not result["uuid"]:
                        result["uuid"] = extract_uuid_from_url(src)
                    img = imgs[i]
                    path = CAPTCHA_DIR / ("shandong_captcha_%d.png" % int(time.time()))
                    img.screenshot(path=str(path))
                    result["path"] = str(path)

        browser.close()

except Exception as e:
    result["error"] = str(e)

print(json.dumps(result, ensure_ascii=False))
""".format(home_url=HOME_URL, captcha_dir=str(CAPTCHA_DIR))

    try:
        with tempfile.NamedTemporaryFile("w", suffix="_shandong_captcha.py", delete=False, encoding="utf-8") as f:
            f.write(child_code)
            temp_script = f.name

        proc = subprocess.run(
            [sys.executable, temp_script],
            capture_output=True,
            text=True,
            timeout=70,
        )

        stdout = (proc.stdout or "").strip()
        if not stdout:
            logging.info("验证码子进程没有输出：stderr=%s", (proc.stderr or "")[:300])
            return "", None

        # 取最后一行 JSON，避免 Playwright 可能输出其它提示。
        last_line = stdout.splitlines()[-1]
        data = json.loads(last_line)
        if data.get("error"):
            logging.info("验证码子进程提示：%s", data.get("error"))

        img_path = data.get("path") or ""
        uuid = data.get("uuid") or ""

        if img_path and Path(img_path).exists():
            logging.info("已通过页面截取验证码图片：%s", img_path)
            return uuid, Path(img_path)

    except Exception as e:
        logging.info("通过子进程截取验证码失败：%s", e)
    finally:
        try:
            if "temp_script" in locals():
                Path(temp_script).unlink(missing_ok=True)
        except Exception:
            pass

    return "", None


def get_captcha_from_terminal(session: requests.Session, label: str) -> Tuple[str, str]:
    """
    获取山东验证码。

    山东验证码接口有两种情况：
    1. 返回 JSON：可自动解析 captchaUuid + 验证码图片，用户只输入图片验证码；
    2. 直接返回图片：captchaUuid 可能在响应头/Cookie 中；如果解析不到，必须手动从 F12 的
       getListByCode 请求里复制 captchaUuid。不能用 captchaCode 冒充 captchaUuid，否则一定会校验失败。
    """
    for attempt in range(1, 8):
        uuid, image_path = try_auto_get_captcha(session)

        if image_path:
            logging.info("验证码图片已保存：%s", image_path)
            open_file_for_user(image_path)
            print(f"\n验证码图片已保存并尝试自动打开：{image_path}")

            if not uuid:
                print("注意：本次验证码接口只返回了图片，程序没有解析到 captchaUuid。")
                print("请在浏览器 F12 的 getListByCode 请求中复制 captchaUuid。")
                manual_uuid = input(f"请输入 captchaUuid（不是图片验证码；{label}；输入 q 跳过）：").strip()
                if manual_uuid.lower() == "q":
                    return "", ""
                if not manual_uuid:
                    continue
                uuid = manual_uuid

            print("请直接输入图片中的验证码；看不清可直接回车刷新。")
            code = input(f"请输入图片验证码 captchaCode（{label}，输入 q 跳过）：").strip()
            if code.lower() == "q":
                return "", ""
            if not code:
                continue

            return uuid, code

        if uuid:
            print(f"\n已获取 captchaUuid：{uuid}")
            code = input(f"请输入网页上看到的图片验证码 captchaCode（{label}，输入 q 跳过）：").strip()
            if code.lower() == "q":
                return "", ""
            if not code:
                continue
            return uuid, code

        print("\n未能自动获取验证码图片和 captchaUuid。")
        print("请打开浏览器访问：", HOME_URL)
        print("在页面搜索一次后，从 F12 的 getListByCode 请求中复制 captchaUuid。")
        manual_uuid = input(f"请输入 captchaUuid（不是图片验证码；{label}；输入 q 跳过）：").strip()
        if manual_uuid.lower() == "q":
            return "", ""
        if not manual_uuid:
            continue

        manual_code = input(f"请输入图片中的验证码 captchaCode（{label}，输入 q 跳过）：").strip()
        if manual_code.lower() == "q":
            return "", ""
        if not manual_code:
            continue

        return manual_uuid, manual_code

    raise RuntimeError(f"多次获取验证码失败：{label}")

def detect_captcha_error(js: Dict[str, Any]) -> bool:
    text = json.dumps(js, ensure_ascii=False)
    return any(x in text for x in ["验证码", "captcha", "Captcha", "校验码"]) and any(
        x in text for x in ["错误", "不正确", "失效", "为空", "过期", "invalid", "error"]
    )


# ========================
# 4. 列表接口处理
# ========================

def find_records_and_total(js: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], int]:
    """
    兼容常见返回结构：
    data.list / data.records / data.rows / data.data / rows / list
    """
    possible_blocks: List[Any] = [
        js,
        js.get("data") if isinstance(js, dict) else None,
        js.get("result") if isinstance(js, dict) else None,
    ]

    records: List[Dict[str, Any]] = []
    total = 0

    for block in possible_blocks:
        if not isinstance(block, dict):
            continue

        for total_key in ["total", "totalCount", "count", "recordsTotal"]:
            if total_key in block:
                try:
                    total = int(block.get(total_key) or 0)
                    break
                except Exception:
                    pass

        for list_key in ["records", "list", "rows", "data", "items", "content"]:
            val = block.get(list_key)
            if isinstance(val, list):
                records = [x for x in val if isinstance(x, dict)]
                if records:
                    if not total:
                        total = len(records)
                    return records, total
            if isinstance(val, dict):
                # 有些接口 data 里面再嵌一层 records/list
                for sub_key in ["records", "list", "rows", "items"]:
                    sub = val.get(sub_key)
                    if isinstance(sub, list):
                        records = [x for x in sub if isinstance(x, dict)]
                        if records:
                            if not total:
                                total = int(val.get("total") or val.get("count") or len(records))
                            return records, total

    return records, total


def make_list_payload(
    task: Dict[str, str],
    keyword: str,
    page: int,
    start_date: str,
    end_date: str,
    captcha_uuid: str,
    captcha_code: str,
) -> Dict[str, Any]:
    return {
        "colCode": task["colCode"],
        "area": task.get("area", ""),
        "title": keyword,
        "projectCode": "",
        "currentPage": page,
        "pageSize": PAGE_SIZE,
        "buyKind": "",
        "buyType": "",
        "startTime": f"{start_date} 00:00:00",
        "oldData": 0,
        "endTime": f"{end_date} 23:59:59",
        "homePage": 0,
        "mergeType": 0,
        "projectType": "",
        "unitName": "",
        "captchaUuid": captcha_uuid,
        "captchaCode": captcha_code,
    }


def post_list(
    session: requests.Session,
    task: Dict[str, str],
    keyword: str,
    page: int,
    start_date: str,
    end_date: str,
    captcha_uuid: str,
    captcha_code: str,
) -> Tuple[List[Dict[str, Any]], int, Dict[str, Any]]:
    payload = make_list_payload(task, keyword, page, start_date, end_date, captcha_uuid, captcha_code)
    js = request_json(session, "POST", LIST_API, json=payload)
    records, total = find_records_and_total(js)
    return records, total, js


def post_list_with_captcha_retry(
    session: requests.Session,
    task: Dict[str, str],
    keyword: str,
    page: int,
    start_date: str,
    end_date: str,
    captcha_uuid: str,
    captcha_code: str,
    category_label: str,
    max_captcha_retry: int = 3,
) -> Tuple[List[Dict[str, Any]], int, Dict[str, Any], str, str, bool]:
    """
    请求某一页列表。如果接口提示验证码失效/错误，则重新获取验证码，并重试“当前页”。

    返回：
        records, total, js, captcha_uuid, captcha_code, ok
    """
    for retry_idx in range(max_captcha_retry + 1):
        records, total, js = post_list(
            session=session,
            task=task,
            keyword=keyword,
            page=page,
            start_date=start_date,
            end_date=end_date,
            captcha_uuid=captcha_uuid,
            captcha_code=captcha_code,
        )

        if not detect_captcha_error(js):
            return records, total, js, captcha_uuid, captcha_code, True

        logging.warning(
            "%s / %s / 第 %d 页提示验证码失效或校验失败，需要重新输入验证码后继续当前页。",
            category_label,
            keyword,
            page,
        )

        captcha_uuid, captcha_code = get_captcha_from_terminal(
            session,
            f"{category_label}/{keyword}/第{page}页重新验证"
        )

        if not captcha_uuid or not captcha_code:
            logging.warning("%s / %s / 第 %d 页未输入有效验证码，停止当前栏目。", category_label, keyword, page)
            return records, total, js, captcha_uuid, captcha_code, False

    logging.warning("%s / %s / 第 %d 页验证码重试次数过多，停止当前栏目。", category_label, keyword, page)
    return records, total, js, captcha_uuid, captcha_code, False



def get_item_id(item: Dict[str, Any]) -> str:
    return get_by_keys(item, [
        "id", "ID", "noticeId", "articleId", "contentId", "uuid", "guid", "dataId",
        "pkId", "rowId", "mainId"
    ])


def get_item_title(item: Dict[str, Any]) -> str:
    return get_by_keys(item, [
        "title", "noticeTitle", "articleTitle", "name", "projectName", "projName",
        "caption", "heading"
    ])


def get_item_date(item: Dict[str, Any]) -> str:
    return get_by_keys(item, [
        "publishTime", "publishDate", "publish_time", "pubDate", "createTime",
        "releaseTime", "date", "time"
    ])


def get_item_unit(item: Dict[str, Any]) -> str:
    return get_by_keys(item, [
        "unitName", "buyerName", "purchaseName", "purchaser", "buyer", "orgName",
        "publishUser", "publisher"
    ])


def get_item_project_code(item: Dict[str, Any]) -> str:
    return get_by_keys(item, ["projectCode", "projCode", "projectNo", "code"])


def build_detail_url(item: Dict[str, Any], task: Dict[str, str]) -> str:
    item_id = get_item_id(item)
    old_data = get_by_keys(item, ["oldData", "olddata"]) or "0"
    col_code = task["colCode"]
    if not item_id:
        return ""
    return f"{BASE_URL}/detail?id={quote_plus(item_id)}&colCode={quote_plus(col_code)}&oldData={quote_plus(str(old_data))}"


# ========================
# 5. 详情解析
# ========================

def text_first_match(text: str, patterns: Iterable[str]) -> str:
    for pat in patterns:
        m = re.search(pat, text, re.S)
        if m:
            val = m.group(1) if m.groups() else m.group(0)
            val = clean_text(val)
            val = re.split(
                r"(?:采购人信息|采购代理机构信息|供应商名称|供应商地址|中标金额|成交金额|预算金额|项目编号|项目名称|联系方式|二、|三、|四、|五、|六、)",
                val,
                maxsplit=1,
            )[0].strip(" ：:，,。；;")
            if val:
                return val
    return ""


def normalize_amount(num: str, unit: str = "") -> str:
    if not num:
        return ""
    s = str(num).replace(",", "").replace("￥", "").replace("¥", "").strip()
    try:
        v = float(s)
        if unit in ["万元", "万"]:
            v *= 10000
        if v.is_integer():
            return str(int(v))
        return f"{v:.2f}".rstrip("0").rstrip(".")
    except Exception:
        return s + unit


def parse_amount(text: str, labels: Iterable[str]) -> str:
    label_part = "|".join(map(re.escape, labels))
    patterns = [
        rf"(?:{label_part})\s*[:：]?\s*[￥¥]?\s*([0-9][0-9,]*(?:\.\d+)?)\s*(万元|万|元)?",
        rf"(?:{label_part})[^0-9]{{0,15}}[￥¥]?\s*([0-9][0-9,]*(?:\.\d+)?)\s*(万元|万|元)?",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.S)
        if m:
            return normalize_amount(m.group(1), m.group(2) or "")
    return ""


def parse_table_rows(soup: BeautifulSoup) -> List[List[str]]:
    rows: List[List[str]] = []
    for tr in soup.find_all("tr"):
        cells = [clean_text(td.get_text(" ", strip=True)) for td in tr.find_all(["td", "th"])]
        cells = [c for c in cells if c]
        if cells:
            rows.append(cells)
    return rows


def parse_intention_items(soup: BeautifulSoup) -> List[Dict[str, str]]:
    """
    采购意向详情一般是表格：
    序号、采购项目名称、采购需求概况、预算金额（万元）、预计采购时间、备注
    """
    rows = parse_table_rows(soup)
    items: List[Dict[str, str]] = []
    if len(rows) < 2:
        return items

    header_idx = -1
    header = []
    for i, row in enumerate(rows):
        joined = "".join(row)
        if "采购项目名称" in joined and "预算金额" in joined:
            header_idx = i
            header = row
            break

    if header_idx < 0:
        return items

    for row in rows[header_idx + 1:]:
        if len(row) < 3:
            continue

        # 兼容 header 被拆分或金额单位单独成列的情况
        seq = row[0] if row else ""
        name = row[1] if len(row) > 1 else ""
        desc = row[2] if len(row) > 2 else ""
        budget = row[3] if len(row) > 3 else ""
        reserve = row[4] if len(row) > 4 else ""
        expected = row[5] if len(row) > 5 else ""
        remark = row[6] if len(row) > 6 else ""

        if not name or "采购项目名称" in name:
            continue

        items.append({
            "intention_seq": seq,
            "intention_project_name": name,
            "intention_need": desc,
            "intention_budget_amount": normalize_amount(budget, "万元") if re.fullmatch(r"[0-9,.]+", budget) else budget,
            "intention_reserve_sme": reserve,
            "intention_expected_time": expected,
            "intention_remark": remark,
        })

    return items


def parse_detail_html(html: str, category: str = "") -> Tuple[Dict[str, str], List[Dict[str, str]]]:
    soup = BeautifulSoup(html or "", "lxml")

    # 去掉翻译插件样式、脚本等
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    title = ""
    node = soup.select_one(".site-title")
    if node:
        title = clean_text(node.get_text(" ", strip=True))
    if not title and soup.title:
        title = clean_text(soup.title.get_text(" ", strip=True))

    text = normalize_text(soup.get_text("\n", strip=True))
    one = clean_text(text)

    publish_time = text_first_match(one, [
        r"发布时间\s*[:：]\s*([0-9]{4}[-年/][0-9]{1,2}[-月/][0-9]{1,2}(?:\s+[0-9]{1,2}[:：][0-9]{1,2}(?::[0-9]{1,2})?)?)",
        r"发布日期\s*[:：]\s*([0-9]{4}[-年/][0-9]{1,2}[-月/][0-9]{1,2}(?:\s+[0-9]{1,2}[:：][0-9]{1,2}(?::[0-9]{1,2})?)?)",
    ])
    publisher = text_first_match(one, [
        r"发布人\s*[:：]\s*([^\s]+(?:\s*[^\s]+){0,8})",
        r"采购人\s*[:：]\s*([^\n\r]+)",
    ])

    project_code = text_first_match(one, [
        r"项目编号\s*[:：]\s*([A-Za-z0-9_\-（）()【】\[\].]+)",
        r"采购项目编号\s*[:：]\s*([A-Za-z0-9_\-（）()【】\[\].]+)",
    ])

    project_name = text_first_match(one, [
        r"项目名称\s*[:：]\s*([^\n\r]+?)\s*(?:采购方式|预算金额|项目编号|二、|三、)",
        r"采购项目名称\s*[:：]\s*([^\n\r]+?)\s*(?:采购方式|预算金额|项目编号|二、|三、)",
    ])

    purchase_name = text_first_match(one, [
        r"采购人信息\s*名\s*称\s*[:：]\s*([^\n\r]+)",
        r"采购单位信息\s*名\s*称\s*[:：]\s*([^\n\r]+)",
        r"采购人\s*[:：]\s*([^\n\r；;，,。]+)",
        r"采购单位\s*[:：]\s*([^\n\r；;，,。]+)",
    ])

    agency_name = text_first_match(one, [
        r"采购代理机构信息\s*名\s*称\s*[:：]\s*([^\n\r]+)",
        r"代理机构信息\s*名\s*称\s*[:：]\s*([^\n\r]+)",
        r"采购代理机构\s*[:：]\s*([^\n\r；;，,。]+)",
        r"代理机构\s*[:：]\s*([^\n\r；;，,。]+)",
    ])

    supplier_name = text_first_match(one, [
        r"供应商名称\s*[:：]\s*([^\n\r；;，,。]+)",
        r"中标供应商\s*[:：]\s*([^\n\r；;，,。]+)",
        r"成交供应商\s*[:：]\s*([^\n\r；;，,。]+)",
        r"中标人\s*[:：]\s*([^\n\r；;，,。]+)",
        r"成交人\s*[:：]\s*([^\n\r；;，,。]+)",
    ])

    budget_amount = parse_amount(one, ["预算金额", "采购包预算金额", "项目预算", "最高限价"])
    win_amount = parse_amount(one, ["中标金额", "成交金额", "中标（成交）金额", "中标(成交)金额", "报价金额", "总价", "投标报价"])

    procurement_method = text_first_match(one, [
        r"采购方式\s*[:：]\s*([^\n\r；;，,。]+)",
    ])

    intention_items = parse_intention_items(soup) if "意向" in category else []

    detail = {
        "detail_title": title,
        "detail_publish_time": publish_time,
        "publisher": publisher,
        "project_code_detail": project_code,
        "project_name_detail": project_name,
        "purchase_name_detail": purchase_name,
        "agency_name_detail": agency_name,
        "supplier_name_detail": supplier_name,
        "budget_amount_detail": budget_amount,
        "win_amount_detail": win_amount,
        "procurement_method_detail": procurement_method,
        "summary": clean_text(one[:800]),
    }
    return detail, intention_items


def detail_has_rendered_content(html: str) -> bool:
    if not html:
        return False
    text = clean_text(html)
    return any(x in text for x in ["公告正文", "发布时间", "采购人信息", "项目编号", "预算金额", "政府采购意向"])


def fetch_detail_by_requests(session: requests.Session, detail_url: str) -> str:
    return request_text(session, detail_url)


class PlaywrightRenderer:
    def __init__(self) -> None:
        self.enabled = False
        self.playwright = None
        self.browser = None
        self.page = None

        if not USE_PLAYWRIGHT_DETAIL_FALLBACK:
            return

        try:
            from playwright.sync_api import sync_playwright  # type: ignore
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage", "--no-proxy-server"],
            )
            self.page = self.browser.new_page(
                viewport={"width": 1366, "height": 900},
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/147.0.0.0 Safari/537.36"
                ),
            )
            self.page.set_default_timeout(45000)
            self.enabled = True
            logging.info("Playwright 详情渲染已启用。")
        except Exception as e:
            logging.warning("Playwright 不可用，详情页只使用 requests 抓取：%s", e)

    def render(self, url: str) -> str:
        if not self.enabled or not self.page:
            return ""
        try:
            self.page.goto(url, wait_until="domcontentloaded", timeout=45000)
            try:
                self.page.wait_for_load_state("networkidle", timeout=12000)
            except Exception:
                pass
            time.sleep(1.0)
            return self.page.content()
        except Exception as e:
            logging.warning("Playwright 渲染详情失败：%s；错误：%s", url, e)
            return ""

    def close(self) -> None:
        try:
            if self.browser:
                self.browser.close()
        except Exception:
            pass
        try:
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass


def fetch_and_parse_detail(
    session: requests.Session,
    renderer: Optional[PlaywrightRenderer],
    detail_url: str,
    category: str,
) -> Tuple[Dict[str, str], List[Dict[str, str]], str]:
    if not detail_url:
        return {}, [], "无详情链接"

    html = ""
    status = "成功"

    try:
        html = fetch_detail_by_requests(session, detail_url)
    except Exception as e:
        status = f"requests详情失败：{e}"

    # Vue 页面如果 requests 返回的是空壳，用 Playwright 渲染
    if (not detail_has_rendered_content(html)) and renderer and renderer.enabled:
        rendered = renderer.render(detail_url)
        if detail_has_rendered_content(rendered):
            html = rendered
            status = "成功"
        elif rendered:
            html = rendered
            status = "详情可能未完全渲染"

    if not html:
        return {}, [], status

    try:
        detail, intention_items = parse_detail_html(html, category)
        return detail, intention_items, status
    except Exception as e:
        raw_path = RAW_DIR / f"detail_parse_fail_{int(time.time())}.html"
        raw_path.write_text(html, encoding="utf-8", errors="ignore")
        return {}, [], f"详情解析失败：{e}；raw={raw_path.name}"


# ========================
# 6. 记录与导出
# ========================

def load_visited() -> set[str]:
    if not VISITED_PATH.exists():
        return set()
    return set(x.strip() for x in VISITED_PATH.read_text(encoding="utf-8").splitlines() if x.strip())


def save_visited(url: str) -> None:
    if not url:
        return
    with VISITED_PATH.open("a", encoding="utf-8") as f:
        f.write(url + "\n")


def write_jsonl(record: Dict[str, Any]) -> None:
    with JSONL_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def read_jsonl() -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if not JSONL_PATH.exists():
        return rows
    for line in JSONL_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            rows.append(json.loads(line))
        except Exception:
            pass
    return rows


def base_record_from_item(
    item: Dict[str, Any],
    task: Dict[str, str],
    keyword: str,
    detail_url: str,
) -> Dict[str, Any]:
    return {
        "省份": "山东",
        "来源网站": "山东省政府采购信息公开平台",
        "关键词": keyword,
        "公告大类": task["category"],
        "层级": task["level"],
        "栏目编码": task["colCode"],
        "标题": get_item_title(item),
        "发布时间": get_item_date(item),
        "地区": get_by_keys(item, ["areaName", "regionName", "districtName", "cityName", "area"]),
        "采购人": get_item_unit(item),
        "代理机构": get_by_keys(item, ["agencyName", "agentName", "agency"]),
        "项目编号": get_item_project_code(item),
        "项目名称": get_by_keys(item, ["projectName", "projName", "projectTitle"]),
        "采购方式": get_by_keys(item, ["buyKindName", "buyTypeName", "purchaseMode", "procurementMethod"]),
        "预算金额_元": get_by_keys(item, ["budgetAmount", "budgetMoney", "budget", "amount"]),
        "中标单位": get_by_keys(item, ["supplierName", "winningSupplier", "bidSupplier", "winner"]),
        "中标金额_元": get_by_keys(item, ["winAmount", "winningAmount", "bidAmount", "dealAmount", "totalAmount"]),
        "采购意向项目名称": "",
        "采购需求概况": "",
        "意向预算金额_元": "",
        "预计采购时间": "",
        "详情链接": detail_url,
        "详情状态": "",
        "正文摘要": "",
        "抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def merge_detail(record: Dict[str, Any], detail: Dict[str, str]) -> Dict[str, Any]:
    if detail.get("detail_title"):
        record["标题"] = detail["detail_title"]
    if detail.get("detail_publish_time") and not record.get("发布时间"):
        record["发布时间"] = detail["detail_publish_time"]
    if detail.get("project_code_detail"):
        record["项目编号"] = detail["project_code_detail"]
    if detail.get("project_name_detail"):
        record["项目名称"] = detail["project_name_detail"]
    if detail.get("purchase_name_detail"):
        record["采购人"] = detail["purchase_name_detail"]
    elif detail.get("publisher") and not record.get("采购人"):
        record["采购人"] = detail["publisher"]
    if detail.get("agency_name_detail"):
        record["代理机构"] = detail["agency_name_detail"]
    if detail.get("supplier_name_detail"):
        record["中标单位"] = detail["supplier_name_detail"]
    if detail.get("budget_amount_detail"):
        record["预算金额_元"] = detail["budget_amount_detail"]
    if detail.get("win_amount_detail"):
        record["中标金额_元"] = detail["win_amount_detail"]
    if detail.get("procurement_method_detail"):
        record["采购方式"] = detail["procurement_method_detail"]
    if detail.get("summary"):
        record["正文摘要"] = detail["summary"]
    return record


def expand_intention_records(base: Dict[str, Any], intention_items: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    if not intention_items:
        return [base]

    rows: List[Dict[str, Any]] = []
    for it in intention_items:
        r = dict(base)
        r["采购意向项目名称"] = it.get("intention_project_name", "")
        r["采购需求概况"] = it.get("intention_need", "")
        r["意向预算金额_元"] = it.get("intention_budget_amount", "")
        r["预计采购时间"] = it.get("intention_expected_time", "")
        # 对采购意向，项目名称优先使用表格内的项目名
        if r["采购意向项目名称"]:
            r["项目名称"] = r["采购意向项目名称"]
        if r["意向预算金额_元"]:
            r["预算金额_元"] = r["意向预算金额_元"]
        rows.append(r)
    return rows


def export_outputs() -> None:
    rows = read_jsonl()
    if not rows:
        logging.warning("没有记录可导出。")
        return

    # 按 详情链接 + 项目名称 去重
    dedup: Dict[str, Dict[str, Any]] = {}
    for r in rows:
        key = f"{r.get('详情链接','')}|{r.get('项目名称','')}|{r.get('标题','')}"
        dedup[key] = r
    final_rows = list(dedup.values())

    final_rows.sort(key=lambda x: str(x.get("发布时间", "")), reverse=True)

    JSON_PATH.write_text(json.dumps(final_rows, ensure_ascii=False, indent=2), encoding="utf-8")

    columns = [
        "省份", "来源网站", "关键词", "公告大类", "层级", "栏目编码",
        "标题", "发布时间", "地区",
        "采购人", "代理机构",
        "项目编号", "项目名称", "采购方式",
        "预算金额_元", "中标单位", "中标金额_元",
        "采购意向项目名称", "采购需求概况", "意向预算金额_元", "预计采购时间",
        "详情链接", "详情状态", "正文摘要", "抓取时间",
    ]

    df = pd.DataFrame(final_rows)
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    df = df[columns]
    df.to_excel(EXCEL_PATH, index=False)

    logging.info("JSON 已导出：%s，共 %d 条", JSON_PATH, len(final_rows))
    logging.info("Excel 已导出：%s，共 %d 条", EXCEL_PATH, len(final_rows))


# ========================
# 7. 主流程
# ========================

def should_skip_task(task: Dict[str, str]) -> bool:
    if task["level"] == "市区县" and not INCLUDE_CITY_COUNTY:
        return True
    return False


def crawl_task_keyword(
    session: requests.Session,
    renderer: Optional[PlaywrightRenderer],
    task: Dict[str, str],
    keyword: str,
    start_date: str,
    end_date: str,
    captcha: Tuple[str, str],
    visited: set[str],
) -> Tuple[int, Tuple[str, str]]:
    category_label = f"{task['category']}-{task['level']}"
    captcha_uuid, captcha_code = captcha

    if not captcha_uuid or not captcha_code or CAPTCHA_EVERY_KEYWORD:
        captcha_uuid, captcha_code = get_captcha_from_terminal(session, f"{category_label}/{keyword}")
        if not captcha_uuid or not captcha_code:
            logging.warning("跳过：%s/%s", category_label, keyword)
            return 0, (captcha_uuid, captcha_code)

    total_saved = 0

    # 第一页
    # 对采购意向-市区县这类栏目，山东站可能使用不同 colCode。
    # 如果配置了 colCode_candidates，则依次测试候选编码，选择 total>0 或有列表记录的编码。
    task = dict(task)
    candidates = task.get("colCode_candidates") or [task.get("colCode")]
    candidates = [str(x) for x in candidates if str(x).strip()]

    records: List[Dict[str, Any]] = []
    total = 0
    js: Dict[str, Any] = {}
    selected_code = str(task.get("colCode", ""))

    for idx, code_candidate in enumerate(candidates, start=1):
        task["colCode"] = code_candidate
        try:
            records, total, js, captcha_uuid, captcha_code, ok = post_list_with_captcha_retry(
                session=session,
                task=task,
                keyword=keyword,
                page=1,
                start_date=start_date,
                end_date=end_date,
                captcha_uuid=captcha_uuid,
                captcha_code=captcha_code,
                category_label=category_label,
                max_captcha_retry=3,
            )
            if not ok:
                return 0, (captcha_uuid, captcha_code)
        except Exception as e:
            logging.warning("列表请求失败：%s/%s，候选colCode=%s，第1页，错误：%s",
                            category_label, keyword, code_candidate, e)
            continue

        if records or total:
            selected_code = code_candidate
            logging.info("%s / 关键词=%s 选用 colCode=%s", category_label, keyword, selected_code)
            break

        # 没有候选编码时，只测试一次；有候选编码时继续测试下一个
        if len(candidates) == 1:
            selected_code = code_candidate
            break
    else:
        # 全部候选都没有数据时，保留第一个候选编码，便于保存 raw 排查
        selected_code = candidates[0] if candidates else str(task.get("colCode", ""))
        task["colCode"] = selected_code

    page_count = max(1, math.ceil(total / PAGE_SIZE)) if total else 1
    page_count = min(page_count, MAX_PAGES_PER_QUERY)

    logging.info("%s / 关键词=%s / total=%s / 页数=%s", category_label, keyword, total, page_count)

    for page in range(1, page_count + 1):
        if page == 1:
            page_records = records
        else:
            try:
                page_records, _, js, captcha_uuid, captcha_code, ok = post_list_with_captcha_retry(
                    session=session,
                    task=task,
                    keyword=keyword,
                    page=page,
                    start_date=start_date,
                    end_date=end_date,
                    captcha_uuid=captcha_uuid,
                    captcha_code=captcha_code,
                    category_label=category_label,
                    max_captcha_retry=3,
                )
                if not ok:
                    break
            except Exception as e:
                logging.warning("列表请求失败：%s/%s，第%d页，错误：%s", category_label, keyword, page, e)
                continue

        logging.info("处理 %s / %s / 第 %d 页：%d 条", category_label, keyword, page, len(page_records))

        if not page_records:
            # 保存原始返回，便于排查字段结构
            raw_path = RAW_DIR / f"empty_{safe_filename(category_label)}_{safe_filename(keyword)}_p{page}.json"
            raw_path.write_text(json.dumps(js, ensure_ascii=False, indent=2), encoding="utf-8")
            logging.warning("第 %d 页没有解析到记录，原始响应已保存：%s", page, raw_path)
            if page > 1:
                logging.info("第 %d 页为空，判断为已到末页或接口未返回更多数据，停止当前栏目翻页。", page)
                break
            continue

        for item in page_records:
            detail_url = build_detail_url(item, task)
            title = get_item_title(item) or detail_url

            if detail_url and detail_url in visited:
                logging.info("跳过已抓取：%s", title)
                continue

            logging.info("详情：%s", title)
            base = base_record_from_item(item, task, keyword, detail_url)

            detail, intention_items, detail_status = fetch_and_parse_detail(session, renderer, detail_url, task["category"])
            base = merge_detail(base, detail)
            base["详情状态"] = detail_status

            rows = expand_intention_records(base, intention_items)

            for r in rows:
                write_jsonl(r)
                total_saved += 1

            if detail_url:
                save_visited(detail_url)
                visited.add(detail_url)

            sleep_delay(DETAIL_DELAY)

        export_outputs()
        sleep_delay(REQUEST_DELAY)

    return total_saved, (captcha_uuid, captcha_code)


def main() -> None:
    cfg = load_config()
    ensure_dirs()
    setup_logging()

    keywords = get_keywords(cfg)
    start_date, end_date = get_date_range(cfg)

    logging.info("开始爬取山东省政府采购信息公开平台")
    logging.info("关键词：%s", keywords)
    logging.info("时间范围：%s 至 %s", start_date, end_date)
    logging.info("是否抓取市区县：%s", INCLUDE_CITY_COUNTY)

    session = make_session()
    visited = load_visited()
    renderer: Optional[PlaywrightRenderer] = None

    if USE_PLAYWRIGHT_DETAIL_FALLBACK:
        renderer = PlaywrightRenderer()

    captcha: Tuple[str, str] = ("", "")

    total_saved = 0

    try:
        # 先访问首页，建立会话
        try:
            session.get(HOME_URL, timeout=TIMEOUT)
        except Exception:
            pass

        for keyword in keywords:
            # 如果配置要求每个关键词重新验证码，则清空
            if CAPTCHA_EVERY_KEYWORD:
                captcha = ("", "")

            for task in TASKS:
                if should_skip_task(task):
                    continue

                saved, captcha = crawl_task_keyword(
                    session=session,
                    renderer=renderer,
                    task=task,
                    keyword=keyword,
                    start_date=start_date,
                    end_date=end_date,
                    captcha=captcha,
                    visited=visited,
                )
                total_saved += saved

        export_outputs()
        logging.info("运行完成，本次新增记录：%d 条", total_saved)

    except KeyboardInterrupt:
        logging.warning("用户中断，正在导出已有数据。")
        export_outputs()
    finally:
        if renderer:
            renderer.close()


if __name__ == "__main__":
    main()
