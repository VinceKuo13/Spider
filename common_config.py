# -*- coding: utf-8 -*-
"""
通用配置读取模块。

各省爬虫仍然独立运行，例如：
    python beijing.py
    python tianjin.py

客户只需要修改同目录下的 config.yaml 中的：
    keywords
    start_date
    end_date
其他参数仍保留在各省 .py 文件顶部。
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any, Dict, List

try:
    import yaml  # type: ignore
except Exception:  # pragma: no cover
    yaml = None

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.yaml"


def _as_list(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    s = str(value).strip()
    return [s] if s else []


def read_common_config() -> Dict[str, Any]:
    """读取 config.yaml。end_date 为空时自动使用运行当天。"""
    if not CONFIG_PATH.exists():
        return {}
    if yaml is None:
        raise RuntimeError("检测到 config.yaml，但当前环境没有 pyyaml，请先执行：pip install pyyaml")
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}

    out: Dict[str, Any] = {}
    keywords = _as_list(cfg.get("keywords") or cfg.get("keyword"))
    if keywords:
        out["keywords"] = keywords

    start_date = str(cfg.get("start_date") or "").strip()
    if start_date:
        out["start_date"] = start_date

    # end_date 留空表示当天。
    end_date = str(cfg.get("end_date") or "").strip()
    out["end_date"] = end_date or date.today().strftime("%Y-%m-%d")
    return out


def apply_common_config(namespace: Dict[str, Any]) -> Dict[str, Any]:
    """把 config.yaml 里的关键词、日期覆盖到各省脚本的全局变量。"""
    cfg = read_common_config()
    if not cfg:
        return cfg

    if "keywords" in cfg:
        namespace["KEYWORDS"] = cfg["keywords"]

    if "start_date" in cfg:
        if "START_DATE" in namespace:
            namespace["START_DATE"] = cfg["start_date"]
        if "BEGIN_DATE" in namespace:
            namespace["BEGIN_DATE"] = cfg["start_date"]

    if "end_date" in cfg:
        if "END_DATE" in namespace:
            namespace["END_DATE"] = cfg["end_date"]

    return cfg
