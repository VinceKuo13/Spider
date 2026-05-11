#!/usr/bin/env bash
set -e

echo "[1/3] 安装 Python 依赖..."
pip install -r requirements.txt

echo "[2/3] 安装 Playwright Chromium（新疆爬虫需要；其他省份不会用到）..."
python -m playwright install --with-deps chromium || python -m playwright install chromium

echo "[3/3] 环境准备完成。"
echo "现在可以运行：python tianjin.py 或 python beijing.py 等。"
