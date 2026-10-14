#!/usr/bin/env bash
# 先跑官方範例測試，通過才打包 Claude 網頁版要上傳的 zip，輸出到 dist/。
set -euo pipefail
cd "$(dirname "$0")"
python3 tests/test_official_examples.py
mkdir -p dist
rm -f dist/shopee-fee-calc.zip
find skills -name "__pycache__" -prune -exec rm -rf {} +
(cd skills && zip -rq ../dist/shopee-fee-calc.zip shopee-fee-calc)
echo "已打包：dist/shopee-fee-calc.zip"
