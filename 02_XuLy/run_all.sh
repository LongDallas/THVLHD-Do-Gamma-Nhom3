#!/usr/bin/env bash
# Chạy toàn bộ quy trình: phân tích → sổ số liệu Excel → báo cáo Word.
# Dùng: ./02_XuLy/run_all.sh   (chạy từ bất kỳ đâu)
set -euo pipefail
cd "$(dirname "$0")"
[ -d node_modules/docx ] || npm install --silent
python3 analysis.py
python3 build_excel.py
node build_report.js
