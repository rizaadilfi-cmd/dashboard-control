#!/usr/bin/env bash
# ============================================================
# PT NIFCO INDONESIA - MILKRUN DASHBOARD LAUNCHER
# Linux Terminal / macOS
# ============================================================

set -e

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "============================================================"
echo "       PT NIFCO INDONESIA - MILKRUN DASHBOARD LAUNCHER"
echo "                 Linux / macOS Terminal"
echo "============================================================"
echo ""

# 1. Deteksi Python 3
echo "[*] Mendeteksi Python di sistem..."
if command -v python3 &>/dev/null; then
    PY_CMD=python3
elif command -v python &>/dev/null; then
    PY_CMD=python
else
    echo "[ERROR] Python tidak terdeteksi di sistem!"
    echo "Untuk Ubuntu/Debian, jalankan: sudo apt update && sudo apt install python3 python3-pip python3-venv -y"
    echo "Untuk CentOS/RHEL, jalankan: sudo dnf install python3 python3-pip -y"
    exit 1
fi

PY_VER=$($PY_CMD --version 2>&1)
echo "[OK] Menggunakan: $PY_VER ($PY_CMD)"

# 2. Jalankan launcher.py
$PY_CMD "$DIR/launcher.py"
