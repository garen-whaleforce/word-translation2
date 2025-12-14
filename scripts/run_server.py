#!/usr/bin/env python3
"""啟動伺服器的腳本，自動載入 .env"""
import os
import sys
from pathlib import Path

# 專案根目錄
ROOT = Path(__file__).parent.parent
BACKEND = ROOT / "backend"
ENV_FILE = ROOT / ".env"

# 載入 .env
if ENV_FILE.exists():
    with open(ENV_FILE) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                key, val = line.split('=', 1)
                os.environ[key] = val
    print(f"已載入環境變數：{ENV_FILE}")

# 切換到 backend 目錄
os.chdir(BACKEND)
sys.path.insert(0, str(BACKEND))

# 啟動伺服器
import uvicorn

port = int(os.environ.get("PORT", 8000))
print(f"啟動伺服器 http://localhost:{port}")

uvicorn.run(
    "main:app",
    host="0.0.0.0",
    port=port,
    reload=False
)
