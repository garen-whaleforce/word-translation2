# ==============================================
# CB to CNS Report Generator - Dockerfile
# 適用於 Zeabur 部署
# ==============================================

# 使用官方 Python slim image 作為基礎
FROM python:3.11-slim

# 設定環境變數
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# 設定工作目錄
WORKDIR /app

# 安裝系統依賴（如果需要的話）
# python-docx 不需要額外的系統套件
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# 複製 requirements 並安裝依賴
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製應用程式程式碼
COPY backend/ ./backend/
COPY templates/ ./templates/

# 建立暫存目錄
RUN mkdir -p /tmp/reports

# 設定工作目錄到 backend
WORKDIR /app/backend

# 暴露 port（Zeabur 會自動設定 PORT 環境變數）
EXPOSE 8000

# 啟動命令
# 使用 $PORT 環境變數，讓 Zeabur 可以控制 port
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]
