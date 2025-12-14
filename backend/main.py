"""
==============================================
CB to CNS Report Generator - FastAPI Application
==============================================

主要入口點：提供 API endpoint 將 CB PDF 報告轉換為 CNS Word 報告

Endpoints:
- GET /          : 簡易上傳頁面
- POST /generate-report : 接收 PDF，回傳填好的 Word 檔案
- GET /health    : 健康檢查
"""

import os
import uuid
import tempfile
from datetime import datetime
from typing import Optional
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

# 確保可以 import backend 內的模組
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import settings
from utils.logger import get_logger, setup_logging
from services.adobe_extract import extract_pdf_to_json, create_mock_extract_result, AdobeExtractError
from services.azure_llm import extract_report_schema_from_adobe_json, create_mock_schema
from services.word_filler import fill_cns_template

# 設定 logging
setup_logging()
logger = get_logger(__name__)


# ==============================================
# Lifespan Management
# ==============================================

@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    應用程式生命週期管理
    """
    # Startup
    logger.info("=" * 50)
    logger.info(f"啟動 {settings.app_name}")
    logger.info("=" * 50)

    # 確保暫存目錄存在
    os.makedirs(settings.temp_dir, exist_ok=True)
    logger.info(f"暫存目錄: {settings.temp_dir}")

    # 確保模板目錄存在
    template_dir = os.path.join(os.path.dirname(__file__), "..", settings.template_dir)
    if not os.path.exists(template_dir):
        os.makedirs(template_dir, exist_ok=True)
        logger.warning(f"模板目錄不存在，已建立: {template_dir}")

    yield

    # Shutdown
    logger.info("應用程式關閉")


# ==============================================
# FastAPI App Setup
# ==============================================

app = FastAPI(
    title=settings.app_name,
    description="將 CB Test Report PDF 轉換為 CNS Report Word 文件",
    version="1.0.0",
    lifespan=lifespan
)

# CORS 設定（允許前端跨域存取）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 在正式環境可限制為特定網域
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ==============================================
# HTML Template for Upload Page
# ==============================================

UPLOAD_PAGE_HTML = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CB → CNS 報告轉換器</title>
    <style>
        * {
            box-sizing: border-box;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        }
        body {
            max-width: 800px;
            margin: 0 auto;
            padding: 40px 20px;
            background: #f5f5f5;
        }
        h1 {
            color: #333;
            text-align: center;
            margin-bottom: 10px;
        }
        .subtitle {
            text-align: center;
            color: #666;
            margin-bottom: 40px;
        }
        .card {
            background: white;
            border-radius: 8px;
            padding: 30px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #333;
        }
        input[type="file"] {
            width: 100%;
            padding: 12px;
            border: 2px dashed #ccc;
            border-radius: 4px;
            background: #fafafa;
            cursor: pointer;
        }
        input[type="file"]:hover {
            border-color: #007bff;
        }
        button {
            width: 100%;
            padding: 14px;
            background: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.2s;
        }
        button:hover {
            background: #0056b3;
        }
        button:disabled {
            background: #ccc;
            cursor: not-allowed;
        }
        .status {
            margin-top: 20px;
            padding: 15px;
            border-radius: 4px;
            display: none;
        }
        .status.loading {
            display: block;
            background: #e3f2fd;
            color: #1565c0;
        }
        .status.success {
            display: block;
            background: #e8f5e9;
            color: #2e7d32;
        }
        .status.error {
            display: block;
            background: #ffebee;
            color: #c62828;
        }
        .spinner {
            display: inline-block;
            width: 16px;
            height: 16px;
            border: 2px solid #1565c0;
            border-top-color: transparent;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin-right: 8px;
            vertical-align: middle;
        }
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        .note {
            margin-top: 30px;
            padding: 15px;
            background: #fff3e0;
            border-radius: 4px;
            font-size: 14px;
            color: #e65100;
        }
        .checkbox-group {
            margin-top: 10px;
        }
        .checkbox-group label {
            display: flex;
            align-items: center;
            font-weight: normal;
            cursor: pointer;
        }
        .checkbox-group input[type="checkbox"] {
            margin-right: 8px;
            width: auto;
        }
    </style>
</head>
<body>
    <h1>CB → CNS 報告轉換器</h1>
    <p class="subtitle">上傳 CB Test Report PDF，自動產生 CNS 報告 Word 檔</p>

    <div class="card">
        <form id="uploadForm" enctype="multipart/form-data">
            <div class="form-group">
                <label for="pdfFile">選擇 CB Report PDF 檔案</label>
                <input type="file" id="pdfFile" name="file" accept=".pdf" required>
            </div>

            <div class="form-group checkbox-group">
                <label>
                    <input type="checkbox" id="useMock" name="use_mock">
                    使用模擬資料（測試用，不會呼叫 Adobe/Azure API）
                </label>
            </div>

            <button type="submit" id="submitBtn">開始轉換</button>
        </form>

        <div id="status" class="status"></div>
    </div>

    <div class="note">
        <strong>注意事項：</strong>
        <ul style="margin: 10px 0 0 20px; padding: 0;">
            <li>請確保 PDF 檔案為有效的 CB Test Report</li>
            <li>轉換過程可能需要 1-3 分鐘</li>
            <li>請確保 templates/ 資料夾中有 CNS Word 模板</li>
        </ul>
    </div>

    <script>
        const form = document.getElementById('uploadForm');
        const statusDiv = document.getElementById('status');
        const submitBtn = document.getElementById('submitBtn');

        form.addEventListener('submit', async (e) => {
            e.preventDefault();

            const fileInput = document.getElementById('pdfFile');
            const useMock = document.getElementById('useMock').checked;

            if (!fileInput.files.length) {
                alert('請選擇 PDF 檔案');
                return;
            }

            // 顯示 loading
            statusDiv.className = 'status loading';
            statusDiv.innerHTML = '<span class="spinner"></span>正在處理中，請稍候...';
            submitBtn.disabled = true;

            try {
                const formData = new FormData();
                formData.append('file', fileInput.files[0]);
                formData.append('use_mock', useMock ? 'true' : 'false');

                const response = await fetch('/generate-report', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(errorData.detail || '轉換失敗');
                }

                // 取得檔案名稱
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = 'CNS_Report.docx';
                if (contentDisposition) {
                    const filenameMatch = contentDisposition.match(/filename[^;=\\n]*=(['\"]?)([^'\"\\n]*)\1/);
                    if (filenameMatch) {
                        filename = decodeURIComponent(filenameMatch[2]);
                    }
                }

                // 下載檔案
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);

                statusDiv.className = 'status success';
                statusDiv.textContent = '✓ 轉換成功！檔案已開始下載。';

            } catch (error) {
                statusDiv.className = 'status error';
                statusDiv.textContent = '✗ 錯誤：' + error.message;
            } finally {
                submitBtn.disabled = false;
            }
        });
    </script>
</body>
</html>
"""


# ==============================================
# API Endpoints
# ==============================================

@app.get("/", response_class=HTMLResponse)
async def root():
    """
    首頁：提供簡易的上傳介面
    """
    return UPLOAD_PAGE_HTML


@app.get("/health")
async def health_check():
    """
    健康檢查 endpoint
    """
    return {
        "status": "healthy",
        "app_name": settings.app_name,
        "timestamp": datetime.now().isoformat()
    }


@app.post("/generate-report")
async def generate_report(
    file: UploadFile = File(..., description="CB Report PDF 檔案"),
    use_mock: str = Form(default="false", description="是否使用模擬資料")
):
    """
    主要 API：將 CB PDF 轉換為 CNS Word 報告

    流程：
    1. 讀取上傳的 PDF 檔案
    2. 呼叫 Adobe PDF Extract API 萃取內容
    3. 呼叫 Azure OpenAI 將內容轉換為統一 Schema
    4. 使用 Schema 填寫 CNS Word 模板
    5. 回傳填好的 Word 檔案

    Args:
        file: 上傳的 PDF 檔案
        use_mock: 是否使用模擬資料（用於測試）

    Returns:
        FileResponse: 填好的 Word 檔案
    """
    logger.info("=" * 50)
    logger.info("收到報告轉換請求")
    logger.info(f"檔案名稱: {file.filename}")
    logger.info(f"使用模擬: {use_mock}")
    logger.info("=" * 50)

    # 驗證檔案類型
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(
            status_code=400,
            detail="請上傳 PDF 檔案"
        )

    # 讀取 PDF 內容
    try:
        pdf_bytes = await file.read()
        logger.info(f"PDF 大小: {len(pdf_bytes)} bytes")

        # 檢查檔案大小
        max_size = settings.max_pdf_size_mb * 1024 * 1024
        if len(pdf_bytes) > max_size:
            raise HTTPException(
                status_code=400,
                detail=f"檔案過大，最大允許 {settings.max_pdf_size_mb} MB"
            )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"讀取 PDF 失敗: {e}")
        raise HTTPException(status_code=400, detail=f"讀取 PDF 失敗: {str(e)}")

    # 判斷是否使用模擬資料
    use_mock_data = use_mock.lower() == "true"

    try:
        # Step 1: Adobe PDF Extract
        if use_mock_data:
            logger.info("使用模擬 Adobe Extract 結果")
            adobe_json = create_mock_extract_result()
        else:
            logger.info("呼叫 Adobe PDF Extract API...")
            try:
                adobe_json = await extract_pdf_to_json(pdf_bytes)
            except AdobeExtractError as e:
                logger.error(f"Adobe Extract 失敗: {e}")
                raise HTTPException(
                    status_code=500,
                    detail=f"PDF 解析失敗: {str(e)}"
                )

        # Step 2: Azure OpenAI Schema Extraction
        if use_mock_data:
            logger.info("使用模擬 Schema")
            schema = create_mock_schema()
        else:
            logger.info("呼叫 Azure OpenAI 萃取 Schema...")
            try:
                schema = await extract_report_schema_from_adobe_json(adobe_json)
            except Exception as e:
                logger.error(f"Schema 萃取失敗: {e}")
                raise HTTPException(
                    status_code=500,
                    detail=f"資料萃取失敗: {str(e)}"
                )

        # 設定來源檔名
        schema.source_filename = file.filename

        # Step 3: 尋找 Word 模板
        template_dir = os.path.join(os.path.dirname(__file__), "..", settings.template_dir)
        template_files = [
            f for f in os.listdir(template_dir)
            if f.endswith('.docx') and not f.startswith('~')
        ]

        if not template_files:
            raise HTTPException(
                status_code=500,
                detail="找不到 CNS 報告模板，請在 templates/ 資料夾放置 .docx 模板"
            )

        # 優先使用 placeholder 版本的模板
        placeholder_templates = [f for f in template_files if '.placeholder.' in f]
        if placeholder_templates:
            template_path = os.path.join(template_dir, placeholder_templates[0])
        else:
            template_path = os.path.join(template_dir, template_files[0])
        logger.info(f"使用模板: {template_path}")

        # Step 4: 填寫 Word 模板
        # 產生輸出檔案名稱
        report_no = schema.basic_info.cb_report_no or "Unknown"
        safe_report_no = "".join(c if c.isalnum() or c in "-_" else "_" for c in report_no)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"CNS_Report_{safe_report_no}_{timestamp}.docx"
        output_path = os.path.join(settings.temp_dir, output_filename)

        logger.info(f"填寫 Word 模板，輸出: {output_path}")

        try:
            fill_cns_template(schema, template_path, output_path)
        except FileNotFoundError as e:
            raise HTTPException(status_code=500, detail=str(e))
        except Exception as e:
            logger.error(f"填寫模板失敗: {e}")
            raise HTTPException(
                status_code=500,
                detail=f"填寫模板失敗: {str(e)}"
            )

        # Step 5: 回傳檔案
        logger.info("轉換完成，回傳檔案")

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{output_filename}"'
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"未預期的錯誤: {e}", exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"處理過程發生錯誤: {str(e)}"
        )


@app.get("/api/schema-sample")
async def get_schema_sample():
    """
    取得 Schema 範例（用於開發與測試）
    """
    schema = create_mock_schema()
    return JSONResponse(content=schema.model_dump())


@app.get("/api/template-info")
async def get_template_info():
    """
    取得模板資訊
    """
    template_dir = os.path.join(os.path.dirname(__file__), "..", settings.template_dir)

    if not os.path.exists(template_dir):
        return {
            "status": "error",
            "message": f"模板目錄不存在: {template_dir}"
        }

    template_files = [
        f for f in os.listdir(template_dir)
        if f.endswith('.docx') and not f.startswith('~')
    ]

    return {
        "status": "ok",
        "template_dir": template_dir,
        "templates": template_files,
        "count": len(template_files)
    }


# ==============================================
# Run with Uvicorn (for development)
# ==============================================

if __name__ == "__main__":
    import uvicorn

    # 取得 port（Zeabur 會設定 PORT 環境變數）
    port = int(os.environ.get("PORT", 8000))

    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port,
        reload=settings.debug
    )
