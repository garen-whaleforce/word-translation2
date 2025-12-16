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
import time
from datetime import datetime
from typing import Optional
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, UploadFile, HTTPException, Form, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import asyncio
import json
import base64

# 確保可以 import backend 內的模組
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import settings
from utils.logger import get_logger, setup_logging
from services.adobe_extract import extract_pdf_to_json as adobe_extract_pdf, AdobeExtractError
from services.pymupdf_extract import extract_pdf_to_json as pymupdf_extract_pdf, PyMuPDFExtractError
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
    expose_headers=["Content-Disposition", "X-Processing-Time", "X-PDF-Pages", "X-Total-Tokens", "X-Estimated-Cost"],
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
        input[type="text"] {
            width: 100%;
            padding: 12px;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 14px;
        }
        input[type="text"]:focus {
            outline: none;
            border-color: #007bff;
            box-shadow: 0 0 0 2px rgba(0,123,255,0.1);
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

            <hr style="margin: 20px 0; border: none; border-top: 1px solid #e0e0e0;">
            <p style="font-size: 13px; color: #666; margin-bottom: 15px;">📋 以下為台灣申請者資訊（選填，不填則空白）</p>

            <div class="form-group">
                <label for="applicantName">申請者名稱（選填）</label>
                <input type="text" id="applicantName" name="applicant_name" placeholder="台灣申請者/代理商名稱，如：鼎福科技有限公司">
            </div>

            <div class="form-group">
                <label for="applicantAddress">申請者地址（選填）</label>
                <input type="text" id="applicantAddress" name="applicant_address" placeholder="台灣地址，如：新北市中和區民治街19巷8號">
            </div>

            <div class="form-group">
                <label for="cnsReportNo">CNS 報告編號（選填）</label>
                <input type="text" id="cnsReportNo" name="cns_report_no" placeholder="如：AST-B-25120522-000">
            </div>

            <hr style="margin: 20px 0; border: none; border-top: 1px solid #e0e0e0;">

            <div class="form-group">
                <label for="reportAuthor">報告撰寫人（選填）</label>
                <input type="text" id="reportAuthor" name="report_author" placeholder="請輸入報告撰寫人姓名">
            </div>

            <div class="form-group">
                <label for="reportSigner">報告簽署人（選填）</label>
                <input type="text" id="reportSigner" name="report_signer" placeholder="請輸入報告簽署人姓名">
            </div>

            <div class="form-group">
                <label for="seriesModel">系列型號（選填）</label>
                <input type="text" id="seriesModel" name="series_model" placeholder="多個型號請用逗號分隔，如：MC-601, MC-602">
            </div>

            <button type="submit" id="submitBtn">開始轉換</button>
        </form>

        <div id="status" class="status"></div>
    </div>

    <div class="note">
        <strong>注意事項：</strong>
        <ul style="margin: 10px 0 0 20px; padding: 0;">
            <li>請確保 PDF 檔案為有效的 CB Test Report</li>
            <li>轉換時間依 PDF 頁數而定（約 1-5 分鐘）</li>
            <li>請確保 templates/ 資料夾中有 CNS Word 模板</li>
        </ul>
    </div>

    <script>
        const form = document.getElementById('uploadForm');
        const statusDiv = document.getElementById('status');
        const submitBtn = document.getElementById('submitBtn');
        let startTime = null;
        let timerInterval = null;

        // 更新計時器顯示
        function updateTimer() {
            if (!startTime) return;
            const elapsed = Math.floor((Date.now() - startTime) / 1000);
            const minutes = Math.floor(elapsed / 60);
            const seconds = elapsed % 60;
            const timerSpan = document.getElementById('timer');
            if (timerSpan) {
                timerSpan.textContent = `已執行 ${minutes}:${seconds.toString().padStart(2, '0')}`;
            }
        }

        // 更新進度訊息
        function updateProgress(message, detail = '') {
            const progressMsg = document.getElementById('progressMsg');
            const progressDetail = document.getElementById('progressDetail');
            if (progressMsg) progressMsg.textContent = message;
            if (progressDetail) progressDetail.textContent = detail;
        }

        form.addEventListener('submit', async (e) => {
            e.preventDefault();

            const fileInput = document.getElementById('pdfFile');

            if (!fileInput.files.length) {
                alert('請選擇 PDF 檔案');
                return;
            }

            // 顯示 loading 並開始計時
            statusDiv.className = 'status loading';
            statusDiv.innerHTML = `
                <div style="display: flex; align-items: center; margin-bottom: 10px;">
                    <span class="spinner"></span>
                    <span id="progressMsg">正在準備上傳...</span>
                </div>
                <div id="progressDetail" style="font-size: 13px; color: #666; margin-bottom: 5px;"></div>
                <div id="timer" style="font-size: 12px; color: #999;">已執行 0:00</div>
            `;
            submitBtn.disabled = true;

            // 開始計時
            startTime = Date.now();
            timerInterval = setInterval(updateTimer, 1000);

            try {
                const formData = new FormData();
                formData.append('file', fileInput.files[0]);

                // 台灣申請者資訊
                const applicantName = document.getElementById('applicantName').value.trim();
                const applicantAddress = document.getElementById('applicantAddress').value.trim();
                const cnsReportNo = document.getElementById('cnsReportNo').value.trim();

                if (applicantName) formData.append('applicant_name', applicantName);
                if (applicantAddress) formData.append('applicant_address', applicantAddress);
                if (cnsReportNo) formData.append('cns_report_no', cnsReportNo);

                // 其他選填欄位
                const reportAuthor = document.getElementById('reportAuthor').value.trim();
                const reportSigner = document.getElementById('reportSigner').value.trim();
                const seriesModel = document.getElementById('seriesModel').value.trim();

                if (reportAuthor) formData.append('report_author', reportAuthor);
                if (reportSigner) formData.append('report_signer', reportSigner);
                if (seriesModel) formData.append('series_model', seriesModel);

                // 更新進度
                updateProgress('正在上傳 PDF 檔案...', `檔案大小：${(fileInput.files[0].size / 1024 / 1024).toFixed(2)} MB`);

                // 使用 SSE 串流接收進度和結果
                const response = await fetch('/generate-report', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const errorText = await response.text();
                    try {
                        const errorData = JSON.parse(errorText);
                        throw new Error(errorData.detail || '轉換失敗');
                    } catch {
                        throw new Error(errorText || '轉換失敗');
                    }
                }

                // 讀取 SSE 串流
                const reader = response.body.getReader();
                const decoder = new TextDecoder();
                let buffer = '';
                let stats = {};
                let filename = 'CNS_Report.docx';
                let fileBase64 = null;

                while (true) {
                    const { done, value } = await reader.read();
                    if (done) break;

                    buffer += decoder.decode(value, { stream: true });

                    // 解析 SSE 事件
                    const lines = buffer.split('\\n');
                    buffer = lines.pop() || '';  // 保留未完成的行

                    let eventType = null;
                    let eventData = null;

                    for (const line of lines) {
                        if (line.startsWith('event: ')) {
                            eventType = line.slice(7).trim();
                        } else if (line.startsWith('data: ')) {
                            try {
                                eventData = JSON.parse(line.slice(6));
                            } catch (e) {
                                console.error('Failed to parse SSE data:', line);
                                continue;
                            }

                            // 處理事件
                            if (eventType === 'progress' && eventData) {
                                updateProgress(eventData.message, `進度：${eventData.percent}%`);
                            } else if (eventType === 'error' && eventData) {
                                throw new Error(eventData.message);
                            } else if (eventType === 'complete' && eventData) {
                                filename = eventData.filename;
                                fileBase64 = eventData.file_base64;
                                stats = eventData.stats || {};
                            }

                            eventType = null;
                            eventData = null;
                        }
                    }
                }

                // 檢查是否收到檔案
                if (!fileBase64) {
                    throw new Error('未收到檔案資料');
                }

                // 將 Base64 轉換為 Blob 並下載
                const binaryString = atob(fileBase64);
                const bytes = new Uint8Array(binaryString.length);
                for (let i = 0; i < binaryString.length; i++) {
                    bytes[i] = binaryString.charCodeAt(i);
                }
                const blob = new Blob([bytes], {
                    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                });

                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);

                // 停止計時
                clearInterval(timerInterval);

                statusDiv.className = 'status success';
                statusDiv.innerHTML = `
                    <div style="margin-bottom: 10px;">✓ 轉換成功！檔案已開始下載。</div>
                    <div style="font-size: 13px; color: #2e7d32; border-top: 1px solid #c8e6c9; padding-top: 10px; margin-top: 10px;">
                        <div><strong>執行統計：</strong></div>
                        <div>• 處理時間：${stats.processing_time || 'N/A'} 秒</div>
                        <div>• PDF 頁數：${stats.pdf_pages || 'N/A'} 頁</div>
                        <div>• Token 使用量：${stats.total_tokens ? stats.total_tokens.toLocaleString() : 'N/A'}</div>
                        <div>• 預估成本：${stats.estimated_cost ? '$' + stats.estimated_cost.toFixed(4) : 'N/A'}</div>
                    </div>
                `;

            } catch (error) {
                clearInterval(timerInterval);
                statusDiv.className = 'status error';
                statusDiv.textContent = '✗ 錯誤：' + error.message;
            } finally {
                submitBtn.disabled = false;
                startTime = null;
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
        "pdf_extractor": settings.pdf_extractor,
        "timestamp": datetime.now().isoformat()
    }


@app.post("/generate-report")
async def generate_report(
    file: UploadFile = File(..., description="CB Report PDF 檔案"),
    applicant_name: str = Form(default="", description="台灣申請者名稱"),
    applicant_address: str = Form(default="", description="台灣申請者地址"),
    cns_report_no: str = Form(default="", description="CNS 報告編號"),
    report_author: str = Form(default="", description="報告撰寫人"),
    report_signer: str = Form(default="", description="報告簽署人"),
    series_model: str = Form(default="", description="系列型號（逗號分隔）")
):
    """
    主要 API：將 CB PDF 轉換為 CNS Word 報告
    使用 Server-Sent Events (SSE) 串流回傳進度，避免長時間請求超時

    流程：
    1. 讀取上傳的 PDF 檔案
    2. 呼叫 Adobe PDF Extract API 萃取內容
    3. 呼叫 Azure OpenAI 將內容轉換為統一 Schema
    4. 使用 Schema 填寫 CNS Word 模板
    5. 回傳填好的 Word 檔案（Base64 編碼）

    Returns:
        StreamingResponse: SSE 串流，最後包含 Base64 編碼的 Word 檔案
    """
    start_time = time.time()
    pdf_filename = file.filename

    logger.info("=" * 50)
    logger.info("收到報告轉換請求")
    logger.info(f"檔案名稱: {pdf_filename}")
    logger.info(f"台灣申請者: {applicant_name or '(未填，使用 CB 報告資訊)'}")
    logger.info(f"申請者地址: {applicant_address or '(未填)'}")
    logger.info(f"CNS 報告編號: {cns_report_no or '(未填)'}")
    logger.info(f"報告撰寫人: {report_author or '(未填)'}")
    logger.info(f"報告簽署人: {report_signer or '(未填)'}")
    logger.info(f"系列型號: {series_model or '(未填)'}")
    logger.info("=" * 50)

    # 驗證檔案類型
    if not pdf_filename.lower().endswith('.pdf'):
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

    # 使用 SSE 串流回傳進度
    async def generate_stream():
        """SSE 串流生成器"""
        nonlocal pdf_bytes, pdf_filename, applicant_name, applicant_address
        nonlocal cns_report_no, report_author, report_signer, series_model, start_time

        def send_event(event_type: str, data: dict):
            """發送 SSE 事件"""
            return f"event: {event_type}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"

        try:
            yield send_event("progress", {"stage": "pdf_extract", "message": "正在解析 PDF 內容...", "percent": 10})

            # Step 1: PDF Extract
            extractor = settings.pdf_extractor.lower()
            logger.info(f"使用 PDF 擷取引擎: {extractor}")

            if extractor == "pymupdf":
                logger.info("呼叫 PyMuPDF 擷取 PDF...")
                try:
                    extract_json = await pymupdf_extract_pdf(pdf_bytes)
                except PyMuPDFExtractError as e:
                    logger.error(f"PyMuPDF Extract 失敗: {e}")
                    yield send_event("error", {"message": f"PDF 解析失敗: {str(e)}"})
                    return
            else:
                logger.info("呼叫 Adobe PDF Extract API...")
                try:
                    extract_json = await adobe_extract_pdf(pdf_bytes)
                except AdobeExtractError as e:
                    logger.error(f"Adobe Extract 失敗: {e}")
                    yield send_event("error", {"message": f"PDF 解析失敗: {str(e)}"})
                    return

            pdf_pages = extract_json.get("metadata", {}).get("total_pages", 0)
            yield send_event("progress", {"stage": "llm_start", "message": f"PDF 解析完成（{pdf_pages} 頁），正在進行 AI 翻譯...", "percent": 25})

            # Step 2: Azure OpenAI Schema Extraction（這是最耗時的步驟）
            # 每 10 秒發送一次心跳，保持連線
            llm_stats = None
            llm_task = asyncio.create_task(extract_report_schema_from_adobe_json(extract_json))

            heartbeat_count = 0
            while not llm_task.done():
                await asyncio.sleep(10)
                heartbeat_count += 1
                progress_percent = min(25 + heartbeat_count * 5, 85)
                yield send_event("progress", {
                    "stage": "llm_processing",
                    "message": f"AI 翻譯處理中...（已執行 {heartbeat_count * 10} 秒）",
                    "percent": progress_percent
                })

            try:
                schema, llm_stats = await llm_task
            except Exception as e:
                logger.error(f"Schema 萃取失敗: {e}")
                yield send_event("error", {"message": f"資料萃取失敗: {str(e)}"})
                return

            yield send_event("progress", {"stage": "template", "message": "AI 翻譯完成，正在產生 Word 文件...", "percent": 90})

            # 設定來源檔名
            schema.source_filename = pdf_filename

            # Step 3: 尋找 Word 模板
            template_dir = os.path.join(os.path.dirname(__file__), "..", settings.template_dir)
            template_files = [
                f for f in os.listdir(template_dir)
                if f.endswith('.docx') and not f.startswith('~')
            ]

            if not template_files:
                yield send_event("error", {"message": "找不到 CNS 報告模板"})
                return

            placeholder_templates = [f for f in template_files if '.placeholder.' in f]
            if placeholder_templates:
                template_path = os.path.join(template_dir, placeholder_templates[0])
            else:
                template_path = os.path.join(template_dir, template_files[0])

            # Step 4: 填寫 Word 模板
            pdf_basename = os.path.splitext(pdf_filename)[0]
            safe_basename = "".join(c if c.isalnum() or c in "-_" else "_" for c in pdf_basename)
            output_filename = f"AST-B-{safe_basename}.docx"
            output_path = os.path.join(settings.temp_dir, output_filename)

            user_inputs = {
                "applicant_name": applicant_name.strip() if applicant_name else "",
                "applicant_address": applicant_address.strip() if applicant_address else "",
                "cns_report_no": cns_report_no.strip() if cns_report_no else "",
                "report_author": report_author.strip() if report_author else "",
                "report_signer": report_signer.strip() if report_signer else "",
                "series_model": series_model.strip() if series_model else ""
            }

            try:
                fill_cns_template(schema, template_path, output_path, user_inputs=user_inputs)
            except Exception as e:
                logger.error(f"填寫模板失敗: {e}")
                yield send_event("error", {"message": f"填寫模板失敗: {str(e)}"})
                return

            # Step 5: 讀取檔案並以 Base64 編碼回傳
            with open(output_path, "rb") as f:
                file_content = f.read()
            file_base64 = base64.b64encode(file_content).decode("utf-8")

            processing_time = round(time.time() - start_time, 2)
            logger.info(f"轉換完成，總處理時間: {processing_time} 秒")

            # 發送完成事件，包含檔案資料
            yield send_event("complete", {
                "filename": output_filename,
                "file_base64": file_base64,
                "stats": {
                    "processing_time": processing_time,
                    "pdf_pages": pdf_pages,
                    "total_tokens": llm_stats.get("total_tokens", 0) if llm_stats else 0,
                    "estimated_cost": llm_stats.get("estimated_cost", 0) if llm_stats else 0
                }
            })

        except Exception as e:
            logger.error(f"串流處理錯誤: {e}", exc_info=True)
            yield send_event("error", {"message": f"處理過程發生錯誤: {str(e)}"})

    return StreamingResponse(
        generate_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"  # 禁用 Nginx 緩衝
        }
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
