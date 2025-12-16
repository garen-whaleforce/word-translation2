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

from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

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
            <p style="font-size: 13px; color: #666; margin-bottom: 15px;">📋 以下為台灣申請者資訊（選填，若不填則使用 CB 報告中的製造商資訊）</p>

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

                // 模擬進度更新（因為後端是同步處理，無法取得即時進度）
                const progressStages = [
                    { delay: 2000, msg: '正在解析 PDF 內容...', detail: '使用 PyMuPDF 擷取文字與表格' },
                    { delay: 5000, msg: '正在進行 AI 翻譯...', detail: '使用 Azure OpenAI 分析報告內容' },
                    { delay: 15000, msg: 'AI 翻譯處理中...', detail: '這可能需要 1-3 分鐘，請耐心等待' },
                    { delay: 30000, msg: '仍在處理中...', detail: '大型報告可能需要較長時間' },
                    { delay: 60000, msg: '即將完成...', detail: '正在合併結果並產生 Word 文件' }
                ];

                const progressTimers = progressStages.map(stage =>
                    setTimeout(() => updateProgress(stage.msg, stage.detail), stage.delay)
                );

                const response = await fetch('/generate-report', {
                    method: 'POST',
                    body: formData
                });

                // 清除進度計時器
                progressTimers.forEach(timer => clearTimeout(timer));

                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(errorData.detail || '轉換失敗');
                }

                // 取得統計資訊 (debug: log all headers)
                console.log('Response headers:');
                response.headers.forEach((value, key) => console.log(`  ${key}: ${value}`));

                const stats = {
                    totalTime: response.headers.get('X-Processing-Time') || 'N/A',
                    pdfPages: response.headers.get('X-PDF-Pages') || 'N/A',
                    totalTokens: response.headers.get('X-Total-Tokens') || 'N/A',
                    estimatedCost: response.headers.get('X-Estimated-Cost') || 'N/A'
                };
                console.log('Stats:', stats);

                // 取得檔案名稱
                const contentDisposition = response.headers.get('Content-Disposition');
                console.log('Content-Disposition header:', contentDisposition);
                let filename = 'CNS_Report.docx';
                if (contentDisposition) {
                    // 嘗試多種格式解析
                    // 格式1: filename="xxx.docx"
                    let match = contentDisposition.match(/filename="([^"]+)"/);
                    if (!match) {
                        // 格式2: filename=xxx.docx
                        match = contentDisposition.match(/filename=([^;\\s]+)/);
                    }
                    if (match) {
                        filename = decodeURIComponent(match[1]);
                        console.log('Parsed filename:', filename);
                    } else {
                        console.log('Could not parse filename from Content-Disposition');
                    }
                } else {
                    console.log('Content-Disposition header is null - CORS expose_headers may not be working');
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

                // 停止計時
                clearInterval(timerInterval);
                const elapsed = Math.floor((Date.now() - startTime) / 1000);
                const minutes = Math.floor(elapsed / 60);
                const seconds = elapsed % 60;

                statusDiv.className = 'status success';
                statusDiv.innerHTML = `
                    <div style="margin-bottom: 10px;">✓ 轉換成功！檔案已開始下載。</div>
                    <div style="font-size: 13px; color: #2e7d32; border-top: 1px solid #c8e6c9; padding-top: 10px; margin-top: 10px;">
                        <div><strong>執行統計：</strong></div>
                        <div>• 處理時間：${stats.totalTime !== 'N/A' ? stats.totalTime + ' 秒' : minutes + ':' + seconds.toString().padStart(2, '0')}</div>
                        <div>• PDF 頁數：${stats.pdfPages} 頁</div>
                        <div>• Token 使用量：${stats.totalTokens !== 'N/A' ? parseInt(stats.totalTokens).toLocaleString() : 'N/A'}</div>
                        <div>• 預估成本：${stats.estimatedCost !== 'N/A' ? '$' + stats.estimatedCost : 'N/A'}</div>
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

    流程：
    1. 讀取上傳的 PDF 檔案
    2. 呼叫 Adobe PDF Extract API 萃取內容
    3. 呼叫 Azure OpenAI 將內容轉換為統一 Schema
    4. 使用 Schema 填寫 CNS Word 模板
    5. 回傳填好的 Word 檔案

    Args:
        file: 上傳的 PDF 檔案
        applicant_name: 台灣申請者名稱（覆蓋 CB 報告中的製造商）
        applicant_address: 台灣申請者地址
        cns_report_no: CNS 報告編號

    Returns:
        FileResponse: 填好的 Word 檔案
    """
    start_time = time.time()

    logger.info("=" * 50)
    logger.info("收到報告轉換請求")
    logger.info(f"檔案名稱: {file.filename}")
    logger.info(f"台灣申請者: {applicant_name or '(未填，使用 CB 報告資訊)'}")
    logger.info(f"申請者地址: {applicant_address or '(未填)'}")
    logger.info(f"CNS 報告編號: {cns_report_no or '(未填)'}")
    logger.info(f"報告撰寫人: {report_author or '(未填)'}")
    logger.info(f"報告簽署人: {report_signer or '(未填)'}")
    logger.info(f"系列型號: {series_model or '(未填)'}")
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

    try:
        # Step 1: PDF Extract（根據設定選擇 PyMuPDF 或 Adobe）
        extractor = settings.pdf_extractor.lower()
        logger.info(f"使用 PDF 擷取引擎: {extractor}")

        if extractor == "pymupdf":
            # 使用免費的 PyMuPDF
            logger.info("呼叫 PyMuPDF 擷取 PDF...")
            try:
                extract_json = await pymupdf_extract_pdf(pdf_bytes)
            except PyMuPDFExtractError as e:
                logger.error(f"PyMuPDF Extract 失敗: {e}")
                raise HTTPException(
                    status_code=500,
                    detail=f"PDF 解析失敗: {str(e)}"
                )
        else:
            # 使用 Adobe PDF Extract API
            logger.info("呼叫 Adobe PDF Extract API...")
            try:
                extract_json = await adobe_extract_pdf(pdf_bytes)
            except AdobeExtractError as e:
                logger.error(f"Adobe Extract 失敗: {e}")
                raise HTTPException(
                    status_code=500,
                    detail=f"PDF 解析失敗: {str(e)}"
                )

        # Step 2: Azure OpenAI Schema Extraction
        llm_stats = None
        logger.info("呼叫 Azure OpenAI 萃取 Schema...")
        try:
            schema, llm_stats = await extract_report_schema_from_adobe_json(extract_json)
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
        # 產生輸出檔案名稱 - 使用上傳的 PDF 檔名
        # 例如：上傳 DYS830.pdf -> AST-B-DYS830.docx
        pdf_basename = os.path.splitext(file.filename)[0]  # 移除 .pdf 副檔名
        # 清理檔名，只保留安全字元
        safe_basename = "".join(c if c.isalnum() or c in "-_" else "_" for c in pdf_basename)
        output_filename = f"AST-B-{safe_basename}.docx"
        output_path = os.path.join(settings.temp_dir, output_filename)

        logger.info(f"填寫 Word 模板，輸出: {output_path}")

        # 準備前端傳入的額外欄位
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
        except FileNotFoundError as e:
            raise HTTPException(status_code=500, detail=str(e))
        except Exception as e:
            logger.error(f"填寫模板失敗: {e}")
            raise HTTPException(
                status_code=500,
                detail=f"填寫模板失敗: {str(e)}"
            )

        # Step 5: 回傳檔案
        processing_time = round(time.time() - start_time, 2)
        logger.info(f"轉換完成，總處理時間: {processing_time} 秒")

        # 取得 PDF 頁數
        pdf_pages = extract_json.get("metadata", {}).get("total_pages", 0)

        # 準備回應 headers
        response_headers = {
            "Content-Disposition": f'attachment; filename="{output_filename}"',
            "X-Processing-Time": str(processing_time),
            "X-PDF-Pages": str(pdf_pages),
            "Access-Control-Expose-Headers": "Content-Disposition, X-Processing-Time, X-PDF-Pages, X-Total-Tokens, X-Estimated-Cost"
        }

        # 如果有 LLM 統計資訊，加入 headers
        if llm_stats:
            response_headers["X-Total-Tokens"] = str(llm_stats.get("total_tokens", 0))
            response_headers["X-Estimated-Cost"] = str(llm_stats.get("estimated_cost", 0))
            logger.info(f"  - Token 使用量: {llm_stats.get('total_tokens', 0):,}")
            logger.info(f"  - 預估成本: ${llm_stats.get('estimated_cost', 0):.4f}")

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=response_headers
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
