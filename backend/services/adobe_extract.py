"""
==============================================
Adobe PDF Extract Service
使用 Adobe PDF Services API 萃取 PDF 內容
==============================================

此模組負責：
1. 取得 Adobe API access token
2. 上傳 PDF 並觸發 Extract 作業
3. 輪詢作業狀態並取得結果
4. 解析結果 JSON（文字 + 表格）
"""

import httpx
import json
import time
import zipfile
import io
from typing import Optional
from tenacity import retry, stop_after_attempt, wait_exponential

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import settings
from utils.logger import get_logger

logger = get_logger(__name__)


# ==============================================
# Adobe API Constants
# ==============================================

# Adobe IMS Token Endpoint
ADOBE_IMS_TOKEN_URL = "https://ims-na1.adobelogin.com/ims/token/v3"

# Adobe PDF Services API Endpoints
ADOBE_PDF_SERVICES_BASE = settings.adobe_pdf_services_base_url


class AdobeExtractError(Exception):
    """Adobe PDF Extract 相關錯誤"""
    pass


# ==============================================
# Token Management
# ==============================================

class AdobeTokenManager:
    """
    管理 Adobe API Access Token
    實作 token 快取與自動更新
    """

    def __init__(self):
        self._token: Optional[str] = None
        self._token_expires_at: float = 0

    async def get_access_token(self) -> str:
        """
        取得有效的 access token
        如果 token 即將過期或不存在，自動取得新 token
        """
        current_time = time.time()

        # 如果 token 還有效（預留 60 秒緩衝），直接回傳
        if self._token and current_time < (self._token_expires_at - 60):
            return self._token

        # 取得新 token
        await self._refresh_token()
        return self._token

    async def _refresh_token(self):
        """
        從 Adobe IMS 取得新的 access token
        """
        logger.info("正在取得 Adobe API access token...")

        async with httpx.AsyncClient() as client:
            try:
                response = await client.post(
                    ADOBE_IMS_TOKEN_URL,
                    data={
                        "grant_type": "client_credentials",
                        "client_id": settings.adobe_client_id,
                        "client_secret": settings.adobe_client_secret,
                        "scope": "openid,AdobeID,read_organizations"
                    },
                    headers={
                        "Content-Type": "application/x-www-form-urlencoded"
                    },
                    timeout=30.0
                )

                if response.status_code != 200:
                    error_msg = f"取得 Adobe token 失敗: {response.status_code} - {response.text}"
                    logger.error(error_msg)
                    raise AdobeExtractError(error_msg)

                token_data = response.json()
                self._token = token_data.get("access_token")
                # Adobe token 通常有效期為 24 小時（86400 秒）
                expires_in = token_data.get("expires_in", 86400)
                self._token_expires_at = time.time() + expires_in

                logger.info(f"成功取得 Adobe token，有效期 {expires_in} 秒")

            except httpx.RequestError as e:
                error_msg = f"連接 Adobe IMS 時發生錯誤: {str(e)}"
                logger.error(error_msg)
                raise AdobeExtractError(error_msg)


# 全域 token manager
_token_manager = AdobeTokenManager()


# ==============================================
# PDF Extract Functions
# ==============================================

@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=2, max=10)
)
async def _upload_pdf_and_create_job(pdf_bytes: bytes, access_token: str) -> str:
    """
    Step 1: 上傳 PDF 並建立 Extract 作業

    Args:
        pdf_bytes: PDF 檔案的 bytes
        access_token: Adobe API access token

    Returns:
        job_id: 作業 ID，用於後續查詢狀態
    """
    logger.info(f"正在上傳 PDF（{len(pdf_bytes)} bytes）並建立 Extract 作業...")

    async with httpx.AsyncClient() as client:
        # Step 1a: 取得 upload presigned URL
        # 參考: https://developer.adobe.com/document-services/docs/apis/#tag/PDF-Extract
        create_job_url = f"{ADOBE_PDF_SERVICES_BASE}/operation/extractpdf"

        headers = {
            "Authorization": f"Bearer {access_token}",
            "x-api-key": settings.adobe_client_id,
            "Content-Type": "application/json"
        }

        # 建立 Extract 作業的 payload
        # 包含要萃取的元素類型
        job_payload = {
            "elementsToExtract": ["text", "tables"],
            "tableOutputFormat": "csv",  # 或 "xlsx"
            "renditionsToExtract": [],  # 可選：萃取圖片
            "notifiers": []
        }

        # 首先需要上傳 PDF 到 Adobe 的 asset 服務
        # Step 1a: 建立 asset
        asset_url = f"{ADOBE_PDF_SERVICES_BASE}/assets"

        asset_response = await client.post(
            asset_url,
            headers={
                "Authorization": f"Bearer {access_token}",
                "x-api-key": settings.adobe_client_id,
                "Content-Type": "application/json"
            },
            json={
                "mediaType": "application/pdf"
            },
            timeout=60.0
        )

        if asset_response.status_code not in [200, 201]:
            raise AdobeExtractError(
                f"建立 asset 失敗: {asset_response.status_code} - {asset_response.text}"
            )

        asset_data = asset_response.json()
        upload_uri = asset_data.get("uploadUri")
        asset_id = asset_data.get("assetID")

        logger.info(f"已建立 asset，ID: {asset_id}")

        # Step 1b: 上傳 PDF 到 presigned URL
        upload_response = await client.put(
            upload_uri,
            content=pdf_bytes,
            headers={
                "Content-Type": "application/pdf"
            },
            timeout=120.0
        )

        if upload_response.status_code not in [200, 201]:
            raise AdobeExtractError(
                f"上傳 PDF 失敗: {upload_response.status_code} - {upload_response.text}"
            )

        logger.info("PDF 上傳成功")

        # Step 1c: 使用 asset ID 建立 Extract 作業
        extract_response = await client.post(
            create_job_url,
            headers=headers,
            json={
                "assetID": asset_id,
                "elementsToExtract": ["text", "tables"],
                "tableOutputFormat": "csv"
            },
            timeout=60.0
        )

        if extract_response.status_code not in [200, 201]:
            raise AdobeExtractError(
                f"建立 Extract 作業失敗: {extract_response.status_code} - {extract_response.text}"
            )

        # 從 response header 取得 job location
        job_location = extract_response.headers.get("x-request-id") or extract_response.headers.get("location")

        # 或從 response body 取得
        job_data = extract_response.json() if extract_response.text else {}
        job_id = job_data.get("jobId") or job_location

        logger.info(f"Extract 作業已建立，Job ID: {job_id}")

        return job_id


async def _poll_job_status(job_id: str, access_token: str, max_wait_seconds: int = 300) -> dict:
    """
    Step 2: 輪詢作業狀態直到完成

    Args:
        job_id: 作業 ID
        access_token: Adobe API access token
        max_wait_seconds: 最大等待秒數

    Returns:
        完成的作業資訊（包含下載 URL）
    """
    logger.info(f"正在等待 Extract 作業完成（Job ID: {job_id}）...")

    status_url = f"{ADOBE_PDF_SERVICES_BASE}/operation/extractpdf/{job_id}/status"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "x-api-key": settings.adobe_client_id
    }

    start_time = time.time()
    poll_interval = 3  # 每 3 秒檢查一次

    async with httpx.AsyncClient() as client:
        while True:
            elapsed = time.time() - start_time
            if elapsed > max_wait_seconds:
                raise AdobeExtractError(f"Extract 作業逾時（已等待 {max_wait_seconds} 秒）")

            try:
                response = await client.get(status_url, headers=headers, timeout=30.0)

                if response.status_code == 200:
                    status_data = response.json()
                    status = status_data.get("status", "").lower()

                    if status == "done" or status == "succeeded":
                        logger.info("Extract 作業完成")
                        return status_data

                    elif status in ["failed", "error"]:
                        error_msg = status_data.get("error", {}).get("message", "未知錯誤")
                        raise AdobeExtractError(f"Extract 作業失敗: {error_msg}")

                    elif status in ["in progress", "running", "pending"]:
                        logger.debug(f"作業進行中... ({elapsed:.0f}s)")

                    else:
                        logger.warning(f"未知狀態: {status}")

            except httpx.RequestError as e:
                logger.warning(f"輪詢時發生連接錯誤: {e}")

            await asyncio_sleep(poll_interval)


async def asyncio_sleep(seconds: float):
    """非同步 sleep"""
    import asyncio
    await asyncio.sleep(seconds)


async def _download_and_parse_result(result_data: dict, access_token: str) -> dict:
    """
    Step 3: 下載並解析 Extract 結果

    Adobe Extract 會回傳一個 ZIP 檔案，內含：
    - structuredData.json: 結構化的文字與表格資料
    - 可能的 CSV 檔案: 各個表格的原始資料

    Args:
        result_data: 完成的作業資訊
        access_token: Adobe API access token

    Returns:
        解析後的 JSON 字典
    """
    logger.info("正在下載並解析 Extract 結果...")

    # 取得結果下載 URL
    download_uri = result_data.get("content", {}).get("downloadUri")
    if not download_uri:
        # 嘗試其他可能的欄位名稱
        download_uri = result_data.get("asset", {}).get("downloadUri")

    if not download_uri:
        raise AdobeExtractError("無法取得結果下載 URL")

    async with httpx.AsyncClient() as client:
        # 注意：Adobe 回傳的 downloadUri 是 presigned URL（已含簽名）
        # 不需要加 Authorization header，否則會衝突
        response = await client.get(
            download_uri,
            timeout=120.0
        )

        if response.status_code != 200:
            raise AdobeExtractError(
                f"下載結果失敗: {response.status_code} - {response.text}"
            )

        content = response.content
        content_type = response.headers.get("content-type", "")

        # 檢查回傳格式：可能是 ZIP 或 JSON
        if content_type.startswith("application/json") or content.startswith(b'{'):
            # 直接是 JSON 格式
            logger.info("收到 JSON 格式結果")
            structured_data = json.loads(content)
            raw_text = _extract_text_from_structured_data(structured_data)
            return {
                "structured_data": structured_data,
                "tables": [],
                "raw_text": raw_text
            }
        else:
            # 解析 ZIP 檔案
            return _parse_extract_zip(content)


def _parse_extract_zip(zip_content: bytes) -> dict:
    """
    解析 Adobe Extract 回傳的 ZIP 檔案

    Args:
        zip_content: ZIP 檔案的 bytes

    Returns:
        解析後的結構化資料
    """
    result = {
        "structured_data": None,
        "tables": [],
        "raw_text": ""
    }

    try:
        with zipfile.ZipFile(io.BytesIO(zip_content), 'r') as zip_ref:
            file_list = zip_ref.namelist()
            logger.info(f"ZIP 內含檔案: {file_list}")

            # 讀取主要的結構化資料 JSON
            for filename in file_list:
                if filename.endswith("structuredData.json"):
                    with zip_ref.open(filename) as f:
                        result["structured_data"] = json.load(f)
                        logger.info("成功解析 structuredData.json")

                # 讀取表格 CSV
                elif filename.endswith(".csv"):
                    with zip_ref.open(filename) as f:
                        csv_content = f.read().decode('utf-8')
                        result["tables"].append({
                            "filename": filename,
                            "content": csv_content
                        })
                        logger.info(f"成功讀取表格: {filename}")

    except zipfile.BadZipFile as e:
        logger.error(f"ZIP 檔案格式錯誤: {e}")
        raise AdobeExtractError(f"無法解析 ZIP 檔案: {e}")

    # 從 structured_data 中提取純文字
    if result["structured_data"]:
        result["raw_text"] = _extract_text_from_structured_data(result["structured_data"])

    return result


def _extract_text_from_structured_data(structured_data: dict) -> str:
    """
    從 Adobe 的 structuredData.json 中提取純文字

    Adobe Extract 的 structured data 格式包含 elements 陣列，
    每個 element 有 Text 屬性包含該區塊的文字
    """
    text_parts = []

    elements = structured_data.get("elements", [])

    for element in elements:
        # 提取文字內容
        text = element.get("Text", "")
        if text:
            text_parts.append(text)

    return "\n".join(text_parts)


# ==============================================
# Main Export Function
# ==============================================

def _try_unlock_pdf(pdf_bytes: bytes) -> bytes:
    """
    嘗試使用 qpdf 移除 PDF 權限限制
    如果 qpdf 不可用或失敗，回傳原始 bytes
    """
    import subprocess
    import tempfile

    try:
        # 檢查 qpdf 是否可用
        result = subprocess.run(["which", "qpdf"], capture_output=True)
        if result.returncode != 0:
            logger.debug("qpdf 未安裝，跳過解鎖")
            return pdf_bytes

        # 建立暫存檔案
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as input_file:
            input_file.write(pdf_bytes)
            input_path = input_file.name

        output_path = input_path.replace(".pdf", "_unlocked.pdf")

        # 執行 qpdf 解鎖
        result = subprocess.run(
            ["qpdf", "--decrypt", input_path, output_path],
            capture_output=True,
            timeout=30
        )

        if result.returncode == 0:
            with open(output_path, "rb") as f:
                unlocked_bytes = f.read()
            logger.info("成功移除 PDF 權限限制")

            # 清理暫存檔
            import os
            os.unlink(input_path)
            os.unlink(output_path)

            return unlocked_bytes
        else:
            logger.debug(f"qpdf 解鎖失敗: {result.stderr.decode()}")
            import os
            os.unlink(input_path)
            return pdf_bytes

    except Exception as e:
        logger.debug(f"PDF 解鎖過程發生錯誤: {e}")
        return pdf_bytes


async def extract_pdf_to_json(pdf_bytes: bytes) -> dict:
    """
    主要函式：將 PDF 轉換為結構化 JSON

    這是此模組的主要入口點。

    Args:
        pdf_bytes: PDF 檔案的 bytes

    Returns:
        結構化的 JSON 字典，包含：
        - structured_data: Adobe Extract 的原始結構化資料
        - tables: 表格資料（CSV 格式）
        - raw_text: 提取的純文字
        - elements: 依頁分組的元素（方便後續處理）

    Raises:
        AdobeExtractError: 當 Extract 過程發生錯誤時

    Usage:
        >>> pdf_content = open("report.pdf", "rb").read()
        >>> result = await extract_pdf_to_json(pdf_content)
        >>> print(result["raw_text"][:500])
    """
    logger.info("開始 PDF Extract 流程...")

    # 嘗試解鎖 PDF（移除權限限制）
    pdf_bytes = _try_unlock_pdf(pdf_bytes)

    try:
        # 取得 access token
        access_token = await _token_manager.get_access_token()

        # 上傳 PDF 並建立作業
        job_id = await _upload_pdf_and_create_job(pdf_bytes, access_token)

        # 輪詢作業狀態
        result_data = await _poll_job_status(job_id, access_token)

        # 下載並解析結果
        extracted_data = await _download_and_parse_result(result_data, access_token)

        # 後處理：依頁分組元素
        if extracted_data.get("structured_data"):
            extracted_data["elements_by_page"] = _group_elements_by_page(
                extracted_data["structured_data"]
            )

        logger.info("PDF Extract 流程完成")
        return extracted_data

    except AdobeExtractError:
        raise
    except Exception as e:
        error_msg = f"PDF Extract 過程發生未預期錯誤: {str(e)}"
        logger.error(error_msg)
        raise AdobeExtractError(error_msg)


def _group_elements_by_page(structured_data: dict) -> dict:
    """
    將 elements 依頁碼分組

    Args:
        structured_data: Adobe Extract 的原始結構化資料

    Returns:
        以頁碼為 key 的字典
    """
    pages = {}

    elements = structured_data.get("elements", [])

    for element in elements:
        # Adobe Extract 的 element 通常有 Page 屬性
        page_num = element.get("Page", 0)

        if page_num not in pages:
            pages[page_num] = {
                "texts": [],
                "tables": []
            }

        # 判斷是文字還是表格
        if element.get("Table"):
            pages[page_num]["tables"].append(element)
        elif element.get("Text"):
            pages[page_num]["texts"].append(element)

    return pages


# ==============================================
# Utility Functions for Testing / Development
# ==============================================

def create_mock_extract_result() -> dict:
    """
    建立模擬的 Extract 結果（用於開發測試）

    當 Adobe API 尚未設定或想要測試下游流程時，
    可以使用此函式產生假資料。
    """
    return {
        "structured_data": {
            "elements": [
                {
                    "Page": 0,
                    "Text": "CB TEST REPORT",
                    "Bounds": [50, 700, 300, 730]
                },
                {
                    "Page": 0,
                    "Text": "Report Number: TW-12345-UL",
                    "Bounds": [50, 680, 300, 695]
                },
                {
                    "Page": 0,
                    "Text": "Standard: IEC 62368-1:2018",
                    "Bounds": [50, 660, 300, 675]
                },
                {
                    "Page": 1,
                    "Text": "Applicant: ABC Technology Co., Ltd.",
                    "Bounds": [50, 700, 400, 715]
                },
                {
                    "Page": 1,
                    "Text": "Address: No. 123, Tech Road, Taipei, Taiwan",
                    "Bounds": [50, 680, 450, 695]
                },
                {
                    "Page": 1,
                    "Text": "Manufacturer: XYZ Manufacturing Inc.",
                    "Bounds": [50, 650, 400, 665]
                },
                {
                    "Page": 2,
                    "Table": True,
                    "Text": "Model|Vout|Iout|Pout\nPA-120W-A|12V|10A|120W\nPA-120W-B|24V|5A|120W"
                }
            ]
        },
        "tables": [
            {
                "filename": "table_1.csv",
                "content": "Model,Vout,Iout,Pout\nPA-120W-A,12V,10A,120W\nPA-120W-B,24V,5A,120W"
            }
        ],
        "raw_text": """CB TEST REPORT
Report Number: TW-12345-UL
Standard: IEC 62368-1:2018

Applicant: ABC Technology Co., Ltd.
Address: No. 123, Tech Road, Taipei, Taiwan
Manufacturer: XYZ Manufacturing Inc.
""",
        "elements_by_page": {
            0: {
                "texts": [
                    {"Text": "CB TEST REPORT"},
                    {"Text": "Report Number: TW-12345-UL"},
                    {"Text": "Standard: IEC 62368-1:2018"}
                ],
                "tables": []
            },
            1: {
                "texts": [
                    {"Text": "Applicant: ABC Technology Co., Ltd."},
                    {"Text": "Address: No. 123, Tech Road, Taipei, Taiwan"},
                    {"Text": "Manufacturer: XYZ Manufacturing Inc."}
                ],
                "tables": []
            }
        }
    }
