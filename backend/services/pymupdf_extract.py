"""
==============================================
PyMuPDF PDF Extract Service
使用 PyMuPDF (fitz) 擷取 PDF 內容
作為 Adobe PDF Extract API 的免費替代方案
==============================================

此模組負責：
1. 讀取 PDF 檔案
2. 擷取文字內容（按頁分組）
3. 擷取表格內容
4. 回傳與 Adobe Extract 相容的 JSON 結構
"""

import fitz  # PyMuPDF
import subprocess
import tempfile
from typing import Optional
from pathlib import Path

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.logger import get_logger

logger = get_logger(__name__)


class PyMuPDFExtractError(Exception):
    """PyMuPDF 擷取錯誤"""
    pass


def _try_unlock_pdf(pdf_bytes: bytes) -> bytes:
    """
    嘗試使用 qpdf 移除 PDF 權限限制

    Args:
        pdf_bytes: 原始 PDF bytes

    Returns:
        解鎖後的 PDF bytes（如果成功），否則回傳原始 bytes
    """
    try:
        # 檢查 qpdf 是否可用
        result = subprocess.run(['which', 'qpdf'], capture_output=True)
        if result.returncode != 0:
            logger.debug("qpdf 未安裝，跳過解鎖步驟")
            return pdf_bytes

        # 建立暫存檔案
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as input_file:
            input_file.write(pdf_bytes)
            input_path = input_file.name

        output_path = input_path.replace('.pdf', '_unlocked.pdf')

        try:
            # 執行 qpdf 解鎖
            result = subprocess.run(
                ['qpdf', '--decrypt', input_path, output_path],
                capture_output=True,
                timeout=30
            )

            if result.returncode == 0 and os.path.exists(output_path):
                with open(output_path, 'rb') as f:
                    unlocked_bytes = f.read()
                logger.info("成功移除 PDF 權限限制")
                return unlocked_bytes
            else:
                logger.debug(f"qpdf 執行失敗: {result.stderr.decode()}")
                return pdf_bytes

        finally:
            # 清理暫存檔案
            if os.path.exists(input_path):
                os.unlink(input_path)
            if os.path.exists(output_path):
                os.unlink(output_path)

    except Exception as e:
        logger.debug(f"PDF 解鎖失敗: {e}")
        return pdf_bytes


def extract_pdf_with_pymupdf(pdf_bytes: bytes) -> dict:
    """
    使用 PyMuPDF 擷取 PDF 內容

    Args:
        pdf_bytes: PDF 檔案的 bytes

    Returns:
        與 Adobe Extract 相容的 JSON 結構:
        {
            "elements_by_page": {
                0: {"texts": [...], "tables": [...]},
                1: {"texts": [...], "tables": [...]},
                ...
            },
            "raw_text": "完整文字內容",
            "tables": [...],
            "metadata": {...}
        }
    """
    logger.info("開始使用 PyMuPDF 擷取 PDF...")

    # 嘗試解鎖 PDF
    pdf_bytes = _try_unlock_pdf(pdf_bytes)

    try:
        # 開啟 PDF
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise PyMuPDFExtractError(f"無法開啟 PDF: {e}")

    total_pages = len(doc)
    logger.info(f"PDF 共 {total_pages} 頁")

    elements_by_page = {}
    all_texts = []
    all_tables = []

    for page_num in range(total_pages):
        page = doc[page_num]

        # 擷取文字
        text_blocks = []

        # 使用 get_text("dict") 取得結構化文字
        text_dict = page.get_text("dict")

        for block in text_dict.get("blocks", []):
            if block.get("type") == 0:  # 文字區塊
                block_text = ""
                for line in block.get("lines", []):
                    line_text = ""
                    for span in line.get("spans", []):
                        line_text += span.get("text", "")
                    block_text += line_text + "\n"

                if block_text.strip():
                    text_blocks.append({
                        "Text": block_text.strip(),
                        "Bounds": [
                            block.get("bbox", [0, 0, 0, 0])[0],
                            block.get("bbox", [0, 0, 0, 0])[1],
                            block.get("bbox", [0, 0, 0, 0])[2],
                            block.get("bbox", [0, 0, 0, 0])[3]
                        ]
                    })
                    all_texts.append(f"[Page {page_num}] {block_text.strip()}")

        # 擷取表格
        page_tables = []
        try:
            tables = page.find_tables()
            for table in tables:
                # 將表格轉為文字
                table_data = table.extract()
                if table_data:
                    table_text = "\n".join([
                        "\t".join([str(cell) if cell else "" for cell in row])
                        for row in table_data
                    ])
                    page_tables.append({
                        "Text": table_text,
                        "Bounds": list(table.bbox) if hasattr(table, 'bbox') else [0, 0, 0, 0]
                    })
                    all_tables.append({
                        "page": page_num,
                        "content": table_text
                    })
        except Exception as e:
            logger.debug(f"頁面 {page_num} 表格擷取失敗: {e}")

        elements_by_page[page_num] = {
            "texts": text_blocks,
            "tables": page_tables
        }

    doc.close()

    # 組合原始文字
    raw_text = "\n\n".join(all_texts)

    result = {
        "elements_by_page": elements_by_page,
        "raw_text": raw_text,
        "tables": all_tables,
        "metadata": {
            "total_pages": total_pages,
            "extractor": "PyMuPDF",
            "version": fitz.version[0]
        }
    }

    logger.info(f"PyMuPDF 擷取完成: {total_pages} 頁, {len(all_texts)} 個文字區塊, {len(all_tables)} 個表格")

    return result


async def extract_pdf_to_json(pdf_bytes: bytes) -> dict:
    """
    非同步版本的 PDF 擷取（保持與 Adobe Extract 介面相容）

    Args:
        pdf_bytes: PDF 檔案的 bytes

    Returns:
        擷取結果的 dict
    """
    return extract_pdf_with_pymupdf(pdf_bytes)


def create_mock_extract_result() -> dict:
    """
    建立模擬的擷取結果（用於測試）
    """
    return {
        "elements_by_page": {
            0: {
                "texts": [
                    {"Text": "CB Test Report", "Bounds": [0, 0, 100, 20]},
                    {"Text": "Model: MC-601", "Bounds": [0, 30, 100, 50]},
                ],
                "tables": []
            }
        },
        "raw_text": "CB Test Report\nModel: MC-601\nTest Result: Pass",
        "tables": [],
        "metadata": {
            "total_pages": 1,
            "extractor": "Mock",
            "version": "1.0"
        }
    }


# 測試用
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python pymupdf_extract.py <pdf_file>")
        sys.exit(1)

    pdf_path = sys.argv[1]

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    result = extract_pdf_with_pymupdf(pdf_bytes)

    print(f"總頁數: {result['metadata']['total_pages']}")
    print(f"文字區塊數: {sum(len(p['texts']) for p in result['elements_by_page'].values())}")
    print(f"表格數: {len(result['tables'])}")
    print(f"\n前 500 字元:\n{result['raw_text'][:500]}")
