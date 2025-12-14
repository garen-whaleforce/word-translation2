# scripts/convert_pdf_to_text.py
"""
使用 Adobe PDF Extract API 將 CB PDF 轉成純文字
輸出到 artifacts/cb_mc_601_text.txt
"""
import asyncio
import sys
import os
from pathlib import Path

# 加入 backend 路徑以使用現有服務
sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from services.adobe_extract import extract_pdf_to_json


PDF_PATH = Path("templates/CB MC-601.pdf")
OUT_PATH = Path("artifacts/cb_mc_601_text.txt")


async def main():
    # 確保輸出目錄存在
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    # 確認 PDF 存在
    if not PDF_PATH.exists():
        print(f"錯誤：找不到 PDF 檔案 {PDF_PATH}")
        sys.exit(1)

    print(f"正在讀取 PDF：{PDF_PATH}")
    pdf_bytes = PDF_PATH.read_bytes()
    print(f"PDF 大小：{len(pdf_bytes):,} bytes")

    print("正在呼叫 Adobe PDF Extract API...")
    result = await extract_pdf_to_json(pdf_bytes)

    # 提取純文字
    raw_text = result.get("raw_text", "")

    if not raw_text:
        # 嘗試從 structured_data 提取
        structured = result.get("structured_data", {})
        elements = structured.get("elements", [])
        text_parts = []
        for elem in elements:
            t = elem.get("Text", "")
            if t:
                text_parts.append(t)
        raw_text = "\n".join(text_parts)

    if not raw_text:
        print("警告：無法從 PDF 提取文字")
        sys.exit(1)

    # 寫入輸出檔案
    OUT_PATH.write_text(raw_text, encoding="utf-8")
    print(f"成功！已寫入 {OUT_PATH}")
    print(f"文字長度：{len(raw_text):,} 字元")
    print(f"行數：{len(raw_text.splitlines()):,} 行")

    # 顯示前 500 字元預覽
    print("\n--- 文字預覽（前 500 字元）---")
    print(raw_text[:500])


if __name__ == "__main__":
    asyncio.run(main())
