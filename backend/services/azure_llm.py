"""
==============================================
Azure OpenAI LLM Service
使用 Azure OpenAI 將 Adobe Extract 結果轉換為統一 Schema
==============================================

此模組負責：
1. 準備 system prompt 與 user prompt
2. 將 Adobe JSON 分 chunk 送給 LLM
3. 解析 LLM 回應並轉換為 ReportSchema
4. 合併多次 LLM 呼叫的結果
5. 執行必要的繁中翻譯
"""

import json
import re
import asyncio
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from typing import Optional, List, Dict, Any, Tuple
from openai import AzureOpenAI
from tenacity import retry, stop_after_attempt, wait_exponential

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import settings
from schemas.report_schema import (
    ReportSchema,
    BasicInfo,
    TestItemParticulars,
    SeriesModel,
    ClauseVerdict,
    KeyTables,
    Translations,
    CheckboxFlags,
    merge_schemas,
    create_empty_schema
)
from utils.logger import get_logger

logger = get_logger(__name__)

# ==============================================
# Token Usage Tracking
# ==============================================

class TokenUsageTracker:
    """追蹤 token 使用量"""
    def __init__(self):
        self.total_prompt_tokens = 0
        self.total_completion_tokens = 0
        self.call_count = 0

    def add(self, prompt_tokens: int, completion_tokens: int):
        self.total_prompt_tokens += prompt_tokens
        self.total_completion_tokens += completion_tokens
        self.call_count += 1

    @property
    def total_tokens(self) -> int:
        return self.total_prompt_tokens + self.total_completion_tokens

    def calculate_cost(self) -> float:
        """
        計算成本（基於 Azure OpenAI gpt-4o 定價）
        價格來源：https://azure.microsoft.com/en-us/pricing/details/cognitive-services/openai-service/
        gpt-4o: $5.00 / 1M input tokens, $15.00 / 1M output tokens
        """
        input_cost = (self.total_prompt_tokens / 1_000_000) * 5.00
        output_cost = (self.total_completion_tokens / 1_000_000) * 15.00
        return round(input_cost + output_cost, 4)

# 全域 token tracker（每次請求會重置）
_token_tracker = TokenUsageTracker()


# ==============================================
# Azure OpenAI Client Setup
# ==============================================

def get_azure_client() -> AzureOpenAI:
    """
    建立 Azure OpenAI 客戶端

    Returns:
        AzureOpenAI client instance
    """
    return AzureOpenAI(
        api_key=settings.azure_openai_api_key,
        api_version=settings.azure_openai_api_version,
        azure_endpoint=settings.azure_openai_endpoint
    )


# ==============================================
# Prompt Templates
# ==============================================

SYSTEM_PROMPT = """你是一位專業的電子產品安全測試工程師，專門負責分析 IEC 62368-1 CB Test Report。

你的任務是從提供的 PDF 萃取內容中，精確地提取以下資訊並以 JSON 格式回傳：

## 需要提取的資訊類別：

### 1. basic_info（基本資料）
- cb_report_no: CB 報告編號
- standard: 適用標準（例如 IEC 62368-1:2018）
- applicant_en: 申請人名稱（英文）
- applicant_address_en: 申請人地址（英文）
- manufacturer_en: 製造商名稱（英文）
- manufacturer_address_en: 製造商地址（英文）
- product_name_en: 產品名稱（英文）
- model_main: 主型號
- ratings_input: 輸入額定值（如 100-240Vac, 50/60Hz, 2A）
- ratings_output: 輸出額定值（如 12Vdc, 5A）
- issue_date: 報告發行日期（YYYY-MM-DD）
- receive_date: 試驗件收件日（YYYY-MM-DD）
- test_date_from: 測試開始日期（YYYY-MM-DD）
- test_date_to: 測試結束日期（YYYY-MM-DD）
- equipment_mass: 設備質量（如 "0.5 kg"）
- protection_rating: 保護裝置額定電流（如 "10A"）
- brand: 品牌名稱
- trademark: 商標

### 2. test_item_particulars（試驗樣品特性）
- product_group: 產品群組（AV / ICT / Telecom）
- classification_of_use: 使用分類（Ordinary / Skilled / Instructed）- 陣列
- supply_connection: 電源連接（Class I / II / III）- 陣列
- ovc: 過電壓類別（如 OVC II）
- pollution_degree: 污染等級（如 2）
- ip_code: IP 防護等級（如 IP20）
- tma: 最高環境溫度（如 40°C）
- altitude_limit_m: 海拔高度限制（數字，單位 m，如 2000）
- mobility: 移動性（Portable / Stationary / Fixed）
- mains_supply: 主電源類型（AC / DC / AC+DC）
- rated_voltage: 額定電壓
- rated_frequency: 額定頻率
- rated_current: 額定電流

### 3. series_models（系列型號）- 陣列
每個型號包含：
- model: 型號名稱
- vout: 輸出電壓
- iout: 輸出電流
- pout: 輸出功率
- vin: 輸入電壓
- case_type: 外殼類型（Metal / Plastic / Open Frame）
- differences: 與主型號的差異說明

### 4. clause_verdicts（條文判定）- 陣列
每個條文包含：
- clause: 條文編號（例如 "4.1.1"）
- verdict: 判定結果（P / N/A / F）
- comment_en: 英文備註

### 5. checkbox_flags（勾選狀態）
根據 test_item_particulars 的內容設定對應的布林值：
- is_av / is_ict / is_telecom
- is_ordinary / is_skilled / is_instructed
- is_class_i / is_class_ii / is_class_iii
- is_portable / is_stationary / is_fixed

## 輸出格式要求：

1. 只輸出純 JSON，不要有任何其他文字或 markdown 標記
2. 如果某個欄位在文件中找不到，使用 null 或空字串
3. 陣列欄位如果沒有資料，使用空陣列 []
4. 保持英文資料的原始大小寫和格式
5. 日期格式使用 YYYY-MM-DD
6. 數值欄位（如 altitude_limit_m）請使用數字，不要包含單位

## 特別注意：

- CB 報告通常在前幾頁有 basic_info
- Test Item Particulars 表格通常在報告開頭，包含產品分類、環境條件等
- 系列型號表格可能跨多頁
- 條文判定（Clause verdicts）通常是報告的主體部分
- 溫升測試表格通常在 Clause 5 或 Annex 中
- 設備質量（Mass）通常在 Product Information 或 Test Item Particulars 中
- 移動性（Mobility）可能標示為 Portable, Stationary, 或 Fixed

請仔細閱讀並提取所有能找到的資訊。"""


TRANSLATION_PROMPT = """你是一位專業的翻譯人員，專門翻譯電子產品安全測試報告。

請將以下英文內容翻譯成繁體中文：

{content_to_translate}

## 翻譯要求：

1. 公司名稱：
   - 如果是知名公司，使用其官方中文名稱
   - 如果無法確定，保留英文名稱或音譯
   - 例如：Apple Inc. → 蘋果公司

2. 產品名稱：
   - 使用業界常用的繁體中文術語
   - 例如：Power Adapter → 電源供應器
   - 例如：Switching Power Supply → 交換式電源供應器

3. 地址：
   - 保留原有格式，只翻譯國家名稱
   - 例如：Taiwan → 台灣
   - 例如：China → 中國

4. 技術術語：
   - 使用 CNS 標準的官方繁體中文譯名
   - 例如：Touch Current → 觸及電流
   - 例如：Dielectric Strength → 耐電壓

請以 JSON 格式回傳翻譯結果：
{{
    "applicant_zh": "申請人中文名稱",
    "applicant_address_zh": "申請人中文地址",
    "manufacturer_zh": "製造商中文名稱",
    "manufacturer_address_zh": "製造商中文地址",
    "product_name_zh": "產品中文名稱"
}}

只輸出 JSON，不要有其他文字。"""


CLAUSE_TRANSLATION_PROMPT = """請將以下 CB 報告條文備註翻譯成簡潔的繁體中文。

原文：{comment_en}

要求：
1. 使用 CNS 標準術語
2. 翻譯要簡潔明瞭
3. 保留關鍵數值和單位
4. 只輸出翻譯後的中文，不要其他說明"""


# ==============================================
# LLM Interaction Functions
# ==============================================

@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=2, max=10)
)
def _call_llm(messages: List[Dict[str, str]], temperature: float = None) -> str:
    """
    呼叫 Azure OpenAI LLM

    Args:
        messages: 訊息列表
        temperature: 溫度參數（預設使用 settings）

    Returns:
        LLM 回應的文字內容
    """
    global _token_tracker
    client = get_azure_client()

    temp = temperature if temperature is not None else settings.llm_temperature

    logger.debug(f"呼叫 LLM，deployment: {settings.azure_openai_deployment}")

    response = client.chat.completions.create(
        model=settings.azure_openai_deployment,
        messages=messages,
        temperature=temp,
        max_completion_tokens=settings.llm_max_tokens,
        response_format={"type": "json_object"}  # 強制 JSON 輸出
    )

    # 追蹤 token 使用量
    if response.usage:
        _token_tracker.add(
            prompt_tokens=response.usage.prompt_tokens,
            completion_tokens=response.usage.completion_tokens
        )

    content = response.choices[0].message.content
    logger.debug(f"LLM 回應長度: {len(content)} 字元")

    return content


def _parse_llm_json_response(response_text: str, return_empty_on_fail: bool = False) -> dict:
    """
    解析 LLM 回傳的 JSON 字串

    Args:
        response_text: LLM 回傳的原始文字
        return_empty_on_fail: 如果解析失敗是否回傳空 dict（而非拋出錯誤）

    Returns:
        解析後的 dict
    """
    if not response_text:
        logger.warning("LLM 回傳空字串")
        return {}

    # 清理常見問題
    cleaned = response_text.strip()

    # 嘗試直接解析
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        logger.debug(f"直接解析失敗: {e}")

    # 嘗試提取 JSON 區塊（如果 LLM 加了 markdown 標記）
    json_match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', cleaned)
    if json_match:
        try:
            return json.loads(json_match.group(1))
        except json.JSONDecodeError as e:
            logger.debug(f"從 markdown 區塊解析失敗: {e}")

    # 嘗試找到第一個 { 和最後一個 }
    start = cleaned.find('{')
    end = cleaned.rfind('}')
    if start != -1 and end != -1 and end > start:
        json_str = cleaned[start:end + 1]
        try:
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            logger.debug(f"從括號範圍解析失敗: {e}")

            # 嘗試修復常見的 JSON 問題
            try:
                # 移除尾部多餘的逗號
                fixed = re.sub(r',\s*([}\]])', r'\1', json_str)
                # 修復單引號
                fixed = fixed.replace("'", '"')
                return json.loads(fixed)
            except json.JSONDecodeError:
                pass

            # 嘗試使用更寬鬆的解析
            try:
                import ast
                # 將 null 轉為 None, true/false 轉為 True/False
                fixed = json_str.replace('null', 'None').replace('true', 'True').replace('false', 'False')
                result = ast.literal_eval(fixed)
                if isinstance(result, dict):
                    return result
            except (ValueError, SyntaxError):
                pass

    # 解析失敗
    logger.error(f"無法解析 LLM 回應為 JSON: {response_text[:500]}")

    if return_empty_on_fail:
        logger.warning("回傳空 dict 以允許流程繼續")
        return {}

    raise ValueError(f"JSON 解析失敗: {response_text[:100]}")


# ==============================================
# Chunk Processing
# ==============================================

def _prepare_chunks(adobe_json: dict, pages_per_chunk: int = None) -> List[dict]:
    """
    將 Adobe Extract 結果分成多個 chunk

    Args:
        adobe_json: Adobe Extract 的結果
        pages_per_chunk: 每個 chunk 包含的頁數

    Returns:
        chunk 列表，每個 chunk 是一個 dict
    """
    pages_per_chunk = pages_per_chunk or settings.llm_chunk_pages

    elements_by_page = adobe_json.get("elements_by_page", {})

    if not elements_by_page:
        # 如果沒有按頁分組，直接回傳整個內容作為單一 chunk
        return [{
            "pages": [0],
            "content": adobe_json.get("raw_text", ""),
            "tables": adobe_json.get("tables", [])
        }]

    # 取得所有頁碼並排序
    page_numbers = sorted(elements_by_page.keys())
    total_pages = len(page_numbers)

    chunks = []

    for i in range(0, total_pages, pages_per_chunk):
        chunk_pages = page_numbers[i:i + pages_per_chunk]

        # 收集這幾頁的文字
        texts = []
        tables = []

        for page_num in chunk_pages:
            page_data = elements_by_page.get(page_num, {})

            # 收集文字
            for text_element in page_data.get("texts", []):
                text = text_element.get("Text", "")
                if text:
                    texts.append(f"[Page {page_num}] {text}")

            # 收集表格
            for table_element in page_data.get("tables", []):
                tables.append({
                    "page": page_num,
                    "content": table_element.get("Text", "")
                })

        chunks.append({
            "pages": chunk_pages,
            "content": "\n".join(texts),
            "tables": tables
        })

    logger.info(f"已將 {total_pages} 頁分成 {len(chunks)} 個 chunks")
    return chunks


def _process_chunk(chunk: dict, chunk_index: int, total_chunks: int) -> ReportSchema:
    """
    處理單一 chunk，呼叫 LLM 萃取資料

    Args:
        chunk: chunk 資料
        chunk_index: chunk 索引
        total_chunks: 總 chunk 數

    Returns:
        從此 chunk 萃取的 ReportSchema
    """
    logger.info(f"處理 chunk {chunk_index + 1}/{total_chunks}，頁面: {chunk['pages']}")

    # 準備 user prompt
    user_content = f"""以下是 CB Test Report 第 {chunk['pages']} 頁的內容：

=== 文字內容 ===
{chunk['content']}

=== 表格內容 ===
{json.dumps(chunk['tables'], ensure_ascii=False, indent=2)}

請提取所有能找到的資訊，以 JSON 格式回傳。"""

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content}
    ]

    # 呼叫 LLM
    response = _call_llm(messages)

    # 解析回應
    try:
        extracted_data = _parse_llm_json_response(response)
    except ValueError as e:
        logger.warning(f"Chunk {chunk_index + 1} JSON 解析失敗: {e}")
        return create_empty_schema()

    # 轉換為 ReportSchema
    return _dict_to_schema(extracted_data)


def _dict_to_schema(data: dict) -> ReportSchema:
    """
    將 LLM 回傳的 dict 轉換為 ReportSchema

    Args:
        data: LLM 回傳的 dict

    Returns:
        ReportSchema 物件
    """
    schema = create_empty_schema()

    # 處理 basic_info
    if "basic_info" in data:
        bi = data["basic_info"] or {}
        # 使用 or "" 確保 None 值轉為空字串
        schema.basic_info = BasicInfo(
            cb_report_no=bi.get("cb_report_no") or "",
            standard=bi.get("standard") or "",
            applicant_en=bi.get("applicant_en") or "",
            applicant_address_en=bi.get("applicant_address_en") or "",
            manufacturer_en=bi.get("manufacturer_en") or "",
            manufacturer_address_en=bi.get("manufacturer_address_en") or "",
            product_name_en=bi.get("product_name_en") or "",
            model_main=bi.get("model_main") or "",
            ratings_input=bi.get("ratings_input") or "",
            ratings_output=bi.get("ratings_output") or "",
            issue_date=bi.get("issue_date"),
            receive_date=bi.get("receive_date"),
            test_date_from=bi.get("test_date_from"),
            test_date_to=bi.get("test_date_to"),
            equipment_mass=bi.get("equipment_mass"),
            protection_rating=bi.get("protection_rating"),
            test_lab=bi.get("test_lab"),
            brand=bi.get("brand"),
            trademark=bi.get("trademark")
        )

    # 處理 test_item_particulars
    if "test_item_particulars" in data:
        tip = data["test_item_particulars"] or {}
        # 確保 tma 是字串（LLM 可能回傳數字）
        tma_value = tip.get("tma")
        if tma_value is not None and not isinstance(tma_value, str):
            tma_value = str(tma_value)
        schema.test_item_particulars = TestItemParticulars(
            product_group=tip.get("product_group"),
            classification_of_use=tip.get("classification_of_use", []),
            supply_connection=tip.get("supply_connection", []),
            ovc=tip.get("ovc"),
            pollution_degree=tip.get("pollution_degree"),
            ip_code=tip.get("ip_code"),
            tma=tma_value,
            altitude_limit_m=tip.get("altitude_limit_m"),
            mobility=tip.get("mobility"),
            mains_supply=tip.get("mains_supply"),
            rated_voltage=tip.get("rated_voltage"),
            rated_frequency=tip.get("rated_frequency"),
            rated_current=tip.get("rated_current")
        )

    # 處理 series_models
    if "series_models" in data:
        for model_data in data["series_models"]:
            if model_data.get("model"):
                schema.series_models.append(SeriesModel(
                    model=model_data.get("model", ""),
                    vout=model_data.get("vout"),
                    iout=model_data.get("iout"),
                    pout=model_data.get("pout"),
                    case_type=model_data.get("case_type"),
                    differences=model_data.get("differences")
                ))

    # 處理 clause_verdicts
    if "clause_verdicts" in data:
        for verdict_data in data["clause_verdicts"]:
            if verdict_data.get("clause"):
                schema.clause_verdicts.append(ClauseVerdict(
                    clause=verdict_data.get("clause", ""),
                    verdict=verdict_data.get("verdict", "P"),
                    comment_en=verdict_data.get("comment_en"),
                    clause_title=verdict_data.get("clause_title")
                ))

    # 處理 checkbox_flags
    if "checkbox_flags" in data:
        flags = data["checkbox_flags"]
        schema.checkbox_flags = CheckboxFlags(
            is_av=flags.get("is_av", False),
            is_ict=flags.get("is_ict", False),
            is_av_ict=flags.get("is_av_ict", False),
            is_telecom=flags.get("is_telecom", False),
            is_ordinary=flags.get("is_ordinary", False),
            is_skilled=flags.get("is_skilled", False),
            is_instructed=flags.get("is_instructed", False),
            is_class_i=flags.get("is_class_i", False),
            is_class_ii=flags.get("is_class_ii", False),
            is_class_iii=flags.get("is_class_iii", False)
        )

    return schema


# ==============================================
# Translation Functions
# ==============================================

def _translate_to_chinese(schema: ReportSchema) -> ReportSchema:
    """
    翻譯 schema 中的英文欄位為繁體中文

    Args:
        schema: 原始 schema

    Returns:
        包含翻譯的 schema
    """
    logger.info("開始進行繁體中文翻譯...")

    # 準備需要翻譯的內容
    try:
        content_to_translate = {
            "applicant_en": schema.basic_info.applicant_en or "",
            "applicant_address_en": schema.basic_info.applicant_address_en or "",
            "manufacturer_en": schema.basic_info.manufacturer_en or "",
            "manufacturer_address_en": schema.basic_info.manufacturer_address_en or "",
            "product_name_en": schema.basic_info.product_name_en or ""
        }
    except Exception as e:
        logger.error(f"準備翻譯內容時發生錯誤: {e}")
        schema.translations = Translations()
        return schema

    # 過濾掉空值
    content_to_translate = {k: v for k, v in content_to_translate.items() if v}

    if not content_to_translate:
        logger.warning("沒有需要翻譯的內容")
        schema.translations = Translations()
        return schema

    # 呼叫 LLM 進行翻譯
    try:
        translation_prompt = TRANSLATION_PROMPT.format(
            content_to_translate=json.dumps(content_to_translate, ensure_ascii=False, indent=2)
        )
    except Exception as e:
        logger.error(f"準備翻譯 prompt 時發生錯誤: {e}")
        schema.translations = Translations()
        return schema

    messages = [
        {"role": "system", "content": "你是專業的翻譯人員，專門翻譯電子產品安全測試報告。只輸出 JSON，不要有其他文字。"},
        {"role": "user", "content": translation_prompt}
    ]

    try:
        response = _call_llm(messages, temperature=0.3)

        if response:
            logger.debug(f"翻譯 LLM 回應長度: {len(response)}")
            # 使用 return_empty_on_fail=True 確保即使解析失敗也不會中斷
            translations = _parse_llm_json_response(response, return_empty_on_fail=True)
        else:
            translations = {}

        # 更新 schema 的翻譯欄位
        schema.translations = Translations(
            applicant_zh=translations.get("applicant_zh") if translations else None,
            applicant_address_zh=translations.get("applicant_address_zh") if translations else None,
            manufacturer_zh=translations.get("manufacturer_zh") if translations else None,
            manufacturer_address_zh=translations.get("manufacturer_address_zh") if translations else None,
            product_name_zh=translations.get("product_name_zh") if translations else None
        )

        logger.info("翻譯完成")

    except Exception as e:
        logger.error(f"翻譯過程發生錯誤: {e}", exc_info=True)
        # 設定空翻譯，不讓錯誤中斷流程
        schema.translations = Translations()

    # 翻譯條文備註（也加上錯誤處理）
    try:
        schema = _translate_clause_comments(schema)
    except Exception as e:
        logger.error(f"條文備註翻譯過程發生錯誤: {e}", exc_info=True)

    return schema


def _translate_clause_comments(schema: ReportSchema) -> ReportSchema:
    """
    翻譯條文備註

    Args:
        schema: 原始 schema

    Returns:
        包含翻譯備註的 schema
    """
    # 收集需要翻譯的備註
    comments_to_translate = []
    for i, verdict in enumerate(schema.clause_verdicts):
        if verdict.comment_en and not verdict.comment_zh:
            comments_to_translate.append({
                "index": i,
                "clause": verdict.clause,
                "comment_en": verdict.comment_en
            })

    if not comments_to_translate:
        return schema

    logger.info(f"翻譯 {len(comments_to_translate)} 個條文備註...")

    # 批次翻譯（避免太多 API 呼叫）
    batch_size = 10
    for batch_start in range(0, len(comments_to_translate), batch_size):
        batch = comments_to_translate[batch_start:batch_start + batch_size]

        # 準備批次翻譯 prompt
        batch_content = "\n".join([
            f"[{item['clause']}] {item['comment_en']}"
            for item in batch
        ])

        messages = [
            {"role": "system", "content": "請將以下 CB 報告條文備註翻譯成簡潔的繁體中文。每行格式為 [條文編號] 英文備註。請以相同格式回傳翻譯結果，用 JSON 格式 {\"translations\": [{\"clause\": \"條文編號\", \"comment_zh\": \"中文翻譯\"}]}。"},
            {"role": "user", "content": batch_content}
        ]

        try:
            response = _call_llm(messages, temperature=0.3)
            # 使用 return_empty_on_fail=True 確保即使解析失敗也不會中斷
            result = _parse_llm_json_response(response, return_empty_on_fail=True)

            if not result:
                logger.warning("條文備註翻譯結果為空，跳過此批次")
                continue

            # 更新翻譯
            translations_list = result.get("translations", [])
            if isinstance(translations_list, list):
                translations_map = {}
                for t in translations_list:
                    if isinstance(t, dict) and "clause" in t and "comment_zh" in t:
                        translations_map[t["clause"]] = t["comment_zh"]

                for item in batch:
                    if item["clause"] in translations_map:
                        schema.clause_verdicts[item["index"]].comment_zh = translations_map[item["clause"]]

        except Exception as e:
            logger.error(f"翻譯條文備註時發生錯誤: {e}")

    return schema


# ==============================================
# Checkbox Flags Inference
# ==============================================

def _infer_checkbox_flags(schema: ReportSchema) -> ReportSchema:
    """
    根據 test_item_particulars 推斷 checkbox_flags

    Args:
        schema: 原始 schema

    Returns:
        更新後的 schema
    """
    tip = schema.test_item_particulars
    flags = schema.checkbox_flags

    # 產品群組
    product_group = (tip.product_group or "").upper()
    if "AV" in product_group and "ICT" in product_group:
        flags.is_av_ict = True
    elif "AV" in product_group:
        flags.is_av = True
    elif "ICT" in product_group:
        flags.is_ict = True
    elif "TELECOM" in product_group:
        flags.is_telecom = True

    # 使用分類
    for classification in tip.classification_of_use:
        classification_upper = classification.upper()
        if "ORDINARY" in classification_upper:
            flags.is_ordinary = True
        if "SKILLED" in classification_upper:
            flags.is_skilled = True
        if "INSTRUCTED" in classification_upper:
            flags.is_instructed = True

    # 電源等級
    for connection in tip.supply_connection:
        connection_upper = connection.upper()
        if "CLASS I" in connection_upper or "CLASS 1" in connection_upper:
            flags.is_class_i = True
        if "CLASS II" in connection_upper or "CLASS 2" in connection_upper:
            flags.is_class_ii = True
        if "CLASS III" in connection_upper or "CLASS 3" in connection_upper:
            flags.is_class_iii = True

    # 移動性
    mobility = (tip.mobility or "").upper()
    if "PORTABLE" in mobility:
        flags.is_portable = True
    if "STATIONARY" in mobility:
        flags.is_stationary = True
    if "FIXED" in mobility:
        flags.is_fixed = True

    schema.checkbox_flags = flags
    return schema


# ==============================================
# Main Export Function
# ==============================================

async def extract_report_schema_from_adobe_json(
    adobe_json: dict,
    max_concurrent: int = None
) -> Tuple[ReportSchema, dict]:
    """
    主要函式：將 Adobe Extract 結果轉換為統一 Schema

    這是此模組的主要入口點。

    流程：
    1. 將 Adobe JSON 分成多個 chunks
    2. 並發處理多個 chunks（加速處理）
    3. 合併所有 chunk 的結果
    4. 執行繁體中文翻譯
    5. 推斷 checkbox flags

    Args:
        adobe_json: Adobe Extract 的結果（來自 adobe_extract.py）
        max_concurrent: 最大並發數（預設 5，避免 API rate limit）

    Returns:
        Tuple[ReportSchema, dict]: (完整的 ReportSchema 物件, 統計資訊)

    Usage:
        >>> adobe_result = await extract_pdf_to_json(pdf_bytes)
        >>> schema, stats = await extract_report_schema_from_adobe_json(adobe_result)
        >>> print(schema.basic_info.cb_report_no)
        >>> print(f"Token 使用量: {stats['total_tokens']}")
    """
    global _token_tracker

    # 重置 token tracker
    _token_tracker = TokenUsageTracker()

    # 使用設定值或預設值
    if max_concurrent is None:
        max_concurrent = settings.llm_max_concurrent

    logger.info("開始從 Adobe JSON 萃取 Report Schema...")
    logger.info(f"並發設定: max_concurrent={max_concurrent}")

    # Step 1: 分 chunks
    chunks = _prepare_chunks(adobe_json)
    total_chunks = len(chunks)

    # Step 2: 並發處理 chunks
    merged_schema = create_empty_schema()

    # 使用 ThreadPoolExecutor 進行並發處理（因為 _call_llm 是同步的）
    loop = asyncio.get_event_loop()

    # 分批處理以控制並發數
    for batch_start in range(0, total_chunks, max_concurrent):
        batch_end = min(batch_start + max_concurrent, total_chunks)
        batch_chunks = chunks[batch_start:batch_end]

        logger.info(f"並發處理 chunks {batch_start + 1}-{batch_end}/{total_chunks}")

        # 使用 ThreadPoolExecutor 並發呼叫
        with ThreadPoolExecutor(max_workers=max_concurrent) as executor:
            futures = []
            for i, chunk in enumerate(batch_chunks):
                chunk_index = batch_start + i
                future = loop.run_in_executor(
                    executor,
                    _process_chunk,
                    chunk,
                    chunk_index,
                    total_chunks
                )
                futures.append(future)

            # 等待所有 futures 完成
            results = await asyncio.gather(*futures, return_exceptions=True)

        # 合併結果
        for i, result in enumerate(results):
            chunk_index = batch_start + i
            if isinstance(result, Exception):
                logger.error(f"處理 chunk {chunk_index + 1} 時發生錯誤: {result}")
                continue
            merged_schema = merge_schemas(merged_schema, result)

    # Step 3: 推斷 checkbox flags
    merged_schema = _infer_checkbox_flags(merged_schema)

    # Step 4: 翻譯成繁體中文
    merged_schema = _translate_to_chinese(merged_schema)

    # 設定 metadata
    merged_schema.extraction_timestamp = datetime.now().isoformat()
    merged_schema.extraction_version = "1.0.0"

    # 準備統計資訊
    stats = {
        "total_tokens": _token_tracker.total_tokens,
        "prompt_tokens": _token_tracker.total_prompt_tokens,
        "completion_tokens": _token_tracker.total_completion_tokens,
        "llm_calls": _token_tracker.call_count,
        "estimated_cost": _token_tracker.calculate_cost(),
        "total_chunks": total_chunks
    }

    logger.info("Report Schema 萃取完成")
    logger.info(f"  - 基本資料: {merged_schema.basic_info.cb_report_no}")
    logger.info(f"  - 系列型號數: {len(merged_schema.series_models)}")
    logger.info(f"  - 條文判定數: {len(merged_schema.clause_verdicts)}")
    logger.info(f"  - Token 使用量: {stats['total_tokens']:,}")
    logger.info(f"  - 預估成本: ${stats['estimated_cost']:.4f}")

    return merged_schema, stats


# ==============================================
# Synchronous Wrapper
# ==============================================

def extract_report_schema_from_adobe_json_sync(adobe_json: dict) -> Tuple[ReportSchema, dict]:
    """
    同步版本的 extract_report_schema_from_adobe_json

    用於不支援 async 的環境

    Args:
        adobe_json: Adobe Extract 的結果

    Returns:
        Tuple[ReportSchema, dict]: (完整的 ReportSchema 物件, 統計資訊)
    """
    import asyncio

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        return loop.run_until_complete(
            extract_report_schema_from_adobe_json(adobe_json)
        )
    finally:
        loop.close()


# ==============================================
# Testing / Development Utilities
# ==============================================

def create_mock_schema() -> ReportSchema:
    """
    建立模擬的 ReportSchema（用於開發測試）
    """
    return ReportSchema(
        basic_info=BasicInfo(
            cb_report_no="TW-12345-UL",
            standard="IEC 62368-1:2018",
            applicant_en="ABC Technology Co., Ltd.",
            applicant_address_en="No. 123, Tech Road, Hsinchu, Taiwan",
            manufacturer_en="XYZ Manufacturing Inc.",
            manufacturer_address_en="No. 456, Industry Blvd, Shenzhen, China",
            product_name_en="Switching Power Supply",
            model_main="SPS-120W",
            ratings_input="100-240Vac, 50/60Hz, 2A",
            ratings_output="12Vdc, 10A",
            issue_date="2024-01-15"
        ),
        test_item_particulars=TestItemParticulars(
            product_group="ICT",
            classification_of_use=["Ordinary"],
            supply_connection=["Class I"],
            ovc="OVC II",
            pollution_degree="2",
            ip_code="IP20",
            tma="40°C",
            altitude_limit_m=2000
        ),
        series_models=[
            SeriesModel(model="SPS-120W-A", vout="12V", iout="10A", pout="120W", case_type="Metal"),
            SeriesModel(model="SPS-120W-B", vout="24V", iout="5A", pout="120W", case_type="Metal"),
            SeriesModel(model="SPS-120W-C", vout="48V", iout="2.5A", pout="120W", case_type="Plastic"),
        ],
        clause_verdicts=[
            ClauseVerdict(clause="4.1.1", verdict="P", comment_en="Product properly classified"),
            ClauseVerdict(clause="4.2.1", verdict="P", comment_en="Energy sources identified"),
            ClauseVerdict(clause="5.4.2", verdict="P", comment_en="Temperature rise within limits"),
        ],
        translations=Translations(
            applicant_zh="ABC 科技股份有限公司",
            manufacturer_zh="XYZ 製造有限公司",
            product_name_zh="交換式電源供應器"
        ),
        checkbox_flags=CheckboxFlags(
            is_ict=True,
            is_ordinary=True,
            is_class_i=True
        ),
        extraction_timestamp=datetime.now().isoformat()
    )
