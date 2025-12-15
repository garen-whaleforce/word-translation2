"""
==============================================
Configuration Module
讀取 .env 環境變數，提供全域 settings 物件
==============================================
"""

from functools import lru_cache
from pydantic_settings import BaseSettings
from pydantic import Field
from typing import Optional


class Settings(BaseSettings):
    """
    應用程式設定
    所有敏感資訊皆從 .env 檔案讀取
    """

    # ==============================================
    # Azure OpenAI 設定
    # ==============================================
    azure_openai_endpoint: str = Field(
        ...,
        description="Azure OpenAI 端點 URL，例如 https://your-resource.openai.azure.com/"
    )
    azure_openai_api_key: str = Field(
        ...,
        description="Azure OpenAI API Key"
    )
    azure_openai_deployment: str = Field(
        default="gpt-5.1",
        description="Azure OpenAI 模型部署名稱（例如 gpt-5.1, gpt-4o）"
    )
    azure_openai_api_version: str = Field(
        default="2024-12-01-preview",
        description="Azure OpenAI API 版本"
    )

    # ==============================================
    # PDF 擷取設定
    # ==============================================
    pdf_extractor: str = Field(
        default="pymupdf",
        description="PDF 擷取引擎：'pymupdf'（免費）或 'adobe'（需 API key）"
    )

    # ==============================================
    # Adobe PDF Services 設定（可選，當 pdf_extractor=adobe 時需要）
    # ==============================================
    adobe_client_id: Optional[str] = Field(
        default=None,
        description="Adobe PDF Services Client ID"
    )
    adobe_client_secret: Optional[str] = Field(
        default=None,
        description="Adobe PDF Services Client Secret"
    )
    adobe_pdf_services_base_url: str = Field(
        default="https://pdf-services.adobe.io",
        description="Adobe PDF Services API Base URL"
    )

    # ==============================================
    # 應用程式設定
    # ==============================================
    app_name: str = Field(
        default="CB to CNS Report Generator",
        description="應用程式名稱"
    )
    debug: bool = Field(
        default=False,
        description="除錯模式"
    )
    template_dir: str = Field(
        default="templates",
        description="Word 模板資料夾路徑"
    )
    temp_dir: str = Field(
        default="/tmp/reports",
        description="暫存檔案資料夾"
    )
    max_pdf_size_mb: int = Field(
        default=50,
        description="最大允許的 PDF 檔案大小 (MB)"
    )

    # ==============================================
    # LLM 相關設定
    # ==============================================
    llm_max_tokens: int = Field(
        default=16384,
        description="LLM 回應的最大 token 數"
    )
    llm_temperature: float = Field(
        default=0.1,
        description="LLM 溫度參數（越低越穩定）"
    )
    llm_chunk_pages: int = Field(
        default=5,
        description="每次送給 LLM 的頁數"
    )
    llm_max_concurrent: int = Field(
        default=5,
        description="LLM 並發呼叫數量（避免 API rate limit）"
    )

    # ==============================================
    # 實驗室固定資訊
    # ==============================================
    lab_name: str = Field(
        default="安捷檢測有限公司",
        description="實驗室名稱"
    )
    lab_address: str = Field(
        default="新北市新店區寶興路45巷8弄16號4樓",
        description="實驗室地址"
    )
    lab_accreditation_no: str = Field(
        default="SL2-IN/VA-T-0157",
        description="標準檢驗局指定試驗室認可編號"
    )
    lab_altitude: str = Field(
        default="約 50 m",
        description="實驗室海拔高度"
    )

    # ==============================================
    # 報告預設值
    # ==============================================
    default_test_type: str = Field(
        default="型式試驗",
        description="預設試驗方式"
    )
    default_cns_standard: str = Field(
        default="CNS 15598-1",
        description="預設 CNS 標準"
    )
    default_cns_standard_version: str = Field(
        default="109年版",
        description="預設 CNS 標準版本"
    )

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"
        case_sensitive = False  # 環境變數名稱不區分大小寫


@lru_cache()
def get_settings() -> Settings:
    """
    取得設定物件（使用 lru_cache 確保只初始化一次）
    """
    return Settings()


# 全域 settings 物件，方便直接 import 使用
settings = get_settings()
