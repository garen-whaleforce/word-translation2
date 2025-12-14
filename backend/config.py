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
    # Adobe PDF Services 設定
    # ==============================================
    adobe_client_id: str = Field(
        ...,
        description="Adobe PDF Services Client ID"
    )
    adobe_client_secret: str = Field(
        ...,
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
