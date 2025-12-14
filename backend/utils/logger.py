"""
==============================================
Logger Configuration
簡單的 logging 設定模組
==============================================
"""

import logging
import sys
from typing import Optional


# 全域 logging 設定狀態
_logging_configured = False


def setup_logging(
    level: int = logging.INFO,
    format_string: Optional[str] = None
) -> None:
    """
    設定全域 logging 格式

    Args:
        level: logging 等級（預設 INFO）
        format_string: 自訂格式字串
    """
    global _logging_configured

    if _logging_configured:
        return

    if format_string is None:
        format_string = (
            "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
        )

    # 建立 handler
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(level)

    # 設定格式
    formatter = logging.Formatter(format_string, datefmt="%Y-%m-%d %H:%M:%S")
    handler.setFormatter(formatter)

    # 設定 root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    # 清除既有的 handlers（避免重複輸出）
    root_logger.handlers = []
    root_logger.addHandler(handler)

    # 降低第三方 library 的 logging 等級
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("httpcore").setLevel(logging.WARNING)
    logging.getLogger("openai").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)

    _logging_configured = True


def get_logger(name: str) -> logging.Logger:
    """
    取得指定名稱的 logger

    會自動初始化 logging 設定

    Args:
        name: logger 名稱（通常使用 __name__）

    Returns:
        logging.Logger 物件

    Usage:
        >>> from utils.logger import get_logger
        >>> logger = get_logger(__name__)
        >>> logger.info("Hello, world!")
    """
    # 確保 logging 已設定
    setup_logging()

    return logging.getLogger(name)


# ==============================================
# Convenience Functions
# ==============================================

def log_separator(logger: logging.Logger, char: str = "=", length: int = 50) -> None:
    """輸出分隔線"""
    logger.info(char * length)


def log_section(logger: logging.Logger, title: str) -> None:
    """輸出區段標題"""
    logger.info("")
    log_separator(logger)
    logger.info(f"  {title}")
    log_separator(logger)
