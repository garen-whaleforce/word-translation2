"""
Schemas Package
定義所有資料結構與 JSON Schema
"""

from .report_schema import (
    ReportSchema,
    BasicInfo,
    TestItemParticulars,
    SeriesModel,
    ClauseVerdict,
    KeyTables,
    InputTestRow,
    TemperatureRiseRow,
    EnergySourceRow,
    Translations,
    CheckboxFlags,
)

__all__ = [
    "ReportSchema",
    "BasicInfo",
    "TestItemParticulars",
    "SeriesModel",
    "ClauseVerdict",
    "KeyTables",
    "InputTestRow",
    "TemperatureRiseRow",
    "EnergySourceRow",
    "Translations",
    "CheckboxFlags",
]
