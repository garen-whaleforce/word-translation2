"""
==============================================
Report Schema Definition
CB 報告 → CNS 報告的統一 JSON Schema
==============================================

此模組定義了從 CB 報告萃取出來的所有結構化資料。
使用 Pydantic models 確保型別安全與資料驗證。
"""

from pydantic import BaseModel, Field, field_validator
from typing import List, Optional, Any
from enum import Enum


# ==============================================
# Enums 列舉定義
# ==============================================

class VerdictType(str, Enum):
    """條文判定結果"""
    PASS = "P"
    FAIL = "F"
    NOT_APPLICABLE = "N/A"
    NOT_TESTED = "NT"
    CONDITIONAL = "C"


class ProductGroup(str, Enum):
    """產品群組分類"""
    AV = "AV"
    ICT = "ICT"
    AV_ICT = "Audio/Video & ICT"
    TELECOM = "Telecom"
    OTHER = "Other"


class SupplyConnection(str, Enum):
    """電源連接方式"""
    CLASS_I = "Class I"
    CLASS_II = "Class II"
    CLASS_III = "Class III"


class ClassificationOfUse(str, Enum):
    """使用分類"""
    ORDINARY = "Ordinary"
    SKILLED = "Skilled"
    INSTRUCTED = "Instructed"


# ==============================================
# 子 Schema 定義
# ==============================================

class BasicInfo(BaseModel):
    """
    基本資料區塊
    包含報告編號、標準、申請人、製造商、產品資訊等
    """
    # 報告資訊
    cb_report_no: str = Field(default="", description="CB 報告編號")
    ast_report_no: Optional[str] = Field(default=None, description="AST 報告編號（用於頁首/頁尾）")
    cns_report_no: Optional[str] = Field(default=None, description="CNS 報告編號（由我方填寫）")
    bsmi_designated_report_no: Optional[str] = Field(default=None, description="BSMI 指定報告編號，如 SL2INT0157250509")
    standard: str = Field(default="", description="適用標準，例如 IEC 62368-1:2018")
    standard_version: Optional[str] = Field(default=None, description="標準版本")
    cns_standard: Optional[str] = Field(default=None, description="CNS 標準，如 CNS 15598-1")
    cns_standard_version: Optional[str] = Field(default=None, description="CNS 標準版本，如 109年版")
    national_differences: Optional[str] = Field(default=None, description="國家差異")
    test_lab: Optional[str] = Field(default=None, description="測試實驗室名稱")
    test_lab_country: Optional[str] = Field(default=None, description="測試實驗室國家")
    cb_scheme_member: Optional[str] = Field(default=None, description="CB Scheme 會員機構")

    # 申請人資訊
    applicant_en: str = Field(default="", description="申請人名稱（英文）")
    applicant_address_en: str = Field(default="", description="申請人地址（英文）")

    # 製造商資訊
    manufacturer_en: str = Field(default="", description="製造商名稱（英文）")
    manufacturer_address_en: str = Field(default="", description="製造商地址（英文）")
    factory_name_en: Optional[str] = Field(default=None, description="工廠名稱（英文）")
    factory_address_en: Optional[str] = Field(default=None, description="工廠地址（英文）")

    # 產品資訊
    product_name_en: str = Field(default="", description="產品名稱（英文）")
    model_main: str = Field(default="", description="主型號")
    brand: Optional[str] = Field(default=None, description="品牌")
    trademark: Optional[str] = Field(default=None, description="商標")

    # 額定值
    ratings_input: str = Field(default="", description="輸入額定值，例如 100-240Vac, 50/60Hz, 2A")
    ratings_output: str = Field(default="", description="輸出額定值，例如 12Vdc, 5A")
    ratings_power: Optional[str] = Field(default=None, description="功率額定值")
    rated_output_lines: Optional[List[str]] = Field(default=None, description="輸出額定值（多行列表，用於區塊填寫）")
    # 最大輸出（供 sanity 檢核與文案）
    max_output_v: Optional[str] = Field(default=None, description="最大輸出電壓")
    max_output_a: Optional[str] = Field(default=None, description="最大輸出電流")
    max_output_w: Optional[str] = Field(default=None, description="最大輸出功率")

    # 報告日期
    issue_date: Optional[str] = Field(default=None, description="報告發行日期")
    issue_date_short: Optional[str] = Field(default=None, description="報告發行日期（短格式，如 113.03.27）")
    receive_date: Optional[str] = Field(default=None, description="試驗件收件日")
    test_date_from: Optional[str] = Field(default=None, description="測試開始日期")
    test_date_to: Optional[str] = Field(default=None, description="測試結束日期")

    # CB 報告資訊
    cb_test_lab: Optional[str] = Field(default=None, description="CB 測試實驗室")
    cb_certificate_no: Optional[str] = Field(default=None, description="CB 證書編號")
    cb_standard: Optional[str] = Field(default=None, description="CB 適用標準")

    # 設備資訊
    equipment_mass: Optional[str] = Field(default=None, description="設備質量")
    protection_rating: Optional[str] = Field(default=None, description="保護裝置額定電流")

    # 試驗相關資訊
    test_type: Optional[str] = Field(default=None, description="試驗方式：型式試驗/監督試驗")
    overall_result: Optional[str] = Field(default=None, description="整體試驗結果：符合/不符合")
    sample_conforms: Optional[str] = Field(default=None, description="符合項目說明")
    sample_not_conforms: Optional[str] = Field(default=None, description="不符合項目說明")
    not_applicable_items: Optional[str] = Field(default=None, description="不適用項目說明")
    special_installation: Optional[str] = Field(default=None, description="特殊安裝要求")
    national_differences_summary: Optional[str] = Field(default=None, description="國別差異摘要")
    model_differences: Optional[str] = Field(default=None, description="型號差異說明")
    cb_report_note: Optional[str] = Field(default=None, description="CB 報告備註/引用來源")
    attachment_list: Optional[List[str]] = Field(default=None, description="附件清單")
    temperature_requirements_text: Optional[str] = Field(default=None, description="溫度/負載條件敘述，用於替換舊案33W段落")


class TestItemParticulars(BaseModel):
    """
    試驗樣品特性（Test Item Particulars）
    通常位於 CB 報告的前幾頁
    """
    # 產品分類
    product_group: Optional[str] = Field(default=None, description="產品群組：AV / ICT / Telecom 等")

    # 使用分類（可多選）
    classification_of_use: List[str] = Field(
        default_factory=list,
        description="使用分類：Ordinary / Skilled / Instructed"
    )

    # 電源連接（可多選）
    supply_connection: List[str] = Field(
        default_factory=list,
        description="電源連接方式：Class I / II / III"
    )

    # 環境條件
    ovc: Optional[str] = Field(default=None, description="過電壓類別 (OVC)，例如 OVC II")
    pollution_degree: Optional[str] = Field(default=None, description="污染等級，例如 2")
    ip_code: Optional[str] = Field(default=None, description="IP 防護等級，例如 IP20")
    tma: Optional[str] = Field(default=None, description="最高環境溫度 Tma，例如 40°C")
    altitude_limit_m: Optional[int] = Field(default=None, description="海拔高度限制 (m)，例如 2000")

    # 安裝與操作
    installation_type: Optional[str] = Field(default=None, description="安裝類型")
    operating_conditions: Optional[str] = Field(default=None, description="操作條件")

    # 電氣特性
    mains_supply: Optional[str] = Field(default=None, description="主電源類型：AC / DC / AC+DC")
    rated_voltage: Optional[str] = Field(default=None, description="額定電壓")
    rated_frequency: Optional[str] = Field(default=None, description="額定頻率")
    rated_current: Optional[str] = Field(default=None, description="額定電流")

    # 其他特性
    protection_class: Optional[str] = Field(default=None, description="保護等級")
    insulation_type: Optional[str] = Field(default=None, description="絕緣類型")
    mobility: Optional[str] = Field(default=None, description="移動性：Portable / Stationary / Fixed")

    # 備註
    additional_info: Optional[str] = Field(default=None, description="其他特性說明")

    # Validators 處理 LLM 回傳的非預期格式
    @field_validator('pollution_degree', 'ovc', 'ip_code', 'tma', 'mobility',
                     'installation_type', 'operating_conditions', 'mains_supply',
                     'rated_voltage', 'rated_frequency', 'rated_current',
                     'protection_class', 'insulation_type', 'additional_info',
                     'product_group', mode='before')
    @classmethod
    def convert_to_string(cls, v: Any) -> Optional[str]:
        """將各種類型轉換為字串"""
        if v is None:
            return None
        if isinstance(v, list):
            return ', '.join(str(item) for item in v)
        return str(v)


class RevisionRecord(BaseModel):
    """
    報告修訂記錄
    """
    item: str = Field(default="01", description="項次")
    date: Optional[str] = Field(default=None, description="發行日期")
    report_no: Optional[str] = Field(default=None, description="報告編號")
    description: str = Field(default="主報告", description="修訂內容")


class SeriesModel(BaseModel):
    """
    系列型號資訊
    一份 CB 報告可能涵蓋多個系列型號
    """
    model: str = Field(default="", description="型號名稱")
    vout: Optional[str] = Field(default=None, description="輸出電壓")
    iout: Optional[str] = Field(default=None, description="輸出電流")
    pout: Optional[str] = Field(default=None, description="輸出功率")
    vin: Optional[str] = Field(default=None, description="輸入電壓")
    iin: Optional[str] = Field(default=None, description="輸入電流")
    case_type: Optional[str] = Field(default=None, description="外殼類型：Metal / Plastic / Open Frame")
    connector_type: Optional[str] = Field(default=None, description="連接器類型")
    differences: Optional[str] = Field(default=None, description="與主型號的差異說明")
    remarks: Optional[str] = Field(default=None, description="備註")


class ClauseVerdict(BaseModel):
    """
    條文判定結果
    記錄每個條文的測試結果
    """
    clause: str = Field(default="", description="條文編號，例如 4.1.1")
    clause_title: Optional[str] = Field(default=None, description="條文標題")
    verdict: str = Field(default="P", description="判定結果：P / N/A / F / NT / C")
    comment_en: Optional[str] = Field(default=None, description="英文備註")
    comment_zh: Optional[str] = Field(default=None, description="繁體中文備註（由 LLM 翻譯）")
    test_method: Optional[str] = Field(default=None, description="測試方法")
    reference: Optional[str] = Field(default=None, description="參考資料")


# ==============================================
# 關鍵測試表格的 Row 定義
# ==============================================

class InputTestRow(BaseModel):
    """輸入測試表格的一行資料"""
    test_condition: Optional[str] = Field(default=None, description="測試條件")
    voltage: Optional[str] = Field(default=None, description="電壓")
    current: Optional[str] = Field(default=None, description="電流")
    power: Optional[str] = Field(default=None, description="功率")
    frequency: Optional[str] = Field(default=None, description="頻率")
    power_factor: Optional[str] = Field(default=None, description="功率因數")
    remarks: Optional[str] = Field(default=None, description="備註")


class TemperatureRiseRow(BaseModel):
    """溫升測試表格的一行資料"""
    location: Optional[str] = Field(default=None, description="量測位置")
    component: Optional[str] = Field(default=None, description="元件名稱")
    measured_temp: Optional[str] = Field(default=None, description="量測溫度 (°C)")
    ambient_temp: Optional[str] = Field(default=None, description="環境溫度 (°C)")
    temp_rise: Optional[str] = Field(default=None, description="溫升 (K)")
    limit: Optional[str] = Field(default=None, description="限值 (K)")
    verdict: Optional[str] = Field(default=None, description="判定")
    remarks: Optional[str] = Field(default=None, description="備註")


class EnergySourceRow(BaseModel):
    """能量來源表格的一行資料"""
    energy_source: Optional[str] = Field(default=None, description="能量來源類型")
    class_level: Optional[str] = Field(default=None, description="等級 (ES1/ES2/ES3)")
    voltage: Optional[str] = Field(default=None, description="電壓")
    current: Optional[str] = Field(default=None, description="電流")
    power: Optional[str] = Field(default=None, description="功率")
    location: Optional[str] = Field(default=None, description="位置")
    safeguard: Optional[str] = Field(default=None, description="防護措施")
    remarks: Optional[str] = Field(default=None, description="備註")


class FactoryInfo(BaseModel):
    """工廠資訊"""
    name: str = Field(default="", description="工廠名稱")
    address: str = Field(default="", description="工廠地址")


class KeyTables(BaseModel):
    """
    關鍵測試表格
    從 CB 報告中萃取的重要測試數據表格
    """
    input_tests: List[InputTestRow] = Field(
        default_factory=list,
        description="輸入測試數據"
    )
    temperature_rise: List[TemperatureRiseRow] = Field(
        default_factory=list,
        description="溫升測試數據"
    )
    energy_sources: List[EnergySourceRow] = Field(
        default_factory=list,
        description="能量來源分類表"
    )
    # 原樣表格（若解析成列/行）
    input_test_raw: Optional[List[List[Any]]] = Field(
        default=None,
        description="輸入試驗原始表格資料（行列表）"
    )
    abnormal_fault_raw: Optional[List[List[Any]]] = Field(
        default=None,
        description="異常/故障試驗原始表格資料（行列表）"
    )

    # 可擴充其他表格類型
    dielectric_test: Optional[List[dict]] = Field(
        default=None,
        description="耐壓測試數據"
    )
    leakage_current: Optional[List[dict]] = Field(
        default=None,
        description="漏電流測試數據"
    )
    abnormal_test: Optional[List[dict]] = Field(
        default=None,
        description="異常測試數據"
    )
    component_list: Optional[List[dict]] = Field(
        default=None,
        description="關鍵零組件清單"
    )


class Translations(BaseModel):
    """
    繁體中文翻譯欄位
    由 LLM 自動翻譯產生
    """
    applicant_zh: Optional[str] = Field(default=None, description="申請人名稱（繁中）")
    applicant_address_zh: Optional[str] = Field(default=None, description="申請人地址（繁中）")
    manufacturer_zh: Optional[str] = Field(default=None, description="製造商名稱（繁中）")
    manufacturer_address_zh: Optional[str] = Field(default=None, description="製造商地址（繁中）")
    product_name_zh: Optional[str] = Field(default=None, description="產品名稱（繁中）")
    factory_name_zh: Optional[str] = Field(default=None, description="工廠名稱（繁中）")
    factory_address_zh: Optional[str] = Field(default=None, description="工廠地址（繁中）")
    # 多工廠支援
    factory_name_1: Optional[str] = Field(default=None, description="工廠1名稱")
    factory_address_1: Optional[str] = Field(default=None, description="工廠1地址")
    factory_name_2: Optional[str] = Field(default=None, description="工廠2名稱")
    factory_address_2: Optional[str] = Field(default=None, description="工廠2地址")
    additional_translations: Optional[dict] = Field(
        default=None,
        description="其他需要翻譯的欄位"
    )


class CheckboxFlags(BaseModel):
    """
    Checkbox / Flag 狀態
    用於控制 Word 模板中的勾選框
    """
    # 產品群組
    is_av: bool = Field(default=False, description="是否為 AV 產品")
    is_ict: bool = Field(default=False, description="是否為 ICT 產品")
    is_av_ict: bool = Field(default=False, description="是否為 AV & ICT 產品")
    is_telecom: bool = Field(default=False, description="是否為 Telecom 產品")

    # 使用分類
    is_ordinary: bool = Field(default=False, description="一般使用者")
    is_skilled: bool = Field(default=False, description="專業人員")
    is_instructed: bool = Field(default=False, description="受指導人員")

    # 電源等級
    is_class_i: bool = Field(default=False, description="Class I")
    is_class_ii: bool = Field(default=False, description="Class II")
    is_class_iii: bool = Field(default=False, description="Class III")

    # 移動性 / 設備移動性
    is_direct_plugin: bool = Field(default=False, description="直插式設備")
    is_stationary: bool = Field(default=False, description="放置式設備")
    is_building_in: bool = Field(default=False, description="崁入式設備")
    is_wall_ceiling: bool = Field(default=False, description="壁面/天花板安裝式")
    is_rack_mounted: bool = Field(default=False, description="SRME/機架安裝")
    is_portable: bool = Field(default=False, description="可攜式/可移動式")
    is_fixed: bool = Field(default=False, description="固定式/永久安裝")

    # 其他常見選項
    is_pluggable_a: bool = Field(default=False, description="Pluggable Type A")
    is_pluggable_b: bool = Field(default=False, description="Pluggable Type B")
    is_permanently_connected: bool = Field(default=False, description="永久連接")


# ==============================================
# 主 Schema：整合所有子區塊
# ==============================================

class ReportSchema(BaseModel):
    """
    CB 報告完整 Schema
    整合所有子區塊，作為系統內部的統一資料格式
    """
    # 基本資料
    basic_info: BasicInfo = Field(
        default_factory=BasicInfo,
        description="基本資料區塊"
    )

    # 試驗樣品特性
    test_item_particulars: TestItemParticulars = Field(
        default_factory=TestItemParticulars,
        description="試驗樣品特性"
    )

    # 系列型號清單
    series_models: List[SeriesModel] = Field(
        default_factory=list,
        description="系列型號清單"
    )

    # 修訂記錄
    revision_records: List[RevisionRecord] = Field(
        default_factory=list,
        description="報告修訂記錄清單"
    )

    # 條文判定清單
    clause_verdicts: List[ClauseVerdict] = Field(
        default_factory=list,
        description="條文判定結果清單"
    )

    # 關鍵測試表格
    key_tables: KeyTables = Field(
        default_factory=KeyTables,
        description="關鍵測試表格"
    )

    # 工廠清單
    factories: List[FactoryInfo] = Field(
        default_factory=list,
        description="工廠清單（名稱與地址）"
    )

    # 附件/附件描述
    attachments: Optional[List[str]] = Field(
        default=None,
        description="附件清單"
    )

    # 繁中翻譯
    translations: Translations = Field(
        default_factory=Translations,
        description="繁體中文翻譯"
    )

    # Checkbox 狀態
    checkbox_flags: CheckboxFlags = Field(
        default_factory=CheckboxFlags,
        description="勾選框狀態"
    )

    # 元資料
    extraction_version: str = Field(
        default="1.0.0",
        description="Schema 版本號"
    )
    extraction_timestamp: Optional[str] = Field(
        default=None,
        description="萃取時間戳記"
    )
    source_filename: Optional[str] = Field(
        default=None,
        description="來源 PDF 檔名"
    )
    extraction_notes: Optional[str] = Field(
        default=None,
        description="萃取過程備註"
    )

    class Config:
        """Pydantic 設定"""
        json_schema_extra = {
            "example": {
                "basic_info": {
                    "cb_report_no": "TW-12345-UL",
                    "standard": "IEC 62368-1:2018",
                    "applicant_en": "ABC Technology Co., Ltd.",
                    "manufacturer_en": "XYZ Manufacturing Inc.",
                    "product_name_en": "Power Adapter",
                    "model_main": "PA-120W",
                    "ratings_input": "100-240Vac, 50/60Hz, 2A",
                    "ratings_output": "12Vdc, 10A"
                },
                "series_models": [
                    {
                        "model": "PA-120W-A",
                        "vout": "12V",
                        "iout": "10A",
                        "pout": "120W"
                    }
                ],
                "translations": {
                    "applicant_zh": "ABC 科技股份有限公司",
                    "product_name_zh": "電源供應器"
                }
            }
        }


# ==============================================
# Helper Functions
# ==============================================

def create_empty_schema() -> ReportSchema:
    """建立一個空的 ReportSchema 物件"""
    return ReportSchema()


def merge_schemas(base: ReportSchema, update: ReportSchema) -> ReportSchema:
    """
    合併兩個 Schema（用於多次 LLM 呼叫後合併結果）

    規則：
    - basic_info: 取 update 中非空的欄位覆蓋 base
    - series_models: 累加（去重）
    - clause_verdicts: 以 clause 為 key 合併
    - key_tables: 累加
    - translations: 取 update 中非空的欄位覆蓋 base
    - checkbox_flags: OR 運算（任一為 True 則為 True）
    """
    merged = base.model_copy(deep=True)

    # 合併 basic_info
    base_dict = merged.basic_info.model_dump()
    update_dict = update.basic_info.model_dump()
    for key, value in update_dict.items():
        if value is not None and value != "":
            base_dict[key] = value
    merged.basic_info = BasicInfo(**base_dict)

    # 合併 test_item_particulars
    base_tip = merged.test_item_particulars.model_dump()
    update_tip = update.test_item_particulars.model_dump()
    for key, value in update_tip.items():
        if value is not None and value != "" and value != []:
            if isinstance(value, list):
                # List 欄位做合併
                existing = base_tip.get(key, [])
                base_tip[key] = list(set(existing + value))
            else:
                base_tip[key] = value
    merged.test_item_particulars = TestItemParticulars(**base_tip)

    # 合併 series_models（以 model 名稱去重）
    existing_models = {m.model: m for m in merged.series_models}
    for model in update.series_models:
        if model.model and model.model not in existing_models:
            existing_models[model.model] = model
    merged.series_models = list(existing_models.values())

    # 合併 clause_verdicts（以 clause 為 key）
    existing_clauses = {c.clause: c for c in merged.clause_verdicts}
    for verdict in update.clause_verdicts:
        if verdict.clause:
            existing_clauses[verdict.clause] = verdict
    merged.clause_verdicts = list(existing_clauses.values())

    # 合併 key_tables
    merged.key_tables.input_tests.extend(update.key_tables.input_tests)
    merged.key_tables.temperature_rise.extend(update.key_tables.temperature_rise)
    merged.key_tables.energy_sources.extend(update.key_tables.energy_sources)
    if update.key_tables.input_test_raw:
        merged.key_tables.input_test_raw = update.key_tables.input_test_raw
    if update.key_tables.abnormal_fault_raw:
        merged.key_tables.abnormal_fault_raw = update.key_tables.abnormal_fault_raw

    # 合併 translations
    base_trans = merged.translations.model_dump()
    update_trans = update.translations.model_dump()
    for key, value in update_trans.items():
        if value is not None and value != "":
            base_trans[key] = value
    merged.translations = Translations(**base_trans)

    # 合併 checkbox_flags（OR 運算）
    base_flags = merged.checkbox_flags.model_dump()
    update_flags = update.checkbox_flags.model_dump()
    for key, value in update_flags.items():
        if value:  # 如果 update 中為 True
            base_flags[key] = True
    merged.checkbox_flags = CheckboxFlags(**base_flags)

    # 合併工廠清單（按 name/address 去重）
    if update.factories:
        existing = {(f.name, f.address) for f in merged.factories}
        for f in update.factories:
            key = (f.name, f.address)
            if key not in existing:
                merged.factories.append(f)
                existing.add(key)

    # 合併附件
    if update.attachments:
        if merged.attachments:
            merged.attachments = list(dict.fromkeys(merged.attachments + update.attachments))
        else:
            merged.attachments = update.attachments

    return merged
