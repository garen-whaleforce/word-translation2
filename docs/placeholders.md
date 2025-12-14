# AST-B 報告 Placeholder 對照表

此文件記錄 AST-B CNS 報告模板中可回填欄位與 CB 報告來源的對應關係。

## CB 來源可對應的欄位

| placeholder | AST 定位 | AST 原文字 | CB 來源關鍵字/行 | type |
|-------------|----------|-----------|-----------------|------|
| `report_number` | table:0, row:3, col:0 | 報告編號……………..………: | Report Number. ..............................: 2025112058855971-00 | string |
| `applicant_name` | table:0, row:4, col:0 | 申請者名稱……………..……: | Applicant's name (參見 CB 報告首頁) | string |
| `applicant_address` | table:0, row:5, col:0 | 地址………………..…………: | Address (參見 CB 報告首頁 Applicant 區塊) | string |
| `factory_name` | table:0, row:6, col:0 | 生產廠場……………..……: | Name and address of factory (ies): 1) Dongguan Aohai Technology Co.,Ltd. | string |
| `factory_address` | table:0, row:7, col:0 | 地址………………..…………: | (工廠地址，參見 CB 報告 factory 區塊) | string |
| `test_standard` | table:0, row:8, col:0 | 試驗標準(規範)……………: | Standard: IEC 62368-1:2018 | string |
| `product_name` | table:0, row:10, col:0 | 品名…………..…………………: | Test item description: AC POWER SUPPLY | string |
| `main_model` | table:0, row:11, col:0 | 主型號……………..……………: | Model/Type reference: MC-601, A1231-200300C-US1 | string |
| `series_models` | table:0, row:12, col:0 | 系列型號…………..……………: | Model/Type reference: MC-601, A1231-200300C-US1 (系列部分) | array |
| `trademark` | table:0, row:13, col:0 | 廠牌/商標………………….……: | Trade Mark(s): (參見 CB 報告) | string |
| `ratings_input` | table:0, row:14, col:2 | 輸入: 100-240 V～, 50/60 Hz, 1.7 A | Input: 100-240V~, 50/60Hz, 1.7A | string |
| `ratings_output` | table:0, row:14, col:2 | 輸出: 5.0 V 3.0 A 15.0 W or 9.0 V 3.0 A 27.0 W or... | Output: 5.0V 3.0A 15.0W or 9.0V 3.0A 27.0W or 15.0V 3.0A 45.0W or 20.0V 3.0A 60.0W or 5.0-20.0V... | string |
| `sample_receipt_date` | table:0, row:18, col:0 | 試驗件收件日……..……………: | Date of receipt of test item: 2025-11-07 | string |
| `test_date` | table:0, row:19, col:0 | 執行測試日………………...…...: | (參見 CB 報告測試日期) | string |
| `report_issue_date` | table:0, row:20, col:0 | 報告發行日…………..…………: | Date of issue: 2025-12-02 | string |
| `testing_lab` | table:0, row:21, col:0 | 測試單位…………..………: | Name of Testing Laboratory: Keyway Testing Technology (Guangdong) Co., Ltd. | string |
| `testing_lab_address` | table:0, row:22, col:0 | 地址…………………………: | Testing location/address: 21/F., Building 6, Dongyi Intelligent Equipment New Energy Vehicle Park... | string |
| `tma` | table:3, row:12, col:1 | 45 °C 室外:最低 °C | Manufacturer's specified Tma: 45°C | string |
| `ip_rating` | table:3, row:13, col:1 | IPX0 IP___ | IP protection class: IPX0 | string |
| `altitude` | table:3, row:15, col:0 | 設備適用的海拔高度 | Altitude during operation (m): 2000 m or less / 5000 m | string |
| `test_lab_altitude` | table:3, row:16, col:0 | 測試實驗室海拔高度 | Altitude of test laboratory (m): 2000 m or less | string |
| `equipment_mass` | table:3, row:17, col:0 | 設備質量(kg) | Mass of equipment (kg): Approx. 0.072kg. | string |
| `equipment_class` | table:3, row:8, col:1 | OVC I OVC II OVC III OVC IV 其他: | Class of equipment: Class I / PD 3 | string |
| `supply_connection_type` | table:3, row:5, col:1 | A 型插接式設備 不可分離式電源線 分離式電源線 直插式設備... | Supply connection – type: pluggable equipment type A - non-detachable supply cord... direct plug-in | string |
| `equipment_mobility` | table:3, row:7, col:1 | 移動式設備 手持式設備 可攜式設備 直插式設備 放置式設備... | Equipment mobility: movable | string |
| `cb_certificate_number` | table:3, row:18 (備註) | 參考證書號碼為DK-174052-UL | (需從 CB 證書取得) | string |
| `cb_report_number` | table:3, row:18 (備註) | 報告號碼為2025112058855971-00 | Report Number: 2025112058855971-00 | string |
| `limiting_component_voltage` | table:7, row:11, col:2 | 限制元件：U2在 T1 腳位 5至6之後，電壓為 46.8Vpk (ES1) | Limiting component: U2 After T1 pin 5 to 6, voltage is 46.8Vpk (ES1) (for output 20.0Vdc, 3.0A) | string |
| `insulation_resistance` | table:7, row:72, col:2 | AC 插頭與輸出端子之間：>500 M | Insulation resistance (MΩ): Between AC plug and output terminals: >500 MΩ. | string |
| `multiplication_factor` | table:7, row:46, col:1 | 超過海拔2,000 m之乘數因子 | Multiplication factors for clearances and test voltages: 1.48 | number |

## 需人工填寫的欄位 (needs_manual)

以下欄位在 CB 報告中找不到直接對應，需人工填寫或從其他來源取得：

| placeholder | AST 定位 | AST 原文字 | 說明 |
|-------------|----------|-----------|------|
| `bsmi_report_number` | table:0, row:2, col:0 | 標準檢驗局試驗報告指定編號: | 標檢局專用編號，非 CB 來源 |
| `bsmi_lab_code` | table:0, row:0 header | 標準檢驗局指定試驗室認可編號: SL2-IN/VA-T-0157 | 台灣實驗室認可編號 |
| `test_method` | table:0, row:9, col:0 | 試驗方式…………….……: | 需依實際情況填寫 |
| `test_result` | table:0, row:23, col:0 | 試驗結果………..…………: | 需依實際測試結果填寫 |
| `test_pass_items` | table:0, row:16, col:0 | 測試樣品符合要求……………..: | 需彙整測試結果 |
| `test_fail_items` | table:0, row:17, col:0 | 測試樣品不符合要求.………….: | 需彙整測試結果 |
| `test_na_items` | table:0, row:15, col:0 | 測試項目不適用……………..…: | 需彙整不適用項目 |
| `report_author` | table:0, row:25, col:0 | 報告製作者: | 需人工填寫 |
| `report_signer` | table:0, row:25, col:3 | 報告簽署人: | 需人工填寫 |
| `revision_history` | table:2, row:0, col:0 | 報告修訂紀錄: | 新報告通常為空 |
| `outdoor_min_temp` | table:3, row:12, col:1 | 室外:最低 °C | CB 報告僅提供 Tma 45°C，無室外最低溫 |
| `special_installation` | table:3, row:10, col:1 | 無特殊安裝 限制存取區 戶外位置 其他: | 需依產品確認 |
| `protection_device_rating` | table:3, row:6, col:1 | 20 A建築 位置: 建築 設備內 無保護裝置 | 需依產品確認 |

## 欄位類型說明

- **string**: 單一文字值
- **number**: 數值
- **array**: 多值陣列（如系列型號）

## 備註

1. CB 報告中 `Manufacturer: Same as applicant` 表示製造商與申請者相同
2. 額定值 (Ratings) 包含輸入與輸出兩部分，建議分開處理
3. 部分欄位可能需要格式轉換（如日期格式、電壓單位等）
4. 工廠資訊在 CB 報告中標示為 `Dongguan Aohai Technology Co.,Ltd.`
5. 海拔高度有兩種選項：2000m 以下 或 5000m，此產品適用 5000m（報告中提到使用 1.48 倍數因子）
