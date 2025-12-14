# CB to CNS Report Generator

將 CB Test Report PDF 自動轉換為 CNS Report Word 文件。

## 功能概述

- **PDF 解析**：使用 Adobe PDF Extract API 萃取 CB Report PDF 的文字與表格
- **智慧萃取**：使用 Azure OpenAI 將 PDF 內容轉換為結構化 JSON Schema
- **自動翻譯**：自動將英文內容翻譯為繁體中文
- **Word 填寫**：將萃取的資料填入 CNS 報告 Word 模板（不破壞原有格式）

## 專案結構

```
word-translation2/
├── backend/
│   ├── main.py                 # FastAPI 主程式
│   ├── config.py               # 設定檔（讀取 .env）
│   ├── schemas/
│   │   └── report_schema.py    # JSON Schema 定義
│   ├── services/
│   │   ├── adobe_extract.py    # Adobe PDF Extract 服務
│   │   ├── azure_llm.py        # Azure OpenAI 服務
│   │   └── word_filler.py      # Word 模板填寫服務
│   ├── utils/
│   │   └── logger.py           # Logging 設定
│   └── requirements.txt        # Python 依賴
├── templates/                  # Word 模板資料夾
│   └── (放置您的 CNS 報告模板 .docx)
├── .env.example               # 環境變數範本
├── Dockerfile                 # Docker 設定（Zeabur 部署用）
└── README.md                  # 本文件
```

## 快速開始

### 1. 安裝依賴

```bash
cd backend
pip install -r requirements.txt
```

### 2. 設定環境變數

```bash
# 複製範本
cp .env.example .env

# 編輯 .env 填入實際值
```

需要設定的變數：

| 變數名稱 | 說明 |
|---------|------|
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI 端點 URL |
| `AZURE_OPENAI_API_KEY` | Azure OpenAI API Key |
| `AZURE_OPENAI_DEPLOYMENT` | 部署名稱（如 gpt-4o） |
| `ADOBE_CLIENT_ID` | Adobe PDF Services Client ID |
| `ADOBE_CLIENT_SECRET` | Adobe PDF Services Client Secret |

### 3. 準備 Word 模板

在 `templates/` 資料夾中放置您的 CNS 報告 Word 模板（.docx 格式）。

模板中使用 `{{placeholder}}` 格式標記需要填寫的欄位，例如：

- `{{report_no}}` - 報告編號
- `{{applicant_zh}}` - 申請人（中文）
- `{{product_name_zh}}` - 產品名稱（中文）
- `{{series_model_1}}` - 第一個系列型號

詳細欄位清單請參考 `backend/services/word_filler.py`。

### 4. 啟動服務

```bash
cd backend
uvicorn main:app --reload
```

服務將在 http://localhost:8000 啟動。

### 5. 使用方式

1. 開啟瀏覽器訪問 http://localhost:8000
2. 上傳 CB Report PDF 檔案
3. 等待處理完成
4. 自動下載填好的 CNS Report Word 檔案

## API 文件

### `POST /generate-report`

上傳 CB PDF 並產生 CNS Word 報告。

**Request:**
- Content-Type: `multipart/form-data`
- Body:
  - `file`: PDF 檔案（必填）
  - `use_mock`: 是否使用模擬資料（選填，預設 false）

**Response:**
- Content-Type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- 直接回傳 Word 檔案

### `GET /health`

健康檢查。

### `GET /api/template-info`

取得模板資訊。

### `GET /api/schema-sample`

取得 Schema 範例（JSON 格式）。

## Zeabur 部署

### 1. 推送到 Git Repository

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin <your-repo-url>
git push -u origin main
```

### 2. 在 Zeabur 建立專案

1. 登入 [Zeabur](https://zeabur.com)
2. 建立新專案
3. 從 Git Repository 部署

### 3. 設定環境變數

在 Zeabur 專案設定中加入所有必要的環境變數（參考 `.env.example`）。

### 4. 注意事項

- Zeabur 會自動設定 `PORT` 環境變數
- `templates/` 資料夾會被包含在 Docker image 中
- 如果需要動態更新模板，考慮使用外部儲存（如 S3）

## Word 模板 Placeholder 參考

### 基本資料

| Placeholder | 說明 |
|-------------|------|
| `{{report_no}}` | 報告編號 |
| `{{cb_report_no}}` | CB 報告編號 |
| `{{standard}}` | 適用標準 |
| `{{applicant_en}}` | 申請人（英文） |
| `{{applicant_zh}}` | 申請人（中文） |
| `{{manufacturer_en}}` | 製造商（英文） |
| `{{manufacturer_zh}}` | 製造商（中文） |
| `{{product_name_en}}` | 產品名稱（英文） |
| `{{product_name_zh}}` | 產品名稱（中文） |
| `{{model_main}}` | 主型號 |
| `{{ratings_input}}` | 輸入額定值 |
| `{{ratings_output}}` | 輸出額定值 |

### 系列型號（1-60）

| Placeholder | 說明 |
|-------------|------|
| `{{series_model_N}}` | 第 N 個型號 |
| `{{series_model_N_vout}}` | 輸出電壓 |
| `{{series_model_N_iout}}` | 輸出電流 |
| `{{series_model_N_pout}}` | 輸出功率 |

### Checkbox 處理

模板中使用 `□` 符號表示 checkbox，程式會根據報告內容自動將對應的 `□` 改成 `■`。

支援的 checkbox：
- 產品群組：AV、ICT、Audio/Video & ICT、Telecom
- 使用分類：Ordinary、Skilled、Instructed
- 電源等級：Class I、Class II、Class III

## 開發與測試

### 使用模擬資料測試

在上傳頁面勾選「使用模擬資料」可以跳過 Adobe/Azure API 呼叫，使用內建的假資料測試 Word 填寫功能。

### 查看未替換的 Placeholder

```python
from services.word_filler import find_unreplaced_placeholders

unreplaced = find_unreplaced_placeholders("output.docx")
print(unreplaced)
```

## 常見問題

### Q: Word 模板中的 placeholder 沒有被替換？

A: Word 可能將 `{{name}}` 切成多個 run（例如 `{{`、`name`、`}}`）。程式已處理這種情況，但如果仍有問題：
1. 在 Word 中選取整個 placeholder
2. 剪下後直接貼上（清除格式）
3. 或使用「清除格式」功能

### Q: Adobe API 回傳錯誤？

A: 請確認：
1. `ADOBE_CLIENT_ID` 和 `ADOBE_CLIENT_SECRET` 正確
2. Adobe 帳號有足夠的 API 配額
3. PDF 檔案格式正確且未加密

### Q: Azure OpenAI 回傳錯誤？

A: 請確認：
1. `AZURE_OPENAI_ENDPOINT` 格式正確（結尾有 `/`）
2. `AZURE_OPENAI_DEPLOYMENT` 是實際存在的部署名稱
3. API Key 有效且有權限存取該部署

## 授權

MIT License
