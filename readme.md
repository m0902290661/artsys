# 弘明實驗高級中學 - 美術歷程檔案管理系統 (v4.1)

本系統為 **弘明實驗高級中學 115 級美術歷程** 所開發的線上作業繳交、審核與專案管理系統。前端採用純 HTML/JS/CSS 打造響應式介面（RWD），後端基於 **Google Apps Script (GAS)** 結合 **Google Sheets** 與 **Google Drive**，提供無伺服器（Serverless）架構的完整作業管理與權限控管方案。

---

## 📁 系統架構與檔案說明

專案由前端靜態網頁、API 設定檔以及 Google Apps Script 後端程式碼組成：

```text
├── index.html       # 學生端首頁 (作業項目選單、PDF 上傳/補繳、繳交撤銷)
├── login.html       # 系統登入與學生帳號開通頁面
├── manage.html      # 管理員後台 (時程設定、繳交/缺交追蹤、退件審核、批量匯出 QR Code、帳號管理)
├── question.html    # 問題通報頁面 (密碼重設、無法登入、上傳問題回報)
├── api.json         # 前端 API 設定檔 (記錄 Google Apps Script 部署 URL)
└── 程式碼.js         # 後端核心邏輯 (GAS 伺服器端腳本)
```

---

## 🌟 核心功能特點

### 1. 學生端 (Student Portal)
- **身分驗證與 Session 控管**：採用 8 小時時效性 Session Token，保障連線安全。
- **作業繳交與狀態即時同步**：
  - 上傳 PDF 檔案（限制 20MB 以內，具備寫入進度模擬）。
  - 自動偵測是否「逾時遲交」，逾時未獲授權將自動阻斷上傳。
- **即時退件通知與說明**：檢視管理員給予的退件評語與原因。
- **自主撤銷重傳**：截止時間前可自由撤銷舊檔並重新提交。
- **問題回報功能**：可線上填報密碼重設、登入障礙或系統異常。

### 2. 管理員後台 (Admin Console)
- **時程項目管理**：
  - 新增/編輯/刪除美術作業項目與截止日期。
  - 支援附加範本檔案上傳（儲存於 Google Drive 獨立目錄）。
- **繳交與缺交專案管理**：
  - 多維度篩選：依專案項目、班級、社團、繳交狀態（有效/補繳/退件/缺交）進行精準檢索。
  - 統計卡片：即時計算應繳人數、已繳人數、缺交人數與繳交率（%）。
  - 單筆/批量退件：可附帶私人訊息通知學生。
  - 手動/快速解鎖：針對個別學生授予逾時「補繳權限」。
- **批量列印與 QR Code 匯出**：
  - 自動產生包含學生資訊、檔案名稱與雲端連結 QR Code 的卡片版面。
  - 利用 `jsPDF` 與 `html2canvas` 支援批量導出標準 **A4 PDF 文件**。
- **帳號管理與批量匯入**：
  - 支援 CSV 檔案批量導入學生名冊（欄位包含學號、姓名、班級、社團、預設密碼）。
- **問題單據管理**：處理並更新學生通報的各類系統單據狀態（待處理 / 處理中 / 已完成）。

---

## 🔧 後端部署說明 (Google Apps Script)

### 1. Google 雲端資源準備
1. **Google Sheets (資料庫)**：建立一個新的 Google 試算表，將 Sheet ID 填入 `程式碼.js` 的 `SPREADSHEET_ID` 變數中。
2. **Google Drive (檔案庫)**：在 Google 雲端硬碟建立一個主要資料夾，將 Folder ID 填入 `程式碼.js` 的 `MAIN_FOLDER_ID` 變數中。

### 2. 試算表工作表結構 (自動生成)
系統首次執行時會自動建立以下分頁與表頭：
- `users`：`studentId`, `name`, `className`, `club`, `password`, `sessionToken`, `sessionExpiresAt`
- `schedule`：`task`, `deadline`, `note`, `attachmentUrl`, `attachmentName`
- `submissions`：`task`, `studentId`, `studentName`, `className`, `filename`, `url`, `createdAt`, `status`, `revokedAt`, `rejectReason`
- `unlocks`：`studentId`, `task`
- `reports`：`createdAt`, `reportType`, `studentId`, `studentName`, `className`, `contact`, `subject`, `description`, `status`, `handledAt`, `handledBy`
- `options`：`type`, `value` (定義班級與社團下拉選單選項)

### 3. 發布網路應用程式 (Web App)
1. 開啟 Google Apps Script 編輯器，貼上 `程式碼.js`。
2. 點擊 **部署** > **新增部署**。
3. 部署類型選擇 **網路應用程式 (Web App)**：
   - **執行者 (Execute as)**：`我 (Me)`
   - **誰有存取權限 (Who has access)**：`所有人 (Anyone)`
4. 複製部署後取得的 Web App URL。

### 4. 前端 API 配置
將取得的 Web App URL 填入專案根目錄的 `api.json` 檔案中：

```json
{
  "gasUrl": "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec"
}
```

---

## 🛠️ 技術棧 (Tech Stack)

- **前端 (Frontend)**: HTML5, CSS3 (RWD Flexbox/Grid), Modern Vanilla JavaScript (ES6+)
- **第三方套件 (CDNs)**:
  - `QRCode.js` - 前端動態生成 QR Code 碼
  - `jsPDF` - 前端生成與導出 PDF
  - `html2canvas` - 將 HTML 節點渲染為 Canvas 圖像
  - `PapaParse` - CSV 檔案高效解析
- **後端 (Backend)**: Google Apps Script (Node/JS-like), Google Sheets API, Google Drive API