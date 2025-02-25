# 📱 App Store 價格追蹤工具

本工具使用 **Google Apps Script** 來自動追蹤 **App Store 應用程式的價格變動**，並在價格下降時發送 Email 通知。

---

## 🚀 **功能介紹**

### 📄 **產生空白表格**
- 建立 App Tracking (應用程式追蹤) 和 Settings (設定) 兩個工作表。
- 如果 App Tracking 已有數據，則會自動備份為 Backup_YYYYMMDDHHMM 表格。
- 詢問並設定 預設的 Email 通知地址。
- 自動 刪除其他不相關的工作表，保持工作表整潔。
- 自動儲存試算表 ID，確保 checkAppPrices() 可以獨立運行。

### 🔍 **查詢 App ID**
- 允許輸入 **App 名稱 或 App ID** 來查詢。
- 若查詢結果包含 **多個應用程式**，會讓使用者選擇正確的 App。
- 選擇後，自動將 **App ID、名稱、價格、新增到 `App Tracking` 表格**。

### 🔄 **更新 App 價格並發信**
- 每天自動從 **App Store API** 取得 **最新價格**。
- 若價格下降，會自動 **發送 Email 通知**。
- **使用 Settings 工作表中的試算表 ID**，確保即使 Google Sheets 未開啟，也能正常運行。
### ✉️ **設定通知 Email**
- 允許手動修改 Email，並存入 `Settings` 工作表。

### ⚙️ **自動化功能**
- 可設定 **每天自動更新價格**，並在價格下降時發送通知。
- **確保函式可獨立執行，不依賴試算表是否開啟**。
---

## 📋 **使用方式**

### 1️⃣ **新增 Google Apps Script**
1. **開啟 Google Sheets**
2. 點選 **`擴充功能` → `Apps Script`**
3. **刪除原始內容**，貼上完整程式碼。
4. 點選 **「執行」**，授權腳本存取 Google Sheets。

### 2️⃣ **執行「產生空白表格」**
1. 進入 Google Sheets
2. 點選 **`📱 追蹤 App 價格` → `📄 產生空白表格`**
3. **輸入通知 Email**，確保未來能收到降價通知。
4. 如果 App Tracking 已有數據，會自動備份為 Backup_YYYYMMDDHHMM 表格。
5. 系統會自動建立 **`App Tracking`** 和 **`Settings`** 兩個工作表。

### 3️⃣ **查詢並新增 App**
1. 點選 **`📱 追蹤 App 價格` → `🔍 查詢 App ID`**
2. **輸入 App 名稱 或 App ID 及國家代碼** (例如 `XCOM 2, tw` 或 `1458655678, tw`)
3. 如果找到 **多個應用程式**，系統會要求選擇。
4. 選擇後，系統會將 **App ID、名稱、價格** 新增到 `App Tracking`。

### 4️⃣ **更新 App 價格並發送通知**
1. 點選 **`📱 追蹤 App 價格` → `🔄 更新 App 價格並發信`**
2. 系統會檢查最新價格，若 **價格下降**，則發送 Email 通知。
3. 若無變化，系統會顯示 **「今日沒有 App 降價」**。

---

## ⏰ **設定每天自動回報價格**
1. **進入 Apps Script**
2. 點選 **「時鐘 ⏰ 圖示（觸發器）」**
3. **新增觸發器**
4. **函數選擇** `checkAppPrices`
5. **觸發時間選擇** `時間驅動` → `每天` (可選擇早上 8:00 執行)
6. **儲存觸發器設定**

這樣，每天都會自動更新價格，並在降價時發送通知！

