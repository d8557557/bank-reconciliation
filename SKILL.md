---
name: bank-reconciliation
description: 銀行匯款對帳技能（全域可用）。把官網後台匯出的「匯款訂單」CSV（cp950）整理進銀行匯款對帳總表：交易日（匯款日期）、存款金額（匯款金額）、摘要（轉帳後五碼，遇賣家備註寫「正確後5碼」時改用正確碼覆蓋）、備註（匯款銀行）、部門（固定電圖部）、網站訂單編號、賣家備註；可用 --lookup 讀「客戶訂單｜客戶代號」對照檔比對回填客戶代號；並可用 --push 把 A~G 欄推送到線上 Google 試算表（GAS 網頁應用程式），以網站訂單編號去重、重複自動跳過。摘要自動以文字格式保留前導零。說「銀行對帳」「匯款對帳」「整理匯款訂單」「整理銀行資料表」「匯款資料整理」「比對客戶代號」「補客戶代號」「推匯款資料到線上」「上傳到 Google 試算表」時載入。
---

# 銀行匯款對帳 (bank-reconciliation)

## 用途

把官網後台（綠界／藍新金流）匯出的**匯款訂單 CSV**，整理進**銀行匯款對帳總表**（xlsx）；可從「客戶訂單｜客戶代號」對照檔**回填客戶代號**；並可透過 **GAS 網頁應用程式**把 A~G 欄**推送到線上 Google 試算表**（以網站訂單編號去重）。

## 檔案位置

- 主程式：`~/.config/opencode/skills/bank-reconciliation/bank_recon.py`
- 本機設定：工作目錄下的 `bank_recon.config.json`（**含 token，勿進版控**）

```json
{
  "webapp": "https://script.google.com/macros/s/<deploymentId>/exec",
  "token": "<與 GAS 程式碼一致的權杖>"
}
```

## 使用方式

```powershell
# ① 匯款資料寫進本機總表
python ~/.config/opencode/skills/bank-reconciliation/bank_recon.py <匯款訂單.csv> [--total 對帳總表.xlsx] [--dept 電圖部]

# ② 回填客戶代號（可與 ① 同時）
python ~/.config/opencode/skills/bank-reconciliation/bank_recon.py <匯款訂單.csv> --lookup "☺客戶訂單+代號_銷貨明細 -_YYYYMMDD.xlsx"

# ③ 再推到線上 Google 試算表（A~G）
python ~/.config/opencode/skills/bank-reconciliation/bank_recon.py <匯款訂單.csv> --lookup "..." --push

# ④ 把總表全部資料列推上去（本地已累積多筆時）
python ~/.config/opencode/skills/bank-reconciliation/bank_recon.py --push-all

# ⑤ 只試算不寫入線上（安全驗證）
python ~/.config/opencode/skills/bank-reconciliation/bank_recon.py --push-all --dry-run
```

### 參數

| 參數 | 說明 |
|------|------|
| `匯款訂單.csv` | 官網匯出的匯款訂單 CSV（cp950 / Big5）。可省略，省略時只做 `--lookup` / `--push-all` |
| `--total` | 對帳總表檔名（預設 `銀行匯款對帳資料-總表.xlsx`），**不存在會自動建立** |
| `--dept` | 部門固定值（預設 `電圖部`） |
| `--lookup` | 「客戶訂單｜客戶代號」兩欄 xlsx，用來比對回填客戶代號 |
| `--overwrite` | 清空總表既有資料列後重寫（只保留表頭） |
| `--push` | 把**本次寫入**的資料列推到線上（A~G） |
| `--push-all` | 把總表**全部資料列**推到線上（A~G） |
| `--webapp` | 指定 GAS 網頁應用程式網址（沒給就讀 `bank_recon.config.json`） |
| `--dry-run` | 只試算不真的寫入線上（搭配 `--push` / `--push-all`） |

---

## 總表欄位

| 欄 | 名稱 | 說明 |
|----|------|------|
| A | 交易日 | |
| B | 存款金額 | |
| C | 摘要 | |
| D | 備註 | |
| E | 部門 | |
| F | 網站訂單編號 | |
| G | **客戶代號** | 有 `--lookup` 才會出現 |
| H | 賣家備註 | **不推送到線上** |

## 欄位對應規則（定案）

| 總表欄位 | 來源／規則 |
|----------|-----------|
| 交易日 | 官網「匯款日期」（`YYYY-MM-DD`），輸出為日期格式 `yyyy/mm/dd`；未匯款者留空 |
| 存款金額 | 官網「匯款金額」（數字格式 `#,##0`）；未匯款者為 0 |
| 摘要 | 官網「轉帳後五碼」；**若「賣家備註」含「正確後5碼」等字樣，改用該 5 位數字覆蓋**。一律存成**文字格式**（保留 `00780`、`08986` 前導零） |
| 備註 | 官網「匯款銀行」原文（例：`中國信託`、`中華郵政`、`013`、`郵局`） |
| 部門 | 固定值（預設 `電圖部`） |
| 網站訂單編號 | 官網「訂單編號」（E 開頭，例：`E2609080034`） |
| 客戶代號 | `--lookup` 對照檔：以總表「網站訂單編號」對照檔案「客戶訂單」，帶入「客戶代號」 |
| 賣家備註 | 官網「賣家備註」原文（不推送線上） |

> 來源 CSV 沒有的欄位一律**保留空白**。

---

## 處理流程

1. **讀取 CSV**：依序嘗試 `cp950` → `big5` → `utf-8-sig` → `utf-8`，自動判斷編碼
2. **刪除全空白列**
3. **逐列對應欄位**（規則如上表）
4. **摘要修正**：掃描「賣家備註」，若有「正確」字樣則抓出第一個 5 位數字覆蓋摘要
   - 例：`轉帳後五碼=35538`、`賣家備註=正確後5碼06306` → 摘要輸出 `06306`
5. **寫入總表**：讀入既有總表 → 於最後一列之後**接續寫入**（表頭與既有資料都保留）
6. **客戶代號回填**（有 `--lookup`）：
   - 讀對照檔第一張工作表，取 A 欄「客戶訂單」、B 欄「客戶代號」建立去重字典
   - 確保總表有「客戶代號」欄（沒有就插在「網站訂單編號」右邊；**重跑不會重複插欄**）
   - 逐列以「網站訂單編號」查表帶入，印出 `帶入 N/M 列` 與未命中清單
7. **推送線上**（有 `--push` / `--push-all`）：
   - POST JSON 到 GAS 網頁應用程式：`{token, rows, dryRun}`
   - rows 為 A~G 欄（H 賣家備註不推），日期以 `YYYY-MM-DD` 傳送
   - GAS 端**先讀線上「網站訂單編號」整欄建成集合**，逐筆比對；已存在者**跳過**，同批次內也去重
   - 回傳 `{added, skipped, sheet}`，本機印出 `新增 X 筆、跳過 Y 筆`
8. **格式化**：交易日置中＋日期格式、存款金額靠右＋千分位、摘要文字格式＋置中

---

## GAS 端（線上推送）

線上試算表需搭配一支 GAS 網頁應用程式（獨立腳本，用 `openById` 開啟試算表）。

- 部署設定：執行身分 **我**、存取權 **任何人**
- 腳本常數：`SPREADSHEET_ID`、`SHEET_NAME`（空＝第一個分頁）、`TOKEN`、`COLS = 7`
- `doPost`：驗 token → 找表頭列與「網站訂單編號」欄 → 讀既有鍵去重 → `setValues` 批次寫入 A~G
- `doGet`：回傳目前分頁資訊，方便確認設定
- `dryRun: true` 時只計算不寫入，回傳 `added / skipped / dupKeys`

> ⚠️ **綁定試算表的 GAS 專案無法用 clasp API 建立網頁應用程式部署**（會回 `Only users in the same domain as the script owner may deploy this script`）。請改用**獨立腳本**＋`SpreadsheetApp.openById(SPREADSHEET_ID)`。

> ⚠️ 網頁應用程式網址 = 一把鑰匙，等於可以寫入該試算表。請放在 `bank_recon.config.json`（勿進版控），不要寫在程式碼或公開 repo。

---

## 注意事項

1. **來源 CSV 別用 Excel 開啟後另存**：12 位數的單號（如出貨單號）會被 Excel 轉成科學記號 `9.07873E+11`，末幾碼永久遺失。
2. **摘要是文字格式**：即使 `24768` 這種沒有前導零的值也存成文字，避免 Excel 自動轉數字。
3. **總表可累積**：每次執行都接在最後一列之後。
4. **未匯款列**（匯款方式＝未選擇匯款方式、匯款金額 0）：交易日／摘要／備註留空、存款金額 0，需人工確認。
5. **客戶代號對照檔**：第一張工作表須為兩欄（A=客戶訂單、B=客戶代號），從第 2 列起為資料。
6. **推送前去重是必要的**：重複執行 `--push-all` 是安全的（第二次全部跳過），但務必先用 `--dry-run` 確認筆數。

---

## 依賴

- `openpyxl`（處理 Excel）
- `urllib`（標準庫，推送用，無需額外套件）

```powershell
pip install openpyxl
```
