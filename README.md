# bank-reconciliation 銀行匯款對帳

**把官網「匯款訂單」CSV 整理進對帳總表，回填客戶代號，並可推送到線上 Google 試算表（GAS）。**

## 功能流程

| 步驟 | 說明 |
|------|------|
| **Step 1** | 讀取 CSV（自動判斷 cp950 / big5 / utf-8），刪除全空白列 |
| **Step 2** | 逐列對應：交易日、存款金額、摘要、備註、部門、網站訂單編號、賣家備註 |
| **Step 3** | 摘要智慧修正：賣家備註若寫「正確後5碼」，用正確碼覆蓋轉帳後五碼 |
| **Step 4** | 接續寫入既有總表最後一列之後 |
| **Step 5** | 客戶代號回填：以「客戶訂單｜客戶代號」對照檔比對網站訂單編號 |
| **Step 6** | 推送線上：POST A~G 欄到 GAS 網頁應用程式，以網站訂單編號去重、重複跳過 |
| **Step 7** | 格式化：交易日置中＋日期、存款金額靠右＋千分位、摘要文字格式（保留前導零） |

## 安裝依賴

```bash
pip install openpyxl
```

## 使用方式

```bash
# 匯款資料寫進總表
python bank_recon.py <匯款訂單.csv> [--total 對帳總表.xlsx] [--dept 電圖部]

# 回填客戶代號
python bank_recon.py <匯款訂單.csv> --lookup "☺客戶訂單+代號_銷貨明細 -_YYYYMMDD.xlsx"

# 再推到線上 Google 試算表（A~G）
python bank_recon.py <匯款訂單.csv> --lookup "..." --push

# 把總表全部資料列推上去
python bank_recon.py --push-all

# 只試算不寫入線上（安全驗證）
python bank_recon.py --push-all --dry-run
```

### 參數

| 參數 | 說明 |
|------|------|
| `匯款訂單.csv` | 官網匯出的匯款訂單 CSV（可省略） |
| `--total` | 對帳總表檔名（預設 `銀行匯款對帳資料-總表.xlsx`），不存在會自動建立 |
| `--dept` | 部門固定值（預設 `電圖部`） |
| `--lookup` | 「客戶訂單｜客戶代號」兩欄 xlsx |
| `--overwrite` | 清空總表既有資料列後重寫（保留表頭） |
| `--push` | 推送**本次寫入**的資料列到線上（A~G） |
| `--push-all` | 推送總表**全部資料列**到線上（A~G） |
| `--webapp` | 指定 GAS 網址；沒給就讀 `bank_recon.config.json` |
| `--dry-run` | 只試算不寫入線上 |

### 本機設定檔

工作目錄下的 `bank_recon.config.json`（**含 token，請勿進版控**）：

```json
{
  "webapp": "https://script.google.com/macros/s/<deploymentId>/exec",
  "token": "<與 GAS 一致的權杖>"
}
```

### 範例輸出

```bash
python bank_recon.py "E-0922.csv" --lookup "☺客戶訂單+代號_銷貨明細 -_20260926.xlsx" --push
```

```
來源: E-0922.csv（編碼 cp950）
寫入 20 筆 → 銀行匯款對帳資料-總表.xlsx（第 3~22 列）
摘要以「正確後5碼」覆蓋: 2 筆
客戶代號比對: 查找表 1205 筆（衝突 0）→ 帶入 21/21 列
線上推送: 送出 21 筆 → 新增 2 筆、跳過(重複) 19 筆  [分頁 115(26)]
```

## 總表欄位

| 欄 | 名稱 | 推送線上 |
|----|------|:---:|
| A | 交易日 | ✅ |
| B | 存款金額 | ✅ |
| C | 摘要 | ✅ |
| D | 備註 | ✅ |
| E | 部門 | ✅ |
| F | 網站訂單編號 | ✅ |
| G | 客戶代號 | ✅ |
| H | 賣家備註 | ❌ 不推 |

## GAS 網頁應用程式

- 型別：**獨立腳本**（`SpreadsheetApp.openById(SPREADSHEET_ID)`）
- 部署：執行身分「我」、存取權「任何人」
- `doPost` 收到 `{token, rows, dryRun?}`：驗權杖 → 找表頭列與「網站訂單編號」欄 → 讀既有鍵 → 去重後 `setValues` 寫入 A~G → 回傳 `{ok, added, skipped, sheet}`
- `doGet`：回傳目前分頁、表頭列、欄位位置，方便確認

> ⚠️ 綁定試算表的 GAS 專案無法用 clasp API 建立網頁應用程式部署（會回 `Only users in the same domain as the script owner may deploy this script`），請用**獨立腳本**。
>
> ⚠️ Web App 網址等於可以寫入該試算表，請放在 `bank_recon.config.json` 並排除版控。

## 注意

- 來源 CSV **不要用 Excel 開啟後另存**，12 位數單號會被轉成科學記號而遺失末碼。
- 摘要一律存成**文字格式**，避免 `00780` 這類前導零被吃掉。
- `--lookup` 重複執行不會重複插欄；`--push-all` 重複執行第二次會全部跳過（去重）。
