# teceadraw 活動摸彩系統

報到頁與投影前台放在 GitHub Pages（只有畫面，不含任何資料），後台就是 Google Sheet，資料由 Sheet 綁定的 GAS 提供 JSON API。

```
checkin/      報到頁（員工手機，兼查詢）
stage/        前台舞台（第 3 階段）
gas/          Sheet 綁定的 Apps Script 程式碼
vendor/       QR code 產生程式庫（qrcode-generator，MIT），本機產生、不呼叫外部服務
tools/        壓力測試與測試名單產生器
deploy-test/  第 1 階段開頭的 5 分鐘部署測試
prototype/    第 0 階段前台視覺原型
```

## 目前進度

| 階段 | 狀態 |
|---|---|
| 0 前台視覺原型 | 完成，待選定拉霸樣式與配色 |
| 1 範本 Sheet＋GAS＋報到頁＋初始化 | 程式完成，待部署測試與壓測 |
| 2 Sheet 後台（控制台、完整選單、資料檢查） | 未開始 |
| 3 前台接真實資料 | 未開始 |
| 4 README、操作小抄、`v1` 標籤 | 未開始 |

## 第 1 階段：5 分鐘部署測試

確認共用雲端硬碟裡的 Sheet 可以用「任何人」部署，手機無痕視窗也打得開。

1. 在共用雲端硬碟建一份空白 Sheet，「擴充功能 › Apps Script」。
2. 把 `deploy-test/Code.gs` 整份貼上、儲存。
3. 「部署 › 新增部署作業」→ 網頁應用程式，執行身分「我」，誰可以存取「任何人」→ 部署、授權。
4. 用手機無痕視窗打開 `https://<Pages 網址>/deploy-test/test.html?e=<部署作業 ID>`，按「測試」，GET 與 POST 都要出現 ✅。

「誰可以存取」選單裡沒有「任何人」時，代表 Workspace 管理員關閉了這個選項，要先請管理員開啟，後面都做不下去。

## 建立範本 Sheet

1. 在共用雲端硬碟新建一份 Sheet，「擴充功能 › Apps Script」。
2. 專案設定：時區選 (GMT+08:00) 台北，勾選「在編輯器中顯示 appsscript.json」，把 `gas/appsscript.json` 內容貼上。
3. 依檔名建立 `gas/` 底下每一個檔案（`.gs` 新增「指令碼」、`.html` 新增「HTML」，名稱不含副檔名），內容整份貼上。
4. 在編輯器選 `setupTemplate` 執行一次，完成授權；回到 Sheet 重新整理，會看到所有分頁與「摸彩系統」選單。
5. 「摸彩系統 › 活動前 › 初始化」，照對話框步驟部署並貼回部署作業 ID，等檢查通過。
6. 「摸彩系統 › 活動前 › 產生連結」取得報到連結、QR code 與前台專屬連結。

程式更新時用「管理部署作業 › 編輯 › 版本：新版本」，部署作業 ID 不變，已印出的 QR code 不用重印。

## 壓力測試

需要 Node 18 以上。

```bash
node tools/make-test-people.mjs 600 > test-people.csv
```

1. 把 `test-people.csv` 的內容（不含標題列）貼到「人員」分頁 A2。
2. 「摸彩系統 › 活動當天 › 開始報到」。
3. 執行：

```bash
node tools/stress-test.mjs --e <部署作業ID> --csv test-people.csv --n 300 --window 30
```

驗收標準：30 秒內送出 300 人，全部在 60 秒內成功，「回報已報到」為 0。逐筆耗時寫在 `stress-result.csv`。跑完用「進階 › 清除資料」回到準備中。

## 抽獎原理

（第 3 階段補上白話說明）
