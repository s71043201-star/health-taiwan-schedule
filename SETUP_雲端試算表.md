# 設定「共用 Google Sheet → 自動更新網站」

完成後，任何人編輯共用試算表並存檔，網站會在數秒～1 分鐘內自動更新，
**完全不需要碰 GitHub、也不需要任何人的電腦開著**。

整體流程：
```
共用 Google Sheet（大家可編輯）
   └─(一變動)→ Apps Script → 呼叫 GitHub repository_dispatch
                  └→ GitHub Actions 抓試算表 → 跑 generate.py → 重建 → Pages 更新
```

---

## 一次性設定（約 10 分鐘，只需做一次）

### ① 建立共用試算表
1. 到 Google 試算表新建一份，把下載到的課表資料**整份貼上**。
2. **第一列必須是這 9 個欄位標題**（順序需一致）：

   ```
   課程名稱 | 處方類型 | 日期 | 時段 | 上課地點 | 已報名 | 人數上限 | 是否額滿 | 狀態
   ```
   - 日期格式：`2026-06-01` 或 `2026/6/1` 皆可。
   - 時段格式：`09:00 - 09:30`。
   - 狀態為「取消」者會自動排除。
3. 右上「共用」→ 一般存取權改成 **知道連結的任何人 ▸ 檢視者**。
4. 複製網址中的試算表 ID（`/d/` 與 `/edit` 之間那段）。

### ② 在 GitHub 設定試算表來源
1. 進 repo ▸ **Settings ▸ Secrets and variables ▸ Actions ▸ Variables ▸ New repository variable**。
2. 名稱填 `SHEET_XLSX_URL`，值填（把 `<試算表ID>` 換掉）：

   ```
   https://docs.google.com/spreadsheets/d/<試算表ID>/export?format=xlsx
   ```

   到這裡，定時／手動觸發就已會抓試算表。下面再加「即時觸發」。

### ③ 建立 GitHub 權杖（給 Apps Script 用）
1. GitHub ▸ 右上頭像 ▸ Settings ▸ Developer settings ▸
   **Personal access tokens ▸ Tokens (classic) ▸ Generate new token (classic)**。
2. 勾選 **`repo`** 權限，產生後**複製權杖字串**（只會顯示一次）。

### ④ 在試算表掛 Apps Script（即時觸發）
1. 在試算表 ▸ 上方選單 **擴充功能 ▸ Apps Script**。
2. 把 `apps_script_即時更新.gs` 的內容整個貼進去（`GITHUB_REPO` 已是你的 repo）。
3. 左側 **專案設定（齒輪）▸ 指令碼屬性 ▸ 新增屬性**：
   - 名稱 `GH_TOKEN`，值貼上③的權杖，儲存。
4. 左側 **觸發條件（時鐘圖示）▸ 新增觸發條件**：
   - 函式：`onChangeRebuild`
   - 事件來源：**試算表**
   - 事件類型：**變更時（On change）**
   - 儲存（會要求授權，按同意）。
5. 可先選 `testDispatch` 按「執行」測試 → 到 repo 的 **Actions** 頁應看到「重建課表網站」被觸發。

---

## 之後的日常使用
- 想更新課表 → **打開共用試算表，貼上／修改資料，存檔**。
- 等數秒～1 分鐘，https://s71043201-star.github.io/health-taiwan-schedule/ 就更新。
- 要找人協作 → 直接把試算表分享給對方可編輯即可（不必給 GitHub 權限）。

## 備援與排錯
- 即使 Apps Script 沒觸發，Actions 每天清晨也會定時重建一次。
- 想手動重建：repo ▸ Actions ▸「重建課表網站」▸ Run workflow。
- 沒更新時先看 Actions 頁那次執行的紀錄訊息。
