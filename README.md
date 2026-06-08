# 健康台灣深耕計畫 · 課程表網站

把課程時段 Excel 自動分區、分類、統計，產生各區課程表網頁與資料庫。

## 結構

```
課表網站/
├── generate.py        產生器（重跑就用它）
├── courses.json       結構化資料庫（已排除「取消」）
├── index.html         入口頁（選擇月份）
├── data/              來源 Excel
├── templates/         前端版型（沿用 UI測試 設計）
└── 2026-06/           ← 每個月份一個資料夾（依資料自動產生）
    ├── index.html         該月主選單（全區綜合 + 三區）
    ├── 大字版/            各區電腦/大字版 HTML
    └── 手機版/            各區手機版 HTML
```

> 倒入多個月份的資料時，會自動為每個月各建一個 `YYYY-MM/` 資料夾，
> 根目錄 `index.html` 列出所有月份供選擇。標題、月曆、統計都會依該月份自動標記。

## 分類原則

- **行政區**：依「上課地點」文字以關鍵字判斷（北投／士林／中山）。
  - 晨昕診所 → 中山區、圓山捷運站/夢想館 → 中山區。
  - 對應規則集中在 `generate.py` 的 `classify_district()`，要調整改這裡即可。
- **處方類型**：營養／運動／社會／情緒調適，對應固定色票。
- **狀態為「取消」一律排除**，不顯示也不計入統計。
- 一筆 Excel 列 = 一堂課，忠實呈現，不做人工合併。

## 新增 / 更新課表

> **預設：全自動，免手動** — n8n 每天從健康處方系統把課程時段同步到 Supabase，
> GitHub Actions 每 5 分鐘讀 Supabase 重建。**平常完全不用碰**，課程在原系統異動後，
> 數分鐘內網站自動跟上。架構見下方「資料來源」。

程式資料來源優先序（由 `COURSE_SOURCE` 指定，未指定時自動判斷）：
1. **Supabase**（預設）：`COURSE_SOURCE=supabase`，讀 n8n 已同步好的 `course_slots`。
2. **健康處方系統 API**：`COURSE_SOURCE=api` + `COURSE_API_ACCOUNT` / `COURSE_API_PASSWORD`，
   直接登入 API 逐月抓 `slots/summary`。
3. **共用 Google Sheet**：環境變數 `COURSE_XLSX_URL`（試算表 xlsx 匯出網址）。
4. 否則讀 `data/` 內**修改時間最新**的 `.xlsx`（檔名隨意）。

> 共同規則：英文處方類型（`nutrition`/`exercise`/`social`/`mental`）自動轉中文；
> 狀態 `cancelled` 一律排除（等同 Excel 的「取消」）；行政區靠 `classify_district()` 判斷。

### 資料來源（自動更新架構）

```
健康處方系統 API
   ↓ n8n（本機 Windows service，每天 09:30 / 15:00 抓）
Supabase prescription_data(id='main').course_slots
   ↓ GitHub Actions（build.yml，每 5 分鐘 + 變動觸發）讀 Supabase 跑 generate.py
GitHub Pages 自動更新
```

- **Supabase（預設）**：anon 唯讀金鑰是公開金鑰、RLS 保護，已內建於 `generate.py`，
  公開 repo 不放任何帳密。可用 `COURSE_SUPABASE_URL` / `COURSE_SUPABASE_KEY` /
  `COURSE_SUPABASE_ROW` 覆寫。
- **直接打 API（備援）**：`COURSE_SOURCE=api`。帳密**嚴禁寫進程式碼**，放
  repo ▸ Settings ▸ Secrets and variables ▸ Actions ▸ **Secrets**
  （`COURSE_API_ACCOUNT` / `COURSE_API_PASSWORD`）。範圍用 **Variables** 覆寫
  `COURSE_START_DATE`（預設 `2025-12-20`）/ `COURSE_END_DATE`（預設今天 + 6 個月）。
- 本機測試：`COURSE_SOURCE=supabase python generate.py`。

### 方法 A：在 GitHub 網站上傳（推薦，免裝任何東西、不限本機）

任何有此 repo 權限的人都能更新：

1. 打開 repo 的 `data/` 資料夾 →「Add file ▸ Upload files」。
2. 拖入新的課表 Excel，按「Commit changes」。
3. GitHub Actions（`.github/workflows/build.yml`）會自動跑 `generate.py`、
   重建所有網頁並提交，GitHub Pages 約 1～2 分鐘後完成更新。

> 要讓同事也能上傳：repo ▸ Settings ▸ Collaborators ▸ 邀請對方即可。

### 方法 B：本機執行

```
python generate.py
git add -A && git commit -m "更新課表" && git push
```

## 頁面與導覽

- 根 `index.html`：**選擇月份**。
- 各月 `index.html`：該月主選單（全區綜合 + 三區，各有大字版/手機版）。
- 課表頁左下角「← 主選單」回該月選單；月選單左上「← 選擇月份」回根頁。
- **全區綜合**：三區課程合併於一張表，每筆標出所屬區（［北投］等）。
- 多月份時，每月一個 `YYYY-MM/` 資料夾，標題／月曆／統計自動依月份標記。

## 統計（2026-06）

| 區 | 堂數 | 天數 | 營養 | 運動 | 社會 | 情緒 |
|----|----|----|----|----|----|----|
| 北投 | 345 | 27 | 169 | 81 | 33 | 62 |
| 士林 | 80 | 20 | 7 | 21 | 24 | 28 |
| 中山 | 24 | 6 | 8 | 6 | 5 | 5 |
| **合計** | **449** | | | | | |
