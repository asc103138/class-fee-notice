# 四上班費收費通知單與套印系統 - AGENTS.md

## 專案資訊
- 專案名稱：四上班費收費通知單與套印系統
- 專案用途：國小四年級上學期班費收費通知單排版、全班 28 人批次套印（A4 橫式一頁 6 張、直式 4 合 1）、收費名冊核對與學生姓名自動帶入工具。
- 主要工作目錄：`/Users/tunyuan/anti/班費`
- GitHub Repo：`https://github.com/asc103138/class-fee-notice` (Public)
- 線上網站 (Pages)：`https://asc103138.github.io/class-fee-notice/`
- 關鍵規格：
  - A4 橫式一頁 6 張（2 列 × 3 欄），全班 28 人僅需 5 頁（含 2 張空白備用單），裁切線清晰，零跑版防跨頁設定。
  - 收費合計：300 元
  - 收費明細：運動會道具 50 元、影印及雜支 112 元、學用品 138 元（寒假作業 28 元 + 國語隨堂練習 50 元 + 自然練習簿 60 元）。

## Obsidian 關聯筆記
- Vault 路徑：`/Users/tunyuan/opencode_0715`
- 專案駕駛艙：`/Users/tunyuan/opencode_0715/04-專案/四上班費收費通知-專案駕駛艙.md`

## 工作與安全規則
- 回應使用繁體中文（台灣）。
- 開工時讀本檔、讀 Obsidian 駕駛艙、檢查 Git 狀態。
- 收工時更新 Obsidian，檢查 diff 後只提交本次相關檔案。
- 絕不 commit API Key、Token、密碼或個人隱私資料。
- 學生資料僅記錄班級代號與座號，不公開上傳學生真實姓名。
- Word 排版必須維持 `snapToGrid="0"` 與嚴格列高控制，防止在 Mac 平台或 Pages 開啟時次列被擠壓跨頁。
