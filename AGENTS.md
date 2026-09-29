# 四上班費收費通知單與套印系統 - AGENTS.md

## 專案資訊
- **專案名稱**：四上班費收費通知單與套印系統
- **專案用途**：國小四年級上學期班費收費通知單排版、全班 28 人批次套印（A4 橫式一頁 6 張、直式 4 合 1）、收費名冊核對與學生姓名自動帶入工具。
- **主要工作目錄**：`D:\antui\班費`（Windows 環境）
- **GitHub Repo**：`https://github.com/asc103138/class-fee-notice` (Public)
- **線上展示 (Pages)**：`https://asc103138.github.io/class-fee-notice/`
- **關鍵規格**：
  - A4 橫式一頁 6 張（2 列 × 3 欄），全班 28 人僅需 5 頁（含 2 張空白備用單），裁切線清晰，零跑版防跨頁設定。
  - 收費合計：300 元
  - 收費明細：運動會道具 50 元、影印及雜支 112 元、學用品 138 元（寒假作業 28 元 + 國語隨堂練習 50 元 + 自然練習簿 60 元）。

## 全域技能綁定與使用守則（交代事項）
本專案已連結並強制要求後續所有 Agent／任務**第一時間遵循以下全域技能**，禁止重造輪子或盲目試錯：

1. **`antigravity-advanced-docs`（進階文件處理 Skill）**：
   - 凡涉及 Word (`python-docx`)、PDF (`fitz` / `pypdf` / `reportlab`)、Excel (`openpyxl` / `pandas`) 處理任務，**一律先載入此技能架構**。
   - 遵循「100% 本機端運算」與零個資外洩守則。
2. **`antigravity-workflow`（標準工作流程 Skill）**：
   - 開工先讀本檔、`handoff.md`、檢查 `git status`。
   - 收工檢查敏感資料、更新專案記錄、僅 stage 本次相關變更，絕不未經確認自動 push。

## 技術實現與 Windows 環境標準 SOP
為避免重複踩坑與執行拖延，所有後續操作務必遵照以下標準流程：

1. **強制 UTF-8 編碼防線（防 CP950 崩潰）**：
   - 在 Windows PowerShell 下執行 Python 腳本或處理中文字元（如學生罕見字、Emoji），**必須**在腳本頂部強制配置：
     ```python
     import sys, io
     sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
     sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
     ```
   - 避免直接在終端 inline 執行複雜引號字串，一律寫成獨立 `.py` 檔呼叫。
2. **PDF 匯出最佳實踐（Word COM 原生轉換）**：
   - 本機已具備 Microsoft Word 16.0 引擎。
   - Word 轉 PDF **直接調用 PowerShell Word COM 物件**，100% 呈現原生排版，轉換時間 < 3 秒，禁止耗時尋找其他轉換器：
     ```powershell
     $word = New-Object -ComObject Word.Application; $word.Visible = $false
     $doc = $word.Documents.Open($docPath)
     $doc.SaveAs([ref]$pdfPath, [ref]17) # 17 = wdFormatPDF
     $doc.Close([ref]0); $word.Quit()
     ```
3. **排版防跨頁與儲存格保護**：
   - 橫式一頁 6 張版型必須維持表格 `snapToGrid="0"`、`w:cantSplit` 與每列精確固定高度（`w:trHeight w:val="5300"`）。
   - 套印替換文字時，直接修改第二段的 `runs[0].text`，保留原樣式、字型（標楷體）、粗體與段落間距，禁止重建表格導致破版。
4. **名冊讀取優先級**：
   - 優先讀取：`班級名條.xlsx`（包含完整座號、學號、性別、姓名）。
   - 次要回退：`學生名冊與收費核對表_四上班費.xlsx`。

## 安全與隱私紅線
- 回應使用繁體中文（台灣）。
- 學生資料僅記錄班級代號與座號至公開 Repo。含真實學生姓名之檔案（`班級名條*.xlsx`、`*四年丙班*`、`*含姓名*`）已受 `.gitignore` 保護，**絕不 commit 或 push 至公開 GitHub**。
- 不自動執行 `git pull`、`git commit` 或 `git push`，需由使用者明確指示。
