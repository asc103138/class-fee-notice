# 四上班費收費通知單與套印系統 - AGENTS.md

## 專案資訊
- **專案名稱**：四上班費收費通知單與套印系統
- **專案用途**：國小四年級上學期班費收費通知單排版、全班 28 人批次套印（A4 橫式一頁 6 張、直式 4 合 1）、收費名冊核對與學生姓名自動帶入工具。
- **跨平台支援目錄**：
  - **Windows 11**：`D:\antui\班費`
  - **macOS**：`/Users/tunyuan/anti/115四上班費`（或 `~/anti/115四上班費`）
- **GitHub Repo**：`https://github.com/asc103138/class-fee-notice` (Public)
- **線上展示 (Pages)**：`https://asc103138.github.io/class-fee-notice/`
- **關鍵規格**：
  - A4 橫式一頁 6 張（2 列 × 3 欄），全班 28 人僅需 5 頁（含 2 張空白備用單），裁切線清晰，零跑版防跨頁設定。
  - 收費合計：300 元
  - 收費明細：運動會道具 50 元、影印及雜支 112 元、學用品 138 元（寒假作業 28 元 + 國語隨堂練習 50 元 + 自然練習簿 60 元）。

## 全域技能綁定與使用守則
本專案已連結並強制要求後續所有 Agent／任務**第一時間遵循以下全域技能**，禁止重造輪子或盲目試錯：

1. **`antigravity-advanced-docs`（進階文件處理 Skill）**：
   - 凡涉及 Word (`python-docx`)、PDF (`fitz` / `pypdf` / `reportlab`)、Excel (`openpyxl` / `pandas`) 處理任務，**一律先載入此技能架構**。
   - 遵循「100% 本機端運算」與零個資外洩守則。
2. **`antigravity-workflow`（標準工作流程 Skill）**：
   - 開工先讀本檔、`handoff.md`、檢查 `git status`。
   - 收工檢查敏感資料、更新專案記錄、僅 stage 本次相關變更，絕不未經確認自動 push。

## macOS ↔ Windows 11 跨平台適性化架構規範
為消除跨作業系統切換時的破版、編碼崩潰與 Git 髒污問題，落實以下架構防線：

1. **Git 換行規範（Line Endings Normalization）**：
   - 專案根目錄已建立 `.gitattributes`，文字檔強制 `eol=lf`（檢入時統一 LF，檢出時自動適配本機行尾）。
   - 禁止在 Windows 與 macOS 間因 CRLF/LF 差異產生整檔 diff。
2. **終端與 Python 編碼防護**：
   - 針對 Windows 11 終端預設 CP950，Python 腳本已內建 `sys.platform == 'win32'` 時強制設定 UTF-8 I/O（`io.TextIOWrapper`），防止學生罕見字或 Emoji 噴出 `UnicodeEncodeError`。
   - macOS 系統預設原生 UTF-8，自動支援雙向讀寫無障礙。
3. **原生 PDF 雙引擎極速轉檔**：
   - **Windows 11**：調用 PowerShell `Word.Application` COM 物件（< 3 秒原生無損）。
   - **macOS**：調用 AppleScript 控制 macOS 原生 Microsoft Word（或自動回退 LibreOffice）。
   - 保證產出之 PDF 具備最高解析度與 100% 原始字型（標楷體）排版。
4. **Word / Pages 排版鎖定與防溢出規範**：
   - 橫式一頁 6 張版型必須嚴格維持表格 `snapToGrid="0"`、`w:cantSplit` 與每列精確固定高度（`w:trHeight w:val="5300"`）。
   - 套印替換文字時，直接修改第二段的 `runs[0].text`，保留原樣式、字型（標楷體）、粗體與段落間距，防止在 macOS Word 或 Apple Pages 開啟時因字型度量差造成表格被擠出第 2 頁。
5. **雙系統暫存檔隔離**：
   - `.gitignore` 已完整過濾 macOS（`.DS_Store`、`._*`、`.AppleDouble`）與 Windows（`Thumbs.db`、`Desktop.ini`）系統雜檔及 Office 暫存檔（`~$*.docx`）。

## 安全與隱私紅線
- 回應使用繁體中文（台灣）。
- 學生資料僅記錄班級代號與座號至公開 Repo。含真實學生姓名之檔案（`班級名條*.xlsx`、`*四年丙班*`、`*含姓名*`）已受 `.gitignore` 保護，**絕不 commit 或 push 至公開 GitHub**。
- 不自動執行 `git pull`、`git commit` 或 `git push`，需由使用者明確指示。
