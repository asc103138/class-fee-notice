# 專案交接紀錄 - handoff.md

- **交接日期**：2026-09-29
- **工作目錄**：
  - Windows 11：`D:\antui\班費`
  - macOS：`/Users/tunyuan/anti/115四上班費`（或 `~/anti/115四上班費`）
- **負責班級**：四年丙班（全班共 28 人）
- **關聯 GitHub**：`https://github.com/asc103138/class-fee-notice`

---

## 🎯 本階段跨平台適性優化成果
1. **跨平台換行與 Git 髒污解決**：
   - 建立 `.gitattributes`，文字檔強制 LF 歸一化，消除 macOS 與 Windows 11 間 CRLF/LF 轉換警告與無效 Git diff。
2. **macOS ↔ Windows 11 雙向 PDF 轉檔引擎**：
   - 升級 `自動套印學生姓名.py`，支援：
     - Windows 11：調用 PowerShell Word COM 自動化（< 3 秒極速匯出）。
     - macOS：調用 AppleScript 控制 Word for Mac 原生輸出，若未安裝則自動回退 LibreOffice。
   - 解決 macOS 執行 Windows 專用 PowerShell 指令中斷的跨平台缺陷。
3. **終端與 Python 編碼全平台防護**：
   - Windows 11 自動啟用 UTF-8 I/O 封裝；macOS 保持原生 UTF-8，全平台無損讀寫學生罕用字名條。
4. **雙系統暫存檔過濾整備**：
   - `.gitignore` 擴充過濾 `.DS_Store`、`._*`、`Thumbs.db`、`Desktop.ini` 與 Office 暫存檔，確保兩端同步時庫存乾淨。
5. **已完成套印與名冊產出**：
   - `班費收費通知_四年丙班28人_橫式一頁6張.pdf`（5 頁 A4，28 人姓名座號 + 2 張備用單）。
   - `班費收費通知_四年丙班28人_橫式一頁6張.docx`。
   - `學生名冊與收費核對表_四年丙班.xlsx`。

---

## 📌 下一階段行動建議
- 在 Windows 11 或 macOS 下均可直接執行 `python3 自動套印學生姓名.py`（或 `python 自動套印學生姓名.py`）進行自動套印。
- 學生紙本發放：直接以 100% 實際大小列印 `班費收費通知_四年丙班28人_橫式一頁6張.pdf`（共 5 張 A4）。
- 行政收費核對：使用 `學生名冊與收費核對表_四年丙班.xlsx` 登記勾選。
