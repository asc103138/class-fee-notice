#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
班費收費通知 - 全班學生姓名自動套印與 PDF 匯出工具（跨平台相容版）
相容環境：macOS (Apple Silicon / Intel) 與 Windows 11 / 10

支援功能：
1. 自動讀取「班級名條.xlsx」之班級代號與學生名冊（或回退讀取核對表）
2. 精確套印姓名與座號至「A4 橫式一頁 6 張」及「A4 直式一頁 4 張」版面
3. 產生與更新「學生名冊與收費核對表_{班級}.xlsx」
4. 跨平台無損 PDF 匯出：
   - Windows：調用本機 Word COM 自動化（< 3 秒原生無損）
   - macOS：調用 AppleScript 控制 Word for Mac（或回退 LibreOffice）
"""

import os
import sys
import io
import shutil
import subprocess
import openpyxl
import docx

# Windows 終端強制 UTF-8 輸出，防止 CP950 編碼報錯
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ROSTER_EXCEL = os.path.join(BASE_DIR, "班級名條.xlsx")
CHECKLIST_EXCEL = os.path.join(BASE_DIR, "學生名冊與收費核對表_四上班費.xlsx")
TEMPLATE_LANDSCAPE = os.path.join(BASE_DIR, "班費收費通知_四上全班28人_橫式一頁6張.docx")
TEMPLATE_PORTRAIT = os.path.join(BASE_DIR, "班費收費通知_四上全班28人列印版.docx")

def load_roster():
    students = []
    class_name = "四年丙班"

    if os.path.exists(ROSTER_EXCEL):
        wb = openpyxl.load_workbook(ROSTER_EXCEL)
        ws = wb.active
        c1 = ws.cell(1, 1).value
        if c1:
            class_name = str(c1).strip()
        for r in range(3, ws.max_row + 1):
            seat = ws.cell(r, 1).value
            sid = ws.cell(r, 2).value
            gender = ws.cell(r, 3).value
            name = ws.cell(r, 4).value
            if seat is not None and name is not None:
                students.append({
                    "seat": int(seat),
                    "name": str(name).strip(),
                    "gender": str(gender).strip() if gender else "",
                    "sid": str(sid).strip() if sid else ""
                })
    elif os.path.exists(CHECKLIST_EXCEL):
        wb = openpyxl.load_workbook(CHECKLIST_EXCEL)
        ws = wb.active
        for row in range(5, 33):
            seat_val = ws.cell(row=row, column=1).value
            name_val = ws.cell(row=row, column=2).value
            if seat_val:
                students.append({
                    "seat": int(seat_val),
                    "name": str(name_val).strip() if name_val else "",
                    "gender": "",
                    "sid": ""
                })

    students.sort(key=lambda s: s["seat"])
    return class_name, students

def generate_landscape_docx(class_name, students, output_path):
    if not os.path.exists(TEMPLATE_LANDSCAPE):
        raise FileNotFoundError(f"找不到範本檔：{TEMPLATE_LANDSCAPE}")
    
    doc = docx.Document(TEMPLATE_LANDSCAPE)
    for t_idx, t in enumerate(doc.tables):
        for r_idx, r in enumerate(t.rows):
            for c_idx, c in enumerate(r.cells):
                idx = t_idx * 6 + r_idx * 3 + c_idx
                p1 = c.paragraphs[1]
                if idx < len(students):
                    s = students[idx]
                    p1.runs[0].text = f"{class_name}   座號：{s['seat']:02d} 號   姓名：{s['name']}"
                else:
                    p1.runs[0].text = f"{class_name}   座號：____ 號   姓名：____________"
    
    doc.save(output_path)
    print(f"✅ 已成功產出橫式套印 Word：{output_path}")

def generate_portrait_docx(class_name, students, output_path):
    if not os.path.exists(TEMPLATE_PORTRAIT):
        print(f"⚠️ 找不到直式範本檔：{TEMPLATE_PORTRAIT}，略過直式套印。")
        return
    
    doc = docx.Document(TEMPLATE_PORTRAIT)
    for t_idx, t in enumerate(doc.tables):
        for r_idx, r in enumerate(t.rows):
            for c_idx, c in enumerate(r.cells):
                idx = t_idx * 4 + r_idx * 2 + c_idx
                p1 = c.paragraphs[1]
                if idx < len(students):
                    s = students[idx]
                    p1.runs[0].text = f"{class_name}   座號：{s['seat']:02d} 號   姓名：{s['name']}"
                else:
                    p1.runs[0].text = f"{class_name}   座號：____ 號   姓名：____________"
    
    doc.save(output_path)
    print(f"✅ 已成功產出直式套印 Word：{output_path}")

def generate_checklist_excel(class_name, students, output_path):
    if not os.path.exists(CHECKLIST_EXCEL):
        return
    
    wb = openpyxl.load_workbook(CHECKLIST_EXCEL)
    ws = wb.active
    ws.cell(1, 1).value = f"{class_name} 班費收費核對名冊（四上）"
    for s in students:
        row_num = 4 + s["seat"]
        ws.cell(row_num, 2).value = s["name"]
    
    wb.save(output_path)
    print(f"✅ 已成功產出學生收費核對表：{output_path}")

def convert_docx_to_pdf(docx_path, pdf_path):
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path)

    # 1. Windows 平台：調用 PowerShell Word COM 物件
    if sys.platform == "win32":
        ps_cmd = f"""
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        try {{
            $doc = $word.Documents.Open('{docx_path}')
            $doc.SaveAs([ref]'{pdf_path}', [ref]17)
            $doc.Close([ref]0)
        }} finally {{
            $word.Quit()
        }}
        """
        try:
            res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], capture_output=True, text=True)
            if res.returncode == 0 and os.path.exists(pdf_path):
                print(f"✅ [Windows Word COM] 原生 PDF 轉換成功：{pdf_path}")
                return True
        except Exception as e:
            print(f"⚠️ Windows Word COM 執行異常：{e}")

    # 2. macOS 平台：調用 AppleScript 控制 Word for Mac
    elif sys.platform == "darwin":
        applescript = f'''
        tell application "Microsoft Word"
            set wasRunning to running
            open POSIX file "{docx_path}"
            set activeDoc to active document
            save as activeDoc file format format PDF file name "{pdf_path}"
            close activeDoc saving no
            if not wasRunning then
                quit
            end if
        end tell
        '''
        try:
            res = subprocess.run(["osascript", "-e", applescript], capture_output=True, text=True)
            if res.returncode == 0 and os.path.exists(pdf_path):
                print(f"✅ [macOS Word] 原生 PDF 轉換成功：{pdf_path}")
                return True
        except Exception as e:
            print(f"⚠️ macOS AppleScript Word 執行異常：{e}")

    # 3. 跨平台備用：LibreOffice
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin and sys.platform == "darwin":
        mac_lo = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(mac_lo):
            soffice_bin = mac_lo

    if soffice_bin:
        out_dir = os.path.dirname(pdf_path)
        try:
            res = subprocess.run([soffice_bin, "--headless", "--convert-to", "pdf", docx_path, "--outdir", out_dir], capture_output=True, text=True)
            if os.path.exists(pdf_path):
                print(f"✅ [LibreOffice] PDF 轉換成功：{pdf_path}")
                return True
        except Exception as e:
            print(f"⚠️ LibreOffice 轉換異常：{e}")

    print(f"⚠️ 提示：未偵測到可用之本機 Word 或 LibreOffice 轉檔引擎。已產出 Word 檔（{docx_path}），請手動另存為 PDF。")
    return False

def main():
    print("🚀 開始執行班費收費通知單套印程序（macOS / Windows 跨平台）...")
    class_name, students = load_roster()
    print(f"📋 班級：{class_name}，學生人數：{len(students)} 人")

    # 1. 橫式 A4 一頁 6 張（主要首選，全班僅需 5 頁，含 2 張備用單）
    out_landscape_c = os.path.join(BASE_DIR, f"班費收費通知_{class_name}{len(students)}人_橫式一頁6張.docx")
    generate_landscape_docx(class_name, students, out_landscape_c)

    # 匯出橫式 PDF
    pdf_landscape_c = os.path.join(BASE_DIR, f"班費收費通知_{class_name}{len(students)}人_橫式一頁6張.pdf")
    convert_docx_to_pdf(out_landscape_c, pdf_landscape_c)

    # 2. 直式 A4 一頁 4 張（備用）
    out_portrait_c = os.path.join(BASE_DIR, f"班費收費通知_{class_name}{len(students)}人_直式一頁4張_含姓名版.docx")
    generate_portrait_docx(class_name, students, out_portrait_c)
    pdf_portrait_c = os.path.join(BASE_DIR, f"班費收費通知_{class_name}{len(students)}人_直式一頁4張_含姓名版.pdf")
    convert_docx_to_pdf(out_portrait_c, pdf_portrait_c)

    # 3. 更新行政收費核對名冊
    out_checklist_c = os.path.join(BASE_DIR, f"學生名冊與收費核對表_{class_name}.xlsx")
    generate_checklist_excel(class_name, students, out_checklist_c)

    print("🎉 全套班費收費通知單與 PDF 套印檔製作完成！")

if __name__ == "__main__":
    main()
