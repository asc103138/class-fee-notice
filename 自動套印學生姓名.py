#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
班費收費通知 - 全班學生姓名自動套印與 PDF 匯出工具
支援：
1. 自動讀取「班級名條.xlsx」之班級代號與學生名冊（或回退讀取核對表）
2. 精確套印姓名與座號至「A4 橫式一頁 6 張」及「A4 直式一頁 4 張」版面
3. 產生與更新「學生名冊與收費核對表_四年丙班.xlsx」
4. 自動調用 Word 引擎匯出 100% 零跑版之高畫質 PDF 套印檔
"""

import os
import sys
import io
import subprocess
import openpyxl
import docx

# Ensure UTF-8 output on Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

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

    # 方式一：在 Windows 平台直接調用 PowerShell Word COM 物件（最高保真度）
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
    res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], capture_output=True, text=True)
    if res.returncode == 0 and os.path.exists(pdf_path):
        print(f"✅ 已成功轉換原生 PDF：{pdf_path}")
        return True
    else:
        print(f"❌ Word COM 轉換失敗：{res.stderr}")
        return False

def main():
    print("🚀 開始執行班費收費通知單套印程序...")
    class_name, students = load_roster()
    print(f"📋 班級：{class_name}，學生人數：{len(students)} 人")

    # 1. 橫式 A4 一頁 6 張（主要首選）
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
