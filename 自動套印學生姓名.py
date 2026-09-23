#!/usr/bin/env python3
"""
班費收費通知 - 全班學生姓名自動套印工具
使用方式：
1. 在「學生名冊與收費核對表_四上班費.xlsx」填入每位學生的姓名（B欄）
2. 執行本腳本：python3 自動套印學生姓名.py
3. 即會自動產出填好姓名的：
   - 班費收費通知_四上全班28人_橫式一頁6張_含姓名版.docx (推薦！只需5張A4)
   - 班費收費通知_四上全班28人_直式一頁4張_含姓名版.docx (7張A4)
"""

import os
import openpyxl
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "學生名冊與收費核對表_四上班費.xlsx")
OUT_LANDSCAPE_NAMES = os.path.join(BASE_DIR, "班費收費通知_四上全班28人_橫式一頁6張_含姓名版.docx")
OUT_PORTRAIT_NAMES = os.path.join(BASE_DIR, "班費收費通知_四上全班28人_直式一頁4張_含姓名版.docx")

def get_students_from_excel():
    names = []
    if os.path.exists(EXCEL_PATH):
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
        for row in range(5, 33): # 28 students
            val = ws.cell(row=row, column=2).value
            names.append(str(val).strip() if val else "")
    else:
        names = ["" for _ in range(28)]
    return names

def create_compact_slip(cell, seat_num=None, student_name=""):
    p_title = cell.paragraphs[0]
    p_title.text = ""
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("四上預定收費項目")
    r_title.font.name = "標楷體"
    r_title._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    r_title.font.size = Pt(14)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(192, 0, 0)
    
    p_info = cell.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_info.paragraph_format.space_before = Pt(1)
    p_info.paragraph_format.space_after = Pt(3.5)
    name_str = f" {student_name} " if student_name else "____________"
    seat_str = f"{seat_num:02d} 號" if seat_num is not None else "____ 號"
    r_info = p_info.add_run(f"四年____班   座號：{seat_str}   姓名：{name_str}")
    r_info.font.name = "標楷體"
    r_info._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    r_info.font.size = Pt(9.5)
    r_info.font.bold = True

    items = [
        "① 運動會道具：50 元",
        "② 影印及雜支：172 元",
        "③ 學用品：78 元(明細如下)"
    ]
    for i, item in enumerate(items):
        p_item = cell.add_paragraph()
        p_item.paragraph_format.left_indent = Cm(0.2)
        p_item.paragraph_format.space_before = Pt(0)
        p_item.paragraph_format.space_after = Pt(1.5) if i < 2 else Pt(3)
        p_item.paragraph_format.line_spacing = 1.05
        r_item = p_item.add_run(item)
        r_item.font.name = "標楷體"
        r_item._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
        r_item.font.size = Pt(9.5)

    p_tot = cell.add_paragraph()
    p_tot.paragraph_format.left_indent = Cm(0.2)
    p_tot.paragraph_format.space_before = Pt(1)
    p_tot.paragraph_format.space_after = Pt(3.5)
    r_tot = p_tot.add_run("本學期現金收費合計：300 元")
    r_tot.font.name = "標楷體"
    r_tot._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    r_tot.font.size = Pt(10)
    r_tot.font.bold = True
    r_tot.font.highlight_color = WD_COLOR_INDEX.YELLOW

    tbl = cell.add_table(rows=3, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_data = [
        ["學用品名稱", "單價", "數量", "價格"],
        ["寒假作業", "28", "1", "28"],
        ["國語隨堂練習", "50", "1", "50"]
    ]
    col_widths = [Cm(3.4), Cm(1.3), Cm(1.3), Cm(1.3)]
    for row_idx, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:trHeight {} w:val="280" w:hRule="atLeast"/>'.format(nsdecls('w'))))
        for col_idx, c in enumerate(row.cells):
            c.width = col_widths[col_idx]
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = c._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(r'''
                <w:tcBorders {}>
                    <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                </w:tcBorders>
            '''.format(nsdecls('w'))))
            tcPr.append(parse_xml(r'''
                <w:tcMar {}>
                    <w:top w:w="40" w:type="dxa"/>
                    <w:bottom w:w="40" w:type="dxa"/>
                    <w:left w:w="80" w:type="dxa"/>
                    <w:right w:w="80" w:type="dxa"/>
                </w:tcMar>
            '''.format(nsdecls('w'))))
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (col_idx > 0 or row_idx == 0) else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(tbl_data[row_idx][col_idx])
            r.font.name = "標楷體"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
            r.font.size = Pt(8.5)
            if row_idx == 0:
                r.font.bold = True

    p_note = cell.paragraphs[-1]
    p_note.paragraph_format.left_indent = Cm(0.2)
    p_note.paragraph_format.space_before = Pt(4)
    p_note.paragraph_format.space_after = Pt(2.5)
    r_n1 = p_note.add_run("發下 ")
    r_n1.font.name = "標楷體"
    r_n1._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    r_n1.font.size = Pt(8.5)
    r_box = p_note.add_run("現金收費袋")
    r_box.font.name = "標楷體"
    r_box._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    r_box.font.size = Pt(8.5)
    r_box.font.bold = True
    r_box._element.get_or_add_rPr().append(parse_xml(r'<w:bdr {} w:val="single" w:sz="6" w:space="2" w:color="000000"/>'.format(nsdecls('w'))))
    r_n2 = p_note.add_run("，請於本週內交給老師。謝謝！")
    r_n2.font.name = "標楷體"
    r_n2._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    r_n2.font.size = Pt(8.5)

    p_sig = cell.add_paragraph()
    p_sig.paragraph_format.left_indent = Cm(0.2)
    p_sig.paragraph_format.space_before = Pt(2)
    p_sig.paragraph_format.space_after = Pt(2)
    r_sig = p_sig.add_run("家長簽名 【　　　　　　　】")
    r_sig.font.name = "標楷體"
    r_sig._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    r_sig.font.size = Pt(9.5)
    r_sig.font.bold = True

def build_landscape_names():
    names = get_students_from_excel()
    doc = docx.Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(0.8)
    section.bottom_margin = Cm(0.8)
    section.left_margin = Cm(0.8)
    section.right_margin = Cm(0.8)

    total_students = 28
    pages = 5
    for page in range(pages):
        table = doc.add_table(rows=2, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_i, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:trHeight {} w:val="5300" w:hRule="atLeast"/>'.format(nsdecls('w'))))
            trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))
            for c_i, cell in enumerate(row.cells):
                cell.width = Cm(9.2)
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(r'''
                    <w:tcBorders {}>
                        <w:top w:val="dashed" w:sz="6" w:space="0" w:color="A0A0A0"/>
                        <w:left w:val="dashed" w:sz="6" w:space="0" w:color="A0A0A0"/>
                        <w:bottom w:val="dashed" w:sz="6" w:space="0" w:color="A0A0A0"/>
                        <w:right w:val="dashed" w:sz="6" w:space="0" w:color="A0A0A0"/>
                    </w:tcBorders>
                '''.format(nsdecls('w'))))
                tcPr.append(parse_xml(r'''
                    <w:tcMar {}>
                        <w:top w:w="100" w:type="dxa"/>
                        <w:bottom w:w="100" w:type="dxa"/>
                        <w:left w:w="140" w:type="dxa"/>
                        <w:right w:w="140" w:type="dxa"/>
                    </w:tcMar>
                '''.format(nsdecls('w'))))
                
                idx = page * 6 + r_i * 3 + c_i
                seat_num = idx + 1
                if seat_num <= total_students:
                    s_name = names[idx] if idx < len(names) else ""
                    create_compact_slip(cell, seat_num, s_name)
                else:
                    create_compact_slip(cell, None, "")
                    
                if page > 0 and r_i == 0 and c_i == 0:
                    cell.paragraphs[0].paragraph_format.page_break_before = True
                    
    doc.save(OUT_LANDSCAPE_NAMES)
    print(f"✅ 成功產出橫式姓名版：{OUT_LANDSCAPE_NAMES}")

if __name__ == "__main__":
    build_landscape_names()
