# app.py
# 基督教聖約教會堅樂小學 - 陸運會行政自動化與成績計算平台 (All-in-One 完整版)
import streamlit as st
import pandas as pd
import numpy as np
import json
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

st.set_page_config(page_title="堅樂小學陸運會平台", layout="wide")

CONFIG_PATH = "quota_config.json"
QUOTA_MATRIX_PATH = "quota_matrix.json"
STUDENTS_FILE = "saved_students.xlsx"
RESULTS_FILE = "tournament_results.json"
REG_DATA_DIR = "registrations_data"
SCHEDULES_DIR = "schedules_data"

os.makedirs(REG_DATA_DIR, exist_ok=True)
os.makedirs(SCHEDULES_DIR, exist_ok=True)

ZH_TO_NUM = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6"}

def get_grade_digit(text):
    if not text: return "1"
    first = str(text).strip()[0]
    return ZH_TO_NUM.get(first, first if first.isdigit() else "1")

# 決賽種子分配順序 (1名->3線, 2名->4線, 3名->5線, 4名->6線, 5名->7線, 6名->2線, 7名->1線, 8名->8線)
FINAL_SEED_LANES = [3, 4, 5, 6, 7, 2, 1, 8]

MASTER_EVENTS = {
    # 1XX: 徑賽初賽 (不評分，分組排列，選拔前8名進入決賽)
    101: "一年級男子40米跑", 102: "一年級女子40米跑",
    103: "二年級男子40米跑", 104: "二年級女子40米跑",
    105: "三年級男子40米跑", 106: "三年級女子40米跑",
    107: "一年級男子60米跑", 108: "一年級女子60米跑",
    109: "二年級男子60米跑", 110: "二年級女子60米跑",
    111: "三年級男子60米跑", 112: "三年級女子60米跑",
    113: "四年級男子60米跑", 114: "四年級女子60米跑",
    115: "五年級男子60米跑", 116: "五年級女子60米跑",
    117: "六年級男子60米跑", 118: "六年級女子60米跑",
    119: "四年級男子100米跑", 120: "四年級女子100米跑",
    121: "五年級男子100米跑", 122: "五年級女子100米跑",
    123: "六年級男子100米跑", 124: "六年級女子100米跑",

    # 2XX: 200m直接決賽及全校田賽決賽 (給分 9-1)
    225: "五年級男子200米跑", 226: "五年級女子200米跑",
    227: "六年級男子200米跑", 228: "六年級女子200米跑",
    229: "一年級男子擲木球", 230: "一年級女子擲木球",
    231: "二年級男子擲木球", 232: "二年級女子擲木球",
    233: "三年級男子擲木球", 234: "三年級女子擲木球",
    235: "四年級男子擲壘球", 236: "四年級女子擲壘球",
    237: "五年級男子擲壘球", 238: "五年級女子擲壘球",
    239: "六年級男子擲壘球", 240: "六年級女子擲壘球",
    241: "一年級男子立定跳遠", 242: "一年級女子立定跳遠",
    243: "二年級男子立定跳遠", 244: "二年級女子立定跳遠",
    245: "三年級男子立定跳遠", 246: "三年級女子立定跳遠",
    247: "四年級男子助跑跳遠", 248: "四年級女子助跑跳遠",
    249: "五年級男子助跑跳遠", 250: "五年級女子助跑跳遠",
    251: "六年級男子助跑跳遠", 252: "六年級女子助跑跳遠",
    253: "一年級男子單人跳前繩", 254: "一年級女子單人跳前繩",
    255: "二年級男子單人跳前繩", 256: "二年級女子單人跳前繩",
    257: "三年級男子單人跳前繩", 258: "三年級女子單人跳前繩",
    259: "四年級男子單人跳前繩", 260: "四年級女子單人跳前繩",

    # 3XX: 徑賽決賽 (給分 9-1，對應初賽+260)
    361: "一年級男子40米跑(決賽)", 362: "一年級女子40米跑(決賽)",
    363: "二年級男子40米跑(決賽)", 364: "二年級女子40米跑(決賽)",
    365: "三年級男子40米跑(決賽)", 366: "三年級女子40米跑(決賽)",
    367: "一年級男子60米跑(決賽)", 368: "一年級女子60米跑(決賽)",
    369: "二年級男子60米跑(決賽)", 370: "二年級女子60米跑(決賽)",
    371: "三年級男子60米跑(決賽)", 372: "三年級女子60米跑(決賽)",
    373: "四年級男子60米跑(決賽)", 374: "四年級女子60米跑(決賽)",
    375: "五年級男子60米跑(決賽)", 376: "五年級女子60米跑(決賽)",
    377: "六年級男子60米跑(決賽)", 378: "六年級女子60米跑(決賽)",
    379: "四年級男子100米跑(決賽)", 380: "四年級女子100米跑(決賽)",
    381: "五年級男子100米跑(決賽)", 382: "五年級女子100米跑(決賽)",
    383: "六年級男子100米跑(決賽)", 384: "六年級女子100米跑(決賽)",

    # 4XX: 接力決賽 (給分 9-1)
    485: "四年級4×100米接力跑",
    486: "五年級4×100米接力跑",
    487: "六年級4×100米接力跑",
}

POINT_SYSTEM = {1: 9, 2: 7, 3: 6, 4: 5, 5: 4, 6: 3, 7: 2, 8: 1}

GRADE_EVENTS = {
    1: ["擲木球", "立定跳遠", "單人跳前繩", "40米跑", "60米跑"],
    2: ["擲木球", "立定跳遠", "單人跳前繩", "40米跑", "60米跑"],
    3: ["擲木球", "立定跳遠", "單人跳前繩", "40米跑", "60米跑"],
    4: ["擲壘球", "助跑跳遠", "單人跳前繩", "60米跑", "100米跑"],
    5: ["擲壘球", "助跑跳遠", "60米跑", "100米跑", "200米跑"],
    6: ["擲壘球", "助跑跳遠", "60米跑", "100米跑", "200米跑"],
}

# =====================================================================
# 內建 Word 排版輔助引擎
# =====================================================================
def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_padding(cell, top=45, bottom=45, left=50, right=50):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_clean_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="A6ACAF"/>'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="A6ACAF"/>'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="A6ACAF"/>'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="A6ACAF"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D5D8DC"/>'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="D5D8DC"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def make_table_cant_split(table):
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def build_word_document(events_data, school_name, session_title, output_filename):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.45)
        section.right_margin = Inches(0.45)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after = Pt(2)
    
    r1 = title_p.add_run(f"{school_name}\n")
    r1.font.size = Pt(17)
    r1.font.bold = True
    
    sess = session_title if ("屆" in session_title or "年度" in session_title) else f"第{session_title}屆"
    r2 = title_p.add_run(f"{sess}陸運會 比賽程序秩序冊（含檢錄次序表）\n")
    r2.font.size = Pt(13)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(70, 70, 70)

    for ev in events_data:
        eid = ev.get("id", "")
        ename = ev.get("name", "比賽項目")
        etype = ev.get("type", "track")

        p_head = doc.add_paragraph()
        p_head.paragraph_format.space_before = Pt(8)
        p_head.paragraph_format.space_after = Pt(2)
        p_head.paragraph_format.keep_with_next = True

        id_str = f"【場次編號：{eid}】  " if eid else "【示範賽事】  "
        r_id = p_head.add_run(id_str)
        r_id.font.bold = True
        r_id.font.size = Pt(10)
        r_id.font.color.rgb = RGBColor(0, 51, 102)

        r_nm = p_head.add_run(f"項目名稱：{ename}")
        r_nm.font.bold = True
        r_nm.font.size = Pt(10)

        if etype == "relay":
            matrix = ev.get("matrix", [])
            if matrix:
                table = doc.add_table(rows=len(matrix), cols=len(matrix[0]))
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = False
                add_clean_borders(table)
                make_table_cant_split(table)

                for col in table.columns: col.width = Inches(1.4)
                for i, row in enumerate(matrix):
                    for j, val in enumerate(row):
                        cell = table.cell(i, j)
                        cell.text = str(val) if pd.notna(val) else ""
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        set_cell_padding(cell, 45, 45, 60, 60)
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        if p.runs:
                            if i == 0:
                                p.runs[0].font.size = Pt(9.5)
                                p.runs[0].font.bold = True
                                set_cell_background(cell, "EAEDED")
                            elif i == 1:
                                p.runs[0].font.size = Pt(9)
                                p.runs[0].font.bold = True
                                set_cell_background(cell, "F2F4F4")
                            else:
                                p.runs[0].font.size = Pt(8.5)

        elif etype == "field":
            athletes = ev.get("athletes", [])
            chunk_size = 10
            chunks = [athletes[i:i + chunk_size] for i in range(0, max(1, len(athletes)), chunk_size)]
            total_rows = max(1, len(chunks)) * 2

            table = doc.add_table(rows=total_rows, cols=10)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False
            add_clean_borders(table)
            make_table_cant_split(table)

            for c_idx, chunk in enumerate(chunks):
                h_r = c_idx * 2
                d_r = h_r + 1
                start_n = c_idx * chunk_size + 1

                for j in range(10):
                    h_cell = table.cell(h_r, j)
                    d_cell = table.cell(d_r, j)
                    set_cell_padding(h_cell, 35, 35, 40, 40)
                    set_cell_padding(d_cell, 35, 35, 40, 40)

                    if j < len(chunk):
                        h_cell.text = f"次序{start_n + j}"
                        d_cell.text = str(chunk[j])
                        set_cell_background(h_cell, "F2F4F4")
                    else:
                        h_cell.text = ""
                        d_cell.text = ""

                    hp = h_cell.paragraphs[0]
                    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if hp.runs:
                        hp.runs[0].font.size = Pt(7.5)
                        hp.runs[0].font.bold = True

                    dp = d_cell.paragraphs[0]
                    dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if dp.runs: dp.runs[0].font.size = Pt(8)

        else:
            grid = ev.get("grid", [])
            if grid:
                num_cols = min(9, max(len(row) for row in grid))
                table = doc.add_table(rows=len(grid), cols=num_cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = False
                add_clean_borders(table)
                make_table_cant_split(table)

                for i, row in enumerate(grid):
                    for j in range(num_cols):
                        cell = table.cell(i, j)
                        val = row[j] if j < len(row) else ""
                        val = "" if str(val) == "nan" else val
                        cell.text = str(val)
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        set_cell_padding(cell, 40, 40, 50, 50)

                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        if p.runs:
                            p.runs[0].font.size = Pt(8)
                            if i == 0 or (j == 0 and "組別" in str(val)):
                                p.runs[0].font.bold = True

                        if i == 0: set_cell_background(cell, "F2F4F4")
                        elif j == 0 and "組別" in str(val): set_cell_background(cell, "FAFAFA")

        doc.add_paragraph().paragraph_format.space_after = Pt(3)

    doc.save(output_filename)

def build_score_sheets_doc(events_data, school_name, session_title, output_filename):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    sess = session_title if ("屆" in session_title or "年度" in session_title) else f"第{session_title}屆"
    total_events = len(events_data)

    for ev_idx, ev in enumerate(events_data):
        eid = ev.get("id", "")
        ename = ev.get("name", "比賽項目")
        etype = ev.get("type", "track")

        if etype in ["track", "track_final"]:
            grid = ev.get("grid", [])
            data_rows = grid[1:] if len(grid) > 1 else []

            for r_idx, row in enumerate(data_rows):
                heat_name = row[0] if row else f"組別{r_idx+1}"

                p_top = doc.add_paragraph()
                p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_t1 = p_top.add_run(f"{school_name}\n{sess}陸運會 - 徑賽裁判執法記錄表\n")
                r_t1.font.bold = True
                r_t1.font.size = Pt(14)

                p_info = doc.add_paragraph()
                p_info.paragraph_format.space_after = Pt(4)
                p_info.add_run(f"【場次編號：{eid}】  項目：{ename}  ｜  組別：{heat_name}  ｜  地點：跑道\n").bold = True
                p_info.add_run("比賽時間：_______:_______    天氣：□ 晴  □ 陰  □ 雨    風速：_______ m/s").font.size = Pt(9)

                table = doc.add_table(rows=9, cols=8)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                add_clean_borders(table)
                make_table_cant_split(table)

                col_titles = ["線道", "班別", "學號", "選手姓名", "終點時間 (秒)", "複查時間", "名次", "備註 (DNS/DQ)"]
                col_widths = [0.8, 0.9, 0.9, 1.6, 1.4, 1.4, 0.8, 1.4]

                for c_i, (t_txt, w_val) in enumerate(zip(col_titles, col_widths)):
                    cell = table.cell(0, c_i)
                    cell.width = Inches(w_val)
                    cell.text = t_txt
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cell.paragraphs[0].runs[0].font.size = Pt(8.5)
                    cell.paragraphs[0].runs[0].font.bold = True
                    set_cell_background(cell, "EAEDED")
                    set_cell_padding(cell, 50, 50, 40, 40)

                for lane in range(1, 9):
                    cell_r = table.cell(lane, 0)
                    cell_r.text = f"第 {lane} 線"
                    cell_r.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cell_r.paragraphs[0].runs[0].font.size = Pt(8.5)
                    cell_r.paragraphs[0].runs[0].font.bold = True
                    set_cell_background(cell_r, "F8F9F9")

                    runner_val = row[lane] if lane < len(row) else ""
                    c_class, c_name = "", ""
                    if runner_val and str(runner_val) != "nan":
                        c_class = runner_val[:2] if len(runner_val) >= 2 else ""
                        c_name = runner_val[2:] if len(runner_val) >= 2 else runner_val

                    table.cell(lane, 1).text = c_class
                    table.cell(lane, 2).text = ""
                    table.cell(lane, 3).text = c_name

                    for col_j in range(1, 8):
                        c = table.cell(lane, col_j)
                        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        if c.paragraphs[0].runs: c.paragraphs[0].runs[0].font.size = Pt(8.5)
                        set_cell_padding(c, 50, 50, 40, 40)

                p_sign = doc.add_paragraph()
                p_sign.paragraph_format.space_before = Pt(14)
                p_sign.paragraph_format.keep_with_next = True
                p_sign.add_run("發令員簽署：____________________    終點裁判長簽署：____________________    計時主任簽署：____________________").font.size = Pt(9)
                doc.add_page_break()

        elif etype == "field":
            athletes = ev.get("athletes", [])
            p_top = doc.add_paragraph()
            p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_t1 = p_top.add_run(f"{school_name}\n{sess}陸運會 - 田賽裁判執法記錄表\n")
            r_t1.font.bold = True
            r_t1.font.size = Pt(14)

            p_info = doc.add_paragraph()
            p_info.paragraph_format.space_after = Pt(4)
            p_info.add_run(f"【場次編號：{eid}】  項目：{ename}  ｜  地點：田賽場區\n").bold = True
            p_info.add_run("比賽規則：每人試跳／試擲 2 次，取最佳成績；平手比次佳成績。犯規記 X，缺席記 -。").font.size = Pt(8.5)

            total_athletes = max(10, len(athletes))
            table = doc.add_table(rows=total_athletes + 1, cols=9)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            add_clean_borders(table)
            make_table_cant_split(table)

            f_titles = ["次序", "班別", "姓名", "第 1 次試跳/擲", "第 2 次試跳/擲", "最佳成績", "次佳成績", "名次", "備註"]
            f_widths = [0.8, 1.0, 1.6, 1.4, 1.4, 1.2, 1.2, 0.8, 1.2]

            for c_i, (t_txt, w_val) in enumerate(zip(f_titles, f_widths)):
                cell = table.cell(0, c_i)
                cell.width = Inches(w_val)
                cell.text = t_txt
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.paragraphs[0].runs[0].font.size = Pt(8.5)
                cell.paragraphs[0].runs[0].font.bold = True
                set_cell_background(cell, "EAEDED")
                set_cell_padding(cell, 45, 45, 35, 35)

            for a_idx in range(total_athletes):
                row_idx = a_idx + 1
                table.cell(row_idx, 0).text = str(a_idx + 1)
                ath_val = athletes[a_idx] if a_idx < len(athletes) else ""
                c_class = ath_val[:2] if len(ath_val) >= 2 else ""
                c_name = ath_val[2:] if len(ath_val) >= 2 else ath_val

                table.cell(row_idx, 1).text = c_class
                table.cell(row_idx, 2).text = c_name

                for c_j in range(9):
                    c = table.cell(row_idx, c_j)
                    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if c.paragraphs[0].runs: c.paragraphs[0].runs[0].font.size = Pt(8.5)
                    set_cell_padding(c, 45, 45, 35, 35)

            p_sign = doc.add_paragraph()
            p_sign.paragraph_format.space_before = Pt(14)
            p_sign.paragraph_format.keep_with_next = True
            p_sign.add_run("田賽裁判長簽署：____________________        記錄員簽署：____________________        賽事覆核員簽署：____________________").font.size = Pt(9)
            if ev_idx < total_events - 1: doc.add_page_break()

        elif etype == "relay":
            matrix = ev.get("matrix", [])
            p_top = doc.add_paragraph()
            p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_t1 = p_top.add_run(f"{school_name}\n{sess}陸運會 - 班際接力裁判執法記錄表\n")
            r_t1.font.bold = True
            r_t1.font.size = Pt(14)

            p_info = doc.add_paragraph()
            p_info.paragraph_format.space_after = Pt(4)
            p_info.add_run(f"【場次編號：{eid}】  項目：{ename}  ｜  地點：跑道\n").bold = True
            p_info.add_run("規則：A班3線、B班4線、C班5線、D班6線、E班7線。第1至8名分別得 9, 7, 6, 5, 4, 3, 2, 1 分。").font.size = Pt(8.5)

            table = doc.add_table(rows=6, cols=8)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            add_clean_borders(table)
            make_table_cant_split(table)

            r_titles = ["線道", "班別", "第 1 棒", "第 2 棒", "第 3 棒", "第 4 棒", "大會成績 (秒)", "名次"]
            r_widths = [0.9, 0.9, 1.4, 1.4, 1.4, 1.4, 1.6, 0.8]

            for c_i, (t_txt, w_val) in enumerate(zip(r_titles, r_widths)):
                cell = table.cell(0, c_i)
                cell.width = Inches(w_val)
                cell.text = t_txt
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.paragraphs[0].runs[0].font.size = Pt(8.5)
                cell.paragraphs[0].runs[0].font.bold = True
                set_cell_background(cell, "EAEDED")
                set_cell_padding(cell, 50, 50, 40, 40)

            classes = matrix[0] if len(matrix) > 0 else ["4A", "4B", "4C", "4D", "4E"]
            l1 = matrix[2] if len(matrix) > 2 else ["", "", "", "", ""]
            l2 = matrix[3] if len(matrix) > 3 else ["", "", "", "", ""]
            l3 = matrix[4] if len(matrix) > 4 else ["", "", "", "", ""]
            l4 = matrix[5] if len(matrix) > 5 else ["", "", "", "", ""]

            for idx, ln in enumerate([3, 4, 5, 6, 7]):
                r_num = idx + 1
                table.cell(r_num, 0).text = f"第 {ln} 線"
                table.cell(r_num, 1).text = classes[idx] if idx < len(classes) else ""
                table.cell(r_num, 2).text = l1[idx] if idx < len(l1) else ""
                table.cell(r_num, 3).text = l2[idx] if idx < len(l2) else ""
                table.cell(r_num, 4).text = l3[idx] if idx < len(l3) else ""
                table.cell(r_num, 5).text = l4[idx] if idx < len(l4) else ""

                for c_k in range(8):
                    c = table.cell(r_num, c_k)
                    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if c.paragraphs[0].runs: c.paragraphs[0].runs[0].font.size = Pt(8.5)
                    set_cell_padding(c, 50, 50, 40, 40)

            p_sign = doc.add_paragraph()
            p_sign.paragraph_format.space_before = Pt(14)
            p_sign.paragraph_format.keep_with_next = True
            p_sign.add_run("接力檢查長簽署：____________________        終點裁判長簽署：____________________        計時主任簽署：____________________").font.size = Pt(9)
            if ev_idx < total_events - 1: doc.add_page_break()

    doc.save(output_filename)

def build_awards_doc(standing_df, school_name, session_title, output_filename):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.45)
        section.bottom_margin = Inches(0.45)
        section.left_margin = Inches(0.45)
        section.right_margin = Inches(0.45)

    sess = session_title if ("屆" in session_title or "年度" in session_title) else f"第{session_title}屆"

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    
    r_t1 = p_title.add_run(f"{school_name}\n")
    r_t1.font.bold = True
    r_t1.font.size = Pt(16)

    r_t2 = p_title.add_run(f"{sess}陸運會 團體獎頒獎名冊（級際及初／高級組錦標）\n")
    r_t2.font.bold = True
    r_t2.font.size = Pt(13)
    r_t2.font.color.rgb = RGBColor(0, 51, 102)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(8)
    p_sub.add_run("（得獎依據：平均分數 = (比賽得分 - 棄權扣分) ÷ 實際參賽人數）").font.size = Pt(8.5)

    # 1. 分組團體錦標 (初級組 P1-3、高級組 P4-6)
    p_h1 = doc.add_paragraph()
    p_h1.paragraph_format.space_before = Pt(4)
    p_h1.paragraph_format.space_after = Pt(3)
    p_h1.paragraph_format.keep_with_next = True
    r_h1 = p_h1.add_run("🏆 分組團體錦標（初級組 P1–3、高級組 P4–6）")
    r_h1.font.bold = True
    r_h1.font.size = Pt(11)
    r_h1.font.color.rgb = RGBColor(180, 40, 40)

    groups_info = [
        ("初級組團體錦標 (P1–3)", "初級組 (P1-3)"),
        ("高級組團體錦標 (P4–6)", "高級組 (P4-6)")
    ]

    table_div = doc.add_table(rows=3, cols=4)
    table_div.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_clean_borders(table_div)
    make_table_cant_split(table_div)

    d_headers = ["組別名銜", "團體冠軍 🥇", "團體亞軍 🥈", "團體季軍 🥉"]
    d_widths = [2.2, 1.8, 1.8, 1.8]
    for c_i, (t_txt, w_val) in enumerate(zip(d_headers, d_widths)):
        c = table_div.cell(0, c_i)
        c.width = Inches(w_val)
        c.text = t_txt
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_background(c, "FADBD8")
        set_cell_padding(c, 40, 40, 30, 30)

    for g_idx, (g_display, g_key) in enumerate(groups_info):
        r_i = g_idx + 1
        table_div.cell(r_i, 0).text = g_display
        table_div.cell(r_i, 0).paragraphs[0].runs[0].font.bold = True
        set_cell_background(table_div.cell(r_i, 0), "FDEDEC")

        sub_df = standing_df[standing_df["組別"] == g_key].sort_values(by="組排名").reset_index(drop=True)
        for place in range(3):
            col_target = place + 1
            cell = table_div.cell(r_i, col_target)
            if place < len(sub_df):
                c_data = sub_df.iloc[place]
                cell.text = f"{c_data['班別']} ({c_data['平均分數']:.3f}分)"
                cell.paragraphs[0].runs[0].font.bold = True
            else:
                cell.text = "—"
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_padding(cell, 35, 35, 30, 30)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 2. 級際團體獎 (一至六年級 P1-P6 各級際名次)
    p_h2 = doc.add_paragraph()
    p_h2.paragraph_format.space_before = Pt(6)
    p_h2.paragraph_format.space_after = Pt(3)
    p_h2.paragraph_format.keep_with_next = True
    r_h2 = p_h2.add_run("🏅 級際團體獎（各年級 P1–P6 優勝班別）")
    r_h2.font.bold = True
    r_h2.font.size = Pt(11)
    r_h2.font.color.rgb = RGBColor(0, 102, 153)

    table_grade = doc.add_table(rows=7, cols=4)
    table_grade.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_clean_borders(table_grade)
    make_table_cant_split(table_grade)

    g_titles = ["年級", "級際冠軍 🥇", "級際亞軍 🥈", "級際季軍 🥉"]
    g_widths = [1.6, 2.0, 2.0, 2.0]
    for c_i, (t_txt, w_val) in enumerate(zip(g_titles, g_widths)):
        c = table_grade.cell(0, c_i)
        c.width = Inches(w_val)
        c.text = t_txt
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_background(c, "D4E6F1")
        set_cell_padding(c, 35, 35, 25, 25)

    for g_i in range(1, 7):
        r_i = g_i
        table_grade.cell(r_i, 0).text = f"{g_i}年級 (P{g_i})"
        table_grade.cell(r_i, 0).paragraphs[0].runs[0].font.bold = True
        set_cell_background(table_grade.cell(r_i, 0), "EBF5FB")

        g_df = standing_df[standing_df["年級"] == f"{g_i}年級"].sort_values(by="級排名").reset_index(drop=True)
        for place in range(3):
            col_target = place + 1
            cell = table_grade.cell(r_i, col_target)
            if place < len(g_df):
                c_data = g_df.iloc[place]
                cell.text = f"{c_data['班別']} ({c_data['平均分數']:.3f}分)"
                cell.paragraphs[0].runs[0].font.bold = True
            else:
                cell.text = "—"
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_padding(cell, 35, 35, 25, 25)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 3. 全校各班總成績詳細審核大表
    p_h3 = doc.add_paragraph()
    p_h3.paragraph_format.space_before = Pt(6)
    p_h3.paragraph_format.space_after = Pt(3)
    p_h3.paragraph_format.keep_with_next = True
    r_h3 = p_h3.add_run("📊 各班團體總成績詳細審核總表（按年級排序）")
    r_h3.font.bold = True
    r_h3.font.size = Pt(11)

    table_all = doc.add_table(rows=len(standing_df) + 1, cols=10)
    table_all.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_clean_borders(table_all)
    make_table_cant_split(table_all)

    col_titles = ["年級", "班別", "所屬組別", "級排名", "組排名", "比賽得分", "棄權扣減", "實得分", "參賽人數", "平均分"]
    col_widths = [0.8, 0.7, 1.3, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7, 1.0]

    for c_i, (t_txt, w_val) in enumerate(zip(col_titles, col_widths)):
        c = table_all.cell(0, c_i)
        c.width = Inches(w_val)
        c.text = t_txt
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8)
        set_cell_background(c, "EAEDED")
        set_cell_padding(c, 35, 35, 20, 20)

    df_display = standing_df.sort_values(by=["年級", "級排名"]).reset_index(drop=True)
    for r_i, r in df_display.iterrows():
        row_idx = r_i + 1
        vals = [
            str(r["年級"]),
            str(r["班別"]),
            str(r["組別"]),
            f"第 {r['級排名']} 名",
            f"第 {r['組排名']} 名",
            str(r["得分"]),
            str(r["棄權扣減"]),
            str(r["實得總分"]),
            str(r["實際參賽人數"]),
            f"{r['平均分數']:.3f}"
        ]
        for c_j, val in enumerate(vals):
            cell = table_all.cell(row_idx, c_j)
            cell.text = val
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.size = Pt(7.5)
            set_cell_padding(cell, 30, 30, 15, 15)

            if r["級排名"] == 1:
                set_cell_background(cell, "EBF5FB" if c_j in [1, 3, 9] else "FAFAFA")
                if c_j in [1, 3, 9]: cell.paragraphs[0].runs[0].font.bold = True

    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(18)
    p_sign.paragraph_format.keep_with_next = True
    r_s = p_sign.add_run("體育科主任簽署：____________________        大會裁判長簽署：____________________        校長簽核：____________________")
    r_s.font.size = Pt(9.5)

    doc.save(output_filename)

# =====================================================================
# 系統輔助與資料庫函式
# =====================================================================
def get_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f: return json.load(f)
    return {"school_name": "基督教聖約教會堅樂小學", "session_title": "第二十四屆", "relay_min_boys": 1, "relay_min_girls": 1}

def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f: json.dump(cfg, f, ensure_ascii=False, indent=2)

def get_quota_matrix():
    if os.path.exists(QUOTA_MATRIX_PATH):
        with open(QUOTA_MATRIX_PATH, "r", encoding="utf-8") as f: return json.load(f)
    return {}

def save_quota_matrix(data):
    with open(QUOTA_MATRIX_PATH, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)

def load_saved_students():
    if os.path.exists(STUDENTS_FILE):
        try: return pd.read_excel(STUDENTS_FILE)
        except Exception: return None
    return None

def save_students_to_disk(df): df.to_excel(STUDENTS_FILE, index=False)

def save_class_reg(class_name, records):
    path = os.path.join(REG_DATA_DIR, f"reg_{class_name}.json")
    with open(path, "w", encoding="utf-8") as f: json.dump(records, f, ensure_ascii=False, indent=2)

def load_class_reg(class_name):
    path = os.path.join(REG_DATA_DIR, f"reg_{class_name}.json")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f: return json.load(f)
    return []

def save_class_relay(class_name, relay_dict):
    path = os.path.join(REG_DATA_DIR, f"relay_{class_name}.json")
    with open(path, "w", encoding="utf-8") as f: json.dump(relay_dict, f, ensure_ascii=False, indent=2)

def load_class_relay(class_name):
    path = os.path.join(REG_DATA_DIR, f"relay_{class_name}.json")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f: return json.load(f)
    return None

def load_all_registrations():
    all_regs = []
    if os.path.exists(REG_DATA_DIR):
        for f in os.listdir(REG_DATA_DIR):
            if f.startswith("reg_") and f.endswith(".json"):
                with open(os.path.join(REG_DATA_DIR, f), "r", encoding="utf-8") as fp:
                    try: all_regs.extend(json.load(fp))
                    except: pass
    return pd.DataFrame(all_regs) if all_regs else pd.DataFrame(columns=["班別", "學號", "姓名", "性別", "項目"])

def save_custom_schedule(eid, df):
    path = os.path.join(SCHEDULES_DIR, f"sch_{eid}.json")
    df.to_json(path, orient="records", force_ascii=False, indent=2)

def load_custom_schedule(eid):
    path = os.path.join(SCHEDULES_DIR, f"sch_{eid}.json")
    if os.path.exists(path):
        try: return pd.read_json(path)
        except: return None
    return None

def get_tournament_results():
    if os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {"results": [], "absences": [], "qualifiers": {}}

def save_tournament_results(data):
    with open(RESULTS_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)

def parse_track_time(raw_val):
    if raw_val is None: return None, ""
    s = str(raw_val).strip().upper()
    if not s or s == "NAN": return None, ""
    if s in ["DNS", "DQ", "DNF", "X", "-", "棄權", "犯規"]: return None, s
    clean_s = s.replace("秒", "").replace('"', "").replace("'", "").replace("S", "")
    try:
        sec = float(clean_s) if "." in clean_s else int(clean_s) / 100.0
        return sec, f"{sec:.2f}"
    except ValueError:
        return None, s

def parse_field_attempt(val):
    if val is None: return np.nan
    s = str(val).strip().upper()
    if not s or s in ["-", "DNS", "棄權"]: return np.nan
    if s in ["X", "DQ", "犯規"]: return 0.0
    try: return float(s)
    except ValueError: return np.nan

def process_field_event_results(athletes_df):
    df = athletes_df.copy()
    df["試1_數值"] = df["第一次"].apply(parse_field_attempt)
    df["試2_數值"] = df["第二次"].apply(parse_field_attempt)
    df["最佳成績"] = df[["試1_數值", "試2_數值"]].max(axis=1)
    df["次佳成績"] = df[["試1_數值", "試2_數值"]].min(axis=1)

    valid_mask = df["最佳成績"].notna() & (df["最佳成績"] > 0)
    df_valid = df[valid_mask].copy()
    df_invalid = df[~valid_mask].copy()

    df_valid.sort_values(by=["最佳成績", "次佳成績"], ascending=[False, False], inplace=True)
    df_valid.reset_index(drop=True, inplace=True)
    df_valid["名次"] = df_valid.index + 1
    df_valid["得分"] = df_valid["名次"].map(POINT_SYSTEM).fillna(0).astype(int)

    df_invalid["名次"] = "-"
    df_invalid["得分"] = 0
    final_df = pd.concat([df_valid, df_invalid], ignore_index=True)
    return final_df.drop(columns=["試1_數值", "試2_數值", "次佳成績"])

def generate_auto_track_schedule(eid, ename):
    g_num = get_grade_digit(ename)
    g = int(g_num)
    gender = "男" if "男子" in ename else ("女" if "女子" in ename else "混合")
    clean_item = ename.replace("一年級", "").replace("二年級", "").replace("三年級", "")
    clean_item = clean_item.replace("四年級", "").replace("五年級", "").replace("六年級", "")
    clean_item = clean_item.replace("男子", "").replace("女子", "").replace("(決賽)", "")

    df_reg = load_all_registrations()
    matches = df_reg[
        (df_reg["性別"] == gender) & 
        (df_reg["項目"] == clean_item) & 
        (df_reg["班別"].str.startswith(str(g), na=False))
    ].copy()

    lane_map = {"A": 1, "B": 2, "C": 3, "D": 4, "E": 5}
    rows = []

    if not matches.empty:
        matches["學號_數值"] = pd.to_numeric(matches["學號"], errors="coerce").fillna(99)
        matches.sort_values(by=["班別", "學號_數值"], inplace=True)
        for c, c_df in matches.groupby("班別"):
            class_letter = c[-1].upper() if len(c) > 1 else "A"
            assigned_lane = lane_map.get(class_letter, 1)
            for heat_idx, (_, r) in enumerate(c_df.iterrows()):
                rows.append({
                    "組別": f"組別{heat_idx + 1}",
                    "線道": assigned_lane,
                    "班別": r["班別"],
                    "學號": str(r["學號"]),
                    "姓名": str(r["姓名"])
                })

    if not rows:
        for grp_num in [1, 2]:
            for idx, l in enumerate(["A", "B", "C", "D", "E"]):
                rows.append({"組別": f"組別{grp_num}", "線道": idx + 1, "班別": f"{g_num}{l}", "學號": "", "姓名": "（尚無報名）"})

    res_df = pd.DataFrame(rows)
    return res_df.sort_values(by=["組別", "線道"]).reset_index(drop=True)

def generate_auto_field_schedule(eid, ename):
    g_num = get_grade_digit(ename)
    g = int(g_num)
    gender = "男" if "男子" in ename else ("女" if "女子" in ename else "混合")
    clean_item = ename.replace("一年級", "").replace("二年級", "").replace("三年級", "")
    clean_item = clean_item.replace("四年級", "").replace("五年級", "").replace("六年級", "")
    clean_item = clean_item.replace("男子", "").replace("女子", "")

    df_reg = load_all_registrations()
    matches = df_reg[
        (df_reg["性別"] == gender) & 
        (df_reg["項目"] == clean_item) & 
        (df_reg["班別"].str.startswith(str(g), na=False))
    ].copy()

    rows = []
    if not matches.empty:
        matches["學號_數值"] = pd.to_numeric(matches["學號"], errors="coerce").fillna(99)
        matches.sort_values(by=["班別", "學號_數值"], inplace=True)
        for idx, (_, r) in enumerate(matches.iterrows()):
            rows.append({"出場次序": idx + 1, "班別": r["班別"], "學號": str(r["學號"]), "姓名": str(r["姓名"])})

    if not rows:
        for idx in range(5): rows.append({"出場次序": idx + 1, "班別": f"{g_num}A", "學號": f"{idx+1}", "姓名": "（尚無報名）"})

    return pd.DataFrame(rows)

def generate_auto_relay_schedule(eid, ename):
    g_num = str(eid - 481)
    class_list = [f"{g_num}A", f"{g_num}B", f"{g_num}C", f"{g_num}D", f"{g_num}E"]
    lanes = [3, 4, 5, 6, 7]
    rows = []

    for c, l in zip(class_list, lanes):
        relay_data = load_class_relay(c) or {}
        legs = relay_data.get("legs", ["", "", "", ""])
        while len(legs) < 4: legs.append("")
        rows.append({
            "線道": l, "班別": c,
            "第1棒": legs[0] if legs[0] else "未指派",
            "第2棒": legs[1] if legs[1] else "未指派",
            "第3棒": legs[2] if legs[2] else "未指派",
            "第4棒": legs[3] if legs[3] else "未指派",
        })
    return pd.DataFrame(rows)

def auto_compile_events_from_database():
    events_data = []
    t_data = get_tournament_results()

    for eid in sorted(MASTER_EVENTS.keys()):
        ename = MASTER_EVENTS[eid]

        if 485 <= eid <= 487:
            sch = load_custom_schedule(eid)
            if sch is None: sch = generate_auto_relay_schedule(eid, ename)
            row0 = sch["班別"].tolist()
            row1 = [f"線道{x}" for x in sch["線道"].tolist()]
            matrix = [row0, row1, sch["第1棒"].tolist(), sch["第2棒"].tolist(), sch["第3棒"].tolist(), sch["第4棒"].tolist()]
            events_data.append({"id": str(eid), "name": ename, "type": "relay", "matrix": matrix})

        elif 229 <= eid <= 260:
            sch = load_custom_schedule(eid)
            if sch is None: sch = generate_auto_field_schedule(eid, ename)
            athletes = [f"{r['班別']}{r['姓名']}" for _, r in sch.iterrows() if r["姓名"] != "（尚無報名）"]
            events_data.append({"id": str(eid), "name": ename, "type": "field", "athletes": athletes})

        elif (101 <= eid <= 124) or (225 <= eid <= 228):
            sch = load_custom_schedule(eid)
            if sch is None: sch = generate_auto_track_schedule(eid, ename)
            header = ["組別", "線道1", "線道2", "線道3", "線道4", "線道5", "線道6", "線道7", "線道8"]
            grid = [header]
            for grp_name, grp_df in sch.groupby("組別"):
                row = [grp_name]
                lane_dict = {int(r["線道"]): f"{r['班別']}{r['姓名']}" for _, r in grp_df.iterrows() if r["姓名"] != "（尚無報名）"}
                for ln in range(1, 9): row.append(lane_dict.get(ln, ""))
                grid.append(row)
            events_data.append({"id": str(eid), "name": ename, "type": "track", "grid": grid})

        else:
            header = ["組別", "線道1", "線道2", "線道3", "線道4", "線道5", "線道6", "線道7", "線道8"]
            row = ["決賽", "", "", "", "", "", "", "", ""]
            qual_list = t_data.get("qualifiers", {}).get(str(eid), [])
            if qual_list:
                for q in qual_list:
                    ln = int(q.get("lane", 1))
                    if 1 <= ln <= 8: row[ln] = f"{q['class']}{q['name']}"
            events_data.append({"id": str(eid), "name": ename, "type": "track", "grid": [header, row]})

    return events_data

cfg = get_config()

if "df_students" not in st.session_state:
    saved_df = load_saved_students()
    if saved_df is not None: st.session_state["df_students"] = saved_df

st.sidebar.title("🏃 陸運會行政平台")
st.sidebar.caption(f"{cfg['school_name']}\n{cfg['session_title']}")

if "df_students" in st.session_state:
    stu_count = len(st.session_state["df_students"])
    st.sidebar.success(f"🟢 已載入學生名冊\n（共 {stu_count} 人，免再上傳）")
else:
    st.sidebar.warning("🟠 尚未建立學生名冊")

menu = st.sidebar.radio(
    "功能模組導航",
    [
        "1. 體育科主任：各班各項目配額設定",
        "2. 班主任：學生選拔報名登記",
        "3. 自動排線與主任線上微調",
        "4. 一鍵匯出 Word 秩序冊",
        "5. 現場成績錄入與班名次統計",
        "⚙️ 大會全域設定"
    ]
)

# ----------------------------------------------------
# 模組 1：體育科主任配額設定
# ----------------------------------------------------
if menu == "1. 體育科主任：各班各項目配額設定":
    st.title("⚙️ 體育科主任：各班各項目參賽人數比例設定")
    has_students = "df_students" in st.session_state
    with st.expander("📂 學生名冊管理（上傳一次永久儲存，若需更換名冊請在此操作）", expanded=not has_students):
        uploaded_file = st.file_uploader("上傳全校學生名冊 Excel/CSV (含欄位：班別、學號、姓名、性別)", type=["xlsx", "csv"])
        if uploaded_file:
            new_df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(".xlsx") else pd.read_csv(uploaded_file)
            save_students_to_disk(new_df)
            st.session_state["df_students"] = new_df
            st.success(f"🎉 學生名冊已永久儲存至本機硬碟，共 {len(new_df)} 人！日後免再上傳。")
            st.rerun()

        if has_students:
            c_info, c_del = st.columns([4, 1])
            with c_info: st.caption(f"目前儲存檔案：`{STUDENTS_FILE}` ｜ 學生人數：{len(st.session_state['df_students'])} 人")
            with c_del:
                if st.button("🗑️ 清除現存名單"):
                    if os.path.exists(STUDENTS_FILE): os.remove(STUDENTS_FILE)
                    del st.session_state["df_students"]
                    st.rerun()

    if "df_students" in st.session_state:
        df_stu = st.session_state["df_students"]
        class_stats = df_stu.groupby(["班別", "性別"]).size().unstack(fill_value=0)
        for col in ["男", "女"]:
            if col not in class_stats.columns: class_stats[col] = 0

        target_grade = st.radio("選擇設定年級", [1, 2, 3, 4, 5, 6], horizontal=True, format_func=lambda x: f"{x}年級")
        ev_list = GRADE_EVENTS[target_grade]
        classes = sorted([c for c in df_stu["班別"].unique() if str(c).startswith(str(target_grade))])
        quotas = get_quota_matrix()
        
        tab_titles = ["👦 男子組項目配額", "👧 女子組項目配額"]
        if target_grade in [4, 5, 6]: tab_titles.append("🎽 班際 4×100米 接力選項設定")
        tabs = st.tabs(tab_titles)
        
        with tabs[0]:
            b_rows = []
            for c in classes:
                r = {"班別": c, "男生人數": int(class_stats.loc[c, "男"]) if c in class_stats.index else 0}
                for ev in ev_list: r[ev] = quotas.get(f"{c}_男_{ev}", 2)
                b_rows.append(r)
            b_df = pd.DataFrame(b_rows)
            ed_boys = st.data_editor(b_df, disabled=["班別", "男生人數"], hide_index=True, key=f"b_{target_grade}", use_container_width=True)

        with tabs[1]:
            g_rows = []
            for c in classes:
                r = {"班別": c, "女生人數": int(class_stats.loc[c, "女"]) if c in class_stats.index else 0}
                for ev in ev_list: r[ev] = quotas.get(f"{c}_女_{ev}", 2)
                g_rows.append(r)
            g_df = pd.DataFrame(g_rows)
            ed_girls = st.data_editor(g_df, disabled=["班別", "女生人數"], hide_index=True, key=f"g_{target_grade}", use_container_width=True)

        if target_grade in [4, 5, 6]:
            with tabs[2]:
                st.subheader(f"🎽 {target_grade} 年級班際 4×100米 混合接力規則設定")
                c_r1, c_r2 = st.columns(2)
                with c_r1: r_b_min = st.number_input(f"{target_grade}年級 正選男生最低人數", min_value=0, max_value=3, value=quotas.get(f"relay_g{target_grade}_min_b", 1), key=f"rg_b_{target_grade}")
                with c_r2: r_g_min = st.number_input(f"{target_grade}年級 正選女生最低人數", min_value=0, max_value=3, value=quotas.get(f"relay_g{target_grade}_min_g", 1), key=f"rg_g_{target_grade}")

                relay_class_status = [{"班別": c, "開放接力報名": quotas.get(f"relay_enabled_{c}", True), "正選人數": 4, "後備名額 (男女各1)": 2} for c in classes]
                df_rc = pd.DataFrame(relay_class_status)
                ed_relay_class = st.data_editor(df_rc, disabled=["班別", "正選人數", "後備名額 (男女各1)"], hide_index=True, key=f"rc_ed_{target_grade}", use_container_width=True)

        st.divider()
        if st.button(f"💾 儲存 {target_grade} 年級全部配額設定", type="primary"):
            for _, r in ed_boys.iterrows():
                for ev in ev_list: quotas[f"{r['班別']}_男_{ev}"] = int(r[ev])
            for _, r in ed_girls.iterrows():
                for ev in ev_list: quotas[f"{r['班別']}_女_{ev}"] = int(r[ev])
            
            if target_grade in [4, 5, 6]:
                quotas[f"relay_g{target_grade}_min_b"] = int(r_b_min)
                quotas[f"relay_g{target_grade}_min_g"] = int(r_g_min)
                for _, r in ed_relay_class.iterrows(): quotas[f"relay_enabled_{r['班別']}"] = bool(r["開放接力報名"])

            save_quota_matrix(quotas)
            st.success(f"{target_grade} 年級個人賽配額及班際接力設定已成功儲存！")

# ----------------------------------------------------
# 模組 2：班主任報名登記 (滿額取消顯示項目)
# ----------------------------------------------------
elif menu == "2. 班主任：學生選拔報名登記":
    st.title("👨‍🏫 班主任：學生選拔報名登記")
    if "df_students" not in st.session_state:
        st.warning("⚠️ 系統尚未建立學生名冊，請先在「模組 1」上傳學生名冊！")
    else:
        df_stu = st.session_state["df_students"]
        all_classes = sorted(df_stu["班別"].unique())
        c_pick = st.selectbox("請選擇登記班別", all_classes)
        grade = int(c_pick[0]) if c_pick[0].isdigit() else 1
        class_df = df_stu[df_stu["班別"] == c_pick].sort_values(by="學號")
        quotas = get_quota_matrix()
        
        saved_reg = load_class_reg(c_pick)
        saved_reg_map = {str(r["學號"]): r["項目"] for r in saved_reg}

        tab_list = ["個人項目報名"]
        if grade in [4, 5, 6]: tab_list.append("🎽 班際 4×100米 接力")
        tabs = st.tabs(tab_list)

        with tabs[0]:
            st.subheader(f"📌 {c_pick} 個人項目登記（每人限報最多 1 項）")
            events = GRADE_EVENTS.get(grade, [])

            current_selections = {}
            for _, s in class_df.iterrows():
                sid = str(s["學號"])
                k = f"reg_{c_pick}_{sid}"
                cur_val = st.session_state.get(k, saved_reg_map.get(sid, "-- 不參賽 --"))
                current_selections[sid] = cur_val

            live_counts = {"男": {ev: 0 for ev in events}, "女": {ev: 0 for ev in events}}
            for _, s in class_df.iterrows():
                sid, g = str(s["學號"]), str(s["性別"])
                cur_val = current_selections.get(sid, "-- 不參賽 --")
                if cur_val in events: live_counts[g][cur_val] += 1

            st.markdown("##### 📊 本班各項目名額即時動態看板（滿額項目將自動從未報名同學的選單中隱藏）：")
            board_tabs = st.tabs(["👦 男子組名額狀態", "👧 女子組名額狀態"])
            for g_idx, g_name in enumerate(["男", "女"]):
                with board_tabs[g_idx]:
                    cols = st.columns(len(events))
                    for idx, ev in enumerate(events):
                        c_used = live_counts[g_name][ev]
                        c_max = quotas.get(f"{c_pick}_{g_name}_{ev}", 2)
                        c_remain = c_max - c_used
                        with cols[idx]:
                            if c_used >= c_max: st.warning(f"**{ev}**\n\n**{c_used} / {c_max}** (滿額 🔒)\n\n*(選單已隱藏)*")
                            else: st.success(f"**{ev}**\n\n**{c_used} / {c_max}**\n\n*(剩 {c_remain} 席 🟢)*")

            st.write("---")
            records_to_save = []

            for _, s in class_df.iterrows():
                sid, name, g = str(s["學號"]), str(s["姓名"]), str(s["性別"])
                cur_chosen = current_selections.get(sid, "-- 不參賽 --")

                available_events = []
                for ev in events:
                    max_limit = quotas.get(f"{c_pick}_{g}_{ev}", 2)
                    others_taken = live_counts[g][ev] - (1 if cur_chosen == ev else 0)
                    if others_taken < max_limit: available_events.append(ev)

                opts = ["-- 不參賽 --"] + available_events
                if cur_chosen in events and cur_chosen not in opts: opts.append(cur_chosen)

                init_idx = opts.index(cur_chosen) if cur_chosen in opts else 0

                c_stu, c_status = st.columns([3, 2])
                with c_stu:
                    chosen_ev = st.selectbox(f"【{sid}號】{name} ({g})", options=opts, index=init_idx, key=f"reg_{c_pick}_{sid}")
                with c_status:
                    st.write("")
                    st.write("")
                    if chosen_ev != "-- 不參賽 --":
                        u = live_counts[g][chosen_ev]
                        m = quotas.get(f"{c_pick}_{g}_{chosen_ev}", 2)
                        if u == m: st.caption(f"🔒 【{g}子{chosen_ev}】滿額 ({u}/{m})")
                        else: st.caption(f"✅ 【{g}子{chosen_ev}】有效 ({u}/{m}，剩 {m-u} 席)")
                        records_to_save.append({"班別": c_pick, "學號": sid, "姓名": name, "性別": g, "項目": chosen_ev})
                    else:
                        st.caption("⚪ 未選報個人賽")

            st.write("---")
            if st.button(f"💾 儲存 {c_pick} 個人賽報名名單", type="primary"):
                save_class_reg(c_pick, records_to_save)
                st.success(f"🎉 {c_pick} 個人賽報名名冊已成功儲存至伺服器！（共登記 {len(records_to_save)} 人次）")

        if grade in [4, 5, 6]:
            with tabs[1]:
                st.subheader(f"🎽 {c_pick} 班際 4×100米 接力名單（棒次自選 + 彈性比例）")
                min_b = quotas.get(f"relay_g{grade}_min_b", cfg.get("relay_min_boys", 1))
                min_g = quotas.get(f"relay_g{grade}_min_g", cfg.get("relay_min_girls", 1))
                st.info(f"📋 **大會規則**：4 位正選中，男生最少 **{min_b}** 人，女生最少 **{min_g}** 人。棒次完全自由自選。")
                
                saved_relay = load_class_relay(c_pick) or {}
                saved_legs = saved_relay.get("legs", ["", "", "", ""])

                names_opts = ["-- 請選擇 --"] + [f"{r['學號']}號 {r['姓名']} ({r['性別']})" for _, r in class_df.iterrows()]
                
                def get_leg_default_idx(leg_idx):
                    if leg_idx < len(saved_legs):
                        target = saved_legs[leg_idx]
                        for idx, opt in enumerate(names_opts):
                            if target and target in opt: return idx
                    return 0

                col1, col2, col3, col4 = st.columns(4)
                with col1: l1 = st.selectbox("🏃 第 1 棒", names_opts, index=get_leg_default_idx(0), key=f"leg1_{c_pick}")
                with col2: l2 = st.selectbox("🏃 第 2 棒", names_opts, index=get_leg_default_idx(1), key=f"leg2_{c_pick}")
                with col3: l3 = st.selectbox("🏃 第 3 棒", names_opts, index=get_leg_default_idx(2), key=f"leg3_{c_pick}")
                with col4: l4 = st.selectbox("🏃 第 4 棒", names_opts, index=get_leg_default_idx(3), key=f"leg4_{c_pick}")

                st.write("---")
                c_sb, c_sg = st.columns(2)
                b_pool = ["-- 無 --"] + [f"{r['學號']}號 {r['姓名']}" for _, r in class_df[class_df["性別"] == "男"].iterrows()]
                g_pool = ["-- 無 --"] + [f"{r['學號']}號 {r['姓名']}" for _, r in class_df[class_df["性別"] == "女"].iterrows()]
                
                sb_def_idx = b_pool.index(saved_relay.get("sub_b")) if saved_relay.get("sub_b") in b_pool else 0
                sg_def_idx = g_pool.index(saved_relay.get("sub_g")) if saved_relay.get("sub_g") in g_pool else 0

                with c_sb: sub_b = st.selectbox("男子後備 (1名)", b_pool, index=sb_def_idx, key=f"sub_b_{c_pick}")
                with c_sg: sub_g = st.selectbox("女子後備 (1名)", g_pool, index=sg_def_idx, key=f"sub_g_{c_pick}")

                chosen_legs = [x for x in [l1, l2, l3, l4] if x != "-- 請選擇 --"]
                if len(chosen_legs) == 4:
                    if len(set(chosen_legs)) < 4:
                        st.error("❌ 錯誤：同一位學生不可重複兼任多個棒次！")
                    else:
                        b_cnt = sum(1 for x in chosen_legs if "(男)" in x)
                        g_cnt = sum(1 for x in chosen_legs if "(女)" in x)
                        if b_cnt >= min_b and g_cnt >= min_g:
                            st.success(f"✅ 名單合規！陣容：{b_cnt} 男 {g_cnt} 女。")
                            if st.button("💾 儲存接力名單", key=f"save_relay_{c_pick}"):
                                save_class_relay(c_pick, {
                                    "legs": [x.split(" ")[1].replace("(", "").replace(")", "").replace("男", "").replace("女", "") for x in chosen_legs],
                                    "sub_b": sub_b,
                                    "sub_g": sub_g
                                })
                                st.success(f"🎉 {c_pick} 接力隊伍名單已成功儲存至伺服器！")
                        else:
                            st.error(f"❌ 性別比例不符合規定！目前選擇：{b_cnt} 男、{g_cnt} 女。（規定：男生至少 {min_b} 人，女生至少 {min_g} 人）")
                else:
                    st.warning("⚠️ 請完整指定第 1 至第 4 棒選手。")

# ----------------------------------------------------
# 模組 3：自動排線與主任線上微調
# ----------------------------------------------------
elif menu == "3. 自動排線與主任線上微調":
    st.title("⚙️ 體育科主任：自動排線與線上微調")
    event_options = [f"{eid} ｜ {mname}" for eid, mname in sorted(MASTER_EVENTS.items())]
    selected_event_str = st.selectbox("🎯 請選擇要檢視或微調的場次", event_options)
    
    sel_eid = int(selected_event_str.split(" ｜ ")[0])
    sel_ename = selected_event_str.split(" ｜ ")[1]

    c_m3_title, c_m3_print = st.columns([3, 1])
    with c_m3_title: st.write(f"### 🎯 【場次 {sel_eid}】{sel_ename}")
    with c_m3_print:
        single_doc_name = f"裁判記錄表_場次{sel_eid}_{sel_ename}.docx"
        if st.button(f"🖨️ 產生【場次 {sel_eid}】裁判記錄紙", type="secondary"):
            single_ev_data = [ev for ev in auto_compile_events_from_database() if str(ev["id"]) == str(sel_eid)]
            if single_ev_data:
                build_score_sheets_doc(single_ev_data, cfg["school_name"], cfg["session_title"], single_doc_name)
                with open(single_doc_name, "rb") as fp:
                    st.download_button(
                        label=f"📥 下載場次 {sel_eid} 記錄紙 (.docx)",
                        data=fp.read(),
                        file_name=single_doc_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

    if sel_eid in [485, 486, 487]:
        st.info("💡 規則：A班-3線、B班-4線、C班-5線、D班-6線、E班-7線。棒次自動同步各班提交之名單。")
        custom_df = load_custom_schedule(sel_eid)
        if custom_df is None or st.button("🔄 重新從各班報名同步最新接力棒次"):
            custom_df = generate_auto_relay_schedule(sel_eid, sel_ename)

        edited_relay = st.data_editor(custom_df, hide_index=True, use_container_width=True)
        if st.button(f"🚀 確認定稿並儲存【場次 {sel_eid}】接力排線", type="primary"):
            save_custom_schedule(sel_eid, edited_relay)
            st.success(f"🎉 【場次 {sel_eid}】接力名單已正式定稿儲存！")

    elif 229 <= sel_eid <= 260:
        st.info("💡 規則：全級選手依「班別＋學號」遞增嚴格排序為出場次序 1 至 N。")
        custom_df = load_custom_schedule(sel_eid)
        if custom_df is None or st.button("🔄 重新從各班報名名冊一鍵自動排序"):
            custom_df = generate_auto_field_schedule(sel_eid, sel_ename)

        edited_field = st.data_editor(
            custom_df,
            column_config={
                "出場次序": st.column_config.NumberColumn(min_value=1, step=1),
                "班別": st.column_config.TextColumn(disabled=True),
                "學號": st.column_config.TextColumn(disabled=True),
                "姓名": st.column_config.TextColumn(disabled=True),
            },
            hide_index=True,
            use_container_width=True
        )
        if st.button(f"🚀 確認定稿並儲存【場次 {sel_eid}】田賽出場次序", type="primary"):
            save_custom_schedule(sel_eid, edited_field.sort_values(by="出場次序").reset_index(drop=True))
            st.success(f"🎉 【場次 {sel_eid}】田賽次序已定稿儲存！")

    elif (101 <= sel_eid <= 124) or (225 <= sel_eid <= 228):
        st.info("💡 規則：A班-1線、B班-2線、C班-3線、D班-4線、E班-5線。同班學號小者入第 1 組，大者入第 2 組。主任可微調各組線道。")
        custom_df = load_custom_schedule(sel_eid)
        if custom_df is None or st.button("🔄 重新從各班報名名冊一鍵自動排線"):
            custom_df = generate_auto_track_schedule(sel_eid, sel_ename)

        unique_heats = sorted(custom_df["組別"].unique())
        heat_edit_tabs = st.tabs([f"📋 {h} 道次" for h in unique_heats])
        all_edited_parts = []

        has_lane_conflict = False
        for h_idx, h_name in enumerate(unique_heats):
            with heat_edit_tabs[h_idx]:
                sub_df = custom_df[custom_df["組別"] == h_name].sort_values(by="線道").reset_index(drop=True)
                ed_sub = st.data_editor(
                    sub_df,
                    column_config={
                        "組別": st.column_config.TextColumn(disabled=True),
                        "線道": st.column_config.NumberColumn("線道 (1-8)", min_value=1, max_value=8, step=1),
                        "班別": st.column_config.TextColumn(disabled=True),
                        "學號": st.column_config.TextColumn(disabled=True),
                        "姓名": st.column_config.TextColumn(disabled=True),
                    },
                    hide_index=True,
                    use_container_width=True,
                    key=f"sch_heat_ed_{sel_eid}_{h_name}"
                )
                all_edited_parts.append(ed_sub)

                valid_runners = ed_sub[ed_sub["姓名"] != "（尚無報名）"]
                dups = valid_runners["線道"].value_counts()
                c_lanes = dups[dups > 1].index.tolist()
                if c_lanes:
                    st.error(f"❌ 警告：【{h_name}】的第 {c_lanes} 線道有重複選手！請調開線道。")
                    has_lane_conflict = True

        combined_edited_track = pd.concat(all_edited_parts, ignore_index=True)

        if not has_lane_conflict:
            st.success("✅ 所有組別與線道設定良好，無重複衝突。")
            if st.button(f"🚀 確認定稿並儲存【場次 {sel_eid}】排線", type="primary"):
                save_custom_schedule(sel_eid, combined_edited_track.sort_values(by=["組別", "線道"]).reset_index(drop=True))
                st.success(f"🎉 【場次 {sel_eid}】徑賽分組排線已定稿儲存！")
    else:
        st.subheader(f"🏁 【場次 {sel_eid}】{sel_ename} 決賽道次")
        t_data_temp = get_tournament_results()
        qual_list = t_data_temp.get("qualifiers", {}).get(str(sel_eid), [])
        if qual_list:
            st.success("📋 已成功載入初賽晉級之 8 強選手，道次已按大會種子順序 (3,4,5,6,7,2,1,8) 排定：")
            df_q_preview = pd.DataFrame(qual_list).sort_values(by="lane")
            st.dataframe(df_q_preview[["lane", "rank", "class", "name", "heat_time"]].rename(
                columns={"lane": "線道", "rank": "初賽名次", "class": "班別", "name": "姓名", "heat_time": "初賽秒數"}
            ), hide_index=True, use_container_width=True)
        else:
            st.info("💡 此場為徑賽決賽，道次將在初賽成績錄入完畢後，由系統自動依種子規則 (3,4,5,6,7,2,1,8) 自動產生！")

# ----------------------------------------------------
# 模組 4：一鍵匯出 Word 秩序冊與裁判執法表 (只匯出 100, 200, 400 項目)
# ----------------------------------------------------
elif menu == "4. 一鍵匯出 Word 秩序冊":
    st.title("📄 匯出全校比賽程序秩序冊與裁判執法表")
    st.write(f"當前大會抬頭：**{cfg['school_name']}** ｜ **{cfg['session_title']}**")

    # 1. 大會秩序冊 (87 場全套)
    st.subheader("📘 大會程序秩序冊（全校 87 場次完整編排）")
    st.info("💡 **全自動模式**：系統彙整全校各班名冊，自動完成 87 場次分組排線，生成完整 Word 秩序冊。")
    out_doc_name = f"{cfg['school_name']}_{cfg['session_title']}_陸運會程序.docx"

    if st.button("🚀 一鍵自動排線並生成 Word 秩序冊", type="primary"):
        with st.spinner("正在自動排線並編排 87 場次 Word 秩序冊中，請稍候..."):
            auto_events = auto_compile_events_from_database()
            build_word_document(auto_events, cfg["school_name"], cfg["session_title"], out_doc_name)
            with open(out_doc_name, "rb") as fp: st.session_state["generated_docx_bytes"] = fp.read()
            st.success("🎉 Word 秩序冊已成功全自動編排完成！請點擊下方按鈕下載：")

    if "generated_docx_bytes" in st.session_state:
        st.download_button(
            label="📥 點擊下載完整 Word 秩序冊 (.docx)",
            data=st.session_state["generated_docx_bytes"],
            file_name=out_doc_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    st.write("---")
    # 2. 裁判執法夾板記錄表 (依規程：只匯出 100, 200, 400 項目)
    st.subheader("📋 裁判執法夾板記錄表（賽前批次匯出：100 / 200 / 400 項目）")
    st.info(
        "💡 **賽前夾板列印專用**：依賽會規程，此處**僅匯出已確定選手之 100系列（初賽）、200系列（田賽／200米）及 400系列（接力）**。\n"
        "🚨 **300系列（徑賽決賽）**因需視現場初賽晉級結果排定種子道次，**請於『模組 5：現場成績錄入』內即時生成與列印！**"
    )

    all_sheets_doc_name = f"{cfg['school_name']}_{cfg['session_title']}_裁判執法記錄表(100_200_400項目).docx"

    if st.button("🖨️ 一鍵生成賽前裁判記錄表 (100, 200, 400 項目獨立分頁)", type="primary"):
        with st.spinner("正在篩選並排版 100、200、400 項目裁判記錄表..."):
            all_events = auto_compile_events_from_database()
            # 依要求只過濾 100、200、400 項目
            filtered_events = [ev for ev in all_events if int(ev["id"]) // 100 in [1, 2, 4]]
            build_score_sheets_doc(filtered_events, cfg["school_name"], cfg["session_title"], all_sheets_doc_name)
            with open(all_sheets_doc_name, "rb") as fp: st.session_state["generated_all_sheets_docx"] = fp.read()
            st.success("🎉 100、200、400 項目裁判記錄表已成功生成！請點擊下方按鈕下載：")

    if "generated_all_sheets_docx" in st.session_state:
        st.download_button(
            label="📥 下載裁判執法記錄表套裝 (100, 200, 400 項目) (.docx)",
            data=st.session_state["generated_all_sheets_docx"],
            file_name=all_sheets_doc_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ----------------------------------------------------
# 模組 5：現場成績錄入、缺席登記與班名次統計 (300項目在此處理列印)
# ----------------------------------------------------
elif menu == "5. 現場成績錄入與班名次統計":
    st.title("⏱️ 現場成績錄入、初賽晉級與班名次總榜")
    t_data = get_tournament_results()
    if "absences" not in t_data: t_data["absences"] = []
    if "qualifiers" not in t_data: t_data["qualifiers"] = {}

    tab_track, tab_field, tab_absent, tab_standing = st.tabs([
        "🏃 徑賽成績錄入 (初賽分組選拔 / 決賽評分)",
        "🎯 田賽成績錄入 (每人2次機會)",
        "📋 學生缺席與棄權登記台",
        "🏆 團體獎總榜 (級際 ＋ 初/高級組錦標)"
    ])

    with tab_track:
        track_ids = [eid for eid in sorted(MASTER_EVENTS.keys()) if (101 <= eid <= 124) or (225 <= eid <= 228) or (361 <= eid <= 384) or (485 <= eid <= 487)]
        chosen_track = st.selectbox("🎯 選擇徑賽場次", [f"{eid} ｜ {MASTER_EVENTS[eid]}" for eid in track_ids], key="track_sel")
        t_eid = int(chosen_track.split(" ｜ ")[0])
        t_ename = chosen_track.split(" ｜ ")[1]
        t_g_num = get_grade_digit(t_ename)

        is_heat = (101 <= t_eid <= 124)
        final_target_eid = t_eid + 260 if is_heat else None

        if is_heat:
            st.subheader(f"🏃 【場次 {t_eid} 初賽】{t_ename} 成績錄入（依組別排列）")
            st.info(
                f"💡 **初賽規則**：**不予評分**！請在各組別標籤中鍵入各組選手秒數（例: `1654` $\rightarrow$ `16.54秒`）。\n"
                f"全部組別輸入完畢後，點擊下方「跨組別綜合排名」按鈕，系統將自動按時間全場總排行，並將前 8 名依種子線道 (3,4,5,6,7,2,1,8) 推入【場次 {final_target_eid} 決賽】。"
            )

            sch_track = load_custom_schedule(t_eid)
            if sch_track is None: sch_track = generate_auto_track_schedule(t_eid, t_ename)

            sch_track = sch_track.sort_values(by=["組別", "線道"]).reset_index(drop=True)
            heats = sorted(sch_track["組別"].unique())

            heat_input_tabs = st.tabs([f"🏃 {h}" for h in heats])
            heat_dfs = []

            for h_idx, h_name in enumerate(heats):
                with heat_input_tabs[h_idx]:
                    st.caption(f"🏁 此為【{h_name}】出賽選手名冊，請輸入終點時間：")
                    sub_heat_df = sch_track[sch_track["組別"] == h_name][["組別", "線道", "班別", "學號", "姓名"]].copy()
                    sub_heat_df["成績快打"] = ""

                    ed_heat = st.data_editor(
                        sub_heat_df,
                        column_config={
                            "組別": st.column_config.TextColumn(disabled=True),
                            "線道": st.column_config.NumberColumn(disabled=True),
                            "班別": st.column_config.TextColumn(disabled=True),
                            "學號": st.column_config.TextColumn(disabled=True),
                            "姓名": st.column_config.TextColumn(disabled=True),
                            "成績快打": st.column_config.TextColumn("成績輸入 (例: 1654)", required=False)
                        },
                        hide_index=True,
                        use_container_width=True,
                        key=f"heat_runner_ed_{t_eid}_{h_name}"
                    )
                    heat_dfs.append(ed_heat)

            st.write("---")
            if st.button("⚡ 跨組別綜合排名並選拔前 8 名（依 3,4,5,6,7,2,1,8 排入決賽）", type="primary", key=f"heat_btn_{t_eid}"):
                all_entered_runners = pd.concat(heat_dfs, ignore_index=True)
                parsed_rows = []
                for _, r in all_entered_runners.iterrows():
                    sec_v, disp_s = parse_track_time(r["成績快打"])
                    parsed_rows.append({
                        "組別": r["組別"], "線道": r["線道"],
                        "班別": r["班別"], "學號": r["學號"], "姓名": r["姓名"],
                        "秒數": sec_v, "大會成績": disp_s
                    })
                
                res_df = pd.DataFrame(parsed_rows)
                v_runners = res_df[res_df["秒數"].notna()].sort_values(by="秒數").reset_index(drop=True)
                v_runners["名次"] = v_runners.index + 1

                def get_heat_status_and_lane(rank):
                    if rank <= 8:
                        assigned_lane = FINAL_SEED_LANES[rank - 1]
                        return f"🟢 晉級決賽 (第{assigned_lane}線)", assigned_lane
                    elif rank <= 10: return f"🟡 候補選手 R{rank-8}", None
                    else: return "⚪ 初賽淘汰", None

                statuses = []
                assigned_lanes = []
                for r in v_runners["名次"]:
                    st_str, ln_val = get_heat_status_and_lane(r)
                    statuses.append(st_str)
                    assigned_lanes.append(ln_val)

                v_runners["晉級狀態與決賽線道"] = statuses
                v_runners["決賽線道"] = assigned_lanes
                v_runners["備註"] = "初賽不評分"

                iv_runners = res_df[res_df["秒數"].isna()].copy()
                iv_runners["名次"] = "-"
                iv_runners["晉級狀態與決賽線道"] = "❌ 缺席/犯規 (無資格)"
                iv_runners["決賽線道"] = None
                iv_runners["備註"] = "初賽不評分"

                final_heat_table = pd.concat([v_runners, iv_runners], ignore_index=True)

                st.write(f"### 📋 【場次 {t_eid} 初賽】跨組別全場總排名與晉級次賽名冊：")
                st.dataframe(
                    final_heat_table[["名次", "組別", "線道", "班別", "學號", "姓名", "大會成績", "晉級狀態與決賽線道", "備註"]],
                    hide_index=True,
                    use_container_width=True
                )

                top8_qualifiers = []
                for _, row in v_runners[v_runners["名次"] <= 8].iterrows():
                    top8_qualifiers.append({
                        "rank": int(row["名次"]),
                        "lane": int(row["決賽線道"]),
                        "class": str(row["班別"]),
                        "name": str(row["姓名"]),
                        "heat_time": str(row["大會成績"])
                    })

                t_data["qualifiers"][str(final_target_eid)] = top8_qualifiers
                save_tournament_results(t_data)
                st.success(f"🎉 初賽前 8 名選手已自動完成跨組別排序，並依種子線道 (3,4,5,6,7,2,1,8) 鎖定推入【場次 {final_target_eid} 決賽】！")

        else:
            # 300 系列徑賽決賽處理區：增加即時列印按鈕
            c_f_head, c_f_print = st.columns([3, 1.2])
            with c_f_head:
                st.subheader(f"🏁 【場次 {t_eid} 決賽】{t_ename} 成績錄入與給分")
            with c_f_print:
                # 300 項目在此即時生成專屬裁判記錄紙 (印有 8 強姓名與種子道次)
                final_sheet_name = f"決賽裁判記錄表_場次{t_eid}_{t_ename}.docx"
                if st.button(f"🖨️ 產生【場次 {t_eid}】決賽裁判記錄紙", key=f"print_f_{t_eid}"):
                    final_ev_data = [ev for ev in auto_compile_events_from_database() if str(ev["id"]) == str(t_eid)]
                    if final_ev_data:
                        build_score_sheets_doc(final_ev_data, cfg["school_name"], cfg["session_title"], final_sheet_name)
                        with open(final_sheet_name, "rb") as fp:
                            st.download_button(
                                label=f"📥 下載場次 {t_eid} 決賽記錄紙",
                                data=fp.read(),
                                file_name=final_sheet_name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"dl_btn_{t_eid}"
                            )

            st.info(f"💡 此場為【決賽】，第 1 至 8 名將依大會規程核發 **9, 7, 6, 5, 4, 3, 2, 1 分**，並即時累計至全校班名次總榜。")

            qual_list = t_data.get("qualifiers", {}).get(str(t_eid), [])

            if qual_list:
                st.success(f"📋 系統已自動載入從初賽脫穎而出的 8 強選手，道次已按種子規則 (3,4,5,6,7,2,1,8) 排定！")
                initial_runners = []
                sorted_qual = sorted(qual_list, key=lambda x: x.get("lane", 1))
                for q in sorted_qual:
                    initial_runners.append({
                        "線道": int(q["lane"]), "班別": q["class"], "姓名": q["name"],
                        "種子資訊": f"初賽第{q['rank']}名 ({q['heat_time']}秒)", "成績快打": ""
                    })
            else:
                sch_track = load_custom_schedule(t_eid)
                if sch_track is not None:
                    initial_runners = []
                    for _, r in sch_track.iterrows():
                        initial_runners.append({"線道": int(r["線道"]), "班別": str(r["班別"]), "姓名": str(r["姓名"]), "種子資訊": "決賽排線", "成績快打": ""})
                else:
                    initial_runners = [
                        {"線道": 3, "班別": f"{t_g_num}A", "姓名": "選手甲", "種子資訊": "決賽排線", "成績快打": "1532"},
                        {"線道": 4, "班別": f"{t_g_num}B", "姓名": "選手乙", "種子資訊": "決賽排線", "成績快打": "1666"},
                        {"線道": 5, "班別": f"{t_g_num}C", "姓名": "選手丙", "種子資訊": "決賽排線", "成績快打": "1022"},
                        {"線道": 6, "班別": f"{t_g_num}D", "姓名": "選手丁", "種子資訊": "決賽排線", "成績快打": "1556"},
                        {"線道": 7, "班別": f"{t_g_num}E", "姓名": "選手戊", "種子資訊": "決賽排線", "成績快打": "1855"},
                    ]

            df_track_in = st.data_editor(
                pd.DataFrame(initial_runners),
                column_config={
                    "線道": st.column_config.NumberColumn(disabled=True),
                    "班別": st.column_config.TextColumn(disabled=True),
                    "姓名": st.column_config.TextColumn(disabled=True),
                    "種子資訊": st.column_config.TextColumn(disabled=True),
                    "成績快打": st.column_config.TextColumn("成績輸入 (例: 1654 即 16.54秒)", required=False)
                },
                hide_index=True,
                use_container_width=True,
                key=f"final_te_editor_{t_eid}"
            )

            if st.button("⚡ 解析決賽成績並核算名次積分 (9-1分)", type="primary", key=f"final_btn_{t_eid}"):
                parsed_rows = []
                for _, r in df_track_in.iterrows():
                    sec_v, disp_s = parse_track_time(r["成績快打"])
                    parsed_rows.append({"線道": r["線道"], "班別": r["班別"], "姓名": r["姓名"], "秒數": sec_v, "大會成績": disp_s})
                res_df = pd.DataFrame(parsed_rows)
                v_runners = res_df[res_df["秒數"].notna()].sort_values(by="秒數").reset_index(drop=True)
                v_runners["名次"] = v_runners.index + 1
                v_runners["得分"] = v_runners["名次"].map(POINT_SYSTEM).fillna(0).astype(int)

                iv_runners = res_df[res_df["秒數"].isna()].copy()
                iv_runners["名次"] = "-"
                iv_runners["得分"] = 0
                final_track_res = pd.concat([v_runners, iv_runners], ignore_index=True)

                st.write(f"### 🏆 【場次 {t_eid} 決賽】最終成績與團體積分榜：")
                st.dataframe(final_track_res[["名次", "線道", "班別", "姓名", "大會成績", "得分"]], hide_index=True, use_container_width=True)

                t_data["results"] = [r for r in t_data["results"] if str(r.get("event_id")) != str(t_eid)]
                for _, row in final_track_res[final_track_res["得分"] > 0].iterrows():
                    t_data["results"].append({
                        "event_id": t_eid, "event_name": t_ename,
                        "rank": int(row["名次"]), "class": str(row["班別"]).strip().upper(),
                        "points": int(row["得分"])
                    })
                save_tournament_results(t_data)
                st.success(f"🎉 【場次 {t_eid} 決賽】成績已確認，團體分數已即時累計入全校班名次榜！")

    with tab_field:
        st.subheader("🎯 田賽決賽成績錄入 (每人2次試投/跳機會)")
        st.info("📋 **規則指引**：田賽直接為決賽，每人 2 次試跳／試擲機會。系統取最佳成績判定名次，前 8 名給予 9 至 1 分。")
        field_ids = [eid for eid in sorted(MASTER_EVENTS.keys()) if 229 <= eid <= 260]
        chosen_field = st.selectbox("🎯 選擇田賽場次", [f"{eid} ｜ {MASTER_EVENTS[eid]}" for eid in field_ids], key="field_sel")
        f_eid = int(chosen_field.split(" ｜ ")[0])
        f_ename = chosen_field.split(" ｜ ")[1]
        f_g_num = get_grade_digit(f_ename)

        sch_field = load_custom_schedule(f_eid)
        if sch_field is not None:
            sample_field_roster = []
            for _, r in sch_field.iterrows():
                if r["姓名"] != "（尚無報名）":
                    sample_field_roster.append({
                        "出場次序": int(r["出場次序"]), "班別": str(r["班別"]), "學號": str(r["學號"]), "姓名": str(r["姓名"]),
                        "第一次": "", "第二次": ""
                    })
        else:
            sample_field_roster = [
                {"出場次序": 1, "班別": f"{f_g_num}A", "學號": "03", "姓名": "陳芷澄", "第一次": "2.35", "第二次": "2.48"},
                {"出場次序": 2, "班別": f"{f_g_num}A", "學號": "12", "姓名": "林秀美", "第一次": "2.48", "第二次": "2.40"},
                {"出場次序": 3, "班別": f"{f_g_num}B", "學號": "05", "姓名": "劉曦儀", "第一次": "2.10", "第二次": "X"},
                {"出場次序": 4, "班別": f"{f_g_num}B", "學號": "08", "姓名": "彭一心", "第一次": "X",    "第二次": "2.55"},
                {"出場次序": 5, "班別": f"{f_g_num}C", "學號": "02", "姓名": "朱迪",   "第一次": "-",    "第二次": "-"},
            ]

        df_field_in = st.data_editor(
            pd.DataFrame(sample_field_roster),
            column_config={
                "出場次序": st.column_config.NumberColumn(disabled=True),
                "班別": st.column_config.TextColumn(disabled=True),
                "學號": st.column_config.TextColumn(disabled=True),
                "姓名": st.column_config.TextColumn(disabled=True),
                "第一次": st.column_config.TextColumn("第 1 次試投/跳", required=False),
                "第二次": st.column_config.TextColumn("第 2 次試投/跳", required=False),
            },
            hide_index=True,
            use_container_width=True,
            key=f"fe_{f_eid}"
        )

        if st.button("⚡ 計算田賽最佳成績並計入班名次", type="primary", key=f"fbtn_{f_eid}"):
            res_field = process_field_event_results(df_field_in)
            st.dataframe(
                res_field[["名次", "出場次序", "班別", "學號", "姓名", "第一次", "第二次", "最佳成績", "得分"]],
                hide_index=True,
                use_container_width=True
            )

            t_data["results"] = [r for r in t_data["results"] if str(r.get("event_id")) != str(f_eid)]
            for _, row in res_field[res_field["得分"] > 0].iterrows():
                t_data["results"].append({
                    "event_id": f_eid, "event_name": f_ename,
                    "rank": int(row["名次"]), "class": str(row["班別"]).strip().upper(),
                    "points": int(row["得分"])
                })
            save_tournament_results(t_data)
            st.success(f"【場次 {f_eid}】{f_ename} 田賽成績與積分已成功累計入班名次榜！")

    with tab_absent:
        st.subheader("📋 學生缺席與無理棄權登記台")
        st.markdown(
            """
            * 🟢 **普通缺席（病假／事假）**：**不扣減分數**，但**減少該班參與人數（分母減 1）**，避免拉低班級平均分。
            * 🔴 **無理棄權（有到校但無故缺賽）**：依規程 4.3 **扣減團體總分 1 分**，參與人數（分母）不減少。
            """
        )
        c_a1, c_a2 = st.columns(2)
        with c_a1:
            absent_type = st.radio("選擇缺席性質", ["普通缺席 (不扣分，參與人數減 1)", "無理棄權 (扣減 1 分，參與人數不變)"], horizontal=True)
            f_class = st.text_input("學生班別 (例: 1A)", key="fc")
            f_sid = st.text_input("學號 / 姓名", key="fs")
        with c_a2:
            all_opts = ["-- 全日請假 --"] + [f"{eid} ｜ {m}" for eid, m in sorted(MASTER_EVENTS.items())]
            f_event_str = st.selectbox("相關項目 / 假別", all_opts, key="fe_box")
            f_reason = st.text_input("備註 / 原因 (例: 醫生證明發燒、無故未點名檢錄)", key="fr")

        if st.button("📝 確認登記缺席／棄權記錄", type="primary"):
            if f_class.strip():
                is_unexcused = ("無理棄權" in absent_type)
                t_data["absences"].append({
                    "class": f_class.strip().upper(),
                    "sid_name": f_sid.strip(),
                    "event": f_event_str,
                    "type": "無理棄權" if is_unexcused else "普通缺席",
                    "penalty": 1 if is_unexcused else 0,
                    "deduct_headcount": 0 if is_unexcused else 1,
                    "reason": f_reason.strip()
                })
                save_tournament_results(t_data)
                if is_unexcused: st.warning(f"已記錄 {f_class.upper()} 學生【{f_sid}】無理棄權，該班扣減 1 分！")
                else: st.success(f"已記錄 {f_class.upper()} 學生【{f_sid}】普通缺席，不扣分，該班參與人數扣減 1 人！")
            else:
                st.error("請填寫學生班別！")

        if t_data["absences"]:
            st.write("#### 📋 目前已登記之缺席／棄權名冊：")
            df_ab_display = pd.DataFrame(t_data["absences"])
            st.dataframe(df_ab_display[["class", "sid_name", "event", "type", "penalty", "deduct_headcount", "reason"]], use_container_width=True)
            if st.button("🗑️ 清空所有缺席／棄權記錄"):
                t_data["absences"] = []
                save_tournament_results(t_data)
                st.rerun()

    with tab_standing:
        st.subheader("🏆 團體獎總榜（級際 ＋ 初級組 P1–3 / 高級組 P4–6 錦標）")
        if "df_students" not in st.session_state:
            st.warning("⚠️ 請先在「模組 1」上傳學生名冊以獲取各班實際人數！")
        else:
            df_stu = st.session_state["df_students"]
            all_classes = sorted(df_stu["班別"].unique())
            orig_counts = df_stu.groupby("班別").size().to_dict()

            class_scores = {c: 0 for c in all_classes}
            for res in t_data["results"]:
                c = res["class"]
                if c in class_scores: class_scores[c] += res["points"]

            class_penalties = {c: 0 for c in all_classes}
            class_excused_deduct = {c: 0 for c in all_classes}

            for ab in t_data["absences"]:
                c = ab["class"]
                if c in all_classes:
                    class_penalties[c] += ab.get("penalty", 0)
                    class_excused_deduct[c] += ab.get("deduct_headcount", 0)

            def get_division_group(class_name):
                g = int(str(class_name)[0]) if str(class_name)[0].isdigit() else 1
                return "初級組 (P1-3)" if g in [1, 2, 3] else "高級組 (P4-6)"

            summary_list = []
            for c in all_classes:
                raw_pt = class_scores[c]
                deduct_pt = class_penalties[c]
                net_pt = raw_pt - deduct_pt

                orig_head = orig_counts.get(c, 1)
                excused_cnt = class_excused_deduct[c]
                final_head = max(1, orig_head - excused_cnt)
                avg_score = round(net_pt / final_head, 3)

                summary_list.append({
                    "班別": c,
                    "年級": f"{c[0]}年級",
                    "組別": get_division_group(c),
                    "得分": raw_pt,
                    "棄權扣減": deduct_pt,
                    "實得總分": net_pt,
                    "全班原有人數": orig_head,
                    "普通缺席人數": excused_cnt,
                    "實際參賽人數": final_head,
                    "平均分數": avg_score
                })

            df_standing = pd.DataFrame(summary_list)
            df_standing["級排名"] = df_standing.groupby("年級")["平均分數"].rank(method="min", ascending=False).astype(int)
            df_standing["組排名"] = df_standing.groupby("組別")["平均分數"].rank(method="min", ascending=False).astype(int)

            st_tab_grade, st_tab_junior, st_tab_senior, st_tab_all = st.tabs([
                "🏅 級際團體榜 (P1–P6)",
                "🟢 初級組錦標 (P1–3)",
                "🟣 高級組錦標 (P4–6)",
                "📊 全校 24 班審核總表"
            ])

            show_cols = [
                "班別", "年級", "組別", "得分", "棄權扣減", "實得總分",
                "全班原有人數", "普通缺席人數", "實際參賽人數",
                "平均分數", "級排名", "組排名"
            ]

            with st_tab_grade:
                st.markdown("##### 🏅 級際團體名次（各年級冠、亞、季軍）")
                for g_num in range(1, 7):
                    g_label = f"{g_num}年級"
                    sub_g = df_standing[df_standing["年級"] == g_label].sort_values(by="級排名")
                    with st.expander(f"📌 {g_label} 級際排行（第 1 名：{sub_g.iloc[0]['班別']}，平均分：{sub_g.iloc[0]['平均分數']}）", expanded=(g_num==1)):
                        st.dataframe(sub_g[show_cols], hide_index=True, use_container_width=True)

            with st_tab_junior:
                st.markdown("##### 🟢 初級組團體錦標 (P1–3)（前三名：冠軍、亞軍、季軍）")
                df_jun_view = df_standing[df_standing["組別"] == "初級組 (P1-3)"].sort_values(by="組排名")
                st.dataframe(df_jun_view[show_cols], hide_index=True, use_container_width=True)

            with st_tab_senior:
                st.markdown("##### 🟣 高級組團體錦標 (P4–6)（前三名：冠軍、亞軍、季軍）")
                df_sen_view = df_standing[df_standing["組別"] == "高級組 (P4-6)"].sort_values(by="組排名")
                st.dataframe(df_sen_view[show_cols], hide_index=True, use_container_width=True)

            with st_tab_all:
                st.markdown("##### 📊 全校 24 班詳細審核總表（按年級排序）")
                df_all_view = df_standing.sort_values(by=["年級", "級排名"])
                st.dataframe(df_all_view[show_cols], hide_index=True, use_container_width=True)

            st.write("---")
            c_csv, c_word_awards = st.columns(2)
            with c_csv:
                csv_data = df_standing.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 匯出級際及分組團體獎統計報表 (CSV/Excel)",
                    data=csv_data,
                    file_name=f"{cfg['school_name']}_陸運會團體成績總表.csv",
                    mime="text/csv",
                    use_container_width=True
                )

            with c_word_awards:
                awards_doc_name = f"{cfg['school_name']}_{cfg['session_title']}_團體獎頒獎名冊(級際及初高級錦標).docx"
                if st.button("🖨️ 產生正式「團體獎頒獎名冊（級際及初高級錦標）」(.docx)", type="primary", use_container_width=True):
                    with st.spinner("正在排版司儀榮譽榜與級際錦標表..."):
                        build_awards_doc(df_standing, cfg["school_name"], cfg["session_title"], awards_doc_name)
                        with open(awards_doc_name, "rb") as fp:
                            st.session_state["generated_awards_docx"] = fp.read()
                        st.success("🎉 團體獎頒獎名冊已成功生成！請點擊下方按鈕下載：")

            if "generated_awards_docx" in st.session_state:
                st.download_button(
                    label="📥 下載團體獎頒獎名冊 (.docx)",
                    data=st.session_state["generated_awards_docx"],
                    file_name=f"{cfg['school_name']}_{cfg['session_title']}_團體獎頒獎名冊.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

# ----------------------------------------------------
# 模組 6：大會全域設定
# ----------------------------------------------------
elif menu == "⚙️ 大會全域設定":
    st.title("⚙️ 體育科主任：大會全域設定")
    s_name = st.text_input("學校全銜", value=cfg["school_name"])
    s_sess = st.text_input("運動會屆別／年度", value=cfg["session_title"])
    min_b = st.number_input("接力正選男生預設最低人數", min_value=0, max_value=3, value=cfg.get("relay_min_boys", 1))
    min_g = st.number_input("接力正選女生預設最低人數", min_value=0, max_value=3, value=cfg.get("relay_min_girls", 1))
    
    if st.button("💾 儲存全域設定", type="primary"):
        save_config({
            "school_name": s_name,
            "session_title": s_sess,
            "relay_min_boys": int(min_b),
            "relay_min_girls": int(min_g)
        })
        st.success("大會設定已更新！全系統即時生效。")