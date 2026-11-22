"""워드(.docx) 문서 만들기

app.py에서 나눈 모듈입니다. 아래 층(dogae.base, dogae.ai, dogae.storage, dogae.exam_data, dogae.questions)만 가져다 씁니다."""
import sys
import io
import streamlit as st
import ast
import hmac
import math
import shutil
import tempfile
import time
import pymupdf
import os
import re
import difflib
import sqlite3
import json
import uuid
import datetime
import requests
import zipfile
import pandas as pd
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import hashlib as _hashlib
import secrets as _pysecrets
import wave as _wave
from dogae.base import *  # noqa: F401,F403
from dogae.ai import *  # noqa: F401,F403
from dogae.storage import *  # noqa: F401,F403
from dogae.exam_data import *  # noqa: F401,F403
from dogae.questions import *  # noqa: F401,F403


# -------------------------------------------------------------------------
# [2] 워드 표 레이아웃 및 디자인 무너짐 완벽 방지 엔진
# -------------------------------------------------------------------------
def set_document_font(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = '맑은 고딕'
    font.size = Pt(11)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

def set_cell_background(cell, fill_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=200, right=200):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_parsed_text_to_cell(cell, text):
    set_cell_margins(cell)
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    parts = text.split("**")
    for i, part in enumerate(parts):
        run = p.add_run(part)
        if i % 2 != 0:
            run.bold = True
            run.font.color.rgb = RGBColor(0, 51, 102)

def create_word_files(content, student_name, interview_type, target_desc):
    type_label = "생기부면접" if "생기부" in interview_type else "제시문면접"
    is_jesimun = "제시문" in interview_type
    doc_student = Document()
    doc_teacher = Document()

    for doc, is_teacher in [(doc_student, False), (doc_teacher, True)]:
        set_document_font(doc)
        title_text = f"🎓 [{student_name}] {target_desc} 도개고 맞춤 모의면접 {'지침서 (교사용)' if is_teacher else '워크북 (학생용)'}"
        doc.add_heading(title_text, level=1)
        doc.add_paragraph(f"[{type_label}] 본 문서는 도개고등학교 진로진학 지도 기준에 맞춰 생성되었습니다.\n")

        briefing_match = re.search(r'(### 🔍 \[생기부 심층 분석 브리핑 리포트\].*?)(?=### 📌|### 📄|$)', content, re.DOTALL)
        if briefing_match:
            add_intro_paragraphs(doc, briefing_match.group(1).strip())
            doc.add_paragraph()

        blocks = content.split('### 📌')
        for block in blocks[1:]:
            lines = block.strip().split('\n')
            if not lines: continue

            title = lines[0].strip()
            body = "\n".join(lines[1:])
            table = doc.add_table(rows=0, cols=1)
            table.style = 'Table Grid'
            for row in table.rows:
                trPr = row._tr.get_or_add_trPr()
                trPr.append(OxmlElement('w:cantSplit'))

            row_title = table.add_row()
            cell_title = row_title.cells[0]
            set_cell_background(cell_title, "EBF1FA")
            add_parsed_text_to_cell(cell_title, f"📌 {title}")

            if is_jesimun:
                rc_match = re.search(r'\[실제 기출\](.*?)(?=\[제시문\]|\[문제 1\]|$)', body, re.DOTALL)
                js_match = re.search(r'\[제시문\](.*?)(?=\[문제 1\]|$)', body, re.DOTALL)
                q1_match = re.search(r'\[문제 1\](.*?)(?=\[평가요소 1\]|\[평가의도 1\]|\[문제 2\]|$)', body, re.DOTALL)
                e1_match = re.search(r'\[평가요소 1\](.*?)(?=\[평가의도 1\]|$)', body, re.DOTALL)
                i1_match = re.search(r'\[평가의도 1\](.*?)(?=\[모범답안 1\]|$)', body, re.DOTALL)
                a1_match = re.search(r'\[모범답안 1\](.*?)(?=\[꼬리질문 1\]|$)', body, re.DOTALL)
                f1_match = re.search(r'\[꼬리질문 1\](.*?)(?=\[문제 2\]|$)', body, re.DOTALL)

                q2_match = re.search(r'\[문제 2\](.*?)(?=\[평가요소 2\]|\[평가의도 2\]|$)', body, re.DOTALL)
                e2_match = re.search(r'\[평가요소 2\](.*?)(?=\[평가의도 2\]|$)', body, re.DOTALL)
                i2_match = re.search(r'\[평가의도 2\](.*?)(?=\[모범답안 2\]|$)', body, re.DOTALL)
                a2_match = re.search(r'\[모범답안 2\](.*?)(?=\[꼬리질문 2\]|$)', body, re.DOTALL)
                f2_match = re.search(r'\[꼬리질문 2\](.*?)(?=$)', body, re.DOTALL)

                js_text = js_match.group(1).strip() if js_match else ""
                q1_text = q1_match.group(1).strip() if q1_match else ""
                e1_text = e1_match.group(1).strip() if e1_match else ""
                i1_text = i1_match.group(1).strip() if i1_match else ""
                a1_text = a1_match.group(1).strip() if a1_match else ""
                f1_text = f1_match.group(1).strip() if f1_match else ""

                q2_text = q2_match.group(1).strip() if q2_match else ""
                e2_text = e2_match.group(1).strip() if e2_match else ""
                i2_text = i2_match.group(1).strip() if i2_match else ""
                a2_text = a2_match.group(1).strip() if a2_match else ""
                f2_text = f2_match.group(1).strip() if f2_match else ""

                rc_text = rc_match.group(1).strip() if rc_match else ""
                if rc_text:
                    row_rc = table.add_row()
                    set_cell_background(row_rc.cells[0], "FFF8E1")
                    add_parsed_text_to_cell(row_rc.cells[0], f"**[📚 실제 기출 사례]**\n{rc_text}")

                if js_text:
                    row_js = table.add_row()
                    set_cell_background(row_js.cells[0], "F4F6F9")
                    add_parsed_text_to_cell(row_js.cells[0], f"**[응용 제시문]**\n{js_text}")

                row_q1 = table.add_row()
                add_parsed_text_to_cell(row_q1.cells[0], f"**[문제 1]**\n{q1_text}")
                if is_teacher:
                    if e1_text:
                        row_e1 = table.add_row()
                        set_cell_background(row_e1.cells[0], "F3F0FA")
                        add_parsed_text_to_cell(row_e1.cells[0], f"**[🏷 평가 요소 1]**\n{e1_text}")
                    row_i1 = table.add_row()
                    set_cell_background(row_i1.cells[0], "F9F9F9")
                    add_parsed_text_to_cell(row_i1.cells[0], f"**[평가 의도 1]**\n{i1_text}")
                    row_a1 = table.add_row()
                    add_parsed_text_to_cell(row_a1.cells[0], f"**[모범 답안 가이드 1]**\n{a1_text}")
                    row_f1 = table.add_row()
                    set_cell_background(row_f1.cells[0], "FFF4F4")
                    add_parsed_text_to_cell(row_f1.cells[0], f"**[압박용 꼬리질문 1]**\n{f1_text}")

                if q2_text:
                    row_q2 = table.add_row()
                    add_parsed_text_to_cell(row_q2.cells[0], f"**[문제 2]**\n{q2_text}")
                    if is_teacher:
                        if e2_text:
                            row_e2 = table.add_row()
                            set_cell_background(row_e2.cells[0], "F3F0FA")
                            add_parsed_text_to_cell(row_e2.cells[0], f"**[🏷 평가 요소 2]**\n{e2_text}")
                        row_i2 = table.add_row()
                        set_cell_background(row_i2.cells[0], "F9F9F9")
                        add_parsed_text_to_cell(row_i2.cells[0], f"**[평가 의도 2]**\n{i2_text}")
                        row_a2 = table.add_row()
                        add_parsed_text_to_cell(row_a2.cells[0], f"**[모범 답안 가이드 2]**\n{a2_text}")
                        row_f2 = table.add_row()
                        set_cell_background(row_f2.cells[0], "FFF4F4")
                        add_parsed_text_to_cell(row_f2.cells[0], f"**[압박용 꼬리질문 2]**\n{f2_text}")
            else:
                q_match = re.search(r'\[질문\](.*?)(?=\[평가요소\]|\[평가의도\]|\[모범답안\]|\[꼬리질문\]|$)', body, re.DOTALL)
                e_match = re.search(r'\[평가요소\](.*?)(?=\[평가의도\]|\[모범답안\]|\[꼬리질문\]|$)', body, re.DOTALL)
                i_match = re.search(r'\[평가의도\](.*?)(?=\[모범답안\]|\[꼬리질문\]|$)', body, re.DOTALL)
                a_match = re.search(r'\[모범답안\](.*?)(?=\[꼬리질문\]|$)', body, re.DOTALL)
                f_match = re.search(r'\[꼬리질문\](.*?)(?=$)', body, re.DOTALL)

                q_text = q_match.group(1).strip() if q_match else "내용 없음"
                e_text = e_match.group(1).strip() if e_match else ""
                i_text = i_match.group(1).strip() if i_match else ""
                a_text = a_match.group(1).strip() if a_match else ""
                f_text = f_match.group(1).strip() if f_match else ""

                row_q = table.add_row()
                add_parsed_text_to_cell(row_q.cells[0], f"**[면접 질문]**\n{q_text}")

                if is_teacher:
                    if e_text:
                        row_e = table.add_row()
                        set_cell_background(row_e.cells[0], "F3F0FA")
                        add_parsed_text_to_cell(row_e.cells[0], f"**[🏷 평가 요소]**\n{e_text}")
                    row_i = table.add_row()
                    set_cell_background(row_i.cells[0], "F9F9F9")
                    add_parsed_text_to_cell(row_i.cells[0], f"**[평가 의도]**\n{i_text}")
                    row_a = table.add_row()
                    add_parsed_text_to_cell(row_a.cells[0], f"**[모범 답안 가이드]**\n{a_text}")
                    row_f = table.add_row()
                    set_cell_background(row_f.cells[0], "FFF4F4")
                    add_parsed_text_to_cell(row_f.cells[0], f"**[압박용 꼬리질문]**\n{f_text}")
            doc.add_paragraph()

    student_path = _out_path(f"{student_name}_{target_desc}_{type_label}_학생용.docx")
    teacher_path = _out_path(f"{student_name}_{target_desc}_{type_label}_교사용.docx")
    doc_student.save(student_path)
    doc_teacher.save(teacher_path)
    return student_path, teacher_path

def create_chat_history_word(chat_history, student_name):
    doc = Document()
    set_document_font(doc)
    doc.add_heading(f"💬 [{student_name}] 면접 문항 피드백 대화 내역", level=1)
    doc.add_paragraph("AI 출제위원과의 피드백 기록입니다.\n" + "="*50)
    for msg in chat_history:
        role_title = "👤 선생님/학생 (요청)" if msg["role"] == "user" else "🤖 AI 출제위원 (답변/평가)"
        doc.add_heading(role_title, level=2)
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.3
        parts = msg["content"].split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            if i % 2 != 0: run.bold = True
        doc.add_paragraph("-" * 50)
    file_path = _out_path(f"{student_name}_피드백_대화내역.docx")
    doc.save(file_path)
    return file_path

def create_easy_explanation_word(explanation_text, student_name, target_desc):
    """생기부 원문을 고1 눈높이로 해설한 내용을 워드 문서로 만듭니다."""
    doc = Document()
    set_document_font(doc)
    doc.add_heading(f"📚 [{student_name}] {target_desc} 생기부 쉬운 해설 (고등학교 1학년 눈높이)", level=1)
    doc.add_paragraph(
        "이 문서는 학생 본인이 자신의 생활기록부(생기부) 내용을 스스로 읽고 이해할 수 있도록, "
        "어려운 용어와 활동명을 고등학교 1학년 눈높이로 쉽게 풀어 설명한 자료입니다.\n"
    )
    add_intro_paragraphs(doc, explanation_text)
    file_path = _out_path(f"{student_name}_생기부_쉬운해설.docx")
    doc.save(file_path)
    return file_path

def _set_col_widths(table, widths_cm):
    """표의 열 너비를 고정합니다(워드가 자동으로 넓히지 않게)."""
    from docx.shared import Cm
    table.autofit = False
    for col, w in zip(table.columns, widths_cm):
        col.width = Cm(w)                      # 표 격자(gridCol) 너비
    for row in table.rows:
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))   # 한 행이 쪽 사이에서 잘리지 않게
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)                 # 칸(tcW) 너비

def create_highlight_word(hl, student_name, target_desc, is_teacher=False):
    """형광펜 생기부 워드 문서. 칠한 문장 바로 아래에 질문이 붙습니다.
    학생용은 질문과 자가 체크칸, 교사용은 답변 가이드와 칠한 이유까지 넣습니다."""
    from docx.shared import Cm
    try:
        unmatched_sets = attach_set_details(hl, st.session_state.get("last_result_text", ""))
    except Exception:
        unmatched_sets = []
    doc = Document()
    set_document_font(doc)
    doc.add_heading(
        f"🎓 [{student_name}] {target_desc} 모의면접 {'지침서 (교사용)' if is_teacher else '워크북 (학생용)'}", level=1
    )
    doc.add_paragraph(
        f"면접관이 짚을 만한 문장 {len(hl['items'])}곳에 역량별로 색을 칠하고, 바로 아래에 예상 질문 {hl['q_total']}개를 붙였습니다. "
        + ("질문마다 평가 의도·모범 답안·꼬리질문을 함께 넣었습니다." if is_teacher
           else "질문마다 먼저 소리 내어 답해 보고 자신 있음·애매·모름에 체크하세요. 맨 뒤에 생기부 쉬운 해설이 있습니다.")
    )
    if is_teacher:
        _bm = re.search(r'### 🔍 \[생기부 심층 분석 브리핑 리포트\](.*?)(?=### 📌|### 📄|$)',
                        str(st.session_state.get("last_result_text", "")), re.DOTALL)
        if _bm and _bm.group(1).strip():
            doc.add_heading("생기부 분석 브리핑", level=2)
            add_intro_paragraphs(doc, _bm.group(1).strip())
    legend = doc.add_paragraph()
    for k, v in HL_COMPS.items():
        r = legend.add_run(f" {v['label']} ({hl['count'].get(k, 0)}) ")
        _shade_run(r, v["fill"])
        legend.add_run("   ")
    r = legend.add_run(f"{HL_DEFEND['label']} ({hl['defend']})")
    _wavy_red(r)
    if hl.get("keywords"):
        doc.add_paragraph("핵심 키워드: " + " · ".join(hl["keywords"]))

    if hl["gaps"]:
        doc.add_heading("영역별 준비 상태", level=2)
        cols = ["영역", "칠한 문장", "질문", "상태"]
        t = doc.add_table(rows=1, cols=len(cols))
        t.style = "Table Grid"
        for i, h in enumerate(cols):
            t.rows[0].cells[i].text = h
            set_cell_background(t.rows[0].cells[i], "EBF1FA")
        for g in hl["gaps"]:
            cells = t.add_row().cells
            for i, key in enumerate(cols):
                cells[i].text = str(g[key])
            if g["상태"].startswith("⚠️"):
                set_cell_background(cells[3], "FDECEA")
        _set_col_widths(t, [4.2, 2.4, 2.0, 7.4])

    doc.add_heading("생기부 원문과 예상 질문", level=2)
    text, items = hl["text"], hl["items"]
    for pi, pa in enumerate(hl["paras"]):
        start, end = pa["start"], pa["end"]
        if start == end:
            continue
        cards = [c for c in items if c["para"] == pi]
        p = doc.add_paragraph()
        cur = start
        for c in cards:
            s, e = max(c["start"], start), min(c["end"], end)
            if s > cur:
                p.add_run(text[cur:s])
            m = p.add_run(_marker(c["no"]))
            m.font.superscript = True
            m.font.bold = True
            m.font.color.rgb = RGBColor(0xC6, 0x28, 0x28)
            r = p.add_run(text[s:e])
            _shade_run(r, HL_COMPS[c["comp"]]["fill"])
            if c["defend"]:
                _wavy_red(r)
            cur = e
        if cur < end:
            p.add_run(text[cur:end])
        for c in cards:
            for q in c["qs"]:
                qp = doc.add_paragraph()
                qp.paragraph_format.left_indent = Cm(0.8)
                qp.paragraph_format.space_after = Pt(0)
                head = qp.add_run(f"{q['no']}. {_marker(c['no'])} ")
                head.bold = True
                qp.add_run(q["q"])
                tp = doc.add_paragraph()
                tp.paragraph_format.left_indent = Cm(0.8)
                tp.paragraph_format.space_after = Pt(0 if is_teacher else 6)
                tr = tp.add_run(_q_tags(c, q) + ("" if is_teacher else "     □ 자신 있음  □ 애매  □ 모름"))
                tr.font.size = Pt(8.5)
                tr.font.color.rgb = RGBColor(0x75, 0x75, 0x75)
                if is_teacher and q.get("guide"):
                    gp = doc.add_paragraph()
                    gp.paragraph_format.left_indent = Cm(0.8)
                    gp.paragraph_format.space_after = Pt(4)
                    gr = gp.add_run("답변 가이드: " + q["guide"])
                    gr.font.size = Pt(9)
                    gr.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
                if is_teacher:
                    _det = q.get("detail") or {}
                    for key, val in (("평가의도", _det.get("평가의도")), ("모범답안", answer_of(q)),
                                     ("꼬리질문", _det.get("꼬리질문"))):
                        if val:
                            dp = doc.add_paragraph()
                            dp.paragraph_format.left_indent = Cm(0.8)
                            dp.paragraph_format.space_after = Pt(4)
                            dr = dp.add_run(f"{key}: " + re.sub(r"\s+", " ", val))
                            dr.font.size = Pt(9)
            if is_teacher and c.get("reason"):
                rp = doc.add_paragraph()
                rp.paragraph_format.left_indent = Cm(0.8)
                rr = rp.add_run(f"{_marker(c['no'])} 칠한 이유: {c['reason']}")
                rr.font.size = Pt(8.5)
                rr.italic = True
    if unmatched_sets:
        doc.add_heading("추가 예상 질문 (특정 문장에 붙지 않은 문항)", level=2)
        for d in unmatched_sets:
            qp = doc.add_paragraph()
            qp.add_run(f"5세트 {d['no']}번 · {d['title']}  ").bold = True
            qp.add_run(re.sub(r"\s+", " ", d["질문"]))
            if is_teacher:
                for key in ("평가의도", "모범답안", "꼬리질문"):
                    if d.get(key):
                        dr = doc.add_paragraph().add_run(f"{key}: " + re.sub(r"\s+", " ", d[key]))
                        dr.font.size = Pt(9)
    _easy = str(st.session_state.get("easy_explanation_text") or "").strip()
    if not is_teacher and _easy:
        from docx.enum.text import WD_BREAK
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        doc.add_heading("부록. 생기부 쉬운 해설 (고등학교 1학년 눈높이)", level=2)
        add_intro_paragraphs(doc, _easy)
    suffix = "교사용_지침서" if is_teacher else "학생용_워크북"
    path = _out_path(f"{student_name}_{target_desc}_{suffix}.docx")
    doc.save(path)
    return path

def create_qa_sheet_word(hl, student_name, target_desc):
    """질문만 모은 1부와 모범답안 2부로 된 연습지."""
    from docx.enum.text import WD_BREAK
    try:
        unmatched_sets = attach_set_details(hl, st.session_state.get("last_result_text", ""))
    except Exception:
        unmatched_sets = []
    rows = [(c, q) for c in hl["items"] for q in c["qs"]]
    doc = Document()
    set_document_font(doc)
    doc.add_heading(f"❓ [{student_name}] {target_desc} 예상 질문 · 모범답안지", level=1)
    doc.add_paragraph("1부는 질문만 있습니다. 먼저 소리 내어 답해 본 뒤 2부의 모범답안과 비교하세요.")
    doc.add_heading("1부. 예상 질문", level=2)
    last_area = None
    for c, q in rows:
        area = hl["paras"][c["para"]]["area"] or c["area_key"]
        if area != last_area:
            doc.add_heading(f"📂 {area}", level=3)
            last_area = area
        qp = doc.add_paragraph()
        qp.paragraph_format.space_after = Pt(2)
        qp.add_run(f"{q['no']}. ").bold = True
        qp.add_run(q["q"])
        tr = doc.add_paragraph().add_run(f"근거 {_marker(c['no'])} “{c['quote']}”  ·  {_q_tags(c, q)}     □ 자신 있음  □ 애매  □ 모름")
        tr.font.size = Pt(8.5)
        tr.font.color.rgb = RGBColor(0x75, 0x75, 0x75)
    for d in unmatched_sets:
        qp = doc.add_paragraph()
        qp.add_run(f"5세트 {d['no']}번. ").bold = True
        qp.add_run(re.sub(r"\s+", " ", d["질문"]))
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.add_heading("2부. 모범답안", level=2)
    for c, q in rows:
        qp = doc.add_paragraph()
        qp.paragraph_format.space_after = Pt(2)
        qp.add_run(f"{q['no']}. {q['q']}").bold = True
        ans = answer_of(q)
        ap = doc.add_paragraph(ans if ans else "(아직 모범답안이 없습니다. 앱의 '질문만 보기'에서 모범답안을 만들 수 있습니다.)")
        if not ans and q.get("guide"):
            gr = doc.add_paragraph().add_run("답변 가이드: " + q["guide"])
            gr.font.size = Pt(9)
        tail = (q.get("detail") or {}).get("꼬리질문")
        if tail:
            tr = doc.add_paragraph().add_run("꼬리질문: " + re.sub(r"\s+", " ", tail))
            tr.font.size = Pt(9)
    for d in unmatched_sets:
        qp = doc.add_paragraph()
        qp.add_run(f"5세트 {d['no']}번. " + re.sub(r"\s+", " ", d["질문"])).bold = True
        doc.add_paragraph(re.sub(r"\s+", " ", d.get("모범답안", "")))
    path = _out_path(f"{student_name}_질문_모범답안지.docx")
    doc.save(path)
    return path

def _card_cell(cell, paras, fill=None):
    """paras: [(글, 크기, 굵게, 색)] — 한 칸에 문단 여러 개."""
    set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
    if fill:
        set_cell_background(cell, fill)
    for i, (text, size, bold, color) in enumerate(paras):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(str(text))
        r.font.size = Pt(size)
        r.bold = bold
        if color is not None:
            r.font.color.rgb = color

def _repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)

_GRAY = RGBColor(0x60, 0x60, 0x60)
_NAVY = RGBColor(0, 51, 102)
_RED = RGBColor(0xC6, 0x28, 0x28)

def create_interview_card_word(card, student_name, target_desc, info_row=None, keywords=None):
    from docx.shared import Cm
    from docx.enum.text import WD_BREAK
    doc = Document()
    set_document_font(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.4)
    sec.top_margin = sec.bottom_margin = Cm(1.3)

    doc.add_heading(f"🗂️ [{student_name}] {card.get('uni', '')} {card.get('major', '')} 면접 전날 최종 요약카드", level=1)
    meta = doc.add_paragraph()
    run = meta.add_run(f"핵심 개념 {len(card.get('concepts', []))}개 · 예상 질문 {len(card.get('qas', []))}선 · 만든 날 {card.get('made', '')}")
    run.font.size = Pt(9)
    run.font.color.rgb = _GRAY
    if info_row is not None:
        bits = []
        for col in ("면접유형", "면접시간", "면접위원"):
            v = str(info_row.get(col, "") or "").strip()
            if v and v.lower() != "nan":
                bits.append(f"{col} {v}")
        if bits:
            p = doc.add_paragraph()
            r = p.add_run("📋 면접 형식: " + " · ".join(bits))
            r.font.size = Pt(9.5)
            r.bold = True
    if keywords:
        p = doc.add_paragraph()
        r = p.add_run("🔑 핵심 키워드: " + " · ".join(keywords))
        r.font.size = Pt(9.5)
        r.bold = True
        r.font.color.rgb = _NAVY

    concepts = card.get("concepts") or []
    if concepts:
        doc.add_heading("1. 생기부 핵심 개념 카드", level=2)
        note = doc.add_paragraph()
        nr = note.add_run("면접관이 \"그게 무엇인가요?\"라고 되물을 수 있는 용어입니다. 뜻을 소리 내어 30초 안에 말할 수 있는지 확인하세요.")
        nr.font.size = Pt(8.5)
        nr.font.color.rgb = _GRAY
        t = doc.add_table(rows=1, cols=4)
        t.style = "Table Grid"
        for i, h in enumerate(["개념", "뜻", "어디서 나왔나 (학년 · 과목 · 내용)", "답변에 덧붙일 말 · 꼬리질문"]):
            _card_cell(t.rows[0].cells[i], [(h, 9, True, RGBColor(255, 255, 255))], fill="1C4532")
        _repeat_header(t.rows[0])
        for n, c in enumerate(concepts, 1):
            cells = t.add_row().cells
            term_paras = [(f"{n}. {c['term']}", 9.5, True, _NAVY)]
            if c.get("kind"):
                term_paras.append((c["kind"], 7.5, False, _GRAY))
            if not c.get("found"):
                term_paras.append(("⚠️ 원문 표기 확인", 7.5, False, _RED))
            _card_cell(cells[0], term_paras, fill="F4F6F9")
            _card_cell(cells[1], [(c.get("desc") or "-", 8.5, False, None)])
            src_paras = []
            for s in c.get("srcs") or []:
                src_paras.append((f"{s['grade']} · {s['area']}".strip(" ·"), 8, True, None))
                if s.get("what"):
                    src_paras.append((s["what"], 8, False, _GRAY))
            _card_cell(cells[2], src_paras or [("-", 8, False, None)])
            add_paras = [(c.get("add") or "-", 8.5, False, None)]
            if c.get("follow"):
                add_paras.append(("🔥 " + c["follow"], 8, False, _RED))
            _card_cell(cells[3], add_paras)
        _set_col_widths(t, [2.8, 5.2, 4.6, 5.6])

    qas = card.get("qas") or []
    if qas:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        doc.add_heading(f"2. {card.get('major', '')} 예상 질문 {len(qas)}선 · 모범답안", level=2)
        note = doc.add_paragraph()
        nr = note.add_run("앞쪽은 학과 면접 빈출 질문, 뒤쪽은 생기부 질문입니다. [본인 경험: …] 부분은 자기 말로 채워 연습하세요.")
        nr.font.size = Pt(8.5)
        nr.font.color.rgb = _GRAY
        t = doc.add_table(rows=1, cols=3)
        t.style = "Table Grid"
        for i, h in enumerate(["번호", "예상 질문", "모범답안"]):
            _card_cell(t.rows[0].cells[i], [(h, 9, True, RGBColor(255, 255, 255))], fill="1C4532")
        _repeat_header(t.rows[0])
        for n, qa in enumerate(qas, 1):
            cells = t.add_row().cells
            _card_cell(cells[0], [(str(n), 10, True, _NAVY)], fill="F4F6F9")
            q_paras = [(qa["q"], 9, True, None), (qa.get("type", ""), 7.5, False, _GRAY)]
            if qa.get("basis") and qa["basis"] != "없음":
                q_paras.append(("근거: " + qa["basis"], 7.5, False, _GRAY))
            if qa.get("ref"):
                q_paras.append(("기출 참고: " + qa["ref"], 7.5, False, _GRAY))
            _card_cell(cells[1], q_paras)
            a_paras = [(qa["answer"], 8.5, False, None)]
            if qa.get("keys"):
                a_paras.append(("핵심: " + " · ".join(qa["keys"]), 8, True, _NAVY))
            _card_cell(cells[2], a_paras)
        _set_col_widths(t, [1.0, 5.4, 11.8])

        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        doc.add_heading("3. 질문만 보기 (마지막 자가 점검)", level=2)
        for n, qa in enumerate(qas, 1):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.add_run(f"□ {n}. ").bold = True
            p.add_run(qa["q"])
        for c in concepts:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            p.add_run("□ 개념: ").bold = True
            p.add_run(f"{c['term']} — 30초 안에 뜻을 말할 수 있나요?")

    path = _out_path(f"{student_name}_{target_desc}_면접전날_요약카드.docx")
    doc.save(path)
    return path

def create_question_sheet_word(sheet, student_name, target_desc, is_teacher=False, info_row=None, file_tag=""):
    from docx.shared import Cm
    doc = Document()
    set_document_font(doc)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.4)
    sec.top_margin = sec.bottom_margin = Cm(1.3)
    uni, major, qs = sheet.get("uni", ""), sheet.get("major", ""), sheet.get("qs") or []
    doc.add_heading(f"📋 [{student_name}] {uni} {major} 면접 질문지 ({'교사용 · 모범답안 포함' if is_teacher else '학생용'})", level=1)
    meta = doc.add_paragraph()
    mr = meta.add_run(f"질문 {len(qs)}개 · 출제 기준: {sheet.get('basis', '')} · 만든 날 {sheet.get('made', '')}")
    mr.font.size = Pt(8.5)
    mr.font.color.rgb = _GRAY
    if info_row is not None:
        bits = []
        for col in ("면접유형", "면접시간", "면접위원"):
            v = str(info_row.get(col, "") or "").strip()
            if v and v.lower() != "nan":
                bits.append(f"{col} {v}")
        if bits:
            r = doc.add_paragraph().add_run("📋 면접 형식: " + " · ".join(bits))
            r.font.size = Pt(9.5)
            r.bold = True
    if is_teacher and sheet.get("style_line"):
        r = doc.add_paragraph().add_run("🏫 대학 스타일: " + sheet["style_line"])
        r.font.size = Pt(9)
        r.font.color.rgb = _NAVY
    guide = doc.add_paragraph()
    gr = guide.add_run(
        "질문 순서는 실제 면접 흐름대로입니다. 질문마다 근거 생기부 문장(📍)을 함께 적었으니 그 부분을 다시 읽고 소리 내어 답해 보세요."
        if not is_teacher else
        "학생용과 같은 순서입니다. 모의면접 때 질문을 읽고, 답변 뒤 🔥 꼬리질문으로 압박해 보세요. ✖ 피할 답변이 나오면 바로 짚어 주세요.")
    gr.font.size = Pt(8.5)
    gr.font.color.rgb = _GRAY

    if is_teacher:
        t = doc.add_table(rows=1, cols=3)
        heads, widths = ["번호", "질문 · 생기부 출처", "모범답안 · 꼬리질문"], [1.0, 6.4, 10.8]
    else:
        t = doc.add_table(rows=1, cols=3)
        heads, widths = ["번호", "예상 질문", "이 질문이 나온 생기부 (📍 학년 · 영역 · 원문)"], [1.0, 9.0, 8.2]
    t.style = "Table Grid"
    for i, h in enumerate(heads):
        _card_cell(t.rows[0].cells[i], [(h, 9, True, RGBColor(255, 255, 255))], fill="1F3864")
    _repeat_header(t.rows[0])
    for n, q in enumerate(qs, 1):
        cells = t.add_row().cells
        _card_cell(cells[0], [(str(n), 10, True, _NAVY)], fill="F4F6F9")
        src_paras = []
        if q.get("quote") or _src_line(q):
            src_paras.append(("📍 " + (_src_line(q) or "생기부") + _found_mark(q), 8, True,
                              _RED if q.get("found") is False else None))
            if q.get("quote"):
                src_paras.append((f"“{q['quote']}”", 8, False, _GRAY))
        if is_teacher:
            paras = [(q["q"], 9, True, None),
                     (" · ".join(x for x in (q.get("type"), q.get("eval")) if x), 7.5, False, _GRAY)]
            paras += src_paras
            if q.get("style"):
                paras.append(("🏫 " + q["style"], 7.5, False, _NAVY))
            _card_cell(cells[1], paras)
            a = [(q.get("answer") or "-", 8.5, False, None)]
            if q.get("keys"):
                a.append(("핵심: " + " · ".join(q["keys"]), 8, True, _NAVY))
            if q.get("follow"):
                a.append(("🔥 꼬리질문: " + q["follow"], 8, False, _RED))
            if q.get("avoid"):
                a.append(("✖ 피할 답변: " + q["avoid"], 7.5, False, _GRAY))
            _card_cell(cells[2], a)
        else:
            _card_cell(cells[1], [(q["q"], 9.5, True, None), ("□ 자신 있음   □ 애매   □ 모름", 7.5, False, _GRAY)])
            _card_cell(cells[2], src_paras or [("생기부 전체 (지원 동기·마무리 질문)", 8, False, _GRAY)])
    _set_col_widths(t, widths)
    tag = f"_{file_tag}" if file_tag else ""
    path = _out_path(f"{student_name}_{target_desc}_면접질문지{tag}_{'교사용' if is_teacher else '학생용'}.docx")
    doc.save(path)
    return path

def create_summary_card_word(result_text, student_name, university, major, interview_type, max_items=6):
    """면접 직전에 훑어볼 수 있는 A4 한 장짜리 핵심 요약카드를 만듭니다."""
    pairs = extract_qa_pairs(result_text)[:max_items]

    doc = Document()
    set_document_font(doc)
    doc.add_heading(f"🗂️ [{student_name}] 면접 직전 핵심 요약카드", level=1)
    info_p = doc.add_paragraph()
    info_run = info_p.add_run(f"{university} · {major}  |  {interview_type}")
    info_run.bold = True
    info_run.font.color.rgb = RGBColor(0, 51, 102)
    doc.add_paragraph("면접장 들어가기 직전, 예상 질문과 답변 핵심만 빠르게 훑어보세요.\n")

    if not pairs:
        doc.add_paragraph("요약할 문항을 찾지 못했습니다. 문항을 먼저 생성한 뒤 다시 시도해 주세요.")
    else:
        for i, pair in enumerate(pairs, start=1):
            q_text = re.sub(r'\s+', ' ', pair["q"]).strip()
            a_text = re.sub(r'\s+', ' ', pair["a"]).strip()
            point = a_text[:80] + ("…" if len(a_text) > 80 else "")

            table = doc.add_table(rows=0, cols=1)
            table.style = 'Table Grid'
            row_q = table.add_row()
            set_cell_background(row_q.cells[0], "EBF1FA")
            add_parsed_text_to_cell(row_q.cells[0], f"**Q{i}.** {q_text}")
            row_a = table.add_row()
            add_parsed_text_to_cell(row_a.cells[0], f"**핵심 포인트:** {point}")
            doc.add_paragraph()

    file_path = _out_path(f"{student_name}_면접직전_요약카드.docx")
    doc.save(file_path)
    return file_path

def _set_table_fixed_layout(table, widths=None):
    """표의 열 너비를 지정한 값 그대로 고정합니다 (자동 맞춤 해제).
    widths를 주면 표의 그리드(gridCol)와 각 셀 너비를 모두 같은 값으로 맞춥니다."""
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)

    if widths:
        for c, width in enumerate(widths):
            try:
                table.columns[c].width = width
            except (IndexError, ValueError):
                pass
        for row in table.rows:
            for c, width in enumerate(widths):
                try:
                    row.cells[c].width = width
                except IndexError:
                    pass

def _set_cell_text(cell, text, bold=False, size=9.5, align=None, color=None):
    """평가표용 셀 텍스트 입력 헬퍼 (글자 크기·정렬 지정 가능)"""
    set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        if i > 0:
            p.add_run("\n")
        run = p.add_run(line)
        run.bold = bold
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color
    return p

def _small_spacer(doc, size=6):
    """문서 사이 여백을 최소한으로 넣기 위한 작은 빈 줄"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run("")
    run.font.size = Pt(size)
    return p

def _add_write_box(doc, title, height_inches=0.9, fill="F7F9FC", spacer=True, content=""):
    """제목 줄 + 손으로 적을 수 있는 빈 칸 한 세트를 추가합니다(content가 있으면 미리 채움)."""
    table = doc.add_table(rows=2, cols=1)
    table.style = 'Table Grid'
    _set_table_fixed_layout(table, [Inches(6.1)])
    set_cell_background(table.rows[0].cells[0], fill)
    _set_cell_text(table.rows[0].cells[0], title, bold=True, size=10, color=RGBColor(0, 51, 102))
    _set_cell_text(table.rows[1].cells[0], str(content or ""), size=10)
    table.rows[1].height = Inches(height_inches)
    table.rows[1].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    if spacer:
        _small_spacer(doc)
    return table

def create_evaluation_sheet_word(result_text, student_name, university, major, interview_type,
                                  evaluator="", interview_date="", include_questions=True,
                                  run_rows=None, opinions=None):
    """교사가 면접 현장에서 바로 채점할 수 있는 인쇄용 평가표를 만듭니다."""
    doc = Document()
    set_document_font(doc)

    doc.add_heading("📝 모의면접 평가표 (교사용)", level=1)
    sub = doc.add_paragraph()
    sub_run = sub.add_run("도개고등학교 진로진학 모의면접 · 항목당 10점 만점 / 총 100점")
    sub_run.bold = True
    sub_run.font.size = Pt(10)
    sub_run.font.color.rgb = RGBColor(0, 51, 102)

    # ── 기본 정보 ──
    info_table = doc.add_table(rows=2, cols=4)
    info_table.style = 'Table Grid'
    info_pairs = [
        ("학생명", student_name), ("지원 대학", university),
        ("지원 학과", major), ("면접 유형", interview_type),
    ]
    _set_table_fixed_layout(info_table, [Inches(1.525)] * 4)
    for idx, (label, value) in enumerate(info_pairs):
        cell_label = info_table.rows[0].cells[idx]
        set_cell_background(cell_label, "EBF1FA")
        _set_cell_text(cell_label, label, bold=True, size=9.5, align="center")
        _set_cell_text(info_table.rows[1].cells[idx], value or "", size=9.5, align="center")
    _small_spacer(doc)

    info2 = doc.add_table(rows=1, cols=4)
    info2.style = 'Table Grid'
    _set_table_fixed_layout(info2, [Inches(1.2), Inches(1.85), Inches(1.2), Inches(1.85)])
    set_cell_background(info2.rows[0].cells[0], "EBF1FA")
    _set_cell_text(info2.rows[0].cells[0], "면접 일자", bold=True, size=9.5, align="center")
    _set_cell_text(info2.rows[0].cells[1], interview_date or "          .     .     .", size=9.5, align="center")
    set_cell_background(info2.rows[0].cells[2], "EBF1FA")
    _set_cell_text(info2.rows[0].cells[2], "평가 교사", bold=True, size=9.5, align="center")
    _set_cell_text(info2.rows[0].cells[3], evaluator or "", size=9.5, align="center")
    _small_spacer(doc)

    # ── 채점 척도 안내 ──
    scale_p = doc.add_paragraph()
    scale_run = scale_p.add_run(
        "▣ 채점 척도  |  9~10점 매우 우수    7~8점 우수    5~6점 보통    3~4점 미흡    1~2점 매우 미흡"
    )
    scale_run.bold = True
    scale_run.font.size = Pt(9)

    # ── 평가 항목 표 ──
    total_items = sum(len(items) for _, items in EVALUATION_RUBRIC)
    table = doc.add_table(rows=1 + total_items + 2, cols=5)
    table.style = 'Table Grid'

    headers = ["평가 영역", "평가 항목", "세부 평가 기준", "배점", "점수"]
    widths = [Inches(0.82), Inches(1.38), Inches(2.8), Inches(0.5), Inches(0.6)]
    for c, head in enumerate(headers):
        cell = table.rows[0].cells[c]
        set_cell_background(cell, "1C4532")
        _set_cell_text(cell, head, bold=True, size=9.5, align="center", color=RGBColor(255, 255, 255))

    row_idx = 1
    merge_ranges = []
    for area_name, items in EVALUATION_RUBRIC:
        start_row = row_idx
        for item_name, criterion in items:
            _set_cell_text(table.rows[row_idx].cells[1], item_name, bold=True, size=9.5)
            _set_cell_text(table.rows[row_idx].cells[2], criterion, size=9)
            _set_cell_text(table.rows[row_idx].cells[3], "10", size=9.5, align="center")
            _set_cell_text(table.rows[row_idx].cells[4], "", size=9.5, align="center")
            table.rows[row_idx].height = Inches(0.42)
            table.rows[row_idx].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            row_idx += 1
        merge_ranges.append((area_name, start_row, row_idx - 1))

    # 열 너비 고정 (병합 전에 먼저 지정)
    _set_table_fixed_layout(table, widths)

    # 합계 / 10점 환산 행 (병합을 먼저 하고 그 뒤에 글자를 채워 빈 줄이 생기지 않게 함)
    total_row = table.rows[row_idx]
    score_cell = total_row.cells[4]
    point_cell = total_row.cells[3]
    merged_total = total_row.cells[0].merge(total_row.cells[2])
    set_cell_background(merged_total, "F4F6F9")
    _set_cell_text(merged_total, "합    계", bold=True, size=10, align="center")
    _set_cell_text(point_cell, "100", bold=True, size=9.5, align="center")
    _set_cell_text(score_cell, "", size=9.5, align="center")
    row_idx += 1

    conv_row = table.rows[row_idx]
    conv_score_cell = conv_row.cells[4]
    merged_conv = conv_row.cells[0].merge(conv_row.cells[3])
    set_cell_background(merged_conv, "F4F6F9")
    _set_cell_text(merged_conv,
                   "10점 환산 점수 (합계 ÷ 10)   |   환산 등급  A: 90↑   B: 80~89   C: 70~79   D: 60~69   E: 60 미만",
                   bold=True, size=9, align="center")
    _set_cell_text(conv_score_cell, "", size=9.5, align="center")

    # 영역 셀 세로 병합 (내용 영역 / 태도·표현 영역)
    for area_name, start_row, end_row in merge_ranges:
        merged = table.rows[start_row].cells[0].merge(table.rows[end_row].cells[0])
        set_cell_background(merged, "F4F6F9")
        _set_cell_text(merged, area_name, bold=True, size=9.5, align="center", color=RGBColor(0, 51, 102))

    _small_spacer(doc, size=8)

    # ── 문항별 평가란 ──
    pairs = extract_qa_pairs(result_text) if include_questions else []
    if include_questions and run_rows:       # 🎙️ 모의면접 진행에서 매긴 상·중·하와 메모
        pairs = [{"q": r["q"], "grade": r.get("grade", ""), "memo": r.get("memo", "")} for r in run_rows]
    if pairs:
        q_head = doc.add_paragraph()
        q_head_run = q_head.add_run("▣ 문항별 답변 평가 (모의면접 진행 결과: 상·중·하)" if run_rows else "▣ 문항별 답변 평가 (각 10점)")
        q_head_run.bold = True
        q_head_run.font.size = Pt(10)
        q_head_run.font.color.rgb = RGBColor(0, 51, 102)

        pairs = pairs[:30] if run_rows else pairs[:6]
        q_table = doc.add_table(rows=1 + len(pairs), cols=4)
        q_table.style = 'Table Grid'
        for c, head in enumerate(["번호", "면접 질문", "평가" if run_rows else "점수", "특이사항 · 메모"]):
            cell = q_table.rows[0].cells[c]
            set_cell_background(cell, "EBF1FA")
            _set_cell_text(cell, head, bold=True, size=9.5, align="center")

        q_widths = [Inches(0.4), Inches(2.9), Inches(0.55), Inches(2.25)]
        _set_table_fixed_layout(q_table, q_widths)
        for i, pair in enumerate(pairs, start=1):
            q_text = re.sub(r'\s+', ' ', pair["q"]).strip()
            if len(q_text) > 90:
                q_text = q_text[:90] + "…"
            row = q_table.rows[i]
            _set_cell_text(row.cells[0], str(i), size=9.5, align="center")
            _set_cell_text(row.cells[1], q_text, size=9)
            _set_cell_text(row.cells[2], pair.get("grade", ""), size=10, bold=True, align="center")
            _set_cell_text(row.cells[3], pair.get("memo", ""), size=9)
            row.height = Inches(0.55)
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        _small_spacer(doc, size=8)

    # ── 종합 의견란 ──
    opinion_head = doc.add_paragraph()
    opinion_run = opinion_head.add_run("▣ 종합 의견")
    opinion_run.bold = True
    opinion_run.font.size = Pt(10)
    opinion_run.font.color.rgb = RGBColor(0, 51, 102)

    ops = list(opinions or []) + ["", "", ""]
    _add_write_box(doc, "① 잘한 점 (강점) — 실제 답변 내용을 근거로 구체적으로 기재", height_inches=0.95, content=ops[0])
    _add_write_box(doc, "② 보완할 점 (약점) — 내용·태도 측면 모두 기재", height_inches=0.95, content=ops[1])
    _add_write_box(doc, "③ 다음 연습까지의 지도 방향 및 과제", height_inches=0.9, spacer=False, content=ops[2])
    _small_spacer(doc, size=8)

    # ── 서명란 ──
    sign_table = doc.add_table(rows=1, cols=4)
    sign_table.style = 'Table Grid'
    _set_table_fixed_layout(sign_table, [Inches(1.0), Inches(2.05), Inches(1.5), Inches(1.55)])
    set_cell_background(sign_table.rows[0].cells[0], "EBF1FA")
    _set_cell_text(sign_table.rows[0].cells[0], "평가일", bold=True, size=9.5, align="center")
    _set_cell_text(sign_table.rows[0].cells[1], interview_date or "         .      .      .", size=9.5, align="center")
    set_cell_background(sign_table.rows[0].cells[2], "EBF1FA")
    _set_cell_text(sign_table.rows[0].cells[2], "평가 교사 (서명)", bold=True, size=9.5, align="center")
    _set_cell_text(sign_table.rows[0].cells[3], f"{evaluator}                    (인)" if evaluator else "                             (인)", size=9.5, align="center")

    file_path = _out_path(f"{student_name}_면접평가표.docx")
    doc.save(file_path)
    return file_path

def create_growth_report_word(report_text, student_name):
    doc = Document()
    set_document_font(doc)
    doc.add_heading(f"📈 [{student_name}] 학생 모의면접 성장 리포트", level=1)
    doc.add_paragraph("여러 차례의 모의면접 연습 기록을 바탕으로 정리한 성장 리포트입니다. 학부모 상담 자료로도 활용하실 수 있습니다.\n")
    add_intro_paragraphs(doc, report_text.replace("### 📈", "### 🔍"))
    file_path = _out_path(f"{student_name}_성장리포트.docx")
    doc.save(file_path)
    return file_path


__all__ = ["set_document_font", "set_cell_background", "set_cell_margins", "add_parsed_text_to_cell", "create_word_files", "create_chat_history_word", "create_easy_explanation_word", "_set_col_widths", "create_highlight_word", "create_qa_sheet_word", "_card_cell", "_repeat_header", "_GRAY", "_NAVY", "_RED", "create_interview_card_word", "create_question_sheet_word", "create_summary_card_word", "_set_table_fixed_layout", "_set_cell_text", "_small_spacer", "_add_write_box", "create_evaluation_sheet_word", "create_growth_report_word"]
