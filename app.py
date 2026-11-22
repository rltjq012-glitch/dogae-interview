import sys
import io

# 파이썬 표준 입출력 인코딩을 UTF-8로 완전히 강제 재설정
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

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

# 페이지 기본 설정
st.set_page_config(page_title="도개고 면접 마스터", layout="wide", page_icon="🎓")

# 🌟 트렌디한 고급 UI/UX 및 다크모드 대응 CSS 주입 🌟
custom_css = """
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
html, body, [class*="css"] {
    font-family: 'Pretendard', sans-serif !important;
}
@media (prefers-color-scheme: light) {
    [data-testid="stAppViewContainer"] { background-color: #F9F8F3 !important; }
    [data-testid="stSidebar"] { background-color: #F0EFEA !important; }
}
@media (prefers-color-scheme: dark) {
    [data-testid="stAppViewContainer"] { background-color: #121212 !important; }
    [data-testid="stSidebar"] { background-color: #1A1A1A !important; }
}

/* 🎓 헤더 배너 */
.hero-banner {
    background: linear-gradient(135deg, #192c23 0%, #294435 100%);
    border-radius: 16px;
    padding: 2.5rem 3rem;
    color: white;
    margin-bottom: 2rem;
    box-shadow: 0 10px 25px rgba(0,0,0,0.15);
}
.hero-badge {
    display: inline-block;
    border: 1px solid rgba(255,255,255,0.3);
    padding: 0.3rem 1rem;
    border-radius: 20px;
    font-size: 0.85rem;
    color: #cbd5e1;
    margin-bottom: 1rem;
    letter-spacing: 1px;
}
.hero-title {
    font-size: 2.4rem;
    font-weight: 800;
    margin: 0 0 0.5rem 0;
    color: #ffffff;
}
.hero-subtitle {
    font-size: 1.05rem;
    color: #a7f3d0;
    margin: 0;
}
.hero-line {
    width: 50px;
    height: 3px;
    background-color: #d4af37;
    margin-top: 20px;
    border-radius: 2px;
}

/* STEP 텍스트 디자인 */
.step-text {
    color: #d4af37;
    font-weight: 800;
    font-size: 0.95rem;
    margin-bottom: -15px;
    letter-spacing: 1.5px;
}

/* 채팅 메시지 박스 커스텀 */
[data-testid="stChatMessage"] {
    background-color: rgba(255,255,255,0.7);
    border-radius: 16px;
    padding: 1.5rem;
    margin-bottom: 15px;
    border: 1px solid rgba(0,0,0,0.05);
    box-shadow: 0 2px 10px rgba(0,0,0,0.02);
}
@media (prefers-color-scheme: dark) {
    [data-testid="stChatMessage"] {
        background-color: rgba(255,255,255,0.05);
        border: 1px solid rgba(255,255,255,0.1);
    }
}

/* 둥근 테두리의 다운로드 버튼 */
.stDownloadButton > button {
    background-color: transparent !important;
    border: 1.5px solid #94a3b8 !important;
    color: inherit !important;
    border-radius: 30px !important;
    font-weight: 600 !important;
    padding: 0.5rem 1rem !important;
    width: 100%;
}
.stDownloadButton > button:hover {
    background-color: #e2e8f0 !important;
    border-color: #475569 !important;
    transform: translateY(-2px);
}
@media (prefers-color-scheme: dark) {
    .stDownloadButton > button:hover {
        background-color: #334155 !important;
        border-color: #f8fafc !important;
    }
}

/* 메인 동작 버튼 */
div.stButton > button:first-child {
    background: linear-gradient(135deg, #192c23 0%, #294435 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 0.7rem 1.2rem !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
    transition: all 0.3s ease !important;
    width: 100%;
}
div.stButton > button:first-child:hover {
    transform: translateY(-2px) !important;
}

.stTextInput>div>div>input { border-radius: 8px !important; }
.stFileUploader>div>div { border-radius: 12px !important; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# -------------------------------------------------------------------------
# 상단 헤더 배너 (HTML)
# -------------------------------------------------------------------------
st.markdown("""
<div class="hero-banner">
    <div class="hero-badge">도개고등학교 - 김기섭</div>
    <h1 class="hero-title">🎓 도개고 대입 모의면접 마스터 솔루션</h1>
    <p class="hero-subtitle">생기부와 실제 대학 기출 형식을 정교하게 분석해, 실전과 같은 모의면접 세트를 설계합니다</p>
    <div class="hero-line"></div>
</div>
""", unsafe_allow_html=True)

# ── 기능 모듈 (dogae/ 폴더) ─────────────────────────────────────────────
from dogae.base import *  # noqa: F401,F403
from dogae.ai import *  # noqa: F401,F403
from dogae.storage import *  # noqa: F401,F403
from dogae.exam_data import *  # noqa: F401,F403
from dogae.questions import *  # noqa: F401,F403
from dogae.word_docs import *  # noqa: F401,F403
from dogae.screens import *  # noqa: F401,F403


if not (_appsscript_enabled() or _sheets_enabled()):
    init_records_db()

def build_explain_and_highlight_prompt(student_record_text, questions, priority_text):
    """'생기부 쉬운 해설'과 '질문 카드'를 한 번의 호출로 요청합니다."""
    return build_easy_explanation_prompt(trim_record_for_review(student_record_text)) + f"""

    [추가 작업: 형광펜 질문 카드 — 해설을 모두 쓴 뒤에 이어서 작성]
    해설 맨 끝에 '{HL_MARKER}' 한 줄을 쓰고, 그 아래에 칠할 문장을 한 줄에 하나씩 JSON으로 쓰세요.
    {_hl_card_rules(priority_text)}

    [이미 만든 면접 질문]
    {_q_lines(questions)}
    """

def run_explain_and_highlight(student_record, result_text, profile, uni_name, student_name, target_desc, api_key):
    """'생기부 쉬운 해설'과 '형광펜 질문 카드'를 만들어 세션에 넣습니다.
    실패해도 화면을 다시 그리면 사라지지 않도록 이유를 st.session_state.highlight_error에 남깁니다."""
    st.session_state.highlight_data = None
    st.session_state.highlight_files = None
    st.session_state.highlight_raw_items = []
    st.session_state.highlight_error = ""
    questions = [qa["q"] for qa in extract_qa_pairs(result_text)]
    priority = highlight_priority_text(profile, uni_name)
    try:
        raw = _call_gemini_retry(build_explain_and_highlight_prompt(student_record, questions, priority), api_key)
        easy_explanation, hl_items = split_explanation_and_highlights(raw)
        st.session_state.easy_explanation_text = easy_explanation
        st.session_state.easy_explanation_file = create_easy_explanation_word(easy_explanation, student_name, target_desc)
        hl = build_highlight_data(student_record, hl_items) if hl_items else None
        if not (hl and hl["items"]):
            first_note = (
                "해설 응답에 질문 카드가 없었습니다" if not hl_items
                else "해설 응답의 문장이 원문과 맞지 않았습니다"
            )
            # 해설이 길면 AI가 카드 부분을 빼먹는 일이 잦아, 카드만 JSON으로 따로 한 번 더 받습니다.
            raw2 = _call_gemini_retry(build_highlight_only_prompt(student_record, questions, priority), api_key, json_mode=True)
            hl_items = _extract_hl_objects(raw2)
            if not any("구절" in it for it in hl_items):
                head = re.sub(r"\s+", " ", str(raw2))[:120]
                st.session_state.highlight_error = (
                    f"{first_note}. 카드만 다시 요청했지만 목록을 읽지 못했습니다. (AI 응답 앞부분: {head})"
                )
                return False
            hl = build_highlight_data(student_record, hl_items)
            if not hl["items"]:
                sample = next((str(it.get("구절", "")) for it in hl_items if "구절" in it), "")[:40]
                st.session_state.highlight_error = (
                    f"{first_note}. 다시 요청한 문장도 모두 원문과 글자가 달라 칠하지 못했습니다. (예: {sample})"
                )
                return False
        set_highlight_state(hl, hl_items, student_name, target_desc)
        return True
    except Exception as e:
        st.session_state.highlight_error = _friendly_error(e)
        return False

# -------------------------------------------------------------------------
# [3] 메인 UI 설정
# -------------------------------------------------------------------------
if "chat_history" not in st.session_state: st.session_state.chat_history = []
if "word_files" not in st.session_state: st.session_state.word_files = None
if "last_audio_size" not in st.session_state: st.session_state.last_audio_size = 0
if "last_result_text" not in st.session_state: st.session_state.last_result_text = ""
if "last_examples" not in st.session_state: st.session_state.last_examples = None
if "current_record_id" not in st.session_state: st.session_state.current_record_id = None
if "current_record_key" not in st.session_state: st.session_state.current_record_key = None
if "record_key_map" not in st.session_state: st.session_state.record_key_map = {}
if "loaded_student_record_text" not in st.session_state: st.session_state.loaded_student_record_text = ""
if "easy_explanation_text" not in st.session_state: st.session_state.easy_explanation_text = ""
if "easy_explanation_file" not in st.session_state: st.session_state.easy_explanation_file = None
if "highlight_data" not in st.session_state: st.session_state.highlight_data = None
if "highlight_files" not in st.session_state: st.session_state.highlight_files = None
if "highlight_error" not in st.session_state: st.session_state.highlight_error = ""
if "highlight_raw_items" not in st.session_state: st.session_state.highlight_raw_items = []
if "card_checks" not in st.session_state: st.session_state.card_checks = {}
if "summary_card_file" not in st.session_state: st.session_state.summary_card_file = None
if "growth_report_file" not in st.session_state: st.session_state.growth_report_file = None
if "evaluation_sheet_file" not in st.session_state: st.session_state.evaluation_sheet_file = None
if "interview_card" not in st.session_state: st.session_state.interview_card = None
if "interview_card_file" not in st.session_state: st.session_state.interview_card_file = None
if "interview_card_error" not in st.session_state: st.session_state.interview_card_error = ""
if "qsheet" not in st.session_state: st.session_state.qsheet = None
if "qsheet_files" not in st.session_state: st.session_state.qsheet_files = None
if "qsheet_error" not in st.session_state: st.session_state.qsheet_error = ""
if "qbank_sel" not in st.session_state: st.session_state.qbank_sel = None
if "qbank_runs" not in st.session_state: st.session_state.qbank_runs = []
if "run" not in st.session_state: st.session_state.run = None
if "last_save_info" not in st.session_state: st.session_state.last_save_info = None

# =========================================================================
# 📱 학생 연습 — 학생별 링크로 휴대폰·패드에서 답변 작성 → AI 첨삭
#   저장은 기존 기록 저장소(구글시트 Apps Script / 서비스 계정 / 로컬)를 그대로 씁니다.
#   어느 저장 방식에서나 확실히 남는 칸(id·이름·대학·학과·result_text·chat_history)만 사용합니다.
#     STU-<묶음키>.<토큰>  : 학생 링크 (PIN, 하루 첨삭 횟수, 학생에게 보낼 질문·근거·모범답안)
#     ANS-<묶음키>         : 학생 답안·첨삭 기록 (학생 기기에서 저장)
#     CMT-<묶음키>         : 선생님 코멘트
#   <묶음키>는 선생님 기록 id에서 만든 해시라서, 링크를 새로 만들어도 답안 기록은 이어집니다.
# =========================================================================
import hashlib as _hashlib
import secrets as _pysecrets
import wave as _wave


# 학생 링크(?s=...)로 들어오면 교사 로그인 없이 그 학생 화면만 보여 주고 멈춥니다.
_stu_param = st.query_params.get("s")
if _stu_param:
    render_student_page(_stu_param)
    st.stop()

# 면접관 링크(?iv=...)로 들어오면 패드·폰용 모의면접 진행 화면만 보여 줍니다(링크 + PIN).
_iv_param = st.query_params.get("iv")
if _iv_param:
    render_interviewer_page(_iv_param)
    st.stop()

require_login()

exam_db, exam_db_files = load_exam_db(APP_DIR)
univ_info_db, univ_info_files = load_univ_info(APP_DIR)
RT.exam_db, RT.univ_info_db = exam_db, univ_info_db   # 모듈 함수가 쓰는 자료

def _rec_get(row, key, default=""):
    """sqlite3.Row와 Apps Script/Sheets에서 온 dict를 모두 안전하게 다루기 위한 헬퍼.
    (Apps Script 백엔드는 생기부 원문·면접유형·난이도를 저장하지 않으므로 해당 키가 없을 수 있음)"""
    try:
        val = row[key]
        return val if val not in (None, "") else default
    except (KeyError, IndexError):
        return default

def render_record_loader(api_key):
    """사이드바: 저장된 학생 기록 불러오기 · 성장 리포트 · 삭제."""
    saved_records = list_records()
    if not saved_records:
        st.caption("아직 저장된 기록이 없습니다. 문항을 한 번 생성하면 이 학생의 문항·대화가 자동으로 저장됩니다.")
    else:
        _show_owner = is_admin()
        record_options = {
            (f"[{record_owner(r['id']) or '예전 기록'}] " if _show_owner else "")
            + f"{r['student_name']} · {r['university']}_{r['major']} — "
            + f"{str(r['updated_at'])[:16].replace('T', ' ')}": r["id"]
            for r in saved_records
        }
        selected_label = st.selectbox("불러올 기록", list(record_options.keys()), key="record_select_box",
                                      label_visibility="collapsed", placeholder="학생 기록 선택")
        with st.container():
            if st.button("📥 이 기록 불러오기", width="stretch", type="primary", key="load_record_btn"):
                row = get_record(record_options[selected_label])
                if row:
                    st.session_state["region_select"] = "직접 입력"
                    st.session_state["uni_direct_input"] = row["university"]
                    st.session_state["major_select"] = MAJOR_DIRECT
                    st.session_state["major_input"] = row["major"]
                    st.session_state["student_name_input"] = row["student_name"]
                    st.session_state["interview_type_radio"] = _rec_get(row, "interview_type", "생기부 기반 면접")
                    st.session_state["difficulty_radio"] = _rec_get(row, "difficulty", "중 (표준)")
                    st.session_state["loaded_student_record_text"] = _rec_get(row, "student_record_text", "")
                    st.session_state["last_result_text"] = row["result_text"] or ""
                    _hist, _extras = load_saved_history(_rec_get(row, "chat_history", ""))
                    st.session_state["chat_history"] = _hist
                    st.session_state["current_record_id"] = row["id"]
                    _loaded_key = (
                        f"{row['student_name']}|{row['university']}|{row['major']}|"
                        f"{_rec_get(row, 'interview_type', '생기부 기반 면접')}"
                    )
                    st.session_state["current_record_key"] = _loaded_key
                    st.session_state.setdefault("record_key_map", {})[_loaded_key] = row["id"]
                    # 생기부 쉬운 해설은 저장소에 보관하지 않으므로, 기록을 불러올 때는 일단 비워두고
                    # 필요하면 '면접 패키지 생성 시작'을 다시 눌러 새로 만들 수 있게 합니다.
                    st.session_state["easy_explanation_text"] = ""
                    st.session_state["easy_explanation_file"] = None
                    st.session_state["highlight_data"] = None
                    st.session_state["highlight_files"] = None
                    st.session_state["highlight_raw_items"] = []
                    st.session_state["highlight_error"] = (
                        "이 기록에는 형광펜이 저장되어 있지 않습니다(형광펜 저장 기능을 넣기 전에 만든 기록). 아래 버튼으로 다시 만들어 주세요."
                    )
                    # 저장된 형광펜 목록·쉬운 해설을 복원 (AI를 다시 부르지 않음)
                    _ld_name = row["student_name"]
                    _ld_desc = f"{row['university']}_{row['major']}"
                    _ld_rec = st.session_state["loaded_student_record_text"]
                    _easy = str(_extras.get("easy_explanation") or "")
                    if _easy.strip():
                        st.session_state["easy_explanation_text"] = _easy
                        try:
                            st.session_state["easy_explanation_file"] = create_easy_explanation_word(_easy, _ld_name, _ld_desc)
                        except Exception:
                            pass
                    st.session_state["card_checks"] = dict(_extras.get("card_checks") or {})
                    for _k in [k for k in list(st.session_state.keys()) if str(k).startswith("chk_")]:
                        del st.session_state[_k]
                    _items = [it for it in (_extras.get("highlight_items") or []) if isinstance(it, dict)]
                    if _items and _ld_rec.strip():
                        try:
                            _hl_loaded = build_highlight_data(_ld_rec, _items)
                            if _hl_loaded["items"]:
                                st.session_state["highlight_data"] = _hl_loaded
                                st.session_state["highlight_raw_items"] = _items
                                st.session_state["highlight_error"] = ""
                                st.session_state["highlight_files"] = (
                                    create_highlight_word(_hl_loaded, _ld_name, _ld_desc, is_teacher=False),
                                    create_highlight_word(_hl_loaded, _ld_name, _ld_desc, is_teacher=True),
                                    create_qa_sheet_word(_hl_loaded, _ld_name, _ld_desc),
                                )
                        except Exception as e:
                            st.session_state["highlight_error"] = f"저장된 형광펜을 복원하지 못했습니다: {e}"
                    st.session_state["summary_card_file"] = None
                    st.session_state["growth_report_file"] = None
                    st.session_state["evaluation_sheet_file"] = None
                    st.session_state["interview_card_error"] = ""
                    st.session_state["qsheet_error"] = ""
                    _qb = load_qbank_state(row["id"])
                    st.session_state["qbank_sel"] = _qb.get("sel") or None
                    st.session_state["qbank_runs"] = _qb.get("runs") or []
                    st.session_state["run"] = None
                    st.session_state["last_run"] = None
                    st.session_state.pop("_qbank_cache", None)
                    try:
                        set_question_sheet_state(load_question_sheet(row["id"]), _ld_name, _ld_desc,
                                                 row["university"], row["major"])
                    except Exception as e:
                        st.session_state["qsheet"] = None
                        st.session_state["qsheet_files"] = None
                        st.session_state["qsheet_error"] = f"저장된 질문지를 불러오지 못했습니다: {e}"
                    try:
                        set_interview_card_state(load_interview_card(row["id"]), _ld_name, _ld_desc,
                                                 row["university"], row["major"])
                    except Exception as e:
                        st.session_state["interview_card"] = None
                        st.session_state["interview_card_file"] = None
                        st.session_state["interview_card_error"] = f"저장된 요약카드를 불러오지 못했습니다: {e}"
                    if row["result_text"]:
                        loaded_type = _rec_get(row, "interview_type", "생기부 기반 면접")
                        stu_path, tea_path = create_word_files(
                            row["result_text"], row["student_name"], loaded_type,
                            f"{row['university']}_{row['major']}"
                        )
                        st.session_state["word_files"] = (stu_path, tea_path)
                        # 평가표는 AI 호출 없이 만들 수 있으므로 불러올 때 함께 다시 생성
                        try:
                            st.session_state["evaluation_sheet_file"] = create_evaluation_sheet_word(
                                row["result_text"], row["student_name"], row["university"], row["major"], loaded_type
                            )
                        except Exception:
                            pass
                    if not _rec_get(row, "student_record_text", ""):
                        st.info("ℹ️ 이 저장 방식은 생기부 원문은 따로 저장하지 않습니다. 문항/대화 내용은 그대로 불러왔고, 생기부 재분석이 필요하면 PDF를 다시 업로드해주세요.")
                    st.session_state.pop("result_section", None)   # 새 학생은 첫 화면부터
                    st.toast(f"'{row['student_name']}' 학생의 기록을 불러왔습니다.", icon="📥")
                    st.rerun()
        _c_grow, _c_del = st.columns([3, 2])
        with _c_del:
            with st.popover("🗑️ 삭제", width="stretch"):
                st.caption("선택한 기록을 지웁니다. 되돌릴 수 없습니다.")
                if st.button("삭제 확인", type="primary", width="stretch", key="confirm_delete_record"):
                    delete_record(record_options[selected_label])
                    st.toast("삭제했습니다.", icon="🗑️")
                    st.rerun()
        with _c_grow:
            if st.button("📈 성장 리포트", width="stretch", key="growth_report_btn"):
                if not api_key:
                    st.error("API 키를 입력해 주세요.")
                else:
                    target_id = record_options[selected_label]
                    target_row = next((r for r in saved_records if r["id"] == target_id), None)
                    target_name = target_row["student_name"] if target_row else None
                    matching = [r for r in saved_records if r.get("student_name") == target_name]
                    matching.sort(key=lambda r: r.get("updated_at", ""))
                    with st.spinner("휴대폰 답안 연습 기록을 확인하는 중입니다..."):
                        try:
                            _practice_txt = practice_summary_for_growth(target_name)
                        except Exception:
                            _practice_txt = ""
                    if len(matching) < 2 and not _practice_txt:
                        st.warning(f"'{target_name}' 학생은 기록이 {len(matching)}개뿐이라 성장 비교가 어렵습니다.")
                    else:
                        with st.spinner(f"'{target_name}' 학생의 {len(matching)}개 회차 기록을 분석해 성장 리포트를 작성하는 중입니다..."):
                            sessions = []
                            for r in matching[-8:]:  # 너무 길어지지 않도록 최근 8회차까지만
                                full = get_record(r["id"])
                                if not full:
                                    continue
                                chat_hist, _ = load_saved_history(_rec_get(full, "chat_history", ""))
                                sessions.append({
                                    "updated_at": full["updated_at"],
                                    "university": full["university"],
                                    "major": full["major"],
                                    "chat_history": chat_hist,
                                })
                            try:
                                report_text = call_gemini(build_growth_report_prompt(target_name, sessions, _practice_txt), api_key)
                                report_path = create_growth_report_word(report_text, target_name)
                                st.session_state.growth_report_file = report_path
                                st.session_state.growth_report_text = report_text
                                st.session_state.growth_report_for = target_name
                                st.session_state.result_section = "📊 평가·리포트"
                                st.toast(f"'{target_name}' 학생의 성장 리포트를 만들었습니다.", icon="📈")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ 성장 리포트 생성 실패: {e}")


with st.sidebar:
    _teacher = current_teacher()
    if _teacher:
        _tc1, _tc2 = st.columns([3, 2], vertical_alignment="center")
        _tc1.markdown(f"👤 **{_teacher}** 선생님" + (" · 관리자" if is_admin() else ""))
        if _tc2.button("로그아웃", width="stretch", key="logout_button"):
            logout()
    else:
        st.warning("🔓 로그인이 설정되지 않았습니다. 주소를 아는 누구나 저장된 학생 기록을 볼 수 있습니다.")

    # Streamlit Cloud의 Secrets(GEMINI_API_KEY)에 키가 등록되어 있으면 자동으로 사용하고,
    # 없을 경우에만 직접 입력창을 보여줍니다.
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.caption("🔑 API 키 연결됨" + (f" · 🤖 {st.session_state['gemini_last_model']}"
                                       if st.session_state.get("gemini_last_model") else ""))
    except (KeyError, FileNotFoundError, AttributeError):
        api_key = st.text_input("🔑 Gemini API Key", type="password")

    # 꼭 알아야 하는 문제만 바깥에 보여 줍니다(평소에는 ⚙️ 관리 안에 접어 둠).
    if exam_db.empty:
        st.warning("⚠️ 실제 기출 DB(master_interview_qa.csv)가 없습니다.")
    _last_save = st.session_state.get("last_save_info")
    if _last_save and not _last_save.get("ok"):
        st.error("⚠️ 마지막 저장이 실패했습니다. ⚙️ 관리에서 연결을 확인해 주세요.")

    st.markdown("#### 📂 학생 기록")
    render_record_loader(api_key)

    st.divider()
    with st.expander("⚙️ 관리 · 도구", expanded=False):
        st.toggle("👥 여러 학생 일괄 생성 열기", key="show_batch",
                  help="켜면 본문에 '여러 학생 한 번에 처리' 화면이 나타납니다.")
        st.markdown("**저장소**")
        if _appsscript_enabled():
            st.caption("☁️ Google Sheets(Apps Script)에 반영구 저장 중")
        elif _sheets_enabled():
            st.caption("☁️ Google Sheets(서비스 계정)에 반영구 저장 중")
        else:
            st.caption("💾 로컬 임시 저장 중 (재배포 시 초기화될 수 있음)")
        render_save_status()
        # 구글시트 연결이 실제로 살아있는지 직접 확인 (토큰이 어긋나면 저장이 조용히 실패할 수 있음)
        if _appsscript_enabled():
            if st.button("🔌 구글시트 연결 테스트", width="stretch", key="conn_test_btn"):
                try:
                    test_payload = {
                        "token": st.secrets["APPS_SCRIPT_TOKEN"], "action": "save",
                        "id": "__connection_test__", "student_name": "(연결테스트)",
                        "university": "-", "major": "-", "result_text": "", "chat_history": "[]",
                        "updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
                    }
                    r = requests.post(st.secrets["APPS_SCRIPT_URL"], json=test_payload, timeout=30)
                    r.raise_for_status()
                    body = r.json()
                    if isinstance(body, dict) and body.get("status") == "ok":
                        requests.post(
                            st.secrets["APPS_SCRIPT_URL"],
                            json={"token": st.secrets["APPS_SCRIPT_TOKEN"], "action": "delete", "id": "__connection_test__"},
                            timeout=30,
                        )
                        st.success("✅ 구글시트 저장이 정상 동작합니다.")
                    else:
                        st.error(f"❌ 저장 실패 응답: {body}\n\n토큰이 서로 다르거나 Apps Script를 새 버전으로 재배포하지 않았을 수 있습니다.")
                except Exception as e:
                    st.error(f"❌ 연결 실패: {e}")

        st.markdown("**데이터**")
        if not exam_db.empty:
            st.caption(f"📊 기출 DB {len(exam_db):,}건 · {exam_db['_uk'].nunique()}개 대학"
                       + (f" ({', '.join(exam_db_files)})" if len(exam_db_files) > 1 else ""))
        if univ_info_db.empty:
            st.caption("📋 대학별 면접 형식 정보(univ_interview_info.csv)가 없습니다.")
        else:
            st.caption(f"📋 2026 면접 형식 {univ_info_db['대학'].nunique()}개 대학 · {len(univ_info_db)}개 학과")

        st.markdown("**AI 모델**")
        st.caption(f"🤖 최근 사용: {st.session_state.get('gemini_last_model') or '아직 없음'} (최신 모델 자동 선택)")
        if api_key and st.button("🩺 AI 모델 연결 점검", width="stretch", key="gemini_diag_btn",
                                 help="최신 모델부터 아주 짧은 요청을 보내, 이 API 키로 실제로 쓸 수 있는 모델을 확인합니다."):
            with st.spinner("모델마다 짧은 요청을 보내 확인하는 중입니다..."):
                st.session_state["gemini_diag"] = diagnose_gemini_models(api_key)
        if st.session_state.get("gemini_diag"):
            st.dataframe(pd.DataFrame(st.session_state["gemini_diag"]), hide_index=True, width="stretch")
            st.caption('✅ 모델 중 맨 위의 것이 자동으로 쓰입니다. 고정하려면 Secrets에 GEMINI_MODEL = "모델이름"')

    with st.popover("📖 사용 설명서", width="stretch"):
        st.markdown("""
        ### 🔑 Google Gemini API 키 발급 방법
        1. **Google AI Studio 접속:** [Google AI Studio](https://aistudio.google.com/)에 접속합니다.
        2. **구글 계정 로그인:** 평소 사용하는 구글 계정으로 로그인합니다.
        3. **API 키 생성:** 좌측 상단 'Get API key' 버튼을 클릭하여 키 생성 후 복사합니다.

        ### 📂 [중요] 생기부 PDF는 반드시 '텍스트 추출(OCR)'된 파일이어야 합니다!
        * **왜 필요한가요?** 단순 이미지(스캔본) PDF는 AI가 글자를 읽지 못하므로, **마우스로 글자가 드래그되거나 텍스트로 인식되는 PDF**여야만 정상 분석이 가능합니다.

        ### 🎙️ [신규] 휴대폰 음성 인식(STT) 면접 평가 기능 사용법!
        * **1단계 (키보드 활용):** 휴대폰으로 접속 시, 하단 채팅창을 누른 후 **휴대폰 키보드에 있는 '마이크(🎤)' 버튼**을 누르고 말하면 텍스트로 바로 입력됩니다.
        * **2단계 (무인 AI 면접관 모드):** 하단의 **[🎙️ 음성으로 면접 답변하기]** 버튼을 눌러 직접 녹음해 보세요. AI가 음성을 듣고 즉각적인 평가와 꼬리질문을 던져줍니다!

        ### 📊 전국 140개 대학 실제 면접 기출 데이터베이스 연동!
        * `master_interview_qa.csv` 파일을 이 앱과 같은 폴더에 넣어두면, 지원 학과와 가장 유사한 **실제 대학 기출 질의응답 약 1.6만 건**을 자동으로 찾아 문항 생성에 참고합니다.
        * 현재 DB 로딩 상태: **{db_status}**

        ### 💾 [신규] 학생 기록 자동 저장 및 불러오기
        * 문항을 한 번 생성하면 해당 학생의 **생기부 텍스트 · 생성된 문항 · 피드백 대화 내역**이 자동으로 저장됩니다.
        * 다음에 다시 접속했을 때는 왼쪽 사이드바 **'📂 학생 기록'**에서 이름을 선택해 불러오면, PDF를 다시 업로드하지 않고 이어서 진행할 수 있습니다.
        * ⚠️ 단, 이 저장 방식은 앱이 실행 중인 서버의 파일에 저장되는 방식입니다. 앱을 재배포(GitHub에 새로 커밋)하거나 서버가 완전히 재시작되면 저장된 기록이 초기화될 수 있으니, 중요한 학생 기록은 워드 파일로 다운로드해 별도 보관하시길 권장합니다.
        """.format(db_status=f"✅ {len(exam_db):,}건 로드 완료 ({exam_db['_uk'].nunique() if not exam_db.empty else 0}개 대학)" if not exam_db.empty else "⚠️ master_interview_qa.csv 파일을 찾지 못해 기본 학습 패턴만 사용 중입니다."))

render_class_practice_stats()

if st.session_state.get("growth_report_text") and not st.session_state.get("chat_history"):
    with st.expander(f"📈 {st.session_state.get('growth_report_for', '')} 학생 성장 리포트", expanded=True):
        st.markdown(st.session_state.growth_report_text)

def _reset_if_invalid(key, options):
    """선택지가 바뀌어 저장된 값이 더 이상 목록에 없으면 선택을 초기화합니다."""
    if key in st.session_state and st.session_state[key] not in options:
        del st.session_state[key]

col1, col2 = st.columns(2)
with col1:
    interview_type = st.radio("🎯 면접 방식", ["생기부 기반 면접", "상위권 대학 제시문 기반 면접"], horizontal=True, key="interview_type_radio")
    _db_unis = cached_db_universities(APP_DIR)
    _region_unis = build_region_universities(_db_unis)
    _regions = [r for r in UNI_REGION_ORDER if _region_unis.get(r)]
    if _region_unis.get("기타 대학"):
        _regions.append("기타 대학")
    if _db_unis:
        _region_unis[ALL_UNI_REGION] = _db_unis
        _regions.append(ALL_UNI_REGION)
    _regions.append(DIRECT_REGION)
    _reset_if_invalid("region_select", _regions)
    region = st.selectbox("📍 권역 선택", _regions, key="region_select")
    if region == DIRECT_REGION:
        uni = st.text_input("🏫 대학 직접 입력", value="한국대", key="uni_direct_input")
    else:
        _uni_list = _region_unis[region]
        _uni_counts = dict(_uni_list)
        _uni_options = [u for u, _ in _uni_list]
        _reset_if_invalid("uni_select", _uni_options)
        uni = st.selectbox(
            f"🏫 대학 선택 ({len(_uni_options)}개 · 기출 많은 순)", _uni_options,
            format_func=lambda u: f"{u} · 기출 {_uni_counts.get(u, 0):,}건" if _uni_counts.get(u, 0) else f"{u} · 기출 없음(타 대학 기출로 재구성)",
            key="uni_select",
        )
with col2:
    # 🎓 학과: 이 대학에 기출이 있는 학과를 먼저 보여주고, 없으면 전국 학과 목록이나 직접 입력으로 고릅니다.
    _depts = cached_uni_departments(APP_DIR, uni)
    _dept_counts = dict(_depts)
    _major_options = [d for d, _ in _depts] + [MAJOR_OTHER, MAJOR_DIRECT]
    _reset_if_invalid("major_select", _major_options)
    _major_choice = st.selectbox(
        f"🎓 지원 학과 선택 (이 대학 기출 보유 {len(_depts)}개 학과)" if _depts else "🎓 지원 학과 선택 (이 대학은 기출 보유 학과가 없습니다)",
        _major_options,
        format_func=lambda d: f"{d} · 기출 {_dept_counts[d]}건" if d in _dept_counts else d,
        key="major_select",
    )
    if _major_choice == MAJOR_DIRECT:
        major = st.text_input("🎓 학과 직접 입력", placeholder="예: 철학과", key="major_input")
    elif _major_choice == MAJOR_OTHER:
        _all_depts = [x for x in cached_all_departments(APP_DIR) if x[0] not in _dept_counts]
        _all_info = {d: (n, k) for d, n, k in _all_depts}
        _other_options = [d for d, _, _ in _all_depts]
        _reset_if_invalid("major_other_select", _other_options)
        major = st.selectbox(
            f"🔎 전국 학과 목록 ({len(_other_options)}개 · 기출 많은 순, 글자를 입력해 검색)", _other_options,
            format_func=lambda d: f"{d} · {_all_info[d][1]}개 대학 {_all_info[d][0]}건",
            key="major_other_select",
        ) if _other_options else ""
    else:
        major = _major_choice
    major = str(major or "").strip()
    student_name = st.text_input("👤 지원자 성명", placeholder="예: 홍길동", key="student_name_input").strip()
    difficulty = st.radio("⚙️ 난이도 선택", ["하 (기초)", "중 (표준)", "상 (압박)"], horizontal=True, index=1, key="difficulty_radio")

uploaded_file = None
if interview_type == "생기부 기반 면접":
    uploaded_file = st.file_uploader(
        "📂 학생 생기부 PDF 업로드 (OCR 변환 필수 · 위에서 기존 기록을 불러왔다면 다시 올리지 않아도 됩니다)",
        type=["pdf"]
    )
    if st.session_state.get("loaded_student_record_text") and not uploaded_file:
        st.info("📌 불러온 학생의 생기부 텍스트를 그대로 사용합니다. 다른 PDF를 새로 올리면 그것으로 대체됩니다.")

# 📋 선택한 대학의 실제 면접 형식 (2026학년도 합격생 후기 기준)
_info_row = get_univ_info(univ_info_db, uni, major)
if _info_row is not None:
    def _iv(key):
        v = _info_row.get(key, "")
        return "" if (v is None or str(v).strip().lower() in ("", "nan")) else str(v).strip()

    _head = " · ".join(x for x in [
        f"⏱️ {_iv('면접시간')}" if _iv('면접시간') else "",
        f"👥 면접위원 {_iv('면접위원')}" if _iv('면접위원') else "",
        f"📄 {_iv('면접유형')}" if _iv('면접유형') else "",
    ] if x)
    with st.expander(f"📋 {_iv('대학')} 실제 면접 형식 보기 (2026학년도 후기 · {_iv('학과')}) — {_head}", expanded=False):
        st.markdown(f"**전형:** {_iv('전형') or '-'}")
        if _iv("면접절차"):
            st.markdown(f"**면접 절차:** {_iv('면접절차')}")
        if _iv("유의사항"):
            st.markdown(f"**유의사항:** {_iv('유의사항')}")
        if _iv("선배조언"):
            st.markdown(f"**합격 선배의 조언:** {_iv('선배조언')}")
        st.caption("※ 2026학년도 면접 후기 자료에서 정리한 내용이며, 이 정보는 문항 생성 시 AI에게도 함께 전달되어 면접 시간에 맞는 분량으로 설계됩니다.")

# 📊 선택한 대학·학과가 실제로 무엇을 어떤 비중으로 물었는지 (기출 DB 직접 집계)
if str(major or "").strip() and not exam_db.empty:
    _preview_cases = cached_real_cases(APP_DIR, uni, major) if "제시문" in interview_type else None
    render_interview_profile(
        cached_interview_profile(APP_DIR, uni, major), uni, major, interview_type,
        _preview_cases, cached_reference_preview(APP_DIR, uni, major),
    )

st.markdown("---")

# -------------------------------------------------------------------------
# [4] 데이터 기반 프롬프트 및 생성 로직
# -------------------------------------------------------------------------
target_desc = f"{uni}_{major}"

# 🔥 선생님이 주신 20종 기출문제 빅데이터 패턴 완벽 통합 (프롬프트 주입용) 🔥
PAST_EXAM_DATA = """
[도개고 선배들의 20종 실제 합격 기출문제 데이터베이스 (학습용)]
AI는 아래의 실제 대입 기출문제들의 말투, 꼬리질문 방식, 생기부 파고들기 패턴을 완벽히 학습하여 문항을 생성해야 합니다.

<기출 패턴 1. 진위 여부, 실험 과정 및 세특 딥다이브>
- "사회문화 세특에서 사회복지 자율 연구 동아리를 결성하여 활동했다고 했는데, 이에 관해 구체적으로 설명해 줄 수 있나요?"
- "혈당량 측정 실험을 진행했는데 전체적인 실험의 과정과 이 연구의 시간 차이는 어땠는지 말해보세요."
- "(압박 꼬리질문) 그 실험을 진행할 때 도르래의 마찰과 같은 통제 변인을 고려했나요? 아니면 고려하지 않았나요?"

<기출 패턴 2. 개념 설명 및 기술적/논리적 문제 해결>
- "그렇다면 앱을 개발하면서 디바이스에 따른 화질 문제가 발생하였을 것인데 이는 기술적으로 어떻게 해결하였나요?"
- "빅데이터가 인공지능에 왜 중요하다고 생각하는지 본인의 탐구 내용과 엮어서 설명해볼까요?"
- "유전과 암 발병, 비타민 관계 등에서 다양한 탐구를 한 것 같은데, 상관관계를 먼저 설명하고 어떻게 예방할 수 있는지 말해볼까요?"

<기출 패턴 3. 독서 및 매체 연계 심화 검증>
- "발달장애 아동에 대해 탐구하면서 '뇌를 알면 아이가 보인다'를 읽고 뇌의 특징에 대해 알아보았다고 하는데 자세히 말해줄래요?"
- "“도시는 역사다”를 읽었는데 책에서 다룬 많은 도시 중 가장 인상 깊게 읽은 도시는 무엇인가요?"

<기출 패턴 4. 가치관 및 리더십>
- "부회장으로 활동하면서 의견수렴으로 힘들었던 것과 극복 사례를 말씀해 주시고, 많은 봉사를 할 수 있던 원동력에 대해 말씀해 주세요."
- "갈등이 발생했을 때 본인만의 대처 방법은 무엇이며, 그 안에서 학생이 생각하는 리더십이란 무엇이었나요?"
"""

TEMPLATE_SANGBU = """
### 🔍 [생기부 심층 분석 브리핑 리포트]
- **지원자 핵심 역량 및 강점 심층 분석:** (생기부 내 구체적인 활동명, 교과목, 에피소드를 근거로 전공 적합성과 학업 역량을 3~4문장으로 상세히 분석)
- **아쉬운 점 및 면접 방어 전략 (약점 분석):** (다소 부족한 부분이나 활동의 끊김, 성적 추이 등을 구체적으로 짚고, 면접 시 이를 공격받았을 때 방어할 논리적 전략을 상세히 제시)
- **면접관 집중 공략 대상 (최우선 심화 탐구 포인트):** (학생부에서 가장 수준 높은 탐구 보고서나 독서, 동아리 활동을 2개 이상 꼽고, 면접 전 반드시 복습해야 할 학술적/전공 개념 상세 안내)

### 📌 [영역: OOO] **[과목명/활동명]**
[질문]
(내용 작성)
[평가요소]
(이 질문이 학생부종합전형의 4대 평가요소 중 무엇을 보는지: 학업역량 / 전공적합성 / 인성 / 발전가능성 중 해당하는 것을 1~2개 골라 쉼표로 작성)
[평가의도]
(내용 작성)
[모범답안]
(내용 작성)
[꼬리질문]
(내용 작성)
"""

TEMPLATE_JESIMUN = """
### 📌 [세트 1] **[학술 및 전공 딜레마 주제 1]**
[제시문]
(첫 줄에 '※ 기출 응용 포인트: …' 한 문장을 쓰고, 이어서 응용 제시문 (가), (나), (다) 작성)
[문제 1]
(문제 1 내용)
[평가요소 1]
(학업역량 / 전공적합성 / 인성 / 발전가능성 중 해당하는 것을 1~2개 골라 쉼표로 작성)
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)
[문제 2]
(문제 2 내용)
[평가요소 2]
(학업역량 / 전공적합성 / 인성 / 발전가능성 중 해당하는 것을 1~2개 골라 쉼표로 작성)
[평가의도 2]
(내용 작성)
[모범답안 2]
(내용 작성)
[꼬리질문 2]
(내용 작성)

### 📌 [세트 2] **[학술 및 전공 딜레마 주제 2]**
[제시문]
(첫 줄에 '※ 기출 응용 포인트: …' 한 문장을 쓰고, 이어서 응용 제시문 (가), (나), (다) 작성)
[문제 1]
(문제 1 내용)
[평가요소 1]
(학업역량 / 전공적합성 / 인성 / 발전가능성 중 해당하는 것을 1~2개 골라 쉼표로 작성)
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)
[문제 2]
(문제 2 내용)
[평가요소 2]
(학업역량 / 전공적합성 / 인성 / 발전가능성 중 해당하는 것을 1~2개 골라 쉼표로 작성)
[평가의도 2]
(내용 작성)
[모범답안 2]
(내용 작성)
[꼬리질문 2]
(내용 작성)

### 📌 [세트 3] **[학술 및 전공 딜레마 주제 3]**
[제시문]
(첫 줄에 '※ 기출 응용 포인트: …' 한 문장을 쓰고, 이어서 응용 제시문 (가), (나), (다) 작성)
[문제 1]
(문제 1 내용)
[평가요소 1]
(학업역량 / 전공적합성 / 인성 / 발전가능성 중 해당하는 것을 1~2개 골라 쉼표로 작성)
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)
[문제 2]
(문제 2 내용)
[평가요소 2]
(학업역량 / 전공적합성 / 인성 / 발전가능성 중 해당하는 것을 1~2개 골라 쉼표로 작성)
[평가의도 2]
(내용 작성)
[모범답안 2]
(내용 작성)
[꼬리질문 2]
(내용 작성)
"""

EASY_EXPLAIN_TEMPLATE = """
### 📖 [영역: OOO] **[원문 속 활동명/과목명]**
**원문 요약:** (해당 부분의 생기부 원문을 짧게 그대로 인용하거나 간추림)
**쉬운 해설:** (고등학교 1학년이 읽어도 이해되도록, 어려운 전문 용어·개념은 괄호로 뜻을 풀어주고, 이 활동이 구체적으로 무엇을 한 것인지, 왜 의미가 있는지 친절하게 설명)
"""

def build_easy_explanation_prompt(student_record_text):
    return f"""
    당신은 고등학교 1학년 학생에게 생활기록부(생기부)를 눈높이에 맞춰 쉽게 설명해주는 친절한 담임 선생님입니다.
    아래 학생의 생기부 원문을 처음부터 끝까지, 항목(교과 세특, 창의적 체험활동, 동아리활동, 진로활동, 행동특성 및 종합의견, 독서활동 등) 순서를 따라가며
    학생 본인이 "아, 내가 이런 활동을 했고 이런 의미가 있었구나"를 이해할 수 있도록 해설을 붙여주세요.

    [지시사항]
    1. 원문을 요약만 하지 말고, 반드시 각 항목마다 "쉬운 해설"을 덧붙이세요.
    2. 전문 용어, 이론명, 어려운 한자어, 대회/프로그램 이름 등이 나오면 괄호 안에 고1 수준의 쉬운 뜻풀이를 넣어주세요. 예: "메타인지(자기 생각을 스스로 점검하고 조절하는 능력)"
    3. 너무 어린 말투나 반말은 쓰지 말고, 정중하지만 쉬운 문장으로 설명하세요.
    4. 원문에 실제로 있는 내용만 다루고, 없는 내용을 지어내지 마세요.
    5. 서론이나 인사말 없이 바로 '### 📖 [영역: ...]' 부터 출력하세요.

    [출력 템플릿 엄수 - 파싱을 위해 키워드는 그대로 유지]
    {EASY_EXPLAIN_TEMPLATE}

    [생기부 원문]
    {student_record_text}
    """

def build_sangbu_prompt(student_name, uni, major, student_record, combined_exam_data):
    return f"""
    당신은 도개고등학교의 진학 지도 노하우와 {uni} {major} 입학사정관의 시각을 겸비한 최고급 면접 출제위원입니다.
    지원자 '{student_name}' 학생의 생기부를 면밀히 분석하여 다음 작업을 수행하세요.

    {combined_exam_data}

    [지시사항]
    1. **출력의 맨 첫 부분**에 반드시 **[생기부 심층 분석 브리핑 리포트]**를 작성하세요. 단순 요약이 아닌, 실제 입학사정관의 눈으로 학생의 생기부를 현미경처럼 해부하여 구체적인 활동명과 과목명을 직접 언급하며 **매우 디테일하고 상세하게 분량 있게** 분석해야 합니다. 단점 방어 전략도 필수로 기재하세요.
    2. 생기부 5대 영역(교과세특, 창체, 동아리, 행특, 독서 등)을 모두 분석하여 총 5세트의 면접 문항을 만드세요.
    3. 과목명이나 주요 활동명은 반드시 **[생활과 윤리]** 처럼 볼드체로 묶어주고 학습된 기출 데이터 패턴 수준의 날카로운 꼬리질문을 포함하세요.
    4. 위 [참고자료 사용 지침]을 반드시 그대로 따르세요. A(지원 대학의 출제 스타일)는 **질문하는 방식·말투·길이·압박 수위**의 기준이고, B(지원 학과 전공 기출)는 **전공적합성과 학술적 깊이**의 기준입니다. 둘을 결합해 '이 대학이 실제로 묻는 방식으로, 이 학과 전공 수준에 맞게, 이 학생의 생기부 내용을 파고드는' 질문을 만드세요. A에 나온 다른 학과의 전공 내용은 절대 가져오지 마세요.
    5. 각 질문마다 [평가요소]에는 그 질문이 학생부종합전형 평가요소(학업역량/전공적합성/인성/발전가능성) 중 실제로 무엇을 검증하려는 질문인지 정확하게 판단해서 적으세요.
    6. 참고자료에 [C. 실제 출제 유형 분석]과 [5세트 유형 배분]이 있으면 그 배분을 그대로 따르세요. 5세트를 생기부 5대 영역에 고르게 걸치되, 각 세트의 질문 유형은 배분표대로 정하고 [평가의도] 첫 문장에 '(출제 유형: ○○형)'을 적으세요.

    [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
    {TEMPLATE_SANGBU}

    [생기부 내용]
    {student_record}
    """

# -------------------------------------------------------------------------
# 👥 여러 학생 한 번에 처리 (일괄 생성) — 방과후반처럼 여러 학생을 한 번에 지도할 때 사용
# -------------------------------------------------------------------------
if st.session_state.get("show_batch"):
    with st.expander("👥 여러 학생 한 번에 처리 (일괄 생성)", expanded=True):
        st.caption(
            "생기부 PDF 여러 개를 한 번에 올리면 학생별로 면접 패키지를 순서대로 생성하고, "
            "학생/교사용 워드 문서를 모두 묶어 zip 파일 하나로 내려받을 수 있습니다. "
            "(속도를 위해 일괄 처리에서는 '생기부 쉬운 해설'과 '평가요소 태깅'을 제외한 기본 문항만 생성하며, 각 학생 기록은 개별적으로도 자동 저장됩니다.)"
        )
        batch_files = st.file_uploader(
            "생기부 PDF 여러 개 선택 (여러 파일 선택 가능)", type=["pdf"], accept_multiple_files=True, key="batch_pdf_uploader"
        )
        if batch_files:
            default_rows = pd.DataFrame({
                "파일명": [f.name for f in batch_files],
                "학생명": [os.path.splitext(f.name)[0] for f in batch_files],
                "대학": [uni] * len(batch_files),
                "학과": [major if major else ""] * len(batch_files),
            })
            st.markdown("아래 표에서 학생명·대학·학과를 학생별로 확인/수정한 뒤 생성해 주세요.")
            edited_rows = st.data_editor(
                default_rows, key="batch_edit_table", width="stretch", hide_index=True,
                disabled=["파일명"]
            )
            batch_difficulty = st.radio("⚙️ 일괄 적용 난이도", ["하 (기초)", "중 (표준)", "상 (압박)"], horizontal=True, index=1, key="batch_difficulty_radio")

            if st.button("🚀 일괄 생성 시작", width="stretch"):
                if not api_key:
                    st.error("API 키를 입력해 주세요.")
                elif edited_rows["학생명"].isna().any() or (edited_rows["학생명"].astype(str).str.strip() == "").any():
                    st.error("모든 행의 학생명을 입력해 주세요.")
                elif edited_rows["학과"].isna().any() or (edited_rows["학과"].astype(str).str.strip() == "").any():
                    st.error("모든 행의 학과를 입력해 주세요.")
                else:
                    file_map = {f.name: f for f in batch_files}
                    zip_buffer = io.BytesIO()
                    progress = st.progress(0, text="일괄 생성을 시작합니다...")
                    success_count = 0
                    fail_names = []

                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        total = len(edited_rows)
                        for idx, row in edited_rows.reset_index(drop=True).iterrows():
                            b_name = str(row["학생명"]).strip()
                            b_uni = str(row["대학"]).strip() or uni
                            b_major = str(row["학과"]).strip()
                            progress.progress((idx) / total, text=f"({idx+1}/{total}) '{b_name}' 학생 처리 중...")
                            try:
                                b_file = file_map[row["파일명"]]
                                b_file.seek(0)
                                b_student_record = extract_text_from_pdf(b_file)

                                b_strategy = build_reference_strategy(exam_db, b_uni, b_major, top_n=12, style_n=8)
                                b_combined = PAST_EXAM_DATA + ("\n" + b_strategy["prompt_text"] if b_strategy["prompt_text"] else "")
                                b_info_text = format_univ_info_for_prompt(get_univ_info(univ_info_db, b_uni, b_major))
                                if b_info_text:
                                    b_combined += "\n" + b_info_text

                                b_prompt = build_sangbu_prompt(b_name, b_uni, b_major, b_student_record, b_combined)
                                b_result = call_gemini(b_prompt, api_key)

                                b_stu_path, b_tea_path = create_word_files(
                                    b_result, b_name, "생기부 기반 면접", f"{b_uni}_{b_major}"
                                )
                                zf.write(b_stu_path, arcname=os.path.basename(b_stu_path))
                                zf.write(b_tea_path, arcname=os.path.basename(b_tea_path))

                                # 학생별 인쇄용 평가표도 함께 묶어줌 (AI 호출 없이 양식만 생성)
                                try:
                                    b_eval_path = create_evaluation_sheet_word(
                                        b_result, b_name, b_uni, b_major, "생기부 기반 면접"
                                    )
                                    zf.write(b_eval_path, arcname=os.path.basename(b_eval_path))
                                except Exception:
                                    pass

                                b_record_id = new_record_id()
                                save_record(
                                    b_record_id, b_name, b_uni, b_major, "생기부 기반 면접", batch_difficulty,
                                    b_student_record, b_result,
                                    [{"role": "assistant", "content": b_result}]
                                )
                                success_count += 1
                            except Exception as e:
                                fail_names.append(f"{b_name} ({e})")
                            progress.progress((idx + 1) / total, text=f"({idx+1}/{total}) 처리 완료")

                    progress.empty()
                    if success_count > 0:
                        st.success(f"✅ {success_count}명 처리 완료! 아래에서 zip 파일로 한 번에 내려받으세요.")
                        st.download_button(
                            "📦 전체 학생 워드 문서 zip 다운로드", zip_buffer.getvalue(),
                            file_name="일괄_면접패키지.zip", mime="application/zip", width="stretch"
                        )
                    if fail_names:
                        st.error("❌ 다음 학생은 처리에 실패했습니다:\n" + "\n".join(fail_names))

if st.button("🚀 면접 패키지 생성 시작"):
    if not api_key: st.error("API 키를 입력해 주세요."); st.stop()
    if not major: st.error("지원 학과를 입력해 주세요."); st.stop()
    if not student_name: st.error("지원자 성명을 입력해 주세요. (문서 이름과 저장 기록을 학생별로 구분하는 데 쓰입니다)"); st.stop()
    if interview_type == "생기부 기반 면접" and not uploaded_file and not st.session_state.get("loaded_student_record_text"):
        st.error("생기부 파일을 업로드하거나, 위에서 저장된 기록을 먼저 불러와 주세요.")
        st.stop()

    if uploaded_file:
        with st.spinner("📄 PDF에서 텍스트를 확인하는 중입니다... (스캔본이면 자동으로 OCR을 시도합니다)"):
            student_record = extract_text_from_pdf(uploaded_file)
    else:
        student_record = st.session_state.get("loaded_student_record_text", "")
    st.session_state.loaded_student_record_text = student_record

    # 🔎 참고자료를 'A: 지원 대학 출제 스타일'과 'B: 지원 학과 전공 기출' 두 갈래로 준비
    is_sangbu_mode = interview_type == "생기부 기반 면접"
    strategy = build_reference_strategy(exam_db, uni, major, top_n=12, style_n=8, include_profile=is_sangbu_mode)
    relevant_examples = strategy["examples"]

    # 📚 제시문 면접: 실제 기출 사례를 DB에서 직접 골라 '먼저' 보여주고, 각 세트는 그 응용으로 만듭니다.
    real_cases, case_blocks = [], []
    if not is_sangbu_mode:
        real_cases = select_real_cases(exam_db, uni, major, n=3)
        case_blocks = [format_real_case_block(c, uni, major) for c in real_cases]

    # 📋 2026학년도 후기 기준 '그 대학의 실제 면접 형식'도 함께 반영 (시간·면접위원·유의사항)
    univ_info_row = get_univ_info(univ_info_db, uni, major)
    univ_info_text = format_univ_info_for_prompt(univ_info_row)
    st.session_state.last_univ_info = univ_info_row

    combined_exam_data = PAST_EXAM_DATA
    if strategy["prompt_text"]:
        combined_exam_data += "\n" + strategy["prompt_text"]
    if univ_info_text:
        combined_exam_data += "\n" + univ_info_text
    if not is_sangbu_mode:
        combined_exam_data += "\n" + format_real_cases_for_prompt(case_blocks, uni, major, n_sets=3)
    st.session_state.last_examples = relevant_examples
    st.session_state.last_strategy = {
        "mode": strategy["mode"],
        "style": strategy["style"],
        "style_examples": strategy["style_examples"],
        "has_same_dept_at_uni": strategy["has_same_dept_at_uni"],
        "alloc": strategy["alloc"],
        "case_sources": [
            f"{c['대학']} · {c['학과']} · {c['전형']} ({_CASE_TIER_LABEL.get(c['tier'], '')})" for c in real_cases
        ],
        # 생성 당시의 대학·학과를 함께 보관 (이후 드롭다운을 바꿔도 안내문이 어긋나지 않도록)
        "uni": uni,
        "major": major,
    }

    if interview_type == "생기부 기반 면접":
        prompt = f"""
        당신은 도개고등학교의 진학 지도 노하우와 {uni} {major} 입학사정관의 시각을 겸비한 최고급 면접 출제위원입니다.
        지원자 '{student_name}' 학생의 생기부를 면밀히 분석하여 다음 작업을 수행하세요.

        {combined_exam_data}

        [지시사항]
        1. **출력의 맨 첫 부분**에 반드시 **[생기부 심층 분석 브리핑 리포트]**를 작성하세요. 단순 요약이 아닌, 실제 입학사정관의 눈으로 학생의 생기부를 현미경처럼 해부하여 구체적인 활동명과 과목명을 직접 언급하며 **매우 디테일하고 상세하게 분량 있게** 분석해야 합니다. 단점 방어 전략도 필수로 기재하세요.
        2. 생기부 5대 영역(교과세특, 창체, 동아리, 행특, 독서 등)을 모두 분석하여 총 5세트의 면접 문항을 만드세요.
        3. 과목명이나 주요 활동명은 반드시 **[생활과 윤리]** 처럼 볼드체로 묶어주고 학습된 기출 데이터 패턴 수준의 날카로운 꼬리질문을 포함하세요.
        4. 위 [참고자료 사용 지침]을 반드시 그대로 따르세요. A(지원 대학의 출제 스타일)는 **질문하는 방식·말투·길이·압박 수위**의 기준이고, B(지원 학과 전공 기출)는 **전공적합성과 학술적 깊이**의 기준입니다. 둘을 결합해 '이 대학이 실제로 묻는 방식으로, 이 학과 전공 수준에 맞게, 이 학생의 생기부 내용을 파고드는' 질문을 만드세요. A에 나온 다른 학과의 전공 내용은 절대 가져오지 마세요.
        5. 각 질문마다 [평가요소]에는 그 질문이 학생부종합전형 평가요소(학업역량/전공적합성/인성/발전가능성) 중 실제로 무엇을 검증하려는 질문인지 정확하게 판단해서 적으세요. (형식적으로 아무거나 적지 말고, 질문 내용과 실제로 맞는 요소를 고르세요)
        6. 참고자료에 [C. 실제 출제 유형 분석]과 [5세트 유형 배분]이 있으면 그 배분을 그대로 따르세요. 5세트를 생기부 5대 영역에 고르게 걸치되, 각 세트의 질문 유형은 배분표대로 정하고 [평가의도] 첫 문장에 '(출제 유형: ○○형)'을 적으세요.

        [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
        {TEMPLATE_SANGBU}

        [생기부 내용]
        {student_record}
        """
    else:
        prompt = f"""
        당신은 {uni} {major}의 제시문 기반 면접·구술고사 출제위원입니다. {major} 전공적합성과 종합적 사고력, 논리적 추론 능력을 평가하는 제시문 면접 문항을 출제하세요.
        면접 난이도: {difficulty}

        {combined_exam_data}

        [지시사항]
        1. 생기부 내용은 무시하세요. **완전 독립된 3개의 세트**를 만드세요.
        2. 세트 1·2·3은 각각 위 D의 <기출 1>·<기출 2>·<기출 3>을 **응용한 문제**입니다. D의 [응용 규칙]을 반드시 지키세요. 대응하는 기출이 없는 세트는 {major}의 핵심 쟁점으로 창작하세요.
        3. **각 세트마다 복수의 제시문((가), (나), (다) 형태)과 [문제 1], [문제 2] (각각 평가요소, 평가의도, 모범답안, 압박 꼬리질문 포함)**가 유기적으로 묶여야 합니다. 제시문의 형식(글·자료·표나 그래프의 설명·수식)은 응용 대상 기출의 형식을 따르세요.
        4. 위 [참고자료 사용 지침]을 따르세요. A(지원 대학 출제 스타일)에서 **제시문·문제의 화법과 난이도**를, B(지원 학과 전공 기출)에서 **전공 개념의 깊이**를 가져오세요.
        5. 서론이나 인사말은 절대 쓰지 말고, 바로 '### 📌 [세트 1]' 부터 출력하세요.
        6. 각 문제마다 [평가요소 N]에는 그 문제가 학생부종합전형 평가요소(학업역량/전공적합성/인성/발전가능성) 중 실제로 무엇을 검증하려는지 정확하게 판단해서 적으세요.
        7. [모범답안 N]에는 결론만 쓰지 말고, 제시문의 어느 부분을 근거로 삼았는지를 함께 밝히세요.

        [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
        {TEMPLATE_JESIMUN}
        """

    with st.spinner(f"⏳ 로딩중... 실제 기출 DB {len(relevant_examples)}건을 참고하여 문항을 정밀 조립하고 있습니다."):
        try:
            result_text = call_gemini(prompt, api_key)
            if case_blocks:
                # 기출 원문은 AI가 아니라 DB에서 그대로 가져와 각 세트의 맨 앞에 끼워 넣습니다.
                result_text = inject_real_cases(result_text, case_blocks)
            st.session_state.last_result_text = result_text

            stu_path, tea_path = create_word_files(result_text, student_name, interview_type, target_desc)
            st.session_state.word_files = (stu_path, tea_path)

            # 📚 생기부 기반 면접일 때는 '생기부 쉬운 해설'과 '🖍️ 형광펜 생기부'를 한 번의 호출로 함께 생성
            if interview_type == "생기부 기반 면접" and student_record.strip():
                with st.spinner("📚 생기부 해설과 🖍️ 형광펜 표시를 만드는 중입니다..."):
                    run_explain_and_highlight(
                        student_record, result_text, strategy.get("profile"), uni, student_name, target_desc, api_key
                    )
            else:
                st.session_state.easy_explanation_text = ""
                st.session_state.easy_explanation_file = None
                st.session_state.highlight_data = None
                st.session_state.highlight_files = None
                st.session_state.highlight_raw_items = []
                st.session_state.highlight_error = ""

            # (요약카드는 워크북과 내용이 겹쳐 더 이상 만들지 않습니다)
            st.session_state.summary_card_file = None

            # 📝 교사용 인쇄 평가표 (AI 호출 없이 양식만 생성)
            try:
                st.session_state.evaluation_sheet_file = create_evaluation_sheet_word(
                    result_text, student_name, uni, major, interview_type
                )
            except Exception as e:
                st.session_state.evaluation_sheet_file = None
                st.warning(f"⚠️ 평가표 생성에 실패했습니다: {e}")

            # 새로 문항을 생성했으니 이전 학생의 성장 리포트 파일은 초기화
            st.session_state.growth_report_file = None
            st.session_state.interview_card = None
            st.session_state.interview_card_file = None
            st.session_state.interview_card_error = ""
            st.session_state.qsheet = None
            st.session_state.qsheet_files = None
            st.session_state.qsheet_error = ""
            st.session_state.qbank_sel = None
            st.session_state.qbank_runs = []
            st.session_state.run = None
            st.session_state.last_run = None
            st.session_state.pop("result_section", None)

            display_text = to_display_text(result_text)

            full_display_text = display_text + "\n\n---\n💬 **방금까지 나눈 문항 내용과 피드백 대화 내용을 한글 문서(.docx)로 만들어 드릴까요?** (원하시면 **'그래 만들어줘'**라고 말씀해 주세요!)"

            st.session_state.chat_history = [{"role": "assistant", "content": full_display_text}]

            # 📌 이 학생의 생기부·문항·대화를 저장해서, 다음에 다시 열어도 이어서 볼 수 있게 함
            #    ⚠️ 학생이 바뀌었는데 같은 기록 번호를 그대로 쓰면 앞 학생 기록을 덮어쓰게 되므로,
            #       (학생명·대학·학과·면접유형)이 달라지면 자동으로 '새 기록'을 만듭니다.
            #    같은 세션에서 학생을 바꿨다가 되돌아와도 원래 기록에 이어 쓰도록
            #    (학생 구분키 → 기록번호) 표를 세션에 들고 다닙니다.
            new_record_key = f"{student_name}|{uni}|{major}|{interview_type}"
            key_map = st.session_state.setdefault("record_key_map", {})
            if st.session_state.get("current_record_key") != new_record_key:
                st.session_state.current_record_id = key_map.get(new_record_key) or new_record_id()
            elif not st.session_state.get("current_record_id"):
                st.session_state.current_record_id = new_record_id()
            st.session_state.current_record_key = new_record_key
            key_map[new_record_key] = st.session_state.current_record_id
            st.session_state.loaded_student_record_text = student_record
            save_info = save_record(
                st.session_state.current_record_id, student_name, uni, major, interview_type, difficulty,
                student_record, result_text, history_for_save()
            )

            if save_info.get("ok") and "구글시트" in save_info.get("where", ""):
                st.success(
                    f"🎉 면접 패키지 생성 완료! ☁️ 학생 기록도 {save_info['where']}에 자동 저장되었습니다. ({save_info['at']})"
                )
            elif save_info.get("ok"):
                st.success("🎉 면접 패키지 생성 완료! 💾 학생 기록은 로컬에 임시 저장되었습니다.")
            else:
                st.warning(
                    "🎉 면접 패키지는 생성되었지만 ⚠️ **구글시트 자동 저장에 실패**해 로컬에만 임시 저장했습니다.\n\n"
                    f"사유: {save_info.get('error', '')}\n\n"
                    "아래 결과 화면의 **'💾 지금 저장하기'** 버튼으로 다시 시도해 보세요."
                )

        except Exception as e:
            st.error(f"❌ 생성 실패: {e}")

# 이번 생성에 실제로 참고된 기출 사례를 투명하게 보여줌
if st.session_state.get("last_examples") is not None and not st.session_state.last_examples.empty:
    _ex = st.session_state.last_examples
    _strat = st.session_state.get("last_strategy") or {}
    _mode = _strat.get("mode", "")
    # 생성 당시의 대학·학과를 사용 (생성 후 드롭다운을 바꿔도 안내문이 어긋나지 않게)
    _g_uni = _strat.get("uni") or uni
    _g_major = _strat.get("major") or major
    _mode_label = {
        "exact": "✅ 지원 대학·학과 기출 직접 활용",
        "style_transfer": "🔀 대학 스타일 + 타 대학 같은 학과 전공 결합",
        "dept_only": "📘 타 대학 같은 학과 기출 기반",
        "uni_only": "🏫 지원 대학 스타일 기반",
    }.get(_mode, "기출 참고")

    with st.expander(f"🔎 이번 문항 생성 참고자료 보기 — {_mode_label} (학과 기출 {len(_ex)}건)", expanded=False):
        _same_cnt = sum(1 for u in _ex["대학"] if _uni_matches(_g_uni, str(u)))
        _depts = sorted(set(str(d) for d in _ex["학과"]))

        if _mode == "exact":
            st.success(
                f"**{_g_uni} {_g_major}의 실제 기출이 DB에 있습니다.** 그 문항들의 화법·난이도를 그대로 기준 삼아 "
                f"이 학생의 생기부 내용으로 재창작하도록 지시했습니다.\n\n"
                f"- {_g_uni} 기출 **{_same_cnt}건** (⭐ 표시) / 타 대학 같은·유사 학과 {len(_ex) - _same_cnt}건\n"
                f"- 참고 학과: {', '.join(_depts)}"
            )
        elif _mode == "style_transfer":
            st.warning(
                f"**{_g_uni} {_g_major} 기출은 DB에 없습니다.** 그래서 두 갈래로 나눠 결합했습니다.\n\n"
                f"- **A. 출제 스타일** ← {_g_uni}의 다른 학과 기출에서 말투·질문 유형·압박 수위만 가져옴\n"
                f"- **B. 전공 깊이** ← 타 대학 {', '.join(_depts)} 기출 {len(_ex)}건에서 전공적합성 기준을 가져옴\n\n"
                f"→ AI에게 **'{_g_uni}의 말투로 묻는 {_g_major} 전공 질문'**을 만들도록 지시했습니다."
            )
        elif _mode == "uni_only":
            st.warning(
                f"**{_g_major}와 유사한 학과 기출이 DB에 없습니다.** {_g_uni}의 출제 스타일만 참고하고, "
                f"전공 내용은 고교 교육과정 연계 개념에서 직접 설계하도록 지시했습니다."
            )
        else:
            st.info(
                f"지원 학과({_g_major})와 같거나 가장 가까운 학과 기출을 사용했습니다. "
                f"({_g_uni} 기출 {_same_cnt}건 / 타 대학 {len(_ex) - _same_cnt}건)"
            )

        _alloc = _strat.get("alloc") or ([], "")
        if _alloc[0]:
            st.info(
                "📊 **실제 출제 비중에 맞춘 5세트 유형 배분** — "
                + " · ".join(f"{t} {k}세트" for t, k in _alloc[0]) + f"\n\n기준: {_alloc[1]}"
            )
        if _strat.get("case_sources"):
            st.info("📚 **먼저 제시한 실제 기출 사례**\n\n" + "\n".join(f"- {x}" for x in _strat["case_sources"]))

        # A. 대학 스타일 분석 결과
        _style = _strat.get("style")
        if _style:
            _types = " · ".join(f"{k} {v}%" for k, v in _style["대표유형"])
            st.markdown(
                f"##### 🏫 A. {_style['대학']} 출제 스타일 분석\n"
                f"기출 {_style['기출수']}건({_style['학과수']}개 학과) 분석 · 평균 질문 길이 약 {_style['평균질문길이']}자\n\n"
                f"**자주 쓰는 질문 유형:** {_types}"
            )
            _sx = _strat.get("style_examples")
            if _sx is not None and len(_sx):
                st.caption("이 대학의 실제 질문 화법 예시 (전공 내용이 아니라 '묻는 방식'만 참고)")
                for _, r in _sx.iterrows():
                    st.markdown(f"　· *[{r['학과']}]* {str(r['질문'])[:150]}")

        # B. 학과 전공 기출
        st.markdown(f"##### 📘 B. {_g_major} 전공 기출 ({len(_ex)}건)")
        for _, r in _ex.iterrows():
            _mark = "⭐ " if _uni_matches(_g_uni, str(r["대학"])) else ""
            st.markdown(f"{_mark}**[{r['대학']} · {r['학과']}]** {r['질문']}")
            st.caption(str(r['답변'])[:200] + ("..." if len(str(r['답변'])) > 200 else ""))

# -------------------------------------------------------------------------
# [5] 결과 대시보드 및 실시간 피드백
# -------------------------------------------------------------------------
# -------------------------------------------------------------------------
# [5-0] 결과 화면 도우미: 학생 정보 줄 · 내려받기 묶음 · 화면 목록
#   결과 화면은 한 줄로 길게 늘어놓지 않고, 위쪽 '화면 고르기'로 한 화면씩 봅니다.
#   (고른 화면만 그리므로 빠르고, 버튼을 눌러도 보던 화면이 유지됩니다)
# -------------------------------------------------------------------------
RESULT_SECTIONS_SB = ["🖍️ 형광펜·질문 카드", "🏦 질문 은행·질문지", "🎙️ 모의면접 진행", "🎤 연습·첨삭",
                      "📊 평가·리포트", "🗂️ 면접 전날"]
RESULT_SECTIONS_JS = ["📝 문항", "🎙️ 모의면접 진행", "🎤 연습·첨삭", "📊 평가·리포트"]

def _interview_dday():
    d = st.session_state.get("interview_date_input")
    try:
        gap = (d - datetime.date.today()).days
    except Exception:
        return ""
    return f"면접 D-{gap}" if gap > 0 else ("면접 당일" if gap == 0 and st.session_state.get("evaluation_sheet_file") else "")

def render_student_header(student_name, uni, major, interview_type, difficulty, is_sb, cards_on):
    """맨 위 학생 정보 줄: 누구의 어떤 면접인지, 어디까지 했는지, 저장 상태."""
    with st.container(border=True):
        left, right = st.columns([3, 2], vertical_alignment="center")
        with left:
            dday = _interview_dday()
            st.markdown(f"#### 🎓 {_md_safe(student_name or '학생')} · {_md_safe(uni or '')} {_md_safe(major or '')}"
                        + (f" · :red[{dday}]" if dday else ""))
            chips = [f"{'생기부' if is_sb else '제시문'} 면접 · 난이도 {str(difficulty).split(' ')[0]}", "📝 문항 ✅"]
            if is_sb:
                hl = st.session_state.get("highlight_data") or {}
                if cards_on:
                    keys = hl.get("q_keys") or []
                    checks = st.session_state.get("card_checks") or {}
                    done = sum(1 for k in keys if checks.get(k) == "자신 있음")
                    chips.append(f"🖍️ 형광펜 ✅ (자신 있음 {done}/{len(keys)})")
                else:
                    chips.append("🖍️ 형광펜 ⬜")
                chips.append("📋 질문지 " + ("✅" if st.session_state.get("qsheet") else "⬜"))
                _b = current_bank()
                if _b:
                    chips.append(f"🏦 은행 {sum(1 for e in _b if not e.get('dup_of'))}개 · 고른 {len(bank_selection(_b))}개")
                chips.append("🗂️ 요약카드 " + ("✅" if st.session_state.get("interview_card") else "⬜"))
            n_talk = sum(1 for m in st.session_state.get("chat_history", []) if m.get("role") == "user")
            chips.append(f"💬 연습·첨삭 {n_talk}회")
            n_run = len(st.session_state.get("qbank_runs") or []) + sum(
                1 for r in device_runs_cached(st.session_state.get("current_record_id")) if r.get("status") == "완료")
            if n_run:
                chips.append(f"🎙️ 모의면접 {n_run}회")
            st.caption(" · ".join(chips))
        with right:
            b1, b2 = st.columns(2)
            with b1:
                if st.button("💾 저장", width="stretch", key="manual_save_btn",
                             help="자동 저장되지만, 확실히 남겨 두고 싶을 때 누르세요."):
                    info = save_current_session(student_name, uni, major, interview_type, difficulty, quiet=True)
                    if info["ok"] and "구글시트" in info["where"]:
                        st.toast(f"{info['where']}에 저장했습니다. ({info['at']})", icon="✅")
                    elif info["ok"]:
                        st.toast(f"로컬에 임시 저장했습니다. ({info['at']}) — 구글시트 연동이 설정되지 않았습니다.", icon="💾")
                    else:
                        st.error(f"❌ 구글시트 저장 실패: {info['error']}  \n로컬에만 임시 저장했습니다. 사이드바 ⚙️ 관리에서 연결을 확인해 주세요.")
            with b2:
                if st.button("🆕 새 회차로", width="stretch", key="manual_save_new_btn",
                             help="같은 학생이라도 이번 연습을 '새 회차'로 따로 남깁니다. (성장 리포트용 회차가 쌓입니다)"):
                    st.session_state.current_record_id = new_record_id()
                    _new_key = f"{student_name}|{uni}|{major}|{interview_type}"
                    st.session_state.current_record_key = _new_key
                    st.session_state.setdefault("record_key_map", {})[_new_key] = st.session_state.current_record_id
                    info = save_current_session(student_name, uni, major, interview_type, difficulty, quiet=True)
                    if info["ok"]:
                        st.toast(f"새 회차로 저장했습니다. ({info['at']} · {info['where']})", icon="🆕")
                    else:
                        st.error(f"❌ 저장 실패: {info['error']}")
            render_save_status()

def _zip_bytes(items):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        used = set()
        for _, path in items:
            name = os.path.basename(path)
            if name in used:
                continue
            used.add(name)
            z.write(path, arcname=name)
    return buf.getvalue()

def collect_downloads(student_name, interview_type, is_sb):
    """(학생에게 줄 것, 교사용) 두 묶음으로 나눈 내려받기 목록."""
    stu, tea = [], []
    wf = st.session_state.get("word_files")
    if wf:
        stu_path, tea_path = wf
        # 형광펜 카드가 있으면 그것이 '생기부 문장 + 질문 + 5세트 문항'을 모두 합친 최종본입니다.
        hf = st.session_state.get("highlight_files")
        if hf and is_sb:
            stu_path, tea_path = hf[0], hf[1]
        stu.append(("📘 학생용 워크북", stu_path))
        tea.append(("📕 교사용 지침서", tea_path))
        if hf and is_sb and len(hf) > 2:
            stu.append(("📝 질문·답안 연습지", hf[2]))
    if is_sb:
        if st.session_state.get("easy_explanation_file"):
            stu.append(("📚 생기부 쉬운 해설", st.session_state.easy_explanation_file))
        if st.session_state.get("qsheet_files"):
            qs_stu, qs_tea = st.session_state.qsheet_files
            stu.append(("📋 면접 질문지 (학생용)", qs_stu))
            tea.append(("📋 면접 질문지 (교사용·모범답안)", qs_tea))
        if st.session_state.get("interview_card_file"):
            stu.append(("🗂️ 면접 전날 요약카드", st.session_state.interview_card_file))
    if st.session_state.get("evaluation_sheet_file"):
        tea.append(("📝 면접 평가표", st.session_state.evaluation_sheet_file))
    if st.session_state.get("growth_report_file") and st.session_state.get("growth_report_for") in (None, "", student_name):
        tea.append(("📈 성장 리포트", st.session_state.growth_report_file))
    # 피드백 대화는 실제로 주고받은 내용이 있을 때만
    if any(m.get("role") == "user" for m in st.session_state.get("chat_history", [])):
        try:
            tea.append(("💬 피드백 대화 내역", create_chat_history_word(st.session_state.chat_history, student_name)))
        except Exception:
            pass
    ok = lambda items: [(l, p) for l, p in items if p and os.path.exists(p)]
    return ok(stu), ok(tea)

def render_download_center(student_name, interview_type, is_sb):
    stu, tea = collect_downloads(student_name, interview_type, is_sb)
    safe = re.sub(r'[\\/:*?"<>|]', "_", student_name or "학생")
    c1, c2 = st.columns(2)
    for col, title, items, tag in ((c1, "📘 학생에게 줄 것", stu, "학생용"), (c2, "📕 교사용", tea, "교사용")):
        with col, st.container(border=True):
            st.markdown(f"**{title}** · {len(items)}개")
            if not items:
                st.caption("아직 만든 파일이 없습니다.")
                continue
            st.download_button(f"📦 {tag} 한 번에 받기 (.zip)", _zip_bytes(items),
                               file_name=f"{safe}_{tag}_자료.zip", mime="application/zip",
                               type="primary", width="stretch", key=f"dl_zip_{tag}")
            with st.expander("파일 하나씩 받기"):
                for i, (label, path) in enumerate(items):
                    with open(path, "rb") as f:
                        st.download_button(f"{label} (.docx)", f, file_name=os.path.basename(path),
                                           width="stretch", key=f"dl_{tag}_{i}_{os.path.basename(path)}")

if st.session_state.chat_history:
    st.markdown("<div class='step-text'>STEP 2 · 결과 확인 및 피드백</div>", unsafe_allow_html=True)
    _is_sb = interview_type == "생기부 기반 면접"
    _hl = st.session_state.get("highlight_data")
    _cards_on = _is_sb and bool((_hl or {}).get("items"))

    render_student_header(student_name, uni, major, interview_type, difficulty, _is_sb, _cards_on)
    render_download_center(student_name, interview_type, _is_sb)

    _sections = RESULT_SECTIONS_SB if _is_sb else RESULT_SECTIONS_JS
    _goto = st.session_state.pop("_goto_section", None)
    if _goto in _sections:
        st.session_state["result_section"] = _goto
    if st.session_state.get("result_section") not in _sections:
        st.session_state["result_section"] = st.session_state.get("_last_result_section") \
            if st.session_state.get("_last_result_section") in _sections else _sections[0]
    _sec = st.segmented_control("결과 화면", _sections, key="result_section", label_visibility="collapsed")
    _sec = _sec or st.session_state.get("_last_result_section") or _sections[0]
    st.session_state["_last_result_section"] = _sec

    _first = st.session_state.chat_history[0] if st.session_state.chat_history else {}
    _first_is_q = _first.get("role") == "assistant" and "### 📌" in str(_first.get("content", ""))

    # ── 🖍️ 형광펜 · 질문 카드 (생기부) / 📝 문항 (제시문) ─────────────────────
    if _sec in ("🖍️ 형광펜·질문 카드", "📝 문항"):
        if _is_sb:
            _hl = st.session_state.get("highlight_data")
            if interview_type == "생기부 기반 면접" and not (_hl and _hl.get("items")):
                with st.container(border=True):
                    st.markdown("**🖍️ 형광펜 생기부가 아직 없습니다**")
                    if st.session_state.get("highlight_error"):
                        st.caption(f"이유: {st.session_state.highlight_error}")
                    _has_record = bool(st.session_state.get("loaded_student_record_text", "").strip()) or bool(uploaded_file)
                    if not _has_record:
                        st.caption("생기부 원문이 없습니다. 위에서 생기부 PDF를 다시 올린 뒤 아래 버튼을 눌러 주세요.")
                    if st.button("🖍️ 형광펜·쉬운 해설 다시 만들기", disabled=not _has_record or not api_key,
                                 width="stretch", key="rerun_highlight"):
                        _rec = st.session_state.get("loaded_student_record_text", "")
                        if uploaded_file and not _rec.strip():
                            _rec = extract_text_from_pdf(uploaded_file)
                            st.session_state.loaded_student_record_text = _rec
                        with st.spinner("📚 생기부 해설과 🖍️ 형광펜 표시를 만드는 중입니다... (30초~1분)"):
                            if run_explain_and_highlight(
                                _rec, st.session_state.get("last_result_text", ""),
                                cached_interview_profile(APP_DIR, uni, major), uni, student_name, target_desc, api_key,
                            ):
                                # 다시 만든 형광펜도 기록에 남겨, 다음에 불러올 때 그대로 나오게 합니다.
                                save_current_session(student_name, uni, major, interview_type, difficulty, quiet=True)
                        st.rerun()

            if _cards_on:
                render_card_view(_hl, student_name, uni, major, interview_type, difficulty, target_desc, api_key)
            elif _first_is_q:
                with st.expander("📌 생성된 면접 문항", expanded=True):
                    st.markdown(_first["content"])
            if st.session_state.get("easy_explanation_text"):
                with st.expander("📚 생기부 쉬운 해설 미리보기 (고등학교 1학년 눈높이)", expanded=False):
                    st.markdown(st.session_state.easy_explanation_text)
        elif _first:
            st.markdown(_first.get("content", ""))

    # ── 📋 질문지 ─────────────────────────────────────────────────────────────
    elif _sec == "🏦 질문 은행·질문지":
        try:
            render_question_bank_section(student_name, uni, major, target_desc, api_key)
        except Exception as e:
            st.warning(f"질문 은행 화면을 그리지 못했습니다: {e}")

    # ── 🎙️ 모의면접 진행 ──────────────────────────────────────────────────────
    elif _sec == "🎙️ 모의면접 진행":
        try:
            render_run_mode(student_name, uni, major, interview_type, target_desc)
        except Exception as e:
            st.warning(f"모의면접 진행 화면을 그리지 못했습니다: {e}")

    # ── 🗂️ 면접 전날 ─────────────────────────────────────────────────────────
    elif _sec == "🗂️ 면접 전날":
        try:
            render_interview_card_section(student_name, uni, major, target_desc, api_key)
        except Exception as e:
            st.warning(f"요약카드 화면을 그리지 못했습니다: {e}")

    # ── 📊 평가 · 리포트 ──────────────────────────────────────────────────────
    elif _sec == "📊 평가·리포트":
        with st.container(border=True):
            st.markdown("**📝 교사용 면접 평가표**")
            st.caption(
                "실제 대학 학생부종합전형 면접 평가요소(학업역량·전공적합성·인성·발전가능성)와 면접 태도 평가 기준을 바탕으로 "
                "**항목당 10점 만점 / 총 100점** 채점표와 종합의견란이 들어간 A4 양식을 만듭니다. 인쇄해서 면접 현장에서 바로 채점하실 수 있습니다."
            )
            col_ev1, col_ev2 = st.columns(2)
            with col_ev1:
                evaluator_name = st.text_input("평가 교사 성명", value="", placeholder="예: 김기섭", key="evaluator_name_input")
            with col_ev2:
                interview_date_input = st.date_input("면접 일자", value=datetime.date.today(), key="interview_date_input")
            include_q = st.checkbox("생성된 면접 질문을 '문항별 평가란'에 함께 넣기", value=True, key="eval_include_questions")
            _lr = st.session_state.get("last_run")
            _lr = _lr if _lr and _lr.get("for") == student_name else None
            if _lr:
                st.caption(f"🎙️ {_lr.get('at', '')} 모의면접 진행 결과(질문 {len(_lr['rows'])}개 · 상·중·하·메모)가 문항별 평가란에 들어갑니다.")

            if st.button("📝 평가표에 반영하기", width="stretch"):
                try:
                    date_str = interview_date_input.strftime("%Y. %m. %d.") if interview_date_input else ""
                    st.session_state.evaluation_sheet_file = create_evaluation_sheet_word(
                        st.session_state.get("last_result_text", ""), student_name, uni, major, interview_type,
                        evaluator=evaluator_name.strip(), interview_date=date_str, include_questions=include_q,
                        run_rows=(_lr or {}).get("rows"), opinions=(_lr or {}).get("ops"),
                    )
                    st.toast("평가표에 반영했습니다. 위쪽 📕 교사용 묶음에서 받으세요.", icon="📝")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 평가표 생성에 실패했습니다: {e}")
        _runs = [dict(r, device="컴퓨터") for r in (st.session_state.get("qbank_runs") or [])] + [
            dict(r, grades=r.get("counts") or {}) for r in device_runs_cached(st.session_state.get("current_record_id"))
            if r.get("status") == "완료"]
        _runs.sort(key=lambda r: str(r.get("at", "")))
        if _runs:
            with st.container(border=True):
                st.markdown("**🎙️ 모의면접 진행 기록**")
                st.dataframe(pd.DataFrame([{
                    "회차": i, "일시": r.get("at", ""), "기기": r.get("device", "컴퓨터"), "질문": r.get("n", 0),
                    "상": r.get("grades", {}).get("상", 0), "중": r.get("grades", {}).get("중", 0),
                    "하": r.get("grades", {}).get("하", 0),
                    "걸린 시간": f"{int(r.get('sec', 0)) // 60}:{int(r.get('sec', 0)) % 60:02d}"}
                    for i, r in enumerate(_runs, 1)]), hide_index=True, width="stretch")
        _gr = st.session_state.get("growth_report_text")
        with st.container(border=True):
            st.markdown("**📈 학생 성장 리포트**")
            if _gr and st.session_state.get("growth_report_for") in (None, "", student_name):
                st.markdown(_gr)
            elif _gr:
                st.caption(f"지금 있는 리포트는 '{st.session_state.get('growth_report_for')}' 학생 것입니다. "
                           "이 학생 리포트는 왼쪽 📂 학생 기록에서 '📈 성장 리포트'를 눌러 만드세요.")
            else:
                st.caption("왼쪽 📂 학생 기록에서 학생을 고르고 '📈 성장 리포트'를 누르면 여기에 나옵니다. "
                           "저장된 회차와 휴대폰 연습 기록을 비교해 무엇이 늘었는지 정리합니다.")

    # ── 🎤 연습 · 첨삭 ────────────────────────────────────────────────────────
    elif _sec == "🎤 연습·첨삭":
        _bank_now = current_bank() if _is_sb else []
        if _is_sb and (_cards_on or _bank_now):
            try:
                render_teacher_practice(_hl, student_name, uni, major, bank=_bank_now)
            except Exception as e:
                st.warning(f"학생 연습 현황을 불러오지 못했습니다: {e}")
            st.divider()
        elif _is_sb:
            st.caption("📱 질문지나 형광펜 질문 카드를 만들면 학생 휴대폰 연습 링크도 여기서 만들 수 있습니다.")
        st.markdown("### 💬 답변 연습 · 문항 고치기")
        _talk = st.session_state.chat_history[1:] if _first_is_q else st.session_state.chat_history
        if not _talk:
            st.caption("아직 주고받은 내용이 없습니다. 녹음하거나, 써온 답변을 첨삭받거나, 아래 입력창에 문항 수정 요청을 적어 보세요.")
        for message in _talk:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        st.markdown("💡 **Tip:** 아래 채팅창에서 📱휴대폰 키보드 마이크(🎤)를 눌러 말하거나, 우측 🎙️ 녹음 버튼을 활용해 보세요!")

        audio_value = st.audio_input("🎙️ 음성으로 면접 답변하기 (녹음 버튼을 누르고 답변을 말해보세요!)")

        if audio_value is not None:
            if st.session_state.last_audio_size != audio_value.size:
                if not api_key:
                    st.error("API 키를 입력해 주세요.")
                else:
                    with st.spinner("AI 면접관이 학생의 음성을 분석하고 답변을 평가 중입니다..."):
                        audio_bytes = audio_value.read()
                        eval_result = call_gemini_audio_eval(audio_bytes, api_key)

                        st.session_state.chat_history.append({"role": "user", "content": "[🎙️ 음성 답변 제출 완료]"})
                        st.session_state.chat_history.append({"role": "assistant", "content": eval_result})

                        st.session_state.last_audio_size = audio_value.size

                        if st.session_state.get("current_record_id"):
                            save_record(
                                st.session_state.current_record_id, student_name, uni, major, interview_type, difficulty,
                                st.session_state.get("loaded_student_record_text", ""),
                                st.session_state.get("last_result_text", ""), st.session_state.chat_history
                            )

                        st.rerun()

        # -------------------------------------------------------------------
        # ✍️ 글로 써온 답변 첨삭 받기 (음성 없이, 미리 작성해온 답변을 첨삭)
        #    새 저장소를 따로 만들지 않고, 기존 chat_history/save_record를 그대로 재사용합니다.
        # -------------------------------------------------------------------
        with st.expander("✍️ 미리 써온 답변, 글로 첨삭받기", expanded=False):
            qa_pairs_for_feedback = extract_qa_pairs(st.session_state.get("last_result_text", ""))
            if not qa_pairs_for_feedback:
                st.caption("첨삭받을 문항이 없습니다. 먼저 면접 패키지를 생성해 주세요.")
            else:
                q_labels = [
                    (re.sub(r'\s+', ' ', p["q"]).strip()[:50] + ("…" if len(p["q"]) > 50 else "")) or f"질문 {i+1}"
                    for i, p in enumerate(qa_pairs_for_feedback)
                ]
                selected_q_idx = st.selectbox(
                    "첨삭받을 질문을 선택하세요", range(len(q_labels)),
                    format_func=lambda i: f"Q{i+1}. {q_labels[i]}", key="written_answer_q_select"
                )
                written_answer = st.text_area("학생이 직접 작성한 답변을 붙여넣어 주세요", height=150, key="written_answer_text")
                if st.button("✍️ 이 답변 첨삭 받기", width="stretch"):
                    if not api_key:
                        st.error("API 키를 입력해 주세요.")
                    elif not written_answer.strip():
                        st.error("첨삭받을 답변을 먼저 입력해 주세요.")
                    else:
                        with st.spinner("AI 면접관이 답변을 첨삭하는 중입니다..."):
                            selected_q = qa_pairs_for_feedback[selected_q_idx]["q"]
                            critique_prompt = f"""
                            당신은 대입 면접을 지도하는 날카로운 면접관입니다. 아래 면접 질문에 대해 학생이 직접 글로 작성해온 답변을 첨삭해 주세요.

                            [면접 질문]
                            {selected_q}

                            [학생이 작성한 답변]
                            {written_answer}

                            [지시사항]
                            1. **[논리성]**: 답변의 논리적 흐름이 자연스러운지, 비약은 없는지 평가하세요.
                            2. **[구체성]**: 추상적인 말에 그치지 않고 본인의 경험·탐구 내용을 구체적으로 담았는지 평가하세요.
                            3. **[전공 연계성]**: 지원 전공/학과와의 연결고리가 잘 드러나는지 평가하세요.
                            4. **[수정 제안]**: 위 문제점을 반영해 더 나은 답변 예시를 직접 다시 써서 보여주세요.
                            5. 칭찬만 늘어놓지 말고, 실제 입학사정관처럼 냉정하고 건설적으로 첨삭하세요.
                            """
                            critique_result = call_gemini(critique_prompt, api_key)

                        st.session_state.chat_history.append({
                            "role": "user",
                            "content": f"[✍️ 서술형 답변 제출]\nQ. {selected_q}\nA. {written_answer}"
                        })
                        st.session_state.chat_history.append({"role": "assistant", "content": critique_result})

                        if st.session_state.get("current_record_id"):
                            save_record(
                                st.session_state.current_record_id, student_name, uni, major, interview_type, difficulty,
                                st.session_state.get("loaded_student_record_text", ""),
                                st.session_state.get("last_result_text", ""), st.session_state.chat_history
                            )
                        st.rerun()

        if user_feedback := st.chat_input("질문을 더 어렵게 하거나 답변을 입력해보세요 (키보드 마이크🎤 활용 가능)"):
            st.session_state.chat_history.append({"role": "user", "content": user_feedback})
            with st.chat_message("user"):
                st.markdown(user_feedback)

            with st.chat_message("assistant"):
                with st.spinner("요청하신 내용을 처리 중입니다..."):
                    doc_request_words = ["만들어", "생성", "다운", "파일", "문서로", "저장", "그래", "응", "네", "해줘"]
                    is_doc_request = any(w in user_feedback for w in doc_request_words) and len(user_feedback.strip()) < 15

                    if is_doc_request:
                        chat_path = create_chat_history_word(st.session_state.chat_history, student_name)
                        response_text = "네! 지금까지 나눈 대화 내용을 깔끔한 워드 문서로 생성했습니다. 상단 또는 아래의 **'💬 피드백 대화 내역 (.docx)'** 다운로드 버튼을 클릭해 주세요!"
                        st.markdown(response_text)
                        st.session_state.chat_history.append({"role": "assistant", "content": response_text})
                        st.rerun()
                    else:
                        required_template = TEMPLATE_SANGBU if interview_type == "생기부 기반 면접" else TEMPLATE_JESIMUN
                        add_instruction = "\n(주의: 독립 세트를 최소 2세트 이상 추가로 더 생성해 주세요!)" if any(w in user_feedback for w in ["더", "추가", "많이", "늘려", "또"]) else ""

                        # 이전에 생성했던 문항 원문을 함께 넘겨서 "무엇을 수정해야 하는지" AI가 알 수 있게 함
                        # '[실제 기출]' 칸은 AI에게 보내지 않고 떼어 두었다가, 수정본에 그대로 다시 붙입니다.
                        previous_full = st.session_state.get("last_result_text", "")
                        previous_content = strip_real_cases(previous_full)

                        feedback_prompt = f"""
                        당신은 면접 출제위원입니다. 아래는 방금 전 당신이 생성했던 면접 문항 원문입니다.
                        사용자의 피드백에 맞게 이 내용을 수정/보완하되,
                        심도 있는 학술적/실천적 깊이를 유지하면서 **반드시 다음 템플릿 구조와 [키워드]를 토씨 하나 틀리지 말고 유지**해 주세요.
                        {add_instruction}

                        [직전에 생성했던 문항 원문]
                        {previous_content}

                        [강제 유지 템플릿]
                        {required_template}

                        사용자 피드백: "{user_feedback}"
                        """
                        try:
                            new_result = inject_real_cases(
                                call_gemini(feedback_prompt, api_key), extract_real_cases(previous_full)
                            )
                            st.session_state.last_result_text = new_result
                            display_text = to_display_text(new_result)

                            full_response = display_text + "\n\n---\n💬 **방금까지 나눈 문항 내용과 피드백 대화 내용을 한글 문서(.docx)로 만들어 드릴까요?** (원하시면 **'그래 만들어줘'**라고 말씀해 주세요!)"

                            st.markdown(full_response)
                            st.session_state.chat_history.append({"role": "assistant", "content": full_response})

                            stu_path, tea_path = create_word_files(new_result, student_name, interview_type, target_desc)
                            st.session_state.word_files = (stu_path, tea_path)
                            # 형광펜 워크북·지침서도 수정된 문항을 반영해 다시 만듭니다.
                            if st.session_state.get("highlight_data"):
                                try:
                                    set_highlight_state(st.session_state.highlight_data,
                                                        st.session_state.get("highlight_raw_items", []),
                                                        student_name, target_desc)
                                except Exception:
                                    pass
                            try:
                                st.session_state.evaluation_sheet_file = create_evaluation_sheet_word(
                                    new_result, student_name, uni, major, interview_type
                                )
                            except Exception:
                                pass

                            if st.session_state.get("current_record_id"):
                                save_record(
                                    st.session_state.current_record_id, student_name, uni, major, interview_type, difficulty,
                                    st.session_state.get("loaded_student_record_text", ""),
                                    new_result, st.session_state.chat_history
                                )

                            st.rerun()

                        except Exception as e:
                            st.error(f"피드백 반영 중 오류가 발생했습니다: {e}")
