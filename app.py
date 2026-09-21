import sys
import io

# 파이썬 표준 입출력 인코딩을 UTF-8로 완전히 강제 재설정
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import streamlit as st
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

# -------------------------------------------------------------------------
# [0] 실제 대학 면접 기출 통합 DB (전국 94개 대학 / 약 1.1만 개 질의응답)
#     master_interview_qa.csv 파일을 app.py와 같은 폴더에 두면 자동으로 로드됩니다.
# -------------------------------------------------------------------------
MASTER_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "master_interview_qa.csv")

@st.cache_data(show_spinner=False)
def load_exam_db(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=["대학", "학과", "전형", "질문", "답변", "원본파일"])
    try:
        df = pd.read_csv(path, encoding="utf-8-sig")
        df = df.dropna(subset=["질문", "답변"])
        return df
    except Exception:
        return pd.DataFrame(columns=["대학", "학과", "전형", "질문", "답변", "원본파일"])

def _normalize_dept(name):
    """'간호학과' -> '간호' 처럼 학과명에서 흔한 접미사를 제거해 비교하기 쉽게 만듭니다."""
    if not isinstance(name, str):
        return ""
    s = name.strip()
    for suf in ["학과", "학부", "전공", "과", "학"]:
        if s.endswith(suf) and len(s) > len(suf):
            return s[: -len(suf)]
    return s

def _normalize_uni(name):
    if not isinstance(name, str):
        return ""
    base = name.split("(")[0].strip()
    for suf in ["대학교", "대학"]:
        if base.endswith(suf) and len(base) > len(suf):
            base = base[: -len(suf)]
    return base

def get_relevant_examples(df, major, uni_name, top_n=12):
    """선택한 전공/대학과 가장 유사한 실제 기출 질의응답을 DB에서 골라옵니다."""
    if df.empty or not major:
        return df.head(0)

    df = df.copy()
    major_norm = _normalize_dept(major)
    df["_dept_norm"] = df["학과"].apply(_normalize_dept)

    exact = df[df["학과"] == major.strip()]
    partial = df[df["_dept_norm"].apply(lambda d: bool(d) and (d in major_norm or major_norm in d))]

    dept_pool = df["_dept_norm"].dropna().unique().tolist()
    close = difflib.get_close_matches(major_norm, dept_pool, n=6, cutoff=0.5)
    fuzzy = df[df["_dept_norm"].isin(close)]

    combined = pd.concat([exact, partial, fuzzy]).drop_duplicates(subset=["대학", "학과", "질문"])

    if combined.empty:
        # 전공 일치 항목이 전혀 없을 때는 실전 감각을 위해 DB 전체에서 무작위 샘플을 보여줍니다.
        combined = df.sample(min(top_n, len(df)), random_state=42)
    else:
        uni_key = _normalize_uni(uni_name)
        if uni_key:
            same_uni = combined[combined["대학"].apply(lambda u: uni_key in _normalize_uni(u))]
            rest = combined.drop(same_uni.index)
            combined = pd.concat([same_uni, rest])

    return combined.head(top_n)

def format_examples_for_prompt(examples_df):
    if examples_df.empty:
        return ""
    lines = ["\n[실전 데이터베이스: 지원 전공과 유사한 전국 대학 실제 면접 기출 사례]"]
    lines.append("(AI는 아래 실제 사례의 질문 난이도, 표현 방식, 꼬리질문 패턴을 참고하되 문장을 그대로 베끼지 말고 새로 창작하세요.)")
    for _, r in examples_df.iterrows():
        q = str(r["질문"]).strip().replace("\n", " ")
        a = str(r["답변"]).strip().replace("\n", " ")
        if len(a) > 220:
            a = a[:220] + "..."
        lines.append(f"- [{r['대학']} · {r['학과']}] Q: {q}\n  A: {a}")
    return "\n".join(lines)

# -------------------------------------------------------------------------
# [0-1] 학생별 저장/불러오기 저장소
#   - 생기부 텍스트, 생성된 문항, 대화 내역을 저장해두고 다음에 다시 열었을 때
#     PDF를 재업로드하지 않고 이어서 볼 수 있게 합니다.
#   - Streamlit Secrets에 구글 서비스 계정("gcp_service_account")이 등록되어 있으면
#     Google Sheets에 반영구적으로 저장합니다(앱이 재배포돼도 사라지지 않음).
#   - 등록되어 있지 않으면 예전처럼 로컬 SQLite 파일에 저장합니다(재배포 시 초기화될 수 있음).
# -------------------------------------------------------------------------
RECORD_COLUMNS = [
    "id", "student_name", "university", "major", "interview_type",
    "difficulty", "student_record_text", "result_text", "chat_history", "updated_at"
]
RECORDS_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "interview_records.db")
GOOGLE_SHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def _sheets_enabled():
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False

@st.cache_resource(show_spinner=False)
def _get_records_worksheet():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=GOOGLE_SHEET_SCOPES)
    client = gspread.authorize(creds)
    sheet_name = st.secrets.get("GOOGLE_SHEET_NAME", "도개고_면접기록_DB")
    try:
        sh = client.open(sheet_name)
    except gspread.SpreadsheetNotFound:
        sh = client.create(sheet_name)
    try:
        ws = sh.worksheet("records")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="records", rows=2000, cols=len(RECORD_COLUMNS))
        ws.append_row(RECORD_COLUMNS)
    return ws

# ---- SQLite 백업 저장소 (Google Sheets 미설정 시 사용) ----
def init_records_db():
    conn = sqlite3.connect(RECORDS_DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id TEXT PRIMARY KEY, student_name TEXT, university TEXT, major TEXT,
            interview_type TEXT, difficulty TEXT, student_record_text TEXT,
            result_text TEXT, chat_history TEXT, updated_at TEXT
        )
    """)
    conn.commit()
    conn.close()

def _sqlite_save_record(record_id, student_name, university, major, interview_type, difficulty,
                         student_record_text, result_text, chat_history):
    conn = sqlite3.connect(RECORDS_DB_PATH)
    conn.execute("""
        INSERT INTO records (id, student_name, university, major, interview_type, difficulty,
                              student_record_text, result_text, chat_history, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            student_name=excluded.student_name, university=excluded.university, major=excluded.major,
            interview_type=excluded.interview_type, difficulty=excluded.difficulty,
            student_record_text=excluded.student_record_text, result_text=excluded.result_text,
            chat_history=excluded.chat_history, updated_at=excluded.updated_at
    """, (
        record_id, student_name, university, major, interview_type, difficulty,
        student_record_text, result_text, json.dumps(chat_history, ensure_ascii=False),
        datetime.datetime.now().isoformat(timespec="seconds")
    ))
    conn.commit()
    conn.close()

def _sqlite_list_records():
    conn = sqlite3.connect(RECORDS_DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT id, student_name, university, major, interview_type, updated_at FROM records ORDER BY updated_at DESC"
    ).fetchall()
    conn.close()
    return rows

def _sqlite_get_record(record_id):
    conn = sqlite3.connect(RECORDS_DB_PATH)
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT * FROM records WHERE id = ?", (record_id,)).fetchone()
    conn.close()
    return row

def _sqlite_delete_record(record_id):
    conn = sqlite3.connect(RECORDS_DB_PATH)
    conn.execute("DELETE FROM records WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()

# ---- Google Sheets 저장소 ----
def _sheets_all_values():
    ws = _get_records_worksheet()
    return ws, ws.get_all_values()

def _sheets_save_record(record_id, student_name, university, major, interview_type, difficulty,
                         student_record_text, result_text, chat_history):
    ws, all_values = _sheets_all_values()
    row_data = [
        record_id, student_name, university, major, interview_type, difficulty,
        student_record_text, result_text, json.dumps(chat_history, ensure_ascii=False),
        datetime.datetime.now().isoformat(timespec="seconds")
    ]
    row_idx = None
    for i, row in enumerate(all_values[1:], start=2):
        if row and row[0] == record_id:
            row_idx = i
            break
    if row_idx:
        ws.update(f"A{row_idx}:J{row_idx}", [row_data])
    else:
        ws.append_row(row_data)

def _sheets_list_records():
    _, all_values = _sheets_all_values()
    records = [dict(zip(RECORD_COLUMNS, row)) for row in all_values[1:] if row]
    records.sort(key=lambda r: r.get("updated_at", ""), reverse=True)
    return records

def _sheets_get_record(record_id):
    _, all_values = _sheets_all_values()
    for row in all_values[1:]:
        if row and row[0] == record_id:
            return dict(zip(RECORD_COLUMNS, row))
    return None

def _sheets_delete_record(record_id):
    ws, all_values = _sheets_all_values()
    for i, row in enumerate(all_values[1:], start=2):
        if row and row[0] == record_id:
            ws.delete_rows(i)
            return

# ---- Google Apps Script 웹앱 저장소 (GCP 콘솔/서비스 계정 없이 구글 시트만으로 연동) ----
def _appsscript_enabled():
    try:
        return "APPS_SCRIPT_URL" in st.secrets and "APPS_SCRIPT_TOKEN" in st.secrets
    except Exception:
        return False

def _appsscript_save_record(record_id, student_name, university, major, interview_type, difficulty,
                             student_record_text, result_text, chat_history):
    payload = {
        "token": st.secrets["APPS_SCRIPT_TOKEN"], "action": "save",
        "id": record_id, "student_name": student_name, "university": university, "major": major,
        "interview_type": interview_type, "difficulty": difficulty,
        "student_record_text": student_record_text, "result_text": result_text,
        "chat_history": json.dumps(chat_history, ensure_ascii=False),
        "updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    resp = requests.post(st.secrets["APPS_SCRIPT_URL"], json=payload, timeout=30)
    resp.raise_for_status()

def _appsscript_list_records():
    resp = requests.get(st.secrets["APPS_SCRIPT_URL"], params={"token": st.secrets["APPS_SCRIPT_TOKEN"]}, timeout=30)
    resp.raise_for_status()
    records = resp.json()
    if isinstance(records, dict) and records.get("error"):
        raise RuntimeError(records["error"])
    records.sort(key=lambda r: r.get("updated_at", ""), reverse=True)
    return records

def _appsscript_get_record(record_id):
    resp = requests.get(
        st.secrets["APPS_SCRIPT_URL"],
        params={"token": st.secrets["APPS_SCRIPT_TOKEN"], "id": record_id}, timeout=30
    )
    resp.raise_for_status()
    records = resp.json()
    if isinstance(records, dict) and records.get("error"):
        raise RuntimeError(records["error"])
    return records[0] if records else None

def _appsscript_delete_record(record_id):
    payload = {"token": st.secrets["APPS_SCRIPT_TOKEN"], "action": "delete", "id": record_id}
    resp = requests.post(st.secrets["APPS_SCRIPT_URL"], json=payload, timeout=30)
    resp.raise_for_status()

# ---- 공용 인터페이스: Secrets 설정 여부에 따라 자동으로 저장소를 선택 ----
#   우선순위: Apps Script(가장 간단, GCP 콘솔 불필요) > Google Sheets 서비스 계정 > 로컬 SQLite
def save_record(*args, **kwargs):
    if _appsscript_enabled():
        try:
            return _appsscript_save_record(*args, **kwargs)
        except Exception as e:
            st.warning(f"⚠️ Apps Script 저장에 실패해 로컬에만 임시 저장합니다: {e}")
            return _sqlite_save_record(*args, **kwargs)
    if _sheets_enabled():
        try:
            return _sheets_save_record(*args, **kwargs)
        except Exception as e:
            st.warning(f"⚠️ Google Sheets 저장에 실패해 로컬에만 임시 저장합니다: {e}")
    return _sqlite_save_record(*args, **kwargs)

def list_records():
    if _appsscript_enabled():
        try:
            return _appsscript_list_records()
        except Exception as e:
            st.warning(f"⚠️ Apps Script 조회에 실패했습니다: {e}")
            return []
    if _sheets_enabled():
        try:
            return _sheets_list_records()
        except Exception as e:
            st.warning(f"⚠️ Google Sheets 조회에 실패했습니다: {e}")
            return []
    return _sqlite_list_records()

def get_record(record_id):
    if _appsscript_enabled():
        try:
            return _appsscript_get_record(record_id)
        except Exception as e:
            st.warning(f"⚠️ Apps Script 조회에 실패했습니다: {e}")
            return None
    if _sheets_enabled():
        try:
            return _sheets_get_record(record_id)
        except Exception as e:
            st.warning(f"⚠️ Google Sheets 조회에 실패했습니다: {e}")
            return None
    return _sqlite_get_record(record_id)

def delete_record(record_id):
    if _appsscript_enabled():
        try:
            return _appsscript_delete_record(record_id)
        except Exception as e:
            st.warning(f"⚠️ Apps Script 삭제에 실패했습니다: {e}")
            return
    if _sheets_enabled():
        try:
            return _sheets_delete_record(record_id)
        except Exception as e:
            st.warning(f"⚠️ Google Sheets 삭제에 실패했습니다: {e}")
            return
    return _sqlite_delete_record(record_id)

if not (_appsscript_enabled() or _sheets_enabled()):
    init_records_db()

# -------------------------------------------------------------------------
# [1] 스마트 PDF 및 제미나이 통신 함수 (텍스트 및 오디오)
# -------------------------------------------------------------------------
def extract_text_from_pdf(uploaded_file):
    """
    PDF에서 텍스트를 추출합니다.
    - 텍스트 레이어가 있는 일반 PDF(나이스 출력본, 한글/워드로 저장한 PDF 등)는 바로 추출됩니다.
    - 스캔본(이미지로만 된) PDF라서 추출된 글자가 거의 없으면, 서버에 Tesseract OCR이 설치되어
      있는 경우 자동으로 페이지를 이미지로 렌더링해 OCR을 시도합니다. 그래서 선생님이 미리
      수동으로 OCR 변환을 해두지 않아도 됩니다.
    """
    doc = pymupdf.open(stream=uploaded_file.read(), filetype="pdf")
    full_text = ""
    for page in doc:
        full_text += page.get_text()

    # 페이지당 평균 글자 수가 너무 적으면 스캔본(이미지) PDF로 판단하고 자동 OCR 시도
    avg_chars_per_page = len(full_text.strip()) / max(len(doc), 1)
    if avg_chars_per_page < 30:
        try:
            import pytesseract
            from PIL import Image

            ocr_text = ""
            for page in doc:
                pix = page.get_pixmap(dpi=300)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                ocr_text += pytesseract.image_to_string(img, lang="kor+eng")

            if len(ocr_text.strip()) > len(full_text.strip()):
                st.info("📸 스캔본(이미지) PDF로 보여 자동으로 OCR 텍스트 인식을 실행했습니다.")
                full_text = ocr_text
        except Exception:
            # Tesseract가 서버에 설치되어 있지 않거나 OCR이 실패한 경우:
            # 원래 추출된 텍스트(비어 있을 수 있음)라도 그대로 사용합니다.
            if not full_text.strip():
                st.warning(
                    "⚠️ 이 PDF는 텍스트 레이어가 없는 스캔본으로 보이는데, 서버에 OCR 엔진이 "
                    "설치되어 있지 않아 자동 인식에 실패했습니다. packages.txt에 tesseract-ocr을 "
                    "추가했는지 확인해 주세요."
                )
    return full_text

def call_gemini(prompt, api_key):
    if isinstance(prompt, bytes):
        prompt = prompt.decode('utf-8', errors='ignore')
    elif not isinstance(prompt, str):
        prompt = str(prompt)

    client = genai.Client(api_key=api_key)
    available_models = []
    try:
        for m in client.models.list():
            name = m.name.replace("models/", "") if m.name.startswith("models/") else m.name
            if "gemini-1.5" in name or "gemini-flash" in name:
                available_models.append(name)
    except:
        available_models = ["gemini-1.5-flash", "gemini-1.5-pro"]

    if not available_models: available_models = ["gemini-1.5-flash-latest"]
    models_to_try = sorted(available_models, key=lambda x: "flash" not in x)

    last_error = ""
    for target_model in models_to_try[:3]:
        try:
            response = client.models.generate_content(model=target_model, contents=prompt)
            return response.text
        except Exception as e:
            last_error = str(e)
            time.sleep(2)
    raise Exception(f"AI 모델 통신 실패 (마지막 에러: {last_error})")

def call_gemini_audio_eval(audio_bytes, api_key):
    client = genai.Client(api_key=api_key)
    prompt = """
    당신은 도개고등학교의 날카롭고 전문적인 면접관입니다. 다음은 학생이 면접 질문에 대해 직접 스마트폰으로 녹음한 음성 답변입니다.
    아래 3가지 항목을 반드시 포함하여 분석 및 평가 리포트를 작성해 주세요.
    1. 🗣️ **[답변 내용 변환 (STT)]**: 학생의 음성을 텍스트로 정확하게 받아적어 주세요.
    2. 📊 **[면접관의 평가]**: 학생의 답변을 '논리성, 표현력, 전공적합성'을 기준으로 분석하고 종합 평가를 [상 / 중 / 하]로 매겨주세요.
    3. 🔥 **[추가 압박 꼬리질문]**: 학생의 답변 내용 중 논리적 비약이 있거나 더 깊이 파고들 만한 날카로운 꼬리질문을 하나 던져주세요.
    """
    available_models = []
    try:
        for m in client.models.list():
            name = m.name.replace("models/", "") if m.name.startswith("models/") else m.name
            if "gemini-1.5" in name or "gemini-flash" in name:
                available_models.append(name)
    except:
        pass
    if not available_models: available_models = ["gemini-1.5-flash-latest", "gemini-1.5-flash"]
    models_to_try = sorted(available_models, key=lambda x: "flash" not in x)

    last_error = ""
    for target_model in models_to_try[:3]:
        try:
            audio_part = types.Part.from_bytes(data=audio_bytes, mime_type='audio/wav')
            response = client.models.generate_content(model=target_model, contents=[audio_part, prompt])
            return response.text
        except Exception as e:
            last_error = str(e)
            time.sleep(2)
    return f"음성 분석 실패: {last_error}"

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

def add_intro_paragraphs(doc, text):
    for line in text.split('\n'):
        line = line.strip()
        if not line: continue
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.3
        if line.startswith("### 📄") or line.startswith("### 📌") or line.startswith("### 🔍") or line.startswith("### 📖"):
            run = p.add_run(line)
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(0, 51, 102)
        else:
            parts = line.split("**")
            for i, part in enumerate(parts):
                run = p.add_run(part)
                if i % 2 != 0: run.bold = True

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

                if js_text:
                    row_js = table.add_row()
                    set_cell_background(row_js.cells[0], "F4F6F9")
                    add_parsed_text_to_cell(row_js.cells[0], f"**[서울대 스타일 구술 제시문 (가, 나, 다)]**\n{js_text}")

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

    student_path = f"{student_name}_{target_desc}_{type_label}_학생용.docx"
    teacher_path = f"{student_name}_{target_desc}_{type_label}_교사용.docx"
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
    file_path = f"{student_name}_피드백_대화내역.docx"
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
    file_path = f"{student_name}_생기부_쉬운해설.docx"
    doc.save(file_path)
    return file_path

# -------------------------------------------------------------------------
# [2-1] 생성된 문항에서 [질문]/[문제 N] ↔ [모범답안]/[모범답안 N] 쌍만 순서대로 뽑아내기
#   (구글 시트 Apps Script의 _extractQAPairs와 동일한 로직의 파이썬 버전)
# -------------------------------------------------------------------------
def extract_qa_pairs(result_text):
    if not result_text:
        return []
    tag_pattern = re.compile(r'\[([^\]]+)\]')
    matches = list(tag_pattern.finditer(result_text))
    sections = []
    for idx, m in enumerate(matches):
        start = m.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(result_text)
        sections.append((m.group(1).strip(), result_text[start:end].strip()))

    pairs = []
    pending_q = None
    for tag, content in sections:
        if tag.startswith("질문") or tag.startswith("문제"):
            pending_q = content
        elif tag.startswith("모범답안"):
            pairs.append({"q": pending_q or "", "a": content})
            pending_q = None
    return pairs

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

    file_path = f"{student_name}_면접직전_요약카드.docx"
    doc.save(file_path)
    return file_path

# -------------------------------------------------------------------------
# [2-2] 여러 회차의 저장 기록을 모아 학생의 "성장 리포트"를 만들기
#   (새 저장소를 만들지 않고, 기존 list_records()/get_record()로 이미 쌓인 기록을 재사용합니다)
# -------------------------------------------------------------------------
def build_growth_report_prompt(student_name, sessions):
    session_blocks = []
    for i, s in enumerate(sessions, start=1):
        chat_summary = "\n".join(
            f"- {'학생/선생님' if m['role'] == 'user' else 'AI 면접관'}: {re.sub(r'[#*`]', '', m['content'])[:300]}"
            for m in s.get("chat_history", [])
        )
        session_blocks.append(
            f"[{i}회차 — {s.get('updated_at', '')}, {s.get('university', '')} {s.get('major', '')}]\n"
            f"{chat_summary if chat_summary else '(대화 기록 없음)'}"
        )
    joined = "\n\n".join(session_blocks)
    return f"""
    당신은 학생의 모의면접 연습 기록을 시간 순서대로 검토하고 성장 과정을 분석하는 입시 지도 전문가입니다.
    아래는 '{student_name}' 학생이 여러 차례에 걸쳐 진행한 모의면접 연습 기록(회차 순서대로)입니다.

    {joined}

    [지시사항]
    1. **[전반적인 성장 흐름]**: 회차를 거치며 답변의 논리성, 구체성, 자신감, 전공 이해도 등이 어떻게 달라졌는지 서술하세요.
    2. **[처음보다 좋아진 점]**: 구체적인 회차와 내용을 근거로 들어 설명하세요.
    3. **[아직 반복되는 약점]**: 여러 회차에 걸쳐 계속 나타나는 문제점이 있다면 짚어주세요.
    4. **[다음 연습에서 집중할 부분]**: 다음 모의면접에서 우선적으로 보완해야 할 부분을 2~3가지 구체적으로 제안하세요.
    5. 근거 없이 막연하게 칭찬하지 말고, 실제 기록에 있는 내용을 근거로 구체적으로 작성하세요. 기록이 부실해 판단하기 어려운 부분은 솔직하게 "판단하기 어렵다"고 밝히세요.
    6. 서론 없이 바로 '### 📈 [전반적인 성장 흐름]' 부터 출력하세요.
    """

def create_growth_report_word(report_text, student_name):
    doc = Document()
    set_document_font(doc)
    doc.add_heading(f"📈 [{student_name}] 학생 모의면접 성장 리포트", level=1)
    doc.add_paragraph("여러 차례의 모의면접 연습 기록을 바탕으로 정리한 성장 리포트입니다. 학부모 상담 자료로도 활용하실 수 있습니다.\n")
    add_intro_paragraphs(doc, report_text.replace("### 📈", "### 🔍"))
    file_path = f"{student_name}_성장리포트.docx"
    doc.save(file_path)
    return file_path

# -------------------------------------------------------------------------
# [3] 메인 UI 설정
# -------------------------------------------------------------------------
if "chat_history" not in st.session_state: st.session_state.chat_history = []
if "word_files" not in st.session_state: st.session_state.word_files = None
if "last_audio_size" not in st.session_state: st.session_state.last_audio_size = 0
if "last_result_text" not in st.session_state: st.session_state.last_result_text = ""
if "last_examples" not in st.session_state: st.session_state.last_examples = None
if "current_record_id" not in st.session_state: st.session_state.current_record_id = None
if "loaded_student_record_text" not in st.session_state: st.session_state.loaded_student_record_text = ""
if "easy_explanation_text" not in st.session_state: st.session_state.easy_explanation_text = ""
if "easy_explanation_file" not in st.session_state: st.session_state.easy_explanation_file = None
if "summary_card_file" not in st.session_state: st.session_state.summary_card_file = None
if "growth_report_file" not in st.session_state: st.session_state.growth_report_file = None

exam_db = load_exam_db(MASTER_DB_PATH)

with st.sidebar:
    # Streamlit Cloud의 Secrets(GEMINI_API_KEY)에 키가 등록되어 있으면 자동으로 사용하고,
    # 없을 경우에만 직접 입력창을 보여줍니다.
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("🔑 API 키가 자동으로 연결되었습니다.")
    except (KeyError, FileNotFoundError, AttributeError):
        api_key = st.text_input("🔑 Gemini API Key", type="password")

    if exam_db.empty:
        st.warning("⚠️ 실제 기출 DB(master_interview_qa.csv)가 없습니다.\n앱과 같은 폴더에 파일을 넣어주세요.")
    else:
        st.success(f"📊 실제 기출 DB 연동됨\n{len(exam_db):,}건 / {exam_db['대학'].nunique()}개 대학")

    if _appsscript_enabled():
        st.success("☁️ 학생 기록: Google Sheets(Apps Script)에 반영구 저장 중")
    elif _sheets_enabled():
        st.success("☁️ 학생 기록: Google Sheets(서비스 계정)에 반영구 저장 중")
    else:
        st.info("💾 학생 기록: 로컬 임시 저장 중 (재배포 시 초기화될 수 있음)")

# -------------------------------------------------------------------------
# 📂 저장된 학생 기록 불러오기 (PDF 재업로드 없이 이전 작업 이어서 하기)
# -------------------------------------------------------------------------
def _rec_get(row, key, default=""):
    """sqlite3.Row와 Apps Script/Sheets에서 온 dict를 모두 안전하게 다루기 위한 헬퍼.
    (Apps Script 백엔드는 생기부 원문·면접유형·난이도를 저장하지 않으므로 해당 키가 없을 수 있음)"""
    try:
        val = row[key]
        return val if val not in (None, "") else default
    except (KeyError, IndexError):
        return default

with st.expander("📂 저장된 학생 기록 불러오기 / 관리", expanded=False):
    saved_records = list_records()
    if not saved_records:
        st.caption("아직 저장된 기록이 없습니다. 문항을 한 번 생성하면 이 학생의 문항·대화가 자동으로 저장됩니다.")
    else:
        record_options = {
            f"{r['student_name']} · {r['university']}_{r['major']} — "
            f"{r['updated_at'][:16].replace('T', ' ')}": r["id"]
            for r in saved_records
        }
        selected_label = st.selectbox("불러올 기록을 선택하세요", list(record_options.keys()), key="record_select_box")
        col_load, col_delete, col_grow = st.columns(3)
        with col_load:
            if st.button("📥 이 기록 불러오기", use_container_width=True):
                row = get_record(record_options[selected_label])
                if row:
                    st.session_state["region_select"] = "직접 입력"
                    st.session_state["uni_direct_input"] = row["university"]
                    st.session_state["major_input"] = row["major"]
                    st.session_state["student_name_input"] = row["student_name"]
                    st.session_state["interview_type_radio"] = _rec_get(row, "interview_type", "생기부 기반 면접")
                    st.session_state["difficulty_radio"] = _rec_get(row, "difficulty", "중 (표준)")
                    st.session_state["loaded_student_record_text"] = _rec_get(row, "student_record_text", "")
                    st.session_state["last_result_text"] = row["result_text"] or ""
                    st.session_state["chat_history"] = json.loads(row["chat_history"]) if row["chat_history"] else []
                    st.session_state["current_record_id"] = row["id"]
                    # 생기부 쉬운 해설은 저장소에 보관하지 않으므로, 기록을 불러올 때는 일단 비워두고
                    # 필요하면 '면접 패키지 생성 시작'을 다시 눌러 새로 만들 수 있게 합니다.
                    st.session_state["easy_explanation_text"] = ""
                    st.session_state["easy_explanation_file"] = None
                    st.session_state["summary_card_file"] = None
                    st.session_state["growth_report_file"] = None
                    if row["result_text"]:
                        stu_path, tea_path = create_word_files(
                            row["result_text"], row["student_name"], _rec_get(row, "interview_type", "생기부 기반 면접"),
                            f"{row['university']}_{row['major']}"
                        )
                        st.session_state["word_files"] = (stu_path, tea_path)
                    if not _rec_get(row, "student_record_text", ""):
                        st.info("ℹ️ 이 저장 방식은 생기부 원문은 따로 저장하지 않습니다. 문항/대화 내용은 그대로 불러왔고, 생기부 재분석이 필요하면 PDF를 다시 업로드해주세요.")
                    st.success(f"'{row['student_name']}' 학생의 기록을 불러왔습니다.")
                    st.rerun()
        with col_delete:
            if st.button("🗑️ 이 기록 삭제", use_container_width=True):
                delete_record(record_options[selected_label])
                st.success("삭제되었습니다.")
                st.rerun()
        with col_grow:
            if st.button("📈 이 학생 성장 리포트", use_container_width=True):
                if not api_key:
                    st.error("API 키를 입력해 주세요.")
                else:
                    target_id = record_options[selected_label]
                    target_row = next((r for r in saved_records if r["id"] == target_id), None)
                    target_name = target_row["student_name"] if target_row else None
                    matching = [r for r in saved_records if r.get("student_name") == target_name]
                    matching.sort(key=lambda r: r.get("updated_at", ""))
                    if len(matching) < 2:
                        st.warning(f"'{target_name}' 학생은 저장된 기록이 {len(matching)}개뿐이라 성장 비교가 어렵습니다. 연습을 몇 번 더 진행한 뒤 다시 시도해 주세요.")
                    else:
                        with st.spinner(f"'{target_name}' 학생의 {len(matching)}개 회차 기록을 분석해 성장 리포트를 작성하는 중입니다..."):
                            sessions = []
                            for r in matching[-8:]:  # 너무 길어지지 않도록 최근 8회차까지만
                                full = get_record(r["id"])
                                if not full:
                                    continue
                                try:
                                    chat_hist = json.loads(full["chat_history"]) if full["chat_history"] else []
                                except Exception:
                                    chat_hist = []
                                sessions.append({
                                    "updated_at": full["updated_at"],
                                    "university": full["university"],
                                    "major": full["major"],
                                    "chat_history": chat_hist,
                                })
                            try:
                                report_text = call_gemini(build_growth_report_prompt(target_name, sessions), api_key)
                                report_path = create_growth_report_word(report_text, target_name)
                                st.session_state.growth_report_file = report_path
                                st.success(f"'{target_name}' 학생의 성장 리포트가 생성되었습니다.")
                                with open(report_path, "rb") as f:
                                    st.download_button(
                                        "📈 성장 리포트 다운로드 (.docx)", f, file_name=report_path,
                                        use_container_width=True, key="dl_growth_report_inline"
                                    )
                                st.markdown(report_text)
                            except Exception as e:
                                st.error(f"❌ 성장 리포트 생성 실패: {e}")

with st.expander("📖 [클릭] 프로그램 사용 설명서 및 PDF OCR 변환 방법", expanded=False):
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

    ### 📊 [신규] 전국 94개 대학 실제 면접 기출 데이터베이스 연동!
    * `master_interview_qa.csv` 파일을 이 앱과 같은 폴더에 넣어두면, 지원 학과와 가장 유사한 **실제 대학 기출 질의응답 약 1.1만 건**을 자동으로 찾아 문항 생성에 참고합니다.
    * 현재 DB 로딩 상태: **{db_status}**

    ### 💾 [신규] 학생 기록 자동 저장 및 불러오기
    * 문항을 한 번 생성하면 해당 학생의 **생기부 텍스트 · 생성된 문항 · 피드백 대화 내역**이 자동으로 저장됩니다.
    * 다음에 다시 접속했을 때는 위쪽 **'📂 저장된 학생 기록 불러오기'**에서 이름을 선택해 불러오면, PDF를 다시 업로드하지 않고 이어서 진행할 수 있습니다.
    * ⚠️ 단, 이 저장 방식은 앱이 실행 중인 서버의 파일에 저장되는 방식입니다. 앱을 재배포(GitHub에 새로 커밋)하거나 서버가 완전히 재시작되면 저장된 기록이 초기화될 수 있으니, 중요한 학생 기록은 워드 파일로 다운로드해 별도 보관하시길 권장합니다.
    """.format(db_status=f"✅ {len(exam_db):,}건 로드 완료 ({exam_db['대학'].nunique() if not exam_db.empty else 0}개 대학)" if not exam_db.empty else "⚠️ master_interview_qa.csv 파일을 찾지 못해 기본 학습 패턴만 사용 중입니다."))

UNIVERSITIES = {
    "서울권": ["서울대", "연세대", "고려대", "성균관대", "서강대", "한양대", "중앙대", "경희대", "한국외대", "서울시립대", "이화여대"],
    "충청권": ["카이스트(KAIST)", "충남대", "충북대", "고려대(세종)"],
    "경상권": ["경북대", "부산대", "UNIST", "영남대", "계명대"]
}

col1, col2 = st.columns(2)
with col1:
    interview_type = st.radio("🎯 면접 방식", ["생기부 기반 면접", "상위권 대학 제시문 기반 면접"], horizontal=True, key="interview_type_radio")
    region = st.selectbox("📍 권역 선택", ["서울권", "충청권", "경상권", "직접 입력"], key="region_select")
    uni = st.text_input("🏫 대학 직접 입력", value="한국대", key="uni_direct_input") if region == "직접 입력" else st.selectbox("🏫 대학 선택", UNIVERSITIES[region], key="uni_select")
with col2:
    major = st.text_input("🎓 지원 학과/전공", placeholder="예: 철학과", key="major_input")
    student_name = st.text_input("👤 지원자 성명", value="김기섭", key="student_name_input")
    difficulty = st.radio("⚙️ 난이도 선택", ["하 (기초)", "중 (표준)", "상 (압박)"], horizontal=True, index=1, key="difficulty_radio")

uploaded_file = None
if interview_type == "생기부 기반 면접":
    uploaded_file = st.file_uploader(
        "📂 학생 생기부 PDF 업로드 (OCR 변환 필수 · 위에서 기존 기록을 불러왔다면 다시 올리지 않아도 됩니다)",
        type=["pdf"]
    )
    if st.session_state.get("loaded_student_record_text") and not uploaded_file:
        st.info("📌 불러온 학생의 생기부 텍스트를 그대로 사용합니다. 다른 PDF를 새로 올리면 그것으로 대체됩니다.")

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
(여기에 서울대 구술고사 스타일의 다중 제시문 (가), (나), (다) 작성)
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
(여기에 서울대 구술고사 스타일의 다중 제시문 (가), (나), (다) 작성)
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
(여기에 서울대 구술고사 스타일의 다중 제시문 (가), (나), (다) 작성)
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
    4. 위 [실전 데이터베이스]에 제시된 실제 사례가 있다면, 그 질문의 깊이와 화법을 반드시 참고하여 이 학생의 활동에 맞게 재창작하세요.
    5. 각 질문마다 [평가요소]에는 그 질문이 학생부종합전형 평가요소(학업역량/전공적합성/인성/발전가능성) 중 실제로 무엇을 검증하려는 질문인지 정확하게 판단해서 적으세요.

    [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
    {TEMPLATE_SANGBU}

    [생기부 내용]
    {student_record}
    """

# -------------------------------------------------------------------------
# 👥 여러 학생 한 번에 처리 (일괄 생성) — 방과후반처럼 여러 학생을 한 번에 지도할 때 사용
# -------------------------------------------------------------------------
with st.expander("👥 여러 학생 한 번에 처리 (일괄 생성)", expanded=False):
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
            default_rows, key="batch_edit_table", use_container_width=True, hide_index=True,
            disabled=["파일명"]
        )
        batch_difficulty = st.radio("⚙️ 일괄 적용 난이도", ["하 (기초)", "중 (표준)", "상 (압박)"], horizontal=True, index=1, key="batch_difficulty_radio")

        if st.button("🚀 일괄 생성 시작", use_container_width=True):
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

                            b_examples = get_relevant_examples(exam_db, b_major, b_uni, top_n=12)
                            b_dynamic_text = format_examples_for_prompt(b_examples)
                            b_combined = PAST_EXAM_DATA + ("\n" + b_dynamic_text if b_dynamic_text else "")

                            b_prompt = build_sangbu_prompt(b_name, b_uni, b_major, b_student_record, b_combined)
                            b_result = call_gemini(b_prompt, api_key)

                            b_stu_path, b_tea_path = create_word_files(
                                b_result, b_name, "생기부 기반 면접", f"{b_uni}_{b_major}"
                            )
                            zf.write(b_stu_path, arcname=os.path.basename(b_stu_path))
                            zf.write(b_tea_path, arcname=os.path.basename(b_tea_path))

                            b_record_id = str(uuid.uuid4())
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
                        file_name="일괄_면접패키지.zip", mime="application/zip", use_container_width=True
                    )
                if fail_names:
                    st.error("❌ 다음 학생은 처리에 실패했습니다:\n" + "\n".join(fail_names))

if st.button("🚀 면접 패키지 생성 시작"):
    if not api_key: st.error("API 키를 입력해 주세요."); st.stop()
    if not major: st.error("지원 학과를 입력해 주세요."); st.stop()
    if interview_type == "생기부 기반 면접" and not uploaded_file and not st.session_state.get("loaded_student_record_text"):
        st.error("생기부 파일을 업로드하거나, 위에서 저장된 기록을 먼저 불러와 주세요.")
        st.stop()

    if uploaded_file:
        with st.spinner("📄 PDF에서 텍스트를 확인하는 중입니다... (스캔본이면 자동으로 OCR을 시도합니다)"):
            student_record = extract_text_from_pdf(uploaded_file)
    else:
        student_record = st.session_state.get("loaded_student_record_text", "")
    st.session_state.loaded_student_record_text = student_record

    # 🔎 실제 기출 DB에서 지원 학과와 유사한 사례를 찾아 프롬프트에 결합
    relevant_examples = get_relevant_examples(exam_db, major, uni, top_n=12)
    dynamic_exam_text = format_examples_for_prompt(relevant_examples)
    combined_exam_data = PAST_EXAM_DATA + ("\n" + dynamic_exam_text if dynamic_exam_text else "")
    st.session_state.last_examples = relevant_examples

    if interview_type == "생기부 기반 면접":
        prompt = f"""
        당신은 도개고등학교의 진학 지도 노하우와 {uni} {major} 입학사정관의 시각을 겸비한 최고급 면접 출제위원입니다.
        지원자 '{student_name}' 학생의 생기부를 면밀히 분석하여 다음 작업을 수행하세요.

        {combined_exam_data}

        [지시사항]
        1. **출력의 맨 첫 부분**에 반드시 **[생기부 심층 분석 브리핑 리포트]**를 작성하세요. 단순 요약이 아닌, 실제 입학사정관의 눈으로 학생의 생기부를 현미경처럼 해부하여 구체적인 활동명과 과목명을 직접 언급하며 **매우 디테일하고 상세하게 분량 있게** 분석해야 합니다. 단점 방어 전략도 필수로 기재하세요.
        2. 생기부 5대 영역(교과세특, 창체, 동아리, 행특, 독서 등)을 모두 분석하여 총 5세트의 면접 문항을 만드세요.
        3. 과목명이나 주요 활동명은 반드시 **[생활과 윤리]** 처럼 볼드체로 묶어주고 학습된 기출 데이터 패턴 수준의 날카로운 꼬리질문을 포함하세요.
        4. 위 [실전 데이터베이스]에 제시된 실제 사례가 있다면, 그 질문의 깊이와 화법을 반드시 참고하여 이 학생의 활동에 맞게 재창작하세요.
        5. 각 질문마다 [평가요소]에는 그 질문이 학생부종합전형 평가요소(학업역량/전공적합성/인성/발전가능성) 중 실제로 무엇을 검증하려는 질문인지 정확하게 판단해서 적으세요. (형식적으로 아무거나 적지 말고, 질문 내용과 실제로 맞는 요소를 고르세요)

        [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
        {TEMPLATE_SANGBU}

        [생기부 내용]
        {student_record}
        """
    else:
        prompt = f"""
        당신은 서울대학교 면접 및 구술고사 출제위원입니다. {major} 전공적합성과 종합적 사고력, 논리적 추론 능력을 평가하기 위한 고난도 제시문 기반 구술고사를 출제하세요.
        면접 난이도: {difficulty}

        {combined_exam_data}

        [지시사항]
        1. 생기부 내용은 무시하세요. {major} 학과와 관련된 학술적 딜레마와 심층 개념을 담은 **완전 독립된 3개의 주제 세트**를 창작하세요.
        2. **각 세트마다 복수의 제시문((가), (나), (다) 형태)과 [문제 1], [문제 2] (각각 평가의도, 모범답안, 압박 꼬리질문 포함)**가 유기적으로 묶인 **총 3개의 독립 세트**를 엄격히 만드세요.
        3. 위 [실전 데이터베이스]의 실제 사례가 있다면 질문의 수준과 화법을 참고해 {major}에 맞게 새로 창작하세요.
        4. 서론이나 인사말은 절대 쓰지 말고, 바로 '### 📌 [세트 1]' 부터 출력하세요.
        5. 각 문제마다 [평가요소 N]에는 그 문제가 학생부종합전형 평가요소(학업역량/전공적합성/인성/발전가능성) 중 실제로 무엇을 검증하려는지 정확하게 판단해서 적으세요.

        [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
        {TEMPLATE_JESIMUN}
        """

    with st.spinner(f"⏳ 로딩중... 실제 기출 DB {len(relevant_examples)}건을 참고하여 문항을 정밀 조립하고 있습니다."):
        try:
            result_text = call_gemini(prompt, api_key)
            st.session_state.last_result_text = result_text

            stu_path, tea_path = create_word_files(result_text, student_name, interview_type, target_desc)
            st.session_state.word_files = (stu_path, tea_path)

            # 📚 생기부 기반 면접일 때는 학생이 스스로 읽을 수 있는 '생기부 쉬운 해설'도 함께 생성
            if interview_type == "생기부 기반 면접" and student_record.strip():
                try:
                    with st.spinner("📚 생기부 내용을 고등학교 1학년 눈높이로 해설하는 중입니다..."):
                        easy_explanation = call_gemini(build_easy_explanation_prompt(student_record), api_key)
                    st.session_state.easy_explanation_text = easy_explanation
                    st.session_state.easy_explanation_file = create_easy_explanation_word(
                        easy_explanation, student_name, target_desc
                    )
                except Exception as e:
                    st.session_state.easy_explanation_text = ""
                    st.session_state.easy_explanation_file = None
                    st.warning(f"⚠️ 생기부 쉬운 해설 생성에 실패했습니다: {e}")
            else:
                st.session_state.easy_explanation_text = ""
                st.session_state.easy_explanation_file = None

            # 🗂️ 면접 직전에 훑어볼 A4 한 장 요약카드 (별도 AI 호출 없이 방금 생성된 문항에서 바로 추출)
            try:
                st.session_state.summary_card_file = create_summary_card_word(
                    result_text, student_name, uni, major, interview_type
                )
            except Exception as e:
                st.session_state.summary_card_file = None
                st.warning(f"⚠️ 요약카드 생성에 실패했습니다: {e}")

            # 새로 문항을 생성했으니 이전 학생의 성장 리포트 파일은 초기화
            st.session_state.growth_report_file = None

            display_text = result_text.replace('[문제 1]', '\n**💡 [문제 1]**\n').replace('[평가요소 1]', '\n**🏷️ [평가 요소 1]**\n').replace('[평가의도 1]', '\n**🎯 [평가 의도 1]**\n').replace('[모범답안 1]', '\n**✅ [모범 답안 가이드 1]**\n').replace('[꼬리질문 1]', '\n**🔥 [압박용 꼬리질문 1]**\n')
            display_text = display_text.replace('[문제 2]', '\n**💡 [문제 2]**\n').replace('[평가요소 2]', '\n**🏷️ [평가 요소 2]**\n').replace('[평가의도 2]', '\n**🎯 [평가 의도 2]**\n').replace('[모범답안 2]', '\n**✅ [모범 답안 가이드 2]**\n').replace('[꼬리질문 2]', '\n**🔥 [압박용 꼬리질문 2]**\n')
            display_text = display_text.replace('[질문]', '\n**💡 [면접 질문]**\n').replace('[평가요소]', '\n**🏷️ [평가 요소]**\n').replace('[평가의도]', '\n**🎯 [평가 의도]**\n').replace('[모범답안]', '\n**✅ [모범 답안 가이드]**\n').replace('[꼬리질문]', '\n**🔥 [압박용 꼬리질문]**\n')
            display_text = display_text.replace('[제시문]', '\n**📄 [서울대 스타일 구술 제시문 (가, 나, 다)]**\n')

            full_display_text = display_text + "\n\n---\n💬 **방금까지 나눈 문항 내용과 피드백 대화 내용을 한글 문서(.docx)로 만들어 드릴까요?** (원하시면 **'그래 만들어줘'**라고 말씀해 주세요!)"

            st.session_state.chat_history = [{"role": "assistant", "content": full_display_text}]

            # 📌 이 학생의 생기부·문항·대화를 저장해서, 다음에 다시 열어도 이어서 볼 수 있게 함
            if not st.session_state.get("current_record_id"):
                st.session_state.current_record_id = str(uuid.uuid4())
            save_record(
                st.session_state.current_record_id, student_name, uni, major, interview_type, difficulty,
                student_record, result_text, st.session_state.chat_history
            )

            st.success("🎉 면접 패키지 및 워드 문서 생성이 완료되었습니다! (이 학생 기록은 자동 저장되었습니다)")

        except Exception as e:
            st.error(f"❌ 생성 실패: {e}")

# 이번 생성에 실제로 참고된 기출 사례를 투명하게 보여줌
if st.session_state.get("last_examples") is not None and not st.session_state.last_examples.empty:
    with st.expander(f"🔎 이번 문항 생성에 참고한 실제 기출 사례 보기 ({len(st.session_state.last_examples)}건)", expanded=False):
        for _, r in st.session_state.last_examples.iterrows():
            st.markdown(f"**[{r['대학']} · {r['학과']}]** {r['질문']}")
            st.caption(str(r['답변'])[:200] + ("..." if len(str(r['답변'])) > 200 else ""))

# -------------------------------------------------------------------------
# [5] 결과 대시보드 및 실시간 피드백
# -------------------------------------------------------------------------
if st.session_state.chat_history:
    st.markdown("<div class='step-text'>STEP 2 · 결과 확인 및 피드백</div>", unsafe_allow_html=True)
    st.markdown("## 📋 면접 문항 대시보드 및 실시간 피드백")

    if st.session_state.word_files:
        stu_path, tea_path = st.session_state.word_files
        chat_path = create_chat_history_word(st.session_state.chat_history, student_name)

        downloadable = [
            ("📥 학생용 워크북 (.docx)", stu_path),
            ("📥 교사용 지침서 (.docx)", tea_path),
            ("💬 피드백 대화 내역 (.docx)", chat_path),
        ]
        if st.session_state.get("easy_explanation_file"):
            downloadable.append(("📚 생기부 쉬운 해설 (.docx)", st.session_state.easy_explanation_file))
        if st.session_state.get("summary_card_file"):
            downloadable.append(("🗂️ 면접직전 요약카드 (.docx)", st.session_state.summary_card_file))
        if st.session_state.get("growth_report_file"):
            downloadable.append(("📈 학생 성장 리포트 (.docx)", st.session_state.growth_report_file))

        download_cols = st.columns(min(len(downloadable), 4))
        for i, (label, path) in enumerate(downloadable):
            with download_cols[i % len(download_cols)]:
                with open(path, "rb") as f:
                    st.download_button(label, f, file_name=path, use_container_width=True, key=f"dl_{i}_{path}")

    if st.session_state.get("easy_explanation_text"):
        with st.expander("📚 생기부 쉬운 해설 미리보기 (고등학교 1학년 눈높이)", expanded=False):
            st.markdown(st.session_state.easy_explanation_text)

    st.divider()

    for message in st.session_state.chat_history:
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
            if st.button("✍️ 이 답변 첨삭 받기", use_container_width=True):
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
                    previous_content = st.session_state.get("last_result_text", "")

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
                        new_result = call_gemini(feedback_prompt, api_key)
                        st.session_state.last_result_text = new_result
                        display_text = new_result.replace('[문제 1]', '\n**💡 [문제 1]**\n').replace('[평가요소 1]', '\n**🏷️ [평가 요소 1]**\n').replace('[평가의도 1]', '\n**🎯 [평가 의도 1]**\n').replace('[모범답안 1]', '\n**✅ [모범 답안 가이드 1]**\n').replace('[꼬리질문 1]', '\n**🔥 [압박용 꼬리질문 1]**\n')
                        display_text = display_text.replace('[문제 2]', '\n**💡 [문제 2]**\n').replace('[평가요소 2]', '\n**🏷️ [평가 요소 2]**\n').replace('[평가의도 2]', '\n**🎯 [평가 의도 2]**\n').replace('[모범답안 2]', '\n**✅ [모범 답안 가이드 2]**\n').replace('[꼬리질문 2]', '\n**🔥 [압박용 꼬리질문 2]**\n')
                        display_text = display_text.replace('[질문]', '\n**💡 [면접 질문]**\n').replace('[평가요소]', '\n**🏷️ [평가 요소]**\n').replace('[평가의도]', '\n**🎯 [평가 의도]**\n').replace('[모범답안]', '\n**✅ [모범 답안 가이드]**\n').replace('[꼬리질문]', '\n**🔥 [압박용 꼬리질문]**\n')
                        display_text = display_text.replace('[제시문]', '\n**📄 [서울대 스타일 구술 제시문 (가, 나, 다)]**\n')

                        full_response = display_text + "\n\n---\n💬 **방금까지 나눈 문항 내용과 피드백 대화 내용을 한글 문서(.docx)로 만들어 드릴까요?** (원하시면 **'그래 만들어줘'**라고 말씀해 주세요!)"

                        st.markdown(full_response)
                        st.session_state.chat_history.append({"role": "assistant", "content": full_response})

                        stu_path, tea_path = create_word_files(new_result, student_name, interview_type, target_desc)
                        st.session_state.word_files = (stu_path, tea_path)

                        if st.session_state.get("current_record_id"):
                            save_record(
                                st.session_state.current_record_id, student_name, uni, major, interview_type, difficulty,
                                st.session_state.get("loaded_student_record_text", ""),
                                new_result, st.session_state.chat_history
                            )

                        st.rerun()

                    except Exception as e:
                        st.error(f"피드백 반영 중 오류가 발생했습니다: {e}")
