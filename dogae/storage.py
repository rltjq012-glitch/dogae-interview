"""기록 저장소(구글시트 Apps Script · 서비스 계정 · 로컬 SQLite) · 로그인 설정 · 학생 링크·답안 저장 줄

app.py에서 나눈 모듈입니다. 아래 층(dogae.base, dogae.ai)만 가져다 씁니다."""
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



# -------------------------------------------------------------------------
# [0-1] 학생별 저장/불러오기 저장소
#   - 생기부 텍스트, 생성된 문항, 대화 내역을 저장해두고 다음에 다시 열었을 때
#     PDF를 재업로드하지 않고 이어서 볼 수 있게 합니다.
#   - Streamlit Secrets에 구글 서비스 계정("gcp_service_account")이 등록되어 있으면
#     Google Sheets에 반영구적으로 저장합니다(앱이 재배포돼도 사라지지 않음).
#   - 등록되어 있지 않으면 예전처럼 로컬 SQLite 파일에 저장합니다(재배포 시 초기화될 수 있음).
# -------------------------------------------------------------------------

# -------------------------------------------------------------------------
# [0-A] 교사 로그인 / 기록 소유권 / 출력 파일 경로  (교내 시범운영용)
#   - Streamlit Secrets에 [teachers] 표가 있으면 로그인이 켜집니다.
#         [teachers]
#         "김기섭" = "비밀번호"
#         "홍길동" = "비밀번호"
#   - [teachers]가 없으면 로그인 없이 예전과 똑같이 동작합니다.
#   - 관리자는 ADMIN_TEACHERS = "김기섭" 으로 지정합니다(쉼표로 여러 명).
#     지정하지 않으면 [teachers]의 첫 번째 계정이 관리자입니다.
#   - 기록 id를 "교사아이디::uuid" 형태로 만들어, 구글시트 열 구조를 바꾸지 않고도
#     교사별로 기록을 나눕니다. 예전 기록(id에 "::"가 없음)은 관리자에게만 보입니다.
# -------------------------------------------------------------------------
OWNER_SEP = "::"
_ADMIN_KEY = "ADMIN_TEACHERS"

def _split_names(value):
    if isinstance(value, (list, tuple)):
        return [str(v).strip() for v in value if str(v).strip()]
    return [v.strip() for v in str(value or "").split(",") if v.strip()]

def _auth_config():
    """(계정 dict, 관리자 아이디 목록)을 돌려줍니다. 설정이 없으면 ({}, [])."""
    try:
        if "teachers" not in st.secrets:
            return {}, []
        table = dict(st.secrets["teachers"])
        admins = []
        # TOML에서는 [teachers] 아래에 적은 키가 전부 그 표에 들어가므로,
        # ADMIN_TEACHERS를 표 안에 적었더라도 계정이 아니라 설정으로 취급합니다.
        if _ADMIN_KEY in table:
            admins = _split_names(table.pop(_ADMIN_KEY))
        if _ADMIN_KEY in st.secrets:
            admins = _split_names(st.secrets[_ADMIN_KEY]) or admins
        accounts = {
            str(k).strip(): str(v) for k, v in table.items()
            if str(k).strip() and OWNER_SEP not in str(k) and str(v)
        }
        if not admins and accounts:
            admins = [next(iter(accounts))]
        return accounts, admins
    except Exception:
        return {}, []

def auth_enabled():
    return bool(_auth_config()[0])

def current_teacher():
    """로그인한 교사 아이디. 로그인이 꺼져 있으면 None."""
    accounts, _ = _auth_config()
    teacher = st.session_state.get("auth_teacher")
    return teacher if (accounts and teacher in accounts) else None

def is_admin():
    teacher = current_teacher()
    return bool(teacher) and teacher in _auth_config()[1]

def new_record_id():
    """새 기록 id. 로그인 중이면 소유 교사를 id 앞에 붙입니다."""
    rid = str(uuid.uuid4())
    teacher = current_teacher()
    return f"{teacher}{OWNER_SEP}{rid}" if teacher else rid

def record_owner(record_id):
    rid = str(record_id or "")
    return rid.rsplit(OWNER_SEP, 1)[0] if OWNER_SEP in rid else ""

def can_access_record(record_id):
    if not auth_enabled():
        return True
    teacher = current_teacher()
    if not teacher:
        return False
    return is_admin() or record_owner(record_id) == teacher

# ---- 출력 파일 경로: 접속(세션)마다 다른 임시 폴더를 써서,
#      여러 교사가 동시에 같은 이름의 학생을 작업해도 서로의 파일을 덮어쓰지 않게 합니다.
_OUT_DIR_PREFIX = "interview_out_"

def _prune_old_out_dirs(max_age_hours=24):
    """하루 넘게 지난 임시 폴더를 지워, 학생 이름이 들어간 파일이 서버에 쌓이지 않게 합니다."""
    base = tempfile.gettempdir()
    cutoff = time.time() - max_age_hours * 3600
    try:
        for entry in os.listdir(base):
            path = os.path.join(base, entry)
            if entry.startswith(_OUT_DIR_PREFIX) and os.path.isdir(path) and os.path.getmtime(path) < cutoff:
                shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass

def _out_path(filename):
    out_dir = st.session_state.get("_out_dir")
    if not out_dir or not os.path.isdir(out_dir):
        _prune_old_out_dirs()
        out_dir = tempfile.mkdtemp(prefix=_OUT_DIR_PREFIX)
        st.session_state["_out_dir"] = out_dir
    return os.path.join(out_dir, _safe_filename(filename))

# -------------------------------------------------------------------------
# [0-0] JSON 안전 파서 / 직렬화 헬퍼
#   구글시트(Apps Script) 셀에는 공백, "None", "nan" 같은 값이나 잘린 문자열이
#   들어가 있을 수 있어서, json.loads()를 그대로 쓰면 앱 전체가 멈춥니다.
#   아래 헬퍼는 어떤 값이 와도 예외 없이 기본값으로 돌려줍니다.
# -------------------------------------------------------------------------
SHEET_CELL_LIMIT = 45000  # 구글시트 셀 상한(5만 자)보다 여유 있게

def load_saved_history(value):
    """저장된 대화 내역 칸을 (대화 목록, 부가자료 dict)로 읽습니다. 예전 형식(목록만)도 그대로 읽힙니다."""
    data = load_json_safe(value, [])
    if isinstance(data, dict):
        hist = data.get("history") if isinstance(data.get("history"), list) else []
        extras = data.get("extras") if isinstance(data.get("extras"), dict) else {}
        return hist, extras
    return (data if isinstance(data, list) else []), {}

def _current_card_checks():
    """지금 화면의 질문에 대한 자가 체크만 저장합니다(지난 질문의 체크는 버림)."""
    hl = st.session_state.get("highlight_data") or {}
    keys = set(hl.get("q_keys") or [])
    return {k: v for k, v in (st.session_state.get("card_checks") or {}).items()
            if k in keys and v and v != "미확인"}

def history_for_save(chat_history=None):
    """저장할 때 대화 내역과 함께 형광펜 목록·쉬운 해설을 묶습니다."""
    if chat_history is None:
        chat_history = st.session_state.get("chat_history", [])
    return {
        "history": list(chat_history or []),
        "extras": {
            "highlight_items": st.session_state.get("highlight_raw_items") or [],
            "easy_explanation": st.session_state.get("easy_explanation_text") or "",
            "card_checks": _current_card_checks(),
        },
    }

def load_json_safe(value, default=None):
    """어떤 값이 와도 예외를 던지지 않는 json.loads."""
    if default is None:
        default = []
    if value is None:
        return default
    if isinstance(value, (list, dict)):   # 이미 파싱된 경우
        return value
    if isinstance(value, float) and math.isnan(value):  # pandas NaN
        return default
    text = str(value).strip()
    if not text or text.lower() in ("none", "nan", "null", "-", "#error!", "#n/a"):
        return default
    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError, ValueError):
        pass
    try:  # str(list)처럼 작은따옴표로 저장된 경우 구제
        parsed = ast.literal_eval(text)
        if isinstance(parsed, (list, dict)):
            return parsed
    except Exception:
        pass
    return default

def _fit_history(history, limit):
    history = list(history or [])
    text = json.dumps(history, ensure_ascii=False)
    while len(text) > limit and len(history) > 1:
        history = history[1:]
        text = json.dumps(history, ensure_ascii=False)
    return history if len(text) <= limit else []

def dump_chat_history(chat_history, limit=SHEET_CELL_LIMIT):
    """대화 내역을 JSON 문자열로 만들되, 구글시트 셀 상한을 넘지 않게
    오래된 메시지부터 잘라냅니다. (중간에서 잘린 JSON은 다시 못 읽기 때문)
    chat_history가 {"history": [...], "extras": {...}} 이면 형광펜·쉬운 해설도 같은 칸에 함께 저장합니다.
    칸이 모자라면 쉬운 해설 → 오래된 대화 → 형광펜 순서로 덜어냅니다."""
    if isinstance(chat_history, dict):
        history = list(chat_history.get("history") or [])
        extras = {k: v for k, v in (chat_history.get("extras") or {}).items() if v}
        if extras:
            def pack(h, ex):
                return json.dumps({"v": 2, "history": h, "extras": ex}, ensure_ascii=False)
            text = pack(history, extras)
            if len(text) <= limit:
                return text
            extras.pop("easy_explanation", None)
            if extras:
                room = limit - len(pack([], extras))
                if room > 2:
                    return pack(_fit_history(history, room), extras)
        chat_history = history
    history = list(chat_history or [])
    text = json.dumps(history, ensure_ascii=False)
    while len(text) > limit and len(history) > 1:
        history = history[1:]
        text = json.dumps(history, ensure_ascii=False)
    if len(text) > limit:   # 메시지 한 개가 통째로 너무 길 때
        history = []
        text = "[]"
    return text

RECORD_COLUMNS = [
    "id", "student_name", "university", "major", "interview_type",
    "difficulty", "student_record_text", "result_text", "chat_history", "updated_at"
]
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
        student_record_text, result_text, dump_chat_history(chat_history),
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
        student_record_text, result_text, dump_chat_history(chat_history),
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
        "chat_history": dump_chat_history(chat_history),
        "updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    resp = requests.post(st.secrets["APPS_SCRIPT_URL"], json=payload, timeout=30)
    resp.raise_for_status()
    # ⚠️ Apps Script는 인증 실패 같은 오류도 HTTP 200으로 돌려주기 때문에,
    #    응답 본문까지 확인해야 "저장된 줄 알았는데 실제로는 저장 안 된" 상황을 막을 수 있습니다.
    try:
        body = resp.json()
    except Exception:
        raise RuntimeError("Apps Script가 예상과 다른 응답을 보냈습니다. 웹 앱 URL과 재배포 상태를 확인해 주세요.")
    if isinstance(body, dict) and body.get("error"):
        raise RuntimeError(body["error"])
    if not (isinstance(body, dict) and body.get("status") == "ok"):
        raise RuntimeError(f"Apps Script 저장 응답이 올바르지 않습니다: {str(body)[:200]}")
    return body

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
    try:
        body = resp.json()
    except Exception:
        return
    if isinstance(body, dict) and body.get("error"):
        raise RuntimeError(body["error"])

# ---- 공용 인터페이스: Secrets 설정 여부에 따라 자동으로 저장소를 선택 ----
#   우선순위: Apps Script(가장 간단, GCP 콘솔 불필요) > Google Sheets 서비스 계정 > 로컬 SQLite
def save_record(*args, quiet=False, **kwargs):
    _clear_records_cache()
    try:
        return _save_record_impl(*args, quiet=quiet, **kwargs)
    finally:
        _clear_records_cache()

def _save_record_impl(*args, quiet=False, **kwargs):
    """기록을 저장하고 '어디에 저장되었는지'를 딕셔너리로 돌려줍니다.
    반환값: {"ok": bool, "where": "구글시트(Apps Script)"/"구글시트(서비스 계정)"/"로컬 임시저장",
             "at": "HH:MM:SS", "error": "(실패 시 사유)"}
    """
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if _appsscript_enabled():
        try:
            _appsscript_save_record(*args, **kwargs)
            info = {"ok": True, "where": "구글시트(Apps Script)", "at": now_str, "error": ""}
            st.session_state["last_save_info"] = info
            return info
        except Exception as e:
            if not quiet:
                st.warning(f"⚠️ 구글시트 저장에 실패해 로컬에만 임시 저장합니다: {e}")
            _sqlite_save_record(*args, **kwargs)
            info = {"ok": False, "where": "로컬 임시저장", "at": now_str, "error": str(e)}
            st.session_state["last_save_info"] = info
            return info

    if _sheets_enabled():
        try:
            _sheets_save_record(*args, **kwargs)
            info = {"ok": True, "where": "구글시트(서비스 계정)", "at": now_str, "error": ""}
            st.session_state["last_save_info"] = info
            return info
        except Exception as e:
            if not quiet:
                st.warning(f"⚠️ 구글시트 저장에 실패해 로컬에만 임시 저장합니다: {e}")
            _sqlite_save_record(*args, **kwargs)
            info = {"ok": False, "where": "로컬 임시저장", "at": now_str, "error": str(e)}
            st.session_state["last_save_info"] = info
            return info

    _sqlite_save_record(*args, **kwargs)
    info = {"ok": True, "where": "로컬 임시저장", "at": now_str, "error": ""}
    st.session_state["last_save_info"] = info
    return info

# 속도 개선: 화면을 클릭할 때마다(=Streamlit이 스크립트를 다시 실행할 때마다) 구글시트 전체를
# 다시 내려받지 않도록, 기록 목록을 2분간 캐시합니다. 저장·삭제 시에는 캐시를 즉시 비웁니다.
@st.cache_data(ttl=120, show_spinner=False)
def _fetch_records_cached(backend):
    if backend == "appsscript":
        return _appsscript_list_records()
    if backend == "sheets":
        return _sheets_list_records()
    return [dict(r) for r in _sqlite_list_records()]

def _clear_records_cache():
    try:
        _fetch_records_cached.clear()
    except Exception:
        pass

def _list_records_raw():
    if _appsscript_enabled():
        try:
            return _fetch_records_cached("appsscript")
        except Exception as e:
            st.warning(f"⚠️ Apps Script 조회에 실패했습니다: {e}")
            return []
    if _sheets_enabled():
        try:
            return _fetch_records_cached("sheets")
        except Exception as e:
            st.warning(f"⚠️ Google Sheets 조회에 실패했습니다: {e}")
            return []
    return _fetch_records_cached("sqlite")

def _get_record_raw(record_id):
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

def _delete_record_raw(record_id):
    try:
        return _delete_record_impl(record_id)
    finally:
        _clear_records_cache()

def _delete_record_impl(record_id):
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

def list_records():
    """로그인한 교사가 볼 수 있는 기록만 돌려줍니다(관리자는 전체)."""
    return [r for r in _list_records_raw() if not is_practice_id(r["id"]) and can_access_record(r["id"])]

def get_record(record_id):
    if not can_access_record(record_id):
        return None
    return _get_record_raw(record_id)

def delete_record(record_id):
    if not can_access_record(record_id):
        st.warning("⚠️ 이 기록을 삭제할 권한이 없습니다.")
        return
    return _delete_record_raw(record_id)

def save_current_session(student_name, university, major, interview_type, difficulty, quiet=False):
    """지금 화면에 올라와 있는 내용(생기부 텍스트·문항·대화 내역)을 그대로 저장합니다.
    자동 저장과 수동 저장('💾 지금 저장하기') 모두 이 함수를 사용합니다."""
    if not st.session_state.get("current_record_id"):
        st.session_state.current_record_id = new_record_id()
    return save_record(
        st.session_state.current_record_id, student_name, university, major, interview_type, difficulty,
        st.session_state.get("loaded_student_record_text", ""),
        st.session_state.get("last_result_text", ""),
        history_for_save(),
        quiet=quiet,
    )

def save_interview_card(record_id, student_name, uni, major, card):
    if not record_id or not card:
        return
    try:
        _save_json_row(f"{CARD_PREFIX}{_bundle_key(record_id)}", student_name, uni, major, "면접카드", card)
    except Exception as e:
        st.session_state.interview_card_error = f"요약카드는 만들었지만 저장하지 못했습니다(다시 불러오면 사라질 수 있음): {e}"

def load_interview_card(record_id):
    try:
        _, obj = _load_json_row(f"{CARD_PREFIX}{_bundle_key(record_id)}")
    except Exception:
        return None
    return obj if (obj.get("concepts") or obj.get("qas")) else None

def save_question_sheet(record_id, student_name, uni, major, sheet):
    if not record_id or not sheet:
        return
    try:
        _save_json_row(f"{QSHEET_PREFIX}{_bundle_key(record_id)}", student_name, uni, major, "면접질문지", sheet)
    except Exception as e:
        st.session_state.qsheet_error = f"질문지는 만들었지만 저장하지 못했습니다(다시 불러오면 사라질 수 있음): {e}"

def load_question_sheet(record_id):
    try:
        _, obj = _load_json_row(f"{QSHEET_PREFIX}{_bundle_key(record_id)}")
    except Exception:
        return None
    return obj if obj.get("qs") else None


def load_qbank_state(record_id):
    try:
        _, obj = _load_json_row(f"{QBANK_PREFIX}{_bundle_key(record_id)}")
    except Exception:
        obj = {}
    return obj or {}


def save_qbank_state(record_id, student_name, uni, major):
    if not record_id:
        return False
    obj = {"sel": st.session_state.get("qbank_sel") or [], "runs": (st.session_state.get("qbank_runs") or [])[-8:]}
    try:
        _save_json_row(f"{QBANK_PREFIX}{_bundle_key(record_id)}", student_name, uni, major, "질문은행", obj)
        return True
    except Exception as e:
        st.warning(f"질문 은행 선택을 저장하지 못했습니다: {e}")
        return False

STU_PREFIX, ANS_PREFIX, CMT_PREFIX = "STU-", "ANS-", "CMT-"
CARD_PREFIX = "CARD-"            # 면접 전날 요약카드
QSHEET_PREFIX = "QSHEET-"        # 대학 맞춤 면접 질문지
QBANK_PREFIX = "QBANK-"          # 질문 은행 선택 · 모의면접 진행 기록
IVW_PREFIX = "IVW-"              # 면접관(패드·폰) 진행 링크
IVR_PREFIX = "IVR-"              # 패드·폰에서 진행한 모의면접 결과
FB_CRITERIA = ["근거", "구조", "구체성", "전공연결"]
FB_CRITERIA_HELP = {
    "근거": "생기부 속 활동을 근거로 들었나",
    "구조": "결론부터 말하고 근거 → 배운 점 순서인가",
    "구체성": "과정·맡은 역할·수치가 구체적인가",
    "전공연결": "지원 학과·진로와 이어지는가",
}
_ROW_CHUNK = 38000             # 저장 칸 하나에 넣을 최대 글자 수
_ROW_TOTAL = 76000             # 한 줄(result_text + chat_history)에 넣을 최대 글자 수


def is_practice_id(rid):
    return str(rid or "").startswith((STU_PREFIX, ANS_PREFIX, CMT_PREFIX, CARD_PREFIX, QSHEET_PREFIX, QBANK_PREFIX,
                                    IVW_PREFIX, IVR_PREFIX))


def _bundle_key(record_id):
    return _hashlib.sha256(str(record_id).encode("utf-8")).hexdigest()[:16]


def storage_is_shared():
    """학생 기기와 선생님 기기가 같은 기록을 보려면 구글시트 저장이 필요합니다."""
    return _appsscript_enabled() or _sheets_enabled()


def _save_json_row(row_id, name, uni, major, kind, obj):
    """JSON을 result_text(1칸)와 chat_history(넘치는 부분)에 나눠 저장합니다.
    각 칸 앞에 'J'를 붙여, 구글시트가 '='·'+'로 시작하는 조각을 수식으로 읽지 않게 합니다."""
    s = json.dumps(obj, ensure_ascii=False)
    if len(s) > _ROW_TOTAL:
        raise RuntimeError(f"저장할 내용이 너무 많습니다({len(s):,}자).")
    a, b = s[:_ROW_CHUNK], s[_ROW_CHUNK:]
    return save_record(row_id, name, uni, major, kind, "", "", "J" + a,
                       ([{"c": "J" + b}] if b else []), quiet=True)


def _load_json_row(row_id):
    """(행, JSON dict)을 돌려줍니다. 없으면 (None, {})."""
    try:
        row = _get_record_raw(row_id)
    except Exception:
        row = None
    if not row:
        return None, {}
    a = str(_row_get(row, "result_text", ""))
    hist, _ = load_saved_history(_row_get(row, "chat_history", ""))
    b = str(hist[0].get("c", "")) if hist and isinstance(hist[0], dict) else ""
    a, b = (a[1:] if a.startswith("J") else a), (b[1:] if b.startswith("J") else b)
    try:
        obj = json.loads(a + b) if (a + b).strip() else {}
    except Exception:
        obj = {}
    return row, (obj if isinstance(obj, dict) else {})


# ---------------------------------------------------------------- 저장 행 다루기
def practice_link_rows(record_id=None):
    """학생 링크 줄 목록(id·이름·대학·학과). record_id를 주면 그 기록의 링크만."""
    try:
        rows = _list_records_raw()
    except Exception:
        rows = []
    out = [r for r in rows if str(_row_get(r, "id", "")).startswith(STU_PREFIX)]
    if record_id:
        pre = f"{STU_PREFIX}{_bundle_key(record_id)}."
        out = [r for r in out if str(_row_get(r, "id", "")).startswith(pre)]
    return out


def save_student_link(token, record_id, name, uni, major, settings, snapshot):
    obj = dict(settings)
    obj.update(snapshot)
    obj["record_id"] = record_id
    obj["synced"] = _kst_stamp()
    # 너무 크면 대조용 생기부 → 문단 순서로 줄입니다.
    for cut in (15000, 9000, 5000, 2500):
        obj["record"] = str(snapshot.get("record", ""))[:cut]
        if len(json.dumps(obj, ensure_ascii=False)) <= _ROW_TOTAL:
            break
    while len(json.dumps(obj, ensure_ascii=False)) > _ROW_TOTAL and any(len(q["para"]) > 150 for q in obj["qs"]):
        for q in obj["qs"]:
            q["para"] = q["para"][:150]
    return _save_json_row(f"{STU_PREFIX}{_bundle_key(record_id)}.{token}", name, uni, major, "학생링크", obj)


def load_answers(record_id):
    _, obj = _load_json_row(f"{ANS_PREFIX}{_bundle_key(record_id)}")
    obj.setdefault("q", {})
    obj.setdefault("usage", {})
    return obj


def _prune_answers(ans):
    """저장 칸이 넘치지 않게, 오래된 첨삭의 긴 글부터 줄입니다(점수는 남김)."""
    def size():
        return len(json.dumps(ans, ensure_ascii=False))
    steps = [
        lambda t, last: [t["fb"].pop(k, None) for k in ("better", "good", "transcript")] if not last else None,
        lambda t, last: [f.get("fb", {}).pop(k, None) for f in t.get("follow", []) for k in ("better", "good", "fixes")] if not last else None,
        lambda t, last: t["fb"].pop("fixes", None) if not last else None,
        lambda t, last: t.__setitem__("text", t.get("text", "")[:200]) if not last else None,
    ]
    for step in steps:
        if size() <= _ROW_TOTAL - 2000:
            return ans
        for item in ans.get("q", {}).values():
            tries = item.get("tries", [])
            for i, t in enumerate(tries):
                t.setdefault("fb", {})
                step(t, i == len(tries) - 1)
    for item in ans.get("q", {}).values():          # 그래도 넘치면 오래된 시도를 덜어냄
        if size() <= _ROW_TOTAL - 2000:
            break
        item["tries"] = item.get("tries", [])[-3:]
    days = sorted(ans.get("usage", {}))
    ans["usage"] = {d: ans["usage"][d] for d in days[-7:]}
    return ans


def save_answers(record_id, name, uni, major, ans, owner=""):
    ans["owner"] = owner or ans.get("owner", "")
    ans["record_id"] = record_id
    ans["updated"] = _kst_stamp()
    return _save_json_row(f"{ANS_PREFIX}{_bundle_key(record_id)}", name, uni, major, "학생답안", _prune_answers(ans))


def load_comments(record_id):
    _, obj = _load_json_row(f"{CMT_PREFIX}{_bundle_key(record_id)}")
    obj.setdefault("c", {})
    return obj


def save_comments(record_id, name, uni, major, obj):
    return _save_json_row(f"{CMT_PREFIX}{_bundle_key(record_id)}", name, uni, major, "교사코멘트", obj)


def _student_record_tab(qs, ans):
    rows, firsts, lasts = [], [], []
    for q in qs:
        tries = ans["q"].get(q["key"], {}).get("tries", [])
        if not tries:
            continue
        f, l = tries[0].get("fb", {}), tries[-1].get("fb", {})
        if f.get("scores"):
            firsts.append(f["scores"])
        if l.get("scores"):
            lasts.append(l["scores"])
        rows.append({"질문": f"{q['no']}. {q['q'][:30]}", "시도": len(tries),
                     "처음": f.get("total"), "최근": l.get("total"),
                     "꼬리질문": len(tries[-1].get("follow", [])), "생기부와 다름": "⚠️" if l.get("mismatch") else ""})
    if not rows:
        st.info("아직 첨삭받은 답변이 없어요. '질문 연습'에서 첫 답변을 써 보세요.")
        return
    avg = lambda lst, k: round(sum(s[k] for s in lst) / len(lst), 1) if lst else 0
    m1, m2, m3 = st.columns(3)
    m1.metric("답한 질문", f"{len(rows)}/{len(qs)}")
    m2.metric("평균 점수(최근)", f"{round(sum(r['최근'] or 0 for r in rows) / len(rows), 1)}/20")
    m3.metric("첨삭 횟수", sum(r["시도"] for r in rows))
    chart = pd.DataFrame({"처음": [avg(firsts, k) for k in FB_CRITERIA], "최근": [avg(lasts, k) for k in FB_CRITERIA]},
                         index=FB_CRITERIA)
    st.markdown("**기준별 평균 (5점 만점)**")
    st.bar_chart(chart, stack=False, height=240, horizontal=True)
    weak = min(FB_CRITERIA, key=lambda k: avg(lasts, k))
    st.caption(f"가장 보완할 점: **{weak}** — {FB_CRITERIA_HELP[weak]}")
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")


# ---------------------------------------------------------------- 선생님: 반 전체 통계·성장 리포트
def practice_answer_rows():
    """로그인한 교사가 볼 수 있는 학생 답안 줄 목록."""
    try:
        rows = _list_records_raw()
    except Exception:
        rows = []
    return [r for r in rows if str(_row_get(r, "id", "")).startswith(ANS_PREFIX)]



# ---------------------------------------------------------------- 📱 패드·폰 면접관 링크 · 결과
def interviewer_link_rows(record_id):
    """이 기록의 면접관 링크 줄(보통 하나)."""
    try:
        rows = _list_records_raw()
    except Exception:
        rows = []
    pre = f"{IVW_PREFIX}{_bundle_key(record_id)}."
    return [r for r in rows if str(_row_get(r, "id", "")).startswith(pre)]


def save_interviewer_link(token, record_id, name, uni, major, obj):
    obj = dict(obj, record_id=record_id, synced=_kst_stamp())
    return _save_json_row(f"{IVW_PREFIX}{_bundle_key(record_id)}.{token}", name, uni, major, "면접관링크", obj)


def load_device_runs(record_id):
    _, obj = _load_json_row(f"{IVR_PREFIX}{_bundle_key(record_id)}")
    return list(obj.get("runs") or [])


def save_device_run(record_id, name, uni, major, rec):
    """패드·폰 진행 결과 한 건을 넣거나 바꿉니다(같은 id면 바꿈, 최근 10건만 보관)."""
    runs = [r for r in load_device_runs(record_id) if r.get("id") != rec.get("id")] + [rec]
    return _save_json_row(f"{IVR_PREFIX}{_bundle_key(record_id)}", name, uni, major, "면접관결과",
                          {"record_id": record_id, "runs": runs[-10:]})


__all__ = ["IVW_PREFIX", "IVR_PREFIX", "interviewer_link_rows", "save_interviewer_link", "load_device_runs", "save_device_run", "OWNER_SEP", "_ADMIN_KEY", "_split_names", "_auth_config", "auth_enabled", "current_teacher", "is_admin", "new_record_id", "record_owner", "can_access_record", "_OUT_DIR_PREFIX", "_prune_old_out_dirs", "_out_path", "SHEET_CELL_LIMIT", "load_saved_history", "_current_card_checks", "history_for_save", "load_json_safe", "_fit_history", "dump_chat_history", "RECORD_COLUMNS", "GOOGLE_SHEET_SCOPES", "_sheets_enabled", "_get_records_worksheet", "init_records_db", "_sqlite_save_record", "_sqlite_list_records", "_sqlite_get_record", "_sqlite_delete_record", "_sheets_all_values", "_sheets_save_record", "_sheets_list_records", "_sheets_get_record", "_sheets_delete_record", "_appsscript_enabled", "_appsscript_save_record", "_appsscript_list_records", "_appsscript_get_record", "_appsscript_delete_record", "save_record", "_save_record_impl", "_fetch_records_cached", "_clear_records_cache", "_list_records_raw", "_get_record_raw", "_delete_record_raw", "_delete_record_impl", "list_records", "get_record", "delete_record", "save_current_session", "save_interview_card", "load_interview_card", "save_question_sheet", "load_question_sheet", "load_qbank_state", "save_qbank_state", "STU_PREFIX", "ANS_PREFIX", "CMT_PREFIX", "CARD_PREFIX", "QSHEET_PREFIX", "QBANK_PREFIX", "FB_CRITERIA", "FB_CRITERIA_HELP", "_ROW_CHUNK", "_ROW_TOTAL", "is_practice_id", "_bundle_key", "storage_is_shared", "_save_json_row", "_load_json_row", "practice_link_rows", "save_student_link", "load_answers", "_prune_answers", "save_answers", "load_comments", "save_comments", "_student_record_tab", "practice_answer_rows"]
