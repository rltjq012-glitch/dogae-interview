"""공통 상수 · 작은 도우미(글자 정리, 시각, 개인정보 가리기 등)

app.py에서 나눈 모듈입니다. 아래 층(dogae.-)만 가져다 씁니다."""
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


class _Runtime:
    """화면 흐름에서 읽어 들인 자료(기출 DB 등)를 모듈 함수가 쓰도록 담아 두는 곳."""
    exam_db = None
    univ_info_db = None


RT = _Runtime()

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 저장소 맨 위 폴더

# -------------------------------------------------------------------------
# [0-4] 제시문 면접: '실제 기출 사례'를 먼저 보여주고 응용 문제를 붙입니다
#   AI가 기출을 지어내지 않도록, 기출 원문은 DB에서 코드가 직접 골라 그대로 끼워 넣습니다.
# -------------------------------------------------------------------------
REAL_CASE_TAG = "[실제 기출]"

def to_display_text(result_text):
    """파싱용 [키워드]를 화면에 보기 좋은 굵은 제목으로 바꿉니다."""
    text = str(result_text or "")
    for n in ("1", "2"):
        text = (text.replace(f"[문제 {n}]", f"\n**💡 [문제 {n}]**\n")
                    .replace(f"[평가요소 {n}]", f"\n**🏷️ [평가 요소 {n}]**\n")
                    .replace(f"[평가의도 {n}]", f"\n**🎯 [평가 의도 {n}]**\n")
                    .replace(f"[모범답안 {n}]", f"\n**✅ [모범 답안 가이드 {n}]**\n")
                    .replace(f"[꼬리질문 {n}]", f"\n**🔥 [압박용 꼬리질문 {n}]**\n"))
    text = (text.replace("[질문]", "\n**💡 [면접 질문]**\n")
                .replace("[평가요소]", "\n**🏷️ [평가 요소]**\n")
                .replace("[평가의도]", "\n**🎯 [평가 의도]**\n")
                .replace("[모범답안]", "\n**✅ [모범 답안 가이드]**\n")
                .replace("[꼬리질문]", "\n**🔥 [압박용 꼬리질문]**\n"))
    text = text.replace(REAL_CASE_TAG, "\n**📚 [실제 기출 사례]**\n").replace("[제시문]", "\n**📄 [응용 제시문]**\n")
    return text

def _esc_html(text):
    import html as _html_mod
    return _html_mod.escape(str(text or "")).replace("\n", "<br>")

def _md_safe(text):
    """기출 원문을 화면에 그대로 보여줄 때 마크다운 기호가 서식으로 해석되지 않게 합니다."""
    return str(text or "").replace("~", "∼").replace("$", "＄").replace("*", "＊")

def _safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(name)).strip(" .")
    stem, ext = os.path.splitext(name)
    return (stem[:60] or "문서") + ext[:10]
RECORDS_DB_PATH = os.path.join(APP_DIR, "interview_records.db")

def _is_true(value):
    return value is True or str(value).strip().lower() in ("true", "1", "y", "yes", "예", "o")

def _q_norm(text):
    return re.sub(r"[\s\W_]", "", str(text or ""))


def _kst_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=9)


def _kst_today():
    return _kst_now().strftime("%Y-%m-%d")


def _kst_stamp():
    return _kst_now().strftime("%Y-%m-%d %H:%M")


def _row_get(row, key, default=""):
    try:
        v = row[key]
    except Exception:
        return default
    if v is None or str(v).strip().lower() in ("none", "nan"):
        return default
    return v


def mask_personal(text, name):
    """AI로 보내기 전에 학생 이름을 가립니다(성명 전체, 세 글자 이름이면 이름 두 글자도)."""
    t = str(text or "")
    n = re.sub(r"\s+", "", str(name or ""))
    if len(n) >= 2:
        t = re.sub(r"\s*".join(map(re.escape, n)), "○○", t)
        if len(n) == 3 and re.fullmatch(r"[가-힣]{3}", n):
            t = re.sub(re.escape(n[1:]) + r"(?=[이가은는을를의에와과도만,.\s]|$)", "○○", t)
    return t


def _as_list(x):
    if isinstance(x, list):
        return [str(v).strip() for v in x if str(v).strip()]
    x = str(x or "").strip()
    return [x] if x else []


__all__ = ["RT", "_Runtime", "APP_DIR", "REAL_CASE_TAG", "to_display_text", "_esc_html", "_md_safe", "_safe_filename", "RECORDS_DB_PATH", "_is_true", "_q_norm", "_kst_now", "_kst_today", "_kst_stamp", "_row_get", "mask_personal", "_as_list"]
