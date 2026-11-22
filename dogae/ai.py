"""Gemini 호출 엔진(최신 모델 자동 선택·폴백) · AI 응답(JSON) 읽기

app.py에서 나눈 모듈입니다. 아래 층(dogae.base)만 가져다 씁니다."""
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


# -------------------------------------------------------------------------
# [1-1] 제미나이 모델 자동 선택 (최신 버전 우선 → 실패하면 한 단계씩 낮은 모델로)
#   API가 알려주는 '지금 쓸 수 있는 모델 목록'에서 글 생성용 모델만 골라
#   버전이 높은 것부터 차례로 시도합니다. 구글이 새 모델을 내놓으면 코드를 고치지 않아도
#   자동으로 최신 모델을 쓰고, 옛 모델이 없어져도(404) 다음 모델로 넘어갑니다.
#   순서: 일반 모델(최신 버전 → 낮은 버전, 같은 버전이면 Pro → Flash, 정식 → 미리보기)
#         → 그다음 Lite 모델 → 마지막으로 '-latest' 별칭
# -------------------------------------------------------------------------
GEMINI_MAX_REAL_FAILS = 6            # 한 번 호출에서 '진짜 실패'(한도·과부하 등)를 몇 개 모델까지 견딜지
GEMINI_TIME_BUDGET = 240             # 한 번 호출에서 모델을 바꿔 가며 시도할 최대 시간(초)
GEMINI_FALLBACK_MODELS = [           # 목록 조회가 실패했을 때도 쓸 수 있는 별칭(항상 최신을 가리킴)
    "gemini-flash-latest", "gemini-flash-lite-latest", "gemini-pro-latest",
]
# gemini-3.8-flash, gemini-3.1-pro-preview, gemini-3.5-flash-lite, gemini-2.5-flash-preview-09-2025 …
# (tts·live·image·transcribe·embedding 같은 특수 모델은 이 형식에 맞지 않아 자동으로 빠집니다)
_GEMINI_TEXT_MODEL_RX = re.compile(
    r"^gemini-(\d+(?:\.\d+)?)-(pro|flash-lite|flash)(?:-(preview(?:-\d{2}-\d{4})?|latest|\d{3}))?$"
)
_KEY_ERROR_RX = re.compile(r"API_KEY_INVALID|API key not valid|API key expired", re.I)
_MODEL_GONE_RX = re.compile(r"\b404\b|NOT_FOUND|not found|not supported|no longer available|deprecated|shut ?down|PERMISSION_DENIED|\b403\b", re.I)
# ⚠️ 예전 정규식은 'rate'만 있어도 걸려서, 404 메시지 속 'generateContent'까지 사용량 한도로 오인했습니다.
_RATE_LIMIT_RX = re.compile(r"\b429\b|RESOURCE_EXHAUSTED|quota|rate[ _-]?limit", re.I)
_NO_QUOTA_RX = re.compile(r"limit:\s*0\b", re.I)          # 이 키(무료 등급 등)로는 아예 쓸 수 없는 모델
_DAILY_RX = re.compile(r"PerDay", re.I)                  # 하루 한도 소진
_BUSY_RX = re.compile(r"\b50[0-4]\b|UNAVAILABLE|overloaded|INTERNAL|DEADLINE|timed? ?out", re.I)

def _gemini_rank(name):
    """시도 순서: 세대(3.x → 2.x) → 같은 세대 안에서 일반 모델 먼저, Lite는 뒤
    → 높은 버전 먼저 → Pro → Flash → 정식 → 미리보기."""
    m = _GEMINI_TEXT_MODEL_RX.match(name)
    if not m:
        return ((1, 0), 9, (0, 0), 9, 9)
    ver = [int(x) for x in m.group(1).split(".")] + [0]
    tier, suffix = m.group(2), m.group(3) or ""
    is_lite = 1 if tier == "flash-lite" else 0
    tier_rank = {"pro": 0, "flash": 1, "flash-lite": 2}[tier]
    stab = 0 if suffix in ("", "latest") else 1 if suffix.isdigit() else 2 if suffix == "preview" else 3
    return ((0, -ver[0]), is_lite, (-ver[0], -ver[1]), tier_rank, stab)

@st.cache_data(ttl=3600, show_spinner=False)
def _list_gemini_models(api_key):
    """이 API 키로 지금 쓸 수 있는 글 생성 모델을 '시도할 순서'대로 돌려줍니다(1시간 캐시)."""
    client = genai.Client(api_key=api_key)
    found = set()
    for m in client.models.list():
        name = str(getattr(m, "name", "") or "").replace("models/", "")
        actions = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", None)
        if actions and "generateContent" not in actions:
            continue
        if _GEMINI_TEXT_MODEL_RX.match(name):
            found.add(name)
    return sorted(found, key=_gemini_rank)

@st.cache_resource(show_spinner=False)
def _gemini_model_state():
    """서버 전체에서 공유하는 모델 상태: 건너뛸 모델(skip: 모델→다시 써 볼 시각)과 사유(why)."""
    return {"skip": {}, "why": {}}

def _pinned_models():
    """Secrets에 GEMINI_MODEL = "gemini-3.7-flash" (또는 목록)을 넣으면 그 모델을 맨 먼저 씁니다."""
    try:
        v = st.secrets.get("GEMINI_MODEL", "")
    except Exception:
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [x.strip() for x in str(v or "").split(",") if x.strip()]

def gemini_model_ladder(api_key):
    try:
        listed = list(_list_gemini_models(api_key))
    except Exception:
        listed = []
    ladder = []
    for name in _pinned_models() + listed + GEMINI_FALLBACK_MODELS:
        if name not in ladder:
            ladder.append(name)
    return ladder

def _classify_gemini_error(msg):
    """오류를 (종류, 짧은 설명, 건너뛸 초)로 나눕니다."""
    if _KEY_ERROR_RX.search(msg):
        return "key", "API 키 오류", 0
    if _RATE_LIMIT_RX.search(msg):
        if _NO_QUOTA_RX.search(msg):
            return "noquota", "429 이 키로는 사용 불가(한도 0)", 6 * 3600
        if _DAILY_RX.search(msg):
            return "rate", "429 하루 한도 소진", 3600
        return "rate", "429 분당 한도", 60
    if _MODEL_GONE_RX.search(msg):
        return "gone", "404 사용 종료·접근 불가", 6 * 3600
    if _BUSY_RX.search(msg):
        return "busy", "503 서버 혼잡", 30
    return "other", re.sub(r"\s+", " ", msg)[:90], 0

def _gemini_try_once(client, model, contents, cfg):
    resp = client.models.generate_content(model=model, contents=contents, config=cfg)
    text = getattr(resp, "text", None)
    if not (text and str(text).strip()):
        raise RuntimeError("빈 응답(안전 필터 또는 출력 길이 한도)")
    return text

def _gemini_generate(contents, api_key, json_mode=False):
    """최신 모델부터 차례로 호출하고, 실패하면 다음(낮은) 모델로 내려갑니다.
    - 404(사용 종료) · 한도 0(무료 키로 못 쓰는 모델): 6시간 건너뜀. 시도 횟수에 세지 않음.
    - 429(한도) · 503(혼잡): 잠시 뒤로 미루고 바로 다음 모델.
    - API 키 오류: 어떤 모델도 소용없으므로 바로 알려 줍니다."""
    client = genai.Client(api_key=api_key)
    state = _gemini_model_state()
    now = time.time()
    ladder = gemini_model_ladder(api_key)
    ready = [m for m in ladder if state["skip"].get(m, 0) <= now]
    later = sorted((m for m in ladder if state["skip"].get(m, 0) > now), key=lambda m: state["skip"][m])
    order = ready + later           # 쉬는 중인 모델도 맨 끝에 한 번씩은 다시 시도
    cfg = types.GenerateContentConfig(response_mime_type="application/json") if json_mode else None

    log, real_fails, rate_limited, t0 = [], 0, False, time.time()
    for model in order:
        if real_fails >= GEMINI_MAX_REAL_FAILS or time.time() - t0 > GEMINI_TIME_BUDGET:
            break
        try:
            text = _gemini_try_once(client, model, contents, cfg)
            state["skip"].pop(model, None)
            state["why"].pop(model, None)
            try:
                st.session_state["gemini_last_model"] = model
            except Exception:
                pass
            return text
        except Exception as e:
            msg = str(e)
            kind, short, wait = _classify_gemini_error(msg)
            log.append(f"{model}: {short}")
            if kind == "key":
                raise Exception(f"Gemini API 키가 올바르지 않거나 만료되었습니다. ({msg[:200]})")
            if wait:
                state["skip"][model] = time.time() + wait
                state["why"][model] = short
            if kind == "rate":
                rate_limited = True
            if kind not in ("gone", "noquota"):
                real_fails += 1
                time.sleep(1)
    head = "429 RESOURCE_EXHAUSTED 사용량 한도 — " if rate_limited else ""
    raise Exception(
        f"{head}AI 모델 통신 실패. 모델별 결과: " + " / ".join(log or ["시도할 모델이 없습니다"])
        + " — 사이드바의 '🩺 AI 모델 연결 점검'으로 어떤 모델이 되는지 확인해 보세요."
    )

def diagnose_gemini_models(api_key, limit=14):
    """사이드바 점검용: 모델마다 아주 짧은 요청을 보내 실제로 되는지 확인합니다."""
    client = genai.Client(api_key=api_key)
    state = _gemini_model_state()
    try:
        _list_gemini_models.clear()        # 점검할 때는 모델 목록도 새로 받습니다.
    except Exception:
        pass
    rows = []
    for model in gemini_model_ladder(api_key)[:limit]:
        t = time.time()
        try:
            _gemini_try_once(client, model, "OK라고만 답하세요.", None)
            state["skip"].pop(model, None)
            state["why"].pop(model, None)
            rows.append({"모델": model, "결과": "✅ 사용 가능", "응답 시간": f"{time.time() - t:.1f}초"})
        except Exception as e:
            kind, short, wait = _classify_gemini_error(str(e))
            if wait:
                state["skip"][model] = time.time() + wait
                state["why"][model] = short
            rows.append({"모델": model, "결과": f"❌ {short}", "응답 시간": f"{time.time() - t:.1f}초"})
            if kind == "key":
                break
    return rows

def call_gemini(prompt, api_key, json_mode=False):
    if isinstance(prompt, bytes):
        prompt = prompt.decode('utf-8', errors='ignore')
    elif not isinstance(prompt, str):
        prompt = str(prompt)
    return _gemini_generate(prompt, api_key, json_mode=json_mode)

def call_gemini_audio_eval(audio_bytes, api_key):
    prompt = """
    당신은 도개고등학교의 날카롭고 전문적인 면접관입니다. 다음은 학생이 면접 질문에 대해 직접 스마트폰으로 녹음한 음성 답변입니다.
    아래 3가지 항목을 반드시 포함하여 분석 및 평가 리포트를 작성해 주세요.
    1. 🗣️ **[답변 내용 변환 (STT)]**: 학생의 음성을 텍스트로 정확하게 받아적어 주세요.
    2. 📊 **[면접관의 평가]**: 학생의 답변을 '논리성, 표현력, 전공적합성'을 기준으로 분석하고 종합 평가를 [상 / 중 / 하]로 매겨주세요.
    3. 🔥 **[추가 압박 꼬리질문]**: 학생의 답변 내용 중 논리적 비약이 있거나 더 깊이 파고들 만한 날카로운 꼬리질문을 하나 던져주세요.
    """
    try:
        audio_part = types.Part.from_bytes(data=audio_bytes, mime_type='audio/wav')
        return _gemini_generate([audio_part, prompt], api_key)
    except Exception as e:
        return f"음성 분석 실패: {e}"

def _extract_json_objects(text):
    """글 속의 JSON 객체·배열을 모두 찾아 dict 목록으로 돌려줍니다.
    한 줄에 하나씩 쓰지 않고 여러 줄로 펼치거나 [ ... ] 배열·```json 블록·{"items": [...]}로 감싸도 읽습니다."""
    dec = json.JSONDecoder()
    out, i, text = [], 0, str(text or "")
    while i < len(text):
        if text[i] not in "{[":
            i += 1
            continue
        try:
            obj, end = dec.raw_decode(text, i)
        except Exception:
            i += 1
            continue
        objs = obj if isinstance(obj, list) else [obj]
        for o in objs:
            if not isinstance(o, dict):
                continue
            if isinstance(o.get("items"), list):
                out.extend(x for x in o["items"] if isinstance(x, dict))
                if o.get("키워드"):
                    out.append({"키워드": o["키워드"]})
            else:
                out.append(o)
        i = end
    return out

def _call_gemini_retry(prompt, api_key, json_mode=False):
    """무료 등급은 분당 사용량 한도가 낮아, 막히면 잠시 쉬고 한 번 더 부릅니다."""
    try:
        return call_gemini(prompt, api_key, json_mode=json_mode)
    except Exception as e:
        if not _RATE_LIMIT_RX.search(str(e)):
            raise
        time.sleep(25)
        return call_gemini(prompt, api_key, json_mode=json_mode)

def _friendly_error(e):
    msg = str(e)
    if _RATE_LIMIT_RX.search(msg):
        msg = "Gemini 사용량 한도에 걸렸습니다(무료 등급은 분당 한도가 낮습니다). 1~2분 뒤 다시 시도해 주세요. — " + msg[:200]
    return msg


# ---------------------------------------------------------------- AI 첨삭
def call_gemini_parts(contents, api_key, json_mode=True):
    """음성처럼 글이 아닌 자료를 함께 보낼 때 쓰는 호출."""
    return _gemini_generate(contents, api_key, json_mode=json_mode)


__all__ = ["GEMINI_MAX_REAL_FAILS", "GEMINI_TIME_BUDGET", "GEMINI_FALLBACK_MODELS", "_GEMINI_TEXT_MODEL_RX", "_KEY_ERROR_RX", "_MODEL_GONE_RX", "_RATE_LIMIT_RX", "_NO_QUOTA_RX", "_DAILY_RX", "_BUSY_RX", "_gemini_rank", "_list_gemini_models", "_gemini_model_state", "_pinned_models", "gemini_model_ladder", "_classify_gemini_error", "_gemini_try_once", "_gemini_generate", "diagnose_gemini_models", "call_gemini", "call_gemini_audio_eval", "_extract_json_objects", "_call_gemini_retry", "_friendly_error", "call_gemini_parts"]
