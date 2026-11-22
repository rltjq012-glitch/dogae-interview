"""질문 만들기·읽기(5세트 · 형광펜 카드 · 요약카드 · 질문지 · 질문 은행 · 답안 첨삭)

app.py에서 나눈 모듈입니다. 아래 층(dogae.base, dogae.ai, dogae.storage, dogae.exam_data)만 가져다 씁니다."""
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

QTYPE_ORDER = [n for n, _ in QTYPE_RULES] + [QTYPE_ETC]

# -------------------------------------------------------------------------
# [0-5] 출제 분석 화면 (대학·학과를 고르면 문항 생성 전에 바로 볼 수 있음)
# -------------------------------------------------------------------------
ALL_UNI_REGION = "전체 대학 (기출 많은 순)"
DIRECT_REGION = "직접 입력"

# 권역 구분. 교육대·교원대와 사관학교는 소재지와 상관없이 따로 묶습니다.
UNI_REGION_ORDER = [
    "서울권", "경기·인천권", "충청권", "경상권 (대구·경북)", "경상권 (부산·울산·경남)",
    "강원·호남권", "교육대·교원대", "사관학교",
]

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
    doc = pymupdf.open(stream=uploaded_file.getvalue(), filetype="pdf")
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

# -------------------------------------------------------------------------
# [2-0] 🖍️ 형광펜 생기부
#   생기부 원문에 3색 형광펜을 칠해, 면접관이 짚을 만한 곳을 미리 준비하게 합니다.
#     노랑(확정) = 이미 만든 문항에서 다룬 부분
#     주황(예상) = 아직 안 물었지만 나올 가능성이 높은 부분
#     하늘(방어) = 공격받을 수 있는 부분 (과장된 표현, 연결이 약한 활동 등)
#   AI에게는 '칠할 구절을 원문 그대로 인용'만 시키고, 실제 색칠은 코드가 원문에서 그 구절을
#   찾아서 합니다. 원문에서 찾지 못한 구절은 버리므로, AI가 생기부 문장을 바꿔 쓸 수 없습니다.
#   비용을 늘리지 않으려고 '생기부 쉬운 해설' 호출에 함께 요청합니다.
# -------------------------------------------------------------------------
HL_MARKER = "===형광펜==="
HL_COMPS = {
    "학업": {"label": "학업역량", "fill": "FFF59D", "css": "#fff59d", "icon": "🟨"},
    "진로": {"label": "진로역량", "fill": "BBDEFB", "css": "#bbdefb", "icon": "🟦"},
    "공동체": {"label": "공동체역량", "fill": "C8E6C9", "css": "#c8e6c9", "icon": "🟩"},
}
HL_DEFEND = {"label": "압박 대비", "icon": "〰️", "css": "#d32f2f", "hex": "D32F2F"}
HL_Q_TYPES = ["확인", "심화", "연결", "압박"]
HL_MAX_CARDS = 20            # 칠하는 문장 수 상한 (번호 ①~⑳)
_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
_CHECK_OPTS = ["미확인", "자신 있음", "애매", "모름"]

# 생기부 영역 (빈틈 표시용). 원문에 이 영역이 있는지 + 형광펜이 몇 개 닿았는지를 셉니다.
SANGBU_AREAS = [
    ("교과 세특", r"세\s*부\s*능\s*력\s*및\s*특\s*기\s*사\s*항|세특"),
    ("자율활동", r"자\s*율\s*활\s*동|자율"),
    ("동아리활동", r"동\s*아\s*리\s*활\s*동|동아리"),
    ("진로활동", r"진\s*로\s*활\s*동|진로"),
    ("행동특성 및 종합의견", r"행\s*동\s*특\s*성|종\s*합\s*의\s*견|행특"),
    ("독서활동", r"독\s*서\s*활\s*동|독서"),
]
# -------------------------------------------------------------------------
# 형광펜·쉬운 해설에 쓸 생기부 본문만 남기기
#   면접 준비에 쓸모가 적고 분량만 차지하는 부분을 뺍니다.
#   - 인적·학적사항, 출결, 수상경력, 자격증, 학교폭력 조치상황, 봉사활동(실적표 포함)
#   - 교과학습발달상황의 성적표 (세부능력 및 특기사항은 남김)
#   - 쪽마다 찍히는 출력 머리글(학교명/출력시각/IP/교사명), 학생 인적사항 줄
#   ※ 면접 문항 생성에는 원문 전체를 그대로 씁니다(성적 추이 같은 약점 분석에 필요).
# -------------------------------------------------------------------------
_TOP_SECTION_RX = re.compile(
    r"^\s*(\d{1,2})\s*[.．]\s*(인\s*적|학\s*적|출\s*결|수\s*상|자\s*격\s*증|학\s*교\s*폭\s*력|"
    r"창\s*의\s*적|교\s*과\s*학\s*습|독\s*서|행\s*동\s*특\s*성)"
)
_KEEP_TOP = ("창의적", "독서", "행동특성")        # 통째로 남기는 영역
_GRADES_TOP = "교과학습"                          # 성적표는 빼고 세특만 남기는 영역
_PRINT_NOISE_RX = [
    # 출력 머리글: 도개고등학교/2026.08.31 15:54/10.51.***.69/김기섭 / 2
    re.compile(r"[가-힣A-Za-z]+(?:고등학교|학교)\s*/\s*\d{4}\.\d{1,2}\.\d{1,2}\s+\d{1,2}:\d{2}\s*/\s*[\d.*]+\s*/\s*[^/]{1,20}/\s*\d+"),
    # 쪽 위 학생 정보: 도개고등학교 2026년 8월 31일 23 반 2 번호 6 성명 홍길동
    re.compile(r"[가-힣A-Za-z]+(?:고등학교|학교)\s+\d{4}년\s*\d{1,2}월\s*\d{1,2}일\s+(?:\d+\s*)?반\s*\d+\s*번호\s*\d+\s*성명\s*\S+"),
]
_SETEUK_RX = re.compile(r"세\s*부\s*능\s*력\s*및\s*특\s*기\s*사\s*항|개\s*인\s*별\s*세\s*부")
_GRADE_TABLE_RX = re.compile(r"단\s*위\s*수|원\s*점\s*수|석\s*차\s*등\s*급|성\s*취\s*도\s*\(|과\s*목\s*평\s*균|수\s*강\s*자\s*수|이\s*수\s*단\s*위")
_VOLUNTEER_RX = re.compile(r"^\s*봉\s*사\s*활\s*동")
_CHANGCHE_SUB_RX = re.compile(r"^\s*(?:\d\s+)?(?:자\s*율|동\s*아\s*리|진\s*로)")

def trim_record_for_review(text):
    """생기부 원문에서 면접 준비에 필요한 서술형 본문만 남깁니다.
    NEIS 형식의 영역 번호(1. 인적·학적사항 … 9. 행동특성 및 종합의견)를 찾지 못하면 원문을 그대로 돌려줍니다."""
    raw = str(text or "")
    lines = raw.splitlines()
    if not any(_TOP_SECTION_RX.match(ln) for ln in lines):
        cleaned = raw
        for rx in _PRINT_NOISE_RX:
            cleaned = rx.sub(" ", cleaned)
        return cleaned
    out, keep, section, in_volunteer = [], True, "", False
    for ln in lines:
        for rx in _PRINT_NOISE_RX:
            ln = rx.sub(" ", ln)
        if not ln.strip():
            continue
        m = _TOP_SECTION_RX.match(ln)
        if m:
            section = re.sub(r"\s+", "", m.group(2))
            in_volunteer = False
            keep = section.startswith(_KEEP_TOP)
            if section.startswith(_GRADES_TOP):
                keep = False          # 성적표부터 시작 → 세특 머리를 만나면 다시 남김
            if keep:
                out.append(ln)
            continue
        if section.startswith(_GRADES_TOP):
            if _SETEUK_RX.search(ln):
                keep = True
            elif _GRADE_TABLE_RX.search(ln):
                keep = False
        elif section.startswith("창의적"):
            if _VOLUNTEER_RX.match(ln):
                in_volunteer = True
            elif in_volunteer and _CHANGCHE_SUB_RX.match(ln):
                in_volunteer = False
            if in_volunteer:
                continue
        if keep:
            out.append(ln)
    return "\n".join(out)

# 지원 대학이 자주 묻는 유형 → 생기부에서 먼저 칠해야 할 곳
_TYPE_AREA_HINT = {
    "서류 활동 확인형": "교과 세특·동아리의 탐구 과정과 결과",
    "전공 개념·지식형": "교과 세특 속 전공 개념어·이론명·법칙명",
    "지원동기·진로형": "진로활동과 진로 희망이 드러난 문장",
    "인성·공동체형": "행동특성 및 종합의견, 자율·봉사활동의 협력·리더십 문장",
    "독서 확인형": "책 제목과 독서 후 이어진 활동",
    "견해·시사형": "사회 이슈를 다룬 탐구 주제와 주장",
    "학업 태도·성적형": "과목 선택 이유·성적 변화와 관련된 문장",
    "상황·딜레마형": "갈등 상황과 본인의 판단이 드러난 문장",
}

def highlight_priority_text(profile, uni_name):
    """지원 대학의 실제 출제 비중으로 '어디부터 칠할지'를 정해 AI에게 알려줍니다."""
    share, basis = pick_allocation_basis(profile)
    if not share:
        return "지원 대학의 출제 자료가 부족하니, 전공 개념어와 탐구 과정을 우선하세요."
    top = sorted(share.items(), key=lambda kv: -kv[1])[:3]
    mix = ", ".join(f"{t} {int(round(p * 100))}%" for t, p in top)
    order = " → ".join(_TYPE_AREA_HINT.get(t, t) for t, _ in top)
    return f"{uni_name}은(는) 실제로 {mix} 순으로 많이 묻습니다(기준: {basis}). 그러므로 문장은 {order} 순서로 우선 고르세요."

def _hl_card_rules(priority_text, with_count=True, with_keywords=True):
    """질문 카드(문장 → 역량 → 질문 → 답변 가이드) 형식과 규칙."""
    rules = f"""형식: {{"구절": "...", "역량": "학업|진로|공동체", "방어": true|false, "영역": "...", "이유": "...", "질문": [{{"질문": "...", "유형": "확인|심화|연결|압박", "난이도": "기본|심화", "가이드": "...", "기존": true|false}}]}}
    규칙:
    1. "구절"은 생기부 원문에서 **글자 그대로 복사**한 10~70자 문장 일부입니다. 단어를 바꾸거나, 줄이거나, 띄어쓰기를 고치지 마세요. 면접관이 짚을 만한 문장(탐구 과정과 결과, 전공 개념어, 책 제목, 수치, 맡은 역할, 진로 변화)을 고르세요.
    2. "역량": 그 문장이 드러내는 학종 평가역량 하나 — 학업(학업 태도·탐구력·전공 개념) / 진로(진로 탐색·전공 관심·진로 연계 활동) / 공동체(협업·리더십·배려·소통)
    3. "방어": 과장된 표현('완벽히', '깊이 이해' 등), 근거가 약한 결론, 앞뒤 연결이 약한 활동처럼 면접관이 공격할 만한 문장이면 true. true이면 질문 하나는 유형 "압박"으로 쓰세요.
    4. "영역": 교과 세특(과목명 포함, 예: 교과 세특-생명과학Ⅰ) / 자율활동 / 동아리활동 / 진로활동 / 행동특성 및 종합의견 / 독서활동
    5. "질문": 그 문장을 근거로 면접관이 실제로 던질 질문 1~2개. "~했다고 했는데"처럼 그 문장을 짚는 말로 시작하세요. [이미 만든 면접 질문] 중 이 문장에서 나온 질문이 있으면 그 질문을 글자 그대로 먼저 넣고 "기존": true로 표시하세요.
       - "유형": 확인(사실·동기·역할) / 심화(개념·원리·한계) / 연결(다른 활동·진로와 잇기) / 압박(약점 공격)
       - "가이드": 답변에 꼭 들어갈 내용 1~2문장(무엇을 어떤 순서로 말할지). 모범답안 전문은 쓰지 마세요.
    6. "이유": 이 문장을 고른 이유 한 문장.
    7. 우선순위: {priority_text}"""
    if with_count:
        rules += f"""
    8. 문장은 {HL_MAX_CARDS - 6}~{HL_MAX_CARDS - 2}개를 생기부의 여러 영역에서 고르게 고르세요. 교과 세특을 꼼꼼히 보되, 근거 없이 개수를 채우지 마세요."""
    if with_keywords:
        rules += """
    9. 맨 마지막 줄에 이 학생 면접의 핵심 키워드 5~8개를 {"키워드": ["...", "..."]} 형식으로 쓰세요."""
    return rules

def _q_lines(questions):
    flat_qs = [re.sub(r"\s+", " ", str(q))[:160] for q in questions if q]
    return "\n".join("- " + q for q in flat_qs) or "- (없음)"

def build_highlight_only_prompt(student_record_text, questions, priority_text):
    """합친 호출에서 카드가 안 나왔을 때 쓰는 '카드만' 요청(JSON 응답 모드)."""
    return f"""당신은 학생부종합전형 면접 전문가입니다. 아래 [생기부 원문]에서 면접관이 짚을 문장을 골라 질문 카드를 만드세요.
    JSON 하나만 출력하세요: {{"items": [카드, ...], "키워드": ["...", "..."]}} ("키워드"는 이 학생 면접의 핵심 키워드 5~8개)
    카드 {_hl_card_rules(priority_text, with_keywords=False)}

    [이미 만든 면접 질문]
    {_q_lines(questions)}

    [생기부 원문]
    {trim_record_for_review(student_record_text)}
    """

def build_gap_cards_prompt(student_record_text, areas, taken_quotes, priority_text):
    """질문이 하나도 없는 영역에만 카드를 더 만듭니다."""
    taken = "\n".join("- " + q for q in taken_quotes) or "- (없음)"
    return f"""당신은 학생부종합전형 면접 전문가입니다. 아래 [생기부 원문]에서 다음 영역은 아직 면접 질문이 하나도 없습니다: {", ".join(areas)}.
    각 영역에서 면접관이 짚을 만한 문장을 1~2개씩 골라 질문 카드를 만드세요. [이미 고른 문장]은 다시 고르지 마세요.
    JSON 하나만 출력하세요: {{"items": [카드, ...]}}
    카드 {_hl_card_rules(priority_text, with_count=False, with_keywords=False)}

    [이미 고른 문장]
    {taken}

    [생기부 원문]
    {trim_record_for_review(student_record_text)}
    """

def build_more_questions_prompt(card, para_text, uni_name, major, priority_text):
    """한 문장에서만 새 질문 2개를 만듭니다(생기부 전체를 보내지 않아 비용이 적습니다)."""
    existing = "\n".join("- " + q["q"] for q in card["qs"]) or "- (없음)"
    return f"""당신은 {uni_name} {major} 학생부종합전형 면접관입니다. 아래 [근거 문장]만을 근거로 새 면접 질문 2개를 만드세요.
    [이미 있는 질문]과 겹치지 않게, 하나는 더 깊이 파고드는 질문(심화 또는 연결), 하나는 다른 각도의 질문으로 만드세요.
    JSON 하나만 출력하세요: {{"질문": [{{"질문": "...", "유형": "확인|심화|연결|압박", "난이도": "기본|심화", "가이드": "..."}}]}}
    - 질문은 "~했다고 했는데"처럼 근거 문장을 짚는 말로 시작합니다.
    - "가이드"는 답변에 꼭 들어갈 내용 1~2문장입니다. 모범답안 전문은 쓰지 마세요.
    - 참고: {priority_text}

    [근거 문장] {card["quote"]}
    [문장이 들어 있는 문단] {para_text}
    [이미 있는 질문]
    {existing}
    """

def _extract_hl_objects(text):
    """카드(구절이 있는 항목)와 키워드 항목만 남깁니다."""
    return [o for o in _extract_json_objects(text)
            if str(o.get("구절", "")).strip() or (o.get("키워드") and "구절" not in o)]

def split_explanation_and_highlights(raw):
    """AI 응답을 (쉬운 해설, 카드·키워드 항목 목록)으로 나눕니다. 카드 부분이 없으면 빈 목록."""
    text = str(raw or "")
    if HL_MARKER in text:
        explanation, _, tail = text.partition(HL_MARKER)
    else:
        # 표시줄을 빠뜨린 경우: 첫 카드 JSON이 시작되는 곳부터를 목록으로 봅니다.
        m = re.search(r'(?:```(?:json)?\s*)?\[?\s*\{\s*"구절"', text)
        if not m:
            return text.strip(), []
        explanation, tail = text[:m.start()], text[m.start():]
    explanation = re.sub(r"[*#`\s]+$", "", explanation.strip())
    return explanation.strip(), _extract_hl_objects(tail)

def _layout_record_text(student_record_text):
    """원문을 읽기 좋은 문단으로 정리합니다. 글자 위치가 바뀌지 않도록 줄바꿈 한 글자를
    공백 또는 문단 나눔 한 글자로만 바꿉니다(그래서 형광펜 위치를 그대로 쓸 수 있습니다)."""
    lines = [re.sub(r"[ \t　]+", " ", ln).strip() for ln in str(student_record_text or "").splitlines()]
    lines = [ln for ln in lines if ln]
    if not lines:
        return ""
    head_rx = re.compile(
        r"^(?:\d+\.|[가-힣ⅠⅡ·()\s]{1,18}\s*[:：]|\(\d학기\)|\[\d학년\]|\d학년|"
        + "|".join(p for _, p in SANGBU_AREAS) + ")"
    )
    out = [lines[0]]
    for prev, ln in zip(lines, lines[1:]):
        new_para = bool(re.search(r"(?:다|음|함|됨|임)\.?$|[.!?]$", prev)) or bool(head_rx.match(ln))
        out.append(("\n" if new_para else " ") + ln)
    return "".join(out)

_HL_SKIP_RX = re.compile(r"[\s\W_]", re.UNICODE)

def _locate_quote(norm, nospace, ns_map, quote):
    q = re.sub(r"\s+", " ", str(quote or "")).strip().strip("\"'“”‘’「」『』")
    q = re.sub(r"(?:\.\.\.|…)+$", "", q).strip()
    if len(q) < 4:
        return None
    i = norm.find(q)
    if i >= 0:
        return i, i + len(q)
    # 띄어쓰기·문장부호를 모두 뺀 글자끼리 비교 (PDF에서 '그 래프', '·'/'∙', 따옴표 모양이 달라진 경우)
    qn = _HL_SKIP_RX.sub("", q)
    if len(qn) < 4:
        return None
    j = nospace.find(qn)
    if j >= 0:
        return ns_map[j], ns_map[j + len(qn) - 1] + 1
    # AI가 구절 끝을 살짝 바꾼 경우: 앞에서부터 가장 길게 맞는 부분(10자 이상)만 칠함
    for cut in range(len(qn) - 1, 9, -1):
        j = nospace.find(qn[:cut])
        if j >= 0:
            return ns_map[j], ns_map[j + cut - 1] + 1
    return None

def _canonical_area(area):
    s = str(area or "")
    for name, pat in SANGBU_AREAS:
        if re.search(pat, s):
            return name
    return "기타"

def _guess_comp(area):
    a = str(area or "")
    if "진로" in a:
        return "진로"
    if re.search(r"자율|동아리|봉사|행동|종합|행특", a):
        return "공동체"
    return "학업"

def _norm_comp(value, area=""):
    v = str(value or "")
    for k in HL_COMPS:
        if k in v:
            return k
    return _guess_comp(area)

def _qkey(question):
    import hashlib
    return hashlib.md5(str(question).encode("utf-8")).hexdigest()[:12]

def _norm_questions(it):
    """카드의 질문 목록을 정리합니다. 예전 형식(질문 한 문장 + 준비)도 읽습니다."""
    raw = it.get("질문")
    if isinstance(raw, str):
        raw = [{"질문": raw, "가이드": it.get("준비", "")}]
    out, seen = [], set()
    for q in raw or []:
        if isinstance(q, str):
            q = {"질문": q}
        if not isinstance(q, dict):
            continue
        text = re.sub(r"\s+", " ", str(q.get("질문", ""))).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        typ = str(q.get("유형", ""))
        out.append({
            "q": text,
            "type": next((t for t in HL_Q_TYPES if t in typ), "확인"),
            "level": "심화" if "심화" in str(q.get("난이도", "")) else "기본",
            "guide": str(q.get("가이드", "") or q.get("준비", "")).strip(),
            "existing": _is_true(q.get("기존")),
            "answer": str(q.get("모범답안", "") or "").strip(),
        })
    return out

_PARA_AREA_RX = [
    ("교과 세특", re.compile(r"세\s*부\s*능\s*력\s*및\s*특\s*기\s*사\s*항")),
    ("자율활동", re.compile(r"자\s*율\s*활\s*동")),
    ("동아리활동", re.compile(r"동\s*아\s*리\s*활\s*동")),
    ("진로활동", re.compile(r"진\s*로\s*활\s*동")),
    ("행동특성 및 종합의견", re.compile(r"행\s*동\s*특\s*성")),
    ("독서활동", re.compile(r"독\s*서\s*활\s*동")),
]

def build_highlight_data(student_record_text, raw_items):
    """질문 카드를 원문 위치에 맞춰 정리하고, 문단·영역별 빈틈을 계산합니다.
    raw_items는 AI가 낸 카드 목록(+키워드 항목)이며, 저장·질문 추가도 이 목록을 기준으로 합니다."""
    text = _layout_record_text(trim_record_for_review(student_record_text))
    norm = text.replace("\n", " ")
    nospace_chars, ns_map = [], []
    for idx, ch in enumerate(norm):
        if not _HL_SKIP_RX.match(ch):
            nospace_chars.append(ch)
            ns_map.append(idx)
    nospace = "".join(nospace_chars)

    keywords, cleaned, dropped, answers = [], [], 0, {}
    for idx, it in enumerate(raw_items or []):
        if not isinstance(it, dict):
            continue
        if isinstance(it.get("모범답안"), dict) and "구절" not in it:
            answers.update({str(k): str(v) for k, v in it["모범답안"].items() if str(v).strip()})
            continue
        if it.get("키워드") and "구절" not in it:
            kw = it["키워드"]
            for k in (kw if isinstance(kw, list) else re.split(r"[,·/]", str(kw))):
                k = str(k).strip().strip("#")
                if k and k not in keywords:
                    keywords.append(k)
            continue
        span = _locate_quote(norm, nospace, ns_map, it.get("구절", ""))
        qs = _norm_questions(it)
        if not span or not qs:
            dropped += 1
            continue
        area = str(it.get("영역", "")).strip() or "기타"
        defend = _is_true(it.get("방어")) or "방어" in str(it.get("색", ""))
        cleaned.append({
            "raw_idx": idx, "start": span[0], "end": span[1], "quote": norm[span[0]:span[1]],
            "comp": _norm_comp(it.get("역량"), area), "defend": defend,
            "area": area, "area_key": _canonical_area(area),
            "reason": str(it.get("이유", "")).strip(), "qs": qs,
        })

    # 겹치는 문장은 앞 카드에 질문을 합치고, 문장 수 상한을 둡니다.
    accepted = []
    for it in cleaned:
        host = next((a for a in accepted if not (it["end"] <= a["start"] or it["start"] >= a["end"])), None)
        if host:
            have = {q["q"] for q in host["qs"]}
            host["qs"].extend(q for q in it["qs"] if q["q"] not in have)
            host["defend"] = host["defend"] or it["defend"]
            continue
        if len(accepted) >= HL_MAX_CARDS:
            dropped += 1
            continue
        accepted.append(it)
    accepted.sort(key=lambda a: a["start"])

    # 문단 나누기 + 문단마다 어느 영역인지(머리말을 만날 때마다 바뀜)
    paras, pos, cur_area = [], 0, ""
    for para in text.split("\n"):
        head = para[:40]
        for name, rx in _PARA_AREA_RX:
            if rx.search(head):
                cur_area = name
                break
        paras.append({"start": pos, "end": pos + len(para), "area": cur_area})
        pos += len(para) + 1

    q_no = 0
    for n, it in enumerate(accepted, 1):
        it["no"] = n
        it["para"] = next((i for i, pa in enumerate(paras) if pa["start"] <= it["start"] <= pa["end"]), 0)
        # 압박 질문은 카드 안에서 맨 뒤로
        it["qs"].sort(key=lambda q: q["type"] == "압박")
        for q in it["qs"]:
            q_no += 1
            q["no"] = q_no
            q["key"] = _qkey(q["q"])
            q["answer"] = answers.get(q["key"]) or q.get("answer", "")

    count = {k: sum(1 for a in accepted if a["comp"] == k) for k in HL_COMPS}
    gaps, gap_areas = [], []
    for name, pat in SANGBU_AREAS:
        if not re.search(pat, norm):
            continue
        cards = [a for a in accepted if a["area_key"] == name]
        nq = sum(len(a["qs"]) for a in cards)
        if not cards:
            status = "⚠️ 빈틈 (질문 없음)"
            gap_areas.append(name)
        else:
            status = "준비됨"
        gaps.append({"영역": name, "칠한 문장": len(cards), "질문": nq, "상태": status})
    subjects = sorted({a["area"].split("-", 1)[1].strip() for a in accepted
                       if a["area_key"] == "교과 세특" and "-" in a["area"]})
    return {"text": text, "items": accepted, "paras": paras, "dropped": dropped, "gaps": gaps,
            "gap_areas": gap_areas, "subjects": subjects, "count": count,
            "defend": sum(1 for a in accepted if a["defend"]), "q_total": q_no,
            "q_keys": [q["key"] for a in accepted for q in a["qs"]], "keywords": keywords[:8]}

def answer_of(q):
    """질문의 모범답안: 따로 만든 것 → 5세트 원문 순서로 찾습니다."""
    return (q.get("answer") or (q.get("detail") or {}).get("모범답안") or "").strip()

def build_answers_prompt(rows, uni_name, major):
    lines = []
    for r in rows:
        lines.append(
            f"- id: {r['id']}\n  질문: {r['q']}\n  근거 문장: {r['quote']}\n"
            f"  문단: {r['para'][:400]}\n  가이드: {r['guide'] or '(없음)'}"
        )
    return f"""당신은 {uni_name} {major}에 지원하는 고등학생의 면접 코치입니다. 아래 질문마다 학생이 실제로 말할 모범답안을 쓰세요.
    규칙:
    1. 학생의 1인칭 구어체, 말했을 때 40~60초 분량(4~6문장)입니다.
    2. 두괄식: 결론 → 근거(생기부 속 활동과 과정) → 배운 점이나 전공·진로와의 연결.
    3. [근거 문장]과 [문단]에 있는 사실만 쓰세요. 없는 수치·책·활동·결과를 지어내지 마세요. 학생이 직접 채워야 할 부분은 [본인 경험: 무엇을 넣을지]처럼 대괄호로 표시하세요.
    4. 가이드가 있으면 그 순서를 따르세요. 압박 질문은 약점을 인정한 뒤 배운 점과 보완 노력으로 답하세요.
    JSON 하나만 출력하세요: {{"answers": [{{"id": "...", "모범답안": "..."}}]}}

    [질문 목록]
    {chr(10).join(lines)}
    """

def _marker(no):
    return _CIRCLED[no - 1] if 1 <= no <= len(_CIRCLED) else f"({no})"

def _shade_run(run, fill_hex):
    rpr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    rpr.append(shd)

_SET_TAG_RX = re.compile(r"\[(질문|평가요소|평가의도|모범답안|꼬리질문)\]")

def sangbu_sets(result_text):
    """생기부 5세트 문항을 [{no, title, 질문, 평가요소, 평가의도, 모범답안, 꼬리질문}]로 읽습니다."""
    out = []
    for block in str(result_text or "").split("### 📌")[1:]:
        lines = block.strip().split("\n")
        title = re.sub(r"[*]", "", lines[0]).strip() if lines else ""
        body = "\n".join(lines[1:])
        parts = _SET_TAG_RX.split(body)
        d = {"title": title}
        for i in range(1, len(parts) - 1, 2):
            d[parts[i]] = parts[i + 1].strip()
        if d.get("질문"):
            d["no"] = len(out) + 1
            out.append(d)
    return out

def attach_set_details(hl, result_text):
    """카드의 '기존 문항'에 5세트 원문의 평가의도·모범답안·꼬리질문을 붙입니다.
    카드에 붙지 못한 세트는 목록으로 돌려줍니다(화면 맨 아래에 따로 보여 줌)."""
    sets = sangbu_sets(result_text)
    used = set()
    for c in (hl or {}).get("items", []):
        for q in c["qs"]:
            q.pop("detail", None)
            q.pop("set_no", None)
            qn = _q_norm(q["q"])
            if not qn:
                continue
            best, best_r = None, 0.0
            for d in sets:
                if d["no"] in used:
                    continue
                dn = _q_norm(d["질문"])
                if not dn:
                    continue
                r = 1.0 if (qn in dn or dn in qn) and min(len(qn), len(dn)) >= 12 else difflib.SequenceMatcher(None, qn, dn).ratio()
                if r > best_r:
                    best, best_r = d, r
            if best is not None and best_r >= 0.72:
                used.add(best["no"])
                q["detail"] = best
                q["set_no"] = best["no"]
                q["existing"] = True
    return [d for d in sets if d["no"] not in used]

def sangbu_briefing(result_text):
    """브리핑 리포트를 {항목명: 내용}으로 읽습니다."""
    m = re.search(r"### 🔍 \[생기부 심층 분석 브리핑 리포트\](.*?)(?=### 📌|### 📄|$)", str(result_text or ""), re.DOTALL)
    if not m:
        return {}
    out = {}
    for mm in re.finditer(r"-\s*\*\*(.+?)[:：]?\*\*\s*[:：]?\s*(.*?)(?=\n\s*-\s*\*\*|\Z)", m.group(1).strip(), re.DOTALL):
        label = mm.group(1).strip().rstrip(":：")
        body = mm.group(2).strip()
        if body:
            out[label] = body
    return out

def _first_sentence(text, limit=110):
    t = re.sub(r"\s+", " ", str(text or "")).strip()
    m = re.search(r"^(.+?[.!?。]|.+?(?:다|음|함)\.)\s", t + " ")
    t = m.group(1) if m else t
    return t[:limit] + ("…" if len(t) > limit else "")

def _set_detail_md(d):
    parts = []
    for key, label in [("평가요소", "🏷️ 평가 요소"), ("평가의도", "🎯 평가 의도"),
                       ("모범답안", "✅ 모범 답안"), ("꼬리질문", "🔥 꼬리질문")]:
        if d.get(key):
            parts.append(f"**{label}**  \n{_md_safe(d[key])}")
    return "\n\n".join(parts)

def _wavy_red(run):
    rpr = run._r.get_or_add_rPr()
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "wave")
    u.set(qn("w:color"), HL_DEFEND["hex"])
    rpr.append(u)

def _q_tags(card, q):
    tags = [HL_COMPS[card["comp"]]["label"], f"{q['type']} 질문", q["level"]]
    if q.get("set_no"):
        tags.append(f"5세트 {q['set_no']}번")
    elif q.get("existing"):
        tags.append("기존 문항")
    return " · ".join(tags)

def _para_html(text, start, end, cards, show_defend):
    import html as _html
    out, cur = [], start
    for c in cards:
        s, e = max(c["start"], start), min(c["end"], end)
        out.append(_html.escape(text[cur:s]))
        style = f"background:{HL_COMPS[c['comp']]['css']};color:#111;padding:1px 2px;border-radius:3px"
        if c["defend"] and show_defend:
            style += f";text-decoration:underline wavy {HL_DEFEND['css']};text-underline-offset:4px"
        out.append(
            f'<sup style="color:#c62828;font-weight:700">{_marker(c["no"])}</sup>'
            f'<mark style="{style}">{_html.escape(text[s:e])}</mark>'
        )
        cur = e
    out.append(_html.escape(text[cur:end]))
    return f'<div style="line-height:1.9;font-size:0.97rem;margin:4px 0 6px">{"".join(out)}</div>'

def _quiet_save(save_args):
    """카드 화면에서의 자동 저장. 저장이 실패해도 화면은 멈추지 않게 합니다(사이드바 저장 상태에 표시됨)."""
    try:
        save_current_session(*save_args, quiet=True)
    except Exception as e:
        st.session_state["last_save_info"] = {"ok": False, "where": "저장 안 됨", "error": str(e),
                                              "at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

def _save_card_check(qkey, widget_key, save_args):
    st.session_state.setdefault("card_checks", {})[qkey] = st.session_state.get(widget_key, "미확인")
    _quiet_save(save_args)

# -------------------------------------------------------------------------
# [2-0b] 🗂️ 면접 전날 최종 요약카드
#   ① 핵심 개념 카드: 생기부에 나온 용어·개념의 뜻 + 몇 학년 어느 과목·활동에서 나왔는지
#                     + 답변할 때 덧붙이면 좋은 말 + 예상 꼬리질문
#   ② 학과 예상 질문 20선 + 모범답안 20선: 기출 DB의 같은·유사 학과 질문을 참고해
#                     학과 빈출 질문 10개 + 이 학생 생기부 질문 10개
#   AI 호출은 두 번(개념 / 질문)이며, 버튼을 눌렀을 때만 만듭니다(비용 절약).
#   만든 카드는 기록 저장소에 'CARD-<묶음키>' 줄로 따로 저장해, 기록을 불러오면 다시 나옵니다.
# -------------------------------------------------------------------------
CARD_CONCEPT_RANGE = (14, 20)
CARD_Q_TOTAL = 20
CARD_Q_TYPES = ["전공 개념", "지원동기·진로", "생기부 활동", "인성·공동체", "시사·견해", "학업 태도"]

def build_card_concept_prompt(record_text, uni, major):
    lo, hi = CARD_CONCEPT_RANGE
    return f"""당신은 {uni} {major} 학생부종합전형 면접 코치입니다. 아래 [생기부]를 읽고, 학생이 면접 전날 마지막으로 훑어볼 '핵심 개념 카드'를 만드세요.
    JSON 하나만 출력하세요:
    {{"개념": [{{"용어": "...", "분류": "전공 개념|이론·법칙|탐구 방법|책·인물|활동·프로그램", "설명": "...", "출처": [{{"학년": "1학년|2학년|3학년|학년 미상", "영역": "...", "내용": "..."}}], "덧붙일 말": "...", "꼬리질문": "..."}}]}}
    규칙:
    1. 생기부에 실제로 적힌 전문 용어·개념·이론·법칙·탐구 방법·책 제목을 {lo}~{hi}개 고르세요. 면접관이 "그게 무엇인가요?"라고 되물을 만한 것, {major} 전공과 이어지는 것을 우선합니다. "용어"는 생기부에 적힌 표기 그대로 쓰세요.
    2. "설명": 그 개념의 정확한 뜻을 2~3문장으로 씁니다. 학생이 면접에서 그대로 말해도 틀리지 않을 만큼 정확하되, 고등학생이 이해할 수 있는 말로 쓰세요.
    3. "출처": 그 용어가 나온 곳마다 하나씩 — "학년"(원문에서 알 수 없으면 "학년 미상"), "영역"(교과 세특-과목명 / 자율활동 / 동아리활동 / 진로활동 / 행동특성 및 종합의견 / 독서활동), "내용"(어떤 활동에서 어떻게 쓰였는지 한 문장). 원문에 없는 출처를 지어내지 마세요.
    4. "덧붙일 말": 이 개념으로 답할 때 추가로 말하면 좋은 내용 1~2문장(개념의 한계, 실생활 사례, 관련 개념과의 차이, {major} 전공과의 연결 등). 생기부에 없는 활동을 했다고 말하게 하지 마세요.
    5. "꼬리질문": 이 개념에 대해 면접관이 이어서 물을 질문 하나.
    6. {major} 전공과 가까운 개념부터 나열하세요.

    [생기부]
    {record_text}
    """

def build_card_qa_prompt(record_text, uni, major, ref_lines, priority_text, bank_lines=None):
    half = CARD_Q_TOTAL // 2
    refs = "\n".join(ref_lines) if ref_lines else f"- (DB에 {major} 유사 학과 기출이 없습니다. {major} 면접에서 흔히 묻는 내용으로 직접 만드세요.)"
    return f"""당신은 {uni} {major} 학생부종합전형 면접관이자 면접 코치입니다. {major} 지원자가 면접에서 받을 가능성이 높은 예상 질문 {CARD_Q_TOTAL}개와, 이 학생의 모범답안 {CARD_Q_TOTAL}개를 만드세요.
    JSON 하나만 출력하세요:
    {{"예상질문": [{{"질문": "...", "은행": "B3 또는 빈 문자열", "유형": "{'|'.join(CARD_Q_TYPES)}", "참고 기출": "K3", "근거": "...", "모범답안": "...", "핵심": ["...", "...", "..."]}}]}}
    규칙:
    1. 정확히 {CARD_Q_TOTAL}개. 앞의 {half}개는 '{major} 학과 면접 빈출 질문'(전공 기초 개념, 지원 동기, 입학 후 학업 계획, 전공 관련 시사 쟁점), 뒤의 {half}개는 '이 학생의 생기부를 파고드는 질문'입니다.
    2. 아래 [실제 기출]에서 자주 나오는 질문 유형과 주제를 우선 반영하고, 참고한 기출 번호(K1, K2 …)를 "참고 기출"에 적으세요(없으면 빈 문자열). 기출 문장을 그대로 베끼지 말고 이 학생에게 맞게 바꾸세요.
    3. 질문은 실제 면접관처럼 짧은 구어체 한두 문장입니다.
    4. "근거": 모범답안이 기대는 생기부 내용(예: "2학년 생명과학Ⅰ 세특 – 효소 실험"). 생기부와 상관없는 질문이면 "없음".
    5. "모범답안": 학생의 1인칭 구어체, 말했을 때 40~60초(4~6문장), 두괄식(결론 → 근거가 되는 생기부 활동 → 배운 점·전공 연결). 생기부에 있는 사실만 쓰고, 학생이 직접 채워야 할 부분은 [본인 경험: …]처럼 표시하세요. 전공 개념을 묻는 질문은 개념 설명이 정확해야 합니다.
    6. "핵심": 답변에서 꼭 말해야 할 키워드 3개.
    7. 참고: {priority_text}
    8. 뒤의 생기부 질문은 아래 [질문 은행]에서 면접에 나올 가능성이 높은 것을 먼저 골라 다듬어 쓰고 "은행"에 번호(예: "B3")를 적으세요. 새로 만든 질문은 "은행"을 빈 문자열로 두세요.

    [질문 은행 — 이 학생을 위해 이미 만든 질문]
    {chr(10).join(bank_lines) if bank_lines else "- (아직 없음)"}

    [실제 기출 — {major} 및 유사 학과]
    {refs}

    [생기부]
    {record_text}
    """

def _collect_items(raw, list_key, item_test):
    items = []
    for o in _extract_json_objects(raw):
        if isinstance(o.get(list_key), list):
            items.extend(x for x in o[list_key] if isinstance(x, dict))
        elif item_test(o):
            items.append(o)
    return items

def _norm_card_concepts(items, record_text):
    rec_ns = _HL_SKIP_RX.sub("", str(record_text or ""))
    out, seen = [], set()
    for o in items:
        term = re.sub(r"\s+", " ", str(o.get("용어", ""))).strip()
        if not term or term in seen:
            continue
        seen.add(term)
        srcs = []
        for s in o.get("출처") or []:
            if isinstance(s, str):
                s = {"내용": s}
            if not isinstance(s, dict):
                continue
            srcs.append({
                "grade": str(s.get("학년", "") or "학년 미상").strip(),
                "area": str(s.get("영역", "")).strip(),
                "what": str(s.get("내용", "")).strip(),
            })
        core = _HL_SKIP_RX.sub("", re.sub(r"[(\[（].*?[)\]）]", "", term))
        out.append({
            "term": term, "kind": str(o.get("분류", "")).strip(), "desc": str(o.get("설명", "")).strip(),
            "srcs": srcs, "add": str(o.get("덧붙일 말", "")).strip(), "follow": str(o.get("꼬리질문", "")).strip(),
            "found": bool(core) and core in rec_ns,      # 생기부 원문에 실제로 있는 용어인지
        })
    return out

def _norm_card_qas(items, refs, idmap=None):
    out = []
    for o in items:
        q = re.sub(r"\s+", " ", str(o.get("질문", ""))).strip()
        a = str(o.get("모범답안", "")).strip()
        if not q or not a:
            continue
        keys = o.get("핵심")
        keys = _as_list(keys) if isinstance(keys, list) else [k.strip() for k in re.split(r"[,·/]", str(keys or "")) if k.strip()]
        ref_ids = re.findall(r"K\d+", str(o.get("참고 기출", "")))
        typ = str(o.get("유형", ""))
        out.append({
            "q": q, "type": next((t for t in CARD_Q_TYPES if t in typ), typ.strip() or "기타"),
            "basis": str(o.get("근거", "")).strip(), "answer": a, "keys": keys[:4],
            "ref": ", ".join(dict.fromkeys(refs[r] for r in ref_ids if r in refs)),
            "from": _bank_from(o.get("은행"), idmap or {}),
        })
    return out[:CARD_Q_TOTAL]

def make_interview_card(record_text, uni, major, student_name, profile, db, api_key, bank=None):
    """핵심 개념 카드와 예상 질문 20선을 만듭니다. 반환: (card dict, 오류 메시지)."""
    rec = mask_personal(trim_record_for_review(record_text), student_name)[:40000]
    share, _ = pick_allocation_basis(profile)
    ex = get_relevant_examples(db, major, uni, top_n=30, type_share=share or None) if db is not None and not db.empty else None
    refs, ref_lines = {}, []
    if ex is not None and len(ex):
        for i, (_, r) in enumerate(ex.iterrows(), 1):
            k = f"K{i}"
            refs[k] = f"{re.split(r'[(（]', str(r['대학']))[0].strip()} {r['학과']}"
            ref_lines.append(f"- {k} [{r['대학']} · {r['학과']}] {re.sub(r'\s+', ' ', str(r['질문']))[:160]}")
    errors, concepts, qas = [], [], []
    try:
        raw = _call_gemini_retry(build_card_concept_prompt(rec, uni, major), api_key, json_mode=True)
        concepts = _norm_card_concepts(_collect_items(raw, "개념", lambda o: "용어" in o), record_text)
        if not concepts:
            errors.append("핵심 개념을 읽지 못했습니다")
    except Exception as e:
        errors.append("핵심 개념: " + _friendly_error(e))
    try:
        b_lines, b_map = bank_prompt_lines(bank or [], exclude_src=("요약카드",))
        b_lines = [mask_personal(l, student_name) for l in b_lines]
        raw = _call_gemini_retry(
            build_card_qa_prompt(rec, uni, major, ref_lines, highlight_priority_text(profile, uni), bank_lines=b_lines),
            api_key, json_mode=True)
        qas = _norm_card_qas(_collect_items(raw, "예상질문", lambda o: isinstance(o.get("질문"), str)), refs, b_map)
        if not qas:
            errors.append("예상 질문을 읽지 못했습니다")
    except Exception as e:
        errors.append("예상 질문: " + _friendly_error(e))
    if not concepts and not qas:
        return None, " / ".join(errors)
    card = {"v": 1, "uni": uni, "major": major, "made": _kst_stamp(), "concepts": concepts, "qas": qas,
            "n_refs": len(refs)}
    return card, " / ".join(errors)

def _card_keywords():
    return list(((st.session_state.get("highlight_data") or {}).get("keywords")) or [])

# -------------------------------------------------------------------------
# [2-0c] 📋 대학 맞춤 면접 질문지 (학생용: 질문 + 생기부 출처 / 교사용: + 모범답안)
#   지원 대학의 실제 질문 방식(말투·길이·압박 수위)과 지원 학과 기출의 전공 깊이를
#   build_reference_strategy로 가져와, 이 학생의 생기부 문장에서 질문을 뽑습니다.
#   질문마다 근거 생기부 문장을 원문 그대로 인용하고, 원문에 실제로 있는지 확인합니다.
#   만든 질문지는 'QSHEET-<묶음키>' 줄로 따로 저장합니다.
# -------------------------------------------------------------------------
QSHEET_COUNTS = [10, 15, 20]
QSHEET_EVAL = ["학업역량", "전공적합성", "인성", "발전가능성"]

def default_qsheet_count(info_row):
    """면접 시간에 맞춰 기본 문항 수를 정합니다(10분 이하 10개, 20분 이상 20개)."""
    try:
        mins = [int(x) for x in re.findall(r"(\d+)\s*분", str(info_row.get("면접시간", "")))]
    except Exception:
        mins = []
    if not mins:
        return 15
    m = max(mins)
    return 10 if m <= 10 else 20 if m >= 20 else 15

def build_qsheet_prompt(record_text, uni, major, n, ref_text, univ_text, bank_lines=None):
    return f"""당신은 {uni} {major} 학생부종합전형 면접관입니다. 아래 [참고자료]로 이 대학의 실제 질문 방식(말투·길이·압박 수위·꼬리질문 깊이)과 {major} 전공 질문의 깊이를 익힌 뒤, 이 학생의 [생기부]에서 질문 {n}개를 뽑아 '실제 면접 질문지'를 만드세요.
    {univ_text}
    {ref_text}

    JSON 하나만 출력하세요:
    {{"질문지": [{{"질문": "...", "은행": "B3 또는 빈 문자열", "유형": "...", "평가요소": "{'|'.join(QSHEET_EVAL)}", "출처": {{"학년": "1학년|2학년|3학년|학년 미상", "영역": "...", "원문": "..."}}, "대학 스타일": "...", "모범답안": "...", "핵심": ["...", "...", "..."], "꼬리질문": "...", "피할 답변": "..."}}]}}
    규칙:
    1. 정확히 {n}개, 실제 면접 흐름 순서로 나열하세요. 첫 1~2개는 지원 동기·전공 관심(생기부 근거), 이어서 생기부 영역(교과 세특·자율·동아리·진로·행동특성·독서)을 고르게 다루고, 마지막 1개는 입학 후 계획이나 마무리 질문입니다.
    2. "질문": 이 대학 면접관이 실제로 말하듯 쓰세요(참고자료 A의 말투·길이). 다른 대학의 말투나 다른 학과의 전공 내용을 섞지 마세요. 기출 문장을 그대로 베끼지 말고 이 학생의 생기부 내용으로 바꾸세요.
    3. "출처": "원문"은 질문의 근거가 된 생기부 문장을 글자 하나 바꾸지 말고 20~60자 그대로 옮기세요. "영역"은 과목명까지(예: "교과 세특-생명과학Ⅰ", "동아리활동", "행동특성 및 종합의견"). "학년"은 원문에서 확인될 때만 쓰고 아니면 "학년 미상". 생기부와 무관한 질문이면 "원문"을 빈 문자열로 두세요.
    4. "유형": 참고자료의 출제 유형 이름(예: 활동 심화형, 전공 개념형, 지원 동기형 등)으로 적고, [실제 출제 유형 분석]이나 [유형 배분]이 있으면 그 비중을 {n}개 전체에 비례해 반영하세요.
    5. "대학 스타일": 이 질문을 왜 이렇게 물었는지 한 문장(예: "이 대학은 활동의 한계와 개선 방안을 되묻는 경우가 많음").
    6. "모범답안": 학생 1인칭 구어체, 말했을 때 40~60초(4~6문장), 두괄식(결론 → 생기부 근거 → 배운 점·전공 연결). 생기부에 있는 사실만 쓰고, 학생이 채워야 할 부분은 [본인 경험: …]으로 표시하세요.
    7. "핵심": 답변에서 꼭 말할 키워드 3개. "꼬리질문": 답변 직후 이어질 압박 질문 하나. "피할 답변": 감점되는 답변 유형 한 문장.
    8. 아래 [질문 은행]은 이 학생을 위해 이미 만든 질문입니다. 먼저 그중에서 이 대학 스타일에 맞고 생기부 근거가 분명한 질문을 골라 이 대학 말투로 다듬어 쓰고, 그 경우 "은행"에 번호(예: "B3")를 적으세요. 은행 질문으로 채우지 못한 생기부 영역만 새로 만들고, 새로 만든 질문은 "은행"을 빈 문자열로 두세요.

    [질문 은행]
    {chr(10).join(bank_lines) if bank_lines else "- (아직 없음 — 모두 새로 만드세요)"}

    [생기부]
    {record_text}
    """

def _norm_qsheet(items, record_text, idmap=None):
    rec_ns = _HL_SKIP_RX.sub("", str(record_text or ""))
    out = []
    for o in items:
        q = re.sub(r"\s+", " ", str(o.get("질문", ""))).strip()
        if not q:
            continue
        src = o.get("출처") or {}
        if isinstance(src, str):
            src = {"원문": src}
        quote = re.sub(r"\s+", " ", str(src.get("원문", "") or "")).strip().strip("\"'“”‘’")
        qn_ = _HL_SKIP_RX.sub("", quote)
        if not qn_:
            found = None                      # 생기부와 무관한 질문(마무리 등)
        elif qn_ in rec_ns:
            found = True
        else:                                 # 앞뒤 절반 중 한쪽이라도 원문에 있으면 '대체로 일치'
            h = max(8, len(qn_) // 2)
            found = "partial" if (qn_[:h] in rec_ns or qn_[-h:] in rec_ns) else False
        ev = str(o.get("평가요소", ""))
        keys = o.get("핵심")
        keys = _as_list(keys) if isinstance(keys, list) else [k.strip() for k in re.split(r"[,·/]", str(keys or "")) if k.strip()]
        out.append({
            "q": q, "type": str(o.get("유형", "")).strip(),
            "eval": next((e for e in QSHEET_EVAL if e in ev), ev.strip()),
            "grade": str(src.get("학년", "") or "").strip() or ("학년 미상" if quote else ""),
            "area": str(src.get("영역", "") or "").strip(), "quote": quote, "found": found,
            "style": str(o.get("대학 스타일", "")).strip(), "answer": str(o.get("모범답안", "")).strip(),
            "keys": keys[:4], "follow": str(o.get("꼬리질문", "")).strip(), "avoid": str(o.get("피할 답변", "")).strip(),
            "from": _bank_from(o.get("은행"), idmap or {}),
        })
    return out

def make_question_sheet(record_text, uni, major, student_name, n, db, info_row, api_key, bank=None):
    """대학 맞춤 면접 질문지를 만듭니다. 반환: (sheet dict, 오류 메시지)."""
    rec = mask_personal(trim_record_for_review(record_text), student_name)[:40000]
    ref = build_reference_strategy(db, uni, major, top_n=12, style_n=8)
    lines, idmap = bank_prompt_lines(bank or [], exclude_src=("질문지",))
    lines = [mask_personal(l, student_name) for l in lines]
    try:
        raw = _call_gemini_retry(
            build_qsheet_prompt(rec, uni, major, n, ref.get("prompt_text", ""), format_univ_info_for_prompt(info_row),
                                bank_lines=lines),
            api_key, json_mode=True)
    except Exception as e:
        return None, "질문지: " + _friendly_error(e)
    qs = _norm_qsheet(_collect_items(raw, "질문지", lambda o: isinstance(o.get("질문"), str)), record_text, idmap)
    if not qs:
        return None, "AI 응답에서 질문을 읽지 못했습니다. 다시 만들어 주세요."
    style = ref.get("style") or {}
    mode_label = {"exact": f"{uni} {major} 실제 기출 기준",
                  "style_transfer": f"{uni} 다른 학과 기출의 질문 방식 + 타 대학 {major} 기출의 전공 깊이",
                  "uni_only": f"{uni} 질문 방식 기준(유사 학과 기출 없음)",
                  "dept_only": f"타 대학 {major} 기출 기준({uni} 기출 없음)"}.get(ref.get("mode"), "기출 자료 부족 — 생기부 중심")
    style_line = ""
    if style:
        tops = ", ".join(f"{t} {r}%" for t, r in style.get("대표유형", []))
        style_line = f"{uni} 기출 {style.get('기출수', 0)}개 분석 · 자주 묻는 유형: {tops} · 평균 질문 길이 {style.get('평균질문길이', 0)}자"
    sheet = {"v": 1, "uni": uni, "major": major, "made": _kst_stamp(), "n": n, "qs": qs,
             "basis": mode_label, "style_line": style_line}
    miss = sum(1 for q in qs if q["found"] is False)
    warn = f"근거 문장 {miss}개가 생기부 원문과 정확히 맞지 않아 ⚠️ 표시했습니다." if miss else ""
    return sheet, warn

def _src_line(q):
    return " · ".join(x for x in (q.get("grade"), q.get("area")) if x)

def _found_mark(q):
    return {True: "", "partial": " (원문과 일부 다름)", False: " ⚠️ 원문 확인 필요", None: ""}.get(q.get("found"), "")

# -------------------------------------------------------------------------
# [2-0d] 🏦 질문 은행 — 한 학생을 위해 만든 모든 질문을 한곳에
#   질문지(AI) · 형광펜 카드 · 요약카드 20선 · 5세트 문항을 같은 모양으로 모으고,
#   거의 같은 질문은 하나로 묶습니다(dup_of). 교사가 '쓸 질문'을 고르면
#   학생 휴대폰 링크 · 모의면접 진행 · 질문지 내보내기가 모두 이 선택을 씁니다.
#   질문지·요약카드를 AI로 만들 때도 은행 질문을 먼저 보여 주고 다듬어 쓰게 합니다.
#   고른 질문과 모의면접 진행 기록은 'QBANK-<묶음키>' 줄에 저장합니다.
# -------------------------------------------------------------------------
QB_SOURCES = ["질문지", "형광펜", "요약카드", "5세트"]      # 같은 질문이면 앞쪽 출처를 대표로 둡니다
QB_DEFAULT_PICK = 15


def _qb_key(prefix, text):
    return f"{prefix}-" + _hashlib.sha1(_q_norm(text).encode("utf-8")).hexdigest()[:10]


def _qb_para(record_text, quote, width=320):
    """근거 문장이 들어 있는 생기부 부분(앞뒤 문맥 포함). 못 찾으면 근거 문장만."""
    t, quote = str(record_text or ""), str(quote or "").strip()
    if not quote or not t:
        return quote
    idx = [i for i, ch in enumerate(t) if not ch.isspace()]
    comp = "".join(t[i] for i in idx)
    qn_ = re.sub(r"\s+", "", quote)
    j = comp.find(qn_)
    if j < 0:
        j = comp.find(qn_[:12]) if len(qn_) >= 12 else -1
    if j < 0:
        return quote
    s = idx[j]
    e = idx[min(j + len(qn_), len(idx)) - 1] + 1
    a = max(t.rfind("\n", 0, s) + 1, s - width)
    b = t.find("\n", e)
    b = min(len(t) if b < 0 else b, e + width)
    return t[a:b].strip()


def _qb_similar(a, b):
    if not a or not b:
        return False
    if (a in b or b in a) and min(len(a), len(b)) >= 12:
        return True
    sm = difflib.SequenceMatcher(None, a, b)
    return sm.real_quick_ratio() >= 0.72 and sm.quick_ratio() >= 0.72 and sm.ratio() >= 0.72


def build_question_bank(hl, result_text, qsheet, card, record_text):
    """모든 출처의 질문을 [{key, src, no, q, type, eval, grade, area, quote, found, para, answer,
    tail, keys, guide, style, avoid, from, dup_of, based_on}]로 모읍니다."""
    bank = []
    for i, q in enumerate((qsheet or {}).get("qs") or [], 1):
        bank.append({"key": _qb_key("qs", q["q"]), "src": "질문지", "no": i, "q": q["q"],
                     "type": q.get("type", ""), "eval": q.get("eval", ""), "grade": q.get("grade", ""),
                     "area": q.get("area", ""), "quote": q.get("quote", ""), "found": q.get("found"),
                     "answer": q.get("answer", ""), "tail": q.get("follow", ""), "keys": q.get("keys") or [],
                     "guide": "", "style": q.get("style", ""), "avoid": q.get("avoid", ""), "from": q.get("from", "")})
    unattached = []
    if hl and hl.get("items"):
        unattached = attach_set_details(hl, result_text)
        for c in hl.get("items", []):
            pa = hl["paras"][c["para"]]
            para = hl["text"][pa["start"]:pa["end"]]
            for q in c["qs"]:
                d = q.get("detail") or {}
                bank.append({"key": q["key"], "src": "형광펜", "no": q["no"], "q": q["q"],
                             "type": q.get("type", ""), "eval": c.get("comp", ""), "grade": "",
                             "area": pa.get("area") or c.get("area_key", ""), "quote": c.get("quote", ""),
                             "found": True, "para": para, "answer": answer_of(q), "tail": d.get("꼬리질문", ""),
                             "keys": [], "guide": q.get("guide", ""), "level": q.get("level", "")})
    else:
        unattached = sangbu_sets(result_text)
    for i, q in enumerate((card or {}).get("qas") or [], 1):
        basis = q.get("basis", "")
        bank.append({"key": _qb_key("cd", q["q"]), "src": "요약카드", "no": i, "q": q["q"],
                     "type": q.get("type", ""), "eval": "", "grade": "", "area": "" if basis in ("", "없음") else basis,
                     "quote": "", "found": None, "answer": q.get("answer", ""), "tail": "",
                     "keys": q.get("keys") or [], "guide": "", "from": q.get("from", "")})
    sets = unattached
    if not sets and not bank:        # 제시문 면접처럼 '### 📌' 세트가 없으면 [질문]/[모범답안] 쌍으로
        sets = [{"no": i, "질문": p["q"], "모범답안": p["a"]} for i, p in enumerate(extract_qa_pairs(result_text), 1) if p.get("q")]
    for d in sets:
        q = re.sub(r"\s+", " ", str(d.get("질문", ""))).strip()
        if q:
            bank.append({"key": _qb_key("st", q), "src": "5세트", "no": d.get("no", 0), "q": q,
                         "type": d.get("title", ""), "eval": d.get("평가요소", ""), "grade": "", "area": "",
                         "quote": "", "found": None, "answer": d.get("모범답안", ""), "tail": d.get("꼬리질문", ""),
                         "keys": [], "guide": d.get("평가의도", "")})
    # 키가 겹치면(같은 문장이 같은 출처에 두 번) 뒤의 것을 버림
    seen, uniq = set(), []
    for e in bank:
        if e["key"] not in seen:
            seen.add(e["key"])
            uniq.append(e)
    bank = uniq
    by_key = {e["key"]: e for e in bank}
    for e in bank:                    # AI가 '은행 질문을 다듬어 썼다'고 밝힌 경우: 원래 질문을 그 아래로 묶음
        tgt = by_key.get(e.get("from") or "")
        if tgt is None or tgt is e or tgt.get("dup_of") or e.get("dup_of"):
            continue
        if tgt["src"] == "질문지" and e["src"] != "질문지":
            e["dup_of"] = tgt["key"]      # 대학 맞춤 질문지는 늘 대표로 남깁니다
        else:
            tgt["dup_of"] = e["key"]
            e["based_on"] = f"{tgt['src']} {tgt['no']}번"
    norm = {e["key"]: _q_norm(e["q"]) for e in bank}
    prim = []
    for e in bank:
        if e.get("dup_of"):
            continue
        hit = next((p for p in prim if _qb_similar(norm[e["key"]], norm[p["key"]])), None)
        if hit is not None:
            e["dup_of"] = hit["key"]
        else:
            prim.append(e)
    for e in bank:
        if not e.get("para"):
            e["para"] = _qb_para(record_text, e.get("quote", ""))
    return bank


def bank_selection(bank):
    """고른 질문 키 목록(은행 순서). 고른 적이 없으면: 질문지가 있으면 질문지, 없으면 앞쪽 15개."""
    keys = {e["key"] for e in bank}
    raw = st.session_state.get("qbank_sel")
    if raw and "__none__" in raw:            # '모두 풀기'를 누른 상태
        return []
    sel = [k for k in (raw or []) if k in keys]
    if not sel:                              # 고른 적이 없거나, 질문을 새로 만들어 예전 선택이 사라진 경우
        prim = [e for e in bank if not e.get("dup_of")]
        qs = [e["key"] for e in prim if e["src"] == "질문지"]
        sel = qs or [e["key"] for e in prim][:QB_DEFAULT_PICK]
    chosen = set(sel)
    return [e["key"] for e in bank if e["key"] in chosen]


def bank_selected_entries(bank):
    chosen = set(bank_selection(bank))
    return [e for e in bank if e["key"] in chosen]


def bank_prompt_lines(bank, exclude_src=(), limit=40):
    """AI에게 보여 줄 은행 질문 목록과 번호→키 표."""
    prim = [e for e in bank if not e.get("dup_of") and e["src"] not in exclude_src][:limit]
    lines, idmap = [], {}
    for i, e in enumerate(prim, 1):
        idmap[f"B{i}"] = e["key"]
        lines.append(f"- B{i} [{e['src']} · {e.get('area') or '생기부 전체'}] {re.sub(r'[ \t\r\n]+', ' ', e['q'])[:140]}")
    return lines, idmap


def _bank_from(raw_value, idmap):
    ids = re.findall(r"B\d+", str(raw_value or ""))
    return idmap.get(ids[0], "") if ids else ""


def _qb_src_line(e):
    parts = [x for x in (e.get("grade"), e.get("area")) if x and x != "학년 미상"]
    return " · ".join(parts)


def bank_as_sheet(entries, uni, major):
    """은행에서 고른 질문을 질문지 문서 모양으로."""
    return {"uni": uni, "major": major, "made": _kst_stamp(), "basis": "질문 은행에서 고른 질문", "style_line": "",
            "qs": [{"q": e["q"], "type": e.get("type", ""), "eval": e.get("eval", ""), "grade": e.get("grade", ""),
                    "area": e.get("area", ""), "quote": e.get("quote", ""), "found": e.get("found"),
                    "style": e.get("style", ""), "answer": e.get("answer", ""), "keys": e.get("keys") or [],
                    "follow": e.get("tail", ""), "avoid": e.get("avoid", "")} for e in entries]}


def run_rows(run, entries):
    by = {e["key"]: e for e in entries}
    return [{"q": by[k]["q"], "grade": run["grades"].get(k, ""), "memo": run["memos"].get(k, "")}
            for k in run["keys"] if k in by]


# =====================================================================  ⑥ 학생 링크: 은행 질문 보내기
def practice_snapshot_from_bank(entries, record_text, name):
    """질문 은행에서 고른 질문을 학생 링크용 모양으로(키는 은행 키 그대로 — 형광펜 질문은 예전 답안과 이어짐)."""
    qs = []
    for n, e in enumerate(entries, 1):
        qs.append({
            "key": e["key"], "no": n, "q": mask_personal(e["q"], name),
            "quote": mask_personal(e.get("quote", ""), name), "area": e.get("area") or e["src"],
            "para": mask_personal(e.get("para") or e.get("quote", ""), name)[:500], "comp": e.get("eval", ""),
            "type": e.get("type", ""), "level": e.get("level", ""), "guide": e.get("guide", ""),
            "answer": mask_personal(e.get("answer", ""), name), "tail": e.get("tail", ""),
        })
    rec = mask_personal(trim_record_for_review(record_text), name) if str(record_text or "").strip() else ""
    if not rec:
        rec = "\n".join(dict.fromkeys(q["para"] for q in qs if q["para"]))
    return {"qs": qs, "record": rec[:15000]}


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

# -------------------------------------------------------------------------
# [2-3] 교사용 면접 평가표 (인쇄해서 손으로 채점할 수 있는 양식)
#   실제 대학 학생부종합전형 면접 평가요소(학업역량·전공적합성·인성·발전가능성)와
#   면접 태도 평가 관행(목소리·자세·시선·태도)을 바탕으로 항목당 10점 만점으로 구성
# -------------------------------------------------------------------------
EVALUATION_RUBRIC = [
    ("내용 영역\n(60점)", [
        ("질문 이해·논리성", "질문의 의도를 정확히 파악하고, 결론과 근거가 분명한 논리적 구조로 답변하는가"),
        ("전공 적합성", "지원 전공에 대한 이해와 관심, 관련 활동 경험이 답변 속에 구체적으로 드러나는가"),
        ("학업 역량", "교과 지식과 탐구 경험을 바탕으로 개념을 정확하고 깊이 있게 설명할 수 있는가"),
        ("기재 내용 신뢰도", "생기부에 기재된 활동을 본인이 실제 수행했음이 구체적 경험 진술로 확인되는가"),
        ("인성·공동체 역량", "협업·배려·성실성·소통 능력이 막연한 다짐이 아닌 구체적 경험으로 드러나는가"),
        ("발전 가능성", "자기주도성, 문제 해결 경험, 진로 계획의 구체성과 실현 가능성이 나타나는가"),
    ]),
    ("태도·표현 영역\n(40점)", [
        ("목소리·전달력", "목소리 크기와 발음이 명확하고, 말끝을 흐리거나 불필요한 군말을 반복하지 않는가"),
        ("자세·시선 처리", "바른 자세를 유지하고 면접관과 자연스럽게 시선을 맞추는가"),
        ("표정·면접 태도", "안정된 표정과 예의 바른 태도를 유지하며, 질문을 끝까지 경청하는가"),
        ("답변 속도·시간 관리", "적절한 속도로 답변 시간(1~2분)을 지키며 핵심을 우선 전달하는가"),
    ]),
]

# -------------------------------------------------------------------------
# [2-2] 여러 회차의 저장 기록을 모아 학생의 "성장 리포트"를 만들기
#   (새 저장소를 만들지 않고, 기존 list_records()/get_record()로 이미 쌓인 기록을 재사용합니다)
# -------------------------------------------------------------------------
def build_growth_report_prompt(student_name, sessions, practice_text=""):
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
    joined = "\n\n".join(session_blocks) or "(모의면접 대화 기록 없음)"
    if practice_text:
        joined += "\n\n[휴대폰 답안 연습 기록 — 학생이 직접 쓰거나 말한 답을 AI가 근거·구조·구체성·전공연결 4기준(각 5점)으로 첨삭한 결과]\n" + practice_text
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
STU_DAILY_LIMIT = 30
FOLLOW_MAX = 3                 # 꼬리질문 이어 답하기 최대 단계


def practice_snapshot(hl, result_text, record_text, name):
    """학생에게 보낼 질문·근거·모범답안과, 첨삭 때 대조할 생기부(이름 가림)."""
    attach_set_details(hl, result_text)
    qs = []
    for c in hl.get("items", []):
        pa = hl["paras"][c["para"]]
        para = hl["text"][pa["start"]:pa["end"]]
        for q in c["qs"]:
            d = q.get("detail") or {}
            qs.append({
                "key": q["key"], "no": q["no"], "q": mask_personal(q["q"], name),
                "quote": mask_personal(c["quote"], name), "area": pa["area"] or c["area_key"],
                "para": mask_personal(para, name)[:500], "comp": c["comp"], "type": q["type"],
                "level": q["level"], "guide": q.get("guide", ""), "answer": mask_personal(answer_of(q), name),
                "tail": d.get("꼬리질문", ""),
            })
    rec = mask_personal(trim_record_for_review(record_text), name) if str(record_text or "").strip() else ""
    if not rec:
        seen, parts = set(), []
        for q in qs:
            if q["para"] not in seen:
                seen.add(q["para"])
                parts.append(q["para"])
        rec = "\n".join(parts)
    return {"qs": qs, "record": rec[:15000]}


def build_feedback_prompt(ctx, answer, sec=None, follow=None, voice=False):
    """ctx: {q, quote, para, record, uni, major}. follow: {q, prev}이면 꼬리질문에 대한 답 첨삭."""
    if follow:
        qblock = f"[원래 질문] {ctx['q']}\n    [학생의 앞선 답변] {follow['prev']}\n    [꼬리질문] {follow['q']}\n    (학생은 꼬리질문에 답했습니다. 꼬리질문에 대한 답을 첨삭하세요.)"
    else:
        qblock = f"[질문] {ctx['q']}"
    ans_block = "[학생 답변] 함께 보낸 음성 파일입니다. 먼저 학생이 말한 그대로 받아 적으세요." if voice else f"[학생 답변]\n    {answer}"
    time_note = f"\n    [말한 시간] {sec}초 (권장 40~60초)" if sec else ""
    return f"""당신은 {ctx['uni']} {ctx['major']} 학생부종합전형 면접관이자, 고등학생을 지도하는 면접 코치입니다.
    학생의 답변을 첨삭하세요. 학생에게 직접 말하듯 친절하지만 구체적으로 쓰세요.

    {qblock}
    [근거 문장(생기부)] {ctx['quote']}
    [근거 문단] {ctx['para']}
    {ans_block}{time_note}

    평가 기준(각 1~5점, 5가 가장 좋음):
    - 근거: 생기부 속 활동을 근거로 들었는가
    - 구조: 결론을 먼저 말하고 근거 → 배운 점 순서로 말했는가
    - 구체성: 과정·맡은 역할·수치가 구체적인가
    - 전공연결: {ctx['major']}·진로와 자연스럽게 이어지는가
    JSON 하나만 출력하세요:
    {{"받아쓰기": "...", "점수": {{"근거": 3, "구조": 3, "구체성": 3, "전공연결": 3}}, "잘한 점": "...", "고칠 점": ["...", "..."], "다듬은 답안": "...", "생기부와 다른 내용": ["..."], "꼬리질문": "..."}}
    규칙:
    1. "받아쓰기"는 음성 답변일 때만 학생이 말한 그대로 적고, 글 답변이면 빈 문자열로 두세요.
    2. "고칠 점"은 2~3개, 무엇을 어떻게 바꾸면 되는지 한 문장씩 쓰세요.
    3. "다듬은 답안"은 학생 답변의 내용과 말투를 살려 40~60초 분량으로 다듬되, 생기부에 없는 사실을 새로 넣지 마세요. 학생이 채워야 할 곳은 [본인 경험: …]으로 표시하세요.
    4. "생기부와 다른 내용"에는 학생 답변 중 아래 [생기부]에 없거나 어긋나는 사실(활동, 역할, 수치, 책 제목, 결과)만 학생 말 그대로 인용하세요. 느낌·생각·다짐은 넣지 마세요. 없으면 빈 목록입니다.
    5. "꼬리질문"은 이 답변을 듣고 면접관이 바로 이어서 물을 질문 하나입니다.
    6. 답변이 너무 짧거나 질문과 상관없으면 점수를 낮게 주고, '고칠 점'에 무엇을 말해야 하는지 알려 주세요.

    [생기부]
    {ctx['record']}
    """


def normalize_feedback(raw):
    obj = next((o for o in _extract_json_objects(raw) if isinstance(o.get("점수"), dict)), None)
    if not obj:
        return None
    scores = {}
    for k in FB_CRITERIA:
        try:
            v = int(round(float(obj["점수"].get(k, 1))))
        except Exception:
            v = 1
        scores[k] = max(1, min(5, v))
    return {
        "scores": scores, "total": sum(scores.values()),
        "good": str(obj.get("잘한 점", "")).strip(), "fixes": _as_list(obj.get("고칠 점")),
        "better": str(obj.get("다듬은 답안", "")).strip(), "mismatch": _as_list(obj.get("생기부와 다른 내용")),
        "follow": str(obj.get("꼬리질문", "")).strip(), "transcript": str(obj.get("받아쓰기", "")).strip(),
    }


def run_feedback(ctx, api_key, answer="", audio=None, sec=None, follow=None):
    """글 또는 음성 답변을 첨삭합니다. 실패하면 (None, 이유)."""
    try:
        if audio is not None:
            prompt = build_feedback_prompt(ctx, "", sec=sec, follow=follow, voice=True)
            parts = [types.Part.from_bytes(data=audio, mime_type="audio/wav"), prompt]
            try:
                raw = call_gemini_parts(parts, api_key)
            except Exception as e:
                if not _RATE_LIMIT_RX.search(str(e)):
                    raise
                time.sleep(25)
                raw = call_gemini_parts(parts, api_key)
        else:
            raw = _call_gemini_retry(build_feedback_prompt(ctx, answer, follow=follow), api_key, json_mode=True)
    except Exception as e:
        return None, _friendly_error(e)
    fb = normalize_feedback(raw)
    if not fb:
        return None, "AI 첨삭 결과를 읽지 못했습니다. 다시 시도해 주세요."
    return fb, ""


def _wav_seconds(data):
    try:
        with _wave.open(io.BytesIO(data)) as w:
            return int(round(w.getnframes() / float(w.getframerate())))
    except Exception:
        return None


def load_class_practice():
    out = []
    for r in practice_answer_rows():
        _, obj = _load_json_row(_row_get(r, "id", ""))
        if not obj or not can_access_record(obj.get("record_id", "")):
            continue
        out.append({"name": str(_row_get(r, "student_name", "")), "uni": str(_row_get(r, "university", "")),
                    "major": str(_row_get(r, "major", "")), "ans": obj})
    return out


def _practice_stats(ans):
    items = [v for v in ans.get("q", {}).values() if v.get("tries")]
    lasts = [v["tries"][-1].get("fb", {}) for v in items]
    firsts = [v["tries"][0].get("fb", {}) for v in items]
    if not items:
        return None
    avg = lambda fs, k: round(sum(f.get("scores", {}).get(k, 0) for f in fs) / len(fs), 2)
    return {
        "answered": len(items), "tries": sum(len(v["tries"]) for v in items),
        "follow": sum(len(v["tries"][-1].get("follow", [])) for v in items),
        "first_total": round(sum(f.get("total", 0) for f in firsts) / len(firsts), 1),
        "last_total": round(sum(f.get("total", 0) for f in lasts) / len(lasts), 1),
        "crit": {k: avg(lasts, k) for k in FB_CRITERIA},
        "mismatch": sum(1 for f in lasts if f.get("mismatch")),
        "updated": ans.get("updated", ""),
    }


def practice_summary_for_growth(student_name):
    """성장 리포트에 넣을 '휴대폰 답안 연습' 요약 (그 학생 기록이 없으면 빈 문자열)."""
    blocks = []
    for d in load_class_practice():
        if d["name"] != student_name:
            continue
        s = _practice_stats(d["ans"])
        if not s:
            continue
        lines = [f"[{d['uni']} {d['major']}] 답한 질문 {s['answered']}개, 첨삭 {s['tries']}회, "
                 f"평균 점수 처음 {s['first_total']} → 최근 {s['last_total']} (20점 만점), "
                 f"기준별 최근 평균 " + ", ".join(f"{k} {v}" for k, v in s["crit"].items())
                 + f", 생기부와 다른 내용이 지적된 답 {s['mismatch']}개"]
        for v in list(d["ans"].get("q", {}).values())[:12]:
            tries = v.get("tries", [])
            if not tries:
                continue
            fixes = "; ".join(tries[-1].get("fb", {}).get("fixes", [])[:2])
            lines.append(f"- {v.get('q', '')[:50]} | 점수 " + "→".join(str(t.get('fb', {}).get('total', '-')) for t in tries)
                         + (f" | 최근 고칠 점: {fixes}" if fixes else ""))
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks)


__all__ = ["QTYPE_ORDER", "ALL_UNI_REGION", "DIRECT_REGION", "UNI_REGION_ORDER", "extract_text_from_pdf", "add_intro_paragraphs", "HL_MARKER", "HL_COMPS", "HL_DEFEND", "HL_Q_TYPES", "HL_MAX_CARDS", "_CIRCLED", "_CHECK_OPTS", "SANGBU_AREAS", "_TOP_SECTION_RX", "_KEEP_TOP", "_GRADES_TOP", "_PRINT_NOISE_RX", "_SETEUK_RX", "_GRADE_TABLE_RX", "_VOLUNTEER_RX", "_CHANGCHE_SUB_RX", "trim_record_for_review", "_TYPE_AREA_HINT", "highlight_priority_text", "_hl_card_rules", "_q_lines", "build_highlight_only_prompt", "build_gap_cards_prompt", "build_more_questions_prompt", "_extract_hl_objects", "split_explanation_and_highlights", "_layout_record_text", "_HL_SKIP_RX", "_locate_quote", "_canonical_area", "_guess_comp", "_norm_comp", "_qkey", "_norm_questions", "_PARA_AREA_RX", "build_highlight_data", "answer_of", "build_answers_prompt", "_marker", "_shade_run", "_SET_TAG_RX", "sangbu_sets", "attach_set_details", "sangbu_briefing", "_first_sentence", "_set_detail_md", "_wavy_red", "_q_tags", "_para_html", "_quiet_save", "_save_card_check", "CARD_CONCEPT_RANGE", "CARD_Q_TOTAL", "CARD_Q_TYPES", "build_card_concept_prompt", "build_card_qa_prompt", "_collect_items", "_norm_card_concepts", "_norm_card_qas", "make_interview_card", "_card_keywords", "QSHEET_COUNTS", "QSHEET_EVAL", "default_qsheet_count", "build_qsheet_prompt", "_norm_qsheet", "make_question_sheet", "_src_line", "_found_mark", "QB_SOURCES", "QB_DEFAULT_PICK", "_qb_key", "_qb_para", "_qb_similar", "build_question_bank", "bank_selection", "bank_selected_entries", "bank_prompt_lines", "_bank_from", "_qb_src_line", "bank_as_sheet", "run_rows", "practice_snapshot_from_bank", "extract_qa_pairs", "EVALUATION_RUBRIC", "build_growth_report_prompt", "STU_DAILY_LIMIT", "FOLLOW_MAX", "practice_snapshot", "build_feedback_prompt", "normalize_feedback", "run_feedback", "_wav_seconds", "load_class_practice", "_practice_stats", "practice_summary_for_growth"]
