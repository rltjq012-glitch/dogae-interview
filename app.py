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

# -------------------------------------------------------------------------
# [0] 실제 대학 면접 기출 통합 DB (전국 140개 대학 / 약 1.6만 개 질의응답)
#     master_interview_qa.csv 파일을 app.py와 같은 폴더에 두면 자동으로 로드됩니다.
# -------------------------------------------------------------------------
APP_DIR = os.path.dirname(os.path.abspath(__file__))

QA_COLUMNS = ["대학", "학과", "전형", "질문", "답변", "원본파일"]
INFO_COLUMNS = ["대학", "학과", "전형", "면접유형", "면접시간", "면접위원", "면접절차", "유의사항", "선배조언"]

def _find_data_files(folder, stem):
    """stem으로 시작하는 csv를 모두 찾습니다.
    GitHub에 올릴 때 이름이 'master_interview_qa (1).csv'처럼 바뀌어도 자동으로 인식하기 위함."""
    found = []
    for name in sorted(os.listdir(folder)) if os.path.isdir(folder) else []:
        low = name.lower()
        if low.startswith(stem.lower()) and low.endswith(".csv"):
            found.append(os.path.join(folder, name))
    return found

@st.cache_data(show_spinner=False)
def load_exam_db(folder):
    """앱 폴더에 있는 master_interview_qa*.csv 파일을 모두 읽어 하나로 합칩니다.
    (구버전·신버전이 같이 올라가 있어도 중복을 걸러 최신 내용까지 모두 사용)
    반환: (통합 데이터프레임, 읽어들인 파일명 목록)"""
    frames, loaded = [], []
    for path in _find_data_files(folder, "master_interview_qa"):
        try:
            d = pd.read_csv(path, encoding="utf-8-sig").dropna(subset=["질문", "답변"])
            if not d.empty:
                frames.append(d)
                loaded.append(os.path.basename(path))
        except Exception:
            continue
    if not frames:
        return pd.DataFrame(columns=QA_COLUMNS), []
    df = pd.concat(frames, ignore_index=True).drop_duplicates(subset=["대학", "학과", "질문"], keep="first")
    return prepare_exam_db(df), loaded

@st.cache_data(show_spinner=False)
def load_univ_info(folder):
    """대학별 면접 형식 정보(면접 시간 · 면접위원 수 · 절차 · 유의사항 · 선배 조언).
    univ_interview_info*.csv 파일을 모두 읽어 합칩니다."""
    frames, loaded = [], []
    for path in _find_data_files(folder, "univ_interview_info"):
        try:
            d = pd.read_csv(path, encoding="utf-8-sig")
            if not d.empty:
                frames.append(d)
                loaded.append(os.path.basename(path))
        except Exception:
            continue
    if not frames:
        return pd.DataFrame(columns=INFO_COLUMNS), []
    df = pd.concat(frames, ignore_index=True).drop_duplicates(subset=["대학", "학과", "전형"], keep="first")
    return df, loaded

# -------------------------------------------------------------------------
# [0-1] 대학명 · 학과명 정규화
#   기출 DB에는 같은 학교가 '서울대학교(서울)' / '서울대학교' / '서울대' 처럼 여러 표기로
#   들어 있습니다(원래 229개 표기 → 실제 140개 학교). 표기가 달라도 같은 학교로 묶고,
#   반대로 입시가 분리된 캠퍼스(건국대 글로컬, 동국대 WISE 등)는 다른 학교로 구분합니다.
# -------------------------------------------------------------------------
UNI_ALIASES = {
    "카이스트": "한국과학기술원", "kaist": "한국과학기술원",
    "유니스트": "울산과학기술원", "unist": "울산과학기술원",
    "디지스트": "대구경북과학기술원", "dgist": "대구경북과학기술원",
    "지스트": "광주과학기술원", "gist": "광주과학기술원",
    "포스텍": "포항공과", "postech": "포항공과", "포항공": "포항공과",
    "켄텍": "한국에너지공과", "kentech": "한국에너지공과", "한국에너지공": "한국에너지공과",
    "코리아텍": "한국기술교육", "koreatech": "한국기술교육",
    "서울과기": "서울과학기술", "금오공": "금오공과", "안동": "경국",
    "대가": "대구가톨릭", "교원": "한국교원", "항공": "한국항공", "외": "한국외국어", "한국외": "한국외국어",
    "시립": "서울시립", "연": "연세", "고": "고려", "성": "성균관", "건": "건국", "동": "동국", "홍": "홍익",
    "중": "중앙", "숙": "숙명여자", "이": "이화여자", "이화여": "이화여자", "숙명여": "숙명여자",
    "성신여": "성신여자", "덕성여": "덕성여자", "동덕여": "동덕여자", "서울여": "서울여자",
}
# 본교와 입시가 분리된 캠퍼스만 '다른 학교'로 취급합니다. (서울·대구 같은 단순 소재지 표기는 무시)
SEPARATE_CAMPUS = {
    ("건국", "글로컬"), ("동국", "WISE"), ("한양", "ERICA"), ("고려", "세종"),
    ("홍익", "세종"), ("연세", "미래"), ("단국", "천안"), ("상명", "천안"),
}
_CAMPUS_TOKENS = [
    ("글로컬", r"글로컬|충주"), ("WISE", r"wise|와이즈|경주"), ("ERICA", r"erica|에리카|안산"),
    ("세종", r"세종"), ("미래", r"미래|원주"), ("천안", r"천안"),
]

def uni_key(name):
    """대학명을 비교용 키로 바꿉니다.
    예) '서울대학교(서울)'·'서울대' → '서울' / '국립금오공과대학교(구미)'·'금오공대' → '금오공과'
        '건국대학교(글로컬)' → '건국|글로컬' / '카이스트(KAIST)' → '한국과학기술원'"""
    if not isinstance(name, str) or not name.strip():
        return ""
    raw = name.strip()
    low = raw.lower()
    base = re.split(r"[(\[（]", raw)[0].strip()
    tokens = base.split()
    if len(tokens) > 1 and re.search(r"(대학교|대학|대)$", tokens[0]):
        base = tokens[0]                      # '한양대 에리카', '동국대 WISE캠퍼스' → 앞 토큰만 학교명
    base = re.sub(r"\s+", "", base)
    base = re.sub(r"(?i)(wise|erica)?캠퍼스$", "", base)
    base = re.sub(r"(?i)(wise|erica)$", "", base)
    if not base:
        base = re.sub(r"[()\[\]（）\s]", "", raw)
    if base.startswith("국립") and len(base) > 4:
        base = base[2:]
    for suf in ("대학교", "대학", "대"):
        if base.endswith(suf) and len(base) - len(suf) >= 1:
            base = base[: -len(suf)]
            break
    base = UNI_ALIASES.get(base.lower(), base)
    m = re.search(r"[(（]([A-Za-z]+)[)）]", raw)      # '카이스트(KAIST)', '…(KENTECH)'
    if m and m.group(1).lower() in UNI_ALIASES:
        base = UNI_ALIASES[m.group(1).lower()]
    for campus, pat in _CAMPUS_TOKENS:
        if (base, campus) in SEPARATE_CAMPUS and re.search(pat, low):
            return f"{base}|{campus}"
    return base

def uni_keys_match(a, b):
    """두 대학 키가 같은 학교인지 판정합니다. 캠퍼스가 다르면 다른 학교입니다.
    완전히 같거나, 앞부분이 같고 길이 차이가 2글자 이하일 때만 인정합니다.
    (한국외 ↔ 한국외국어 ✅ / 대구교 ↔ 대구교육 ✅ / 서울 ↔ 서울과학기술 ❌)"""
    if not a or not b:
        return False
    if a == b:
        return True
    (ba, _, ca), (bb, _, cb) = a.partition("|"), b.partition("|")
    if ca != cb:
        return False
    short, long_ = (ba, bb) if len(ba) <= len(bb) else (bb, ba)
    return long_.startswith(short) and len(short) >= 3 and (len(long_) - len(short)) <= 2

def _normalize_uni(name):
    return uni_key(name)

def _uni_matches(query_uni, db_uni):
    return uni_keys_match(uni_key(query_uni), uni_key(db_uni))

def dept_key(name):
    """학과명에서 꼬리말을 떼어 비교용 키로 만듭니다. 최소 2글자는 남깁니다.
    예) 간호학과 → 간호 / 화학과 → 화학 (예전에는 '화' 한 글자만 남아 '문화…' 학과와 잘못 묶였습니다)
        소비자아동학부 (소비자학전공) → 소비자아동"""
    if not isinstance(name, str):
        return ""
    s = re.split(r"[(\[（]", name.strip())[0]
    s = re.sub(r"[\s·ㆍ・,/\-]+", "", s)
    for _ in range(2):
        for suf in ("학과", "학부", "전공", "계열", "과", "부", "학"):
            if s.endswith(suf) and len(s) - len(suf) >= 2:
                s = s[: -len(suf)]
                break
        else:
            break
    return s

def _normalize_dept(name):
    return dept_key(name)

# 계열 분류: 더 구체적인 계열을 먼저 검사합니다.
# (수학교육과 → 교육 / 물리치료학과 → 의약·보건 / 화학공학과 → 공학)
FIELD_RULES = [
    ("자유전공·통합", r"모집\s*단위|단일\s*계열|무학과|기초학부|자유전공|자율전공|무은재|학부대학|열린전공|광역"),
    ("교육", r"교육|사범|교직|초등|유아|특수"),
    ("의약·보건", r"의예|의학|의과|한의|치의|치과|약학|제약|수의|간호|보건|의료|물리치료|작업치료|임상|방사선|치위생|응급|재활|언어치료|안경|병원|바이오메디|의생명|헬스"),
    ("예체능", r"미술|음악|체육|디자인|무용|연극|영화|영상|실용|스포츠|조형|회화|공예|만화|애니|패션|의류|의상|뷰티|성악|작곡|태권도|운동|레저|공연|게임|사진|예술|Fine Arts"),
    ("공학", r"공학|공과|기계|전자|전기|컴퓨터|소프트웨어|건축|토목|건설|재료|소재|반도체|통신|AI(?!인문)|인공지능|데이터|로봇|항공우주|항공기|항공운항|항공교통|항공정비|조선|자동차|에너지|화공|고분자|나노|모빌리티|드론|보안|정보보호|ICT|소방|안전|교통|철도|도시공|도시계획|메카|전파|광학|광전|배터리|디스플레이|산업공|MSDE|IT융합|이공"),
    ("농생명", r"농업|농학|농생|원예|축산|동물|산림|식품|영양|조경|식물|수산|해양|바이오|생명자원|조리|외식|환경|스마트팜"),
    ("자연과학", r"수학|수리|물리|화학|생물|생명|지구|지질|통계|천문|대기|과학|자연|화장품"),
    ("사회과학", r"경영|경제|행정|정치|외교|사회|미디어|언론|신문|방송|광고|홍보|무역|통상|법|복지|심리|관광|호텔|부동산|회계|세무|금융|경찰|군사|국방|국제|글로벌|Global|소비자|아동|가족|지리|문헌정보|상담|청소년|물류|유통|비서|항공서비스|항공운송|커뮤니케이션|콘텐츠|문화|인류|공공|인재|기업|비즈니스|Business|GBT|벤처|휴먼|서비스|산업"),
    ("인문·어문", r"국어|국문|영어|영문|불어|불문|독어|독문|중어|중문|중국|일어|일문|일본|노어|노문|러시아|서어|스페인|어문|언어|문학|사학|역사|철학|문헌|고고|인문|한문|종교|신학|기독교|문예|창작|번역|통역|아랍|베트남|인도|태국|한국어|고전|미학|어과|어학|프랑스|몽골|터키|튀르키예|헝가리|우크라이나|루마니아|세르비아|아프리카|유럽|독일|ELLT|TESL|EICC|Language|유산"),
]
_FIELD_COMPILED = [(n, re.compile(p)) for n, p in FIELD_RULES]

def field_of(dept_name):
    """학과명을 큰 계열로 분류합니다."""
    s = str(dept_name or "")
    for name, rx in _FIELD_COMPILED:
        if rx.search(s):
            return name
    return "기타"

def _field_of(dept_name):
    return field_of(dept_name)

def track_key(name):
    """전형명을 비교용으로 줄입니다. '학생부종합(지역균형)전형'·'지역균형전형' → '지역균형'"""
    s = re.sub(r"\s+", "", str(name or ""))
    if not s or s.lower() == "nan":
        return "전형 미상"
    kind = "교과" if "교과" in s else ""
    m = re.search(r"[(（]([^()（）]+)[)）]?", s)
    inner = m.group(1) if m else re.sub(r"학생부(종합|교과)", "", s)
    inner = re.sub(r"(전형|학생부종합|학생부교과)", "", inner)
    inner = re.sub(r"[ⅠⅡ]$|(?<=[가-힣])[I1-2]$", "", inner).strip("-_·")
    inner = re.sub(r"[A-Za-z]+", lambda m: m.group(0).upper(), inner)
    if not inner:
        inner = "학생부교과" if kind else "학생부종합"
        kind = ""
    return f"{inner}(교과)" if kind else inner

# -------------------------------------------------------------------------
# [0-2] 기출 질문 유형 분류
#   수험생이 복기한 질문 문장을 키워드 규칙으로 분류합니다(위에서부터 먼저 걸리는 유형).
#   사람이 하나하나 판정한 것이 아니라 근사치이므로, 비중은 '경향'으로만 읽어야 합니다.
# -------------------------------------------------------------------------
QTYPE_JESIMUN = "제시문·구술 문제형"
QTYPE_INTRO = "자기소개·마무리형"
QTYPE_ETC = "후속·기타"
QTYPE_RULES = [
    (QTYPE_JESIMUN, r"제시문|지문|\((?:가|나|다|라|마|바)\)|[㉠㉡㉢㉣]|밑줄\s*친|(?:다음|아래|주어진)\s*(?:자료|표|그림|글|그래프|상황|사례)|자료를\s*보고|자료의\s*그래프|그래프(?:를|가|는|에서)\s*(?:보|해석|나타|제시)|도표|(?:문제|문항)\s*\d|\d\s*번\s*(?:문제|문항)|구하시오|증명하시오|하시오\."),
    (QTYPE_INTRO, r"자기\s*소개(?!서)|마지막으로|하고\s*싶은\s*말|마지막\s*(?:질문|한\s*마디)|끝으로|(?:1|일)\s*분\s*(?:동안|간|자기|안에|정도)|못\s*다\s*한|준비(?:한|해\s*온)\s*(?:말|답변)|포부"),
    ("상황·딜레마형", r"(?:이|그|이런|그런|다음|아래|주어진|이러한|그러한)\s*상황(?:이라면|에서|일\s*때|에\s*(?:처|놓))|상황이라면|어떻게\s*(?:대처|대응|행동|지도|설득|조치)\s*(?:할|하겠|하시겠|해야|하실)|어떻게\s*(?:하시겠|하겠|할\s*것|할\s*건|하실)|(?:만약|만일).{0,45}(?:라면|다면)|(?:교사|선생님|간호사|의사|담임|팀장|반장|리더|경찰|군인|장교)(?:가|이)\s*(?:된다면|되었을\s*때|됐을\s*때)|라면\s*어떻게"),
    ("독서 확인형", r"책|독서|읽고|읽은|읽었|저자|작가|도서|『|《"),
    ("인성·공동체형", r"갈등|협력|협업|협동|리더십|리더쉽|리더로|배려|소통|봉사|나눔|공동체|희생|양보|(?:힘들|어려웠)던\s*(?:점|경험|일|적|순간|것|부분)|힘든\s*(?:점|일|순간)|극복|실패|좌절|장점|단점|장단점|강점|약점|성격|가치관|좌우명|존경하는|인성|책임감|인간관계|의견\s*(?:차이|충돌|대립)|친구(?:들)?(?:와|과|를|에게|가|이|랑|한테)|스트레스(?:를|는|가)?\s*(?:받|풀|해소)|취미|특기"),
    ("지원동기·진로형", r"지원\s*(?:동기|계기|한\s*이유|하게|했)|왜\s*(?:우리|이|저희|본)\s*(?:학교|학과|대학|학부|전공)|(?:우리|저희|본)\s*(?:학교|학과|대학|학부)|이\s*(?:학과|학부|전공)(?:에|를)|진로|장래|꿈|희망\s*(?:직업|분야|진로)|되고\s*싶|졸업\s*(?:후|하고|하면)|입학\s*(?:후|하면|하게|해서|한다면)|학업\s*계획|(?:대학|학교)에\s*(?:와서|들어와|입학)|인재상|10\s*년\s*(?:후|뒤)|왜\s*[가-힣]{1,8}(?:학과|학부|전공|과)(?:에|를|인|입)|(?:교사|교직|간호사|의사|약사|교육자|공학자|연구자|기자|경찰|군인)(?:에게|의|가|로서)?\s*(?:필요한|중요한|갖춰야|가져야)?\s*(?:자질|역량|덕목)|다른\s*(?:학교|대학).{0,10}(?:지원|합격)|선택한\s*(?:이유|계기)"),
    ("학업 태도·성적형", r"성적|등급|내신|점수|(?:좋아하|좋아했|어려웠|힘들었|자신\s*있|기억에\s*남|잘하|못하|흥미|재미있|부족)[가-힣\s]{0,8}과목|과목\s*(?:중|은|이\s*(?:있|뭐|무엇))|공부\s*(?:방법|법|를\s*어떻게|는\s*어떻게)|학습\s*(?:방법|법|태도)|선택\s*과목|이수(?:하|했|한)|수강|자기\s*주도"),
    ("서류 활동 확인형", r"세특|생기부|생활\s*기록부|학생부|자소서|자기\s*소개서|동아리|탐구|보고서|실험|프로젝트|발표|수행\s*평가|활동|참여|했다고|했는데|하였는데|했던데|했네요|하셨네요|하셨는데|한\s*것\s*같은데|되어\s*있|적혀|기재|쓰여|나와\s*있|라고\s*(?:했|되어|적|하셨)|조사(?:했|한|를)|제작(?:했|한)|설계(?:했|한)|만들었|진행(?:했|한)|수업\s*(?:시간|에서|중)|\d\s*학년|주제(?:로|를|가)|캠프|대회|자율|행특|창체|토론"),
    ("견해·시사형", r"어떻게\s*생각|(?:본인|지원자|학생|자신)의\s*(?:생각|견해|의견|입장)|생각(?:은\s*(?:무엇|어떤|어떠)|하는지|하시나|하나요|합니까|하십니까|해\s*보)|견해|찬성|반대|찬반|바람직|옳다고|옳은|해결\s*(?:방안|책|방법)|방안|문제점|대책|이슈|뉴스|기사|시사|사회\s*(?:문제|현상)|정책|제도|윤리적|논란|쟁점|전망|미래에"),
    ("전공 개념·지식형", r"무엇|뭔지|뭔가|뭐(?:죠|예요|에요|라고|인가|인지|지)|무슨|설명|차이|원리|정의|개념|의미|뜻|이란|(?<=[가-힣])란\s|종류|특징|아는|알고|아시|아세요|왜|이유|어떻게\s*(?:되|작동|구하|만들|이루|일어|발생|다른|다르)|예시|예를\s*들|공식|법칙|이론"),
]
_QTYPE_COMPILED = [(n, re.compile(p)) for n, p in QTYPE_RULES]
QTYPE_ORDER = [n for n, _ in QTYPE_RULES] + [QTYPE_ETC]
QTYPE_DESC = {
    "서류 활동 확인형": "생기부에 적힌 활동의 동기·과정·결과·배운 점을 구체적으로 확인",
    "전공 개념·지식형": "전공 관련 개념·원리·용어를 정확히 알고 있는지 검증",
    "지원동기·진로형": "지원 동기, 학과 선택 이유, 입학 후 학업 계획과 진로",
    "인성·공동체형": "갈등·협업·리더십·봉사 경험, 장단점과 극복 과정",
    "독서 확인형": "읽은 책의 내용과 자신의 생각, 활동으로 이어진 지점",
    "견해·시사형": "전공 관련 이슈·쟁점에 대한 자신의 입장과 근거",
    "학업 태도·성적형": "성적 추이, 과목 선택 이유, 공부 방법",
    "상황·딜레마형": "가상의 상황을 주고 어떻게 판단·행동할지 묻는 질문",
    QTYPE_JESIMUN: "제시문·자료·공통 문항을 읽고 분석·비교·적용해 답하는 구술",
    QTYPE_INTRO: "자기소개, 마지막으로 하고 싶은 말",
    QTYPE_ETC: "앞 질문에 이어지는 짧은 후속 질문·진행 멘트",
}
# 생기부 기반 면접 세트로 출제할 수 있는 유형 (자기소개·제시문·진행 멘트 제외)
SANGBU_QTYPES = [
    "서류 활동 확인형", "전공 개념·지식형", "지원동기·진로형", "인성·공동체형",
    "독서 확인형", "견해·시사형", "학업 태도·성적형", "상황·딜레마형",
]
_FOLLOWUP_RX = re.compile(
    r"^\s*(?:[\[(]?\s*(?:꼬리|추가|후속)\s*질문\s*[\])]?|그럼|그렇다면|그러면|그렇군요|방금|아까|앞서|추가로|한\s*번\s*더|조금\s*더|좀\s*더|더\s*자세|구체적으로)"
)

def classify_question(question):
    s = re.sub(r"\s+", " ", str(question or "")).strip()
    if len(s) < 4:
        return QTYPE_ETC
    for name, rx in _QTYPE_COMPILED:
        if rx.search(s):
            return name
    return QTYPE_ETC

# 자료집 PDF의 쪽 머리말/꼬리말이 본문에 섞여 들어온 것을 지웁니다.
_PAGE_NOISE_RX = re.compile(
    r"(?:◼+\s*[가-힣A-Za-z() ]{0,30}?|[가-힣]\s(?:[가-힣()（）]\s){4,}[가-힣)）]?\s*)?"
    r"경상북도교육청연구원\s*경북진학지원센터\s*_?\s*\d*\s*[ⅠⅡⅢ]?"
)

def _clean_exam_text(value):
    s = "" if value is None else str(value)
    if s.lower() == "nan":
        return ""
    s = _PAGE_NOISE_RX.sub(" ", s).replace("◼", " ")
    return re.sub(r"\s+", " ", s).strip()

def prepare_exam_db(df):
    """기출 DB에 분석용 열을 붙입니다.
    _uk(대학 키) · _dk(학과 키) · _field(계열) · _track(전형) · _session(한 사람의 면접으로 보이는 연속 구간)
    · _type(질문 유형) · _follow(꼬리질문 여부)"""
    df = df.reset_index(drop=True).copy()
    for col in QA_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    for col in ("대학", "학과", "전형", "질문", "답변"):
        df[col] = df[col].map(_clean_exam_text)
    df = df[(df["질문"] != "") & (df["답변"] != "")].reset_index(drop=True)

    uk_map = {u: uni_key(u) for u in df["대학"].unique()}
    dk_map = {d: dept_key(d) for d in df["학과"].unique()}
    fd_map = {d: field_of(d) for d in df["학과"].unique()}
    tk_map = {t: track_key(t) for t in df["전형"].unique()}
    df["_uk"] = df["대학"].map(uk_map)
    df["_dk"] = df["학과"].map(dk_map)
    df["_field"] = df["학과"].map(fd_map)
    df["_track"] = df["전형"].map(tk_map)

    # 파일 안에서 (대학·학과·전형)이 바뀌는 지점을 한 사람의 면접이 끝나는 곳으로 봅니다.
    changed = (df["_uk"] != df["_uk"].shift()) | (df["학과"] != df["학과"].shift()) | (df["_track"] != df["_track"].shift())
    df["_session"] = changed.cumsum()

    own = df["질문"].map(classify_question)
    # 유형을 알 수 없는 짧은 후속 질문은 바로 앞 질문의 유형을 물려받습니다(같은 면접 구간 안에서만).
    inherited = own.where(own != QTYPE_ETC).groupby(df["_session"]).ffill()
    df["_type"] = inherited.fillna(QTYPE_ETC)
    df["_follow"] = (own == QTYPE_ETC) | df["질문"].map(lambda q: bool(_FOLLOWUP_RX.search(q)))
    return df

def _ensure_prepared(df):
    if df is None or df.empty or "_uk" in df.columns:
        return df
    return prepare_exam_db(df)

def _uni_mask(df, uk):
    """대학 키 uk와 같은 학교인 행을 고릅니다."""
    if not uk or df.empty:
        return pd.Series(False, index=df.index)
    matched = {k for k in df["_uk"].unique() if uni_keys_match(uk, k)}
    return df["_uk"].isin(matched)

def list_db_universities(df):
    """기출 DB에 있는 대학을 (표시 이름, 건수)로 돌려줍니다. 건수가 많은 순."""
    df = _ensure_prepared(df)
    if df is None or df.empty:
        return []
    out = []
    for uk, sub in df.groupby("_uk"):
        if not uk:
            continue
        names = sub["대학"].map(lambda s: re.split(r"[(\[（]", s)[0].strip())
        full = names[names.str.endswith("대학교")]
        label = (full if len(full) else names).value_counts().index[0]
        label = re.sub(r"\s+", "", label) if label.endswith("대학교") else label
        base, _, campus = uk.partition("|")
        if campus:
            label = f"{label.split()[0]}({campus})"
        if uni_key(label) != uk:          # 표시 이름으로 다시 찾을 수 없으면 원래 표기를 그대로 사용
            label = sub["대학"].value_counts().index[0]
        out.append((label, len(sub)))
    return sorted(out, key=lambda x: (-x[1], x[0]))

# 참고 예시로 쓰기에 정보가 적은 유형 (다른 것이 없을 때만 사용)
_LOW_VALUE_QTYPES = {QTYPE_INTRO, QTYPE_ETC}

def _diversify_examples(hits, top_n):
    """일치도가 같은 후보 안에서는 질문 유형이 한쪽으로 쏠리지 않게 돌아가며 뽑습니다.
    (예전에는 DB에 먼저 나오는 12건을 그대로 써서 '자기소개 해보세요' 같은 질문이 섞였습니다)"""
    picked = []
    for score in sorted(hits["_score"].unique()):
        tier = hits[hits["_score"] == score]
        good = tier[~tier["_type"].isin(_LOW_VALUE_QTYPES) & (tier["질문"].str.len() >= 12) & ~tier["_follow"]]
        rest = tier.drop(index=good.index)
        groups = sorted((g for _, g in good.groupby("_type", sort=False)), key=len, reverse=True)
        queues = [list(g.index) for g in groups]
        while any(queues) and len(picked) < top_n:
            for q in queues:
                if q and len(picked) < top_n:
                    picked.append(q.pop(0))
        for idx in rest.index:
            if len(picked) >= top_n:
                break
            picked.append(idx)
        if len(picked) >= top_n:
            break
    return hits.loc[picked]

def get_relevant_examples(df, major, uni_name, top_n=12):
    """선택한 전공/대학과 가장 유사한 실제 기출 질의응답을 DB에서 골라옵니다.

    우선순위 (학과 적합성이 대학보다 우선 — 다른 대학의 같은 학과가
    같은 대학의 엉뚱한 학과보다 면접 준비에 훨씬 유용하기 때문):
      0순위: 지원 대학 + 학과 일치          1순위: 다른 대학 + 학과 일치
      2순위: 지원 대학 + 학과 부분 일치     3순위: 다른 대학 + 학과 부분 일치
      4순위: 지원 대학 + 학과 유사(철자)    5순위: 다른 대학 + 학과 유사
    부분 일치·유사는 '같은 계열'일 때만 인정합니다(물리학과 ↔ 물리치료학과 같은 오매칭 방지).
    """
    df = _ensure_prepared(df)
    if df is None or df.empty or not str(major or "").strip():
        return df.head(0) if df is not None else pd.DataFrame(columns=QA_COLUMNS)

    major = str(major).strip()
    dk, fld = dept_key(major), field_of(major)
    same_uni = _uni_mask(df, uni_key(uni_name))
    same_field = (df["_field"] == fld) if fld != "기타" else pd.Series(True, index=df.index)

    is_exact = (df["학과"] == major) | ((df["_dk"] == dk) & bool(dk))
    all_keys = [k for k in df["_dk"].unique() if k]
    partial_keys = {k for k in all_keys if k != dk and len(k) >= 2 and len(dk) >= 2 and (k in dk or dk in k)}
    is_partial = df["_dk"].isin(partial_keys) & same_field
    close = difflib.get_close_matches(dk, all_keys, n=6, cutoff=0.75) if dk else []
    is_fuzzy = df["_dk"].isin(close) & same_field

    NO_MATCH = 99
    score = pd.Series(NO_MATCH, index=df.index)
    score[is_fuzzy] = 5
    score[is_fuzzy & same_uni] = 4
    score[is_partial] = 3
    score[is_partial & same_uni] = 2
    score[is_exact] = 1
    score[is_exact & same_uni] = 0

    hits = df[score < NO_MATCH].copy()
    hits["_score"] = score[hits.index]
    hits = hits.sort_values("_score", kind="stable").drop_duplicates(subset=["_uk", "학과", "질문"])
    if not hits.empty:
        return _diversify_examples(hits, top_n)

    # 유사 학과가 하나도 없을 때: 최소한 '지원 대학의 다른 학과 기출'로 그 대학 화법이라도 반영합니다.
    same_uni_any = df[same_uni & ~df["_type"].isin(_LOW_VALUE_QTYPES)]
    if not same_uni_any.empty:
        return same_uni_any.sample(min(top_n, len(same_uni_any)), random_state=42)
    return df.sample(min(top_n, len(df)), random_state=42)

# -------------------------------------------------------------------------
# [0-3] 대학·학과별 '실제 출제 분석'
#   "이 대학 이 학과는 실제로 무엇을, 어떤 비중으로 물었는가"를 기출에서 직접 셉니다.
# -------------------------------------------------------------------------
PROFILE_MIN_ROWS = 10   # 이보다 적으면 표본이 너무 작아 한 단계 넓은 범위를 씁니다.

def _summarize_level(label, scope, sub):
    n = len(sub)
    counts = sub["_type"].value_counts()
    dist = sorted(
        [(t, int(counts[t]), int(round(counts[t] * 100 / n))) for t in counts.index],
        key=lambda x: -x[1],
    )
    tracks = []
    for tk, g in sub.groupby("_track"):
        tracks.append((tk, len(g), int(round((g["_type"] == QTYPE_JESIMUN).mean() * 100))))
    tracks.sort(key=lambda x: -x[1])
    samples = {}
    for t, _, _ in dist:
        cand = sub[(sub["_type"] == t) & ~sub["_follow"]]
        cand = cand[cand["질문"].str.len().between(18, 140)]
        if not cand.empty:
            # 여러 학생의 질문이 고르게 보이도록 면접 구간별로 하나씩
            samples[t] = cand.drop_duplicates(subset="_session")["질문"].head(3).tolist()
    return {
        "label": label,
        "scope": scope,
        "n": n,
        "sessions": int(sub["_session"].nunique()),
        "dist": dist,
        "follow_pct": int(round(sub["_follow"].mean() * 100)),
        "avg_len": int(round(sub["질문"].str.len().mean())),
        "jesimun_pct": int(round((sub["_type"] == QTYPE_JESIMUN).mean() * 100)),
        "tracks": tracks,
        "samples": samples,
    }

def build_interview_profile(df, uni_name, major):
    """지원 대학·학과의 실제 출제 경향을 범위별로 요약합니다.
    범위: ① 대학+학과 → ② 대학+같은 계열 → ③ 대학 전체 → ④ 전국 같은 학과
    표본이 PROFILE_MIN_ROWS 이상인 가장 좁은 범위를 '기준 범위(primary)'로 삼습니다."""
    df = _ensure_prepared(df)
    empty = {"levels": [], "primary": None, "uni_label": uni_name, "dept_list": [], "n_uni": 0}
    if df is None or df.empty or not str(uni_name or "").strip():
        return empty
    major = str(major or "").strip()
    dk, fld = dept_key(major), field_of(major)
    same_uni = df[_uni_mask(df, uni_key(uni_name))]
    uni_label = uni_name
    levels = []

    if dk:
        exact = same_uni[(same_uni["학과"] == major) | (same_uni["_dk"] == dk)]
        if len(exact):
            levels.append(_summarize_level(f"{uni_label} {major}", "대학+학과", exact))
        if fld != "기타":
            uni_field = same_uni[same_uni["_field"] == fld]
            if len(uni_field):
                levels.append(_summarize_level(f"{uni_label} {fld} 계열", "대학+계열", uni_field))
    if len(same_uni):
        levels.append(_summarize_level(f"{uni_label} 전체", "대학 전체", same_uni))
    if dk:
        dept_all = df[(df["_dk"] == dk) | (df["학과"] == major)]
        if len(dept_all):
            levels.append(_summarize_level(f"전국 {major}", "전국 같은 학과", dept_all))

    primary = next((lv for lv in levels if lv["n"] >= PROFILE_MIN_ROWS), None)
    dept_counts = same_uni["학과"].value_counts()
    return {
        "levels": levels,
        "primary": primary,
        "uni_label": uni_label,
        "dept_list": [(d, int(c)) for d, c in dept_counts.items() if d],
        "n_uni": len(same_uni),
    }

ALLOC_MIN_ROWS = 8      # 생기부형 질문이 이보다 적은 범위는 배분 기준으로 쓰지 않습니다.
ALLOC_SHRINK = 30       # 좁은 범위의 표본이 적을수록 넓은 범위 쪽으로 끌어당기는 정도

def _sangbu_counts(level):
    return {t: c for t, c, _ in level["dist"] if t in SANGBU_QTYPES}

def pick_allocation_basis(profile):
    """유형 배분의 기준이 될 비중을 고릅니다.
    가장 좁은 범위를 쓰되, 표본이 적으면(예: 학생 2명의 후기뿐) 한 단계 넓은 범위와 섞습니다.
    반환: (유형별 비중 dict, 기준 범위 설명) — 쓸 수 있는 표본이 없으면 ({}, "")"""
    levels = [lv for lv in (profile or {}).get("levels", [])
              if lv["n"] >= PROFILE_MIN_ROWS and sum(_sangbu_counts(lv).values()) >= ALLOC_MIN_ROWS]
    if not levels:
        return {}, ""
    first = levels[0]
    c1 = _sangbu_counts(first)
    n1 = sum(c1.values())
    share = {t: c / n1 for t, c in c1.items()}
    if len(levels) == 1 or n1 >= 60:
        return share, f"{first['label']} 기출 {first['n']}건"
    second = levels[1]
    c2 = _sangbu_counts(second)
    n2 = sum(c2.values())
    w = n1 / (n1 + ALLOC_SHRINK)
    blended = {t: w * share.get(t, 0) + (1 - w) * (c2.get(t, 0) / n2) for t in set(share) | set(c2)}
    return blended, (f"{first['label']} 기출 {first['n']}건(표본이 적어 {second['label']} {second['n']}건의 경향을 "
                     f"{int(round((1 - w) * 100))}% 섞음)")

def allocate_question_types(profile, n_slots=5):
    """기출 유형 비중에 맞춰 n_slots개 세트의 유형을 배분합니다(최대잉여법).
    생기부 기반 면접이므로 '서류 활동 확인형'은 최소 1세트를 보장합니다.
    반환: ([(유형, 세트 수)], 기준 설명)"""
    share, basis = pick_allocation_basis(profile)
    if not share:
        return [], ""
    quota = {t: p * n_slots for t, p in share.items()}
    alloc = {t: int(q) for t, q in quota.items()}
    remain = n_slots - sum(alloc.values())
    for t in sorted(quota, key=lambda t: (quota[t] - alloc[t], quota[t]), reverse=True)[:remain]:
        alloc[t] += 1
    base = "서류 활동 확인형"
    if alloc.get(base, 0) == 0:
        donor = max(alloc, key=lambda t: alloc[t])
        if alloc[donor] > 1:
            alloc[donor] -= 1
            alloc[base] = 1
    return sorted([(t, k) for t, k in alloc.items() if k > 0], key=lambda x: -x[1]), basis

def format_profile_for_prompt(profile, n_slots=5):
    """출제 분석 결과를 AI에게 줄 지시문으로 만듭니다(생기부 기반 면접용)."""
    level = (profile or {}).get("primary")
    if not level:
        return ""
    alloc, basis = allocate_question_types(profile, n_slots)
    mix = " · ".join(f"{t} {pct}%" for t, _, pct in level["dist"][:6])
    intro_pct = next((pct for t, _, pct in level["dist"] if t == QTYPE_INTRO), 0)
    lines = [
        f"\n[C. 실제 출제 유형 분석 — {level['label']} 기출 {level['n']}건 (수험생 복기 자료, 범위: {level['scope']})]",
        f"- 실제로 나온 질문 유형 비중: {mix}",
        f"- 꼬리질문 비중: 약 {level['follow_pct']}% / 평균 질문 길이: 약 {level['avg_len']}자 "
        f"(실제 면접관처럼 짧고 구어체로 물으세요. 한 질문에 여러 요구를 길게 나열하지 마세요.)",
    ]
    if level["follow_pct"] >= 30:
        lines.append("- 이 대학은 꼬리질문이 많습니다. [꼬리질문]은 '1차 꼬리질문 → 학생 답변을 가정한 2차 재질문'의 2단계로 쓰세요.")
    if intro_pct >= 5:
        lines.append(f"- 자기소개·마지막 한마디도 약 {intro_pct}% 나오지만, 이것은 세트로 만들지 마세요(별도 준비 사항).")
    if level["jesimun_pct"] >= 15:
        lines.append(f"- ⚠️ 이 범위는 제시문·구술 문제 비중이 {level['jesimun_pct']}%입니다. 전형에 따라 제시문 면접일 수 있습니다.")
    if alloc:
        sample_levels = [lv for lv in profile["levels"] if lv["n"] >= PROFILE_MIN_ROWS]
        lines.append(f"\n[{n_slots}세트 유형 배분 — 기준: {basis}. 반드시 지키세요]")
        for t, k in alloc:
            lines.append(f"- {t} {k}세트: {QTYPE_DESC.get(t, '')}")
        lines.append("- 각 세트의 [평가의도] 첫 문장에 '(출제 유형: ○○형)'을 적으세요.")
        lines.append("\n[유형별 실제 기출 문장 — 말투와 길이의 기준]")
        for t, _ in alloc:
            qs = next((lv["samples"][t] for lv in sample_levels if lv["samples"].get(t)), [])
            for q in qs[:2]:
                lines.append(f"- [{t}] {q[:150]}")
    return "\n".join(lines)

def analyze_university_style(df, uni_name, min_rows=5):
    """지원 대학 기출 전체를 훑어 '이 대학은 어떤 식으로 묻는가'를 수치로 요약합니다."""
    df = _ensure_prepared(df)
    if df is None or df.empty or not uni_name:
        return None
    same = df[_uni_mask(df, uni_key(uni_name))]
    if len(same) < min_rows:
        return None
    ratio = (same["_type"].value_counts(normalize=True) * 100).round().astype(int).to_dict()
    visible = {t: r for t, r in ratio.items() if t != QTYPE_ETC}
    top_types = sorted(visible.items(), key=lambda kv: kv[1], reverse=True)[:3]
    return {
        "대학": uni_name,
        "기출수": len(same),
        "학과수": same["학과"].nunique(),
        "평균질문길이": int(round(same["질문"].str.len().mean())),
        "유형비율": ratio,
        "대표유형": top_types,
    }

def get_university_style_examples(df, uni_name, major, n=8, exclude_index=None):
    """지원 대학의 출제 스타일을 보여줄 기출 (학과는 달라도 됨).
    지원 학과와 같은 계열을 우선하고, 한 학과에 쏠리지 않게 골고루 뽑습니다."""
    df = _ensure_prepared(df)
    if df is None or df.empty or not uni_name:
        return df.head(0) if df is not None else pd.DataFrame(columns=QA_COLUMNS)
    same = df[_uni_mask(df, uni_key(uni_name))]
    if exclude_index is not None and len(same):
        same = same.drop(index=[i for i in exclude_index if i in same.index], errors="ignore")
    same = same[~same["_type"].isin(_LOW_VALUE_QTYPES) & (same["질문"].str.len() >= 12)]
    if same.empty:
        return same

    target_field = field_of(major)
    same = same.copy()
    same["_field_rank"] = (same["_field"] != target_field).astype(int)  # 같은 계열이 0

    picked, per_dept = [], {}
    for _, row in same.sort_values("_field_rank", kind="stable").iterrows():
        dept = str(row["학과"])
        if per_dept.get(dept, 0) >= 2:      # 한 학과에서 최대 2개
            continue
        per_dept[dept] = per_dept.get(dept, 0) + 1
        picked.append(row)
        if len(picked) >= n:
            break
    return pd.DataFrame(picked) if picked else same.head(0)

# -------------------------------------------------------------------------
# [0-4] 제시문 면접: '실제 기출 사례'를 먼저 보여주고 응용 문제를 붙입니다
#   AI가 기출을 지어내지 않도록, 기출 원문은 DB에서 코드가 직접 골라 그대로 끼워 넣습니다.
# -------------------------------------------------------------------------
REAL_CASE_TAG = "[실제 기출]"
_CASE_TIER_LABEL = {
    0: "지원 대학·학과 기출",
    1: "지원 대학 같은 계열 기출",
    2: "다른 대학 같은 학과 기출",
    3: "다른 대학 같은 계열 기출",
    4: "지원 대학 다른 계열 기출",
    5: "다른 계열 기출",
}
_EMPTY_ANSWER_RX = re.compile(r"기억|생략|중략|대답\s*못|못\s*함|모르겠|^\W*$")

def _safe_case_text(text):
    """기출 원문 속 기호가 문항 파싱·화면 표시를 깨뜨리지 않도록 바꿉니다."""
    s = str(text or "")
    s = s.replace("[", "〔").replace("]", "〕").replace("###", "").replace("**", "")
    s = s.replace("~", "∼").replace("$", "＄")
    return re.sub(r"\s+", " ", s).strip()

def _char_trigrams(text):
    s = re.sub(r"[^가-힣A-Za-z0-9]", "", str(text))
    return {s[i:i + 3] for i in range(max(len(s) - 2, 0))}

def select_real_cases(df, uni_name, major, n=3, max_rows=4):
    """제시문·구술 기출을 '사례' 단위로 묶어, 지원 대학·학과에 가까운 순으로 n건 고릅니다.
    사례 = 한 사람의 면접에서 연달아 나온 제시문 문항들(문제 1, 문제 2 …)."""
    df = _ensure_prepared(df)
    if df is None or df.empty:
        return []
    pool = df[df["_type"] == QTYPE_JESIMUN]
    if pool.empty:
        return []
    major = str(major or "").strip()
    dk, fld = dept_key(major), field_of(major)
    uk = uni_key(uni_name)
    uni_keys = {k for k in df["_uk"].unique() if uni_keys_match(uk, k)} if uk else set()

    # 같은 면접 구간 안에서 '연달아 나온' 제시문 행을 하나의 사례로 묶습니다.
    pos = pd.Series(range(len(df)), index=df.index)[pool.index]
    run_id = ((pos.diff() != 1) | (pool["_session"] != pool["_session"].shift())).cumsum()

    cases = []
    for _, g in pool.groupby(run_id):
        g = g.head(max_rows)
        first = g.iloc[0]
        q_len = int(g["질문"].str.len().sum())
        if q_len < 30:
            continue
        answered = int((~g["답변"].map(lambda a: bool(_EMPTY_ANSWER_RX.search(a)) and len(a) < 40)).sum())
        same_uni = first["_uk"] in uni_keys
        d = first["_dk"]
        same_dept = bool(dk) and bool(d) and (d == dk or (len(d) >= 2 and len(dk) >= 2 and (d in dk or dk in d)))
        same_field = fld != "기타" and first["_field"] == fld
        if same_uni and same_dept:
            tier = 0
        elif same_uni and same_field:
            tier = 1
        elif same_dept and (same_field or fld == "기타"):
            tier = 2
        elif same_field:
            tier = 3
        elif same_uni:
            tier = 4
        else:
            tier = 5
        # 복기 내용이 거의 없는 사례는 제외 ("1번 문제는 ~ 입니다" 처럼 내용이 빠진 후기)
        q_text = " ".join(g["질문"])
        a_len = int(g["답변"].str.len().sum())
        placeholders = q_text.count("~") + q_text.count("∼")
        if placeholders >= 2 or (int(g["질문"].str.len().max()) < 40 and a_len < 200):
            continue
        cases.append({
            "tier": tier,
            "quality": min(q_len, 600) + min(a_len, 400) // 2 + 60 * answered,
            "대학": first["대학"], "학과": first["학과"], "전형": first["전형"], "계열": first["_field"],
            "rows": list(zip(g["질문"].tolist(), g["답변"].tolist())),
            "index": list(g.index),
        })

    cases.sort(key=lambda c: (c["tier"], -c["quality"]))
    picked, seen = [], []
    for c in cases:
        grams = _char_trigrams(" ".join(q for q, _ in c["rows"]))
        if not grams:
            continue
        # 같은 제시문을 여러 수험생이 복기한 경우가 많아, 내용이 겹치면 건너뜁니다.
        if any(len(grams & s) / max(min(len(grams), len(s)), 1) > 0.35 for s in seen):
            continue
        seen.append(grams)
        picked.append(c)
        if len(picked) >= n:
            break
    return picked

def format_real_case_block(case, uni_name="", major=""):
    """사례 하나를 문항 세트에 끼워 넣을 글로 만듭니다(대괄호 등은 안전한 기호로 바꿈)."""
    tier_label = _CASE_TIER_LABEL.get(case["tier"], "")
    lines = [
        f"▣ 출처: {_safe_case_text(case['대학'])} · {_safe_case_text(case['학과'])} · {_safe_case_text(case['전형'])} "
        f"— 수험생 면접 후기(복기) / {tier_label}"
    ]
    if case["tier"] >= 2 and uni_name:
        lines.append(f"※ {_safe_case_text(uni_name)} {_safe_case_text(major)}의 제시문 기출이 DB에 없어, 가장 가까운 기출을 대신 제시합니다.")
    lines.append("▣ 실제로 나온 제시문·문제")
    for i, (q, _) in enumerate(case["rows"], 1):
        lines.append(f"  ({i}) {_safe_case_text(q)[:700]}")
    answers = [(i, a) for i, (_, a) in enumerate(case["rows"], 1)
               if not (_EMPTY_ANSWER_RX.search(a) and len(a) < 40)]
    if answers:
        lines.append("▣ 수험생이 실제로 한 답변(요지)")
        for i, a in answers:
            lines.append(f"  ({i}) {_safe_case_text(a)[:350]}")
    return "\n".join(lines)

def format_real_cases_for_prompt(blocks, uni_name, major, n_sets=3):
    """AI에게 '세트 i는 기출 i를 응용하라'고 지시하는 글."""
    if not blocks:
        return (
            f"\n[D. 제시문 기출]\n- DB에 참고할 제시문 기출이 없습니다. {major}의 핵심 쟁점으로 {n_sets}세트를 직접 창작하세요.\n"
        )
    lines = [
        f"\n[D. 실제 제시문 면접 기출 {len(blocks)}건 — 세트별 응용 대상]",
        "아래는 수험생이 복기한 실제 기출입니다. 각 세트는 같은 번호의 기출을 '응용'해서 만드세요.",
    ]
    for i, b in enumerate(blocks, 1):
        lines.append(f"\n<기출 {i} → 세트 {i}의 응용 대상>\n{b}")
    lines.append(
        "\n[응용 규칙 — 반드시 지키세요]\n"
        "1. 기출의 **문제 구조와 사고 과정**(예: 여러 제시문의 공통점·차이점 비교 → 자료 해석 → 다른 사례에 적용 → 자기 견해)은 그대로 살리세요.\n"
        f"2. **소재와 제시문 내용은 새로** 쓰세요. 기출 문장을 베끼지 말고, {major} 지원자에게 맞는 다른 주제·사례·자료로 바꿉니다.\n"
        "3. [제시문]의 첫 줄에는 반드시 '※ 기출 응용 포인트: (기출의 어떤 구조를 살리고 무엇을 바꿨는지 한 문장)'을 적으세요.\n"
        "4. 기출 원문은 출력하지 마세요. 시스템이 각 세트 맨 앞에 자동으로 넣습니다. '[실제 기출]'이라는 표제도 쓰지 마세요.\n"
        "5. 난이도는 기출과 같거나 한 단계 높게 맞추세요."
    )
    if len(blocks) < n_sets:
        lines.append(f"6. 기출이 없는 세트 {len(blocks) + 1}~{n_sets}는 위 기출들의 출제 방식을 본떠 {major}의 다른 쟁점으로 창작하세요.")
    return "\n".join(lines)

_REAL_CASE_RX = re.compile(r"\[실제 기출\].*?(?=\[제시문\]|\[문제 1\]|###\s*📌|\Z)", re.DOTALL)

def strip_real_cases(result_text):
    """문항 원문에서 '[실제 기출]' 블록을 떼어냅니다(AI에게 다시 보낼 때 사용)."""
    return _REAL_CASE_RX.sub("", str(result_text or ""))

def extract_real_cases(result_text):
    """세트 순서대로 '[실제 기출]' 블록 본문을 꺼냅니다. 없는 세트는 None."""
    out = []
    for block in str(result_text or "").split("### 📌")[1:]:
        m = _REAL_CASE_RX.search(block)
        out.append(m.group(0)[len(REAL_CASE_TAG):].strip() if m else None)
    return out

def inject_real_cases(result_text, blocks):
    """AI가 만든 세트 i의 [제시문] 바로 앞에 기출 i를 끼워 넣습니다."""
    text = strip_real_cases(result_text)
    if not blocks or not any(blocks):
        return text
    parts = text.split("### 📌")
    for i in range(1, len(parts)):
        block = blocks[i - 1] if i - 1 < len(blocks) else None
        if not block:
            continue
        insert = f"{REAL_CASE_TAG}\n{block}\n"
        part = parts[i]
        at = part.find("[제시문]")
        if at < 0:
            at = part.find("[문제 1]")
        if at < 0:
            nl = part.find("\n")
            parts[i] = part + "\n" + insert if nl < 0 else part[:nl + 1] + insert + part[nl + 1:]
        else:
            parts[i] = part[:at] + insert + part[at:]
    return "### 📌".join(parts)

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


def format_university_style_for_prompt(style, style_examples, has_same_dept_at_uni):
    """대학 스타일 분석 + 그 대학 기출 예시를 프롬프트 블록으로 만듭니다."""
    if not style and (style_examples is None or style_examples.empty):
        return ""
    lines = [f"\n[A. 지원 대학의 출제 스타일 분석 — {style['대학'] if style else ''}]"]
    if style:
        types = ", ".join(f"{label} {ratio}%" for label, ratio in style["대표유형"])
        lines.append(f"- 분석 표본: 이 대학 실제 기출 {style['기출수']}건 ({style['학과수']}개 학과)")
        lines.append(f"- 평균 질문 길이: 약 {style['평균질문길이']}자")
        lines.append(f"- 이 대학이 가장 자주 쓰는 질문 유형: {types}")
    if style_examples is not None and not style_examples.empty:
        if has_same_dept_at_uni:
            lines.append("- 이 대학의 실제 질문 화법 예시 (지원 학과 기출은 아래 B에 별도로 있음):")
        else:
            lines.append("- ⚠️ 이 대학의 '지원 학과' 기출은 DB에 없습니다. 아래는 같은 대학 다른 학과 기출이며, "
                          "**질문하는 방식·말투·난이도·꼬리질문 습관만** 참고 대상입니다 (전공 내용은 참고하지 마세요):")
        for _, r in style_examples.iterrows():
            q = re.sub(r"\s+", " ", str(r["질문"])).strip()
            lines.append(f"  · [{r['학과']}] {q[:160]}")
    return "\n".join(lines)

def get_univ_info(info_df, uni_name, major):
    """지원 대학의 실제 면접 형식 정보를 찾습니다. 같은 학과가 있으면 그것을 우선합니다."""
    if info_df is None or info_df.empty or not uni_name:
        return None
    uni_key = _normalize_uni(uni_name)
    if not uni_key:
        return None
    same_uni = info_df[info_df["대학"].apply(lambda u: _uni_matches(uni_name, str(u)))]
    if same_uni.empty:
        return None
    major_norm = _normalize_dept(major or "")
    if major_norm:
        same_dept = same_uni[same_uni["학과"].apply(
            lambda d: bool(_normalize_dept(str(d))) and
            (_normalize_dept(str(d)) in major_norm or major_norm in _normalize_dept(str(d)))
        )]
        if not same_dept.empty:
            return same_dept.iloc[0]
    return same_uni.iloc[0]

def format_univ_info_for_prompt(info_row):
    """면접 형식 정보를 프롬프트에 넣을 문장으로 만듭니다."""
    if info_row is None:
        return ""
    def _v(key):
        val = info_row.get(key, "")
        return "" if (val is None or str(val).strip().lower() in ("", "nan")) else str(val).strip()

    parts = [f"\n[지원 대학의 실제 면접 형식 (2026학년도 면접 후기 기준: {_v('대학')} {_v('학과')} / {_v('전형')})]"]
    if _v("면접유형"):
        parts.append(f"- 면접 유형: {_v('면접유형')}")
    if _v("면접시간"):
        parts.append(f"- 면접 시간: {_v('면접시간')} (이 시간 안에 소화 가능한 분량으로 문항을 설계하세요)")
    if _v("면접위원"):
        parts.append(f"- 면접 위원: {_v('면접위원')}")
    if _v("유의사항"):
        parts.append(f"- 실제 면접 특징: {_v('유의사항')[:300]}")
    if _v("선배조언"):
        parts.append(f"- 합격 선배의 조언(출제 경향 참고): {_v('선배조언')[:300]}")
    parts.append("※ 위 형식을 반드시 반영해 문항 수와 깊이를 정하세요. 예를 들어 10분 면접이면 지나치게 많은 세부 질문을 넣지 마세요.")
    return "\n".join(parts)

def format_examples_for_prompt(examples_df, uni_name=""):
    if examples_df.empty:
        return ""
    lines = ["\n[B. 지원 학과 전공 기출 — 전공적합성과 학술적 깊이의 기준]"]
    lines.append("(질문의 전공 깊이·검증 수준을 여기서 가져오되, 문장을 그대로 베끼지 말고 학생 생기부 내용으로 새로 창작하세요.)")
    for _, r in examples_df.iterrows():
        q = str(r["질문"]).strip().replace("\n", " ")
        a = str(r["답변"]).strip().replace("\n", " ")
        if len(a) > 220:
            a = a[:220] + "..."
        same_mark = " ★지원대학" if (uni_name and _uni_matches(uni_name, str(r["대학"]))) else ""
        lines.append(f"- [{r['대학']} · {r['학과']}{same_mark}] Q: {q}\n  A: {a}")
    return "\n".join(lines)

def build_reference_strategy(exam_db, uni_name, major, top_n=12, style_n=8, include_profile=True):
    """면접 문항 생성에 쓸 참고자료를 '대학 스타일'과 '학과 전공' 두 갈래로 준비합니다.

    - 지원 대학 + 지원 학과 기출이 있으면 → 그것을 그대로 주력으로 사용 (mode='exact')
    - 지원 학과 기출은 있는데 그 대학 것이 없으면 → 다른 대학의 같은 학과 기출(전공)
      + 지원 대학의 다른 학과 기출(스타일)을 결합 (mode='style_transfer')
    - 학과 기출 자체가 없으면 → 지원 대학 스타일 위주 (mode='uni_only')
    """
    examples = get_relevant_examples(exam_db, major, uni_name, top_n=top_n)
    has_same_dept_at_uni = bool(len(examples)) and any(
        _uni_matches(uni_name, str(u)) for u in examples["대학"]
    )
    style = analyze_university_style(exam_db, uni_name)
    style_examples = get_university_style_examples(
        exam_db, uni_name, major, n=style_n,
        exclude_index=list(examples.index) if len(examples) else None,
    )

    if has_same_dept_at_uni:
        mode = "exact"
    elif len(examples):
        mode = "style_transfer" if (style or len(style_examples)) else "dept_only"
    else:
        mode = "uni_only" if (style or len(style_examples)) else "none"

    # 프롬프트 지시문: 모드별로 AI가 해야 할 일을 명확히 지정
    if mode == "exact":
        guide = (
            f"\n[참고자료 사용 지침]\n"
            f"- 아래 B에는 **{uni_name} {major}의 실제 기출**이 포함되어 있습니다(★ 표시). "
            f"이 문항들의 화법·난이도·꼬리질문 방식을 최우선 기준으로 삼으세요.\n"
            f"- 단, 문장을 그대로 쓰지 말고 **이 학생의 생기부에 실제로 적힌 활동·개념으로 바꿔 재창작**하세요.\n"
        )
    elif mode == "style_transfer":
        guide = (
            f"\n[참고자료 사용 지침 — ⚠️ 중요]\n"
            f"- DB에 **{uni_name} {major} 기출은 없습니다.** 그래서 두 가지를 결합해야 합니다.\n"
            f"- ① A(대학 스타일): {uni_name}의 실제 출제 방식 — **어떤 말투·길이·유형으로 묻는지**를 그대로 따르세요. "
            f"질문의 형식과 압박 수위는 A를 기준으로 합니다.\n"
            f"- ② B(학과 전공): 다른 대학의 {major} 기출에서 **전공적합성과 학술적 깊이**를 가져오세요. "
            f"어떤 개념을 어느 수준까지 파고드는지는 B를 기준으로 합니다.\n"
            f"- 즉 **'{uni_name}의 말투로 묻는 {major} 전공 질문'**을 만들어야 합니다. "
            f"A의 다른 학과 전공 내용은 절대 가져오지 마세요.\n"
        )
    elif mode == "uni_only":
        guide = (
            f"\n[참고자료 사용 지침 — ⚠️ 중요]\n"
            f"- DB에 {major}와 유사한 학과 기출이 없습니다. A({uni_name}의 출제 스타일)만 참고할 수 있습니다.\n"
            f"- {uni_name}의 질문 방식은 A를 따르고, **전공 내용은 {major}의 고교 교육과정 연계 개념에서 직접 설계**하세요.\n"
        )
    else:
        guide = "\n[참고자료 사용 지침]\n- 참고 기출이 부족하니 학생 생기부 내용과 전공 개념에 근거해 직접 설계하세요.\n"

    style_text = format_university_style_for_prompt(style, style_examples, has_same_dept_at_uni)
    dept_text = format_examples_for_prompt(examples, uni_name=uni_name)
    combined = guide + style_text + "\n" + dept_text

    # C: 이 대학·학과가 실제로 어떤 유형을 어떤 비중으로 물었는지 → 세트별 유형 배분 지시
    profile = build_interview_profile(exam_db, uni_name, major)
    profile_text = format_profile_for_prompt(profile) if include_profile else ""
    if profile_text:
        combined += "\n" + profile_text

    return {
        "mode": mode,
        "examples": examples,
        "style": style,
        "style_examples": style_examples,
        "has_same_dept_at_uni": has_same_dept_at_uni,
        "prompt_text": combined,
        "profile": profile,
        "alloc": allocate_question_types(profile) if include_profile else ([], ""),
    }

# -------------------------------------------------------------------------
# [0-5] 출제 분석 화면 (대학·학과를 고르면 문항 생성 전에 바로 볼 수 있음)
# -------------------------------------------------------------------------
DB_UNI_REGION = "📚 기출 DB 대학 전체"

@st.cache_data(show_spinner=False)
def cached_db_universities(folder):
    df, _ = load_exam_db(folder)
    return list_db_universities(df)

@st.cache_data(show_spinner=False)
def cached_interview_profile(folder, uni_name, major):
    df, _ = load_exam_db(folder)
    return build_interview_profile(df, uni_name, major)

@st.cache_data(show_spinner=False)
def cached_real_cases(folder, uni_name, major, n=3):
    df, _ = load_exam_db(folder)
    return select_real_cases(df, uni_name, major, n=n)

def _md_safe(text):
    """기출 원문을 화면에 그대로 보여줄 때 마크다운 기호가 서식으로 해석되지 않게 합니다."""
    return str(text or "").replace("~", "∼").replace("$", "＄").replace("*", "＊")

def render_interview_profile(profile, uni_name, major, interview_type, real_cases=None):
    levels = (profile or {}).get("levels") or []
    primary = (profile or {}).get("primary")
    is_sangbu = "생기부" in interview_type
    if not levels:
        st.caption(f"📊 기출 DB에 '{uni_name}' 또는 '{major}' 관련 기출이 없어 출제 분석을 표시할 수 없습니다.")
        return
    head = f"{primary['label']} 기출 {primary['n']}건 기준" if primary else "표본이 적어 참고만 가능"
    with st.expander(f"📊 {uni_name} {major} 실제 출제 분석 — {head}", expanded=False):
        st.caption(
            "수험생이 복기한 면접 후기를 키워드 규칙으로 분류한 근사치입니다. "
            "건수가 적은 범위는 '경향'으로만 참고하세요."
        )
        if is_sangbu and primary and primary["jesimun_pct"] >= 40:
            st.warning(
                f"⚠️ {primary['label']} 기출의 {primary['jesimun_pct']}%가 제시문·구술 문제입니다. "
                "이 대학·전형은 **'상위권 대학 제시문 기반 면접'** 방식으로 준비하는 편이 맞습니다."
            )
        alloc, basis = allocate_question_types(profile)
        if is_sangbu and alloc:
            st.info(
                "**이번에 만들 5세트의 유형 배분** — " + " · ".join(f"{t} {k}세트" for t, k in alloc)
                + f"\n\n기준: {basis}"
            )
        if not is_sangbu:
            if real_cases:
                st.info(
                    "**먼저 제시할 실제 기출 사례** (각 세트는 이 기출을 응용해 만듭니다)\n\n"
                    + "\n".join(
                        f"{i}. {c['대학']} · {c['학과']} · {c['전형']} — {_CASE_TIER_LABEL.get(c['tier'], '')}"
                        for i, c in enumerate(real_cases, 1)
                    )
                )
            else:
                st.warning("제시문·구술 기출이 DB에 없어 실제 사례 없이 창작 문제만 만듭니다.")

        tabs = st.tabs([f"{lv['scope']} · {lv['n']}건" for lv in levels])
        for tab, lv in zip(tabs, levels):
            with tab:
                st.markdown(
                    f"**{lv['label']}** — 꼬리질문 약 {lv['follow_pct']}% · 평균 질문 길이 약 {lv['avg_len']}자"
                )
                st.dataframe(
                    pd.DataFrame([{"질문 유형": t, "건수": c, "비중": pct} for t, c, pct in lv["dist"]]),
                    hide_index=True, use_container_width=True,
                    column_config={"비중": st.column_config.ProgressColumn("비중", min_value=0, max_value=100, format="%d%%")},
                )
                if lv["tracks"]:
                    st.caption("전형별 건수: " + " · ".join(
                        f"{tk} {n}건" + (f"(제시문·구술 {jp}%)" if jp >= 10 else "")
                        for tk, n, jp in lv["tracks"][:6]
                    ))
                for t, _, _ in lv["dist"][:3]:
                    qs = lv["samples"].get(t)
                    if qs:
                        st.markdown(f"**{t} 실제 질문**\n" + "\n".join(f"- {_md_safe(q)}" for q in qs))
        depts = profile.get("dept_list") or []
        if depts:
            shown = ", ".join(f"{d}({c})" for d, c in depts[:30])
            more = f" 외 {len(depts) - 30}개" if len(depts) > 30 else ""
            st.caption(f"이 대학 기출 보유 학과({len(depts)}개): {shown}{more}")


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

def logout():
    out_dir = st.session_state.get("_out_dir")
    if out_dir:
        shutil.rmtree(out_dir, ignore_errors=True)
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

def require_login():
    """로그인이 켜져 있는데 아직 로그인하지 않았다면 로그인 화면만 보여주고 멈춥니다."""
    accounts, _ = _auth_config()
    if not accounts:
        return
    if st.session_state.get("auth_teacher") in accounts:
        return
    st.session_state.pop("auth_teacher", None)

    st.title("🎓 도개고 면접 마스터")
    st.caption("교내 시범운영 중입니다. 배정받은 아이디와 비밀번호로 로그인해 주세요.")

    locked_until = st.session_state.get("_login_locked_until", 0)
    if time.time() < locked_until:
        st.error(f"로그인 시도가 너무 많습니다. {int(locked_until - time.time()) + 1}초 뒤에 다시 시도해 주세요.")
        st.stop()

    with st.form("login_form"):
        teacher_id = st.text_input("아이디")
        password = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("로그인", use_container_width=True)

    if submitted:
        teacher_id = (teacher_id or "").strip()
        expected = accounts.get(teacher_id)
        ok = expected is not None and hmac.compare_digest(
            (password or "").encode("utf-8"), expected.encode("utf-8")
        )
        if ok:
            st.session_state["auth_teacher"] = teacher_id
            st.session_state["_login_fails"] = 0
            st.rerun()
        fails = st.session_state.get("_login_fails", 0) + 1
        st.session_state["_login_fails"] = fails
        if fails >= 5:
            st.session_state["_login_locked_until"] = time.time() + 60
            st.session_state["_login_fails"] = 0
        time.sleep(1)
        st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
    st.stop()

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

def _safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(name)).strip(" .")
    stem, ext = os.path.splitext(name)
    return (stem[:60] or "문서") + ext[:10]

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

def dump_chat_history(chat_history, limit=SHEET_CELL_LIMIT):
    """대화 내역을 JSON 문자열로 만들되, 구글시트 셀 상한을 넘지 않게
    오래된 메시지부터 잘라냅니다. (중간에서 잘린 JSON은 다시 못 읽기 때문)"""
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

def _list_records_raw():
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
    return [r for r in _list_records_raw() if can_access_record(r["id"])]

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
        st.session_state.get("chat_history", []),
        quiet=quiet,
    )

def render_save_status(prefix=""):
    """마지막 저장 시각과 저장 위치를 한 줄로 보여줍니다."""
    last = st.session_state.get("last_save_info")
    if not last:
        st.caption(f"{prefix}아직 이번 세션에서 저장된 기록이 없습니다.")
        return
    if last["ok"] and "구글시트" in last["where"]:
        st.caption(f"{prefix}☁️ 마지막 저장: {last['at']} · {last['where']}")
    elif last["ok"]:
        st.caption(f"{prefix}💾 마지막 저장: {last['at']} · {last['where']} (재배포 시 사라질 수 있음)")
    else:
        st.caption(f"{prefix}⚠️ 마지막 저장 실패: {last['at']} · 로컬에만 임시 저장됨")

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

    file_path = _out_path(f"{student_name}_면접직전_요약카드.docx")
    doc.save(file_path)
    return file_path

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

def _add_write_box(doc, title, height_inches=0.9, fill="F7F9FC", spacer=True):
    """제목 줄 + 손으로 적을 수 있는 빈 칸 한 세트를 추가합니다."""
    table = doc.add_table(rows=2, cols=1)
    table.style = 'Table Grid'
    _set_table_fixed_layout(table, [Inches(6.1)])
    set_cell_background(table.rows[0].cells[0], fill)
    _set_cell_text(table.rows[0].cells[0], title, bold=True, size=10, color=RGBColor(0, 51, 102))
    _set_cell_text(table.rows[1].cells[0], "", size=10)
    table.rows[1].height = Inches(height_inches)
    table.rows[1].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    if spacer:
        _small_spacer(doc)
    return table

def create_evaluation_sheet_word(result_text, student_name, university, major, interview_type,
                                  evaluator="", interview_date="", include_questions=True):
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
    if pairs:
        q_head = doc.add_paragraph()
        q_head_run = q_head.add_run("▣ 문항별 답변 평가 (각 10점)")
        q_head_run.bold = True
        q_head_run.font.size = Pt(10)
        q_head_run.font.color.rgb = RGBColor(0, 51, 102)

        pairs = pairs[:6]
        q_table = doc.add_table(rows=1 + len(pairs), cols=4)
        q_table.style = 'Table Grid'
        for c, head in enumerate(["번호", "면접 질문", "점수", "특이사항 · 메모"]):
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
            _set_cell_text(row.cells[2], "", size=9.5, align="center")
            _set_cell_text(row.cells[3], "", size=9)
            row.height = Inches(0.55)
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        _small_spacer(doc, size=8)

    # ── 종합 의견란 ──
    opinion_head = doc.add_paragraph()
    opinion_run = opinion_head.add_run("▣ 종합 의견")
    opinion_run.bold = True
    opinion_run.font.size = Pt(10)
    opinion_run.font.color.rgb = RGBColor(0, 51, 102)

    _add_write_box(doc, "① 잘한 점 (강점) — 실제 답변 내용을 근거로 구체적으로 기재", height_inches=0.95)
    _add_write_box(doc, "② 보완할 점 (약점) — 내용·태도 측면 모두 기재", height_inches=0.95)
    _add_write_box(doc, "③ 다음 연습까지의 지도 방향 및 과제", height_inches=0.9, spacer=False)
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
    file_path = _out_path(f"{student_name}_성장리포트.docx")
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
if "current_record_key" not in st.session_state: st.session_state.current_record_key = None
if "record_key_map" not in st.session_state: st.session_state.record_key_map = {}
if "loaded_student_record_text" not in st.session_state: st.session_state.loaded_student_record_text = ""
if "easy_explanation_text" not in st.session_state: st.session_state.easy_explanation_text = ""
if "easy_explanation_file" not in st.session_state: st.session_state.easy_explanation_file = None
if "summary_card_file" not in st.session_state: st.session_state.summary_card_file = None
if "growth_report_file" not in st.session_state: st.session_state.growth_report_file = None
if "evaluation_sheet_file" not in st.session_state: st.session_state.evaluation_sheet_file = None
if "last_save_info" not in st.session_state: st.session_state.last_save_info = None

require_login()

exam_db, exam_db_files = load_exam_db(APP_DIR)
univ_info_db, univ_info_files = load_univ_info(APP_DIR)

with st.sidebar:
    _teacher = current_teacher()
    if _teacher:
        st.markdown(f"👤 **{_teacher}** 선생님" + (" · 관리자" if is_admin() else ""))
        if st.button("로그아웃", use_container_width=True, key="logout_button"):
            logout()
        st.divider()
    else:
        st.warning("🔓 로그인이 설정되지 않았습니다. 주소를 아는 누구나 저장된 학생 기록을 볼 수 있습니다.")

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
        st.success(f"📊 실제 기출 DB 연동됨\n{len(exam_db):,}건 / {exam_db['_uk'].nunique()}개 대학")
        if len(exam_db_files) > 1:
            st.caption("📎 읽은 파일 " + ", ".join(exam_db_files) + " (중복은 자동 제거)")

    if univ_info_db.empty:
        st.info("📋 대학별 면접 형식 정보(univ_interview_info.csv)가 없습니다.\n파일을 올리면 대학별 면접 시간·위원 수까지 반영됩니다.")
    else:
        st.success(f"📋 2026 면접 형식 정보\n{univ_info_db['대학'].nunique()}개 대학 / {len(univ_info_db)}개 학과")

    if _appsscript_enabled():
        st.success("☁️ 학생 기록: Google Sheets(Apps Script)에 반영구 저장 중")
    elif _sheets_enabled():
        st.success("☁️ 학생 기록: Google Sheets(서비스 계정)에 반영구 저장 중")
    else:
        st.info("💾 학생 기록: 로컬 임시 저장 중 (재배포 시 초기화될 수 있음)")

    render_save_status()

    # 구글시트 연결이 실제로 살아있는지 직접 확인 (토큰이 어긋나면 저장이 조용히 실패할 수 있음)
    if _appsscript_enabled():
        if st.button("🔌 구글시트 연결 테스트", use_container_width=True, key="conn_test_btn"):
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
        _show_owner = is_admin()
        record_options = {
            (f"[{record_owner(r['id']) or '예전 기록'}] " if _show_owner else "")
            + f"{r['student_name']} · {r['university']}_{r['major']} — "
            + f"{str(r['updated_at'])[:16].replace('T', ' ')}": r["id"]
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
                    st.session_state["chat_history"] = load_json_safe(_rec_get(row, "chat_history", ""), [])
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
                    st.session_state["summary_card_file"] = None
                    st.session_state["growth_report_file"] = None
                    st.session_state["evaluation_sheet_file"] = None
                    if row["result_text"]:
                        loaded_type = _rec_get(row, "interview_type", "생기부 기반 면접")
                        stu_path, tea_path = create_word_files(
                            row["result_text"], row["student_name"], loaded_type,
                            f"{row['university']}_{row['major']}"
                        )
                        st.session_state["word_files"] = (stu_path, tea_path)
                        # 요약카드·평가표는 AI 호출 없이 만들 수 있으므로 불러올 때 함께 다시 생성
                        try:
                            st.session_state["summary_card_file"] = create_summary_card_word(
                                row["result_text"], row["student_name"], row["university"], row["major"], loaded_type
                            )
                            st.session_state["evaluation_sheet_file"] = create_evaluation_sheet_word(
                                row["result_text"], row["student_name"], row["university"], row["major"], loaded_type
                            )
                        except Exception:
                            pass
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
                                chat_hist = load_json_safe(_rec_get(full, "chat_history", ""), [])
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
                                        "📈 성장 리포트 다운로드 (.docx)", f, file_name=os.path.basename(report_path),
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

    ### 📊 전국 140개 대학 실제 면접 기출 데이터베이스 연동!
    * `master_interview_qa.csv` 파일을 이 앱과 같은 폴더에 넣어두면, 지원 학과와 가장 유사한 **실제 대학 기출 질의응답 약 1.6만 건**을 자동으로 찾아 문항 생성에 참고합니다.
    * 현재 DB 로딩 상태: **{db_status}**

    ### 💾 [신규] 학생 기록 자동 저장 및 불러오기
    * 문항을 한 번 생성하면 해당 학생의 **생기부 텍스트 · 생성된 문항 · 피드백 대화 내역**이 자동으로 저장됩니다.
    * 다음에 다시 접속했을 때는 위쪽 **'📂 저장된 학생 기록 불러오기'**에서 이름을 선택해 불러오면, PDF를 다시 업로드하지 않고 이어서 진행할 수 있습니다.
    * ⚠️ 단, 이 저장 방식은 앱이 실행 중인 서버의 파일에 저장되는 방식입니다. 앱을 재배포(GitHub에 새로 커밋)하거나 서버가 완전히 재시작되면 저장된 기록이 초기화될 수 있으니, 중요한 학생 기록은 워드 파일로 다운로드해 별도 보관하시길 권장합니다.
    """.format(db_status=f"✅ {len(exam_db):,}건 로드 완료 ({exam_db['_uk'].nunique() if not exam_db.empty else 0}개 대학)" if not exam_db.empty else "⚠️ master_interview_qa.csv 파일을 찾지 못해 기본 학습 패턴만 사용 중입니다."))

UNIVERSITIES = {
    "서울권": ["서울대", "연세대", "고려대", "성균관대", "서강대", "한양대", "중앙대", "경희대", "한국외대", "서울시립대", "이화여대"],
    "충청권": ["카이스트(KAIST)", "충남대", "충북대", "고려대(세종)"],
    "경상권": ["경북대", "부산대", "UNIST", "영남대", "계명대"]
}

col1, col2 = st.columns(2)
with col1:
    interview_type = st.radio("🎯 면접 방식", ["생기부 기반 면접", "상위권 대학 제시문 기반 면접"], horizontal=True, key="interview_type_radio")
    _db_unis = cached_db_universities(APP_DIR)
    _db_uni_counts = dict(_db_unis)
    _regions = ["서울권", "충청권", "경상권"] + ([DB_UNI_REGION] if _db_unis else []) + ["직접 입력"]
    region = st.selectbox("📍 권역 선택", _regions, key="region_select")
    if region == "직접 입력":
        uni = st.text_input("🏫 대학 직접 입력", value="한국대", key="uni_direct_input")
    elif region == DB_UNI_REGION:
        # 기출이 있는 대학 전체를 건수가 많은 순으로 보여줍니다.
        uni = st.selectbox(
            "🏫 대학 선택 (기출 많은 순)", [u for u, _ in _db_unis],
            format_func=lambda u: f"{u} · 기출 {_db_uni_counts.get(u, 0):,}건", key="uni_db_select",
        )
    else:
        uni = st.selectbox("🏫 대학 선택", UNIVERSITIES[region], key="uni_select")
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
    render_interview_profile(cached_interview_profile(APP_DIR, uni, major), uni, major, interview_type, _preview_cases)

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
                student_record, result_text, st.session_state.chat_history
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
        if st.session_state.get("evaluation_sheet_file"):
            downloadable.append(("📝 교사용 평가표 (.docx)", st.session_state.evaluation_sheet_file))
        if st.session_state.get("growth_report_file"):
            downloadable.append(("📈 학생 성장 리포트 (.docx)", st.session_state.growth_report_file))

        downloadable = [(label, path) for label, path in downloadable if path and os.path.exists(path)]
        download_cols = st.columns(max(min(len(downloadable), 4), 1))
        for i, (label, path) in enumerate(downloadable):
            with download_cols[i % len(download_cols)]:
                with open(path, "rb") as f:
                    st.download_button(
                        label, f, file_name=os.path.basename(path),
                        use_container_width=True, key=f"dl_{i}_{os.path.basename(path)}"
                    )

    # -------------------------------------------------------------------
    # 💾 저장 (자동 저장 + 수동 저장 버튼)
    #    문항 생성·음성 평가·답변 첨삭·피드백 때마다 자동 저장되지만,
    #    자동 저장이 실패했거나 확실히 남겨두고 싶을 때 이 버튼으로 직접 저장할 수 있습니다.
    # -------------------------------------------------------------------
    save_col1, save_col_new, save_col2 = st.columns([1, 1, 2])
    with save_col_new:
        if st.button("🆕 새 기록으로 저장", use_container_width=True, key="manual_save_new_btn",
                     help="같은 학생이라도 이번 연습을 '새로운 회차'로 따로 남기고 싶을 때 사용하세요. (성장 리포트용 회차가 쌓입니다)"):
            st.session_state.current_record_id = new_record_id()
            _new_key = f"{student_name}|{uni}|{major}|{interview_type}"
            st.session_state.current_record_key = _new_key
            st.session_state.setdefault("record_key_map", {})[_new_key] = st.session_state.current_record_id
            new_info = save_current_session(student_name, uni, major, interview_type, difficulty, quiet=True)
            if new_info["ok"]:
                st.success(f"✅ 새 회차로 저장했습니다. ({new_info['at']} · {new_info['where']})")
            else:
                st.error(f"❌ 저장 실패: {new_info['error']}")
    with save_col1:
        if st.button("💾 지금 저장하기", use_container_width=True, key="manual_save_btn"):
            manual_info = save_current_session(student_name, uni, major, interview_type, difficulty, quiet=True)
            if manual_info["ok"] and "구글시트" in manual_info["where"]:
                st.success(f"✅ {manual_info['where']}에 저장했습니다. ({manual_info['at']})")
            elif manual_info["ok"]:
                st.info(f"💾 로컬에 임시 저장했습니다. ({manual_info['at']}) — 구글시트 연동이 설정되지 않은 상태입니다.")
            else:
                st.error(
                    f"❌ 구글시트 저장 실패: {manual_info['error']}\n\n"
                    "로컬에만 임시 저장했습니다. 사이드바의 저장 상태와 Apps Script 토큰·재배포 상태를 확인해 주세요."
                )
    with save_col2:
        render_save_status()

    if st.session_state.get("easy_explanation_text"):
        with st.expander("📚 생기부 쉬운 해설 미리보기 (고등학교 1학년 눈높이)", expanded=False):
            st.markdown(st.session_state.easy_explanation_text)

    # -------------------------------------------------------------------
    # 📝 교사용 면접 평가표 (인쇄해서 손으로 채점하는 양식)
    # -------------------------------------------------------------------
    with st.expander("📝 교사용 면접 평가표 만들기 (인쇄용)", expanded=False):
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

        if st.button("📝 평가표 만들기", use_container_width=True):
            try:
                date_str = interview_date_input.strftime("%Y. %m. %d.") if interview_date_input else ""
                st.session_state.evaluation_sheet_file = create_evaluation_sheet_word(
                    st.session_state.get("last_result_text", ""), student_name, uni, major, interview_type,
                    evaluator=evaluator_name.strip(), interview_date=date_str, include_questions=include_q
                )
                st.success("✅ 평가표가 만들어졌습니다. 아래 버튼으로 내려받아 인쇄하세요.")
            except Exception as e:
                st.error(f"❌ 평가표 생성에 실패했습니다: {e}")

        if st.session_state.get("evaluation_sheet_file") and os.path.exists(st.session_state.evaluation_sheet_file):
            with open(st.session_state.evaluation_sheet_file, "rb") as f:
                st.download_button(
                    "📥 평가표 다운로드 (.docx)", f,
                    file_name=os.path.basename(st.session_state.evaluation_sheet_file),
                    use_container_width=True, key="dl_eval_sheet_inline"
                )

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

                        if st.session_state.get("current_record_id"):
                            save_record(
                                st.session_state.current_record_id, student_name, uni, major, interview_type, difficulty,
                                st.session_state.get("loaded_student_record_text", ""),
                                new_result, st.session_state.chat_history
                            )

                        st.rerun()

                    except Exception as e:
                        st.error(f"피드백 반영 중 오류가 발생했습니다: {e}")
