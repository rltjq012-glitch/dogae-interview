"""기출 DB · 대학별 면접 형식 · 출제 참고자료(대학 스타일·학과 기출) · 출제 유형 분석

app.py에서 나눈 모듈입니다. 아래 층(dogae.base, dogae.ai, dogae.storage)만 가져다 씁니다."""
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

def _diversify_examples(hits, top_n, type_share=None):
    """일치도가 같은 후보 안에서 질문 유형과 출처 대학이 한쪽으로 쏠리지 않게 뽑습니다.
    type_share(지원 대학의 실제 유형 비중)가 있으면, 그 비중에 가깝게 유형별 개수를 맞춥니다.
    → 다른 대학의 같은 학과 기출을 가져올 때도 '지원 대학이 자주 묻는 유형' 위주로 선별됩니다."""
    picked = []
    for score in sorted(hits["_score"].unique()):
        if len(picked) >= top_n:
            break
        tier = hits[hits["_score"] == score]
        good = tier[~tier["_type"].isin(_LOW_VALUE_QTYPES) & (tier["질문"].str.len() >= 12) & ~tier["_follow"]]
        rest = tier.drop(index=good.index)
        queues = {}
        for t, g in good.groupby("_type", sort=False):
            # 한 대학의 질문만 연달아 뽑히지 않게, 대학별로 한 건씩 돌아가며 줄을 세웁니다.
            g = g.assign(_r=g.groupby("_uk").cumcount()).sort_values("_r", kind="stable")
            queues[t] = list(g.index)
        total = sum(len(q) for q in queues.values())
        room = top_n - len(picked)
        weight = {}
        for t, q in queues.items():
            own = len(q) / total if total else 0
            weight[t] = 0.6 * type_share.get(t, 0) + 0.4 * own if type_share else own
        order = sorted(queues, key=lambda t: (-weight[t], -len(queues[t])))
        wsum = sum(weight.values())
        if type_share and wsum > 0:
            for t in order:
                k = min(int(round(weight[t] / wsum * room)), len(queues[t]), top_n - len(picked))
                picked.extend(queues[t][:k])
                queues[t] = queues[t][k:]
        while any(queues.values()) and len(picked) < top_n:
            for t in order:
                if queues[t] and len(picked) < top_n:
                    picked.append(queues[t].pop(0))
        for idx in rest.index:
            if len(picked) >= top_n:
                break
            picked.append(idx)
    return hits.loc[picked]

def get_relevant_examples(df, major, uni_name, top_n=12, type_share=None):
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
    # 비슷한 학과가 전혀 없으면 빈 결과를 돌려줍니다. (예전에는 지원 대학의 아무 학과나 무작위로 채워
    # '지원 학과 기출이 있다'고 잘못 안내되는 경우가 있었습니다. 대학 스타일은 A 자료가 따로 담당합니다.)
    if hits.empty:
        return hits
    return _diversify_examples(hits, top_n, type_share=type_share)

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
    # 먼저 이 대학·학과의 실제 출제 비중을 구하고, 그 비중에 맞춰 참고 기출을 선별합니다.
    profile = build_interview_profile(RT.exam_db, uni_name, major)
    type_share, _ = pick_allocation_basis(profile)
    examples = get_relevant_examples(RT.exam_db, major, uni_name, top_n=top_n, type_share=type_share or None)
    # '지원 대학에 지원 학과 기출이 있다'는 판정은 학과가 일치(_score 0)할 때만 내립니다.
    has_same_dept_at_uni = bool(len(examples)) and bool((examples["_score"] == 0).any())
    style = analyze_university_style(RT.exam_db, uni_name)
    style_examples = get_university_style_examples(
        RT.exam_db, uni_name, major, n=style_n,
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
            f"\n[참고자료 사용 지침 — ⚠️ 중요: 타 대학 기출을 {uni_name} 스타일로 재구성]\n"
            f"- DB에 **{uni_name} {major} 기출은 없습니다.** 그래서 아래 순서로 재구성해야 합니다.\n"
            f"- ① 비교: A({uni_name}의 다른 학과 기출)를 서로 비교해, 이 대학이 **학과를 가리지 않고 공통으로 쓰는 "
            f"질문 방식**(무엇을 먼저 묻는지, 말투와 길이, 꼬리질문을 어디까지 파고드는지)을 먼저 파악하세요.\n"
            f"- ② 선별: B(다른 대학의 {major} 기출)에서 **{major} 지원자에게 실제로 묻는 전공 내용과 깊이**만 골라내세요. "
            f"B는 {uni_name}이 자주 묻는 유형 위주로 이미 선별되어 있습니다. B에 ★ 표시된 것은 {uni_name}의 유사 학과 기출입니다.\n"
            f"- ③ 재구성: ②의 내용을 ①의 방식으로 다시 쓰세요. 즉 **'{uni_name} 면접관이 {major} 지원자에게 물었을 법한 질문'**을 "
            f"이 학생의 생기부 내용으로 만듭니다. B의 다른 대학 말투·형식은 따라 하지 마세요.\n"
            f"- A의 다른 학과 전공 내용은 절대 가져오지 마세요. A에서는 '묻는 방식'만 가져옵니다.\n"
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
MAJOR_OTHER = "이 대학 기출에 없는 학과 → 전국 학과 목록에서 선택"
MAJOR_DIRECT = "학과 직접 입력"
_REGION_MEMBERS = {
    "서울권": "서울 동국 경희 건국 숭실 서울시립 한국외국어 국민 서울과학기술 광운 이화여자 명지 중앙 숙명여자 성신여자 "
              "세종 서울여자 고려 연세 상명 삼육 홍익 동덕여자 성균관 덕성여자 서경 한양 장로회신학 강서 성공회 총신 "
              "서강 백석예술",
    "경기·인천권": "가천 아주 경기 가톨릭 단국 인하 한국항공 인천 을지 한국공학 수원 한양|ERICA 협성 신한 대진 차의과학 "
                   "강남 안양 한세 평택 연성 한신",
    "충청권": "공주 충남 충북 건국|글로컬 고려|세종 홍익|세종 순천향 우송 서원 한서 백석 한국기술교육 호서 단국|천안 청주 "
              "한국교통 세명 한국전통문화 대전 상명|천안 중부 중원 남서울 건양 한밭 나사렛 선문 가톨릭꽃동네 목원 유원 "
              "한국과학기술원",
    "경상권 (대구·경북)": "경북 계명 대구가톨릭 대구 경운 경일 경국 대구한의 한동 동국|WISE 영남 위덕 김천 금오공과 "
                          "포항공과 대구경북과학기술원 대구보건 동양 영남이공 구미 선린 계명문화 대경 대구과학",
    "경상권 (부산·울산·경남)": "부산 동아 경상국립 울산 경성 창원 부산가톨릭 인제 신라 고신 경남 연암공과 "
                               "울산과학기술원 부산외국어",
    "강원·호남권": "강원 전남 전북 원광 한림 조선 광주과학기술원 한국에너지공과 연세|미래",
}
UNI_REGION_BY_KEY = {k: region for region, keys in _REGION_MEMBERS.items() for k in keys.split()}
# 기출 DB에는 없지만 선택은 할 수 있어야 하는 대학 (다른 대학 기출로 재구성됩니다)
EXTRA_UNIVERSITIES = {
    "서울권": ["서강대학교"],
    "충청권": ["충북대학교", "고려대학교(세종)"],
    "교육대·교원대": ["경인교육대학교", "춘천교육대학교", "광주교육대학교"],
}

def region_of_university(name):
    uk = uni_key(name)
    base = uk.partition("|")[0]
    known = UNI_REGION_BY_KEY.get(uk) or UNI_REGION_BY_KEY.get(base)
    if known:
        return known
    if base.endswith("교육") or base == "한국교원":      # ○○교육대학교, 한국교원대학교
        return "교육대·교원대"
    if "사관학교" in base:
        return "사관학교"
    return "기타 대학"

def build_region_universities(db_unis):
    """{권역: [(대학 표시명, 기출 건수)]} — 권역 안에서는 기출이 많은 순."""
    out = {}
    seen = set()
    for label, n in db_unis:
        out.setdefault(region_of_university(label), []).append((label, n))
        seen.add(uni_key(label))
    for region, labels in EXTRA_UNIVERSITIES.items():
        for label in labels:
            if uni_key(label) not in seen:
                out.setdefault(region, []).append((label, 0))
    return {r: sorted(v, key=lambda x: (-x[1], x[0])) for r, v in out.items()}

@st.cache_data(show_spinner=False)
def cached_uni_departments(folder, uni_name):
    """이 대학의 기출 보유 학과 [(학과, 건수)] — 건수가 많은 순."""
    df, _ = load_exam_db(folder)
    if df.empty or not uni_name:
        return []
    sub = df[_uni_mask(df, uni_key(uni_name))]
    return [(d, int(c)) for d, c in sub["학과"].value_counts().items() if d]

@st.cache_data(show_spinner=False)
def cached_all_departments(folder):
    """전국 학과 [(학과, 건수, 보유 대학 수)] — 건수가 많은 순."""
    df, _ = load_exam_db(folder)
    if df.empty:
        return []
    g = df[df["학과"] != ""].groupby("학과").agg(n=("질문", "size"), unis=("_uk", "nunique"))
    g = g.sort_values(["n", "unis"], ascending=False)
    return [(d, int(r.n), int(r.unis)) for d, r in g.iterrows()]

@st.cache_data(show_spinner=False)
def cached_reference_preview(folder, uni_name, major):
    """문항 생성 때 실제로 쓰일 참고 방식(어느 대학·학과 기출을 어떻게 결합하는지)을 미리 계산합니다."""
    df, _ = load_exam_db(folder)
    stg = build_reference_strategy(df, uni_name, major, include_profile=False)
    ex, sx = stg["examples"], stg["style_examples"]
    same = _uni_mask(ex, uni_key(uni_name)) if len(ex) else pd.Series(dtype=bool)
    sources = []
    if len(ex):
        other = ex[~same]
        sources = [f"{u.split('(')[0]} {d}" for (u, d), _ in other.groupby(["대학", "학과"], sort=False)]
    return {
        "mode": stg["mode"],
        "n_examples": int(len(ex)),
        "n_same_uni": int(same.sum()) if len(ex) else 0,
        "same_uni_depts": sorted(set(ex[same]["학과"])) if len(ex) else [],
        "sources": list(dict.fromkeys(sources)),
        "style_depts": list(dict.fromkeys(sx["학과"].tolist())) if sx is not None and len(sx) else [],
    }

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


__all__ = ["QA_COLUMNS", "INFO_COLUMNS", "_find_data_files", "load_exam_db", "load_univ_info", "UNI_ALIASES", "SEPARATE_CAMPUS", "_CAMPUS_TOKENS", "uni_key", "uni_keys_match", "_normalize_uni", "_uni_matches", "dept_key", "_normalize_dept", "FIELD_RULES", "_FIELD_COMPILED", "field_of", "_field_of", "track_key", "QTYPE_JESIMUN", "QTYPE_INTRO", "QTYPE_ETC", "QTYPE_RULES", "_QTYPE_COMPILED", "QTYPE_DESC", "SANGBU_QTYPES", "_FOLLOWUP_RX", "classify_question", "_PAGE_NOISE_RX", "_clean_exam_text", "prepare_exam_db", "_ensure_prepared", "_uni_mask", "list_db_universities", "_LOW_VALUE_QTYPES", "_diversify_examples", "get_relevant_examples", "PROFILE_MIN_ROWS", "_summarize_level", "build_interview_profile", "ALLOC_MIN_ROWS", "ALLOC_SHRINK", "_sangbu_counts", "pick_allocation_basis", "allocate_question_types", "format_profile_for_prompt", "analyze_university_style", "get_university_style_examples", "_CASE_TIER_LABEL", "_EMPTY_ANSWER_RX", "_safe_case_text", "_char_trigrams", "select_real_cases", "format_real_case_block", "format_real_cases_for_prompt", "_REAL_CASE_RX", "strip_real_cases", "extract_real_cases", "inject_real_cases", "format_university_style_for_prompt", "get_univ_info", "format_univ_info_for_prompt", "format_examples_for_prompt", "build_reference_strategy", "MAJOR_OTHER", "MAJOR_DIRECT", "_REGION_MEMBERS", "UNI_REGION_BY_KEY", "EXTRA_UNIVERSITIES", "region_of_university", "build_region_universities", "cached_uni_departments", "cached_all_departments", "cached_reference_preview", "cached_db_universities", "cached_interview_profile", "cached_real_cases"]
