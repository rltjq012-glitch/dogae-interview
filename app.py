import sys
import io

# 파이썬 표준 입출력 인코딩을 UTF-8로 완전히 강제 재설정합니다.
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import streamlit as st
import time
import pymupdf
import os
import re
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
/* 트렌디한 Pretendard 폰트 전역 적용 */
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
html, body, [class*="css"] {
    font-family: 'Pretendard', sans-serif !important;
}

/* 라이트모드 기본 배경색 */
@media (prefers-color-scheme: light) {
    [data-testid="stAppViewContainer"] {
        background-color: #F9F8F3 !important;
    }
    [data-testid="stSidebar"] {
        background-color: #F0EFEA !important;
    }
}

/* 다크모드 기본 배경색 */
@media (prefers-color-scheme: dark) {
    [data-testid="stAppViewContainer"] {
        background-color: #121212 !important;
    }
    [data-testid="stSidebar"] {
        background-color: #1A1A1A !important;
    }
}

/* 🎓 헤더 배너 (딥그린 & 골드 조합) */
.hero-banner {
    background: linear-gradient(135deg, #192c23 0%, #294435 100%);
    border-radius: 16px;
    padding: 2.5rem 3rem;
    color: white;
    margin-bottom: 2rem;
    box-shadow: 0 10px 25px rgba(0,0,0,0.15);
    position: relative;
    overflow: hidden;
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
    letter-spacing: -0.5px;
}
.hero-subtitle {
    font-size: 1.05rem;
    color: #a7f3d0;
    margin: 0;
    font-weight: 400;
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
    transition: all 0.3s ease !important;
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

/* 메인 동작 버튼 (생성 버튼) */
div.stButton > button:first-child {
    background: linear-gradient(135deg, #192c23 0%, #294435 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 0.7rem 1.2rem !important;
    font-weight: 700 !important;
    font-size: 1.1rem !important;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
    transition: all 0.3s ease !important;
    width: 100%;
}
div.stButton > button:first-child:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 12px rgba(0,0,0,0.2) !important;
}

/* 입력 필드 디자인 */
.stTextInput>div>div>input {
    border-radius: 8px !important;
    border: 1px solid #cbd5e1 !important;
}
.stTextInput>div>div>input:focus {
    border-color: #294435 !important;
    box-shadow: 0 0 0 1px #294435 !important;
}
.stFileUploader>div>div {
    border-radius: 12px !important;
    border: 2px dashed #cbd5e1 !important;
    background-color: rgba(255,255,255,0.5) !important;
}
@media (prefers-color-scheme: dark) {
    .stFileUploader>div>div {
        background-color: rgba(0,0,0,0.2) !important;
        border-color: #475569 !important;
    }
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# -------------------------------------------------------------------------
# 상단 헤더 배너 (HTML) - 제작자 명시
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
# [1] 스마트 PDF 및 제미나이 통신 함수 (텍스트 및 오디오 404 방어)
# -------------------------------------------------------------------------
def extract_text_from_pdf(uploaded_file):
    doc = pymupdf.open(stream=uploaded_file.read(), filetype="pdf")
    full_text = ""
    for page in doc:
        full_text += page.get_text()
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
        
    if not available_models:
        available_models = ["gemini-1.5-flash-latest"]
        
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
    2. 📊 **[면접관의 평가]**: 학생의 답변을 '논리성, 표현력, 전공적합성'을 기준으로 분석하고 종합 평가를 [상 / 중 / 하]로 매겨주세요. 구체적인 칭찬과 보완점(피드백)을 작성해 주세요.
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
        
    if not available_models:
        available_models = ["gemini-1.5-flash-latest", "gemini-1.5-flash"]
        
    models_to_try = sorted(available_models, key=lambda x: "flash" not in x)
    
    last_error = ""
    for target_model in models_to_try[:3]:
        try:
            audio_part = types.Part.from_bytes(data=audio_bytes, mime_type='audio/wav')
            response = client.models.generate_content(
                model=target_model,
                contents=[audio_part, prompt]
            )
            return response.text
        except Exception as e:
            last_error = str(e)
            time.sleep(2)
            
    return f"음성 분석 실패: {last_error}\n(API 키 설정이나 모델 상태를 확인해 주세요.)"

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
        
        if line.startswith("### 📄") or line.startswith("### 📌") or line.startswith("### 🔍"):
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
                q1_match = re.search(r'\[문제 1\](.*?)(?=\[평가의도 1\]|\[문제 2\]|$)', body, re.DOTALL)
                i1_match = re.search(r'\[평가의도 1\](.*?)(?=\[모범답안 1\]|$)', body, re.DOTALL)
                a1_match = re.search(r'\[모범답안 1\](.*?)(?=\[꼬리질문 1\]|$)', body, re.DOTALL)
                f1_match = re.search(r'\[꼬리질문 1\](.*?)(?=\[문제 2\]|$)', body, re.DOTALL)
                
                q2_match = re.search(r'\[문제 2\](.*?)(?=\[평가의도 2\]|$)', body, re.DOTALL)
                i2_match = re.search(r'\[평가의도 2\](.*?)(?=\[모범답안 2\]|$)', body, re.DOTALL)
                a2_match = re.search(r'\[모범답안 2\](.*?)(?=\[꼬리질문 2\]|$)', body, re.DOTALL)
                f2_match = re.search(r'\[꼬리질문 2\](.*?)(?=$)', body, re.DOTALL)
                
                js_text = js_match.group(1).strip() if js_match else ""
                q1_text = q1_match.group(1).strip() if q1_match else ""
                i1_text = i1_match.group(1).strip() if i1_match else ""
                a1_text = a1_match.group(1).strip() if a1_match else ""
                f1_text = f1_match.group(1).strip() if f1_match else ""
                
                q2_text = q2_match.group(1).strip() if q2_match else ""
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
                        row_i2 = table.add_row()
                        set_cell_background(row_i2.cells[0], "F9F9F9")
                        add_parsed_text_to_cell(row_i2.cells[0], f"**[평가 의도 2]**\n{i2_text}")
                        row_a2 = table.add_row()
                        add_parsed_text_to_cell(row_a2.cells[0], f"**[모범 답안 가이드 2]**\n{a2_text}")
                        row_f2 = table.add_row()
                        set_cell_background(row_f2.cells[0], "FFF4F4")
                        add_parsed_text_to_cell(row_f2.cells[0], f"**[압박용 꼬리질문 2]**\n{f2_text}")
            else:
                q_match = re.search(r'\[질문\](.*?)(?=\[평가의도\]|\[모범답안\]|\[꼬리질문\]|$)', body, re.DOTALL)
                i_match = re.search(r'\[평가의도\](.*?)(?=\[모범답안\]|\[꼬리질문\]|$)', body, re.DOTALL)
                a_match = re.search(r'\[모범답안\](.*?)(?=\[꼬리질문\]|$)', body, re.DOTALL)
                f_match = re.search(r'\[꼬리질문\](.*?)(?=$)', body, re.DOTALL)
                
                q_text = q_match.group(1).strip() if q_match else "내용 없음"
                i_text = i_match.group(1).strip() if i_match else ""
                a_text = a_match.group(1).strip() if a_match else ""
                f_text = f_match.group(1).strip() if f_match else ""
                
                row_q = table.add_row()
                add_parsed_text_to_cell(row_q.cells[0], f"**[면접 질문]**\n{q_text}")
                
                if is_teacher:
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

# -------------------------------------------------------------------------
# [3] 메인 UI 
# -------------------------------------------------------------------------
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "word_files" not in st.session_state:
    st.session_state.word_files = None
if "last_audio_size" not in st.session_state:
    st.session_state.last_audio_size = 0

with st.expander("📖 [클릭] 프로그램 사용 설명서 및 PDF OCR 변환 방법", expanded=False):
    st.markdown("""
    ### 🔑 Google Gemini API 키 발급 방법
    1. **Google AI Studio 접속:** [Google AI Studio](https://aistudio.google.com/)에 접속합니다.
    2. **구글 계정 로그인:** 평소 사용하는 구글 계정으로 로그인합니다.
    3. **API 키 생성:** 좌측 상단 'Get API key' 버튼을 클릭하여 키 생성 후 복사합니다.

    ### 📂 [중요] 생기부 PDF는 반드시 '텍스트 추출(OCR)'된 파일이어야 합니다!
    * **왜 필요한가요?** 단순 이미지(스캔본) PDF는 AI가 글자를 읽지 못하므로, **마우스로 글자가 드래그되거나 텍스트로 인식되는 PDF**여야만 정상 분석이 가능합니다.
    * **스캔된 PDF를 OCR(텍스트형)로 변환하는 방법:**
      1. **구글 드라이브(Google Drive) 활용 (가장 추천):**
         - 스캔된 생기부 PDF 파일을 구글 드라이브에 업로드합니다.
         - 업로드된 파일 우클릭 ➔ **[연결 앱]** ➔ **[Google 문서]**를 선택하여 엽니다.
         - 구글 문서로 열리면 이미지가 텍스트로 자동 변환됩니다. 상단 메뉴 **[파일] ➔ [다운로드] ➔ [Microsoft Word(.docx)]** 또는 PDF로 저장합니다.
      2. **온라인 무료 툴 활용:** 'ILOVEPDF' 등의 사이트에서 'OCR PDF' 메뉴를 이용해 변환합니다.

    ### 🎙️ [신규] 휴대폰 음성 인식(STT) 면접 평가 기능 사용법!
    * **1단계 (키보드 활용):** 휴대폰으로 접속 시, 하단 채팅창을 누른 후 **휴대폰 키보드에 있는 '마이크(🎤)' 버튼**을 누르고 말하면 텍스트로 바로 입력됩니다.
    * **2단계 (무인 AI 면접관 모드):** 하단의 **[🎙️ 음성으로 면접 답변하기]** 버튼을 눌러 직접 녹음해 보세요. AI가 음성을 듣고 즉각적인 평가와 꼬리질문을 던져줍니다!
    """)

with st.sidebar:
    api_key = st.text_input("🔑 Gemini API Key", type="password")

UNIVERSITIES = {
    "서울권": ["서울대", "연세대", "고려대", "성균관대", "서강대", "한양대", "중앙대", "경희대", "한국외대", "서울시립대", "이화여대"], 
    "충청권": ["카이스트(KAIST)", "충남대", "충북대", "고려대(세종)"], 
    "경상권": ["경북대", "부산대", "UNIST", "영남대", "계명대"]
}

col1, col2 = st.columns(2)
with col1:
    interview_type = st.radio("🎯 면접 방식", ["생기부 기반 면접", "상위권 대학 제시문 기반 면접"], horizontal=True)
    region = st.selectbox("📍 권역 선택", ["서울권", "충청권", "경상권", "직접 입력"])
    if region == "직접 입력":
        uni = st.text_input("🏫 대학 직접 입력", value="한국대")
    else:
        uni = st.selectbox("🏫 대학 선택", UNIVERSITIES[region])

with col2:
    major = st.text_input("🎓 지원 학과/전공", placeholder="예: 철학과")
    student_name = st.text_input("👤 지원자 성명", value="김기섭")
    difficulty = st.radio("⚙️ 난이도 선택", ["하 (기초)", "중 (표준)", "상 (압박)"], horizontal=True, index=1)

uploaded_file = None
if interview_type == "생기부 기반 면접":
    uploaded_file = st.file_uploader("📂 학생 생기부 PDF 업로드 (OCR 변환 필수)", type=["pdf"])

st.markdown("---")

# -------------------------------------------------------------------------
# [4] 생성 버튼 로직
# -------------------------------------------------------------------------
target_desc = f"{uni}_{major}" 

TEMPLATE_SANGBU = """
### 🔍 [생기부 심층 분석 브리핑 리포트]
- **지원자 강점 분석:** (여기에 작성)
- **보완점 및 약점 분석:** (여기에 작성)
- **면접관 집중 공략 포인트 (심화 탐구 상기):** (여기에 작성)

### 📌 [영역: OOO] **[과목명/활동명]**
[질문]
(내용 작성)
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
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)
[문제 2]
(문제 2 내용)
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
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)
[문제 2]
(문제 2 내용)
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
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)
[문제 2]
(문제 2 내용)
[평가의도 2]
(내용 작성)
[모범답안 2]
(내용 작성)
[꼬리질문 2]
(내용 작성)
"""

if st.button("🚀 면접 패키지 생성 시작"):
    if not api_key: st.error("API 키를 입력해 주세요."); st.stop()
    if not major: st.error("지원 학과를 입력해 주세요."); st.stop()
    if interview_type == "생기부 기반 면접" and not uploaded_file: st.error("생기부 파일을 업로드해 주세요."); st.stop()
        
    student_record = extract_text_from_pdf(uploaded_file) if uploaded_file else ""
    
    if interview_type == "생기부 기반 면접":
        prompt = f"""
        당신은 도개고등학교의 진학 지도 노하우와 {uni} {major} 입학사정관의 시각을 겸비한 최고급 면접 출제위원입니다. 
        지원자 '{student_name}' 학생의 생기부를 면밀히 분석하여 다음 작업을 수행하세요.
        
        [지시사항]
        1. **출력의 맨 첫 부분**에 반드시 **[생기부 심층 분석 브리핑 리포트]**를 작성하여 학생의 **강점, 단점(보완점), 그리고 면접관이 예리하게 파고들 수 있는 심화탐구 활동 부분**을 상세히 짚어주어 학생이 면접 전 반드시 상기할 수 있도록 하세요.
        2. 생기부 5대 영역(교과세특, 창체, 동아리, 행특, 독서 등)을 모두 분석하여 총 5세트의 면접 문항을 만드세요.
        3. 과목명이나 주요 활동명은 반드시 **[생활과 윤리]** 처럼 볼드체로 묶어주고 도개고 기출 수준의 날카로운 꼬리질문을 포함하세요.
        
        [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
        {TEMPLATE_SANGBU}
        
        [생기부 내용]
        {student_record}
        """
    else: 
        prompt = f"""
        당신은 서울대학교 면접 및 구술고사 출제위원입니다. {major} 전공적합성과 종합적 사고력, 논리적 추론 능력을 평가하기 위한 고난도 제시문 기반 구술고사를 출제하세요.
        면접 난이도: {difficulty}
        
        [지시사항]
        1. 생기부 내용은 무시하세요. {major} 학과와 관련된 학술적 딜레마와 심층 개념을 담은 **완전 독립된 3개의 주제 세트**를 창작하세요.
        2. **각 세트마다 복수의 제시문((가), (나), (다) 형태)과 [문제 1], [문제 2] (각각 평가의도, 모범답안, 압박 꼬리질문 포함)**가 유기적으로 묶인 **총 3개의 독립 세트**를 엄격히 만드세요.
        3. 서론이나 인사말은 절대 쓰지 말고, 바로 '### 📌 [세트 1]' 부터 출력하세요.
        
        [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
        {TEMPLATE_JESIMUN}
        """
    
    with st.spinner(f"⏳ 로딩중... 생기부 분석 브리핑 및 면접 문항을 정밀 조립하고 있습니다."):
        try:
            result_text = call_gemini(prompt, api_key)
            
            stu_path, tea_path = create_word_files(result_text, student_name, interview_type, target_desc)
            st.session_state.word_files = (stu_path, tea_path)
            
            display_text = result_text.replace('[문제 1]', '\n**💡 [문제 1]**\n').replace('[평가의도 1]', '\n**🎯 [평가 의도 1]**\n').replace('[모범답안 1]', '\n**✅ [모범 답안 가이드 1]**\n').replace('[꼬리질문 1]', '\n**🔥 [압박용 꼬리질문 1]**\n')
            display_text = display_text.replace('[문제 2]', '\n**💡 [문제 2]**\n').replace('[평가의도 2]', '\n**🎯 [평가 의도 2]**\n').replace('[모범답안 2]', '\n**✅ [모범 답안 가이드 2]**\n').replace('[꼬리질문 2]', '\n**🔥 [압박용 꼬리질문 2]**\n')
            display_text = display_text.replace('[질문]', '\n**💡 [면접 질문]**\n').replace('[평가의도]', '\n**🎯 [평가 의도]**\n').replace('[모범답안]', '\n**✅ [모범 답안 가이드]**\n').replace('[꼬리질문]', '\n**🔥 [압박용 꼬리질문]**\n')
            display_text = display_text.replace('[제시문]', '\n**📄 [서울대 스타일 구술 제시문 (가, 나, 다)]**\n')
            
            full_display_text = display_text + "\n\n---\n💬 **방금까지 나눈 문항 내용과 피드백 대화 내용을 한글 문서(.docx)로 만들어 드릴까요?** (원하시면 **'그래 만들어줘'**라고 말씀해 주세요!)"
            
            st.session_state.chat_history = [{"role": "assistant", "content": full_display_text}]
            st.success("🎉 면접 패키지 및 워드 문서 생성이 완료되었습니다!")
            
        except Exception as e:
            st.error(f"❌ 생성 실패: {e}")

# -------------------------------------------------------------------------
# [5] 결과 대시보드 및 실시간 피드백
# -------------------------------------------------------------------------
if st.session_state.chat_history:
    st.markdown("<div class='step-text'>STEP 2 · 결과 확인 및 피드백</div>", unsafe_allow_html=True)
    st.markdown("## 📋 면접 문항 대시보드 및 실시간 피드백")
    
    if st.session_state.word_files:
        stu_path, tea_path = st.session_state.word_files
        chat_path = create_chat_history_word(st.session_state.chat_history, student_name)
        
        col_w1, col_w2, col_w3 = st.columns(3)
        with col_w1:
            with open(stu_path, "rb") as f:
                st.download_button("📥 학생용 워크북 (.docx)", f, file_name=stu_path, use_container_width=True)
        with col_w2:
            with open(tea_path, "rb") as f:
                st.download_button("📥 교사용 지침서 (.docx)", f, file_name=tea_path, use_container_width=True)
        with col_w3:
            with open(chat_path, "rb") as f:
                st.download_button("💬 피드백 대화 내역 (.docx)", f, file_name=chat_path, use_container_width=True)
    
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
                    
                    add_instruction = ""
                    if any(w in user_feedback for w in ["더", "추가", "많이", "늘려", "또"]):
                        add_instruction = "\n(주의: 사용자가 문항 추가를 요청했으므로, 독립 세트를 **최소 2세트 이상 추가로** 더 생성해 주세요!)"
                    
                    feedback_prompt = f"""
                    당신은 면접 출제위원입니다. 이전 내용을 사용자의 피드백에 맞게 수정하되, 
                    심도 있는 학술적/실천적 깊이를 유지하면서 **반드시 다음 템플릿 구조와 [키워드]를 토씨 하나 틀리지 말고 유지**해 주세요.
                    {add_instruction}
                    
                    [강제 유지 템플릿]
                    {required_template}
                    
                    사용자 피드백: "{user_feedback}"
                    """
                    try:
                        new_result = call_gemini(feedback_prompt, api_key)
                        display_text = new_result.replace('[문제 1]', '\n**💡 [문제 1]**\n').replace('[평가의도 1]', '\n**🎯 [평가 의도 1]**\n').replace('[모범답안 1]', '\n**✅ [모범 답안 가이드 1]**\n').replace('[꼬리질문 1]', '\n**🔥 [압박용 꼬리질문 1]**\n')
                        display_text = display_text.replace('[문제 2]', '\n**💡 [문제 2]**\n').replace('[평가의도 2]', '\n**🎯 [평가 의도 2]**\n').replace('[모범답안 2]', '\n**✅ [모범 답안 가이드 2]**\n').replace('[꼬리질문 2]', '\n**🔥 [압박용 꼬리질문 2]**\n')
                        display_text = display_text.replace('[질문]', '\n**💡 [면접 질문]**\n').replace('[평가의도]', '\n**🎯 [평가 의도]**\n').replace('[모범답안]', '\n**✅ [모범 답안 가이드]**\n').replace('[꼬리질문]', '\n**🔥 [압박용 꼬리질문]**\n')
                        display_text = display_text.replace('[제시문]', '\n**📄 [서울대 스타일 구술 제시문 (가, 나, 다)]**\n')
                        
                        full_response = display_text + "\n\n---\n💬 **방금까지 나눈 문항 내용과 피드백 대화 내용을 한글 문서(.docx)로 만들어 드릴까요?** (원하시면 **'그래 만들어줘'**라고 말씀해 주세요!)"
                        
                        st.markdown(full_response)
                        st.session_state.chat_history.append({"role": "assistant", "content": full_response})
                        
                        stu_path, tea_path = create_word_files(new_result, student_name, interview_type, target_desc)
                        st.session_state.word_files = (stu_path, tea_path)
                        st.rerun() 
                        
                    except Exception as e:
                        st.error(f"피드백 반영 중 오류가 발생했습니다: {e}")