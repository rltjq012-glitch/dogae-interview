import streamlit as st
import time
import pymupdf
import os
import re
from google import genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(page_title="도개고 면접 마스터", layout="wide")

# -------------------------------------------------------------------------
# [1] 스마트 PDF 및 제미나이 통신 함수
# -------------------------------------------------------------------------
def extract_text_from_pdf(uploaded_file):
    doc = pymupdf.open(stream=uploaded_file.read(), filetype="pdf")
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    return full_text

def call_gemini(prompt, api_key):
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

# -------------------------------------------------------------------------
# [2] 워드 표 생성 및 가독성/여백 최적화 엔진
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

def add_parsed_text_to_cell(cell, text):
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
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
        
        if line.startswith("### 📄"):
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
        
        blocks = content.split('### 📌')
        intro_text = blocks[0].strip()
        if intro_text:
            add_intro_paragraphs(doc, intro_text)
            doc.add_paragraph() 
        
        for block in blocks[1:]:
            lines = block.strip().split('\n')
            if not lines: continue
            
            title = lines[0].strip()
            body = "\n".join(lines[1:])
            
            table = doc.add_table(rows=0, cols=1)
            table.style = 'Table Grid' 
            
            row_title = table.add_row()
            set_cell_background(row_title.cells[0], "EBF1FA")
            add_parsed_text_to_cell(row_title.cells[0], f"📌 {title}")
            
            if is_jesimun:
                # 제시문 면접: 질문 1, 질문 2 파싱
                q1_match = re.search(r'\[질문 1\](.*?)(?=\[평가의도 1\]|\[질문 2\]|$)', body, re.DOTALL)
                i1_match = re.search(r'\[평가의도 1\](.*?)(?=\[모범답안 1\]|$)', body, re.DOTALL)
                a1_match = re.search(r'\[모범답안 1\](.*?)(?=\[꼬리질문 1\]|$)', body, re.DOTALL)
                f1_match = re.search(r'\[꼬리질문 1\](.*?)(?=\[질문 2\]|$)', body, re.DOTALL)
                
                q2_match = re.search(r'\[질문 2\](.*?)(?=\[평가의도 2\]|$)', body, re.DOTALL)
                i2_match = re.search(r'\[평가의도 2\](.*?)(?=\[모범답안 2\]|$)', body, re.DOTALL)
                a2_match = re.search(r'\[모범답안 2\](.*?)(?=\[꼬리질문 2\]|$)', body, re.DOTALL)
                f2_match = re.search(r'\[꼬리질문 2\](.*?)(?=$)', body, re.DOTALL)
                
                q1_text = q1_match.group(1).strip() if q1_match else ""
                i1_text = i1_match.group(1).strip() if i1_match else ""
                a1_text = a1_match.group(1).strip() if a1_match else ""
                f1_text = f1_match.group(1).strip() if f1_match else ""
                
                q2_text = q2_match.group(1).strip() if q2_match else ""
                i2_text = i2_match.group(1).strip() if i2_match else ""
                a2_text = a2_match.group(1).strip() if a2_match else ""
                f2_text = f2_match.group(1).strip() if f2_match else ""
                
                # 질문 1 행 추가
                row_q1 = table.add_row()
                add_parsed_text_to_cell(row_q1.cells[0], f"**[질문 1]**\n{q1_text}")
                if is_teacher:
                    row_i1 = table.add_row()
                    set_cell_background(row_i1.cells[0], "F9F9F9")
                    add_parsed_text_to_cell(row_i1.cells[0], f"**[평가 의도 1]**\n{i1_text}")
                    
                    row_a1 = table.add_row()
                    add_parsed_text_to_cell(row_a1.cells[0], f"**[모범 답안 가이드 1]**\n{a1_text}")
                    
                    row_f1 = table.add_row()
                    set_cell_background(row_f1.cells[0], "FFF4F4")
                    add_parsed_text_to_cell(row_f1.cells[0], f"**[압박용 꼬리질문 1]**\n{f1_text}")
                
                # 질문 2 행 추가
                if q2_text:
                    row_q2 = table.add_row()
                    add_parsed_text_to_cell(row_q2.cells[0], f"**[질문 2]**\n{q2_text}")
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
                # 생기부 면접 형식
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
        role_title = "👤 선생님 (요청)" if msg["role"] == "user" else "🤖 AI 출제위원 (답변/수정본)"
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

st.title("🎓 도개고 대입 모의면접 마스터 솔루션")

# 🌟 [업데이트] OCR 변환 가이드가 포함된 상세 설명서 🌟
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
      2. **온라인 무료 툴 활용:** 
         - 'ILOVEPDF' 또는 'Smallpdf' 등의 사이트에서 **'OCR PDF(PDF 텍스트 변환)'** 메뉴를 이용해 변환합니다.

    ### 💡 프로그램 사용 순서
    1. 왼쪽 사이드바에 API 키를 입력합니다.
    2. 면접 방식(생기부 기반/제시문 기반)을 선택합니다.
    3. 대학, 학과, 난이도를 설정합니다.
    4. 생기부 면접일 경우 **OCR 변환된 PDF**를 업로드하고 생성 버튼을 누릅니다.
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
### 📄 [제시문 1]
(제시문 1 내용)

### 📄 [제시문 2]
(제시문 2 내용)

### 📄 [제시문 3]
(제시문 3 내용)

### 📌 [세트 1] **[제시문 1 기반 문항]**
[질문 1]
(질문 1 내용)
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)

[질문 2]
(질문 2 내용)
[평가의도 2]
(내용 작성)
[모범답안 2]
(내용 작성)
[꼬리질문 2]
(내용 작성)

### 📌 [세트 2] **[제시문 2 기반 문항]**
[질문 1]
(질문 1 내용)
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)

[질문 2]
(질문 2 내용)
[평가의도 2]
(내용 작성)
[모범답안 2]
(내용 작성)
[꼬리질문 2]
(내용 작성)

### 📌 [세트 3] **[제시문 3 기반 문항]**
[질문 1]
(질문 1 내용)
[평가의도 1]
(내용 작성)
[모범답안 1]
(내용 작성)
[꼬리질문 1]
(내용 작성)

[질문 2]
(질문 2 내용)
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
        지원자 '{student_name}' 학생의 생기부를 분석하여 도개고 선배들이 실제 상위권/지역거점 대학 면접에서 마주했던 수준 높은 기출 문항들의 난이도와 깊이를 반영해 질문을 만드세요[cite: 3, 4].
        면접 난이도: {difficulty}
        
        [지시사항]
        1. 생기부 5대 영역(1. 교과세특, 2. 창체, 3. 동아리, 4. 행특, 5. 독서/기타)을 모두 뒤져서 총 5세트를 만드세요.
        2. 과목명이나 주요 활동명은 반드시 **[생활과 윤리]** 처럼 볼드체로 묶어주세요.
        3. 단순 암기식 질문이 아니라 학생부 활동의 진위 여부, 동기, 심화 탐구 과정, 실생활 적용력을 파악하는 도개고 스타일의 날카로운 꼬리질문을 포함하세요[cite: 3, 4].
        
        [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
        {TEMPLATE_SANGBU}
        
        [생기부 내용]
        {student_record}
        """
    else: 
        prompt = f"""
        당신은 도개고등학교 진학 지도 기준에 맞춘 상위권 대학 제시문 면접 출제위원입니다. {major} 전공적합성과 논리력을 평가하기 위한 고난도 제시문 기반 면접을 출제하세요.
        면접 난이도: {difficulty}
        
        [지시사항]
        1. 생기부 내용은 무시하세요. {major} 학과와 관련된 고난도 딜레마 상황이나 철학적/학술적 주제가 담긴 **제시문 3개**를 도개고 학생들의 수능/구술 면접 수준에 맞춰 먼저 창작하세요[cite: 3, 4].
        2. **제시문 1개당 질문 2개씩** 연결하여 **총 3세트**(세트 1, 세트 2, 세트 3)를 구성하세요.
        3. 서론이나 인사말은 절대 쓰지 말고, 바로 '### 📄 [제시문 1]' 부터 출력하세요.
        
        [출력 템플릿 엄수 - 파싱을 위해 키워드 대괄호를 절대 변경하지 마세요]
        {TEMPLATE_JESIMUN}
        """
    
    with st.spinner(f"⏳ 로딩중... 도개고 면접 기출 수준에 맞춰 {interview_type} 문항을 조립하고 있습니다."):
        try:
            result_text = call_gemini(prompt, api_key)
            
            stu_path, tea_path = create_word_files(result_text, student_name, interview_type, target_desc)
            st.session_state.word_files = (stu_path, tea_path)
            
            display_text = result_text.replace('[질문 1]', '\n**💡 [질문 1]**\n').replace('[평가의도 1]', '\n**🎯 [평가 의도 1]**\n').replace('[모범답안 1]', '\n**✅ [모범 답안 가이드 1]**\n').replace('[꼬리질문 1]', '\n**🔥 [압박용 꼬리질문 1]**\n')
            display_text = display_text.replace('[질문 2]', '\n**💡 [질문 2]**\n').replace('[평가의도 2]', '\n**🎯 [평가 의도 2]**\n').replace('[모범답안 2]', '\n**✅ [모범 답안 가이드 2]**\n').replace('[꼬리질문 2]', '\n**🔥 [압박용 꼬리질문 2]**\n')
            display_text = display_text.replace('[질문]', '\n**💡 [면접 질문]**\n').replace('[평가의도]', '\n**🎯 [평가 의도]**\n').replace('[모범답안]', '\n**✅ [모범 답안 가이드]**\n').replace('[꼬리질문]', '\n**🔥 [압박용 꼬리질문]**\n')
            
            full_display_text = display_text + "\n\n---\n💬 **방금까지 나눈 문항 내용과 피드백 대화 내용을 한글 문서(.docx)로 만들어 드릴까요?** (원하시면 **'그래 만들어줘'**라고 말씀해 주세요!)"
            
            st.session_state.chat_history = [{"role": "assistant", "content": full_display_text}]
            st.success("🎉 도개고 맞춤형 면접 패키지 및 워드 문서 생성이 완료되었습니다!")
            
        except Exception as e:
            st.error(f"❌ 생성 실패: {e}")

# -------------------------------------------------------------------------
# [5] 결과 대시보드
# -------------------------------------------------------------------------
if st.session_state.chat_history:
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
            
    if user_feedback := st.chat_input("질문을 더 어렵게 하거나 특정 활동을 추가해 달라고 피드백을 남겨보세요."):
        st.session_state.chat_history.append({"role": "user", "content": user_feedback})
        with st.chat_message("user"):
            st.markdown(user_feedback)
            
        with st.chat_message("assistant"):
            with st.spinner("요청하신 피드백을 반영하여 처리 중입니다..."):
                
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
                    
                    # 사용자가 문항 추가를 요청한 경우 최소 2세트 이상 추가 생성 지시
                    add_instruction = ""
                    if any(w in user_feedback for w in ["더", "추가", "많이", "늘려", "또"]):
                        add_instruction = "\n(주의: 사용자가 문항 추가를 요청했으므로, **제시문 1개 + 질문 2개로 구성된 세트를 최소 2세트 이상 추가로** 더 생성해 주세요!)"
                    
                    feedback_prompt = f"""
                    당신은 도개고 맞춤형 면접 출제위원입니다. 이전 내용을 사용자의 피드백에 맞게 수정하되, 
                    도개고 기출 수준의 심도 있는 학술적/실천적 깊이를 유지하면서 **반드시 다음 템플릿 구조와 [키워드]를 토씨 하나 틀리지 말고 유지**해 주세요[cite: 3, 4].
                    {add_instruction}
                    
                    [강제 유지 템플릿]
                    {required_template}
                    
                    사용자 피드백: "{user_feedback}"
                    """
                    try:
                        new_result = call_gemini(feedback_prompt, api_key)
                        display_text = new_result.replace('[질문 1]', '\n**💡 [질문 1]**\n').replace('[평가의도 1]', '\n**🎯 [평가 의도 1]**\n').replace('[모범답안 1]', '\n**✅ [모범 답안 가이드 1]**\n').replace('[꼬리질문 1]', '\n**🔥 [압박용 꼬리질문 1]**\n')
                        display_text = display_text.replace('[질문 2]', '\n**💡 [질문 2]**\n').replace('[평가의도 2]', '\n**🎯 [평가 의도 2]**\n').replace('[모범답안 2]', '\n**✅ [모범 답안 가이드 2]**\n').replace('[꼬리질문 2]', '\n**🔥 [압박용 꼬리질문 2]**\n')
                        display_text = display_text.replace('[질문]', '\n**💡 [면접 질문]**\n').replace('[평가의도]', '\n**🎯 [평가 의도]**\n').replace('[모범답안]', '\n**✅ [모범 답안 가이드]**\n').replace('[꼬리질문]', '\n**🔥 [압박용 꼬리질문]**\n')
                        
                        full_response = display_text + "\n\n---\n💬 **방금까지 나눈 문항 내용과 피드백 대화 내용을 한글 문서(.docx)로 만들어 드릴까요?** (원하시면 **'그래 만들어줘'**라고 말씀해 주세요!)"
                        
                        st.markdown(full_response)
                        st.session_state.chat_history.append({"role": "assistant", "content": full_response})
                        
                        stu_path, tea_path = create_word_files(new_result, student_name, interview_type, target_desc)
                        st.session_state.word_files = (stu_path, tea_path)
                        st.rerun() 
                        
                    except Exception as e:
                        st.error(f"피드백 반영 중 오류가 발생했습니다: {e}")
