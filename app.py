import streamlit as st
import pandas as pd
from google import genai
import io
import json
import openpyxl
import os

# 페이지 설정
st.set_page_config(page_title="Standard Taxonomy Mapper (Internal)", layout="wide")

st.title("📊 표준 택소노미(Taxonomy) 기반 재무제표 매핑")
st.markdown("내장된 **2018 표준 Taxonomy**를 기준으로, 업로드한 데이터를 자동으로 분류하고 정렬합니다.")

# --- API Key 관리 ---
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    st.header("설정")
    api_input = st.text_input(
        "Gemini API Key", 
        type="password", 
        value=st.session_state.api_key
    )
    if api_input:
        st.session_state.api_key = api_input
    
    st.info("사용 모델: gemini-3-flash-preview")

# --- [핵심] Taxonomy 내부 파일 로딩 (캐싱 적용) ---
@st.cache_data
def load_internal_taxonomy():
    """
    프로젝트 폴더 내의 '2018taxonomy.xlsx'를 읽어서 텍스트 컨텍스트로 변환
    @st.cache_data를 사용하여 한 번만 읽고 메모리에 저장 (속도 향상)
    """
    file_path = '2018taxonomy.xlsx'
    
    if not os.path.exists(file_path):
        return None

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        all_text_data = []
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if ws.sheet_state == 'hidden' or ws.sheet_state == 'veryHidden':
                continue
            
            # 데이터프레임 변환
            data = ws.values
            try:
                columns = next(data)[0:]
            except StopIteration:
                continue # 빈 시트
                
            df = pd.DataFrame(data, columns=columns)
            
            # CSV 텍스트로 변환
            sheet_csv = df.to_csv(index=False)
            all_text_data.append(f"--- Standard Sheet: {sheet_name} ---\n{sheet_csv}")
            
        return "\n".join(all_text_data)
    except Exception as e:
        st.error(f"Taxonomy 파일 로딩 중 에러: {e}")
        return None

# --- 일반 파일 로딩 함수 ---
def load_target_excel(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    all_text = []
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.sheet_state == 'hidden' or ws.sheet_state == 'veryHidden':
            continue
        
        data = ws.values
        try:
            columns = next(data)[0:]
            df = pd.DataFrame(data, columns=columns)
            all_text.append(f"--- User Data Sheet: {sheet_name} ---\n{df.to_csv(index=False)}")
        except:
            pass
    return "\n".join(all_text)

# --- 메인 로직 ---

# 1. Taxonomy 로드 (자동)
taxonomy_context = load_internal_taxonomy()

if taxonomy_context is None:
    st.error("🚨 **'2018taxonomy.xlsx' 파일을 찾을 수 없습니다!**")
    st.warning("프로젝트 폴더(app.py와 같은 위치)에 엑셀 파일이 있는지 확인하고 GitHub에 올려주세요.")
else:
    st.success("✅ 표준 Taxonomy 데이터 로드 완료")

    # 2. 분석할 파일 업로드
    st.subheader("분석할 재무 데이터 업로드")
    target_files = st.file_uploader("합치고 싶은 엑셀 파일들을 선택하세요", accept_multiple_files=True, type=['xlsx'])

    if target_files and st.session_state.api_key:
        if st.button("표준 양식으로 매핑 시작"):
            status_container = st.container()
            
            try:
                # 타겟 데이터 처리
                target_context_list = []
                with status_container:
                    st.info("📂 업로드된 데이터를 분석 중...")
                    for t_file in target_files:
                        t_context = load_target_excel(t_file)
                        target_context_list.append(t_context)
                
                full_target_context = "\n".join(target_context_list)
                if len(full_target_context) > 100000:
                    full_target_context = full_target_context[:100000] + "\n...(Data Truncated)"

                # AI 요청
                with status_container:
                    st.info("🤖 AI가 표준 Taxonomy에 맞춰 데이터를 끼워 맞추는 중입니다...")
                    
                    client = genai.Client(api_key=st.session_state.api_key)
                    
                    prompt = f"""
                    [Role]
                    당신은 회계 데이터 매핑 시스템입니다. 
                    사용자의 [User Data]를 [Standard Taxonomy]의 구조에 강제로 일치시켜야 합니다.

                    [Input 1: Standard Taxonomy (기준)]
                    이것은 변경할 수 없는 기준입니다.
                    {taxonomy_context}

                    [Input 2: User Data (분석 대상)]
                    {full_target_context}

                    [Mapping Rules]
                    1. **Strict Hierarchy:** 결과의 'Major', 'Medium', 'Account' 컬럼은 오직 [Standard Taxonomy]에 존재하는 명칭만 사용하십시오.
                    2. **Mapping:** User Data의 계정 항목을 가장 의미가 비슷한 Standard Taxonomy 항목에 합산하십시오.
                    3. **Columns:** 연도(2022, 2023 등)는 컬럼으로 분리하십시오.
                    
                    [Output Format]
                    JSON Array Only.
                    [
                        {{
                            "Standard_Major": "자산",
                            "Standard_Medium": "유동자산",
                            "Standard_Account": "현금및현금성자산",
                            "Original_Account_Map": "현금, 보통예금 (매핑된 원본 계정명들)",
                            "2022": 15000,
                            "2023": 20000
                        }},
                        ...
                    ]
                    """

                    response = client.models.generate_content(
                        model="gemini-3-flash-preview",
                        contents=prompt
                    )

                    # 결과 파싱
                    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
                    if "[" in cleaned_text and "]" in cleaned_text:
                        s = cleaned_text.find("[")
                        e = cleaned_text.rfind("]") + 1
                        cleaned_text = cleaned_text[s:e]
                    
                    result_data = json.loads(cleaned_text)
                    result_df = pd.DataFrame(result_data)

                    # 컬럼 정렬 (표준 계정 먼저)
                    cols = result_df.columns.tolist()
                    std_cols = ['Standard_Major', 'Standard_Medium', 'Standard_Account', 'Original_Account_Map']
                    other_cols = [c for c in cols if c not in std_cols]
                    result_df = result_df[std_cols + other_cols]

                    st.success("매핑 완료!")
                    st.subheader("🏆 표준 Taxonomy 매핑 결과")
                    st.dataframe(result_df, use_container_width=True)

                    # 다운로드
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        result_df.to_excel(writer, sheet_name='Mapped_Result', index=False)
                    
                    st.download_button(
                        "📥 결과 엑셀 다운로드",
                        data=buffer.getvalue(),
                        file_name="standardized_financial_statement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"오류 발생: {e}")
                if 'response' in locals():
                    st.expander("오류 상세(AI 응답)").text(response.text)