import streamlit as st
import pandas as pd
from google import genai
import io
import json
import openpyxl

# 페이지 설정
st.set_page_config(page_title="Excel Merger AI", layout="wide")

st.title("📊 Excel Merger & AI Analyzer")
st.markdown("여러 엑셀 파일을 업로드하면 **하나로 합치고**, AI가 **연도별 비교표**를 만들어줍니다.")
st.markdown("ℹ️ **숨겨진 시트나 행은 자동으로 제외**하고, 보이는 데이터만 처리합니다.")

# --- [개선 1] API Key Session State 관리 (캐싱) ---
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    st.header("설정")
    # 입력란의 값을 session_state와 연동
    api_input = st.text_input(
        "Gemini API Key", 
        type="password", 
        placeholder="여기에 키를 입력하세요",
        value=st.session_state.api_key
    )
    
    # 입력된 값이 있으면 업데이트
    if api_input:
        st.session_state.api_key = api_input

    if not st.session_state.api_key:
        st.warning("먼저 API 키를 입력해주세요.")

# --- [개선 2] 정밀 파싱 함수 (숨김 처리 로직 포함) ---
def load_excel_visible_only(file):
    """
    엑셀 파일에서 숨겨진 시트와 숨겨진 행을 제외하고 데이터프레임으로 변환
    """
    # data_only=True: 수식이 아닌 계산된 값만 가져옴 (파싱 오류 방지)
    wb = openpyxl.load_workbook(file, data_only=True)
    all_dfs = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # 1. 숨겨진 시트 건너뛰기
        if ws.sheet_state == 'hidden' or ws.sheet_state == 'veryHidden':
            continue
        
        visible_data = []
        
        # 2. 행 단위로 순회하며 숨겨진 행 제외
        # iter_rows는 1부터 시작하는 인덱스를 사용
        for row_idx, row_cells in enumerate(ws.iter_rows(values_only=True), 1):
            # 행이 숨겨져 있는지 확인
            if ws.row_dimensions[row_idx].hidden:
                continue
            
            # 모든 값이 None인 빈 행은 제외 (선택사항, 파싱 깔끔하게 하기 위함)
            if not any(row_cells):
                continue
                
            visible_data.append(row_cells)
        
        # 데이터가 있다면 DataFrame 생성
        if visible_data:
            # 첫 번째 보이는 행을 헤더로 가정
            headers = visible_data[0]
            # 헤더가 중복되거나 None일 경우 처리
            clean_headers = [str(h) if h is not None else f"Unnamed_{i}" for i, h in enumerate(headers)]
            
            # 데이터프레임 생성 (헤더 다음 줄부터 데이터로 사용)
            df = pd.DataFrame(visible_data[1:], columns=clean_headers)
            
            # 출처 컬럼 추가
            df['Source'] = f"{file.name} - {sheet_name}"
            all_dfs.append(df)
            
    return all_dfs

# --- 메인 로직 ---
uploaded_files = st.file_uploader("엑셀 파일을 드래그하거나 선택하세요", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files and st.session_state.api_key:
    if st.button("데이터 병합 및 분석 시작"):
        all_data = []
        progress_text = st.empty()
        
        try:
            # --- 1단계: 정밀 파싱 로직 적용 ---
            progress_text.text("📂 엑셀 파일(숨김 항목 제외) 읽는 중...")
            
            for file in uploaded_files:
                # 위에서 만든 커스텀 함수 사용
                dfs = load_excel_visible_only(file)
                all_data.extend(dfs)
            
            if not all_data:
                st.error("처리할 데이터가 없습니다. (모든 시트가 비어있거나 숨겨져 있을 수 있습니다)")
            else:
                # 리스트에 모인 데이터프레임을 하나로 합침
                merged_df = pd.concat(all_data, ignore_index=True)
                st.success(f"✅ 총 {len(uploaded_files)}개 파일 병합 완료! ({len(merged_df)}행)")
                
                with st.expander("원본 병합 데이터 보기"):
                    st.dataframe(merged_df)

                # --- 2단계: Gemini AI 분석 ---
                progress_text.text("🤖 AI가 데이터를 분석하고 연도별로 정리하는 중...")
                
                # 데이터 전처리: 너무 크면 자르기
                csv_data = merged_df.to_csv(index=False)
                if len(csv_data) > 50000:
                    csv_data = csv_data[:50000] + "\n...(생략됨)"

                # Client 객체 생성
                client = genai.Client(api_key=st.session_state.api_key)
                
                prompt = f"""
                너는 데이터 분석 전문가야. 아래 CSV 데이터를 분석해서 "연도별 비교(Yearly Comparison)"가 가능한 표로 재구성해줘.
                
                [지시사항]
                1. 'Category'(구분)를 행으로, '2022', '2023', '2024' 등 연도를 열(Column)로 만들어라.
                2. 데이터 안에서 연도를 스스로 추론해서 배치해라.
                3. 숫자는 정확하게 집계하고, 값이 없으면 0으로 채워라.
                4. 결과는 오직 JSON 데이터만 출력해라. (마크다운 ```json 금지)
                5. JSON 형식: [ {{"Category": "매출", "2023": 100, "2024": 120}}, ... ]
                
                [데이터]:
                {csv_data}
                """
                
                response = client.models.generate_content(
                    model="gemini-3-flash-preview",
                    contents=prompt
                )
                
                # 결과 처리
                try:
                    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
                    ai_result_json = json.loads(cleaned_text)
                    ai_df = pd.DataFrame(ai_result_json)
                    
                    st.subheader("🏆 AI 연도별 비교 분석 결과")
                    st.dataframe(ai_df, use_container_width=True)
                    
                    # --- 3단계: 엑셀 다운로드 ---
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        ai_df.to_excel(writer, sheet_name='AI_Analysis', index=False)
                        merged_df.to_excel(writer, sheet_name='Raw_Data', index=False)
                    
                    st.download_button(
                        label="📥 분석 결과 엑셀 다운로드",
                        data=buffer.getvalue(),
                        file_name="merged_analysis_result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except json.JSONDecodeError:
                    st.error("AI 응답 변환 실패. 원본 텍스트:")
                    st.text_area("AI 응답", response.text)
                    
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
        finally:
            progress_text.empty()