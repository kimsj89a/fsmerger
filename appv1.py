import streamlit as st
import pandas as pd
from google import genai
import io
import json
import openpyxl

# 페이지 설정
st.set_page_config(page_title="Excel Merger AI (Expert)", layout="wide")

st.title("📊 재무제표 대/중/소 계정 매핑 (Expert)")
st.markdown("업로드된 데이터를 **[대계정 > 중계정 > 소계정]** 체계로 분류하고, **원본 순서**를 최대한 보존하여 매핑합니다.")
st.markdown("ℹ️ **자동 가나다순 정렬을 하지 않습니다.**")

# --- API Key Session State 관리 ---
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    st.header("설정")
    api_input = st.text_input(
        "Gemini API Key", 
        type="password", 
        placeholder="여기에 키를 입력하세요",
        value=st.session_state.api_key
    )
    if api_input:
        st.session_state.api_key = api_input
    
    st.info("사용 모델: gemini-3-flash-preview")

    if not st.session_state.api_key:
        st.warning("먼저 API 키를 입력해주세요.")

# --- 정밀 파싱 함수 ---
def load_excel_visible_only(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    all_dfs = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.sheet_state == 'hidden' or ws.sheet_state == 'veryHidden':
            continue
        
        visible_data = []
        for row_idx, row_cells in enumerate(ws.iter_rows(values_only=True), 1):
            if ws.row_dimensions[row_idx].hidden:
                continue
            if not any(row_cells):
                continue
            visible_data.append(row_cells)
        
        if visible_data:
            headers = visible_data[0]
            clean_headers = [str(h) if h is not None else f"Unnamed_{i}" for i, h in enumerate(headers)]
            df = pd.DataFrame(visible_data[1:], columns=clean_headers)
            df['Source'] = f"{file.name} - {sheet_name}"
            all_dfs.append(df)
            
    return all_dfs

# --- 메인 로직 ---
uploaded_files = st.file_uploader("엑셀 파일을 드래그하거나 선택하세요", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files and st.session_state.api_key:
    if st.button("계층형 재무제표 생성 시작"):
        all_data = []
        progress_text = st.empty()
        
        try:
            # 1. 파일 읽기
            progress_text.text("📂 엑셀 파일 데이터 추출 중...")
            for file in uploaded_files:
                dfs = load_excel_visible_only(file)
                all_data.extend(dfs)
            
            if not all_data:
                st.error("처리할 데이터가 없습니다.")
            else:
                # concat시 sort=False 옵션으로 순서 유지
                merged_df = pd.concat(all_data, ignore_index=True, sort=False)
                st.success(f"✅ 원본 데이터 병합 완료 ({len(merged_df)}행)")
                
                with st.expander("병합된 원본 데이터 확인"):
                    st.dataframe(merged_df)

                # 2. Gemini AI 분석
                progress_text.text("🤖 AI가 데이터 순서를 유지하며 계정 구조를 생성 중입니다...")
                
                csv_data = merged_df.to_csv(index=False)
                if len(csv_data) > 150000:
                    csv_data = csv_data[:150000] + "\n...(생략됨)"

                client = genai.Client(api_key=st.session_state.api_key)
                
                # --- [핵심 수정] 프롬프트: 정렬 금지 및 순서 보존 명령 ---
                prompt = f"""
                당신은 재무 회계 감사인(Financial Auditor)입니다. 
                제공된 원본 데이터를 분석하여 계층 구조(Hierarchy)를 가진 재무제표를 작성하십시오.

                [작업 순서]
                1. **분류 (Classification):** 각 계정을 [대계정(Major) - 중계정(Medium) - 소계정(Minor)]으로 분류하십시오.
                2. **매핑 (Mapping):** 분류된 소계정을 기준으로 연도별 금액을 매핑하십시오.
                3. **순서 보존 (Order Preservation):** - **절대 계정명(Minor_Category)을 가나다순(Alphabetical)으로 정렬하지 마십시오.**
                   - 가능한 한 입력 데이터(Input Data)의 행 순서를 유지하거나, 표준 재무제표 순서(자산 유동성 배열법 -> 부채 -> 자본 -> 매출 -> 비용)를 따르십시오.

                [강력한 제약사항]
                1. 원본 계정을 생략하거나 통합(Summarize)하지 마십시오.
                2. 금액은 정확히 집계하고, 값이 없으면 0으로 표기하십시오.

                [출력 포맷]
                결과는 오직 **JSON 배열** 형태여야 합니다.
                JSON 구조:
                [
                  {{
                    "Major_Category": "자산",
                    "Medium_Category": "유동자산",
                    "Minor_Category": "현금및현금성자산",
                    "2022": 50000,
                    "2023": 52000,
                    "2024": 55000
                  }},
                  ...
                ]

                [분석할 데이터]:
                {csv_data}
                """
                
                response = client.models.generate_content(
                    model="gemini-3-flash-preview", 
                    contents=prompt
                )
                
                # 3. 결과 처리
                try:
                    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
                    if "[" in cleaned_text and "]" in cleaned_text:
                        start_idx = cleaned_text.find("[")
                        end_idx = cleaned_text.rfind("]") + 1
                        cleaned_text = cleaned_text[start_idx:end_idx]

                    ai_result_json = json.loads(cleaned_text)
                    ai_df = pd.DataFrame(ai_result_json)
                    
                    # [수정] 강제 정렬 코드(sort_values)를 삭제했습니다.
                    # AI가 뱉어준 순서(JSON 리스트 순서) 그대로 출력합니다.
                    
                    # 컬럼 순서만 정리 (대-중-소, 그 뒤에 연도)
                    fixed_cols = ['Major_Category', 'Medium_Category', 'Minor_Category']
                    year_cols = sorted([c for c in ai_df.columns if c not in fixed_cols])
                    final_cols = fixed_cols + year_cols
                    final_cols = [c for c in final_cols if c in ai_df.columns]
                    
                    ai_df = ai_df[final_cols]

                    st.subheader("🏆 계층형 상세 재무제표 (순서 보존)")
                    st.dataframe(ai_df, use_container_width=True)
                    
                    # 4. 엑셀 다운로드
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        ai_df.to_excel(writer, sheet_name='Hierarchical_FS', index=False)
                        merged_df.to_excel(writer, sheet_name='Raw_Data', index=False)
                    
                    st.download_button(
                        label="📥 엑셀 다운로드",
                        data=buffer.getvalue(),
                        file_name="hierarchical_financial_statements.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except json.JSONDecodeError:
                    st.error("결과 변환 중 오류가 발생했습니다. AI 응답 원본을 확인해주세요.")
                    st.text_area("AI Raw Response", response.text, height=300)
                    
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            if "404" in str(e):
                st.warning("⚠️ 모델을 찾을 수 없습니다. (gemini-3-flash-preview). 코드에서 모델명을 'gemini-1.5-flash'로 변경해보세요.")
        finally:
            progress_text.empty()