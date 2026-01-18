import streamlit as st
import pandas as pd
from google import genai
import io
import json
import openpyxl

# 페이지 설정
st.set_page_config(page_title="Excel Merger AI (Expert)", layout="wide")

st.title("📊 재무제표 대/중/소 계정 매핑 (Expert)")
st.markdown("업로드된 데이터를 **[대계정 > 중계정 > 소계정]** 체계로 자동 분류하고, 연도별로 매핑합니다.")
st.markdown("ℹ️ **K-IFRS/일반기업회계기준**을 참고하여 계정 과목의 위계를 자동으로 생성합니다.")

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
    
    # 모델 선택 (기본값 고정이나 필요시 변경 가능하도록 정보 표시)
    st.info("사용 모델: gemini-3-flash-preview")

    if not st.session_state.api_key:
        st.warning("먼저 API 키를 입력해주세요.")

# --- 정밀 파싱 함수 (숨김 항목 제외) ---
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
                merged_df = pd.concat(all_data, ignore_index=True)
                st.success(f"✅ 원본 데이터 병합 완료 ({len(merged_df)}행)")
                
                with st.expander("병합된 원본 데이터 확인"):
                    st.dataframe(merged_df)

                # 2. Gemini AI 분석 (대-중-소 분류 요청)
                progress_text.text("🤖 계정과목을 [대-중-소] 체계로 분류하고 연도별 데이터를 매핑 중입니다...")
                
                csv_data = merged_df.to_csv(index=False)
                # 컨텍스트가 큰 모델이므로 넉넉하게 보냄
                if len(csv_data) > 150000:
                    csv_data = csv_data[:150000] + "\n...(생략됨)"

                client = genai.Client(api_key=st.session_state.api_key)
                
                # --- [핵심 수정] 프롬프트: 계층 구조화 및 연도별 순차 매핑 ---
                prompt = f"""
                당신은 재무 회계 감사인(Financial Auditor)입니다. 
                제공된 원본 데이터를 분석하여 **완벽한 계층 구조(Hierarchy)**를 가진 재무제표를 작성하십시오.

                [작업 순서]
                1. **분류 (Classification):** 원본 데이터의 각 계정(Item)을 표준 회계 기준(K-IFRS 등)에 따라 **대계정(Major) - 중계정(Medium) - 소계정(Minor)**으로 분류하십시오.
                   - 대계정 예시: 자산, 부채, 자본, 매출, 비용
                   - 중계정 예시: 유동자산, 비유동부채, 판매비와관리비, 영업외수익 등
                   - 소계정 예시: (원본 데이터의 계정명, 예: 복리후생비, 미수금 등)
                2. **매핑 (Mapping):** 분류된 소계정을 기준으로, 데이터에 존재하는 모든 연도(Year)의 금액을 찾아 매핑하십시오.
                3. **순서 (Ordering):** 재무제표 표준 순서(자산 -> 부채 -> 자본 -> 매출 -> 비용)대로 데이터를 정렬할 준비를 하십시오.

                [강력한 제약사항]
                1. **절대 원본 계정을 생략하거나 통합(Summarize)하지 마십시오.** 모든 세부 항목이 '소계정'으로 나와야 합니다.
                2. 계정 분류가 불분명하면 가장 적절한 회계 계정으로 추론하여 분류하십시오.
                3. 금액은 정확히 집계하고, 해당 연도에 값이 없으면 0으로 표기하십시오.

                [출력 포맷]
                결과는 오직 **JSON 배열** 형태여야 합니다.
                JSON 구조:
                [
                  {{
                    "Major_Category": "비용",
                    "Medium_Category": "판매비와관리비",
                    "Minor_Category": "급여",
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
                    
                    # [후처리] 컬럼 정렬 및 계층별 정렬
                    # 1. 컬럼 순서 지정
                    fixed_cols = ['Major_Category', 'Medium_Category', 'Minor_Category']
                    year_cols = sorted([c for c in ai_df.columns if c not in fixed_cols])
                    final_cols = fixed_cols + year_cols
                    
                    # 존재하는 컬럼만 선택 (에러 방지)
                    final_cols = [c for c in final_cols if c in ai_df.columns]
                    ai_df = ai_df[final_cols]
                    
                    # 2. 대-중-소 순서로 행 정렬 (가나다 순이 아닌, 회계 표준 순서로 하려면 별도 매핑 필요하지만, 여기선 이름순+AI순서 의존)
                    # AI가 데이터를 순서대로 줬다면 그대로 쓰는 게 낫지만, 혹시 모르니 정렬 옵션 제공
                    ai_df = ai_df.sort_values(by=['Major_Category', 'Medium_Category', 'Minor_Category'])

                    st.subheader("🏆 계층형 상세 재무제표 결과")
                    st.dataframe(ai_df, use_container_width=True)
                    
                    # 4. 엑셀 다운로드
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        ai_df.to_excel(writer, sheet_name='Hierarchical_FS', index=False)
                        merged_df.to_excel(writer, sheet_name='Raw_Data', index=False)
                    
                    st.download_button(
                        label="📥 계층형 재무제표 엑셀 다운로드",
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
                st.warning("⚠️ 'gemini-3-flash-preview' 모델을 사용할 수 없습니다. 구글 정책에 따라 아직 공개되지 않았거나 API 접근 권한이 없을 수 있습니다. 코드를 열어 'gemini-1.5-flash' 등으로 변경해보세요.")
        finally:
            progress_text.empty()