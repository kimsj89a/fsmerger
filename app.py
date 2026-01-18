import streamlit as st
import pandas as pd
# [신버전] google-genai 라이브러리 import
from google import genai
import io
import json

# 페이지 설정
st.set_page_config(page_title="Excel Merger AI", layout="wide")

st.title("📊 Excel Merger & AI Analyzer")
st.markdown("여러 엑셀 파일을 업로드하면 **하나로 합치고**, AI가 **연도별 비교표**를 만들어줍니다.")

# 사이드바: API 키 입력
with st.sidebar:
    st.header("설정")
    api_key = st.text_input("Gemini API Key", type="password", placeholder="여기에 키를 입력하세요")
    if not api_key:
        st.warning("먼저 API 키를 입력해주세요.")

# 1. 파일 업로드
uploaded_files = st.file_uploader("엑셀 파일을 드래그하거나 선택하세요", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files and api_key:
    if st.button("데이터 병합 및 분석 시작"):
        all_data = []
        progress_text = st.empty()
        
        # --- 1단계: Pandas로 엑셀 읽기 및 병합 ---
        try:
            progress_text.text("📂 파일 읽는 중...")
            for file in uploaded_files:
                xls = pd.read_excel(file, sheet_name=None)
                for sheet_name, df in xls.items():
                    df['Source'] = f"{file.name} - {sheet_name}"
                    all_data.append(df)
            
            merged_df = pd.concat(all_data, ignore_index=True)
            st.success(f"✅ 총 {len(uploaded_files)}개 파일 병합 완료! ({len(merged_df)}행)")
            
            with st.expander("원본 병합 데이터 보기"):
                st.dataframe(merged_df)

            # --- 2단계: Gemini AI 분석 (신버전 SDK 적용) ---
            progress_text.text("🤖 AI가 데이터를 분석하고 연도별로 정리하는 중...")
            
            csv_data = merged_df.to_csv(index=False)
            if len(csv_data) > 50000:
                csv_data = csv_data[:50000] + "\n...(생략됨)"

            # [핵심 변경] Client 객체 생성 (신버전 방식)
            client = genai.Client(api_key=api_key)
            
            prompt = f"""
            너는 데이터 분석 전문가야. 아래 CSV 데이터를 분석해서 "연도별 비교(Yearly Comparison)"가 가능한 표로 재구성해줘.
            [지시사항]
            1. 'Category'(구분)를 행으로, '2022', '2023', '2024' 등 연도를 열(Column)로 만들어라.
            2. 결과는 오직 JSON 데이터만 출력해라. (마크다운 ```json 금지)
            3. JSON 형식: [ {{"Category": "매출", "2023": 100, "2024": 120}}, ... ]
            
            [데이터]:
            {csv_data}
            """
            
            # [핵심 변경] generate_content 호출 방식 변경
            response = client.models.generate_content(
                model="gemini-1.5-flash",
                contents=prompt
            )
            
            # 결과 처리
            try:
                # 신버전 SDK에서도 response.text로 텍스트 접근 가능
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