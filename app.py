# app.py
import streamlit as st
import pandas as pd
import io
import logic  # <--- 우리가 만든 logic.py를 가져옵니다

# 페이지 설정
st.set_page_config(page_title="Standard Taxonomy Mapper (Modular)", layout="wide")
st.title("📊 표준 택소노미 기반 재무제표 매핑")
st.markdown("내장된 **2018 Taxonomy**를 기준으로 데이터를 매핑합니다. (로직 분리 버전)")

# --- 설정 ---
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    st.header("설정")
    api_input = st.text_input("Gemini API Key", type="password", value=st.session_state.api_key)
    if api_input:
        st.session_state.api_key = api_input
    
    st.info("Logic Module Loaded")

# --- 메인 실행 ---
target_files = st.file_uploader("분석할 엑셀 파일 업로드", accept_multiple_files=True, type=['xlsx'])

if target_files and st.session_state.api_key:
    if st.button("매핑 시작"):
        # UI용 컨테이너
        status = st.status("작업 진행 중...", expanded=True)
        
        try:
            # 1. 로직 호출 (모든 복잡한 처리는 logic.py가 담당)
            status.write("📂 파일 읽기 및 AI 분석 요청 중...")
            
            # logic.py의 함수 실행
            result_df = logic.process_financial_mapping(
                api_key=st.session_state.api_key,
                target_files=target_files
            )
            
            status.update(label="✅ 작업 완료!", state="complete", expanded=False)

            # 2. 결과 표시
            st.subheader("🏆 매핑 결과")
            st.dataframe(result_df, use_container_width=True)

            # 3. 다운로드
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False)
            
            st.download_button(
                "📥 엑셀 다운로드",
                data=buffer.getvalue(),
                file_name="mapped_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except FileNotFoundError as e:
            status.update(label="🚨 파일 에러", state="error")
            st.error(str(e))
        except Exception as e:
            status.update(label="🚨 실행 에러", state="error")
            st.error(f"오류가 발생했습니다: {e}")