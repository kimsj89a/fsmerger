import streamlit as st
import pandas as pd
import io
import logic # logic.py 임포트

st.set_page_config(page_title="Context-Aware Merger", layout="wide")

st.title("🔗 문맥 기반 재무제표 병합 (Smart Merge)")
st.markdown("""
**순서 보존 병합:** 가나다순 정렬이 아닙니다.  
2022년엔 없고 2023년에만 생긴 계정이 있다면, **2023년의 위치(문맥)를 파악해 2022년 목록 사이사이에 끼워넣습니다.**
""")

# API 키 설정
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    st.header("설정")
    api_key = st.text_input("Gemini API Key", type="password", value=st.session_state.api_key)
    if api_key:
        st.session_state.api_key = api_key

# 파일 업로드
uploaded_files = st.file_uploader("연도별 엑셀 파일을 모두 선택하세요", accept_multiple_files=True, type=['xlsx'])

if uploaded_files and st.session_state.api_key:
    if st.button("스마트 병합 시작"):
        status = st.status("작업 진행 중...", expanded=True)
        
        try:
            status.write("🧠 AI가 파일들의 흐름을 분석하고 있습니다...")
            
            # logic.py의 스마트 병합 함수 호출
            merged_df = logic.process_smart_merge(
                api_key=st.session_state.api_key,
                target_files=uploaded_files
            )
            
            status.update(label="✅ 병합 완료!", state="complete", expanded=False)
            
            st.subheader("📊 병합 결과")
            st.dataframe(merged_df, use_container_width=True)
            
            # 다운로드
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False, sheet_name="Smart_Merged")
                
            st.download_button(
                "📥 엑셀로 다운로드",
                data=buffer.getvalue(),
                file_name="smart_merged_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            status.update(label="❌ 오류 발생", state="error")
            st.error(f"에러 내용: {e}")