import streamlit as st
import pandas as pd
import io
import logic 

st.set_page_config(page_title="Financial Report AI", layout="wide")

st.title("📑 통합 재무제표 보고서 (Smart Format)")
st.markdown("재무제표별로 **탭(Tab)**을 나누고, **계층 구조(들여쓰기)**를 적용하여 보여줍니다.")

# 스타일링 함수 (Level에 따라 배경색 지정)
def highlight_levels(row):
    color = ''
    if row.get('Level') == 1:
        color = 'background-color: #e6f3ff; font-weight: bold; color: #000000;' # 대분류: 파란 배경, 굵게
    elif row.get('Level') == 2:
        color = 'background-color: #ffffff; font-weight: bold; color: #333333;' # 중분류: 흰 배경, 굵게
    else:
        color = 'color: #666666;' # 소분류: 회색 글자
    return [color] * len(row)

if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    st.header("설정")
    api_key = st.text_input("Gemini API Key", type="password", value=st.session_state.api_key)
    if api_key:
        st.session_state.api_key = api_key

uploaded_files = st.file_uploader("연도별 엑셀 파일 업로드", accept_multiple_files=True, type=['xlsx'])

if uploaded_files and st.session_state.api_key:
    if st.button("보고서 생성 시작"):
        status = st.status("AI가 재무제표를 분류하고 서식을 적용 중입니다...", expanded=True)
        
        try:
            # 1. 로직 실행
            df = logic.process_smart_merge(st.session_state.api_key, uploaded_files)
            status.update(label="✅ 생성 완료!", state="complete", expanded=False)

            # 2. 탭 생성 (재무제표 종류별)
            # 데이터에 있는 Statement 종류를 찾음 (BS, IS 등)
            available_types = df['Statement'].unique() if 'Statement' in df.columns else []
            
            # 탭 이름 매핑 (영문 -> 한글)
            type_map = {
                'BS': '재무상태표 (BS)', 
                'IS': '손익계산서 (IS)', 
                'COGM': '제조원가명세서', 
                'CF': '현금흐름표',
                'Unknown': '기타'
            }
            
            # 존재하는 탭만 생성
            tabs = st.tabs([type_map.get(t, t) for t in available_types])

            # 3. 각 탭에 데이터 뿌리기
            for i, stmt_type in enumerate(available_types):
                with tabs[i]:
                    # 해당 재무제표 데이터 필터링
                    sub_df = df[df['Statement'] == stmt_type].copy()
                    
                    # 화면에 보여줄 컬럼 정리 (Account_Name 대신 들여쓰기 된 Display_Name 사용)
                    display_cols = ['Display_Name'] + [c for c in sub_df.columns if c.isdigit()] # 연도 컬럼(숫자)만 가져옴
                    
                    # 데이터프레임 스타일 적용
                    st.dataframe(
                        sub_df[display_cols].style.apply(highlight_levels, axis=1),
                        use_container_width=True,
                        height=600
                    )

            # 4. 통합 엑셀 다운로드
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # 탭별로 시트 나눠서 저장
                for stmt_type in available_types:
                    sheet_name = type_map.get(stmt_type, stmt_type)[:30] # 시트명 길이 제한
                    sub_df = df[df['Statement'] == stmt_type]
                    
                    # 엑셀에는 'Display_Name' (들여쓰기 된 이름)과 연도 데이터만 저장
                    save_cols = ['Display_Name'] + [c for c in sub_df.columns if c.isdigit()]
                    sub_df[save_cols].to_excel(writer, sheet_name=sheet_name, index=False)
                    
            st.download_button(
                "📥 서식이 적용된 엑셀 다운로드",
                data=buffer.getvalue(),
                file_name="Formatted_Financial_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            status.update(label="❌ 오류 발생", state="error")
            st.error(f"에러 내용: {e}")