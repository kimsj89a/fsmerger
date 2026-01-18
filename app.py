# app.py
import streamlit as st
import pandas as pd
import io
import logic 
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Financial Report AI", layout="wide")
st.title("📑 통합 재무제표 보고서 (Pro Format)")
st.markdown("다양한 파일(PDF, Word, Excel 등)을 지원하며, **분기(3Q) 데이터**도 포함합니다.")

# --- 화면용 스타일 함수 (Pandas Styler) ---
def style_dataframe(row):
    # Level에 따른 CSS 스타일 지정
    styles = [''] * len(row)
    level = row.get('Level', 3)
    
    if level == 1:
        # Level 1: 진한 파랑 배경, 흰 글씨, 굵게
        return ['background-color: #1f77b4; color: white; font-weight: bold;'] * len(row)
    elif level == 2:
        # Level 2: 연한 하늘색 배경, 굵게
        return ['background-color: #aec7e8; color: black; font-weight: bold;'] * len(row)
    else:
        # Level 3: 기본 흰 배경
        return ['color: black;'] * len(row)

# --- 엑셀 파일 스타일링 저장 함수 ---
def save_styled_excel(df, sheet_name_map):
    buffer = io.BytesIO()
    
    # Pandas로 먼저 데이터를 씁니다 (Engine: openpyxl)
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Statement별로 시트 나누기
        if 'Statement' in df.columns:
            statements = df['Statement'].unique()
        else:
            statements = ['Result']
            
        for stmt in statements:
            if 'Statement' in df.columns:
                sub_df = df[df['Statement'] == stmt].copy()
            else:
                sub_df = df.copy()
            
            # 저장할 컬럼: Account_Name + 날짜 컬럼들 (Level, Statement 제외)
            cols = [c for c in sub_df.columns if c not in ['Statement', 'Level']]
            # Account_Name을 맨 앞으로
            if 'Account_Name' in cols:
                cols.remove('Account_Name')
                cols = ['Account_Name'] + cols
            
            sheet_title = sheet_name_map.get(stmt, stmt)[:30]
            sub_df[cols].to_excel(writer, sheet_name=sheet_title, index=False)
            
            # --- 엑셀 스타일링 적용 ---
            workbook = writer.book
            worksheet = writer.sheets[sheet_title]
            
            # 스타일 정의
            fill_lv1 = PatternFill(start_color="1F77B4", end_color="1F77B4", fill_type="solid") # 파랑
            font_lv1 = Font(color="FFFFFF", bold=True)
            
            fill_lv2 = PatternFill(start_color="AEC7E8", end_color="AEC7E8", fill_type="solid") # 연하늘
            font_lv2 = Font(color="000000", bold=True)
            
            # 데이터 행 순회하며 스타일 적용
            # sub_df의 인덱스가 섞여있을 수 있으므로 reset_index
            sub_df = sub_df.reset_index(drop=True)
            
            for idx, row in sub_df.iterrows():
                excel_row = idx + 2 # 헤더가 1행이므로 데이터는 2행부터
                level = row.get('Level', 3)
                
                # 행 전체에 스타일 적용
                if level == 1:
                    for col in range(1, len(cols) + 1):
                        cell = worksheet.cell(row=excel_row, column=col)
                        cell.fill = fill_lv1
                        cell.font = font_lv1
                elif level == 2:
                    for col in range(1, len(cols) + 1):
                        cell = worksheet.cell(row=excel_row, column=col)
                        cell.fill = fill_lv2
                        cell.font = font_lv2
                
            # 컬럼 너비 자동 조정 (대략적)
            worksheet.column_dimensions['A'].width = 30

    return buffer

# --- 메인 로직 ---
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    st.header("설정")
    api_key = st.text_input("Gemini API Key", type="password", value=st.session_state.api_key)
    if api_key:
        st.session_state.api_key = api_key

# 1. 파일 업로드 확장 (xls, pdf, word, txt, csv 추가)
uploaded_files = st.file_uploader(
    "파일 업로드 (Excel, PDF, Word, CSV, TXT)", 
    accept_multiple_files=True, 
    type=['xlsx', 'xls', 'csv', 'pdf', 'docx', 'txt']
)

if uploaded_files and st.session_state.api_key:
    if st.button("보고서 생성 시작"):
        status = st.status("파일 분석 및 통합 중...", expanded=True)
        
        try:
            # 로직 실행
            df = logic.process_smart_merge(st.session_state.api_key, uploaded_files)
            status.update(label="✅ 생성 완료!", state="complete", expanded=False)

            # 탭 생성
            available_types = df['Statement'].unique() if 'Statement' in df.columns else []
            type_map = {
                'BS': '재무상태표', 'IS': '손익계산서', 
                'COGM': '제조원가명세서', 'CF': '현금흐름표', 'Other': '기타'
            }
            tabs = st.tabs([type_map.get(t, t) for t in available_types])

            for i, stmt_type in enumerate(available_types):
                with tabs[i]:
                    sub_df = df[df['Statement'] == stmt_type].copy()
                    
                    # 화면 표시용 컬럼 (Level, Statement 숨김)
                    display_cols = [c for c in sub_df.columns if c not in ['Statement', 'Level']]
                    # Account_Name 맨 앞으로
                    if 'Account_Name' in display_cols:
                        display_cols.remove('Account_Name')
                        display_cols = ['Account_Name'] + display_cols
                    
                    # 화면 스타일 적용
                    st.dataframe(
                        sub_df[display_cols].style.apply(style_dataframe, axis=1),
                        use_container_width=True,
                        height=600
                    )

            # 엑셀 다운로드 (스타일 적용된 버전)
            excel_buffer = save_styled_excel(df, type_map)
            
            st.download_button(
                "📥 스타일 적용된 엑셀 다운로드",
                data=excel_buffer.getvalue(),
                file_name="Formatted_Financial_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            status.update(label="❌ 오류 발생", state="error")
            st.error(f"에러 내용: {e}")