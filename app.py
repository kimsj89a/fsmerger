# app.py
import streamlit as st
import pandas as pd
import io
import logic 
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Financial Report AI", layout="wide")

# --- CSS: 파일 목록 박스 스타일 ---
st.markdown("""
    <style>
        .file-list-box {
            border: 1px solid #e6e6e6;
            padding: 10px;
            border-radius: 5px;
            max-height: 200px;
            overflow-y: auto;
            background-color: #f9f9f9;
            margin-bottom: 20px;
        }
        .file-item {
            font-size: 0.9em;
            margin-bottom: 4px;
            padding: 4px;
            background: white;
            border-radius: 3px;
        }
    </style>
""", unsafe_allow_html=True)

# --- 사이드바 설정 ---
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    is_expanded = not bool(st.session_state.api_key)
    with st.expander("🔑 API Key 설정", expanded=is_expanded):
        api_input = st.text_input(
            "Gemini API Key", 
            type="password", 
            value=st.session_state.api_key,
            label_visibility="collapsed"
        )
        if api_input:
            st.session_state.api_key = api_input
    
    st.divider()
    
    # [추가] 단위 선택기
    st.markdown("### 📐 단위 설정")
    unit_option = st.selectbox(
        "출력 단위를 선택하세요",
        ("원", "천원", "백만원", "억원"),
        index=0
    )
    
    # 단위별 나누기 값
    unit_divisors = {
        "원": 1,
        "천원": 1000,
        "백만원": 1000000,
        "억원": 100000000
    }
    divisor = unit_divisors[unit_option]

st.title("📑 통합 재무제표 보고서")

# --- 스타일 함수 ---
def style_dataframe(row):
    # 배경색 스타일
    styles = [''] * len(row)
    level = row.get('Level', 3)
    if level == 1:
        return ['background-color: #1f77b4; color: white; font-weight: bold;'] * len(row)
    elif level == 2:
        return ['background-color: #aec7e8; color: black; font-weight: bold;'] * len(row)
    return ['color: black;'] * len(row)

# --- 엑셀 저장 함수 (단위 포함) ---
def save_styled_excel(df, sheet_name_map, unit_text):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        if 'Statement' in df.columns:
            statements = df['Statement'].unique()
        else:
            statements = ['Result']
            
        for stmt in statements:
            if 'Statement' in df.columns:
                sub_df = df[df['Statement'] == stmt].copy()
            else:
                sub_df = df.copy()
            
            # 저장할 컬럼
            cols = [c for c in sub_df.columns if c not in ['Statement', 'Level']]
            if 'Account_Name' in cols:
                cols.remove('Account_Name')
                cols = ['Account_Name'] + cols
            
            sheet_title = sheet_name_map.get(stmt, stmt)[:30]
            
            # 1행부터 데이터 쓰기 (0행은 단위 표시용으로 비움)
            sub_df[cols].to_excel(writer, sheet_name=sheet_title, index=False, startrow=1)
            
            ws = writer.sheets[sheet_title]
            
            # [추가] 좌측 상단에 단위 표시
            ws['A1'] = f"(단위: {unit_text})"
            ws['A1'].font = Font(bold=True, italic=True)
            
            # 스타일링
            fill_lv1 = PatternFill(start_color="1F77B4", end_color="1F77B4", fill_type="solid")
            font_lv1 = Font(color="FFFFFF", bold=True)
            fill_lv2 = PatternFill(start_color="AEC7E8", end_color="AEC7E8", fill_type="solid")
            font_lv2 = Font(color="000000", bold=True)
            
            # 숫자가 있는 데이터 셀 포맷 (천단위 콤마, 소수점 없음)
            # cols 리스트에서 날짜 컬럼들의 위치를 파악
            date_col_indices = [i+1 for i, c in enumerate(cols) if c != 'Account_Name'] # 1-based index
            
            sub_df = sub_df.reset_index(drop=True)
            for idx, row in sub_df.iterrows():
                excel_row = idx + 3 # 헤더가 2행(startrow=1)이므로 데이터는 3행부터
                level = row.get('Level', 3)
                
                # 행 스타일 적용
                for col_idx in range(1, len(cols) + 1):
                    cell = ws.cell(row=excel_row, column=col_idx)
                    
                    if level == 1:
                        cell.fill = fill_lv1
                        cell.font = font_lv1
                    elif level == 2:
                        cell.fill = fill_lv2
                        cell.font = font_lv2
                        
                    # 숫자 컬럼이면 포맷 적용
                    if col_idx - 1 in date_col_indices: # -1 to match list index
                        cell.number_format = '#,##0' # 정수 포맷

            ws.column_dimensions['A'].width = 30

    return buffer

# --- 메인 로직 ---
uploaded_files = st.file_uploader(
    "파일 업로드 (Excel, PDF, Word, CSV, TXT)", 
    accept_multiple_files=True, 
    type=['xlsx', 'xls', 'csv', 'pdf', 'docx', 'txt']
)

# [UI 개선] 업로드된 파일 목록 커스텀 뷰어
if uploaded_files:
    st.markdown(f"##### 📂 업로드된 파일 목록 ({len(uploaded_files)}개)")
    file_list_html = '<div class="file-list-box">'
    for f in uploaded_files:
        size_kb = f.size / 1024
        file_list_html += f'<div class="file-item">📄 {f.name} ({size_kb:.1f} KB)</div>'
    file_list_html += '</div>'
    st.markdown(file_list_html, unsafe_allow_html=True)

if uploaded_files and st.session_state.api_key:
    if st.button("보고서 생성 시작", type="primary"):
        status = st.status("파일 분석 및 통합 중...", expanded=True)
        
        try:
            # 1. 로직 실행
            raw_df = logic.process_smart_merge(st.session_state.api_key, uploaded_files)
            
            # 2. 단위 변환 (숫자 컬럼만)
            df = raw_df.copy()
            numeric_cols = []
            for col in df.columns:
                if col not in ['Statement', 'Level', 'Account_Name']:
                    # 숫자로 변환 시도
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    if divisor > 1:
                        df[col] = df[col] / divisor
                    numeric_cols.append(col)
            
            status.update(label="✅ 생성 완료!", state="complete", expanded=False)

            # 3. 탭 생성
            available_types = df['Statement'].unique() if 'Statement' in df.columns else []
            type_map = {'BS': '재무상태표', 'IS': '손익계산서', 'COGM': '제조원가명세서', 'CF': '현금흐름표', 'Other': '기타'}
            tabs = st.tabs([type_map.get(t, t) for t in available_types])

            for i, stmt_type in enumerate(available_types):
                with tabs[i]:
                    sub_df = df[df['Statement'] == stmt_type].copy()
                    display_cols = [c for c in sub_df.columns if c not in ['Statement', 'Level']]
                    
                    if 'Account_Name' in display_cols:
                        display_cols.remove('Account_Name')
                        display_cols = ['Account_Name'] + display_cols
                    
                    # [UI 개선] 숫자 포맷팅 (천단위 콤마, 소수점 제거)
                    format_dict = {col: "{:,.0f}" for col in numeric_cols}
                    
                    st.dataframe(
                        sub_df[display_cols].style
                        .apply(style_dataframe, axis=1)
                        .format(format_dict), # 포맷 적용
                        use_container_width=True,
                        height=600
                    )

            # 엑셀 다운로드
            excel_buffer = save_styled_excel(df, type_map, unit_option)
            
            st.download_button(
                f"📥 엑셀 다운로드 (단위: {unit_option})",
                data=excel_buffer.getvalue(),
                file_name=f"Financial_Report_{unit_option}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            status.update(label="❌ 오류 발생", state="error")
            st.error(f"에러 내용: {e}")