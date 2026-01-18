# app.py
import streamlit as st
import pandas as pd
import io
import re  # 정규표현식 (날짜 추출용)
import logic 
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Financial Report AI", layout="wide")

# --- CSS ---
st.markdown("""
    <style>
        .file-list-box {
            border: 1px solid #e6e6e6; padding: 10px; border-radius: 5px;
            max-height: 200px; overflow-y: auto; background-color: #f9f9f9; margin-bottom: 20px;
        }
        .file-item {
            font-size: 0.9em; margin-bottom: 4px; padding: 4px; background: white; border-radius: 3px;
        }
    </style>
""", unsafe_allow_html=True)

# --- [추가] 컬럼 연도순 정렬 함수 ---
def sort_columns_chronologically(columns):
    """
    컬럼 리스트를 받아서 [Account_Name, 2022, 2023, 2024, 2025.1Q ...] 순서로 정렬
    """
    # 고정 컬럼 (앞부분)
    fixed_cols = ['Account_Name']
    
    # 날짜 컬럼만 추출
    date_cols = [c for c in columns if c not in ['Statement', 'Level', 'Account_Name']]
    
    def date_sort_key(col_name):
        # 1. 연도 추출 (4자리 숫자)
        year_match = re.search(r'(\d{4})', str(col_name))
        year = int(year_match.group(1)) if year_match else 9999
        
        # 2. 분기/월 추출 (없으면 0)
        # 3Q, 12M 등의 숫자를 찾아서 보조 정렬 키로 사용
        sub_match = re.search(r'(\d+)[QM]', str(col_name))
        sub_val = int(sub_match.group(1)) if sub_match else 0
        
        # 3. 누적/3개월 구분 (누적이 보통 뒤에 옴)
        is_cum = 1 if '누적' in str(col_name) or 'Cum' in str(col_name) else 0
        
        return (year, sub_val, is_cum, col_name)
    
    # 정렬 실행
    sorted_date_cols = sorted(date_cols, key=date_sort_key)
    
    return fixed_cols + sorted_date_cols

# --- 사이드바 ---
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    is_expanded = not bool(st.session_state.api_key)
    with st.expander("🔑 API Key 설정", expanded=is_expanded):
        api_input = st.text_input("Gemini API Key", type="password", value=st.session_state.api_key, label_visibility="collapsed")
        if api_input: st.session_state.api_key = api_input
    
    st.divider()
    st.markdown("### 📐 단위 설정")
    unit_option = st.selectbox("출력 단위를 선택하세요", ("원", "천원", "백만원", "억원"), index=0)
    
    unit_divisors = {"원": 1, "천원": 1000, "백만원": 1000000, "억원": 100000000}
    divisor = unit_divisors[unit_option]

st.title("📑 통합 재무제표 보고서")

# --- 스타일 함수 ---
def style_dataframe(row):
    styles = [''] * len(row)
    level = row.get('Level', 3)
    if level == 1:
        return ['background-color: #1f77b4; color: white; font-weight: bold;'] * len(row)
    elif level == 2:
        return ['background-color: #aec7e8; color: black; font-weight: bold;'] * len(row)
    return ['color: black;'] * len(row)

# --- 엑셀 저장 함수 ---
def save_styled_excel(df, sheet_name_map, unit_text):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        if 'Statement' in df.columns: statements = df['Statement'].unique()
        else: statements = ['Result']
            
        for stmt in statements:
            if 'Statement' in df.columns: sub_df = df[df['Statement'] == stmt].copy()
            else: sub_df = df.copy()
            
            # [수정] 정렬된 컬럼 순서 적용
            all_cols = sub_df.columns.tolist()
            sorted_cols = sort_columns_chronologically(all_cols)
            # 실제 존재하는 컬럼만 필터링
            final_cols = [c for c in sorted_cols if c in sub_df.columns]
            
            sheet_title = sheet_name_map.get(stmt, stmt)[:30]
            sub_df[final_cols].to_excel(writer, sheet_name=sheet_title, index=False, startrow=1)
            
            ws = writer.sheets[sheet_title]
            ws['A1'] = f"(단위: {unit_text})"
            ws['A1'].font = Font(bold=True, italic=True)
            
            fill_lv1 = PatternFill(start_color="1F77B4", end_color="1F77B4", fill_type="solid")
            font_lv1 = Font(color="FFFFFF", bold=True)
            fill_lv2 = PatternFill(start_color="AEC7E8", end_color="AEC7E8", fill_type="solid")
            font_lv2 = Font(color="000000", bold=True)
            
            # 숫자 컬럼 인덱스 찾기
            numeric_col_indices = [i+1 for i, c in enumerate(final_cols) if c != 'Account_Name']
            
            sub_df = sub_df.reset_index(drop=True)
            for idx, row in sub_df.iterrows():
                excel_row = idx + 3
                level = row.get('Level', 3)
                for col_idx in range(1, len(final_cols) + 1):
                    cell = ws.cell(row=excel_row, column=col_idx)
                    if level == 1:
                        cell.fill = fill_lv1
                        cell.font = font_lv1
                    elif level == 2:
                        cell.fill = fill_lv2
                        cell.font = font_lv2
                    
                    if col_idx - 1 in numeric_col_indices:
                        cell.number_format = '#,##0'

            ws.column_dimensions['A'].width = 30
    return buffer

# --- 메인 로직 ---
uploaded_files = st.file_uploader("파일 업로드 (Drag & Drop)", accept_multiple_files=True, type=['xlsx', 'xls', 'csv', 'pdf', 'docx', 'txt'])

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
            raw_df = logic.process_smart_merge(st.session_state.api_key, uploaded_files)
            for col in raw_df.columns:
                if col not in ['Statement', 'Level', 'Account_Name']:
                    raw_df[col] = pd.to_numeric(raw_df[col], errors='coerce').fillna(0)
            
            st.session_state['raw_data'] = raw_df
            status.update(label="✅ 생성 완료!", state="complete", expanded=False)
        except Exception as e:
            status.update(label="❌ 오류 발생", state="error")
            st.error(f"에러 내용: {e}")

if 'raw_data' in st.session_state:
    st.divider()
    st.subheader(f"📊 분석 결과 (단위: {unit_option})")
    
    display_df = st.session_state['raw_data'].copy()
    
    # [수정] 1. 값이 전부 0인 행 제거 (빈 공간 제거)
    numeric_cols = [c for c in display_df.columns if c not in ['Statement', 'Level', 'Account_Name']]
    # 숫자 컬럼 합계가 0이 아닌 행만 남김 (절대값 합계 사용)
    display_df = display_df[display_df[numeric_cols].abs().sum(axis=1) > 0]
    
    # 2. 단위 변환
    for col in numeric_cols:
        if divisor > 1:
            display_df[col] = display_df[col] / divisor

    available_types = display_df['Statement'].unique() if 'Statement' in display_df.columns else []
    type_map = {'BS': '재무상태표', 'IS': '손익계산서', 'COGM': '제조원가명세서', 'CF': '현금흐름표', 'Other': '기타'}
    
    if len(available_types) > 0:
        tabs = st.tabs([type_map.get(t, t) for t in available_types])

        for i, stmt_type in enumerate(available_types):
            with tabs[i]:
                sub_df = display_df[display_df['Statement'] == stmt_type].copy()
                
                # [수정] 컬럼 정렬 (과거 -> 현재)
                all_cols = sub_df.columns.tolist()
                sorted_cols = sort_columns_chronologically(all_cols)
                final_cols = [c for c in sorted_cols if c in sub_df.columns]

                # 포맷팅
                format_dict = {col: "{:,.0f}" for col in numeric_cols}
                
                st.dataframe(
                    sub_df[final_cols].style
                    .apply(style_dataframe, axis=1)
                    .format(format_dict),
                    use_container_width=True,
                    height=600
                )
    
    excel_buffer = save_styled_excel(display_df, type_map, unit_option)
    st.download_button(
        f"📥 엑셀 다운로드 (현재 단위: {unit_option})",
        data=excel_buffer.getvalue(),
        file_name=f"Financial_Report_{unit_option}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )