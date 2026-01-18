import streamlit as st
import pandas as pd
import io
import re
import logic 
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# 페이지 설정
st.set_page_config(page_title="Financial Report AI", layout="wide")

# --- CSS: 스타일링 ---
st.markdown("""
    <style>
        .file-list-box {
            border: 1px solid #e6e6e6; padding: 10px; border-radius: 5px;
            max-height: 200px; overflow-y: auto; background-color: #f9f9f9; margin-bottom: 20px;
        }
        .file-item {
            font-size: 0.9em; margin-bottom: 4px; padding: 4px; background: white; border-radius: 3px;
        }
        /* 분석 결과 타이틀과 단위 선택기 높이 맞추기 */
        div[data-testid="stVerticalBlock"] > div[style*="flex-direction: column;"] > div[data-testid="stVerticalBlock"] {
            gap: 0rem;
        }
    </style>
""", unsafe_allow_html=True)

# --- 초기화 콜백 함수 ---
def clear_all():
    # 파일 업로더 초기화
    if 'uploader_key' in st.session_state:
        st.session_state['uploader_key'] = []
    # 데이터 초기화
    if 'raw_data' in st.session_state:
        del st.session_state['raw_data']

# --- 컬럼 정렬 함수 ---
def sort_columns_chronologically(columns):
    fixed_cols = ['Account_Name']
    date_cols = [c for c in columns if c not in ['Statement', 'Level', 'Account_Name']]
    
    def date_sort_key(col_name):
        s_name = str(col_name)
        year_match = re.search(r'(\d{4})', s_name)
        year = int(year_match.group(1)) if year_match else 9999
        
        sub_val = 0
        if '1Q' in s_name: sub_val = 1
        elif '2Q' in s_name: sub_val = 4
        elif '3Q' in s_name: sub_val = 7
        elif '4Q' in s_name: sub_val = 10
        
        is_cum = 1 if '누적' in s_name or 'Cum' in s_name or 'Year' in s_name else 0
        return (year, sub_val, is_cum, s_name)
    
    sorted_date_cols = sorted(date_cols, key=date_sort_key)
    return fixed_cols + sorted_date_cols

# --- 스타일 함수 ---
def style_dataframe(row):
    level = row.get('Level', 3)
    if level == 1: return ['background-color: #1f77b4; color: white; font-weight: bold;'] * len(row)
    elif level == 2: return ['background-color: #aec7e8; color: black; font-weight: bold;'] * len(row)
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
            
            all_cols = sub_df.columns.tolist()
            sorted_cols = sort_columns_chronologically(all_cols)
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

# ==========================================
# [UI 1] 타이틀 및 API Key (Body 상단)
# ==========================================
st.title("📑 통합 재무제표 보고서")

if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

# 화면 상단에 API Key 입력 배치
api_key_input = st.text_input(
    "Gemini API Key", 
    type="password", 
    placeholder="sk-...", 
    value=st.session_state.api_key,
    help="Google AI Studio에서 발급받은 API Key를 입력하세요."
)
if api_key_input:
    st.session_state.api_key = api_key_input

st.divider()

# ==========================================
# [UI 2] 파일 업로더 & 초기화 버튼
# ==========================================
col_upload, col_clear = st.columns([0.85, 0.15])

with col_upload:
    uploaded_files = st.file_uploader(
        "분석할 파일들을 업로드하세요 (Excel, PDF, Word 등)", 
        accept_multiple_files=True, 
        type=['xlsx', 'xls', 'csv', 'pdf', 'docx', 'txt'],
        key='uploader_key' # 초기화를 위한 키 설정
    )

with col_clear:
    # 아래 여백을 좀 줘서 업로더 버튼과 라인 맞추기
    st.write("") 
    st.write("")
    if st.button("🗑️ 초기화", type="secondary", use_container_width=True, on_click=clear_all):
        pass # 콜백에서 처리됨

# 파일 목록 뷰어
if uploaded_files:
    file_list_html = '<div class="file-list-box">'
    for f in uploaded_files:
        size_kb = f.size / 1024
        file_list_html += f'<div class="file-item">📄 {f.name} ({size_kb:.1f} KB)</div>'
    file_list_html += '</div>'
    st.markdown(file_list_html, unsafe_allow_html=True)

    if st.session_state.api_key:
        if st.button("🚀 보고서 생성 시작", type="primary", use_container_width=True):
            status = st.status("AI가 모든 시트의 상세 계정을 분석 중입니다...", expanded=True)
            try:
                # 로직 호출
                raw_df = logic.process_smart_merge(st.session_state.api_key, uploaded_files)
                
                # 숫자 변환
                for col in raw_df.columns:
                    if col not in ['Statement', 'Level', 'Account_Name']:
                        raw_df[col] = pd.to_numeric(raw_df[col], errors='coerce').fillna(0)
                
                # 빈 열 삭제
                numeric_cols = [c for c in raw_df.columns if c not in ['Statement', 'Level', 'Account_Name']]
                zero_cols = [c for c in numeric_cols if raw_df[c].abs().sum() == 0]
                if zero_cols:
                    raw_df = raw_df.drop(columns=zero_cols)
                
                st.session_state['raw_data'] = raw_df
                status.update(label="✅ 분석 완료!", state="complete", expanded=False)
            except Exception as e:
                status.update(label="❌ 오류 발생", state="error")
                st.error(f"에러 내용: {e}")
    else:
        st.warning("👆 상단에 API Key를 먼저 입력해주세요.")


# ==========================================
# [UI 3] 분석 결과 & 우측 단위 설정
# ==========================================
if 'raw_data' in st.session_state:
    st.divider()
    
    # [핵심] 타이틀(좌측)과 단위 선택기(우측) 배치
    c_title, c_unit = st.columns([0.7, 0.3])
    
    with c_unit:
        unit_option = st.selectbox(
            "단위 선택", 
            ("원", "천원", "백만원", "억원"), 
            index=0,
            label_visibility="visible" # or "collapsed"
        )
        unit_divisors = {"원": 1, "천원": 1000, "백만원": 1000000, "억원": 100000000}
        divisor = unit_divisors[unit_option]

    with c_title:
        st.subheader(f"📊 분석 결과 (단위: {unit_option})")

    # 데이터 처리 (단위 변환 및 필터링)
    display_df = st.session_state['raw_data'].copy()
    numeric_cols = [c for c in display_df.columns if c not in ['Statement', 'Level', 'Account_Name']]
    
    # 0인 행 제거
    display_df = display_df[display_df[numeric_cols].abs().sum(axis=1) != 0]
    
    # 단위 나눗셈
    for col in numeric_cols:
        if divisor > 1:
            display_df[col] = display_df[col] / divisor

    # 탭 및 테이블 출력
    available_types = display_df['Statement'].unique() if 'Statement' in display_df.columns else []
    type_map = {
        'BS': '재무상태표', 'IS': '손익계산서', 'COGM': '제조원가명세서', 
        'CF': '현금흐름표', 'SCE': '자본변동표', 'RE': '이익잉여금', 'Other': '기타'
    }
    
    if len(available_types) > 0:
        tabs = st.tabs([type_map.get(t, t) for t in available_types])

        for i, stmt_type in enumerate(available_types):
            with tabs[i]:
                sub_df = display_df[display_df['Statement'] == stmt_type].copy()
                
                # 컬럼 정렬
                all_cols = sub_df.columns.tolist()
                sorted_cols = sort_columns_chronologically(all_cols)
                final_cols = [c for c in sorted_cols if c in sub_df.columns]

                # 숫자 포맷
                format_dict = {col: "{:,.0f}" for col in numeric_cols if col in final_cols}
                
                st.dataframe(
                    sub_df[final_cols].style
                    .apply(style_dataframe, axis=1)
                    .format(format_dict),
                    use_container_width=True,
                    height=600
                )
    
    # 엑셀 다운로드
    excel_buffer = save_styled_excel(display_df, type_map, unit_option)
    st.download_button(
        f"📥 엑셀 다운로드 (현재 단위: {unit_option})",
        data=excel_buffer.getvalue(),
        file_name=f"Financial_Report_{unit_option}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )