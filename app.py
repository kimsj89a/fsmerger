import streamlit as st
import pandas as pd
import io
import logic 
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# 페이지 설정
st.set_page_config(page_title="Financial Report AI", layout="wide")

# --- [UI 개선 1] CSS로 파일 업로더 높이 제한 및 스크롤바 적용 ---
st.markdown("""
    <style>
        /* 파일 업로더 내의 파일 목록 영역 타겟팅 */
        [data-testid="stFileUploader"] section[aria-label="file-uploader"] > div:nth-child(2) {
            max-height: 200px; /* 대략 파일 5개 정도 높이 */
            overflow-y: auto;  /* 넘치면 스크롤바 생김 */
        }
        /* 업로더 자체의 불필요한 여백 줄이기 */
        [data-testid="stFileUploader"] {
            padding-top: 10px;
        }
    </style>
""", unsafe_allow_html=True)

# --- [UI 개선 2] API Key 좌측 상단 작게 배치 ---
if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

with st.sidebar:
    # 접이식 메뉴(expander)를 사용하여 작게 만듦
    # 키가 없으면 열려있고(True), 있으면 닫혀있음(False)
    is_expanded = not bool(st.session_state.api_key)
    with st.expander("🔑 API Key 설정", expanded=is_expanded):
        api_input = st.text_input(
            "Gemini API Key", 
            type="password", 
            value=st.session_state.api_key,
            placeholder="sk-...",
            label_visibility="collapsed" # 라벨 숨겨서 더 심플하게
        )
        if api_input:
            st.session_state.api_key = api_input
    
    st.divider() # 구분선
    st.markdown("### ⚙️ 설정 가이드")
    st.info("파일을 업로드하면 AI가 자동으로 분류 및 통합을 시작합니다.")

# --- 메인 타이틀 ---
st.title("📑 통합 재무제표 보고서")
st.markdown("다양한 파일(Excel, PDF, Word 등)을 업로드하면 **분기 데이터**를 포함한 통합 보고서를 생성합니다.")

# --- 스타일 함수들 ---
def style_dataframe(row):
    styles = [''] * len(row)
    level = row.get('Level', 3)
    
    if level == 1:
        return ['background-color: #1f77b4; color: white; font-weight: bold;'] * len(row)
    elif level == 2:
        return ['background-color: #aec7e8; color: black; font-weight: bold;'] * len(row)
    else:
        return ['color: black;'] * len(row)

def save_styled_excel(df, sheet_name_map):
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
            
            cols = [c for c in sub_df.columns if c not in ['Statement', 'Level']]
            if 'Account_Name' in cols:
                cols.remove('Account_Name')
                cols = ['Account_Name'] + cols
            
            sheet_title = sheet_name_map.get(stmt, stmt)[:30]
            sub_df[cols].to_excel(writer, sheet_name=sheet_title, index=False)
            
            workbook = writer.book
            worksheet = writer.sheets[sheet_title]
            
            fill_lv1 = PatternFill(start_color="1F77B4", end_color="1F77B4", fill_type="solid")
            font_lv1 = Font(color="FFFFFF", bold=True)
            fill_lv2 = PatternFill(start_color="AEC7E8", end_color="AEC7E8", fill_type="solid")
            font_lv2 = Font(color="000000", bold=True)
            
            sub_df = sub_df.reset_index(drop=True)
            for idx, row in sub_df.iterrows():
                excel_row = idx + 2
                level = row.get('Level', 3)
                
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
                
            worksheet.column_dimensions['A'].width = 30

    return buffer

# --- 메인 로직 ---
uploaded_files = st.file_uploader(
    "분석할 파일들을 선택하세요 (Drag & Drop)", 
    accept_multiple_files=True, 
    type=['xlsx', 'xls', 'csv', 'pdf', 'docx', 'txt']
)

if uploaded_files and st.session_state.api_key:
    if st.button("보고서 생성 시작", type="primary"):
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
                    
                    display_cols = [c for c in sub_df.columns if c not in ['Statement', 'Level']]
                    if 'Account_Name' in display_cols:
                        display_cols.remove('Account_Name')
                        display_cols = ['Account_Name'] + display_cols
                    
                    st.dataframe(
                        sub_df[display_cols].style.apply(style_dataframe, axis=1),
                        use_container_width=True,
                        height=600
                    )

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