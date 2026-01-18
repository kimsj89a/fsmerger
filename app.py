import streamlit as st
import pandas as pd
import logic 
import ui_results  # [UI 모듈 임포트]

# 페이지 설정
st.set_page_config(page_title="Financial Report AI", layout="wide")

# ==========================================
# [보안 강화] F12, 우클릭, 드래그 방지
# ==========================================
def inject_security_code():
    st.markdown("""
        <style>
            body { user-select: none; -webkit-user-select: none; }
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            input, textarea, [contenteditable] { user-select: text; -webkit-user-select: text; }
            /* 분석 결과 타이틀과 단위 선택기 높이 맞추기 */
            div[data-testid="stVerticalBlock"] > div[style*="flex-direction: column;"] > div[data-testid="stVerticalBlock"] {
                gap: 0rem;
            }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("""
        <script>
            document.addEventListener('DOMContentLoaded', (event) => {
                document.addEventListener('contextmenu', e => e.preventDefault());
                document.addEventListener('keydown', e => {
                    if (e.key === 'F12' || e.keyCode === 123) { e.preventDefault(); return false; }
                    if (e.ctrlKey && e.shiftKey && ['I','J','C','i','j','c'].includes(e.key)) { e.preventDefault(); return false; }
                    if (e.ctrlKey && ['U','u'].includes(e.key)) { e.preventDefault(); return false; }
                });
            });
        </script>
    """, unsafe_allow_html=True)

inject_security_code()

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
    </style>
""", unsafe_allow_html=True)

# ==========================================
# [UI 1] 타이틀 및 API Key
# ==========================================
st.title("📑 통합 재무제표 보고서")

if 'api_key' not in st.session_state:
    st.session_state.api_key = ''

api_key_input = st.text_input(
    "Gemini API Key", 
    type="password", 
    placeholder="sk-...", 
    value=st.session_state.api_key
)
if api_key_input:
    st.session_state.api_key = api_key_input

st.divider()

# ==========================================
# [UI 2] 파일 업로더 (초기화 버튼 삭제됨)
# ==========================================
uploaded_files = st.file_uploader(
    "분석할 파일들을 업로드하세요", 
    accept_multiple_files=True, 
    type=['xlsx', 'xls', 'csv', 'pdf', 'docx', 'txt']
)

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
            status = st.status("AI가 분석 중입니다...", expanded=True)
            try:
                # 1. 로직 실행 (logic.py)
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
                
                # 분석 새로 하면 채팅 기록도 리셋
                if 'messages' in st.session_state:
                    del st.session_state['messages']
                    
                status.update(label="✅ 분석 완료!", state="complete", expanded=False)
            except Exception as e:
                status.update(label="❌ 오류 발생", state="error")
                st.error(f"에러 내용: {e}")
    else:
        st.warning("👆 상단에 API Key를 먼저 입력해주세요.")

# ==========================================
# [UI 3] 분석 결과 및 채팅 (모듈 호출)
# ==========================================
if 'raw_data' in st.session_state:
    # ui_results.py에 있는 함수 호출
    ui_results.render_analysis_result(st.session_state.api_key)