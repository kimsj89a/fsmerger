# logic.py
import pandas as pd
from google import genai
import json
import openpyxl
import os

TAXONOMY_FILE = '2018taxonomy.xlsx'

def load_internal_taxonomy():
    """
    내장된 Taxonomy 파일을 읽어 텍스트로 변환
    """
    if not os.path.exists(TAXONOMY_FILE):
        raise FileNotFoundError(f"'{TAXONOMY_FILE}' 파일이 없습니다. 프로젝트 폴더에 파일을 넣어주세요.")

    try:
        wb = openpyxl.load_workbook(TAXONOMY_FILE, data_only=True)
        all_text_data = []
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if ws.sheet_state == 'hidden' or ws.sheet_state == 'veryHidden':
                continue
            
            data = ws.values
            try:
                columns = next(data)[0:]
            except StopIteration:
                continue
                
            df = pd.DataFrame(data, columns=columns)
            sheet_csv = df.to_csv(index=False)
            all_text_data.append(f"--- Standard Sheet: {sheet_name} ---\n{sheet_csv}")
            
        return "\n".join(all_text_data)
    except Exception as e:
        raise Exception(f"Taxonomy 파일 로딩 실패: {e}")

def load_target_excel_files(uploaded_files):
    """
    업로드된 파일들을 읽어 텍스트 컨텍스트로 변환
    """
    target_context_list = []
    
    for file in uploaded_files:
        try:
            wb = openpyxl.load_workbook(file, data_only=True)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if ws.sheet_state == 'hidden' or ws.sheet_state == 'veryHidden':
                    continue
                
                data = ws.values
                try:
                    columns = next(data)[0:]
                    df = pd.DataFrame(data, columns=columns)
                    # 데이터 출처 표시
                    header = f"--- User Data (File: {file.name}, Sheet: {sheet_name}) ---"
                    target_context_list.append(f"{header}\n{df.to_csv(index=False)}")
                except:
                    continue
        except Exception as e:
            print(f"파일 읽기 에러 ({file.name}): {e}")
            
    return "\n".join(target_context_list)

def process_financial_mapping(api_key, target_files):
    """
    전체 매핑 프로세스 실행 함수
    """
    # 1. Taxonomy 로드
    taxonomy_context = load_internal_taxonomy()

    # 2. 타겟 데이터 로드
    full_target_context = load_target_excel_files(target_files)
    
    # 토큰 제한 처리 (약 10만자)
    if len(full_target_context) > 100000:
        full_target_context = full_target_context[:100000] + "\n...(Data Truncated)"

    # 3. Gemini Client 생성
    client = genai.Client(api_key=api_key)

    # 4. 프롬프트 작성
    prompt = f"""
    [Role]
    당신은 전문 회계 감사 시스템입니다. User Data를 Standard Taxonomy 구조에 매핑하십시오.

    [Input 1: Standard Taxonomy (정답지)]
    {taxonomy_context}

    [Input 2: User Data (문제지)]
    {full_target_context}

    [Mapping Rules]
    1. **Strict Hierarchy:** 결과의 'Major', 'Medium', 'Account'는 오직 [Standard Taxonomy]에 있는 명칭만 사용하십시오.
    2. **Fuzzy Matching:** User Data의 계정명과 100% 일치하지 않더라도, 회계적 의미가 같은 Standard 항목에 합산하십시오.
    3. **Columns:** 연도(Year) 데이터를 찾아 컬럼으로 만드십시오.
    
    [Output Format]
    JSON Array Only.
    [
        {{
            "Standard_Major": "자산",
            "Standard_Medium": "유동자산",
            "Standard_Account": "현금및현금성자산",
            "Original_Accounts": "현금, 예금 (매핑된 원본 계정들)",
            "2022": 15000,
            "2023": 20000
        }},
        ...
    ]
    """

    # 5. AI 호출
    try:
        response = client.models.generate_content(
            model="gemini-3-flash-preview",
            contents=prompt
        )
    except Exception as e:
        if "404" in str(e):
             # 모델명 에러 시 자동 폴백 시도 (선택 사항)
             response = client.models.generate_content(
                model="gemini-1.5-flash",
                contents=prompt
            )
        else:
            raise e

    # 6. JSON 파싱
    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
    if "[" in cleaned_text and "]" in cleaned_text:
        s = cleaned_text.find("[")
        e = cleaned_text.rfind("]") + 1
        cleaned_text = cleaned_text[s:e]
    
    result_data = json.loads(cleaned_text)
    df = pd.DataFrame(result_data)

    # 7. 컬럼 정리
    cols = df.columns.tolist()
    std_cols = ['Standard_Major', 'Standard_Medium', 'Standard_Account', 'Original_Accounts']
    other_cols = [c for c in cols if c not in std_cols]
    
    # 실제 존재하는 컬럼만 선택
    final_std = [c for c in std_cols if c in df.columns]
    final_other = sorted([c for c in other_cols]) # 연도순 정렬 등
    
    return df[final_std + final_other]