# logic.py
import pandas as pd
from google import genai
import json
import openpyxl
import io
import pypdf
import docx

def extract_file_content(file):
    """
    다양한 파일 형식(xls, xlsx, csv, pdf, docx, txt)을 읽어 텍스트 컨텍스트로 변환
    """
    file_ext = file.name.split('.')[-1].lower()
    content_list = []
    
    try:
        # 1. Excel (xlsx, xls)
        if file_ext in ['xlsx', 'xls']:
            # xls 지원을 위해 엔진 분기 처리 (openpyxl or xlrd)
            engine = 'openpyxl' if file_ext == 'xlsx' else 'xlrd'
            # 모든 시트 읽기
            dfs = pd.read_excel(file, sheet_name=None, engine=engine)
            for sheet_name, df in dfs.items():
                df = df.dropna(how='all').dropna(axis=1, how='all')
                csv_text = df.to_csv(index=False)
                content_list.append(f"File: {file.name} | Sheet: {sheet_name}\n{csv_text}")

        # 2. CSV
        elif file_ext == 'csv':
            df = pd.read_csv(file)
            content_list.append(f"File: {file.name}\n{df.to_csv(index=False)}")

        # 3. PDF (텍스트 추출)
        elif file_ext == 'pdf':
            pdf_reader = pypdf.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            content_list.append(f"File: {file.name} (PDF Content)\n{text}")

        # 4. Word (docx)
        elif file_ext in ['docx', 'doc']:
            doc = docx.Document(file)
            text = "\n".join([para.text for para in doc.paragraphs])
            content_list.append(f"File: {file.name} (Word Content)\n{text}")
            
        # 5. Text
        elif file_ext == 'txt':
            text = file.getvalue().decode("utf-8")
            content_list.append(f"File: {file.name}\n{text}")

    except Exception as e:
        return f"Error reading {file.name}: {str(e)}"

    return "\n\n".join(content_list)

def process_smart_merge(api_key, target_files):
    # 1. 모든 파일 텍스트화
    full_context = ""
    for file in target_files:
        full_context += extract_file_content(file) + "\n\n"
    
    if len(full_context) > 200000: # 컨텍스트 길이 확장
        full_context = full_context[:200000] + "\n...(Data Truncated)"

    client = genai.Client(api_key=api_key)

    # 2. 프롬프트 수정: 날짜 인식 강화 및 서식 데이터 요청
    prompt = f"""
    You are a CFO creating a consolidated financial report.

    [Goal]
    Merge data from provided files into a single structured table.

    [Logic 1: Statement Classification]
    Classify each row into: 'BS' (Balance Sheet), 'IS' (Income Statement), 'COGM' (Cost of Goods Manufactured), 'CF' (Cash Flow), or 'Other'.

    [Logic 2: Hierarchy Level]
    Assign a 'Level' (1, 2, 3) for formatting:
    - Level 1: Totals/Majors (e.g., 자산총계, 매출액).
    - Level 2: Sub-totals (e.g., 유동자산, 영업이익).
    - Level 3: Details (e.g., 현금, 접대비).

    [Logic 3: Date Columns (Crucial)]
    - Detect ALL time periods as columns.
    - **Include Quarters:** If data contains '2025.3Q', '2024.1Q', treat them as valid columns just like '2024'.
    - Do not drop any time-related columns.

    [Logic 4: Context-Aware Merge]
    - Preserve logical accounting order (Asset -> Liability -> Equity). Do NOT sort alphabetically.
    - Interleave missing items naturally.

    [Input Data]
    {full_context}

    [Output Format]
    JSON Array Only.
    [
      {{
        "Statement": "BS",
        "Level": 1,
        "Account_Name": "자산총계",
        "2024": 10000,
        "2025.3Q": 12000
      }},
      ...
    ]
    """

    try:
        response = client.models.generate_content(
            model="gemini-3-flash-preview", 
            contents=prompt
        )
    except Exception:
        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=prompt
        )

    # 3. 파싱
    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
    if "[" in cleaned_text and "]" in cleaned_text:
        s = cleaned_text.find("[")
        e = cleaned_text.rfind("]") + 1
        cleaned_text = cleaned_text[s:e]
    
    data_list = json.loads(cleaned_text)
    df = pd.DataFrame(data_list)
    
    # [수정] 공백/아이콘 추가 로직 삭제 (순수 데이터만 반환)

    return df