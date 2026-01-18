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
    파일 내용을 텍스트로 변환 (기존과 동일)
    """
    file_ext = file.name.split('.')[-1].lower()
    content_list = []
    
    try:
        if file_ext in ['xlsx', 'xls']:
            engine = 'openpyxl' if file_ext == 'xlsx' else 'xlrd'
            # header=None으로 읽어서 모든 텍스트를 다 가져옴
            dfs = pd.read_excel(file, sheet_name=None, engine=engine, header=None)
            for sheet_name, df in dfs.items():
                # 빈 행/열 제거
                df = df.dropna(how='all').dropna(axis=1, how='all')
                csv_text = df.to_csv(index=False, header=False)
                content_list.append(f"File: {file.name} | Sheet: {sheet_name}\n{csv_text}")

        elif file_ext == 'csv':
            df = pd.read_csv(file, header=None)
            content_list.append(f"File: {file.name}\n{df.to_csv(index=False, header=False)}")

        elif file_ext == 'pdf':
            pdf_reader = pypdf.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            content_list.append(f"File: {file.name} (PDF)\n{text}")

        elif file_ext in ['docx', 'doc']:
            doc = docx.Document(file)
            text = "\n".join([para.text for para in doc.paragraphs])
            content_list.append(f"File: {file.name} (Word)\n{text}")
            
        elif file_ext == 'txt':
            text = file.getvalue().decode("utf-8")
            content_list.append(f"File: {file.name}\n{text}")

    except Exception as e:
        return f"Error reading {file.name}: {str(e)}"

    return "\n\n".join(content_list)

def process_smart_merge(api_key, target_files):
    # 1. 컨텍스트 생성
    full_context = ""
    for file in target_files:
        full_context += extract_file_content(file) + "\n\n"
    
    # 컨텍스트 길이 제한 (약 30만자까지 늘림 - Flash 모델은 처리 가능)
    if len(full_context) > 300000:
        full_context = full_context[:300000] + "\n...(Truncated)"

    client = genai.Client(api_key=api_key)

    # [핵심 수정] 프롬프트: 완전 보존(Union) 지시
    prompt = f"""
    You are a Forensic Accountant. Your job is to create a consolidated spreadsheet that includes **EVERY SINGLE ACCOUNT ITEM** from the source files.

    [MISSION: ZERO OMISSION]
    1. **List ALL unique account names.** If File A has "Account X" and File B does not, you MUST list "Account X" and put 0 for File B.
    2. **Do NOT Summarize.** Do not merge "Travel Expense" and "Transportation Expense" unless they are exactly the same string. Keep them as separate rows.
    3. **Preserve Granularity.** If the source has detail rows, keep them. Do not just show the Totals.

    [Logic 1: Columns (Time Periods)]
    - Detect ALL time headers (Years, Quarters, Months).
    - Handle Split Headers: "2025.3Q (3M)" and "2025.3Q (Cum)" must be separate columns.

    [Logic 2: Structure]
    - **Statement:** BS, IS, COGM, CF, Other
    - **Level:** 1 (Total), 2 (Subtotal), 3 (Detail)
    - **Account_Name:** The exact name from the source file.

    [Logic 3: Order]
    - Maintain standard accounting order (Assets -> Liabilities -> Equity -> Revenue -> Expense).
    - Do NOT sort alphabetically.

    [Input Data]
    {full_context}

    [Output Format]
    JSON Array Only.
    [
      {{
        "Statement": "IS",
        "Level": 3,
        "Account_Name": "복리후생비",
        "2023": 1000,
        "2024": 1200,
        "2025.3Q(3M)": 300,
        "2025.3Q(Cum)": 900
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
        # 모델 폴백
        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=prompt
        )

    # 파싱 로직
    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
    if "[" in cleaned_text and "]" in cleaned_text:
        s = cleaned_text.find("[")
        e = cleaned_text.rfind("]") + 1
        cleaned_text = cleaned_text[s:e]
    
    data_list = json.loads(cleaned_text)
    df = pd.DataFrame(data_list)
    
    return df