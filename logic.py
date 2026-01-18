import pandas as pd
from google import genai
import json
import openpyxl
import io
import pypdf
import docx

def extract_file_content(file):
    """
    파일 내용을 텍스트로 변환
    """
    file_ext = file.name.split('.')[-1].lower()
    content_list = []
    
    try:
        if file_ext in ['xlsx', 'xls']:
            engine = 'openpyxl' if file_ext == 'xlsx' else 'xlrd'
            # header=None으로 읽어서 모든 데이터를 있는 그대로 가져옴
            dfs = pd.read_excel(file, sheet_name=None, engine=engine, header=None)
            for sheet_name, df in dfs.items():
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
    full_context = ""
    for file in target_files:
        full_context += extract_file_content(file) + "\n\n"
    
    # Context 제한 없음 (Full processing)

    client = genai.Client(api_key=api_key)

    # [프롬프트] 모든 재무제표 식별 + 상세 계정 나열 지시
    prompt = f"""
    You are a Forensic Accountant creating a fully detailed consolidated report.

    [MISSION]
    Extract **EVERY SINGLE ACCOUNT** from ALL sheets found in the files.
    Do NOT summarize. Do NOT omit details.

    [RULE 1: Identify All Statement Types]
    Classify every table into one of these types:
    - **BS**: Balance Sheet (재무상태표)
    - **IS**: Income Statement (손익계산서)
    - **COGM**: Cost of Goods Manufactured (제조원가명세서) - *Crucial to capture detailed costs.*
    - **SCE**: Statement of Changes in Equity (자본변동표)
    - **RE**: Retained Earnings (이익잉여금처분계산서)
    - **CF**: Cash Flow (현금흐름표)

    [RULE 2: Granularity (Lowest Level Details)]
    - For **IS** and **COGM**, you MUST list the lowest level accounts.
    - **Bad:** showing only "Selling & Admin Expenses" (Level 1).
    - **Good:** showing "Salaries", "Rent", "Travel Expense", "Depreciation" (Level 3) under "Selling & Admin Expenses".
    - **Capture EVERYTHING.** If a sheet has 100 rows, I want 100 rows in the output.

    [RULE 3: Date Columns (Annual vs Interim)]
    - **Annual (Year-End):** "2023", "2024" (Simple Year)
    - **Interim (Quarter):** "2025.3Q(3M)", "2025.3Q(Cum)" (Split 3M/Cum)
    - Capture "Previous Period" comparisons if available.

    [Input Data]
    {full_context}

    [Output Format]
    JSON Array Only.
    [
      {{
        "Statement": "COGM",
        "Level": 3,
        "Account_Name": "원재료비",
        "2023": 5000,
        "2024": 6000,
        "2025.3Q(Cum)": 4500
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

    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
    if "[" in cleaned_text and "]" in cleaned_text:
        s = cleaned_text.find("[")
        e = cleaned_text.rfind("]") + 1
        cleaned_text = cleaned_text[s:e]
    
    data_list = json.loads(cleaned_text)
    df = pd.DataFrame(data_list)
    
    return df