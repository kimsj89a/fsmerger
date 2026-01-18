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
                # 데이터가 없는 빈 행/열만 제거하고 전송
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
    
    # [수정] 300,000자 제한 로직 삭제 (Full Context 전송)
    # Gemini 1.5 Flash는 약 100만 토큰(수백만 자)까지 처리가 가능합니다.
    # 단, 데이터가 너무 방대할 경우(수십 MB 텍스트) 처리 속도가 느려질 수는 있습니다.

    client = genai.Client(api_key=api_key)

    # [프롬프트] 연말/분기 구분 + 모든 계정 나열
    prompt = f"""
    You are a Financial Analyst creating a consolidated report.

    [RULE 1: Annual Reports (Year-End)]
    - If the data is for a full year (e.g., 2023, 2024), create simple year columns.
    - **Format:** "2023", "2024"
    - Do NOT split into "3M" or "Cumulative".

    [RULE 2: Interim Reports (Quarterly/Semi-Annual)]
    - If the data is for a quarter (e.g., 2025.1Q, 2025.3Q), you MUST capture BOTH "3 Months" and "Cumulative".
    - **Format:** "2025.3Q(3M)", "2025.3Q(Cum)"
    - Extract "Previous Period" data as well (e.g., "2024.3Q(3M)").

    [RULE 3: Data Integrity (No Summarization)]
    - **List ALL unique accounts.** Even if they look similar, if the text is different, keep them separate rows.
    - **Values:** Extract exact figures. If empty, use 0.
    - **Order:** Assets -> Liabilities -> Equity -> Revenue -> Expense.

    [Input Data]
    {full_context}

    [Output Format]
    JSON Array Only.
    [
      {{
        "Statement": "IS",
        "Level": 3,
        "Account_Name": "매출액",
        "2023": 10000,
        "2024": 12000,
        "2025.3Q(3M)": 3500,
        "2025.3Q(Cum)": 10500
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

    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
    if "[" in cleaned_text and "]" in cleaned_text:
        s = cleaned_text.find("[")
        e = cleaned_text.rfind("]") + 1
        cleaned_text = cleaned_text[s:e]
    
    data_list = json.loads(cleaned_text)
    df = pd.DataFrame(data_list)
    
    return df