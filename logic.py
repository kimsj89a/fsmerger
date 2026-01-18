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
            # header=None으로 읽어서 헤더 구조를 AI가 통째로 보게 함
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
    
    if len(full_context) > 300000:
        full_context = full_context[:300000] + "\n...(Truncated)"

    client = genai.Client(api_key=api_key)

    # [프롬프트] 로직 간소화: 연말 vs 분기 구분 명확화
    prompt = f"""
    You are a Financial Analyst creating a consolidated report.
    Simplify the merging logic based on two report types.

    [RULE 1: Annual Reports (Year-End)]
    - If the data is for a full year (e.g., 2023, 2024), create simple year columns.
    - **Format:** "2023", "2024"
    - Do NOT split into "3M" or "Cumulative" for annual reports (it's always 12M/Cumulative).

    [RULE 2: Interim Reports (Quarterly/Semi-Annual)]
    - If the data is for a quarter (e.g., 2025.1Q, 2025.3Q), you MUST capture BOTH "3 Months" and "Cumulative".
    - **Format:** "2025.3Q(3M)", "2025.3Q(Cum)"
    - Note: Balance Sheet (BS) usually only has "Period End" (treat as Cum). Income Statement (IS) has both.

    [RULE 3: Data Integrity]
    - **List ALL accounts.** Do not group unless they are clearly identical.
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
        "2023": 10000,           // Annual
        "2024": 12000,           // Annual
        "2025.3Q(3M)": 3500,     // Interim (3 Months)
        "2025.3Q(Cum)": 10500    // Interim (Cumulative)
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