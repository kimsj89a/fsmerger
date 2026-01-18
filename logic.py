# logic.py
import pandas as pd
from google import genai
import json
import openpyxl
import io
import pypdf
import docx

def extract_file_content(file):
    file_ext = file.name.split('.')[-1].lower()
    content_list = []
    
    try:
        if file_ext in ['xlsx', 'xls']:
            engine = 'openpyxl' if file_ext == 'xlsx' else 'xlrd'
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
    
    if len(full_context) > 200000:
        full_context = full_context[:200000] + "\n...(Truncated)"

    client = genai.Client(api_key=api_key)

    # [수정] 프롬프트 강화: 동의어 통합 + 빈 줄 방지
    prompt = f"""
    You are an expert CFO consolidating financial reports.

    [Goal]
    Create a SINGLE, DENSE table. Minimize empty cells.

    [CRITICAL RULE 1: Aggressive Synonym Merge]
    - **Do NOT create separate rows** for synonyms. Merge them into one standard account name.
    - Example: If File A has "Sales" and File B has "Revenue", output ONE row "Revenue" (or "Sales").
    - Example: "Ordinary Deposits" and "Bank Deposits" -> Merge to "Cash & Deposits".
    - **Avoid Sparse Matrix:** Ensure year columns are filled in the SAME row as much as possible.

    [CRITICAL RULE 2: Date Columns]
    - Detect ALL periods (Years, Quarters).
    - If a period has "3M" and "Cumulative", keep both as separate columns (e.g., "2025.3Q(3M)", "2025.3Q(Cum)").

    [Logic 3: Hierarchy]
    - Assign 'Level' (1=Total, 2=Subtotal, 3=Detail).
    - Classify 'Statement': BS, IS, COGM, CF, Other.

    [Input Data]
    {full_context}

    [Output Format]
    JSON Array Only.
    [
      {{
        "Statement": "IS",
        "Level": 1,
        "Account_Name": "매출액",
        "2023": 10000,
        "2024": 11000,
        "2025.3Q(Cum)": 12000
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