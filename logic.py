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

    # [프롬프트] 전기(Previous Period) 데이터 보존 및 누락 방지 강화
    prompt = f"""
    You are a Forensic Accountant creating a consolidated financial report.
    
    [MISSION]
    Extract **EVERY** financial figure from the provided files without omission.

    [CRITICAL RULE 1: Capture Comparative Data]
    - Quarterly reports usually compare "Current Period" (Dang-gi) vs "Previous Period" (Jeon-gi).
    - **You MUST extract BOTH.**
    - If "Current" is 2025.3Q, infer that "Previous" is 2024.3Q.
    - Explicitly label them: e.g., "2025.3Q(3M)", "2025.3Q(Cum)", "2024.3Q(3M)", "2024.3Q(Cum)".
    - Do not drop "Previous" columns even if they seem redundant. Users need them for comparison.

    [CRITICAL RULE 2: No Row Summarization]
    - List ALL unique account names.
    - If names differ slightly but mean the same (e.g., "Sales" vs "Revenue"), YOU MAY MERGE THEM.
    - BUT if the account is unique (e.g. "Export Sales"), keep it separate.

    [Logic: Structure]
    - **Statement:** BS, IS, COGM, CF, Other
    - **Level:** 1(Total), 2(Subtotal), 3(Detail)
    - **Account_Name:** Clean name.

    [Input Data]
    {full_context}

    [Output Format]
    JSON Array Only.
    [
      {{
        "Statement": "IS",
        "Level": 3,
        "Account_Name": "매출액",
        "2024.3Q(3M)": 5000,
        "2024.3Q(Cum)": 15000,
        "2025.3Q(3M)": 5500,
        "2025.3Q(Cum)": 16000
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