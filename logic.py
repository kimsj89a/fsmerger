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
    다양한 파일 형식을 텍스트로 변환
    """
    file_ext = file.name.split('.')[-1].lower()
    content_list = []
    
    try:
        # 1. Excel
        if file_ext in ['xlsx', 'xls']:
            engine = 'openpyxl' if file_ext == 'xlsx' else 'xlrd'
            # header=None으로 읽어서 헤더 구조를 AI가 통째로 보게 함 (멀티 헤더 인식률 향상)
            dfs = pd.read_excel(file, sheet_name=None, engine=engine, header=None)
            for sheet_name, df in dfs.items():
                df = df.dropna(how='all').dropna(axis=1, how='all')
                # 데이터가 너무 많으면 상위 100행만 (구조 파악용) + 꼬리 
                csv_text = df.to_csv(index=False, header=False)
                content_list.append(f"File: {file.name} | Sheet: {sheet_name}\n{csv_text}")

        # 2. CSV
        elif file_ext == 'csv':
            df = pd.read_csv(file, header=None)
            content_list.append(f"File: {file.name}\n{df.to_csv(index=False, header=False)}")

        # 3. PDF
        elif file_ext == 'pdf':
            pdf_reader = pypdf.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            content_list.append(f"File: {file.name} (PDF)\n{text}")

        # 4. Word
        elif file_ext in ['docx', 'doc']:
            doc = docx.Document(file)
            text = "\n".join([para.text for para in doc.paragraphs])
            content_list.append(f"File: {file.name} (Word)\n{text}")
            
        # 5. Text
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
    
    if len(full_context) > 200000:
        full_context = full_context[:200000] + "\n...(Truncated)"

    client = genai.Client(api_key=api_key)

    # 2. 프롬프트 수정 (복합 헤더 처리 강조)
    prompt = f"""
    You are a CFO creating a consolidated financial report.

    [Goal]
    Merge data from provided files into a single structured table.

    [Logic 1: Statement Classification]
    Classify each row: 'BS', 'IS', 'COGM', 'CF', 'Other'.

    [Logic 2: Hierarchy Level]
    Assign 'Level' (1, 2, 3) for formatting.

    [Logic 3: Date Columns & Multi-Column Headers (Crucial)]
    - Detect ALL time periods.
    - **Handle Split Periods:** If a year/quarter has sub-columns like "3개월" (3 Months) and "누적" (Cumulative), **KEEP BOTH**.
    - **Naming:** Combine headers to make them unique. 
      - Example: "2025.3Q (3M)", "2025.3Q (Cum)" or "2025.3Q_3개월", "2025.3Q_누적".
    - Do NOT merge them into one. Keep raw granularity.

    [Logic 4: Context-Aware Merge]
    - Preserve logical accounting order (Asset -> Liability -> Equity). 
    - **Interleave:** Insert missing items naturally.

    [Input Data]
    {full_context}

    [Output Format]
    JSON Array Only.
    [
      {{
        "Statement": "IS",
        "Level": 1,
        "Account_Name": "매출액",
        "2024": 10000,
        "2025.3Q_3M": 3000,
        "2025.3Q_Cum": 12000
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
    
    return df