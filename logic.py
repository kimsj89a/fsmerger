import pandas as pd
from google import genai
import json
import openpyxl
import io

def extract_sheet_data(file):
    context_list = []
    wb = openpyxl.load_workbook(file, data_only=True)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.sheet_state == 'hidden' or ws.sheet_state == 'veryHidden':
            continue
        data = ws.values
        try:
            header = next(data)
            columns = [str(h) if h is not None else f"Unnamed_{i}" for i, h in enumerate(header)]
            df = pd.DataFrame(data, columns=columns)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            csv_text = df.to_csv(index=False)
            context_list.append(f"FileName: {file.name} | Sheet: {sheet_name}\n{csv_text}")
        except StopIteration:
            continue
    return "\n\n".join(context_list)

def process_smart_merge(api_key, target_files):
    # 1. 데이터 텍스트화
    full_context = ""
    for file in target_files:
        full_context += extract_sheet_data(file) + "\n\n"
    
    if len(full_context) > 150000:
        full_context = full_context[:150000] + "\n...(Data Truncated)"

    client = genai.Client(api_key=api_key)

    # 2. 프롬프트 강화: 재무제표 타입 분류 & 계층 구조 시각화
    prompt = f"""
    You are a Chief Financial Officer (CFO).
    Analyze the provided multiple Excel files and merge them into a consolidated financial report.

    [Goal]
    Create a unified table with 'Year' columns, structured hierarchically, and separated by Financial Statement Type.

    [Logic 1: Statement Classification]
    For every row, identify which Financial Statement it belongs to:
    - **BS**: Balance Sheet (재무상태표 - 자산, 부채, 자본)
    - **IS**: Income Statement (손익계산서 - 매출, 비용, 이익)
    - **COGM**: Cost of Goods Manufactured (제조원가명세서 - 재료비, 노무비, 경비)
    - **CF**: Cash Flow (현금흐름표)

    [Logic 2: Hierarchy & Styling]
    Classify the 'Level' of each account to create a visual hierarchy:
    - **Level 1 (Major):** Top category (e.g., 자산총계, 유동자산, 부채총계).
    - **Level 2 (Medium):** Sub-category (e.g., 현금및현금성자산, 매출채권).
    - **Level 3 (Detail):** Specific items (e.g., 보통예금, 외상매출금).
    *Tip: If the row is a 'Total' or 'Sum' line, it is usually Level 1.*

    [Logic 3: Context-Aware Merge]
    - **Preserve Order:** Do not sort alphabetically. Keep the logical accounting flow (Asset -> Liability -> Equity).
    - **Interleave:** Insert missing accounts from different years into their logical position.
    
    [Input Data]
    {full_context}

    [Output Format]
    JSON Array Only.
    [
      {{
        "Statement": "BS",
        "Level": 1,
        "Account_Name": "자산총계",
        "2022": 10000,
        "2023": 12000
      }},
      {{
        "Statement": "BS",
        "Level": 2,
        "Account_Name": "유동자산",
        "2022": 5000,
        "2023": 6000
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

    # 4. 시각적 들여쓰기 처리 (Excel/화면용)
    # Level에 따라 Account_Name 앞에 공백 특수문자 추가
    def format_name(row):
        indent = "    " * (int(row.get('Level', 3)) - 1) # 레벨 1=0칸, 2=4칸, 3=8칸
        prefix = "🔹 " if row.get('Level') == 1 else "   " 
        return f"{indent}{prefix}{row['Account_Name']}"

    if 'Level' in df.columns and 'Account_Name' in df.columns:
        df['Display_Name'] = df.apply(format_name, axis=1)

    return df