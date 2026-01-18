import pandas as pd
from google import genai
import json
import openpyxl
import io

def extract_sheet_data(file):
    """
    엑셀 파일을 읽어서 '텍스트 데이터'로 변환 (AI에게 구조를 통째로 넘기기 위함)
    """
    context_list = []
    wb = openpyxl.load_workbook(file, data_only=True)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.sheet_state == 'hidden' or ws.sheet_state == 'veryHidden':
            continue
        
        data = ws.values
        try:
            # 첫 줄을 헤더로 가정
            header = next(data)
            # 빈 컬럼 제외
            columns = [str(h) if h is not None else f"Unnamed_{i}" for i, h in enumerate(header)]
            
            # 데이터프레임 생성
            df = pd.DataFrame(data, columns=columns)
            
            # 너무 많은 빈 행/열 제거
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # CSV 텍스트로 변환
            csv_text = df.to_csv(index=False)
            context_list.append(f"FileName: {file.name} | Sheet: {sheet_name}\n{csv_text}")
        except StopIteration:
            continue
            
    return "\n\n".join(context_list)

def process_smart_merge(api_key, target_files):
    """
    여러 파일의 데이터를 AI에게 주어 '문맥 기반 병합' 수행
    """
    # 1. 모든 엑셀 데이터를 텍스트로 추출
    full_context = ""
    for file in target_files:
        full_context += extract_sheet_data(file) + "\n\n"
    
    # 토큰 제한 (약 15만자 - Gemini 1.5 Flash/Pro는 충분함)
    if len(full_context) > 150000:
        full_context = full_context[:150000] + "\n...(Data Truncated)"

    # 2. Gemini Client 생성
    client = genai.Client(api_key=api_key)

    # 3. 프롬프트: 순서 보존과 끼워넣기(Interleaving) 로직 강조
    prompt = f"""
    You are a specialized Financial Data Merger.
    User provided multiple financial statements (Excel data) from different years or entities.
    
    [Goal]
    Merge all data into a SINGLE table where rows are "Account Items" and columns are "Years" (e.g., 2022, 2023, 2024).

    [Crucial Logic: "Context-Aware Interleaving"]
    1. **Preserve Order (No Alphabetical Sort):** Do NOT sort account names alphabetically. Keep the logical flow of the original files (e.g., Assets -> Liabilities -> Equity).
    2. **Insert Missing Items:** - If File A has [Sales, Operating Profit] and File B has [Sales, COGS, Operating Profit], the result must be [Sales, COGS, Operating Profit].
       - You must detect where a missing account fits based on its neighbors in other files.
    3. **Unify Synonyms:** If File A says "급여" and File B says "임직원급여", merge them into one row (choose the most standard name).
    4. **Columns:** Detect years from the data (headers or values) and create columns like '2022', '2023'.

    [Input Data]
    {full_context}

    [Output Format]
    Return ONLY a JSON Array of objects.
    Example:
    [
      {{
        "Account_Name": "매출액",
        "2022": 1000,
        "2023": 1200
      }},
      {{
        "Account_Name": "매출원가",
        "2022": 0,  <-- If missing in 2022, fill with 0
        "2023": 500
      }}
    ]
    """

    # 4. AI 호출
    try:
        response = client.models.generate_content(
            model="gemini-3-flash-preview", # 없으면 gemini-1.5-flash로 자동 변경 로직은 app.py나 여기서 처리
            contents=prompt
        )
    except Exception:
        # 모델명 에러시 fallback
        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=prompt
        )

    # 5. 결과 파싱
    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
    
    # JSON 부분만 추출
    if "[" in cleaned_text and "]" in cleaned_text:
        s = cleaned_text.find("[")
        e = cleaned_text.rfind("]") + 1
        cleaned_text = cleaned_text[s:e]
    
    data_list = json.loads(cleaned_text)
    df = pd.DataFrame(data_list)
    
    return df