import pandas as pd
from openpyxl import load_workbook

def detect_merged_headers(ws):
    """
    병합 셀 정보를 통해 다단계 컬럼 헤더가 있는지 감지
    """
    header_rows = set()
    for merged in ws.merged_cells.ranges:
        for row in range(merged.min_row, merged.max_row + 1):
            header_rows.add(row - 1)
    if not header_rows:
        return [0]
    
    # 연속된 병합된 행만 반환
    header_rows = sorted(header_rows)
    contiguous = []
    for r in header_rows:
        if not contiguous or r == contiguous[-1] + 1:
            contiguous.append(r)
        else:
            break
    return contiguous if contiguous else [0]

def is_row_index(df):
    """
    첫 번째 열이 row 이름으로 사용될 수 있는지 판단
    - 첫 컬럼이 Unnamed 이거나, 중복이 없고 문자열이면 row로 간주
    """
    first_col = df.columns[0]
    unique_values = df[first_col].dropna().unique()
    return df[first_col].dtype == object and len(unique_values) == len(df)

def flatten_multi_columns(df):
    """MultiIndex 컬럼을 단일 문자열로 평탄화"""
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ['_'.join([str(i) for i in col if str(i) != 'nan']) for col in df.columns.values]
    return df

def parse_sheet(file_path, sheet_name):
    # 우선 header=0 으로 불러옴
    df_try = pd.read_excel(file_path, sheet_name=sheet_name, header=0, nrows=5)

    # Unnamed 컬럼이 많다면 멀티헤더일 가능성 → header=[0,1] 시도
    if df_try.columns.to_series().str.contains("Unnamed").sum() > len(df_try.columns) // 2:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=[0,1])
            df = flatten_multi_columns(df)
        except Exception:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

    # row index 추정
    if is_row_index(df):
        df = df.set_index(df.columns[0])

    return df

def main():
    pd.set_option('display.max_columns', None)
    pd.set_option('display.expand_frame_repr', False)

    file_path = 'input/parsinginput.xlsx'
    wb = load_workbook(file_path, read_only=True)
    sheet_names = wb.sheetnames
    dfs = {}

    for name in sheet_names:
        df = parse_sheet(file_path, sheet_name=name)
        dfs[name] = df
        print(f"\n===== [시트: {name}] =====")
        print(df.to_string())
        print("=" * 40)

if __name__ == "__main__":
    main()
