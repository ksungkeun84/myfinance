import pandas as pd

# 엑셀 파일 읽기
file_path = "../data/7145601085-01.xlsx"
try:
    df_info = pd.read_excel(file_path, nrows=1, header=None)
    account_info = df_info.iloc[0,0]
    trade_period = df_info.iloc[0,1]
    df = pd.read_excel(file_path, header=1)

    rows = len(df)
    cols = len(df.columns)

    print(f'{account_info} | {trade_period}')
    print(f'rows: {rows} cols: {cols}')

    for r in range(0, rows):
        for c in range(0, cols):
            print(f'{df.iloc[r, c]}', end=" | ")
        print('\n')
except Exception as e:
    print(f"에러 발생: {e}")