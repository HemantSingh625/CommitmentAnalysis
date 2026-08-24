import pandas as pd
import glob
import os

def get_col(df, keywords):
    for col in df.columns:
        cl = str(col).lower()
        if all(k in cl for k in keywords):
            return col
    return None

for file in glob.glob('monthly_commitments/*.xlsx'):
    xls = pd.ExcelFile(file)
    best_score = -1
    best_sheet = None
    best_header_row = 0
    for sheet in xls.sheet_names:
        try:
            df_test = pd.read_excel(xls, sheet_name=sheet, nrows=25, header=None)
            for idx, row in df_test.iterrows():
                row_str = ' '.join([str(c).lower() for c in row if pd.notna(c)])
                keywords = ['plant', 'p. code', 'planned gr', 'customer', 'so no', 'item no', 'width', 'thk', 'vertical']
                score = sum(1 for k in keywords if k in row_str)
                if score > best_score:
                    best_score = score
                    best_sheet = sheet
                    best_header_row = idx
        except:
            pass
    if best_score < 3:
        continue
    df = pd.read_excel(file, sheet_name=best_sheet, header=best_header_row)
    p_code_col = get_col(df, ['p', 'code']) or get_col(df, ['p. code'])
    product_col = get_col(df, ['product']) or get_col(df, ['material', 'desc'])
    print(f"{os.path.basename(file)} | P Code: {p_code_col} | Product: {product_col}")
