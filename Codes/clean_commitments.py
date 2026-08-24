import pandas as pd
import glob
import os
import difflib
import json

try:
    with open('config.json', 'r') as f:
        config = json.load(f)
        PLANT_CODES = config.get('plant_codes', ['742'])
        TARGET_PRODUCT = config.get('product_category', 'CRCA')
        P_CODES = config.get('p_codes', [])
        if isinstance(TARGET_PRODUCT, str):
            TARGET_PRODUCTS = [p.strip().upper() for p in TARGET_PRODUCT.split(',')]
        else:
            TARGET_PRODUCTS = [p.upper() for p in TARGET_PRODUCT]
except Exception:
    PLANT_CODES = ['742']
    TARGET_PRODUCTS = ['CRCA']
    P_CODES = []

def standardize_names(names_series, threshold=0.90):
    """Merges customer names that are highly similar"""
    unique_names = names_series.dropna().unique()
    standard_map = {}
    
    for name in unique_names:
        if not str(name).strip():
            continue
        match_found = False
        for std_name in standard_map.values():
            if difflib.SequenceMatcher(None, str(name).upper(), str(std_name).upper()).ratio() >= threshold:
                standard_map[name] = std_name
                match_found = True
                break
        if not match_found:
            standard_map[name] = str(name).strip().upper()
            
    return names_series.map(standard_map).fillna("UNKNOWN")

def clean_commitments_data(input_dir, output_file):
    files = glob.glob(os.path.join(input_dir, '*.xlsx')) + glob.glob(os.path.join(input_dir, '*.xls'))
    files = [f for f in files if not os.path.basename(f).startswith('~')]
    
    if not files:
        print(f"No excel files found in {input_dir}")
        return

    print(f"Found {len(files)} commitment files.")
    
    all_clean_data = []

    for file in files:
        print(f"Processing {os.path.basename(file)}...")
        try:
            xls = pd.ExcelFile(file)
            best_sheet = None
            best_header_row = 0
            best_score = -1
            
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
                except Exception:
                    continue
            
            if best_score < 3:
                print(f"  Skipping {os.path.basename(file)}: Could not confidently identify header row.")
                continue
                
            print(f"  Best sheet: {best_sheet}, Header row: {best_header_row+1}")
            df = pd.read_excel(file, sheet_name=best_sheet, header=best_header_row)
            
            month = os.path.basename(file).split('_')[0].capitalize()
            if month not in ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March']:
                print(f"  WARNING: Could not parse month from filename '{os.path.basename(file)}'. Defaulting to filename.")
                month = os.path.basename(file).split('.')[0]
                
            def get_col(df, keywords):
                for col in df.columns:
                    cl = str(col).lower()
                    if all(k in cl for k in keywords):
                        return col
                return None
                
            col_map = {
                'Plant': get_col(df, ['plant']),
                'SO No': get_col(df, ['so', 'no']),
                'Item No': get_col(df, ['item', 'no']),
                'Combo': get_col(df, ['so-item']) or get_col(df, ['so', 'item', 'combo']),
                'Product': get_col(df, ['product']) or get_col(df, ['material', 'desc']),
                'P Code': get_col(df, ['p', 'code']) or get_col(df, ['p. code']),
                'Customer': get_col(df, ['customer', 'name']) or get_col(df, ['cust', 'name']),
                'Vertical': get_col(df, ['vertical']),
                'Thk': get_col(df, ['thk']) or get_col(df, ['thick']),
                'Width': get_col(df, ['width']),
                'Planned GR Target': get_col(df, ['planned', 'gr']) or get_col(df, ['total', 'gr']),
                'Sales Org': get_col(df, ['sale', 'org'])
            }
            
            col_map = {k: v for k, v in col_map.items() if v is not None}
            
            if 'Product' not in col_map or 'Planned GR Target' not in col_map:
                print(f"  Skipping {os.path.basename(file)}: Missing critical columns (Product or Planned GR Target).")
                continue

            if 'P Code' in col_map and len(P_CODES) > 0:
                df['P Code Clean'] = df[col_map['P Code']].astype(str).str.strip().str.upper()
                df = df[df['P Code Clean'].isin(P_CODES)]
            else:
                df['Product Clean'] = df[col_map['Product']].astype(str).str.strip().str.upper()
                df = df[df['Product Clean'].isin(TARGET_PRODUCTS)]
            
            if len(df) == 0:
                print(f"  Skipping {os.path.basename(file)}: No matching data found (filtered).")
                continue
                
            # FIX INDEX ALIGNMENT BUG: Reset index before creating a new dataframe
            df = df.reset_index(drop=True)
            
            clean_df = pd.DataFrame()
            clean_df['Month'] = [month] * len(df)

            # Do NOT filter by Plant for Commitments, the user wants the total CRCA GR
            if 'Plant' in col_map:
                clean_df['Plant'] = df[col_map['Plant']].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.lstrip('0')
            else:
                clean_df['Plant'] = 'Unknown'

            for tgt, src in col_map.items():
                if tgt == 'Product': continue
                if tgt == 'Plant': continue
                clean_df[tgt] = df[src].copy()
                
            if 'Sales Org' in clean_df.columns:
                clean_df['Sales Org'] = clean_df['Sales Org'].replace(r'^\s*$', 'free stocks', regex=True).fillna('free stocks')
            else:
                clean_df['Sales Org'] = 'free stocks'
                
            if 'Combo' not in clean_df.columns and 'SO No' in clean_df.columns and 'Item No' in clean_df.columns:
                clean_df['Combo'] = clean_df['SO No'].astype(str).str.replace(r'\.0$', '', regex=True) + '-' + clean_df['Item No'].astype(str).str.replace(r'\.0$', '', regex=True)
                clean_df.rename(columns={'Combo': 'SO-Item Combo'}, inplace=True)
            elif 'Combo' in clean_df.columns:
                clean_df.rename(columns={'Combo': 'SO-Item Combo'}, inplace=True)
                
            if 'SO-Item Combo' not in clean_df.columns:
                print(f"  Skipping {os.path.basename(file)}: Could not resolve SO-Item Combo.")
                continue
                
            clean_df['Planned GR Target'] = pd.to_numeric(clean_df['Planned GR Target'], errors='coerce').fillna(0)
            
            before_len = len(clean_df)
            clean_df = clean_df.drop_duplicates(subset=['SO-Item Combo'], keep='last')
            after_len = len(clean_df)
            
            if before_len - after_len > 0:
                print(f"  Removed {before_len - after_len} duplicate SO-Item combinations.")
                
            clean_df = clean_df[clean_df['Planned GR Target'] > 0]
            
            if 'Customer' in clean_df.columns:
                clean_df['Clean Customer Name'] = standardize_names(clean_df['Customer'])
            
            all_clean_data.append(clean_df)
            
        except Exception as e:
            print(f"  Error processing {file}: {e}")

    if not all_clean_data:
        print("No valid commitment data was extracted.")
        return
        
    master_df = pd.concat(all_clean_data, ignore_index=True)
    
    print(f"Writing clean data to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        months = ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March']
        summary_data = []
        
        for m in months:
            m_df = master_df[master_df['Month'] == m]
            qty = m_df['Planned GR Target'].sum()
            summary_data.append({'Month': m, 'Total Planned GR': qty, 'Rows': len(m_df)})
            
            if len(m_df) > 0:
                m_df.to_excel(writer, sheet_name=m, index=False)
                
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="FY Summary", index=False)
        master_df.to_excel(writer, sheet_name="Master Data - All Months", index=False)
        
    print(f"Phase 2 Complete! Total commitment rows saved: {len(master_df)}")

if __name__ == "__main__":
    input_dir = "monthly_commitments"
    output_file = "Cleaned_Commitments_Data.xlsx"
    clean_commitments_data(input_dir, output_file)
