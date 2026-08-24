#!/usr/bin/env python3
import pandas as pd
import json

try:
    with open('config.json', 'r') as f:
        config = json.load(f)
        PLANT_CODES = config.get('plant_codes', ['742'])
        TARGET_PRODUCT = config.get('product_category', 'CRCA')
        P_CODES = config.get('p_codes', [])
        PROD_FILE = config.get('production_file', '742.XLSX')
        if isinstance(TARGET_PRODUCT, str):
            TARGET_PRODUCTS = [p.strip().upper() for p in TARGET_PRODUCT.split(',')]
        else:
            TARGET_PRODUCTS = [p.upper() for p in TARGET_PRODUCT]
except Exception:
    PLANT_CODES = ['742']
    TARGET_PRODUCTS = ['CRCA']
    P_CODES = []
    PROD_FILE = '742.XLSX'

import difflib
import sys
import warnings
warnings.filterwarnings('ignore')

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

def get_col(df, keywords):
    """Strict column finder to avoid picking up description columns"""
    for col in df.columns:
        cl = str(col).lower()
        if all(k in cl for k in keywords):
            return col
    return None

def clean_production_data(input_file, output_file):
    print(f"Reading {input_file}...")
    try:
        xls = pd.ExcelFile(input_file)
        data_sheet = None
        for sheet in xls.sheet_names:
            cols = pd.read_excel(xls, sheet_name=sheet, nrows=0).columns
            if any('posting' in str(c).lower() and 'date' in str(c).lower() for c in cols):
                data_sheet = sheet
                break
                
        if not data_sheet:
            print("CRITICAL ERROR: Could not find any sheet with a 'Posting Date' column.")
            sys.exit(1)
            
        print(f"Found raw data inside sheet: '{data_sheet}'")
        df = pd.read_excel(xls, sheet_name=data_sheet)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)
        
    # STRICT COLUMN MAPPING
    col_map = {
        'Product': get_col(df, ['material', 'description']) or get_col(df, ['product']),
        'P Code': get_col(df, ['product', 'code']) or get_col(df, ['p', 'code']),
        'Plant': get_col(df, ['plant', 'code']) or get_col(df, ['plant']),
        'Quantity': get_col(df, ['qty']) or get_col(df, ['quantity']),
        'Posting Date': get_col(df, ['posting', 'date']),
        'Customer Name': get_col(df, ['cust', 'name']),
        'SO No': get_col(df, ['so', 'no']) or get_col(df, ['so/str']),
        'Item No': get_col(df, ['item', 'no']),
        'Sales Org': get_col(df, ['sales', 'organization']) or get_col(df, ['sale', 'org'])
    }
    col_map = {k: v for k, v in col_map.items() if v is not None}

    print(f"Filtering for Plant {PLANT_CODES} and P Codes {P_CODES}...")
    
    if 'Sales Org' in col_map:
        df['Sales Org'] = df[col_map['Sales Org']].replace(r'^\s*$', 'free stocks', regex=True).fillna('free stocks')
    else:
        df['Sales Org'] = 'free stocks'
    
    # 1. Filter P Code strictly
    if 'P Code' in col_map:
        df['P Code Clean'] = df[col_map['P Code']].astype(str).str.strip().str.upper()
        df = df[df['P Code Clean'].isin(P_CODES)]
        
    # 2. Filter Plant Code perfectly (stripping .0 decimals and zeros)
    if 'Plant' in col_map:
        df['Plant Clean'] = df[col_map['Plant']].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.lstrip('0')
        df = df[df['Plant Clean'].isin(PLANT_CODES)]

    # 3. Clean quantities and parse true Production Month
    if 'Quantity' in col_map:
        df['Quantity'] = pd.to_numeric(df[col_map['Quantity']], errors='coerce').fillna(0)
    
    if 'Posting Date' in col_map:
        df['Posting Date'] = pd.to_datetime(df[col_map['Posting Date']], errors='coerce')
        df['Month Name'] = df['Posting Date'].dt.month_name()
    else:
        print("CRITICAL ERROR: Could not find Posting Date column.")
        sys.exit(1)

    # 4. Standardize overlapping Customer Names
    if 'Customer Name' in col_map:
        print("Standardizing overlapping customer names...")
        df['Clean Customer Name'] = standardize_names(df[col_map['Customer Name']])
        
    # 5. Create SO-Item Combo
    if 'SO No' in col_map and 'Item No' in col_map:
        df['SO-Item Combo'] = df[col_map['SO No']].astype(str).str.replace(r'\.0$', '', regex=True) + '-' + df[col_map['Item No']].astype(str).str.replace(r'\.0$', '', regex=True)

    print(f"Writing clean data to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        months = ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March']
        summary_data = []
        
        for m in months:
            m_df = df[df['Month Name'] == m]
            qty = m_df['Quantity'].sum() if 'Quantity' in col_map else 0
            summary_data.append({'Month': m, 'Total Production (MT)': qty, 'Rows of Data': len(m_df)})
            
            if len(m_df) > 0:
                m_df.to_excel(writer, sheet_name=m, index=False)
                
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="FY Summary", index=False)
        
    print(f"Phase 1 Complete! Total rows saved: {df.shape[0]}")

if __name__ == "__main__":
    input_file = PROD_FILE
    output_file = "Cleaned_Production_Data.xlsx"
    clean_production_data(input_file, output_file)