import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font

INPUT_FILE = 'FY27_Targets_and_Achievements.xlsx'
OUTPUT_FILE = 'FY27 Half AOP Target vs. Achievement.xlsx'

excel_file = pd.ExcelFile(INPUT_FILE)

# Custom geographic mapping rules
CUSTOM_GEO_MAP = {
    'abeokuta north': {'Zone': 'Lagos', 'Region': 'Ogun West'},
    'ijebu east': {'Zone': 'Lagos', 'Region': 'Ogun East'}
}

def clean_area_names(series):
    """
    Standardizes Area names across target and achievement datasets:
    1. Replaces 'Ihala' with 'Ihiala'.
    2. Replaces any area string containing 'Igabi' with 'Kaduna North'.
    """
    s = series.astype(str).str.strip()
    
    # 1. Map Ihala -> Ihiala
    s = s.replace({'Ihala': 'Ihiala', 'ihala': 'Ihiala'})
    
    # 2. Map any Area containing 'Igabi' -> Kaduna North
    mask_igabi = s.str.lower().str.contains('igabi', na=False)
    s.loc[mask_igabi] = 'Kaduna North'
    
    return s

def apply_custom_geo_mapping(df):
    """Overrides Zone and Region mappings for specified areas regardless of raw source values."""
    if 'Area' in df.columns:
        mask = df['Area'].astype(str).str.strip().str.lower().isin(CUSTOM_GEO_MAP.keys())
        for idx in df[mask].index:
            area_key = str(df.loc[idx, 'Area']).strip().lower()
            if 'Zone' in df.columns:
                df.loc[idx, 'Zone'] = CUSTOM_GEO_MAP[area_key]['Zone']
            if 'Region' in df.columns:
                df.loc[idx, 'Region'] = CUSTOM_GEO_MAP[area_key]['Region']
    return df

def clean_and_filter_columns(df):
    """Excludes unwanted columns like Unit Sales, WP, Watt-peak, etc."""
    drop_patterns = ['unit sales', 'unit_sales', 'wp', 'watt', 'unit', 'sales_unit']
    cols_to_keep = [
        col for col in df.columns 
        if not any(pattern in str(col).lower() for pattern in drop_patterns)
    ]
    return df[cols_to_keep]

def process_reconciliation(target_sheet, ach_sheet, header_offset=1):
    df_target = pd.read_excel(excel_file, sheet_name=target_sheet, header=header_offset)
    df_achievement = pd.read_excel(excel_file, sheet_name=ach_sheet)

    # Clean headers
    df_target.columns = [str(c).strip() for c in df_target.columns]
    df_achievement.columns = [str(c).strip() for c in df_achievement.columns]

    # Filter out Unit Sales / WP columns if present
    df_target = clean_and_filter_columns(df_target)
    df_achievement = clean_and_filter_columns(df_achievement)

    # Clean and standardize Area strings (Ihala -> Ihiala, Igabi -> Kaduna North)
    df_target['Area'] = clean_area_names(df_target['Area'])
    df_achievement['Area'] = clean_area_names(df_achievement['Area'])

    # Filter out invalid / summary / variance rows
    invalid_keywords = [
        'total', 'grand total', 'subtotal', 'nan', 'none', '0', '', 'null', 
        'unnamed', 'difference', 'variance', 'check', 'balance', 'unit sales', 'wp'
    ]
    df_target = df_target[~df_target['Area'].str.lower().isin(invalid_keywords) & df_target['Area'].notna()].copy()
    df_achievement = df_achievement[~df_achievement['Area'].str.lower().isin(invalid_keywords) & df_achievement['Area'].notna()].copy()

    # Apply custom geographic classification (Abeokuta North / Ijebu East)
    df_target = apply_custom_geo_mapping(df_target)
    df_achievement = apply_custom_geo_mapping(df_achievement)

    # Dynamic target total calculation
    total_col_name = [c for c in df_target.columns if 'total' in c.lower()]
    if total_col_name:
        target_series = pd.to_numeric(df_target[total_col_name[0]], errors='coerce')
    else:
        target_series = pd.Series(np.nan, index=df_target.index)

    # Monthly sum fallback (columns index 3 through 14)
    monthly_cols = df_target.columns[3:15]
    calculated_totals = df_target[monthly_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    df_target['Target'] = target_series.fillna(calculated_totals)
    
    # Preserve Target geographic mappings
    geo_cols = [c for c in ['Zone', 'Region', 'Area'] if c in df_target.columns]
    df_target_clean = df_target[geo_cols + ['Target']].drop_duplicates(subset=['Area'])

    # Achievement cleanup
    ach_total_col = [c for c in df_achievement.columns if 'total' in c.lower()][0]
    ach_geo_cols = [c for c in ['Zone', 'Region', 'Area'] if c in df_achievement.columns]
    df_ach_clean = df_achievement[ach_geo_cols + [ach_total_col]].copy()
    df_ach_clean.rename(columns={ach_total_col: 'Achievement'}, inplace=True)
    df_ach_clean['Achievement'] = pd.to_numeric(df_ach_clean['Achievement'], errors='coerce').fillna(0)

    # Full Merge (Outer Join to preserve 100% of achievements)
    merged_df = pd.merge(df_ach_clean, df_target_clean, on='Area', how='outer', suffixes=('_ach', '_tgt'))

    # Coalesce Zone and Region columns across target and achievement
    for col in ['Zone', 'Region']:
        if f'{col}_ach' in merged_df.columns and f'{col}_tgt' in merged_df.columns:
            merged_df[col] = merged_df[f'{col}_ach'].fillna(merged_df[f'{col}_tgt'])
            merged_df.drop(columns=[f'{col}_ach', f'{col}_tgt'], inplace=True)
        elif f'{col}_tgt' in merged_df.columns:
            merged_df.rename(columns={f'{col}_tgt': col}, inplace=True)
        elif f'{col}_ach' in merged_df.columns:
            merged_df.rename(columns={f'{col}_ach': col}, inplace=True)

    # Enforce custom geo mapping on final merged output
    merged_df = apply_custom_geo_mapping(merged_df)

    merged_df['Target'] = pd.to_numeric(merged_df['Target'], errors='coerce').fillna(0)
    merged_df['Achievement'] = pd.to_numeric(merged_df['Achievement'], errors='coerce').fillna(0)

    # Filter out residual rows where both Target and Achievement are <= 0 or invalid
    merged_df = merged_df[(merged_df['Target'] > 0) | (merged_df['Achievement'] > 0)].copy()

    # Assign Status Flags
    def assign_status(row):
        tgt = row['Target']
        ach = row['Achievement']
        
        if tgt > 0 and ach > 0:
            return 'Target Achieved' if ach >= tgt else 'Target Not Achieved'
        elif tgt > 0 and ach == 0:
            return 'AHQ Not Launched'
        elif tgt == 0 and ach > 0:
            return 'Missing from Target Sheet'
        return 'Unmapped'

    merged_df['Status'] = merged_df.apply(assign_status, axis=1)

    # Calculate metrics
    merged_df['Deficit'] = merged_df['Target'] - merged_df['Achievement']
    merged_df['% Achieved'] = (
        (merged_df['Achievement'] / merged_df['Target'])
        .replace([np.inf, -np.inf], 0)
        .fillna(0) * 100
    ).round(2)

    cols_order = ['Zone', 'Region', 'Area', 'Target', 'Achievement', 'Deficit', '% Achieved', 'Status']
    merged_df = merged_df[cols_order]

    # Helper function for Zone & Region summaries
    def build_summary(df, group_col):
        summary = df.groupby(group_col, as_index=False).agg({'Target': 'sum', 'Achievement': 'sum'})
        summary['Deficit'] = summary['Target'] - summary['Achievement']
        summary['% Achieved'] = (
            (summary['Achievement'] / summary['Target'])
            .replace([np.inf, -np.inf], 0)
            .fillna(0) * 100
        ).round(2)
        
        g_target = summary['Target'].sum()
        g_ach = summary['Achievement'].sum()
        grand_row = pd.DataFrame([{
            group_col: 'Grand Total',
            'Target': g_target,
            'Achievement': g_ach,
            'Deficit': g_target - g_ach,
            '% Achieved': round((g_ach / g_target * 100) if g_target != 0 else 0, 2)
        }])
        return pd.concat([summary, grand_row], ignore_index=True)

    zone_summary = build_summary(merged_df, 'Zone')
    region_summary = build_summary(merged_df, 'Region')

    return merged_df, zone_summary, region_summary

# Configuration of product lines: (Target Sheet, Achievement Sheet, Output Sheet Name)
PRODUCT_CONFIGS = [
    ('Target', 'Achievement', 'Reconciliation Analysis'),
    ('Inverter Target', 'Inverter Achievement', 'Inverter AOP vs Target'),
    ('PayG Phone Target', 'PayG Phone Achievement', 'PayG Phone AOP vs Target')
]

# Style definitions for status formatting
green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid') # Soft Green
green_font = Font(color='006100', bold=True)

red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')   # Soft Red
red_font = Font(color='9C0006', bold=True)

# Write all tabs to the Excel workbook
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    results_summary = []
    
    for tgt_sheet, ach_sheet, out_sheet_name in PRODUCT_CONFIGS:
        df_detail, zone_sum, reg_sum = process_reconciliation(tgt_sheet, ach_sheet, header_offset=1)
        
        # 1. Main Area Detail Table on the left (Columns A-H)
        df_detail.to_excel(writer, sheet_name=out_sheet_name, startcol=0, index=False)
        
        # 2. Zone Pivot Summary on top right (Columns J-N, row 0)
        zone_sum.to_excel(writer, sheet_name=out_sheet_name, startcol=9, startrow=0, index=False)
        
        # 3. Region Pivot Summary directly beneath Zone Summary (3 blank rows gap)
        reg_sum.to_excel(writer, sheet_name=out_sheet_name, startcol=9, startrow=len(zone_sum) + 3, index=False)
        
        results_summary.append((out_sheet_name, df_detail, zone_sum, reg_sum))

    wb = writer.book
    
    # Format each sheet in the workbook
    for sheet_name, df, z_sum, r_sum in results_summary:
        ws = wb[sheet_name]
        
        # Apply number formatting and color highlights to Area Detail table
        for row_idx in range(2, len(df) + 2):
            # Target Number format '#,##0' (Column D)
            ws[f'D{row_idx}'].number_format = '#,##0'
            
            # Status conditional formatting (Column H)
            status_cell = ws[f'H{row_idx}']
            val = str(status_cell.value)
            
            if val == 'Target Achieved':
                status_cell.fill = green_fill
                status_cell.font = green_font
            elif val in ['Target Not Achieved', 'AHQ Not Launched']:
                status_cell.fill = red_fill
                status_cell.font = red_font

        # Format Pivot Summary Targets (Column K)
        for row_idx in range(2, len(z_sum) + 2):
            ws[f'K{row_idx}'].number_format = '#,##0'
            
        reg_start = len(z_sum) + 5
        for row_idx in range(reg_start, reg_start + len(r_sum)):
            ws[f'K{row_idx}'].number_format = '#,##0'

print("--- RECONCILIATION SUMMARY ---")
for sheet_name, df, _, _ in results_summary:
    print(f"{sheet_name} Total Achievement: {df['Achievement'].sum():,.2f}")

print(f"\nSuccess! Applied area alias cleaning ('Ihala' -> 'Ihiala', 'Igabi' -> 'Kaduna North') and saved to '{OUTPUT_FILE}'.")