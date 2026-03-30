import pandas as pd
import sys
import traceback

file_path = r'C:\Users\lumyzeng\Documents\WXWork\1688858225700575\Cache\File\2026-02\游玩排名前300地图 付费&广告.xlsx'

try:
    print("Starting analysis...", file=sys.stderr)
    
    xl = pd.ExcelFile(file_path)
    print("Excel file loaded successfully", file=sys.stderr)
    
    output = []
    output.append("Sheet names: " + str(xl.sheet_names) + "\n\n")
    output.append("="*80 + "\n\n")
    
    all_sheets = pd.read_excel(file_path, sheet_name=None)
    
    for sheet_name, df in all_sheets.items():
        output.append(f"Sheet: {sheet_name}\n")
        output.append(f"Shape: {df.shape}\n")
        output.append(f"Columns: {df.columns.tolist()}\n")
        output.append(f"\nFirst 5 rows:\n")
        output.append(df.head().to_string() + "\n")
        output.append(f"\nData types:\n")
        output.append(df.dtypes.to_string() + "\n")
        output.append(f"\nBasic statistics:\n")
        output.append(df.describe().to_string() + "\n")
        output.append("\n" + "="*80 + "\n\n")
        
    with open('analysis_result.txt', 'w', encoding='utf-8') as f:
        f.writelines(output)
        
    print("Analysis completed successfully!", file=sys.stderr)
    print("Results saved to analysis_result.txt", file=sys.stderr)
    
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
    traceback.print_exc(file=sys.stderr)
    with open('error_log.txt', 'w', encoding='utf-8') as f:
        f.write(f"Error: {e}\n")
        traceback.print_exc(file=f)
