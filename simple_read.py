import pandas as pd

file_path = 'data.xlsx'

try:
    xl = pd.ExcelFile(file_path)
    with open('data_summary.txt', 'w', encoding='utf-8') as f:
        f.write('Sheet names: ' + str(xl.sheet_names) + '\n\n')
        
        all_sheets = pd.read_excel(file_path, sheet_name=None)
        
        for sheet_name, df in all_sheets.items():
            f.write(f'=== Sheet: {sheet_name} ===\n')
            f.write(f'Shape: {df.shape}\n')
            f.write(f'Columns: {list(df.columns)}\n')
            f.write(f'\nFirst 5 rows:\n')
            f.write(df.head().to_string())
            f.write(f'\n\nData types:\n')
            f.write(df.dtypes.to_string())
            f.write(f'\n\nBasic statistics:\n')
            f.write(df.describe().to_string())
            f.write('\n\n' + '='*80 + '\n\n')
            
    print('SUCCESS: Data summary created')
    print('Check data_summary.txt for details')
    
except Exception as e:
    print(f'ERROR: {e}')
    with open('error_info.txt', 'w', encoding='utf-8') as f:
        f.write('ERROR: ' + str(e) + '\n')
        import traceback
        traceback.print_exc(file=f)
