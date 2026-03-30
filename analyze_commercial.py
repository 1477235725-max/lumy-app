import pandas as pd
import numpy as np

file_path = 'data.xlsx'

try:
    xl = pd.ExcelFile(file_path)
    
    output = []
    output.append("="*80)
    output.append("地图付费和广告数据分析报告")
    output.append("="*80)
    output.append("")
    output.append(f"工作表列表: {xl.sheet_names}")
    output.append("")
    
    all_sheets = pd.read_excel(file_path, sheet_name=None)
    
    for sheet_name, df in all_sheets.items():
        output.append("="*80)
        output.append(f"工作表: {sheet_name}")
        output.append("="*80)
        output.append(f"数据维度: {df.shape[0]} 行 × {df.shape[1]} 列")
        output.append("")
        
        output.append("【字段列表】")
        for i, col in enumerate(df.columns, 1):
            output.append(f"  {i}. {col}")
        output.append("")
        
        output.append("【数据预览（前3行）】")
        output.append(df.head(3).to_string())
        output.append("")
        
        output.append("【数据类型】")
        for col, dtype in df.dtypes.items():
            null_count = df[col].isnull().sum()
            null_pct = (null_count / len(df)) * 100
            output.append(f"  {col:30s} {str(dtype):15s} (缺失: {null_count}, {null_pct:.1f}%)")
        output.append("")
        
        output.append("【数值统计】")
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            output.append(df[numeric_cols].describe().to_string())
        else:
            output.append("  无数值型字段")
        output.append("")
        
        output.append("【商业化分析】")
        
        payment_keywords = ['付费', '支付', '收入', 'revenue', 'payment', 'pay']
        payment_cols = [col for col in df.columns if any(keyword in str(col).lower() for keyword in payment_keywords)]
        
        if payment_cols:
            output.append(f"  发现付费相关字段: {payment_cols}")
            for col in payment_cols:
                if df[col].dtype in [np.int64, np.float64]:
                    total = df[col].sum()
                    mean = df[col].mean()
                    output.append(f"    - {col}: 总计={total:,.2f}, 平均={mean:,.2f}")
        else:
            output.append("  未发现明确的付费相关字段")
        
        ad_keywords = ['广告', '投放', '曝光', '点击', 'ctr', 'cpm', 'ad', 'advertisement']
        ad_cols = [col for col in df.columns if any(keyword in str(col).lower() for keyword in ad_keywords)]
        
        if ad_cols:
            output.append(f"  发现广告相关字段: {ad_cols}")
            for col in ad_cols:
                if df[col].dtype in [np.int64, np.float64]:
                    total = df[col].sum()
                    mean = df[col].mean()
                    output.append(f"    - {col}: 总计={total:,.2f}, 平均={mean:,.2f}")
        else:
            output.append("  未发现明确的广告相关字段")
        
        output.append("")
    
    output.append("="*80)
    output.append("分析完成！")
    output.append("="*80)
    
    with open('commercial_analysis.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(output))
    
    print('SUCCESS: Analysis completed')
    print('Output saved to commercial_analysis.txt')
    
except Exception as e:
    print(f'ERROR: {e}')
    with open('error_log.txt', 'w', encoding='utf-8') as f:
        f.write(f'ERROR: {e}\n')
        import traceback
        traceback.print_exc(file=f)
