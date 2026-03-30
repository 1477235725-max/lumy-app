import pandas as pd
import numpy as np

file_path = 'data.xlsx'
output_file = 'analysis_report.txt'

print("开始分析...")

try:
    xl = pd.ExcelFile(file_path)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("地图付费和广告数据分析报告\n")
        f.write("="*80 + "\n\n")
        
        f.write(f"工作表列表: {xl.sheet_names}\n\n")
        
        all_sheets = pd.read_excel(file_path, sheet_name=None)
        
        for sheet_name, df in all_sheets.items():
            f.write("="*80 + "\n")
            f.write(f"工作表: {sheet_name}\n")
            f.write("="*80 + "\n")
            f.write(f"数据维度: {df.shape[0]} 行 × {df.shape[1]} 列\n\n")
            
            f.write("【字段列表】\n")
            for i, col in enumerate(df.columns, 1):
                f.write(f"  {i}. {col}\n")
            f.write("\n")
            
            f.write("【数据预览（前5行）】\n")
            f.write(df.head().to_string() + "\n\n")
            
            f.write("【数据类型】\n")
            for col, dtype in df.dtypes.items():
                null_count = df[col].isnull().sum()
                null_pct = (null_count / len(df)) * 100
                f.write(f"  {col:30s} {str(dtype):15s} (缺失: {null_count}, {null_pct:.1f}%)\n")
            f.write("\n")
            
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) > 0:
                f.write("【数值统计】\n")
                f.write(df[numeric_cols].describe().to_string() + "\n\n")
            
            f.write("【商业化分析】\n")
            payment_keywords = ['付费', '支付', '收入', 'revenue', 'payment', 'pay']
            payment_cols = [col for col in df.columns if any(keyword in str(col).lower() for keyword in payment_keywords)]
            
            if payment_cols:
                f.write(f"  发现付费相关字段: {payment_cols}\n")
                for col in payment_cols:
                    if df[col].dtype in [np.int64, np.float64]:
                        total = df[col].sum()
                        mean = df[col].mean()
                        f.write(f"    - {col}: 总计={total:,.2f}, 平均={mean:,.2f}\n")
            else:
                f.write("  未发现明确的付费相关字段\n")
            
            ad_keywords = ['广告', '投放', '曝光', '点击', 'ctr', 'cpm', 'ad', 'advertisement']
            ad_cols = [col for col in df.columns if any(keyword in str(col).lower() for keyword in ad_keywords)]
            
            if ad_cols:
                f.write(f"  发现广告相关字段: {ad_cols}\n")
                for col in ad_cols:
                    if df[col].dtype in [np.int64, np.float64]:
                        total = df[col].sum()
                        mean = df[col].mean()
                        f.write(f"    - {col}: 总计={total:,.2f}, 平均={mean:,.2f}\n")
            else:
                f.write("  未发现明确的广告相关字段\n")
            
            f.write("\n")
        
        f.write("="*80 + "\n")
        f.write("分析完成\n")
        f.write("="*80 + "\n")
    
    print("分析完成!")
    print(f"报告已保存到: {output_file}")
    
except Exception as e:
    print(f"错误: {e}")
    import traceback
    traceback.print_exc()
