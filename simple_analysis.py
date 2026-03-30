#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
星世界地图商业化分析 - 简化版
直接运行并保存结果，无需输出捕获
"""

import pandas as pd
import numpy as np

def analyze_and_save():
    """分析数据并保存结果"""
    try:
        # 读取数据
        xl = pd.ExcelFile('data.xlsx')
        all_sheets = pd.read_excel('data.xlsx', sheet_name=None)
        
        # 准备报告内容
        report = []
        report.append("="*80)
        report.append("星世界地图商业化数据分析报告")
        report.append("="*80)
        report.append(f"分析时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("")
        
        # 分析每个工作表
        for sheet_name, df in all_sheets.items():
            report.append("-"*80)
            report.append(f"工作表: {sheet_name}")
            report.append("-"*80)
            report.append(f"数据维度: {df.shape[0]} 行 × {df.shape[1]} 列")
            report.append("")
            
            # 字段列表
            report.append("【字段列表】")
            for i, col in enumerate(df.columns, 1):
                report.append(f"  {i:2d}. {col}")
            report.append("")
            
            # 数据预览
            report.append("【数据预览（前5行）】")
            report.append(df.head().to_string())
            report.append("")
            
            # 数据类型
            report.append("【数据类型】")
            for col, dtype in df.dtypes.items():
                null_count = df[col].isnull().sum()
                null_pct = (null_count / len(df)) * 100
                report.append(f"  {col:30s} {str(dtype):15s} (缺失: {null_count}, {null_pct:.1f}%)")
            report.append("")
            
            # 付费分析
            report.append("【付费分析】")
            payment_keywords = ['付费', '支付', '收入', 'revenue', 'payment', 'pay']
            payment_cols = [col for col in df.columns 
                           if any(keyword in str(col).lower() for keyword in payment_keywords)]
            
            if payment_cols:
                report.append(f"  发现付费相关字段: {payment_cols}")
                for col in payment_cols:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        non_null = df[col].dropna()
                        if len(non_null) > 0:
                            total = non_null.sum()
                            mean = non_null.mean()
                            median = non_null.median()
                            paid_users = len(non_null[non_null > 0])
                            conversion_rate = (paid_users / len(df)) * 100
                            report.append(f"    字段: {col}")
                            report.append(f"      总计: {total:,.2f}")
                            report.append(f"      平均: {mean:,.2f}")
                            report.append(f"      中位数: {median:,.2f}")
                            report.append(f"      付费用户: {paid_users}")
                            report.append(f"      转化率: {conversion_rate:.2f}%")
            else:
                report.append("  未发现明确的付费相关字段")
            report.append("")
            
            # 广告分析
            report.append("【广告分析】")
            ad_keywords = ['广告', '投放', '曝光', '点击', 'ctr', 'cpm', 'ad', 'impression']
            ad_cols = [col for col in df.columns 
                      if any(keyword in str(col).lower() for keyword in ad_keywords)]
            
            if ad_cols:
                report.append(f"  发现广告相关字段: {ad_cols}")
                for col in ad_cols:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        non_null = df[col].dropna()
                        if len(non_null) > 0:
                            total = non_null.sum()
                            mean = non_null.mean()
                            report.append(f"    字段: {col}")
                            report.append(f"      总计: {total:,.2f}")
                            report.append(f"      平均: {mean:,.2f}")
            else:
                report.append("  未发现明确的广告相关字段")
            report.append("")
            
            # 排名分析
            report.append("【排名分析】")
            rank_keywords = ['排名', 'rank', '名次']
            rank_cols = [col for col in df.columns 
                        if any(keyword in str(col).lower() for keyword in rank_keywords)]
            
            if rank_cols:
                for col in rank_cols:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        avg_rank = df[col].mean()
                        min_rank = df[col].min()
                        max_rank = df[col].max()
                        report.append(f"  排名字段: {col}")
                        report.append(f"    平均排名: {avg_rank:.2f}")
                        report.append(f"    最小排名: {min_rank:.0f}")
                        report.append(f"    最大排名: {max_rank:.0f}")
            else:
                report.append("  未发现排名相关字段")
            report.append("")
        
        report.append("="*80)
        report.append("分析完成")
        report.append("="*80)
        
        # 保存报告
        with open('商业化分析报告.txt', 'w', encoding='utf-8') as f:
            f.write('\n'.join(report))
        
        # 同时保存简化版
        with open('analysis_completed.txt', 'w', encoding='utf-8') as f:
            f.write('分析完成！\n')
            f.write(f'工作表数量: {len(all_sheets)}\n')
            f.write('详细报告请查看: 商业化分析报告.txt\n')
        
        return True
        
    except Exception as e:
        with open('error_log.txt', 'w', encoding='utf-8') as f:
            f.write(f'错误: {e}\n')
            import traceback
            traceback.print_exc(file=f)
        return False

if __name__ == "__main__":
    success = analyze_and_save()
    if success:
        with open('run_status.txt', 'w', encoding='utf-8') as f:
            f.write('SUCCESS')
    else:
        with open('run_status.txt', 'w', encoding='utf-8') as f:
            f.write('FAILED')
