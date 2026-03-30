#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
地图付费和广告商业化分析脚本
功能：对地图数据进行全面的商业化分析，包括付费、广告、用户行为等维度
"""

import pandas as pd
import numpy as np
import sys
from datetime import datetime

class CommercialAnalysis:
    """商业化分析类"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.xl = None
        self.all_sheets = {}
        self.output = []
        
    def load_data(self):
        """加载数据"""
        try:
            self.xl = pd.ExcelFile(self.file_path)
            self.all_sheets = pd.read_excel(self.file_path, sheet_name=None)
            self._add_separator()
            self._add_line("数据加载成功！")
            self._add_line(f"文件: {self.file_path}")
            self._add_line(f"工作表数量: {len(self.xl.sheet_names)}")
            self._add_line(f"工作表列表: {self.xl.sheet_names}")
            self._add_separator()
            self._add_line("")
            return True
        except Exception as e:
            self._add_line(f"数据加载失败: {e}")
            return False
    
    def analyze_sheet(self, sheet_name):
        """分析单个工作表"""
        df = self.all_sheets[sheet_name]
        
        self._add_separator()
        self._add_line(f"工作表: {sheet_name}")
        self._add_separator()
        self._add_line(f"数据维度: {df.shape[0]} 行 × {df.shape[1]} 列")
        self._add_line("")
        
        # 1. 基础信息分析
        self._analyze_basic_info(df)
        
        # 2. 数据质量分析
        self._analyze_data_quality(df)
        
        # 3. 付费分析
        self._analyze_payment(df)
        
        # 4. 广告分析
        self._analyze_advertising(df)
        
        # 5. 排名分析
        self._analyze_ranking(df)
        
        # 6. 用户行为分析
        self._analyze_user_behavior(df)
        
        # 7. 商业指标分析
        self._analyze_business_metrics(df)
        
        self._add_line("")
    
    def _analyze_basic_info(self, df):
        """基础信息分析"""
        self._add_line("【1. 基础信息分析】")
        self._add_line(f"字段数量: {len(df.columns)}")
        self._add_line("")
        self._add_line("字段列表:")
        for i, col in enumerate(df.columns, 1):
            self._add_line(f"  {i:2d}. {col}")
        self._add_line("")
        
        # 数据类型分布
        self._add_line("数据类型分布:")
        dtype_counts = df.dtypes.value_counts()
        for dtype, count in dtype_counts.items():
            self._add_line(f"  {str(dtype):20s}: {count} 个字段")
        self._add_line("")
    
    def _analyze_data_quality(self, df):
        """数据质量分析"""
        self._add_line("【2. 数据质量分析】")
        
        # 缺失值分析
        missing = df.isnull().sum()
        missing = missing[missing > 0].sort_values(ascending=False)
        
        if len(missing) > 0:
            self._add_line(f"发现 {len(missing)} 个字段存在缺失值:")
            for col, count in missing.items():
                pct = (count / len(df)) * 100
                if pct > 10:
                    self._add_line(f"    ⚠ {col:30s}: {count:5d} ({pct:5.1f}%) - 缺失率较高")
                else:
                    self._add_line(f"    ✓ {col:30s}: {count:5d} ({pct:5.1f}%)")
        else:
            self._add_line("✓ 无缺失值，数据质量良好")
        self._add_line("")
        
        # 重复值分析
        duplicates = df.duplicated().sum()
        if duplicates > 0:
            self._add_line(f"⚠ 发现 {duplicates} 行重复数据")
        else:
            self._add_line("✓ 无重复数据")
        self._add_line("")
    
    def _analyze_payment(self, df):
        """付费分析"""
        self._add_line("【3. 付费分析】")
        
        payment_keywords = ['付费', '支付', '收入', '收益', '充值', '金额', '价格', 
                          'revenue', 'payment', 'pay', 'income', 'price', 'cost']
        payment_cols = [col for col in df.columns 
                       if any(keyword in str(col).lower() for keyword in payment_keywords)]
        
        if payment_cols:
            self._add_line(f"✓ 发现 {len(payment_cols)} 个付费相关字段:")
            self._add_line("")
            
            for col in payment_cols:
                self._add_line(f"  字段: {col}")
                if pd.api.types.is_numeric_dtype(df[col]):
                    non_null = df[col].dropna()
                    if len(non_null) > 0:
                        total = non_null.sum()
                        mean = non_null.mean()
                        median = non_null.median()
                        std = non_null.std()
                        min_val = non_null.min()
                        max_val = non_null.max()
                        
                        self._add_line(f"    统计信息:")
                        self._add_line(f"      总计: {total:,.2f}")
                        self._add_line(f"      平均: {mean:,.2f}")
                        self._add_line(f"      中位数: {median:,.2f}")
                        self._add_line(f"      标准差: {std:,.2f}")
                        self._add_line(f"      最小值: {min_val:,.2f}")
                        self._add_line(f"      最大值: {max_val:,.2f}")
                        
                        # 付费转化率分析
                        paid_users = len(non_null[non_null > 0])
                        total_users = len(df)
                        conversion_rate = (paid_users / total_users) * 100
                        self._add_line(f"      付费用户数: {paid_users}")
                        self._add_line(f"      付费转化率: {conversion_rate:.2f}%")
                else:
                    self._add_line(f"    类型: {df[col].dtype}")
                    self._add_line(f"    唯一值数量: {df[col].nunique()}")
                self._add_line("")
        else:
            self._add_line("⚠ 未发现明确的付费相关字段")
            self._add_line("")
    
    def _analyze_advertising(self, df):
        """广告分析"""
        self._add_line("【4. 广告分析】")
        
        ad_keywords = ['广告', '投放', '曝光', '点击', '展示', 'ctr', 'cpm', 'cpc',
                      'ad', 'advertisement', 'impression', 'click']
        ad_cols = [col for col in df.columns 
                  if any(keyword in str(col).lower() for keyword in ad_keywords)]
        
        if ad_cols:
            self._add_line(f"✓ 发现 {len(ad_cols)} 个广告相关字段:")
            self._add_line("")
            
            for col in ad_cols:
                self._add_line(f"  字段: {col}")
                if pd.api.types.is_numeric_dtype(df[col]):
                    non_null = df[col].dropna()
                    if len(non_null) > 0:
                        total = non_null.sum()
                        mean = non_null.mean()
                        median = non_null.median()
                        
                        self._add_line(f"    统计信息:")
                        self._add_line(f"      总计: {total:,.2f}")
                        self._add_line(f"      平均: {mean:,.2f}")
                        self._add_line(f"      中位数: {median:,.2f}")
                        
                        # CTR计算（如果有点击和曝光数据）
                        if '点击' in col or 'click' in col.lower():
                            exposure_cols = [c for c in df.columns 
                                          if '曝光' in c or 'impression' in c.lower()]
                            if exposure_cols:
                                exposure_col = exposure_cols[0]
                                clicks = non_null.sum()
                                impressions = df[exposure_col].sum()
                                if impressions > 0:
                                    ctr = (clicks / impressions) * 100
                                    self._add_line(f"      点击率 (CTR): {ctr:.2f}%")
                else:
                    self._add_line(f"    类型: {df[col].dtype}")
                    self._add_line(f"    唯一值数量: {df[col].nunique()}")
                self._add_line("")
        else:
            self._add_line("⚠ 未发现明确的广告相关字段")
            self._add_line("")
    
    def _analyze_ranking(self, df):
        """排名分析"""
        self._add_line("【5. 排名分析】")
        
        rank_keywords = ['排名', 'rank', '顺序', '名次']
        rank_cols = [col for col in df.columns 
                    if any(keyword in str(col).lower() for keyword in rank_keywords)]
        
        if rank_cols:
            for col in rank_cols:
                self._add_line(f"  排名字段: {col}")
                if pd.api.types.is_numeric_dtype(df[col]):
                    # Top 10 分析
                    top_10 = df.nsmallest(10, col)
                    self._add_line(f"  Top 10 数据:")
                    for idx, row in top_10.iterrows():
                        row_str = " | ".join([str(val)[:20] if val is not None else '' 
                                            for val in row.values[:5]])
                        self._add_line(f"    {row_str}")
                    
                    # 排名分布
                    rank_dist = df[col].value_counts().sort_index()
                    if len(rank_dist) > 0:
                        self._add_line(f"  排名分布:")
                        self._add_line(f"    最小排名: {rank_dist.index.min()}")
                        self._add_line(f"    最大排名: {rank_dist.index.max()}")
                        self._add_line(f"    平均排名: {df[col].mean():.2f}")
                self._add_line("")
        else:
            self._add_line("⚠ 未发现排名相关字段")
            self._add_line("")
    
    def _analyze_user_behavior(self, df):
        """用户行为分析"""
        self._add_line("【6. 用户行为分析】")
        
        user_keywords = ['用户', '玩家', 'player', 'user', '活跃', 'active', 'play']
        user_cols = [col for col in df.columns 
                    if any(keyword in str(col).lower() for keyword in user_keywords)]
        
        if user_cols:
            self._add_line(f"✓ 发现 {len(user_cols)} 个用户相关字段:")
            for col in user_cols:
                self._add_line(f"    - {col}")
                if pd.api.types.is_numeric_dtype(df[col]):
                    non_null = df[col].dropna()
                    if len(non_null) > 0:
                        total = non_null.sum()
                        mean = non_null.mean()
                        self._add_line(f"      总计: {total:,.0f}, 平均: {mean:,.2f}")
            
            # 活跃度分析
            active_keywords = ['活跃', 'active', 'play', '游玩']
            active_cols = [col for col in user_cols 
                          if any(keyword in str(col).lower() for keyword in active_keywords)]
            if active_cols:
                self._add_line("")
                self._add_line("  活跃度分析:")
                for col in active_cols:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        non_null = df[col].dropna()
                        if len(non_null) > 0:
                            avg_active = non_null.mean()
                            self._add_line(f"    平均{col}: {avg_active:,.2f}")
        else:
            self._add_line("⚠ 未发现用户相关字段")
        self._add_line("")
    
    def _analyze_business_metrics(self, df):
        """商业指标分析"""
        self._add_line("【7. 商业指标综合分析】")
        
        metrics = {}
        
        # 收入指标
        revenue_keywords = ['收入', 'revenue', '收益', 'earning']
        revenue_cols = [col for col in df.columns 
                       if any(keyword in str(col).lower() for keyword in revenue_keywords)]
        
        if revenue_cols and pd.api.types.is_numeric_dtype(df[revenue_cols[0]]):
            total_revenue = df[revenue_cols[0]].sum()
            metrics['总收入'] = total_revenue
        
        # ARPU分析
        user_keywords = ['用户', 'user', 'player']
        user_cols = [col for col in df.columns 
                    if any(keyword in str(col).lower() for keyword in user_keywords)]
        
        if user_cols and pd.api.types.is_numeric_dtype(df[user_cols[0]]):
            total_users = df[user_cols[0]].sum()
            metrics['总用户数'] = total_users
            
            if '总收入' in metrics and total_users > 0:
                arpu = metrics['总收入'] / total_users
                metrics['ARPU'] = arpu
        
        # 留存相关
        retention_keywords = ['留存', 'retention', '留存率']
        retention_cols = [col for col in df.columns 
                         if any(keyword in str(col).lower() for keyword in retention_keywords)]
        
        if retention_cols:
            metrics['留存字段'] = retention_cols
        
        if metrics:
            self._add_line("核心商业指标:")
            for key, value in metrics.items():
                if isinstance(value, float):
                    self._add_line(f"  {key:10s}: {value:,.2f}")
                elif isinstance(value, list):
                    self._add_line(f"  {key:10s}: {value}")
                else:
                    self._add_line(f"  {key:10s}: {value}")
        else:
            self._add_line("⚠ 未发现明确的商业指标字段")
        
        self._add_line("")
        
        # 商业化建议
        self._add_line("【商业化建议】")
        if 'ARPU' in metrics:
            arpu = metrics['ARPU']
            if arpu < 10:
                self._add_line("  • ARPU较低，建议优化付费点和定价策略")
            elif arpu < 50:
                self._add_line("  • ARPU处于中等水平，可尝试提升付费深度")
            else:
                self._add_line("  • ARPU表现良好，可拓展付费场景")
        
        if revenue_cols:
            self._add_line("  • 建议分析收入构成，优化高收入地图的推广策略")
        
        if ad_cols := [col for col in df.columns 
                      if any(keyword in str(col).lower() for keyword in ['广告', 'ad', 'click'])]:
            self._add_line("  • 建议优化广告位展示，提升广告收入")
        
        self._add_line("  • 建议进行用户分层运营，提升付费转化率")
        self._add_line("  • 建议关注高排名地图的商业化表现，形成标杆")
        self._add_line("")
    
    def generate_summary(self):
        """生成总结报告"""
        self._add_separator()
        self._add_line("【综合分析总结】")
        self._add_separator()
        
        total_rows = sum(df.shape[0] for df in self.all_sheets.values())
        total_cols = sum(df.shape[1] for df in self.all_sheets.values())
        
        self._add_line(f"数据概览:")
        self._add_line(f"  • 工作表数量: {len(self.all_sheets)}")
        self._add_line(f"  • 总数据行数: {total_rows:,}")
        self._add_line(f"  • 总字段数量: {total_cols}")
        self._add_line("")
        
        self._add_line(f"分析完成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self._add_line("")
        self._add_separator()
    
    def _add_line(self, line):
        """添加一行到输出"""
        self.output.append(line)
    
    def _add_separator(self, char="="):
        """添加分隔符"""
        self.output.append(char * 80)
    
    def save_report(self, output_file):
        """保存报告"""
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(self.output))
        print(f"报告已保存到: {output_file}")


def main():
    """主函数"""
    file_path = 'data.xlsx'
    output_file = 'detailed_commercial_analysis.txt'
    
    print("=" * 80)
    print("地图付费和广告商业化分析")
    print("=" * 80)
    print()
    
    # 创建分析对象
    analyzer = CommercialAnalysis(file_path)
    
    # 加载数据
    if not analyzer.load_data():
        print("数据加载失败！")
        sys.exit(1)
    
    # 分析每个工作表
    for sheet_name in analyzer.all_sheets.keys():
        print(f"正在分析工作表: {sheet_name}...")
        analyzer.analyze_sheet(sheet_name)
    
    # 生成总结
    analyzer.generate_summary()
    
    # 保存报告
    analyzer.save_report(output_file)
    
    print()
    print("=" * 80)
    print("分析完成！")
    print("=" * 80)


if __name__ == "__main__":
    main()
