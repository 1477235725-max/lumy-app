#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
星世界地图商业化数据可视化分析
功能：生成可视化图表和商业化效果分析报告
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

class StarWorldCommercialVisualization:
    """星世界地图商业化可视化分析类"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.xl = None
        self.all_sheets = {}
        self.figures = []
        
    def load_data(self):
        """加载数据"""
        try:
            self.xl = pd.ExcelFile(self.file_path)
            self.all_sheets = pd.read_excel(self.file_path, sheet_name=None)
            print(f"数据加载成功！")
            print(f"  工作表数量: {len(self.xl.sheet_names)}")
            print(f"  工作表列表: {self.xl.sheet_names}")
            return True
        except Exception as e:
            print(f"数据加载失败: {e}")
            return False
    
    def analyze_and_visualize(self):
        """分析并可视化数据"""
        for sheet_name in self.all_sheets.keys():
            print(f"\n{'='*80}")
            print(f"正在分析工作表: {sheet_name}")
            print(f"{'='*80}")
            
            df = self.all_sheets[sheet_name]
            print(f"数据维度: {df.shape[0]} 行 × {df.shape[1]} 列")
            print(f"字段列表: {list(df.columns)}")
            
            self._create_visualizations(sheet_name, df)
    
    def _create_visualizations(self, sheet_name, df):
        """创建可视化图表"""
        print(f"\n生成可视化图表...")
        
        sns.set_style("whitegrid")
        sns.set_palette("husl")
        
        self._plot_overview(sheet_name, df)
        self._plot_payment_analysis(sheet_name, df)
        self._plot_advertising_analysis(sheet_name, df)
        self._plot_ranking_analysis(sheet_name, df)
        self._plot_commercial_dashboard(sheet_name, df)
    
    def _plot_overview(self, sheet_name, df):
        """数据概览可视化"""
        print("  生成数据概览图表...")
        
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        fig.suptitle(f'{sheet_name} - 数据概览', fontsize=16, fontweight='bold')
        
        # 数据类型分布
        dtype_counts = df.dtypes.value_counts()
        axes[0, 0].pie(dtype_counts.values, labels=dtype_counts.index, autopct='%1.1f%%')
        axes[0, 0].set_title('数据类型分布', fontsize=12, fontweight='bold')
        
        # 缺失值分析
        missing = df.isnull().sum()
        missing = missing[missing > 0].sort_values(ascending=False)
        if len(missing) > 0:
            y_pos = range(len(missing))
            axes[0, 1].barh(y_pos, missing.values, color='coral')
            axes[0, 1].set_yticks(y_pos)
            axes[0, 1].set_yticklabels(missing.index, fontsize=8)
            axes[0, 1].set_xlabel('缺失值数量')
            axes[0, 1].set_title('字段缺失值分析', fontsize=12, fontweight='bold')
        else:
            axes[0, 1].text(0.5, 0.5, '无缺失值', ha='center', va='center', 
                          transform=axes[0, 1].transAxes, fontsize=14)
            axes[0, 1].set_title('字段缺失值分析', fontsize=12, fontweight='bold')
        
        # 数值型字段统计
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            data_ranges = [df[col].max() - df[col].min() for col in numeric_cols]
            axes[1, 0].bar(range(len(numeric_cols)), data_ranges, color='skyblue')
            axes[1, 0].set_xticks(range(len(numeric_cols)))
            axes[1, 0].set_xticklabels(numeric_cols, rotation=45, ha='right', fontsize=8)
            axes[1, 0].set_ylabel('数据范围')
            axes[1, 0].set_title('数值型字段数据范围', fontsize=12, fontweight='bold')
        else:
            axes[1, 0].text(0.5, 0.5, '无数值型字段', ha='center', va='center',
                          transform=axes[1, 0].transAxes, fontsize=14)
            axes[1, 0].set_title('数值型字段数据范围', fontsize=12, fontweight='bold')
        
        # 数据质量评分
        quality_score = 100
        if len(missing) > 0:
            missing_ratio = missing.sum() / (df.shape[0] * df.shape[1])
            quality_score -= missing_ratio * 100
        
        duplicates = df.duplicated().sum()
        duplicate_ratio = duplicates / df.shape[0]
        quality_score -= duplicate_ratio * 10
        
        quality_score = max(0, quality_score)
        
        categories = ['数据完整性', '数据唯一性', '数据准确性']
        scores = [100 - (missing.sum() / (df.shape[0] * df.shape[1])) * 100 if len(missing) > 0 else 100,
                 100 - duplicate_ratio * 10 if duplicates > 0 else 100,
                 quality_score]
        
        colors = ['#2ecc71' if s >= 80 else '#f1c40f' if s >= 60 else '#e74c3c' for s in scores]
        axes[1, 1].bar(categories, scores, color=colors)
        axes[1, 1].set_ylim(0, 100)
        axes[1, 1].set_ylabel('评分')
        axes[1, 1].set_title('数据质量评分', fontsize=12, fontweight='bold')
        for i, v in enumerate(scores):
            axes[1, 1].text(i, v + 2, f'{v:.1f}', ha='center', fontweight='bold')
        
        plt.tight_layout()
        filename = f'{sheet_name}_01_数据概览.png'
        plt.savefig(filename, dpi=300, bbox_inches='tight')
        self.figures.append(filename)
        plt.close()
        print(f"    {filename}")
    
    def _plot_payment_analysis(self, sheet_name, df):
        """付费分析可视化"""
        print("  生成付费分析图表...")
        
        payment_keywords = ['付费', '支付', '收入', 'revenue', 'payment', 'pay']
        payment_cols = [col for col in df.columns 
                       if any(keyword in str(col).lower() for keyword in payment_keywords)]
        
        if not payment_cols:
            print("    未发现付费相关字段，跳过付费分析")
            return
        
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        fig.suptitle(f'{sheet_name} - 付费分析', fontsize=16, fontweight='bold')
        
        numeric_payment_cols = [col for col in payment_cols if pd.api.types.is_numeric_dtype(df[col])]
        
        if numeric_payment_cols:
            # 付费金额分布
            col = numeric_payment_cols[0]
            data = df[col].dropna()
            axes[0, 0].hist(data, bins=30, color='lightblue', edgecolor='black')
            axes[0, 0].set_xlabel('付费金额')
            axes[0, 0].set_ylabel('频次')
            axes[0, 0].set_title(f'{col} - 金额分布', fontsize=12, fontweight='bold')
            axes[0, 0].axvline(data.mean(), color='red', linestyle='--', linewidth=2, label=f'均值: {data.mean():.2f}')
            axes[0, 0].legend()
            
            # 付费总额对比
            totals = [df[col].sum() for col in numeric_payment_cols]
            axes[0, 1].bar(numeric_payment_cols, totals, color='steelblue')
            axes[0, 1].set_xticklabels(numeric_payment_cols, rotation=45, ha='right', fontsize=9)
            axes[0, 1].set_ylabel('总额')
            axes[0, 1].set_title('各付费字段总额对比', fontsize=12, fontweight='bold')
            for i, v in enumerate(totals):
                axes[0, 1].text(i, v + (max(totals)*0.02), f'{v:,.0f}', ha='center', fontsize=8)
            
            # 付费用户比例
            paid_ratios = []
            for col in numeric_payment_cols:
                paid_users = len(df[df[col] > 0])
                total_users = len(df)
                ratio = (paid_users / total_users) * 100
                paid_ratios.append(ratio)
            
            colors_pie = ['#ff9999' if r < 20 else '#66b3ff' if r < 50 else '#99ff99' for r in paid_ratios]
            wedges, texts, autotexts = axes[1, 0].pie(paid_ratios, labels=numeric_payment_cols, 
                                                      autopct='%1.1f%%', colors=colors_pie)
            for autotext in autotexts:
                autotext.set_color('black')
                autotext.set_fontweight('bold')
            axes[1, 0].set_title('付费用户比例', fontsize=12, fontweight='bold')
            
            # 付费趋势
            if len(df) > 10:
                axes[1, 1].plot(range(len(df)), df[numeric_payment_cols[0]].values, 
                               marker='o', linewidth=2, markersize=4)
                axes[1, 1].set_xlabel('数据索引')
                axes[1, 1].set_ylabel(f'{numeric_payment_cols[0]}')
                axes[1, 1].set_title(f'{numeric_payment_cols[0]} - 趋势图', fontsize=12, fontweight='bold')
                axes[1, 1].grid(True, alpha=0.3)
            else:
                axes[1, 1].text(0.5, 0.5, '数据量不足，无法绘制趋势图', 
                               ha='center', va='center', transform=axes[1, 1].transAxes)
        else:
            for ax in axes.flat:
                ax.text(0.5, 0.5, '无数值型付费字段', ha='center', va='center',
                       transform=ax.transAxes, fontsize=14)
        
        plt.tight_layout()
        filename = f'{sheet_name}_02_付费分析.png'
        plt.savefig(filename, dpi=300, bbox_inches='tight')
        self.figures.append(filename)
        plt.close()
        print(f"    {filename}")
    
    def _plot_advertising_analysis(self, sheet_name, df):
        """广告分析可视化"""
        print("  生成广告分析图表...")
        
        ad_keywords = ['广告', '投放', '曝光', '点击', 'ctr', 'cpm', 'ad', 'impression']
        ad_cols = [col for col in df.columns 
                  if any(keyword in str(col).lower() for keyword in ad_keywords)]
        
        if not ad_cols:
            print("    未发现广告相关字段，跳过广告分析")
            return
        
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        fig.suptitle(f'{sheet_name} - 广告分析', fontsize=16, fontweight='bold')
        
        numeric_ad_cols = [col for col in ad_cols if pd.api.types.is_numeric_dtype(df[col])]
        
        if numeric_ad_cols:
            # 广告指标对比
            means = [df[col].mean() for col in numeric_ad_cols]
            colors_ad = ['#ff6b6b', '#4ecdc4', '#45b7d1', '#96ceb4']
            axes[0, 0].bar(numeric_ad_cols, means, color=colors_ad[:len(numeric_ad_cols)])
            axes[0, 0].set_xticklabels(numeric_ad_cols, rotation=45, ha='right', fontsize=9)
            axes[0, 0].set_ylabel('平均值')
            axes[0, 0].set_title('广告指标平均值对比', fontsize=12, fontweight='bold')
            
            # 点击率分析
            if 'ctr' in str([col.lower() for col in numeric_ad_cols]):
                ctr_col = [col for col in numeric_ad_cols if 'ctr' in col.lower()][0]
                ctr_data = df[ctr_col].dropna()
                axes[0, 1].hist(ctr_data, bins=20, color='lightgreen', edgecolor='black')
                axes[0, 1].set_xlabel('点击率 (%)')
                axes[0, 1].set_ylabel('频次')
                axes[0, 1].set_title('点击率分布', fontsize=12, fontweight='bold')
                axes[0, 1].axvline(ctr_data.mean(), color='red', linestyle='--', 
                                   linewidth=2, label=f'均值: {ctr_data.mean():.2f}%')
                axes[0, 1].legend()
            else:
                axes[0, 1].text(0.5, 0.5, '未找到CTR数据', ha='center', va='center',
                               transform=axes[0, 1].transAxes, fontsize=14)
            
            # 广告性能对比
            if len(numeric_ad_cols) >= 2:
                x = np.arange(len(df.head(20)))
                width = 0.35
                axes[1, 0].bar(x - width/2, df[numeric_ad_cols[0]].head(20), width, 
                             label=numeric_ad_cols[0], color='skyblue')
                axes[1, 0].bar(x + width/2, df[numeric_ad_cols[1]].head(20), width, 
                             label=numeric_ad_cols[1], color='lightcoral')
                axes[1, 0].set_xlabel('样本索引')
                axes[1, 0].set_ylabel('数值')
                axes[1, 0].set_title('前20个样本广告指标对比', fontsize=12, fontweight='bold')
                axes[1, 0].legend()
                axes[1, 0].grid(True, alpha=0.3)
            else:
                axes[1, 0].text(0.5, 0.5, '数据不足，无法对比', ha='center', va='center',
                               transform=axes[1, 0].transAxes, fontsize=14)
            
            # 广告效果热力图
            if len(numeric_ad_cols) > 1:
                corr_matrix = df[numeric_ad_cols].corr()
                im = axes[1, 1].imshow(corr_matrix, cmap='coolwarm', aspect='auto', vmin=-1, vmax=1)
                axes[1, 1].set_xticks(range(len(numeric_ad_cols)))
                axes[1, 1].set_yticks(range(len(numeric_ad_cols)))
                axes[1, 1].set_xticklabels(numeric_ad_cols, rotation=45, ha='right', fontsize=8)
                axes[1, 1].set_yticklabels(numeric_ad_cols, fontsize=8)
                axes[1, 1].set_title('广告指标相关性热力图', fontsize=12, fontweight='bold')
                
                for i in range(len(numeric_ad_cols)):
                    for j in range(len(numeric_ad_cols)):
                        text = axes[1, 1].text(j, i, f'{corr_matrix.iloc[i, j]:.2f}',
                                             ha="center", va="center", color="black", fontsize=9)
            else:
                axes[1, 1].text(0.5, 0.5, '数据不足，无法绘制热力图', ha='center', va='center',
                               transform=axes[1, 1].transAxes, fontsize=14)
        else:
            for ax in axes.flat:
                ax.text(0.5, 0.5, '无数值型广告字段', ha='center', va='center',
                       transform=ax.transAxes, fontsize=14)
        
        plt.tight_layout()
        filename = f'{sheet_name}_03_广告分析.png'
        plt.savefig(filename, dpi=300, bbox_inches='tight')
        self.figures.append(filename)
        plt.close()
        print(f"    {filename}")
    
    def _plot_ranking_analysis(self, sheet_name, df):
        """排名分析可视化"""
        print("  生成排名分析图表...")
        
        rank_keywords = ['排名', 'rank', '名次']
        rank_cols = [col for col in df.columns 
                    if any(keyword in str(col).lower() for keyword in rank_keywords)]
        
        if not rank_cols:
            print("    未发现排名相关字段，跳过排名分析")
            return
        
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        fig.suptitle(f'{sheet_name} - 排名分析', fontsize=16, fontweight='bold')
        
        rank_col = rank_cols[0]
        
        if pd.api.types.is_numeric_dtype(df[rank_col]):
            # 排名分布
            axes[0, 0].hist(df[rank_col], bins=20, color='lightblue', edgecolor='black')
            axes[0, 0].set_xlabel('排名')
            axes[0, 0].set_ylabel('频次')
            axes[0, 0].set_title('排名分布', fontsize=12, fontweight='bold')
            
            # Top 10 性能对比
            if len(df.columns) > 2:
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                other_numeric = [col for col in numeric_cols if col != rank_col]
                
                if other_numeric:
                    top_10 = df.nsmallest(10, rank_col)
                    x = np.arange(10)
                    width = 0.35
                    
                    axes[0, 1].bar(x - width/2, top_10[other_numeric[0]].values, 
                                 width, label=other_numeric[0], color='skyblue')
                    axes[0, 1].set_xticks(x)
                    axes[0, 1].set_xticklabels(top_10[rank_col].values, rotation=45, fontsize=8)
                    axes[0, 1].set_ylabel(other_numeric[0])
                    axes[0, 1].set_title('Top 10 排名对应指标', fontsize=12, fontweight='bold')
                    axes[0, 1].legend()
                    axes[0, 1].grid(True, alpha=0.3)
                else:
                    axes[0, 1].text(0.5, 0.5, '无其他数值字段', ha='center', va='center',
                                   transform=axes[0, 1].transAxes, fontsize=14)
            else:
                axes[0, 1].text(0.5, 0.5, '数据字段不足', ha='center', va='center',
                               transform=axes[0, 1].transAxes, fontsize=14)
            
            # 排名与关键指标关系
            if len(df.columns) > 2:
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                other_numeric = [col for col in numeric_cols if col != rank_col]
                
                if other_numeric:
                    axes[1, 0].scatter(df[rank_col], df[other_numeric[0]], 
                                      alpha=0.6, color='coral')
                    axes[1, 0].set_xlabel('排名')
                    axes[1, 0].set_ylabel(other_numeric[0])
                    axes[1, 0].set_title(f'排名 vs {other_numeric[0]}', fontsize=12, fontweight='bold')
                    axes[1, 0].grid(True, alpha=0.3)
                    
                    z = np.polyfit(df[rank_col], df[other_numeric[0]], 1)
                    p = np.poly1d(z)
                    axes[1, 0].plot(df[rank_col], p(df[rank_col]), 
                                  "r--", alpha=0.8, linewidth=2)
                else:
                    axes[1, 0].text(0.5, 0.5, '无其他数值字段', ha='center', va='center',
                                   transform=axes[1, 0].transAxes, fontsize=14)
            else:
                axes[1, 0].text(0.5, 0.5, '数据字段不足', ha='center', va='center',
                               transform=axes[1, 0].transAxes, fontsize=14)
            
            # 排名分段统计
            rank_bins = pd.cut(df[rank_col], bins=[0, 50, 100, 150, 200, 300, float('inf')],
                              labels=['1-50', '51-100', '101-150', '151-200', '201-300', '300+'])
            rank_counts = rank_bins.value_counts().sort_index()
            colors_rank = ['#e74c3c', '#e67e22', '#f1c40f', '#3498db', '#2ecc71', '#9b59b6']
            axes[1, 1].bar(rank_counts.index, rank_counts.values, color=colors_rank)
            axes[1, 1].set_xlabel('排名区间')
            axes[1, 1].set_ylabel('数量')
            axes[1, 1].set_title('排名分段统计', fontsize=12, fontweight='bold')
            for i, v in enumerate(rank_counts.values):
                axes[1, 1].text(i, v + 0.5, str(v), ha='center', fontweight='bold')
        else:
            for ax in axes.flat:
                ax.text(0.5, 0.5, '排名字段非数值型', ha='center', va='center',
                       transform=ax.transAxes, fontsize=14)
        
        plt.tight_layout()
        filename = f'{sheet_name}_04_排名分析.png'
        plt.savefig(filename, dpi=300, bbox_inches='tight')
        self.figures.append(filename)
        plt.close()
        print(f"    {filename}")
    
    def _plot_commercial_dashboard(self, sheet_name, df):
        """综合商业化分析仪表板"""
        print("  生成商业化分析仪表板...")
        
        fig = plt.figure(figsize=(20, 16))
        gs = fig.add_gridspec(3, 3, hspace=0.3, wspace=0.3)
        fig.suptitle(f'{sheet_name} - 星世界地图商业化综合分析', 
                     fontsize=18, fontweight='bold')
        
        payment_keywords = ['付费', '支付', '收入', 'revenue', 'payment']
        ad_keywords = ['广告', 'ad', 'impression', 'click']
        user_keywords = ['用户', 'user', 'player', '活跃']
        rank_keywords = ['排名', 'rank']
        
        payment_cols = [col for col in df.columns 
                       if any(keyword in str(col).lower() for keyword in payment_keywords)]
        ad_cols = [col for col in df.columns 
                  if any(keyword in str(col).lower() for keyword in ad_keywords)]
        user_cols = [col for col in df.columns 
                    if any(keyword in str(col).lower() for keyword in user_keywords)]
        rank_cols = [col for col in df.columns 
                    if any(keyword in str(col).lower() for keyword in rank_keywords)]
        
        # 关键指标卡片
        metrics_data = []
        metrics_labels = []
        
        if payment_cols and pd.api.types.is_numeric_dtype(df[payment_cols[0]]):
            total_revenue = df[payment_cols[0]].sum()
            metrics_data.append(total_revenue)
            metrics_labels.append('总收入')
        
        if user_cols and pd.api.types.is_numeric_dtype(df[user_cols[0]]):
            total_users = df[user_cols[0]].sum()
            metrics_data.append(total_users)
            metrics_labels.append('总用户数')
        
        if rank_cols and pd.api.types.is_numeric_dtype(df[rank_cols[0]]):
            avg_rank = df[rank_cols[0]].mean()
            metrics_data.append(avg_rank)
            metrics_labels.append('平均排名')
        
        if metrics_data:
            ax1 = fig.add_subplot(gs[0, 0])
            colors_metrics = ['#e74c3c', '#3498db', '#2ecc71'][:len(metrics_data)]
            bars = ax1.bar(metrics_labels, metrics_data, color=colors_metrics)
            ax1.set_title('核心商业指标', fontsize=12, fontweight='bold')
            ax1.set_ylabel('数值')
            
            for bar, value in zip(bars, metrics_data):
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height,
                        f'{value:,.0f}', ha='center', va='bottom', fontweight='bold')
        else:
            ax1 = fig.add_subplot(gs[0, 0])
            ax1.text(0.5, 0.5, '数据不足', ha='center', va='center',
                    transform=ax1.transAxes, fontsize=14)
        
        # 商业化效果评分
        ax2 = fig.add_subplot(gs[0, 1])
        categories = ['付费能力', '广告效果', '用户活跃', '排名表现']
        scores = []
        
        if payment_cols and pd.api.types.is_numeric_dtype(df[payment_cols[0]]):
            revenue = df[payment_cols[0]].sum()
            score = min(100, (revenue / 100000) * 100)
            scores.append(score)
        else:
            scores.append(0)
        
        if ad_cols and len(ad_cols) >= 2:
            scores.append(75)
        else:
            scores.append(0)
        
        if user_cols and pd.api.types.is_numeric_dtype(df[user_cols[0]]):
            users = df[user_cols[0]].sum()
            score = min(100, (users / 10000) * 100)
            scores.append(score)
        else:
            scores.append(0)
        
        if rank_cols and pd.api.types.is_numeric_dtype(df[rank_cols[0]]):
            avg_rank = df[rank_cols[0]].mean()
            score = max(0, 100 - (avg_rank / 3))
            scores.append(score)
        else:
            scores.append(0)
        
        colors_score = ['#2ecc71' if s >= 70 else '#f1c40f' if s >= 40 else '#e74c3c' for s in scores]
        bars = ax2.bar(categories, scores, color=colors_score)
        ax2.set_ylim(0, 100)
        ax2.set_title('商业化效果评分', fontsize=12, fontweight='bold')
        ax2.set_ylabel('评分')
        for bar, score in zip(bars, scores):
            ax2.text(bar.get_x() + bar.get_width()/2., score + 2,
                    f'{score:.0f}', ha='center', fontweight='bold')
        
        # 收入构成
        ax3 = fig.add_subplot(gs[0, 2])
        if len(payment_cols) > 1:
            numeric_payment = [col for col in payment_cols if pd.api.types.is_numeric_dtype(df[col])]
            if numeric_payment:
                revenues = [df[col].sum() for col in numeric_payment]
                ax3.pie(revenues, labels=numeric_payment, autopct='%1.1f%%')
                ax3.set_title('收入构成', fontsize=12, fontweight='bold')
            else:
                ax3.text(0.5, 0.5, '数据不足', ha='center', va='center',
                        transform=ax3.transAxes, fontsize=14)
        else:
            ax3.text(0.5, 0.5, '单一收入来源', ha='center', va='center',
                    transform=ax3.transAxes, fontsize=14)
        
        # 用户活跃度趋势
        ax4 = fig.add_subplot(gs[1, 0])
        if user_cols and pd.api.types.is_numeric_dtype(df[user_cols[0]]):
            ax4.plot(range(len(df)), df[user_cols[0]].values, 
                    marker='o', linewidth=2, markersize=3, color='steelblue')
            ax4.set_xlabel('样本索引')
            ax4.set_ylabel(user_cols[0])
            ax4.set_title('用户活跃度趋势', fontsize=12, fontweight='bold')
            ax4.grid(True, alpha=0.3)
        else:
            ax4.text(0.5, 0.5, '无用户数据', ha='center', va='center',
                    transform=ax4.transAxes, fontsize=14)
        
        # ARPU 分析
        ax5 = fig.add_subplot(gs[1, 1])
        if payment_cols and user_cols:
            payment_col = [col for col in payment_cols if pd.api.types.is_numeric_dtype(df[col])][0]
            user_col = [col for col in user_cols if pd.api.types.is_numeric_dtype(df[col])][0]
            
            arpu_list = []
            for idx in range(len(df)):
                revenue = df[payment_col].iloc[idx] if pd.notnull(df[payment_col].iloc[idx]) else 0
                users = df[user_col].iloc[idx] if pd.notnull(df[user_col].iloc[idx]) else 1
                if users > 0:
                    arpu_list.append(revenue / users)
            
            if arpu_list:
                ax5.hist(arpu_list, bins=20, color='lightcoral', edgecolor='black')
                ax5.set_xlabel('ARPU')
                ax5.set_ylabel('频次')
                ax5.set_title('ARPU 分布', fontsize=12, fontweight='bold')
                ax5.axvline(np.mean(arpu_list), color='red', linestyle='--', 
                           linewidth=2, label=f'均值: {np.mean(arpu_list):.2f}')
                ax5.legend()
            else:
                ax5.text(0.5, 0.5, '数据不足', ha='center', va='center',
                        transform=ax5.transAxes, fontsize=14)
        else:
            ax5.text(0.5, 0.5, '缺少必要字段', ha='center', va='center',
                    transform=ax5.transAxes, fontsize=14)
        
        # 付费转化率
        ax6 = fig.add_subplot(gs[1, 2])
        if payment_cols and pd.api.types.is_numeric_dtype(df[payment_cols[0]]):
            paid_users = len(df[df[payment_cols[0]] > 0])
            total_users = len(df)
            conversion_rate = (paid_users / total_users) * 100
            
            categories_conv = ['付费用户', '非付费用户']
            values_conv = [paid_users, total_users - paid_users]
            colors_conv = ['#e74c3c', '#3498db']
            
            ax6.pie(values_conv, labels=categories_conv, autopct='%1.1f%%', colors=colors_conv)
            ax6.set_title(f'付费转化率: {conversion_rate:.2f}%', fontsize=12, fontweight='bold')
        else:
            ax6.text(0.5, 0.5, '无付费数据', ha='center', va='center',
                    transform=ax6.transAxes, fontsize=14)
        
        # 排名与收入关系
        ax7 = fig.add_subplot(gs[2, 0])
        if rank_cols and payment_cols:
            rank_col = rank_cols[0]
            payment_col = [col for col in payment_cols if pd.api.types.is_numeric_dtype(df[col])][0]
            
            if pd.api.types.is_numeric_dtype(df[rank_col]):
                ax7.scatter(df[rank_col], df[payment_col], alpha=0.6, color='purple')
                ax7.set_xlabel('排名')
                ax7.set_ylabel(payment_col)
                ax7.set_title('排名 vs 收入', fontsize=12, fontweight='bold')
                ax7.grid(True, alpha=0.3)
            else:
                ax7.text(0.5, 0.5, '排名非数值型', ha='center', va='center',
                        transform=ax7.transAxes, fontsize=14)
        else:
            ax7.text(0.5, 0.5, '缺少数据', ha='center', va='center',
                    transform=ax7.transAxes, fontsize=14)
        
        # 商业化建议
        ax8 = fig.add_subplot(gs[2, 1])
        ax8.axis('off')
        
        suggestions = []
        if payment_cols:
            suggestions.append("• 优化付费点设计")
            suggestions.append("• 调整定价策略")
        if ad_cols:
            suggestions.append("• 提升广告位价值")
            suggestions.append("• 优化广告展示")
        if user_cols:
            suggestions.append("• 增强用户粘性")
            suggestions.append("• 提升活跃度")
        if rank_cols:
            suggestions.append("• 关注头部地图表现")
            suggestions.append("• 提升整体排名")
        
        if suggestions:
            ax8.text(0.1, 0.9, '商业化建议：', fontsize=12, fontweight='bold',
                    transform=ax8.transAxes, va='top')
            for i, suggestion in enumerate(suggestions[:8]):
                ax8.text(0.1, 0.85 - i*0.08, suggestion, fontsize=10,
                        transform=ax8.transAxes, va='top')
        else:
            ax8.text(0.5, 0.5, '建议数据完善后再分析', ha='center', va='center',
                    transform=ax8.transAxes, fontsize=14)
        
        # 综合评分
        ax9 = fig.add_subplot(gs[2, 2])
        overall_score = np.mean(scores) if scores else 0
        
        if len(scores) >= 3:
            angles = np.linspace(0, 2 * np.pi, len(scores), endpoint=False).tolist()
            scores_radar = scores + [scores[0]]
            angles += angles[:1]
            
            ax9 = plt.subplot(gs[2, 2], projection='polar')
            ax9.plot(angles, scores_radar, 'o-', linewidth=2)
            ax9.fill(angles, scores_radar, alpha=0.25)
            ax9.set_xticks(angles[:-1])
            ax9.set_xticklabels(categories)
            ax9.set_ylim(0, 100)
            ax9.set_title(f'综合评分: {overall_score:.1f}/100', fontsize=12, fontweight='bold', pad=20)
        else:
            ax9.text(0.5, 0.5, f'综合评分: {overall_score:.1f}/100', 
                    ha='center', va='center', transform=ax9.transAxes, fontsize=16, fontweight='bold')
        
        plt.savefig(f'{sheet_name}_05_商业化综合分析.png', dpi=300, bbox_inches='tight')
        self.figures.append(f'{sheet_name}_05_商业化综合分析.png')
        plt.close()
        print(f"    {sheet_name}_05_商业化综合分析.png")
    
    def generate_report(self):
        """生成分析报告"""
        report_file = 'visualization_report.txt'
        
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("星世界地图商业化数据可视化分析报告\n")
            f.write("=" * 80 + "\n\n")
            
            f.write(f"分析时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"工作表数量: {len(self.all_sheets)}\n")
            f.write(f"生成的图表数量: {len(self.figures)}\n\n")
            
            f.write("=" * 80 + "\n")
            f.write("生成的图表列表:\n")
            f.write("=" * 80 + "\n")
            
            for i, fig_file in enumerate(self.figures, 1):
                f.write(f"{i}. {fig_file}\n")
            
            f.write("\n" + "=" * 80 + "\n")
            f.write("使用说明:\n")
            f.write("=" * 80 + "\n")
            f.write("1. 所有图表已保存为 PNG 格式\n")
            f.write("2. 图表分辨率为 300 DPI，适合打印和展示\n")
            f.write("3. 建议使用图片查看器打开查看\n")
            f.write("4. 可以将图表插入到报告或演示文稿中\n\n")
        
        print(f"\n分析报告已保存到: {report_file}")
        return self.figures


def main():
    """主函数"""
    file_path = 'data.xlsx'
    
    print("=" * 80)
    print("星世界地图商业化数据可视化分析")
    print("=" * 80)
    print()
    
    viz = StarWorldCommercialVisualization(file_path)
    
    if not viz.load_data():
        print("数据加载失败！程序退出。")
        return
    
    print()
    viz.analyze_and_visualize()
    
    print()
    print("=" * 80)
    print("可视化生成完成！")
    print("=" * 80)
    
    figures = viz.generate_report()
    
    print(f"\n共生成 {len(figures)} 个可视化图表")
    print("\n提示：您可以使用图片查看器打开这些 PNG 文件查看详细的分析图表")


if __name__ == "__main__":
    main()
