#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成真实的数据图表
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')  # 使用非交互式后端
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# 设置样式
sns.set_style("whitegrid")

def main():
    """生成真实图表"""
    
    # 读取数据
    df = pd.read_excel('data.xlsx', engine='openpyxl')
    
    # 创建图表目录
    charts_dir = Path('charts')
    charts_dir.mkdir(exist_ok=True)
    
    # 识别数值列
    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
    
    # 图表1: 数据概览
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    # 数据类型分布
    dtype_counts = df.dtypes.value_counts()
    colors = plt.cm.Set3(range(len(dtype_counts)))
    axes[0].pie(dtype_counts.values, labels=[str(x) for x in dtype_counts.index], 
                autopct='%1.1f%%', colors=colors, startangle=90)
    axes[0].set_title('数据类型分布', fontsize=14, fontweight='bold', pad=20)
    
    # 缺失值分布
    missing = df.isnull().sum()
    missing = missing[missing > 0].sort_values(ascending=False).head(10)
    if len(missing) > 0:
        colors = plt.cm.Reds(range(len(missing)))
        axes[1].barh(range(len(missing)), missing.values, color=colors)
        axes[1].set_yticks(range(len(missing)))
        axes[1].set_yticklabels([str(x)[:25] for x in missing.index])
        axes[1].set_xlabel('缺失值数量')
        axes[1].set_title('缺失值分布 (Top 10)', fontsize=14, fontweight='bold', pad=20)
    else:
        axes[1].text(0.5, 0.5, '无缺失值\n数据质量优秀', ha='center', va='center', 
                    transform=axes[1].transAxes, fontsize=16, 
                    bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.5))
        axes[1].set_title('缺失值分布', fontsize=14, fontweight='bold', pad=20)
    
    plt.tight_layout()
    plt.savefig(charts_dir / '1_data_overview.png', dpi=300, bbox_inches='tight')
    plt.close()
    print("图表1生成成功")
    
    # 图表2: 数值列分布
    if len(numeric_cols) > 0:
        n_cols = min(4, len(numeric_cols))
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        axes = axes.flatten()
        
        colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4']
        
        for i, col in enumerate(numeric_cols[:4]):
            data = df[col].dropna()
            if len(data) > 0:
                # 绘制直方图
                axes[i].hist(data, bins=30, edgecolor='black', alpha=0.7, color=colors[i])
                axes[i].set_title(f'{col} 分布', fontsize=12, fontweight='bold')
                axes[i].set_xlabel(col)
                axes[i].set_ylabel('频数')
                
                # 添加统计线
                mean_val = data.mean()
                median_val = data.median()
                axes[i].axvline(mean_val, color='red', linestyle='--', linewidth=2, 
                              label=f'均值: {mean_val:.2f}')
                axes[i].axvline(median_val, color='green', linestyle='--', linewidth=2, 
                              label=f'中位数: {median_val:.2f}')
                axes[i].legend(loc='upper right')
        
        # 隐藏未使用的子图
        for i in range(n_cols, 4):
            axes[i].axis('off')
        
        plt.tight_layout()
        plt.savefig(charts_dir / '2_numeric_distribution.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("图表2生成成功")
    
    # 图表3: 相关性热力图
    if len(numeric_cols) > 1:
        fig, ax = plt.subplots(figsize=(12, 10))
        
        corr_matrix = df[numeric_cols].corr()
        mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
        
        sns.heatmap(corr_matrix, annot=True, cmap='RdYlBu_r', center=0,
                   square=True, linewidths=1, cbar_kws={"shrink": 0.8}, 
                   fmt='.2f', ax=ax, annot_kws={'size': 8}, mask=mask,
                   vmin=-1, vmax=1)
        ax.set_title('数值特征相关性热力图', fontsize=16, fontweight='bold', pad=20)
        
        plt.tight_layout()
        plt.savefig(charts_dir / '3_correlation_heatmap.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("图表3生成成功")
    
    # 图表4: 统计摘要
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.axis('off')
    
    # 准备文本
    stats_text = ["="*60]
    stats_text.append("数据统计概览".center(60))
    stats_text.append("="*60)
    stats_text.append(f"总数据量: {len(df):,}")
    stats_text.append(f"总特征数: {len(df.columns):,}")
    stats_text.append(f"数值特征: {len(numeric_cols):,}")
    stats_text.append(f"缺失值总数: {df.isnull().sum().sum():,}")
    stats_text.append("")
    
    if len(numeric_cols) > 0:
        stats_text.append("-"*60)
        stats_text.append("数值列统计 (前8个)".center(60))
        stats_text.append("-"*60)
        
        for col in numeric_cols[:8]:
            stats = df[col].describe()
            stats_text.append(f"\n【{col}】")
            stats_text.append(f"  均值: {stats['mean']:.2f}")
            stats_text.append(f"  标准差: {stats['std']:.2f}")
            stats_text.append(f"  最小值: {stats['min']:.2f}")
            stats_text.append(f"  25%分位: {stats['25%']:.2f}")
            stats_text.append(f"  中位数: {stats['50%']:.2f}")
            stats_text.append(f"  75%分位: {stats['75%']:.2f}")
            stats_text.append(f"  最大值: {stats['max']:.2f}")
    
    stats_text.append("")
    stats_text.append("="*60)
    
    # 绘制文本
    ax.text(0.05, 0.95, '\n'.join(stats_text), fontsize=10, family='monospace',
            verticalalignment='top', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.7),
            transform=ax.transAxes)
    ax.set_title('关键指标统计', fontsize=16, fontweight='bold', pad=20)
    
    plt.tight_layout()
    plt.savefig(charts_dir / '4_summary_statistics.png', dpi=300, bbox_inches='tight')
    plt.close()
    print("图表4生成成功")
    
    # 图表5: 前10行数据预览
    fig, ax = plt.subplots(figsize=(16, 8))
    ax.axis('tight')
    ax.axis('off')
    
    # 准备表格数据
    display_cols = df.columns[:8]
    table_data = df[display_cols].head(10)
    
    # 创建表格
    table = ax.table(cellText=table_data.values.tolist(),
                     colLabels=display_cols.tolist(),
                     cellLoc='center',
                     loc='center')
    
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1.2, 1.5)
    
    # 设置表头样式
    for i in range(len(display_cols)):
        table[(0, i)].set_facecolor('#667eea')
        table[(0, i)].set_text_props(weight='bold', color='white')
    
    # 设置单元格样式
    for i in range(1, 11):
        for j in range(len(display_cols)):
            if i % 2 == 0:
                table[(i, j)].set_facecolor('#f8f9fa')
    
    ax.set_title('数据预览 (前10行)', fontsize=16, fontweight='bold', pad=20)
    plt.tight_layout()
    plt.savefig(charts_dir / '5_data_preview.png', dpi=300, bbox_inches='tight')
    plt.close()
    print("图表5生成成功")
    
    print(f"\n所有图表已生成到 {charts_dir} 目录")
    print("图表列表:")
    print("  1. 数据概览")
    print("  2. 数值列分布")
    print("  3. 相关性热力图")
    print("  4. 统计摘要")
    print("  5. 数据预览")

if __name__ == "__main__":
    import numpy as np
    main()
