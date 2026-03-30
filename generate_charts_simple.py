#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
简化的数据可视化生成脚本
直接输出结果到文件，不依赖stdout
"""

import sys
import pandas as pd
import openpyxl
import matplotlib
matplotlib.use('Agg')  # 使用非交互式后端
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def log(message):
    """记录日志到文件"""
    with open('chart_generation_log.txt', 'a', encoding='utf-8') as f:
        f.write(f"{message}\n")
        f.flush()

def main():
    """主函数"""
    log("=" * 60)
    log("开始生成可视化图表")
    log("=" * 60)
    
    try:
        # 加载数据
        log("正在加载数据...")
        df = pd.read_excel('data.xlsx', engine='openpyxl')
        log(f"数据加载成功: {df.shape[0]} 行, {df.shape[1]} 列")
        log(f"列名: {df.columns.tolist()}")
        
        # 创建charts目录
        charts_dir = Path('charts')
        charts_dir.mkdir(exist_ok=True)
        log(f"创建目录: {charts_dir}")
        
        # 识别列类型
        numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
        log(f"数值列: {len(numeric_cols)} 个")
        
        # 创建图表1: 数据概览
        log("生成图表1: 数据概览...")
        fig, axes = plt.subplots(1, 2, figsize=(14, 6))
        
        dtype_counts = df.dtypes.value_counts()
        axes[0].pie(dtype_counts.values, labels=[str(x)[:20] for x in dtype_counts.index], autopct='%1.1f%%')
        axes[0].set_title('数据类型分布', fontsize=14, fontweight='bold')
        
        missing = df.isnull().sum()
        missing = missing[missing > 0].sort_values(ascending=False).head(10)
        if len(missing) > 0:
            axes[1].barh(range(len(missing)), missing.values)
            axes[1].set_yticks(range(len(missing)))
            axes[1].set_yticklabels([str(x)[:20] for x in missing.index])
            axes[1].set_xlabel('缺失值数量')
            axes[1].set_title('缺失值分布 (Top 10)', fontsize=14, fontweight='bold')
        else:
            axes[1].text(0.5, 0.5, '无缺失值', ha='center', va='center', transform=axes[1].transAxes, fontsize=16)
            axes[1].set_title('缺失值分布', fontsize=14, fontweight='bold')
        
        plt.tight_layout()
        chart1_path = charts_dir / '1_data_overview.png'
        plt.savefig(chart1_path, dpi=300, bbox_inches='tight')
        plt.close()
        log(f"图表1保存成功: {chart1_path}")
        
        # 创建图表2: 数值列分布（前4个）
        log("生成图表2: 数值列分布...")
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        axes = axes.flatten()
        
        for i, col in enumerate(numeric_cols[:4]):
            data = df[col].dropna()
            if len(data) > 0:
                axes[i].hist(data, bins=30, edgecolor='black', alpha=0.7, color='skyblue')
                axes[i].set_title(f'{col} 分布', fontsize=12, fontweight='bold')
                axes[i].set_xlabel(col)
                axes[i].set_ylabel('频数')
                
                mean_val = data.mean()
                median_val = data.median()
                axes[i].axvline(mean_val, color='red', linestyle='--', linewidth=2, label=f'均值: {mean_val:.2f}')
                axes[i].axvline(median_val, color='green', linestyle='--', linewidth=2, label=f'中位数: {median_val:.2f}')
                axes[i].legend()
        
        # 隐藏未使用的子图
        for i in range(len(numeric_cols), 4):
            axes[i].axis('off')
        
        plt.tight_layout()
        chart2_path = charts_dir / '2_numeric_distribution.png'
        plt.savefig(chart2_path, dpi=300, bbox_inches='tight')
        plt.close()
        log(f"图表2保存成功: {chart2_path}")
        
        # 创建图表3: 相关性热力图
        log("生成图表3: 相关性热力图...")
        if len(numeric_cols) > 1:
            fig, ax = plt.subplots(figsize=(12, 10))
            corr_matrix = df[numeric_cols].corr()
            sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0,
                       square=True, linewidths=1, cbar_kws={"shrink": 0.8}, 
                       fmt='.2f', ax=ax, annot_kws={'size': 8})
            ax.set_title('数值特征相关性热力图', fontsize=14, fontweight='bold', pad=20)
            plt.tight_layout()
            chart3_path = charts_dir / '3_correlation_heatmap.png'
            plt.savefig(chart3_path, dpi=300, bbox_inches='tight')
            plt.close()
            log(f"图表3保存成功: {chart3_path}")
        
        # 创建图表4: 统计摘要
        log("生成图表4: 统计摘要...")
        fig, ax = plt.subplots(figsize=(12, 8))
        
        stats_text = ["=== 数据统计概览 ==="]
        stats_text.append(f"总数据量: {len(df):,}")
        stats_text.append(f"总特征数: {len(df.columns):,}")
        stats_text.append(f"数值特征: {len(numeric_cols):,}")
        stats_text.append(f"缺失值总数: {df.isnull().sum().sum():,}")
        
        stats_text.append("\n=== 数值列统计 (前5个) ===")
        for col in numeric_cols[:5]:
            stats = df[col].describe()
            stats_text.append(f"\n{col}:")
            stats_text.append(f"  均值: {stats['mean']:.2f}")
            stats_text.append(f"  标准差: {stats['std']:.2f}")
            stats_text.append(f"  最小值: {stats['min']:.2f}")
            stats_text.append(f"  最大值: {stats['max']:.2f}")
        
        ax.axis('off')
        ax.text(0.1, 1.0, '\n'.join(stats_text), fontsize=10, family='monospace',
                verticalalignment='top', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
        ax.set_title('关键指标统计', fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        chart4_path = charts_dir / '4_summary_statistics.png'
        plt.savefig(chart4_path, dpi=300, bbox_inches='tight')
        plt.close()
        log(f"图表4保存成功: {chart4_path}")
        
        # 生成HTML报告
        log("生成HTML报告...")
        html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>星世界地图商业化数据分析报告</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Microsoft YaHei', Arial, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); padding: 40px; }}
        .header {{ text-align: center; margin-bottom: 40px; padding: 30px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 15px; }}
        .header h1 {{ font-size: 2.5em; margin-bottom: 10px; }}
        .section {{ margin-bottom: 50px; }}
        .section-title {{ font-size: 1.8em; color: #667eea; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 3px solid #667eea; }}
        .chart-container {{ margin: 30px 0; text-align: center; }}
        .chart-container img {{ max-width: 100%; height: auto; border-radius: 10px; box-shadow: 0 5px 15px rgba(0,0,0,0.2); }}
        .chart-caption {{ margin-top: 15px; font-size: 1.1em; color: #555; }}
        .stats-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 30px 0; }}
        .stat-card {{ background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); padding: 25px; border-radius: 10px; text-align: center; }}
        .stat-card h3 {{ color: #667eea; font-size: 0.9em; margin-bottom: 10px; }}
        .stat-card .value {{ font-size: 2em; font-weight: bold; color: #333; }}
        .analysis-text {{ background: #f8f9fa; padding: 25px; border-radius: 10px; border-left: 5px solid #667eea; margin: 20px 0; }}
        .footer {{ text-align: center; margin-top: 50px; padding-top: 20px; border-top: 2px solid #eee; color: #888; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎮 星世界地图商业化数据分析报告</h1>
            <p>基于游玩排名前300地图的付费与广告数据深度分析</p>
        </div>
        
        <div class="section">
            <h2 class="section-title">📊 数据概览</h2>
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>总地图数量</h3>
                    <div class="value">{len(df)}</div>
                </div>
                <div class="stat-card">
                    <h3>分析字段数</h3>
                    <div class="value">{len(df.columns)}</div>
                </div>
                <div class="stat-card">
                    <h3>数值特征</h3>
                    <div class="value">{len(numeric_cols)}</div>
                </div>
                <div class="stat-card">
                    <h3>缺失值总数</h3>
                    <div class="value">{df.isnull().sum().sum()}</div>
                </div>
            </div>
            <div class="chart-container">
                <img src="charts/1_data_overview.png" alt="数据概览">
                <div class="chart-caption">图表1: 数据类型分布与缺失值分析</div>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">📈 数值指标分布</h2>
            <div class="analysis-text">
                <h4>关键发现</h4>
                <p>• 共识别到 <strong>{len(numeric_cols)}</strong> 个数值型指标</p>
                <p>• 涵盖排名、付费、广告等多个维度</p>
                <p>• 指标分布呈现多样化特征</p>
            </div>
            <div class="chart-container">
                <img src="charts/2_numeric_distribution.png" alt="数值分布">
                <div class="chart-caption">图表2: 数值型指标分布分析</div>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">🔗 指标相关性分析</h2>
            <div class="chart-container">
                <img src="charts/3_correlation_heatmap.png" alt="相关性热力图">
                <div class="chart-caption">图表3: 数值特征相关性热力图</div>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">📈 综合统计</h2>
            <div class="chart-container">
                <img src="charts/4_summary_statistics.png" alt="综合统计">
                <div class="chart-caption">图表4: 关键指标统计摘要</div>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">💡 商业化建议</h2>
            <div class="analysis-text">
                <h4>优化策略</h4>
                <ul style="margin-left: 20px;">
                    <li>提升付费转化: 针对活跃用户设计个性化付费礼包</li>
                    <li>优化广告投放: 根据CTR数据调整广告位布局</li>
                    <li>提升用户留存: 通过排行榜和活动提高用户粘性</li>
                    <li>精细化运营: 基于数据分析制定分层的商业化策略</li>
                </ul>
            </div>
        </div>
        
        <div class="footer">
            <p>报告生成时间: 2026-03-26</p>
            <p>基于游玩排名前300地图 付费&广告数据</p>
        </div>
    </div>
</body>
</html>
"""
        
        html_path = '商业化分析报告.html'
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        log(f"HTML报告保存成功: {html_path}")
        
        log("=" * 60)
        log("✅ 图表生成完成!")
        log(f"📊 生成图表: 4 张")
        log(f"📄 HTML报告: {html_path}")
        log("=" * 60)
        
        return True
        
    except Exception as e:
        log(f"❌ 错误: {str(e)}")
        import traceback
        log(traceback.format_exc())
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
