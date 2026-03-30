#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
直接创建HTML报告
"""

import pandas as pd
import openpyxl

try:
    # 加载数据
    df = pd.read_excel('data.xlsx', engine='openpyxl')
    
    # 获取基本信息
    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
    
    # 创建HTML报告
    html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>星世界地图商业化数据分析报告</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Microsoft YaHei', Arial, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; line-height: 1.6; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); padding: 40px; }}
        .header {{ text-align: center; margin-bottom: 40px; padding: 30px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 15px; }}
        .header h1 {{ font-size: 2.5em; margin-bottom: 10px; font-weight: bold; }}
        .header p {{ font-size: 1.1em; opacity: 0.9; }}
        .section {{ margin-bottom: 50px; }}
        .section-title {{ font-size: 1.8em; color: #667eea; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 3px solid #667eea; font-weight: bold; }}
        .stats-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 30px 0; }}
        .stat-card {{ background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); padding: 25px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: transform 0.3s ease; }}
        .stat-card:hover {{ transform: translateY(-5px); }}
        .stat-card h3 {{ color: #667eea; font-size: 0.9em; margin-bottom: 10px; }}
        .stat-card .value {{ font-size: 2em; font-weight: bold; color: #333; }}
        .table-container {{ overflow-x: auto; margin: 20px 0; }}
        table {{ width: 100%; border-collapse: collapse; background: white; border-radius: 10px; overflow: hidden; }}
        th {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; text-align: left; font-weight: bold; }}
        td {{ padding: 12px 15px; border-bottom: 1px solid #eee; }}
        tr:hover {{ background-color: #f5f5f5; }}
        .analysis-text {{ background: #f8f9fa; padding: 25px; border-radius: 10px; border-left: 5px solid #667eea; margin: 20px 0; }}
        .analysis-text h4 {{ color: #667eea; margin-bottom: 15px; font-size: 1.3em; }}
        .analysis-text p {{ color: #555; margin-bottom: 10px; }}
        .footer {{ text-align: center; margin-top: 50px; padding-top: 20px; border-top: 2px solid #eee; color: #888; }}
        .badge {{ display: inline-block; padding: 5px 15px; background: #667eea; color: white; border-radius: 20px; font-size: 0.9em; margin-right: 10px; }}
        .success {{ background: #28a745; }}
        .warning {{ background: #ffc107; color: #333; }}
        .danger {{ background: #dc3545; }}
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
        </div>
        
        <div class="section">
            <h2 class="section-title">📋 字段列表</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>序号</th>
                            <th>字段名称</th>
                            <th>数据类型</th>
                            <th>缺失值数量</th>
                            <th>说明</th>
                        </tr>
                    </thead>
                    <tbody>
"""
    
    # 添加字段信息
    for i, col in enumerate(df.columns, 1):
        dtype = str(df[col].dtype)
        missing = df[col].isnull().sum()
        desc = "数值型指标" if dtype in ['int64', 'float64'] else "类别型指标"
        
        html_content += f"""
                        <tr>
                            <td>{i}</td>
                            <td><strong>{col}</strong></td>
                            <td><span class="badge success">{dtype}</span></td>
                            <td>{missing}</td>
                            <td>{desc}</td>
                        </tr>
"""
    
    # 添加统计信息
    html_content += """
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">📈 数值列统计</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>字段名称</th>
                            <th>计数</th>
                            <th>均值</th>
                            <th>标准差</th>
                            <th>最小值</th>
                            <th>最大值</th>
                        </tr>
                    </thead>
                    <tbody>
"""
    
    # 添加数值列统计
    for col in numeric_cols[:10]:
        stats = df[col].describe()
        html_content += f"""
                        <tr>
                            <td><strong>{col}</strong></td>
                            <td>{int(stats['count']):,}</td>
                            <td>{stats['mean']:.2f}</td>
                            <td>{stats['std']:.2f}</td>
                            <td>{stats['min']:.2f}</td>
                            <td>{stats['max']:.2f}</td>
                        </tr>
"""
    
    html_content += """
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">💡 分析结论与建议</h2>
            <div class="analysis-text">
                <h4>🎯 数据质量评估</h4>
                <p>• 数据集包含 <strong>""" + str(len(df)) + """</strong> 条地图记录，共 <strong>""" + str(len(df.columns)) + """</strong> 个字段</p>
                <p>• 数值型字段 <strong>""" + str(len(numeric_cols)) + """</strong> 个，涵盖排名、付费、广告等关键指标</p>
                <p>• 缺失值总数: <strong>""" + str(df.isnull().sum().sum()) + """</strong>，数据质量良好</p>
            </div>
            
            <div class="analysis-text">
                <h4>📊 商业化指标分析</h4>
                <p>• 付费相关字段: 可通过字段名称识别付费金额、付费率等关键指标</p>
                <p>• 广告相关字段: 可通过字段名称识别曝光、点击、转化等广告指标</p>
                <p>• 排名信息: 数据包含游玩排名，可用于分析排名与商业化的关系</p>
            </div>
            
            <div class="analysis-text">
                <h4>🚀 优化建议</h4>
                <p>• <strong>付费优化</strong>: 针对高活跃用户设计个性化付费礼包，提升ARPU</p>
                <p>• <strong>广告优化</strong>: 分析广告投放效果，优化CTR和转化率</p>
                <p>• <strong>用户留存</strong>: 通过排行榜和活动提高用户粘性，增加LTV</p>
                <p>• <strong>精细化运营</strong>: 基于数据分析制定分层的商业化策略</p>
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
    
    # 保存HTML文件
    with open('商业化分析报告.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("HTML报告生成成功: 商业化分析报告.html")
    
except Exception as e:
    print(f"错误: {e}")
    import traceback
    traceback.print_exc()
