import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
from plotly.subplots import make_subplots
# import matplotlib.pyplot as plt  # 暂时注释掉，避免导入错误
from collections import defaultdict

def create_pie_chart(data):
    """创建三个独立的环形图 - 修复进度条方向相反问题"""
    # 创建包含三个子图的图表
    fig = make_subplots(rows=1, cols=3, subplot_titles=('时间进度', '年使用率', '月累使用率'))
    
    # 定义每个指标的配置
    indicators = [
        {
            'name': '时间进度',
            'value': data['time_progress'],
            'color': '#1f77b4',  # 蓝色
            'row': 1,
            'col': 1
        },
        {
            'name': '年使用率',
            'value': data['year_usage'],
            'color': '#ff7f0e',  # 橙色
            'row': 1,
            'col': 2
        },
        {
            'name': '月累使用率',
            'value': data['month_usage'],
            'color': '#2ca02c',  # 绿色
            'row': 1,
            'col': 3
        }
    ]
    
    # 为每个指标创建环形图
    for ind in indicators:
        # 计算剩余百分比
        remaining = 100 - ind['value']
        
        # 添加环形图数据
        fig.add_trace(
            go.Pie(
                labels=[ind['name'], '剩余'],
                values=[ind['value'], remaining],
                hole=0.6,
                marker_colors=[ind['color'], '#f0f0f0'],  # 指标颜色和灰色背景
                textinfo='none',
                hoverinfo='label+percent',
                direction='clockwise',  # 确保顺时针方向
                rotation=90  # 从顶部开始
            ),
            row=ind['row'],
            col=ind['col']
        )
        
        # 添加百分比标签
        fig.add_annotation(
            x=0.5, y=0.5, xref=f'x{ind["col"]}', yref=f'y{ind["col"]}',
            text=f"{ind['value']:.1f}%",
            showarrow=False,
            font=dict(size=20, color=ind['color'])
        )
    
    # 更新布局
    fig.update_layout(
        title="项目进度概览",
        height=400,
        showlegend=False,
        margin=dict(l=20, r=20, t=50, b=20)
    )
    
    # 更新每个子图的坐标轴，确保环形图居中
    for i in range(1, 4):
        fig.update_xaxes(showticklabels=False, showgrid=False, zeroline=False, row=1, col=i)
        fig.update_yaxes(showticklabels=False, showgrid=False, zeroline=False, row=1, col=i)
    
    return fig

def create_bar_chart(data):
    """创建柱状图 - 根据输入类型自动选择展示方式"""
    # 检查数据类型
    if isinstance(data, dict):
        # 处理字典类型数据 (二级费项累计对比)
        if not data.get('fee_items'):
            return None
        
        names = [item['name'] for item in data['fee_items']]
        targets = [item['cum_target'] for item in data['fee_items']]
        actuals = [item['cum_actual'] for item in data['fee_items']]
        
        fig = go.Figure()
        
        # 添加柱状图
        fig.add_trace(go.Bar(
            name='目标累计',
            x=names,
            y=targets,
            marker_color='#1f77b4',
            visible=True
        ))
        
        fig.add_trace(go.Bar(
            name='已发生累计',
            x=names,
            y=actuals,
            marker_color='#ff7f0e',
            visible=True
        ))
        
        # 自定义Y轴刻度标签
        def format_y_axis_ticks(tick_values):
            """将数值转换为中文格式"""
            formatted_ticks = []
            for value in tick_values:
                if value >= 1000000:  # 大于等于100万
                    formatted_value = f"{value/1000000:.0f}亿"
                elif value >= 10000:  # 大于等于1万
                    formatted_value = f"{value/10000:.0f}万"
                else:
                    formatted_value = f"{value:.0f}"
                formatted_ticks.append(formatted_value)
            return formatted_ticks
        
        fig.update_layout(
            title={
                'text': "二级费项累计对比",
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18, 'family': 'Microsoft YaHei'}
            },
            height=300,
            plot_bgcolor='white',
            paper_bgcolor='white',
            xaxis=dict(
                title=dict(text="费项", font=dict(family='Microsoft YaHei')),
                showgrid=True,
                gridwidth=1,
                gridcolor='#f0f0f0',
                tickangle=0,
                tickfont=dict(family='Microsoft YaHei')
            ),
            yaxis=dict(
                title=dict(text="累计金额 (万元)", font=dict(family='Microsoft YaHei')),
                showgrid=True,
                gridwidth=1,
                gridcolor='#f0f0f0',
                zeroline=True,
                zerolinecolor='#f0f0f0',
                tickmode='array',
                tickvals=[0, 500000, 1000000, 1500000, 2000000],
                ticktext=format_y_axis_ticks([0, 500000, 1000000, 1500000, 2000000]),
                tickfont=dict(family='Microsoft YaHei')
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5,
                font=dict(family='Microsoft YaHei')
            ),
            barmode='group',
            font=dict(family='Microsoft YaHei')
        )
    else:
        # 处理DataFrame类型数据 (每月费项成本对比)
        if data.empty:
            return None
        
        fig = go.Figure()
        
        # 添加柱状图
        fig.add_trace(go.Bar(
            name='目标成本',
            x=data['月份'],
            y=data['目标成本'],
            marker_color='#1f77b4',
            visible=True
        ))
        
        fig.add_trace(go.Bar(
            name='已发生成本',
            x=data['月份'],
            y=data['已发生成本'],
            marker_color='#ff7f0e',
            visible=True
        ))
        
        # 自定义Y轴刻度标签
        def format_y_axis_ticks(tick_values):
            """将数值转换为中文格式"""
            formatted_ticks = []
            for value in tick_values:
                if value >= 10000:  # 大于等于1万
                    formatted_value = f"{value/10000:.0f}万"
                else:
                    formatted_value = f"{value:.0f}"
                formatted_ticks.append(formatted_value)
            return formatted_ticks
        
        # 计算Y轴范围，留出一些空间
        max_value = max(data['目标成本'].max(), data['已发生成本'].max()) * 1.1
        tick_step = max_value / 5 if max_value > 0 else 1
        tick_vals = [i * tick_step for i in range(6)]
        
        fig.update_layout(
            title={
                'text': "每月费项成本对比",
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18, 'family': 'Microsoft YaHei'}
            },
            height=400,
            plot_bgcolor='white',
            paper_bgcolor='white',
            xaxis=dict(
                title=dict(text="月份", font=dict(family='Microsoft YaHei')),
                showgrid=True,
                gridwidth=1,
                gridcolor='#f0f0f0',
                tickangle=0,
                tickfont=dict(family='Microsoft YaHei')
            ),
            yaxis=dict(
                title=dict(text="金额 (万元)", font=dict(family='Microsoft YaHei')),
                showgrid=True,
                gridwidth=1,
                gridcolor='#f0f0f0',
                zeroline=True,
                zerolinecolor='#f0f0f0',
                # 自定义刻度标签
                tickmode='array',
                tickvals=tick_vals,
                ticktext=format_y_axis_ticks(tick_vals),
                tickfont=dict(family='Microsoft YaHei')
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5,
                font=dict(family='Microsoft YaHei')
            ),
            barmode='group',
            font=dict(family='Microsoft YaHei')
        )
    
    return fig

def create_line_chart(data):
    """创建折线图 - 总成本累计趋势"""
    months = list(range(1, 13))
    
    fig = go.Figure()
    
    # 添加折线图
    fig.add_trace(go.Scatter(
        name='总目标累计',
        x=months,
        y=data['cum_target'],
        mode='lines+markers',
        line=dict(color='#1f77b4', width=3),
        marker=dict(size=8, color='#1f77b4'),
        visible=True
    ))
    
    fig.add_trace(go.Scatter(
        name='总已发生成本',
        x=months,
        y=data['cum_actual'],
        mode='lines+markers',
        line=dict(color='#ff7f0e', width=3),
        marker=dict(size=8, color='#ff7f0e'),
        visible=True
    ))
    
    # 自定义Y轴刻度标签
    def format_y_axis_ticks(tick_values):
        """将数值转换为中文格式"""
        formatted_ticks = []
        for value in tick_values:
            if value >= 10000000:  # 大于等于1000万
                formatted_value = f"{value/10000000:.0f}亿"
            elif value >= 10000:  # 大于等于1万
                formatted_value = f"{value/10000:.0f}万"
            else:
                formatted_value = f"{value:.0f}"
            formatted_ticks.append(formatted_value)
        return formatted_ticks
    
    fig.update_layout(
        title={
            'text': "总成本累计趋势",
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'family': 'Microsoft YaHei'}
        },
        height=300,
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(
            title=dict(text="月份", font=dict(family='Microsoft YaHei')),
            showgrid=True,
            gridwidth=1,
            gridcolor='#f0f0f0',
            tickmode='linear',
            tick0=1,
            dtick=1,
            tickfont=dict(family='Microsoft YaHei')
        ),
        yaxis=dict(
            title=dict(text="累计金额 (万元)", font=dict(family='Microsoft YaHei')),
            showgrid=True,
            gridwidth=1,
            gridcolor='#f0f0f0',
            zeroline=True,
            zerolinecolor='#f0f0f0',
            # 自定义刻度标签 - 增加更多刻度点
            tickmode='array',
            tickvals=[0, 10000000, 20000000, 30000000, 40000000],  # 0, 1000万, 2000万, 3000万, 4000万
            ticktext=format_y_axis_ticks([0, 10000000, 20000000, 30000000, 40000000]),  # 0, 10亿, 20亿, 30亿, 40亿
            tickfont=dict(family='Microsoft YaHei')
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
            font=dict(family='Microsoft YaHei')
        ),
        font=dict(family='Microsoft YaHei')
    )
    
    return fig

def create_bar_line_chart(data):
    """创建上下排布的两张图表"""
    # 创建2行1列的子图布局
    fig = make_subplots(rows=2, cols=1, 
                        subplot_titles=('二级费项累计对比', '总成本累计趋势'),
                        vertical_spacing=0.2)
    
    # 获取柱状图和折线图的数据
    bar_chart = create_bar_chart(data)
    line_chart = create_line_chart(data)
    
    # 如果任一图表为None，返回None
    if bar_chart is None or line_chart is None:
        return None
    
    # 将柱状图的轨迹添加到第一个子图
    for trace in bar_chart.data:
        fig.add_trace(trace, row=1, col=1)
    
    # 将折线图的轨迹添加到第二个子图
    for trace in line_chart.data:
        fig.add_trace(trace, row=2, col=1)
    
    # 更新布局
    fig.update_layout(
        height=700,  # 增加高度以适应两个子图
        showlegend=False,  # 不显示图例（每个子图已有自己的图例）
        margin=dict(l=50, r=50, t=80, b=50)
    )
    
    # 更新x轴和y轴标签
    fig.update_xaxes(title_text="费项", row=1, col=1)
    fig.update_yaxes(title_text="累计金额 (万元)", row=1, col=1)
    fig.update_xaxes(title_text="月份", row=2, col=1)
    fig.update_yaxes(title_text="累计金额 (万元)", row=2, col=1)
    
    return fig

def create_project_comparison_chart(all_data, month):
    """创建项目对比组合图 - 包含使用率柱状图和金额折线图"""
    projects = list(all_data.keys())
    year_usages = [all_data[p]['year_usage'] for p in projects]
    month_usages = [all_data[p]['month_usage'] for p in projects]
    year_targets = [all_data[p]['year_cum_target_wy'] for p in projects]
    # 年累计已发生成本（12月累计）
    year_actuals = [all_data[p]['cum_actual'][-1]/10000 for p in projects]  # 取12月累计值并转换为万元
    
    # 创建图表
    fig = go.Figure()
    
    # 添加年使用率柱状图
    fig.add_trace(go.Bar(
        x=projects, y=year_usages, 
        name='年使用率', 
        marker_color='#1f77b4',
        opacity=0.7,
        width=0.3,
        hovertemplate='<b>%{x}</b><br>年使用率: %{y:.1f}%<extra></extra>',
        yaxis='y2'  # 使用率轴在右边
    ))
    
    # 添加月累使用率柱状图
    fig.add_trace(go.Bar(
        x=projects, y=month_usages, 
        name='月累使用率', 
        marker_color='#ff7f0e',
        opacity=0.7,
        width=0.3,
        hovertemplate='<b>%{x}</b><br>月累使用率: %{y:.1f}%<extra></extra>',
        yaxis='y2'  # 使用率轴在右边
    ))
    
    # 添加年总目标折线
    fig.add_trace(go.Scatter(
        x=projects, y=year_targets, mode='lines+markers', 
        name='年总目标', 
        line=dict(color='#32CD32', width=3),
        marker=dict(size=10, color='#32CD32'),
        hovertemplate='<b>%{x}</b><br>年总目标: %{y:.1f}万元<extra></extra>',
        yaxis='y'  # 金额轴在左边
    ))
    
    # 添加年累计已发生折线
    fig.add_trace(go.Scatter(
        x=projects, y=year_actuals, mode='lines+markers', 
        name='年累计已发生', 
        line=dict(color='#FF0000', width=3),
        marker=dict(size=10, color='#FF0000'),
        hovertemplate='<b>%{x}</b><br>年累计已发生: %{y:.1f}万元<extra></extra>',
        yaxis='y'  # 金额轴在左边
    ))
    
    fig.update_layout(
        title={
            'text': "项目使用率与金额对比",
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'family': 'Microsoft YaHei'}
        },
        xaxis_title="项目",
        height=400,
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(
            showgrid=True,
            gridwidth=1,
            gridcolor='#f0f0f0',
            tickangle=0,
            tickfont=dict(family='Microsoft YaHei')
        ),
        yaxis=dict(
            title=dict(text="金额 (万元)", font=dict(family='Microsoft YaHei')),
            showgrid=True,
            gridwidth=1,
            gridcolor='#f0f0f0',
            zeroline=True,
            zerolinecolor='#f0f0f0',
            tickfont=dict(family='Microsoft YaHei'),
            side='left'
        ),
        yaxis2=dict(
            title=dict(text="使用率 (%)", font=dict(family='Microsoft YaHei')),
            overlaying='y',
            side='right',
            tickfont=dict(family='Microsoft YaHei'),
            range=[0, 100],
            tickmode='linear',
            dtick=20
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
            font=dict(family='Microsoft YaHei')
        ),
        font=dict(family='Microsoft YaHei'),
        barmode='group'  # 柱状图分组显示
    )
    return fig

def create_kpi_display(all_data):
    """创建KPI指标显示"""
    kpi_data = []
    for project_name, data in all_data.items():
        kpi_data.append({
            "项目": project_name,
            "年累计目标(万元)": data['year_cum_target_wy'],
            "年累计已发生(万元)": data['year_cum_actual_wy'],
            "年使用率(%)": data['year_usage'],
            "月累使用率(%)": data['month_usage'],
            "时间进度(%)": data['time_progress']
        })
    return kpi_data 

def create_multi_ring_progress_chart(data, month):
    """创建多环径向进度图，类似用户提供的例图样式"""
    
    # 获取三个进度指标
    year_usage = data['year_usage']
    month_usage = data['month_usage']
    time_progress = data['time_progress']
    
    # 创建三个同心圆环
    fig = go.Figure()
    
    # 定义颜色
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1']  # 红色、青色、蓝色
    
    # 年使用率环（最内层）
    fig.add_trace(go.Scatterpolar(
        r=[year_usage, year_usage, 0, 0],
        theta=[0, 360, 360, 0],
        fill='toself',
        fillcolor=colors[0],
        line=dict(color=colors[0], width=2),
        name='年使用率',
        showlegend=True,
        hovertemplate='年使用率: %{r:.1f}%<extra></extra>'
    ))
    
    # 月累使用率环（中间层）
    fig.add_trace(go.Scatterpolar(
        r=[month_usage, month_usage, 0, 0],
        theta=[0, 360, 360, 0],
        fill='toself',
        fillcolor=colors[1],
        line=dict(color=colors[1], width=2),
        name='月累使用率',
        showlegend=True,
        hovertemplate='月累使用率: %{r:.1f}%<extra></extra>'
    ))
    
    # 时间进度环（最外层）
    fig.add_trace(go.Scatterpolar(
        r=[time_progress, time_progress, 0, 0],
        theta=[0, 360, 360, 0],
        fill='toself',
        fillcolor=colors[2],
        line=dict(color=colors[2], width=2),
        name='时间进度',
        showlegend=True,
        hovertemplate='时间进度: %{r:.1f}%<extra></extra>'
    ))
    
    # 添加刻度线
    for i in range(0, 101, 20):
        fig.add_trace(go.Scatterpolar(
            r=[i, i],
            theta=[0, 360],
            mode='lines',
            line=dict(color='lightgray', width=1, dash='dot'),
            showlegend=False,
            hoverinfo='skip'
        ))
    
    # 添加径向线
    for i in range(0, 360, 30):
        fig.add_trace(go.Scatterpolar(
            r=[0, 100],
            theta=[i, i],
            mode='lines',
            line=dict(color='lightgray', width=1),
            showlegend=False,
            hoverinfo='skip'
        ))
    
    # 更新布局
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 100],
                ticktext=['0%', '20%', '40%', '60%', '80%', '100%'],
                tickvals=[0, 20, 40, 60, 80, 100],
                tickfont=dict(size=10),
                gridcolor='lightgray',
                linecolor='lightgray'
            ),
            angularaxis=dict(
                visible=True,
                ticktext=['0°', '90°', '180°', '270°'],
                tickvals=[0, 90, 180, 270],
                tickfont=dict(size=10),
                gridcolor='lightgray',
                linecolor='lightgray'
            ),
            bgcolor='white'
        ),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        margin=dict(l=50, r=50, t=50, b=50),
        height=400,
        title=dict(
            text="进度指标多环图",
            x=0.5,
            font=dict(size=16)
        )
    )
    
    return fig 

def create_semi_circular_chart(data, title="项目进度概览"):
    """创建半圆形多段式图表，类似用户提供的例图样式"""
    
    # 获取数据
    year_usage = data['year_usage']
    month_usage = data['month_usage']
    time_progress = data['time_progress']
    
    # 创建半圆形图表
    fig = go.Figure()
    
    # 定义颜色
    colors = ['#FFD700', '#C0C0C0', '#FFA500', '#4169E1', '#32CD32', '#1E90FF']  # 黄、灰、橙、蓝、绿、深蓝
    
    # 计算角度范围（半圆180度）
    total_angle = 180
    year_angle = (year_usage / 100) * total_angle
    month_angle = (month_usage / 100) * total_angle
    time_angle = (time_progress / 100) * total_angle
    
    # 年使用率扇形（最大）
    fig.add_trace(go.Scatterpolar(
        r=[1, 1, 0.8, 0.8],
        theta=[0, year_angle, year_angle, 0],
        fill='toself',
        fillcolor=colors[0],
        line=dict(color=colors[0], width=2),
        name=f'年使用率 ({year_usage:.1f}%)',
        showlegend=True,
        hovertemplate=f'年使用率: {year_usage:.1f}%<extra></extra>'
    ))
    
    # 月累使用率扇形
    fig.add_trace(go.Scatterpolar(
        r=[0.8, 0.8, 0.6, 0.6],
        theta=[0, month_angle, month_angle, 0],
        fill='toself',
        fillcolor=colors[1],
        line=dict(color=colors[1], width=2),
        name=f'月累使用率 ({month_usage:.1f}%)',
        showlegend=True,
        hovertemplate=f'月累使用率: {month_usage:.1f}%<extra></extra>'
    ))
    
    # 时间进度扇形
    fig.add_trace(go.Scatterpolar(
        r=[0.6, 0.6, 0.4, 0.4],
        theta=[0, time_angle, time_angle, 0],
        fill='toself',
        fillcolor=colors[2],
        line=dict(color=colors[2], width=2),
        name=f'时间进度 ({time_progress:.1f}%)',
        showlegend=True,
        hovertemplate=f'时间进度: {time_progress:.1f}%<extra></extra>'
    ))
    
    # 添加刻度线
    for i in range(0, 181, 30):
        fig.add_trace(go.Scatterpolar(
            r=[0, 1],
            theta=[i, i],
            mode='lines',
            line=dict(color='lightgray', width=1),
            showlegend=False,
            hoverinfo='skip'
        ))
    
    # 添加环形刻度
    for i in range(1, 6):
        radius = i * 0.2
        fig.add_trace(go.Scatterpolar(
            r=[radius, radius],
            theta=[0, 180],
            mode='lines',
            line=dict(color='lightgray', width=1, dash='dot'),
            showlegend=False,
            hoverinfo='skip'
        ))
    
    # 更新布局
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 1],
                ticktext=['0%', '20%', '40%', '60%', '80%', '100%'],
                tickvals=[0, 0.2, 0.4, 0.6, 0.8, 1],
                tickfont=dict(size=10),
                gridcolor='lightgray',
                linecolor='lightgray'
            ),
            angularaxis=dict(
                visible=True,
                ticktext=['0°', '30°', '60°', '90°', '120°', '150°', '180°'],
                tickvals=[0, 30, 60, 90, 120, 150, 180],
                tickfont=dict(size=10),
                gridcolor='lightgray',
                linecolor='lightgray'
            ),
            bgcolor='white'
        ),
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.1
        ),
        margin=dict(l=50, r=150, t=50, b=50),
        height=400,
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=16)
        )
    )
    
    return fig 

def create_donut_chart(data, title="项目进度概览"):
    """创建蓝色渐变甜甜圈图，采用圆融的圆弧设计和动态效果"""
    
    # 获取三个进度指标
    year_usage = data['year_usage']
    month_usage = data['month_usage']
    time_progress = data['time_progress']
    
    # 创建甜甜圈图
    fig = go.Figure()
    
    # 定义蓝色渐变色彩 - 从浅蓝到深蓝
    colors = ['#E3F2FD', '#90CAF9', '#42A5F5', '#1E88E5', '#1565C0', '#0D47A1']
    
    # 创建动态的蓝色渐变效果
    def create_gradient_colors(base_color, steps=10):
        """创建渐变色彩"""
        import colorsys
        # 将十六进制转换为RGB
        rgb = tuple(int(base_color[i:i+2], 16)/255 for i in (1, 3, 5))
        # 转换为HSL
        h, l, s = colorsys.rgb_to_hls(*rgb)
        # 创建渐变
        gradient = []
        for i in range(steps):
            # 调整亮度和饱和度
            new_l = max(0.1, min(0.9, l + (i - steps/2) * 0.1))
            new_s = max(0.3, min(1.0, s + (i - steps/2) * 0.1))
            # 转换回RGB
            new_rgb = colorsys.hls_to_rgb(h, new_l, new_s)
            # 转换为十六进制
            hex_color = '#{:02x}{:02x}{:02x}'.format(
                int(new_rgb[0]*255), 
                int(new_rgb[1]*255), 
                int(new_rgb[2]*255)
            )
            gradient.append(hex_color)
        return gradient
    
    # 年使用率环（最内层）- 使用蓝色渐变
    year_gradient = create_gradient_colors('#42A5F5')
    fig.add_trace(go.Pie(
        labels=['年使用率', ''],
        values=[year_usage, 100-year_usage],
        hole=0.3,
        marker_colors=[year_gradient[5], '#f8f9fa'],
        textinfo='none',
        hoverinfo='label+percent',
        name='年使用率',
        showlegend=False,
        # 添加圆融效果
        marker=dict(
            line=dict(width=0),
            pattern=dict(
                shape="",
                size=0,
                solidity=0
            )
        )
    ))
    
    # 月累使用率环（中间层）- 使用中蓝色渐变
    month_gradient = create_gradient_colors('#1E88E5')
    fig.add_trace(go.Pie(
        labels=['月累使用率', ''],
        values=[month_usage, 100-month_usage],
        hole=0.5,
        marker_colors=[month_gradient[5], '#f8f9fa'],
        textinfo='none',
        hoverinfo='label+percent',
        name='月累使用率',
        showlegend=False,
        marker=dict(
            line=dict(width=0),
            pattern=dict(
                shape="",
                size=0,
                solidity=0
            )
        )
    ))
    
    # 时间进度环（最外层）- 使用深蓝色渐变
    time_gradient = create_gradient_colors('#1565C0')
    fig.add_trace(go.Pie(
        labels=['时间进度', ''],
        values=[time_progress, 100-time_progress],
        hole=0.7,
        marker_colors=[time_gradient[5], '#f8f9fa'],
        textinfo='none',
        hoverinfo='label+percent',
        name='时间进度',
        showlegend=False,
        marker=dict(
            line=dict(width=0),
            pattern=dict(
                shape="",
                size=0,
                solidity=0
            )
        )
    ))
    
    # 添加动态标签和连接线 - 采用蓝色主题
    # 年使用率标签
    fig.add_annotation(
        x=1.4, y=0.5, xref='paper', yref='paper',
        text=f"{year_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#42A5F5',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#42A5F5',
        borderwidth=2,
        # 添加圆角效果
        align='center',
        valign='middle'
    )
    
    # 月累使用率标签
    fig.add_annotation(
        x=0.5, y=1.3, xref='paper', yref='paper',
        text=f"{month_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#1E88E5',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#1E88E5',
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 时间进度标签
    fig.add_annotation(
        x=0.15, y=0.6, xref='paper', yref='paper',
        text=f"{time_progress:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#1565C0',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#1565C0',
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 更新布局 - 添加动态效果和蓝色主题
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=18, color='#1565C0', family='Arial Black'),
            y=0.95
        ),
        height=400,
        showlegend=False,
        margin=dict(l=60, r=60, t=80, b=60),
        plot_bgcolor='rgba(248,249,250,0.8)',
        paper_bgcolor='white',
        # 添加动态效果
        updatemenus=[{
            'buttons': [],
            'direction': 'left',
            'pad': {'r': 10, 't': 87},
            'showactive': False,
            'type': 'buttons',
            'x': 0.1,
            'xanchor': 'right',
            'y': 0,
            'yanchor': 'top'
        }],
        # 添加动画效果
        sliders=[{
            'active': 0,
            'currentvalue': {'prefix': '动画: '},
            'len': 0.9,
            'pad': {'t': 50},
            'steps': [],
            'x': 0.1,
            'xanchor': 'left',
            'y': 0,
            'yanchor': 'top'
        }]
    )
    
    # 添加悬停效果和动态交互
    fig.update_traces(
        hovertemplate="<b>%{label}</b><br>" +
                     "进度: %{percent:.1%}<br>" +
                     "数值: %{value:.1f}%<extra></extra>",
        hoverlabel=dict(
            bgcolor='rgba(21,101,192,0.9)',
            bordercolor='#1565C0',
            font=dict(color='white', size=12)
        )
    )
    
    return fig 

def create_rounded_donut_chart(data, title="项目进度概览"):
    """创建具有圆融圆弧效果的甜甜圈图，采用蓝色渐变和动态效果"""
    
    # 获取三个进度指标
    year_usage = data['year_usage']
    month_usage = data['month_usage']
    time_progress = data['time_progress']
    
    # 创建甜甜圈图
    fig = go.Figure()
    
    # 定义蓝色渐变色彩
    blue_gradients = {
        'light': ['#E3F2FD', '#BBDEFB', '#90CAF9'],
        'medium': ['#90CAF9', '#64B5F6', '#42A5F5'],
        'dark': ['#42A5F5', '#2196F3', '#1E88E5']
    }
    
    # 年使用率环（最内层）- 浅蓝色渐变
    fig.add_trace(go.Pie(
        labels=['年使用率', ''],
        values=[year_usage, 100-year_usage],
        hole=0.3,
        marker_colors=[blue_gradients['light'][1], '#f8f9fa'],
        textinfo='none',
        hoverinfo='label+percent',
        name='年使用率',
        showlegend=False,
        # 添加圆融效果 - 使用圆形标记
        marker=dict(
            line=dict(width=0),
            pattern=dict(
                shape="",
                size=0,
                solidity=0
            )
        ),
        # 添加渐变效果
        hovertemplate="<b>年使用率</b><br>" +
                     "进度: %{percent:.1%}<br>" +
                     "数值: %{value:.1f}%<extra></extra>"
    ))
    
    # 月累使用率环（中间层）- 中蓝色渐变
    fig.add_trace(go.Pie(
        labels=['月累使用率', ''],
        values=[month_usage, 100-month_usage],
        hole=0.5,
        marker_colors=[blue_gradients['medium'][1], '#f8f9fa'],
        textinfo='none',
        hoverinfo='label+percent',
        name='月累使用率',
        showlegend=False,
        marker=dict(
            line=dict(width=0),
            pattern=dict(
                shape="",
                size=0,
                solidity=0
            )
        ),
        hovertemplate="<b>月累使用率</b><br>" +
                     "进度: %{percent:.1%}<br>" +
                     "数值: %{value:.1f}%<extra></extra>"
    ))
    
    # 时间进度环（最外层）- 深蓝色渐变
    fig.add_trace(go.Pie(
        labels=['时间进度', ''],
        values=[time_progress, 100-time_progress],
        hole=0.7,
        marker_colors=[blue_gradients['dark'][1], '#f8f9fa'],
        textinfo='none',
        hoverinfo='label+percent',
        name='时间进度',
        showlegend=False,
        marker=dict(
            line=dict(width=0),
            pattern=dict(
                shape="",
                size=0,
                solidity=0
            )
        ),
        hovertemplate="<b>时间进度</b><br>" +
                     "进度: %{percent:.1%}<br>" +
                     "数值: %{value:.1f}%<extra></extra>"
    ))
    
    # 添加动态标签和连接线 - 采用蓝色主题
    # 年使用率标签
    fig.add_annotation(
        x=1.35, y=0.5, xref='paper', yref='paper',
        text=f"{year_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#90CAF9',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#90CAF9',
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 月累使用率标签
    fig.add_annotation(
        x=0.5, y=1.25, xref='paper', yref='paper',
        text=f"{month_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#42A5F5',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#42A5F5',
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 时间进度标签
    fig.add_annotation(
        x=0.2, y=0.6, xref='paper', yref='paper',
        text=f"{time_progress:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#1E88E5',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#1E88E5',
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 更新布局 - 确保图表完整显示
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=18, color='#1565C0', family='Arial Black'),
            y=0.95
        ),
        height=450,  # 增加高度确保完整显示
        width=600,   # 设置固定宽度
        showlegend=False,
        margin=dict(l=80, r=80, t=100, b=80),  # 增加边距
        plot_bgcolor='rgba(248,249,250,0.8)',
        paper_bgcolor='white',
        # 添加动画效果
        updatemenus=[{
            'buttons': [],
            'direction': 'left',
            'pad': {'r': 10, 't': 87},
            'showactive': False,
            'type': 'buttons',
            'x': 0.1,
            'xanchor': 'right',
            'y': 0,
            'yanchor': 'top'
        }],
        # 添加滑块效果
        sliders=[{
            'active': 0,
            'currentvalue': {'prefix': '动画: '},
            'len': 0.9,
            'pad': {'t': 50},
            'steps': [],
            'x': 0.1,
            'xanchor': 'left',
            'y': 0,
            'yanchor': 'top'
        }]
    )
    
    # 添加悬停效果和动态交互
    fig.update_traces(
        hoverlabel=dict(
            bgcolor='rgba(21,101,192,0.9)',
            bordercolor='#1565C0',
            font=dict(color='white', size=12)
        )
    )
    
    return fig 

def create_smooth_donut_chart(data, title="项目进度概览"):
    """创建具有真正圆融圆弧效果的甜甜圈图"""
    
    # 获取三个进度指标
    year_usage = data['year_usage']
    month_usage = data['month_usage']
    time_progress = data['time_progress']
    
    # 创建甜甜圈图
    fig = go.Figure()
    
    # 定义蓝色渐变色彩
    blue_gradients = {
        'light': ['#E3F2FD', '#BBDEFB', '#90CAF9'],
        'medium': ['#90CAF9', '#64B5F6', '#42A5F5'],
        'dark': ['#42A5F5', '#2196F3', '#1E88E5']
    }
    
    # 使用Scatterpolar创建圆融的圆弧效果
    # 年使用率环（最内层）
    angles = np.linspace(0, 2*np.pi, 100)
    year_angles = angles[:int(len(angles) * year_usage / 100)]
    year_r = [0.3] * len(year_angles)
    
    fig.add_trace(go.Scatterpolar(
        r=year_r,
        theta=year_angles * 180 / np.pi,
        mode='lines',
        line=dict(color=blue_gradients['light'][1], width=8),
        name='年使用率',
        showlegend=False,
        hovertemplate="<b>年使用率</b><br>" +
                     "进度: " + f"{year_usage:.1f}%" + "<br>" +
                     "数值: " + f"{year_usage:.1f}%" + "<extra></extra>"
    ))
    
    # 月累使用率环（中间层）
    month_angles = angles[:int(len(angles) * month_usage / 100)]
    month_r = [0.5] * len(month_angles)
    
    fig.add_trace(go.Scatterpolar(
        r=month_r,
        theta=month_angles * 180 / np.pi,
        mode='lines',
        line=dict(color=blue_gradients['medium'][1], width=8),
        name='月累使用率',
        showlegend=False,
        hovertemplate="<b>月累使用率</b><br>" +
                     "进度: " + f"{month_usage:.1f}%" + "<br>" +
                     "数值: " + f"{month_usage:.1f}%" + "<extra></extra>"
    ))
    
    # 时间进度环（最外层）
    time_angles = angles[:int(len(angles) * time_progress / 100)]
    time_r = [0.7] * len(time_angles)
    
    fig.add_trace(go.Scatterpolar(
        r=time_r,
        theta=time_angles * 180 / np.pi,
        mode='lines',
        line=dict(color=blue_gradients['dark'][1], width=8),
        name='时间进度',
        showlegend=False,
        hovertemplate="<b>时间进度</b><br>" +
                     "进度: " + f"{time_progress:.1f}%" + "<br>" +
                     "数值: " + f"{time_progress:.1f}%" + "<extra></extra>"
    ))
    
    # 添加背景圆环
    background_angles = np.linspace(0, 360, 100)
    
    # 年使用率背景
    fig.add_trace(go.Scatterpolar(
        r=[0.3] * 100,
        theta=background_angles,
        mode='lines',
        line=dict(color='#f0f0f0', width=8),
        showlegend=False,
        hoverinfo='skip'
    ))
    
    # 月累使用率背景
    fig.add_trace(go.Scatterpolar(
        r=[0.5] * 100,
        theta=background_angles,
        mode='lines',
        line=dict(color='#f0f0f0', width=8),
        showlegend=False,
        hoverinfo='skip'
    ))
    
    # 时间进度背景
    fig.add_trace(go.Scatterpolar(
        r=[0.7] * 100,
        theta=background_angles,
        mode='lines',
        line=dict(color='#f0f0f0', width=8),
        showlegend=False,
        hoverinfo='skip'
    ))
    
    # 添加标签
    # 年使用率标签
    fig.add_annotation(
        x=1.35, y=0.5, xref='paper', yref='paper',
        text=f"{year_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#90CAF9',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#90CAF9',
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 月累使用率标签
    fig.add_annotation(
        x=0.5, y=1.25, xref='paper', yref='paper',
        text=f"{month_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#42A5F5',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#42A5F5',
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 时间进度标签
    fig.add_annotation(
        x=0.2, y=0.6, xref='paper', yref='paper',
        text=f"{time_progress:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor='#1E88E5',
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor='#1E88E5',
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 更新布局
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=18, color='#1565C0', family='Arial Black'),
            y=0.95
        ),
        polar=dict(
            radialaxis=dict(
                visible=False,
                range=[0, 1]
            ),
            angularaxis=dict(
                visible=False
            ),
            bgcolor='rgba(248,249,250,0.8)'
        ),
        height=450,
        width=600,
        showlegend=False,
        margin=dict(l=80, r=80, t=100, b=80),
        plot_bgcolor='rgba(248,249,250,0.8)',
        paper_bgcolor='white'
    )
    
    # 添加悬停效果
    fig.update_traces(
        hoverlabel=dict(
            bgcolor='rgba(21,101,192,0.9)',
            bordercolor='#1565C0',
            font=dict(color='white', size=12)
        )
    )
    
    return fig 

def create_perfect_donut_chart(data, title="项目进度概览"):
    """创建完美的圆融圆弧甜甜圈图"""
    
    # 获取三个进度指标
    year_usage = data['year_usage']
    month_usage = data['month_usage']
    time_progress = data['time_progress']
    
    # 创建甜甜圈图
    fig = go.Figure()
    
    # 定义蓝色渐变色彩
    blue_gradients = {
        'light': '#90CAF9',   # 浅蓝色
        'medium': '#42A5F5',  # 中蓝色
        'dark': '#1E88E5'     # 深蓝色
    }
    
    # 创建完整的圆弧路径
    def create_arc_path(radius, start_angle, end_angle, points=100):
        """创建圆弧路径"""
        angles = np.linspace(start_angle, end_angle, points)
        x = radius * np.cos(angles)
        y = radius * np.sin(angles)
        return x, y
    
    # 年使用率环（最内层）- 浅蓝色
    year_angle = (year_usage / 100) * 2 * np.pi
    year_x, year_y = create_arc_path(0.3, 0, year_angle)
    
    # 创建年使用率圆弧
    fig.add_trace(go.Scatter(
        x=year_x,
        y=year_y,
        mode='lines',
        line=dict(color=blue_gradients['light'], width=12, shape='spline'),
        name='年使用率',
        showlegend=False,
        hovertemplate="<b>年使用率</b><br>" +
                     "进度: " + f"{year_usage:.1f}%" + "<br>" +
                     "数值: " + f"{year_usage:.1f}%" + "<extra></extra>"
    ))
    
    # 月累使用率环（中间层）- 中蓝色
    month_angle = (month_usage / 100) * 2 * np.pi
    month_x, month_y = create_arc_path(0.5, 0, month_angle)
    
    fig.add_trace(go.Scatter(
        x=month_x,
        y=month_y,
        mode='lines',
        line=dict(color=blue_gradients['medium'], width=12, shape='spline'),
        name='月累使用率',
        showlegend=False,
        hovertemplate="<b>月累使用率</b><br>" +
                     "进度: " + f"{month_usage:.1f}%" + "<br>" +
                     "数值: " + f"{month_usage:.1f}%" + "<extra></extra>"
    ))
    
    # 时间进度环（最外层）- 深蓝色
    time_angle = (time_progress / 100) * 2 * np.pi
    time_x, time_y = create_arc_path(0.7, 0, time_angle)
    
    fig.add_trace(go.Scatter(
        x=time_x,
        y=time_y,
        mode='lines',
        line=dict(color=blue_gradients['dark'], width=12, shape='spline'),
        name='时间进度',
        showlegend=False,
        hovertemplate="<b>时间进度</b><br>" +
                     "进度: " + f"{time_progress:.1f}%" + "<br>" +
                     "数值: " + f"{time_progress:.1f}%" + "<extra></extra>"
    ))
    
    # 添加背景圆环
    # 年使用率背景
    bg_year_x, bg_year_y = create_arc_path(0.3, 0, 2*np.pi)
    fig.add_trace(go.Scatter(
        x=bg_year_x,
        y=bg_year_y,
        mode='lines',
        line=dict(color='#f0f0f0', width=12, shape='spline'),
        showlegend=False,
        hoverinfo='skip'
    ))
    
    # 月累使用率背景
    bg_month_x, bg_month_y = create_arc_path(0.5, 0, 2*np.pi)
    fig.add_trace(go.Scatter(
        x=bg_month_x,
        y=bg_month_y,
        mode='lines',
        line=dict(color='#f0f0f0', width=12, shape='spline'),
        showlegend=False,
        hoverinfo='skip'
    ))
    
    # 时间进度背景
    bg_time_x, bg_time_y = create_arc_path(0.7, 0, 2*np.pi)
    fig.add_trace(go.Scatter(
        x=bg_time_x,
        y=bg_time_y,
        mode='lines',
        line=dict(color='#f0f0f0', width=12, shape='spline'),
        showlegend=False,
        hoverinfo='skip'
    ))
    
    # 添加标签
    # 年使用率标签
    fig.add_annotation(
        x=1.35, y=0.5, xref='paper', yref='paper',
        text=f"{year_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=blue_gradients['light'],
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=blue_gradients['light'],
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 月累使用率标签
    fig.add_annotation(
        x=0.5, y=1.25, xref='paper', yref='paper',
        text=f"{month_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=blue_gradients['medium'],
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=blue_gradients['medium'],
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 时间进度标签
    fig.add_annotation(
        x=0.2, y=0.6, xref='paper', yref='paper',
        text=f"{time_progress:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=blue_gradients['dark'],
        font=dict(size=16, color='#1565C0', family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=blue_gradients['dark'],
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 更新布局
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=18, color='#1565C0', family='Arial Black'),
            y=0.95
        ),
        xaxis=dict(
            visible=False,
            range=[-1, 1]
        ),
        yaxis=dict(
            visible=False,
            range=[-1, 1],
            scaleanchor="x",
            scaleratio=1
        ),
        height=450,
        width=600,
        showlegend=False,
        margin=dict(l=80, r=80, t=100, b=80),
        plot_bgcolor='rgba(248,249,250,0.8)',
        paper_bgcolor='white'
    )
    
    # 添加悬停效果
    fig.update_traces(
        hoverlabel=dict(
            bgcolor='rgba(21,101,192,0.9)',
            bordercolor='#1565C0',
            font=dict(color='white', size=12)
        )
    )
    
    return fig 

def create_simple_donut_chart(data, title="项目进度概览"):
    """创建简化的甜甜圈图，避免参数错误"""
    
    # 获取三个进度指标
    year_usage = data['year_usage']
    month_usage = data['month_usage']
    time_progress = data['time_progress']
    
    # 创建甜甜圈图
    fig = go.Figure()
    
    # 定义颜色
    BAR_COLOR = '#42A5F5'  # 蓝色进度条
    BG_COLOR = '#f0f0f0'   # 背景色
    TEXT_COLOR = '#1565C0' # 文字颜色
    
    # 定义三个环的半径
    radii = [0.3, 0.5, 0.7]
    progress_values = [year_usage, month_usage, time_progress]
    labels = ['年使用率', '月累使用率', '时间进度']
    
    # 为每个环创建进度条效果
    for idx, (radius, progress, label) in enumerate(zip(radii, progress_values, labels)):
        # 创建完整的圆环作为背景
        angles = np.linspace(0, 2*np.pi, 200)
        x_bg = radius * np.cos(angles)
        y_bg = radius * np.sin(angles)
        
        # 背景圆环
        fig.add_trace(go.Scatter(
            x=x_bg,
            y=y_bg,
            mode='lines',
            line=dict(color=BG_COLOR, width=15),
            showlegend=False,
            hoverinfo='skip'
        ))
        
        # 创建进度圆弧
        progress_angle = (progress / 100) * 2 * np.pi
        progress_angles = np.linspace(0, progress_angle, max(100, int(progress * 3)))
        x_progress = radius * np.cos(progress_angles)
        y_progress = radius * np.sin(progress_angles)
        
        # 进度圆弧 - 不使用smoothing参数，避免错误
        fig.add_trace(go.Scatter(
            x=x_progress,
            y=y_progress,
            mode='lines',
            line=dict(
                color=BAR_COLOR, 
                width=15
            ),
            name=label,
            showlegend=False,
            hovertemplate=f"<b>{label}</b><br>" +
                         f"进度: {progress:.1f}%<br>" +
                         f"数值: {progress:.1f}%<extra></extra>"
        ))
        
        # 添加圆融的端部效果
        if progress > 0:
            # 末端圆点
            end_x = radius * np.cos(progress_angle)
            end_y = radius * np.sin(progress_angle)
            
            fig.add_trace(go.Scatter(
                x=[end_x],
                y=[end_y],
                mode='markers',
                marker=dict(
                    color=BAR_COLOR,
                    size=18,
                    line=dict(color='white', width=2),
                    symbol='circle'
                ),
                showlegend=False,
                hoverinfo='skip'
            ))
            
            # 起始圆点
            start_x = radius * np.cos(0)
            start_y = radius * np.sin(0)
            
            fig.add_trace(go.Scatter(
                x=[start_x],
                y=[start_y],
                mode='markers',
                marker=dict(
                    color=BAR_COLOR,
                    size=18,
                    line=dict(color='white', width=2),
                    symbol='circle'
                ),
                showlegend=False,
                hoverinfo='skip'
            ))
    
    # 添加标签 - 调整位置避免被截断
    # 年使用率标签
    fig.add_annotation(
        x=1.25, y=0.5, xref='paper', yref='paper',  # 调整x位置
        text=f"{year_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=BAR_COLOR,
        font=dict(size=16, color=TEXT_COLOR, family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=BAR_COLOR,
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 月累使用率标签
    fig.add_annotation(
        x=0.5, y=1.15, xref='paper', yref='paper',  # 调整y位置
        text=f"{month_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=BAR_COLOR,
        font=dict(size=16, color=TEXT_COLOR, family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=BAR_COLOR,
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 时间进度标签
    fig.add_annotation(
        x=0.25, y=0.6, xref='paper', yref='paper',  # 调整x位置
        text=f"{time_progress:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=BAR_COLOR,
        font=dict(size=16, color=TEXT_COLOR, family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=BAR_COLOR,
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 更新布局 - 增加边距避免标签被截断
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=18, color=TEXT_COLOR, family='Arial Black'),
            y=0.95
        ),
        xaxis=dict(
            visible=False,
            range=[-1, 1],
            scaleanchor="y",
            scaleratio=1
        ),
        yaxis=dict(
            visible=False,
            range=[-1, 1],
            scaleanchor="x",
            scaleratio=1
        ),
        height=500,
        width=700,
        showlegend=False,
        margin=dict(l=120, r=120, t=120, b=120),  # 增加边距
        plot_bgcolor='rgba(248,249,250,0.8)',
        paper_bgcolor='white'
    )
    
    # 添加悬停效果
    fig.update_traces(
        hoverlabel=dict(
            bgcolor='rgba(66,165,245,0.9)',
            bordercolor=BAR_COLOR,
            font=dict(color='white', size=12)
        )
    )
    
    return fig 

def create_echarts_style_donut_chart(data, title="项目进度概览"):
    """创建参考ECharts roundCap效果的甜甜圈图"""
    
    # 获取三个进度指标
    year_usage = data['year_usage']
    month_usage = data['month_usage']
    time_progress = data['time_progress']
    
    # 创建甜甜圈图
    fig = go.Figure()
    
    # 定义颜色
    BAR_COLOR = '#42A5F5'  # 蓝色进度条
    BG_COLOR = '#f0f0f0'   # 背景色
    TEXT_COLOR = '#1565C0' # 文字颜色
    
    # 定义三个环的半径（对应ECharts的radiusAxis）
    radii = [0.3, 0.5, 0.7]
    progress_values = [year_usage, month_usage, time_progress]
    labels = ['年使用率', '月累使用率', '时间进度']
    
    # 为每个环创建进度条效果
    for idx, (radius, progress, label) in enumerate(zip(radii, progress_values, labels)):
        # 创建完整的圆环作为背景
        angles = np.linspace(0, 2*np.pi, 200)
        x_bg = radius * np.cos(angles)
        y_bg = radius * np.sin(angles)
        
        # 背景圆环
        fig.add_trace(go.Scatter(
            x=x_bg,
            y=y_bg,
            mode='lines',
            line=dict(color=BG_COLOR, width=15),
            showlegend=False,
            hoverinfo='skip'
        ))
        
        # 创建进度圆弧 - 参考ECharts的roundCap效果
        progress_angle = (progress / 100) * 2 * np.pi
        
        # 使用更密集的点来模拟roundCap效果
        num_points = max(300, int(progress * 6))  # 大幅增加点数
        progress_angles = np.linspace(0, progress_angle, num_points)
        
        # 创建圆弧路径
        x_progress = radius * np.cos(progress_angles)
        y_progress = radius * np.sin(progress_angles)
        
        # 进度圆弧 - 使用spline插值实现圆融效果
        fig.add_trace(go.Scatter(
            x=x_progress,
            y=y_progress,
            mode='lines',
            line=dict(
                color=BAR_COLOR, 
                width=15, 
                shape='spline',  # 使用spline确保平滑
                smoothing=1.3    # 修正为1.3，符合Plotly要求
            ),
            name=label,
            showlegend=False,
            hovertemplate=f"<b>{label}</b><br>" +
                         f"进度: {progress:.1f}%<br>" +
                         f"数值: {progress:.1f}%<extra></extra>"
        ))
        
        # 添加圆融的端部效果 - 模拟ECharts的roundCap
        if progress > 0:
            # 末端圆点 - 使用更大的圆点模拟roundCap
            end_x = radius * np.cos(progress_angle)
            end_y = radius * np.sin(progress_angle)
            
            fig.add_trace(go.Scatter(
                x=[end_x],
                y=[end_y],
                mode='markers',
                marker=dict(
                    color=BAR_COLOR,
                    size=20,  # 增大圆点尺寸
                    line=dict(color='white', width=3),
                    symbol='circle'
                ),
                showlegend=False,
                hoverinfo='skip'
            ))
            
            # 起始圆点 - 确保起始端也是圆融的
            start_x = radius * np.cos(0)
            start_y = radius * np.sin(0)
            
            fig.add_trace(go.Scatter(
                x=[start_x],
                y=[start_y],
                mode='markers',
                marker=dict(
                    color=BAR_COLOR,
                    size=20,
                    line=dict(color='white', width=3),
                    symbol='circle'
                ),
                showlegend=False,
                hoverinfo='skip'
            ))
            
            # 添加额外的圆融效果 - 在圆弧两端添加小圆点
            if progress > 5:  # 只有当进度足够大时才添加
                # 在圆弧的1/4处添加小圆点
                quarter_angle = progress_angle * 0.25
                quarter_x = radius * np.cos(quarter_angle)
                quarter_y = radius * np.sin(quarter_angle)
                
                fig.add_trace(go.Scatter(
                    x=[quarter_x],
                    y=[quarter_y],
                    mode='markers',
                    marker=dict(
                        color=BAR_COLOR,
                        size=8,
                        line=dict(color='white', width=1),
                        symbol='circle'
                    ),
                    showlegend=False,
                    hoverinfo='skip'
                ))
                
                # 在圆弧的3/4处添加小圆点
                three_quarter_angle = progress_angle * 0.75
                three_quarter_x = radius * np.cos(three_quarter_angle)
                three_quarter_y = radius * np.sin(three_quarter_angle)
                
                fig.add_trace(go.Scatter(
                    x=[three_quarter_x],
                    y=[three_quarter_y],
                    mode='markers',
                    marker=dict(
                        color=BAR_COLOR,
                        size=8,
                        line=dict(color='white', width=1),
                        symbol='circle'
                    ),
                    showlegend=False,
                    hoverinfo='skip'
                ))
    
    # 添加标签
    # 年使用率标签
    fig.add_annotation(
        x=1.35, y=0.5, xref='paper', yref='paper',
        text=f"{year_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=BAR_COLOR,
        font=dict(size=16, color=TEXT_COLOR, family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=BAR_COLOR,
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 月累使用率标签
    fig.add_annotation(
        x=0.5, y=1.25, xref='paper', yref='paper',
        text=f"{month_usage:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=BAR_COLOR,
        font=dict(size=16, color=TEXT_COLOR, family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=BAR_COLOR,
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 时间进度标签
    fig.add_annotation(
        x=0.2, y=0.6, xref='paper', yref='paper',
        text=f"{time_progress:.0f}%",
        showarrow=True,
        arrowhead=2,
        arrowsize=1.2,
        arrowwidth=2,
        arrowcolor=BAR_COLOR,
        font=dict(size=16, color=TEXT_COLOR, family='Arial Black'),
        bgcolor='rgba(255,255,255,0.95)',
        bordercolor=BAR_COLOR,
        borderwidth=2,
        align='center',
        valign='middle'
    )
    
    # 更新布局 - 参考ECharts的极坐标系统
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=18, color=TEXT_COLOR, family='Arial Black'),
            y=0.95
        ),
        # 使用极坐标系统，类似ECharts的polar配置
        xaxis=dict(
            visible=False,
            range=[-1, 1],
            scaleanchor="y",
            scaleratio=1
        ),
        yaxis=dict(
            visible=False,
            range=[-1, 1],
            scaleanchor="x",
            scaleratio=1
        ),
        height=500,
        width=700,
        showlegend=False,
        margin=dict(l=100, r=100, t=120, b=100),
        plot_bgcolor='rgba(248,249,250,0.8)',
        paper_bgcolor='white'
    )
    
    # 添加悬停效果
    fig.update_traces(
        hoverlabel=dict(
            bgcolor='rgba(66,165,245,0.9)',
            bordercolor=BAR_COLOR,
            font=dict(color='white', size=12)
        )
    )
    
    return fig 

def create_three_donut_charts(data, month):
    """
    创建3个独立的甜甜圈图，分别显示时间进度、年使用率、月累使用率
    使用苹果风格设计，每个图表独立显示
    """
    import plotly.graph_objects as go
    import numpy as np
    
    # 计算三个指标
    time_progress = data.get('time_progress', (month / 12) * 100)  # 时间进度
    year_usage = data.get('year_usage', 0)  # 年使用率
    month_usage = data.get('month_usage', 0)  # 月累使用率
    
    # 调试信息已移除
    
    # 颜色规则函数
    def get_time_progress_color(progress):
        """时间进度颜色规则：≤50%蓝色，>50%橙色，>90%红色"""
        if progress <= 50:
            return '#007AFF'  # 苹果蓝
        elif progress > 90:
            return '#FF3B30'  # 苹果红
        else:
            return '#FF9500'  # 苹果橙
    
    def get_usage_color(usage):
        """使用率颜色规则：>90%红色，>60%橙色，≤60%绿色"""
        if usage > 90:
            return '#FF3B30'  # 苹果红
        elif usage > 60:
            return '#FF9500'  # 苹果橙
        else:
            return '#34C759'  # 苹果绿
    
    # 数据配置
    chart_configs = [
        {
            'title': '时间进度',
            'progress': time_progress,
            'value': time_progress,
            'unit': '%',
            'color': get_time_progress_color(time_progress),
            'bg_color': '#F2F2F7'  # 苹果浅灰
        },
        {
            'title': '年使用率',
            'progress': year_usage,
            'value': year_usage,
            'unit': '%',
            'color': get_usage_color(year_usage),
            'bg_color': '#F2F2F7'  # 苹果浅灰
        },
        {
            'title': '月累使用率',
            'progress': month_usage,
            'value': month_usage,
            'unit': '%',
            'color': get_usage_color(month_usage),
            'bg_color': '#F2F2F7'  # 苹果浅灰
        }
    ]
    
    # 创建三个独立的图表
    charts = []
    
    for idx, config in enumerate(chart_configs):
        # 创建甜甜圈图数据
        progress = min(config['progress'], 100)
        remaining = 100 - progress
        
        # 调试信息已移除
        
        # 创建独立的图表
        fig = go.Figure()
        
        # 添加甜甜圈图 - 正确的绘制逻辑
        # 起始角度：顶部（-90°），顺时针方向绘制
        fig.add_trace(
            go.Pie(
                values=[progress, remaining],
                hole=0.75,  # 更大的洞，更细的环
                marker_colors=[config['color'], config['bg_color']],
                showlegend=False,
                textinfo='none',
                hoverinfo='skip',
                rotation=-90,  # 从顶部开始（-90度）
                direction='clockwise',  # 顺时针方向
                textposition='inside'
            )
        )
        
        # 添加中心文本
        center_text = f"{config['value']:.1f}<br><span style='font-size: 16px; opacity: 0.7;'>{config['unit']}</span>"
        
        fig.add_annotation(
            x=0.5,
            y=0.5,
            text=center_text,
            showarrow=False,
            font=dict(size=28, color=config['color'], family='SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif', weight='bold'),
            xref='paper',
            yref='paper'
        )
        
        # 苹果风格布局
        fig.update_layout(
            title={
                'text': config['title'],
                'x': 0.5,
                'xanchor': 'center',
                'font': {
                    'size': 18,
                    'color': '#1D1D1F',
                    'family': 'SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif',
                    'weight': 'bold'
                },
                'y': 0.95
            },
            height=280,
            width=280,
            margin=dict(l=20, r=20, t=60, b=20),
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family='SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif')
        )
        
        charts.append(fig)
    
    return charts 

# 暂时注释掉 Matplotlib 版本，避免导入问题
# def create_three_donut_charts_matplotlib(data, month):
#     """
#     使用Matplotlib创建3个独立的甜甜圈图，分别显示时间进度、年使用率、月累使用率
#     使用Matplotlib可以更精确地控制绘制逻辑
#     """
#     pass

def create_tertiary_exception_chart(all_data, month):
    """
    创建三级费项异常项展示图
    显示所有项目的三级费项异常情况，包括年度目标为0的费项
    """
    import plotly.graph_objects as go
    import plotly.express as px
    from collections import defaultdict
    
    # 收集所有异常数据
    all_exceptions = []
    red_count = 0
    yellow_count = 0
    
    for project_name, data in all_data.items():
        if 'tertiary_exceptions' in data:
            for exception in data['tertiary_exceptions']:
                # 只统计到当前选择月份的异常
                if exception['month'] <= month:
                    # 添加项目名称到异常记录
                    exception_record = exception.copy()
                    exception_record['project_name'] = project_name
                    all_exceptions.append(exception_record)
                    
                    # 统计异常数量
                    if exception['exception_type'] == 'red':
                        red_count += 1
                    elif exception['exception_type'] == 'yellow':
                        yellow_count += 1
    
    total_exceptions = red_count + yellow_count
    
    if total_exceptions == 0:
        # 如果没有异常，显示空状态
        fig = go.Figure()
        fig.add_annotation(
            x=0.5, y=0.5,
            text="🎉 恭喜！<br>所有三级费项都在正常范围内",
            showarrow=False,
            font=dict(size=20, color="#34C759"),
            xref="paper", yref="paper"
        )
        fig.update_layout(
            title=f"三级费项异常情况 (第{month}月)",
            xaxis=dict(showgrid=False, showticklabels=False),
            yaxis=dict(showgrid=False, showticklabels=False),
            plot_bgcolor='white',
            height=400
        )
        return fig
    
    # 创建饼图（与前面异常项展示保持一致）
    fig = go.Figure()
    
    # 准备数据
    labels = []
    values = []
    colors = []
    
    if red_count > 0:
        labels.append('红色异常')
        values.append(red_count)
        colors.append('#ef4444')
    
    if yellow_count > 0:
        labels.append('黄色异常')
        values.append(yellow_count)
        colors.append('#f59e0b')
    

    
    # 创建饼图
    fig.add_trace(go.Pie(
        labels=labels,
        values=values,
        hole=0.5,
        marker_colors=colors,
        textinfo='label+value',
        textposition='inside',
        textfont=dict(size=12, color='white')
    ))
    
    # 更新布局
    fig.update_layout(
        title={
            'text': f'三级费项异常统计 (第{month}月)',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#1D1D1F'}
        },
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif"),
        height=400,
        showlegend=True,
        legend=dict(
            orientation="v",  # 垂直布局
            yanchor="top",
            y=1,
            xanchor="left",
            x=0,  # 图例放在左边
            itemsizing='constant'
        )
    )
    
    return fig

def create_tertiary_exception_details_table(all_data, month):
    """
    创建三级费项异常详情表格
    按照用户要求：Excel 文件名 | 月份 | 费项名称 | 异常类型 | 已发生金额累计 | 目标金额累计
    包括年度目标为0的费项
    """
    import plotly.graph_objects as go
    
    # 收集所有异常数据
    all_exceptions = []
    
    for project_name, data in all_data.items():
        if 'tertiary_exceptions' in data:
            for exception in data['tertiary_exceptions']:
                # 只显示到当前选择月份的异常
                if exception['month'] <= month:
                    # 添加项目名称到异常记录
                    exception_record = exception.copy()
                    exception_record['project_name'] = project_name
                    all_exceptions.append(exception_record)
    
    # 合并所有数据
    all_items = all_exceptions
    
    if not all_items:
        return None, None  # 返回None表示没有数据，第二个None是DataFrame
    
    # 按月份升序排序
    all_items.sort(key=lambda x: x.get('month', 0))
    
    # 创建表格数据
    table_data = []
    for item in all_items:
        exception_type_text = "超年度目标" if item['exception_type'] == 'red' else "超月度目标"
        
        table_data.append([
            item['project_name'],  # Excel 文件名
            f"{item.get('month', 'N/A')}月" if item.get('month') else "N/A",  # 月份
            item['fee_name'],      # 费项名称
            exception_type_text,   # 异常类型
            f"{item['cum_actual']:.2f}",  # 已发生金额累计
            f"{item['cum_target']:.2f}"   # 目标金额累计
        ])
    
    # 创建DataFrame用于下载
    df_data = []
    for item in all_items:
        exception_type_text = "超年度目标" if item['exception_type'] == 'red' else "超月度目标"
        
        df_data.append({
            'Excel 文件名': item['project_name'],
            '月份': f"{item.get('month', 'N/A')}月" if item.get('month') else "N/A",
            '费项名称': item['fee_name'],
            '异常类型': exception_type_text,
            '已发生金额累计': round(item['cum_actual'], 2),
            '目标金额累计': round(item['cum_target'], 2)
        })
    
    df = pd.DataFrame(df_data)
    
    # 创建表格
    fig = go.Figure(data=[go.Table(
        header=dict(
            values=['Excel 文件名', '月份', '费项名称', '异常类型', '已发生金额累计', '目标金额累计'],
            fill_color='#007AFF',
            font=dict(color='white', size=12),
            align='left'
        ),
        cells=dict(
            values=list(zip(*table_data)),
            fill_color='white',
            font=dict(color='#1D1D1F', size=11),
            align='left',
            height=30
        )
    )])
    
    fig.update_layout(
        title={
            'text': f'三级费项异常详情 (第{month}月)',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 16, 'color': '#1D1D1F'}
        },
        height=400,
        margin=dict(l=20, r=20, t=60, b=20)
    )
    
    return fig, df
    return fig