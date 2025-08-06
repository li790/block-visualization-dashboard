import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from utils.chart_creator import (
    create_pie_chart, 
    create_bar_line_chart, 
    create_project_comparison_chart,
    create_kpi_display,
    create_multi_ring_progress_chart,
    create_semi_circular_chart,
    create_donut_chart,
    create_rounded_donut_chart,
    create_smooth_donut_chart,
    create_perfect_donut_chart,
    create_echarts_style_donut_chart,
    create_simple_donut_chart,
    create_three_donut_charts, # Added this import
    create_bar_chart, # Added create_bar_chart import
    create_line_chart, # Added create_line_chart import
    create_tertiary_exception_chart, # Added tertiary exception chart
    create_tertiary_exception_details_table, # Added tertiary exception details table
    # create_three_donut_charts_matplotlib # 暂时注释掉 Matplotlib 版本
)

from utils.data_processor import create_summary_excel, merge_project_data, create_client_download_table, create_monthly_fee_summary

def render_kpi_metrics(all_data, all_dfs, month):
    """渲染KPI指标 - 苹果风格卡片设计"""
    # 判断是单项目还是多项目模式
    if len(all_data) == 1:
        # 单项目模式：不显示金额指标卡片，保持与多项目模式一致
        pass
    else:
        # 多项目模式：显示合并后的关键指标
        merged_data = merge_project_data(all_data, all_dfs, month)
        if merged_data:
            st.markdown("### 多项目合并指标")
            
            # 创建上下布局：上面3个甜甜圈图，下面金额指标
            # 上面显示3个独立的甜甜圈图
            three_donut_charts = create_three_donut_charts(merged_data, month)
            
            # 使用3列布局显示独立的甜甜圈图
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.plotly_chart(three_donut_charts[0], use_container_width=True, key="donut_chart_1_merged")
            
            with col2:
                st.plotly_chart(three_donut_charts[1], use_container_width=True, key="donut_chart_2_merged")
            
            with col3:
                st.plotly_chart(three_donut_charts[2], use_container_width=True, key="donut_chart_3_merged")
            
            # 下面显示金额指标 - 使用更紧凑的布局
            st.markdown("#### 金额指标详情")
            
            # 渐变蓝色模块CSS - 统一蓝色渐变风格，减小尺寸
            card_style = """
            <style>
            .gradient-blue-card-compact {
                background: linear-gradient(135deg, #E3F2FD 0%, #90CAF9 25%, #42A5F5 50%, #1E88E5 75%, #1565C0 100%);
                border-radius: 12px;
                padding: 15px;
                margin: 8px 0;
                box-shadow: 0 6px 20px rgba(66,165,245,0.3);
                color: white;
                text-align: center;
                transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
                position: relative;
                overflow: hidden;
                border: 2px solid rgba(255,255,255,0.1);
            }
            .gradient-blue-card-compact::before {
                content: '';
                position: absolute;
                top: 0;
                left: -100%;
                width: 100%;
                height: 100%;
                background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
                transition: left 0.6s ease;
            }
            .gradient-blue-card-compact:hover {
                transform: translateY(-4px) scale(1.01);
                box-shadow: 0 10px 25px rgba(66,165,245,0.4);
            }
            .gradient-blue-card-compact:hover::before {
                left: 100%;
            }
            .metric-value-compact {
                font-size: 20px;
                font-weight: bold;
                margin: 8px 0;
                text-shadow: 0 2px 4px rgba(0,0,0,0.2);
                position: relative;
                z-index: 2;
            }
            .metric-label-compact {
                font-size: 12px;
                opacity: 0.95;
                font-weight: 500;
                position: relative;
                z-index: 2;
            }
            .metric-unit-compact {
                font-size: 10px;
                opacity: 0.8;
                position: relative;
                z-index: 2;
            }
            </style>
            """
            st.markdown(card_style, unsafe_allow_html=True)
            
            # 使用3列布局显示金额指标
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # 年总目标模块
                target_html = f"""
                <div class="gradient-blue-card-compact">
                    <div class="metric-label-compact">年总目标</div>
                    <div class="metric-value-compact">{merged_data['year_cum_target_wy']:.2f}</div>
                    <div class="metric-unit-compact">万元</div>
                </div>
                """
                st.markdown(target_html, unsafe_allow_html=True)
            
            with col2:
                # 月累目标模块
                month_target_html = f"""
                <div class="gradient-blue-card-compact">
                    <div class="metric-label-compact">月累目标</div>
                    <div class="metric-value-compact">{merged_data['cum_target'][month-1]/10000:.2f}</div>
                    <div class="metric-unit-compact">万元</div>
                </div>
                """
                st.markdown(month_target_html, unsafe_allow_html=True)
            
            with col3:
                # 月累已发生模块
                month_actual_html = f"""
                <div class="gradient-blue-card-compact">
                    <div class="metric-label-compact">月累已发生</div>
                    <div class="metric-value-compact">{merged_data['cum_actual'][month-1]/10000:.2f}</div>
                    <div class="metric-unit-compact">万元</div>
                </div>
                """
                st.markdown(month_actual_html, unsafe_allow_html=True)
            
            # 显示参与合并的项目列表
            st.markdown("---")
            st.markdown("**参与合并的项目:**")
            if 'merged_projects' in merged_data:
                project_list = ", ".join(merged_data['merged_projects'])
            else:
                project_list = ", ".join(all_data.keys())
            st.info(project_list)

def render_single_project_analysis(project_name, data, month, show_details=True):
    """渲染单项目详细分析 - 采用苹果风格卡片设计"""
    # 只有在需要显示详细信息时才显示标题
    if show_details:
        st.subheader(f"📊 {project_name} 详细分析")
        
        # 项目详细信息 - 显示甜甜圈图
        st.markdown("#### 项目进度指标")
        
        # 显示3个独立的甜甜圈图
        three_donut_charts = create_three_donut_charts(data, month)
        
        # 使用3列布局显示独立的甜甜圈图
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.plotly_chart(three_donut_charts[0], use_container_width=True, key=f"donut_chart_1_detail_{project_name}")
        
        with col2:
            st.plotly_chart(three_donut_charts[1], use_container_width=True, key=f"donut_chart_2_detail_{project_name}")
        
        with col3:
            st.plotly_chart(three_donut_charts[2], use_container_width=True, key=f"donut_chart_3_detail_{project_name}")
        
        # 显示其他详细信息
        st.markdown("#### 📋 项目详细信息")
        
        # 苹果风格卡片CSS
        detail_card_style = """
        <style>
        .detail-card {
            background: linear-gradient(135deg, #E3F2FD 0%, #90CAF9 25%, #42A5F5 50%, #1E88E5 75%, #1565C0 100%);
            border-radius: 16px;
            padding: 20px;
            margin: 15px 0;
            box-shadow: 0 8px 25px rgba(66,165,245,0.3);
            color: white;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
            border: 2px solid rgba(255,255,255,0.1);
        }
        .detail-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
            transition: left 0.6s ease;
        }
        .detail-card:hover {
            transform: translateY(-6px) scale(1.02);
            box-shadow: 0 12px 30px rgba(66,165,245,0.4);
        }
        .detail-card:hover::before {
            left: 100%;
        }
        .detail-label {
            font-size: 14px;
            opacity: 0.9;
            font-weight: 500;
            margin-bottom: 8px;
            position: relative;
            z-index: 2;
        }
        .detail-value {
            font-size: 24px;
            font-weight: bold;
            margin: 10px 0;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
            position: relative;
            z-index: 2;
        }
        .detail-unit {
            font-size: 12px;
            opacity: 0.8;
            position: relative;
            z-index: 2;
        }
        </style>
        """
        st.markdown(detail_card_style, unsafe_allow_html=True)
        
        # 使用3列布局显示金额指标
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # 年总目标
            target_html = f"""
            <div class="detail-card">
                <div class="detail-label">年总目标</div>
                <div class="detail-value">{data['year_cum_target_wy']:.2f}</div>
                <div class="detail-unit">万元</div>
            </div>
            """
            st.markdown(target_html, unsafe_allow_html=True)
        
        with col2:
            # 月累目标
            month_target_html = f"""
            <div class="detail-card">
                <div class="detail-label">月累目标</div>
                <div class="detail-value">{data['cum_target'][month-1]/10000:.2f}</div>
                <div class="detail-unit">万元</div>
            </div>
            """
            st.markdown(month_target_html, unsafe_allow_html=True)
        
        with col3:
            # 月累已发生
            month_actual_html = f"""
            <div class="detail-card">
                <div class="detail-label">月累已发生</div>
                <div class="detail-value">{data['cum_actual'][month-1]/10000:.2f}</div>
                <div class="detail-unit">万元</div>
            </div>
            """
            st.markdown(month_actual_html, unsafe_allow_html=True)
    
    # 图表控制选项
    st.markdown("#### 二级费项展示")
    control_col1, control_col2, control_col3, control_col4 = st.columns(4)
    
    with control_col1:
        show_target_bars = st.checkbox("显示目标累计柱状图", value=True, key=f"show_target_bars_{project_name}")
    with control_col2:
        show_actual_bars = st.checkbox("显示已发生累计柱状图", value=True, key=f"show_actual_bars_{project_name}")
    with control_col3:
        show_target_line = st.checkbox("显示总目标累计折线", value=True, key=f"show_target_line_{project_name}")
    with control_col4:
        show_actual_line = st.checkbox("显示总已发生成本折线", value=True, key=f"show_actual_line_{project_name}")
    
    # 创建上下排布的图表
    chart = create_bar_line_chart(data)
    if chart is not None:
        # 控制柱状图显示（第一个子图中的轨迹）
        if not show_target_bars and len(chart.data) >= 1:
            chart.data[0].visible = False
        if not show_actual_bars and len(chart.data) >= 2:
            chart.data[1].visible = False
        
        # 控制折线图显示（第二个子图中的轨迹）
        if not show_target_line and len(chart.data) >= 3:
            chart.data[2].visible = False
        if not show_actual_line and len(chart.data) >= 4:
            chart.data[3].visible = False
        
        # 显示组合图表
        st.plotly_chart(chart, use_container_width=True, height=700, key=f"combined_chart_{project_name}")
        

    else:
        st.warning("没有找到二级费项数据")

def render_anomaly_section(anomalies, project_name=None, all_data=None, month=None, section_type="main"):
    """渲染主控费项展示区域"""
    st.subheader("主控费项异常监控")
    
    # 月份筛选器
    all_months = sorted({anomaly['month'] for anomaly in anomalies})
    if all_months:
        selected_month = st.select_slider(
            "选择月份",
            options=list(range(1, 13)),
            value=all_months[-1] if all_months else 1,
            key=f"anomaly_month_slider_{project_name or 'multi'}_{section_type}"
        )
        
        # 过滤到选中月份的数据
        filtered_anomalies = [a for a in anomalies if a['month'] <= selected_month]
        
        # 统计异常数量
        red_count = sum(1 for a in filtered_anomalies if a['exception_type'] == 'red')
        yellow_count = sum(1 for a in filtered_anomalies if a['exception_type'] == 'yellow')
        
        # 创建饼图（与三级费项保持一致）
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
                'text': f'主控费项异常统计 (第{selected_month}月)',
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
        
        # 左右排布显示
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.plotly_chart(fig, use_container_width=True, key=f"main_anomaly_chart_{project_name or 'multi'}_{section_type}")
        
        with col2:
            if filtered_anomalies:
                # 显示带颜色的表格
                st.subheader("主控费项异常详情")
                
                # 准备表格数据
                table_data = []
                for anomaly in filtered_anomalies:
                    row = {
                        '项目名称': anomaly.get('project_name', project_name) if (anomaly.get('project_name') or project_name) else '未知项目',
                        '月份': anomaly['month'],
                        '费项名称': anomaly['fee_name'],
                        '异常类型': '红色异常' if anomaly['exception_type'] == 'red' else '黄色异常',
                        '累计已发生': round(anomaly['cum_actual'], 2),
                        '当月目标': round(anomaly['cum_target'], 2),
                        '年总目标': round(anomaly['year_target'], 2)
                    }
                    table_data.append(row)
                
                # 创建DataFrame并显示
                df = pd.DataFrame(table_data)
                
                # 添加样式
                def highlight_anomaly(row):
                    if row['异常类型'] == '红色异常':
                        return ['background-color: #fee2e2']*len(row)  # 红色背景
                    elif row['异常类型'] == '黄色异常':
                        return ['background-color: #fef3c7']*len(row)  # 黄色背景
                    return ['']*len(row)
                
                styled_df = df.style.apply(highlight_anomaly, axis=1)
                st.dataframe(styled_df, use_container_width=True)

                # 添加异常明细下载按钮
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 下载主控费项异常明细",
                    data=csv,
                    file_name=f"主控费项异常明细_{selected_month}月.csv",
                    mime="text/csv",
                    key=f"download_main_anomaly_detail_{project_name or 'multi'}_{section_type}"
                )
            else:
                st.info("没有找到异常数据")
    else:
        st.info("没有找到异常数据")
    
    # 添加三级费项异常监控作为组成部分
    if all_data and month:
        st.markdown("---")
        st.subheader(" 三级费项异常监控")
        
        # 创建三级费项异常图表和表格
        # 使用汇总的异常数据创建图表
        exception_chart = create_tertiary_exception_chart(all_data, selected_month if all_months else month)
        exception_result = create_tertiary_exception_details_table(all_data, selected_month if all_months else month)
        
        if exception_result:
            exception_table, exception_df = exception_result
        else:
            exception_table, exception_df = None, None
        
        # 显示异常统计图表
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.plotly_chart(exception_chart, use_container_width=True, key=f"tertiary_exception_chart_{project_name or 'multi'}_{section_type}")
        
        with col2:
            if exception_df is not None and not exception_df.empty:
                # 显示带颜色的表格
                st.subheader("三级费项异常详情")
                
                # 添加样式函数
                def highlight_tertiary_anomaly(row):
                    if row['异常类型'] == '超年度目标':
                        return ['background-color: #fee2e2']*len(row)  # 红色背景
                    elif row['异常类型'] == '超月度目标':
                        return ['background-color: #fef3c7']*len(row)  # 黄色背景
                    return ['']*len(row)
                
                # 应用样式
                styled_df = exception_df.style.apply(highlight_tertiary_anomaly, axis=1)
                st.dataframe(styled_df, use_container_width=True)
                
                # 添加下载按钮
                csv = exception_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 下载三级费项异常明细",
                    data=csv,
                    file_name=f"三级费项异常明细_{selected_month if all_months else month}月.csv",
                    mime="text/csv",
                    key=f"download_tertiary_exception_{project_name or 'multi'}_{section_type}"
                )
            else:
                st.info("暂无三级费项异常数据")

def render_multi_project_analysis(all_data, all_main_dfs, all_tertiary_dfs, month):
    """渲染多项目对比分析"""
    st.subheader("📊 项目对比分析")
    
    # 项目使用率对比图
    st.plotly_chart(create_project_comparison_chart(all_data, month), use_container_width=True, key="project_comparison_chart")
    
    # 每月费项数据处理
    monthly_fee_df = create_monthly_fee_summary(all_main_dfs)
    if monthly_fee_df is not None:
        # 格式化数据，保留2位小数
        monthly_fee_df['目标成本'] = monthly_fee_df['目标成本'].round(2)
        monthly_fee_df['已发生成本'] = monthly_fee_df['已发生成本'].round(2)
        
        # 绘制柱状图
        st.subheader("📊 每月费项成本柱状图")
        bar_chart = create_bar_chart(monthly_fee_df)
        st.plotly_chart(bar_chart, use_container_width=True, key="monthly_fee_bar_chart")
        
        # 添加下载按钮
        csv = monthly_fee_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 下载每月费项汇总表",
            data=csv,
            file_name="每月费项成本汇总.csv",
            mime="text/csv",
            key="download_monthly_fee_summary"
        )
    else:
        st.info("没有找到可用的费项数据")
    
    # 详细指标对比表
    st.subheader("📋 详细指标对比")
    comparison_data = create_kpi_display(all_data)
    df_comparison = pd.DataFrame(comparison_data)
    st.dataframe(df_comparison, use_container_width=True)
    
    # 下载选项
    col1, col2 = st.columns(2)
    
    with col1:
        # 下载月度汇总表（CSV格式）
        csv = df_comparison.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 下载月度汇总表",
            data=csv,
            file_name=f"项目汇总_{month}月.csv",
            mime="text/csv",
            key="download_csv_multi"
        )
    
    with col2:
        # 下载原始数据汇总表（Excel格式）
        if all_main_dfs:
            summary_df = create_summary_excel(all_main_dfs)
            if summary_df is not None:
                # 将DataFrame转换为Excel字节流
                import io
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary_df.to_excel(writer, index=False, sheet_name='汇总数据')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 下载原始数据汇总表",
                    data=excel_data,
                    file_name="项目原始数据汇总.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel_multi"
                )
    
    # 异常项展示 - 合并所有项目的异常
    all_anomalies = []
    all_tertiary_exceptions = []
    
    for project_name, data in all_data.items():
        # 收集主控费项异常
        if 'exceptions' in data and data['exceptions']:
            for anomaly in data['exceptions']:
                # 添加项目名称到异常数据
                anomaly_with_project = anomaly.copy()
                anomaly_with_project['project_name'] = project_name
                all_anomalies.append(anomaly_with_project)
        
        # 收集三级费项异常
        if 'tertiary_exceptions' in data and data['tertiary_exceptions']:
            for exception in data['tertiary_exceptions']:
                # 添加项目名称到异常数据
                exception_with_project = exception.copy()
                exception_with_project['project_name'] = project_name
                all_tertiary_exceptions.append(exception_with_project)
    
    if all_anomalies or all_tertiary_exceptions:
        # 创建汇总的异常数据
        combined_data = {}
        for project_name, data in all_data.items():
            combined_data[project_name] = data.copy()
            # 确保每个项目都有异常数据字段
            if 'exceptions' not in combined_data[project_name]:
                combined_data[project_name]['exceptions'] = []
            if 'tertiary_exceptions' not in combined_data[project_name]:
                combined_data[project_name]['tertiary_exceptions'] = []
        
        # 将汇总的三级费项异常数据添加到combined_data中
        for project_name, data in all_data.items():
            if project_name in combined_data:
                # 收集该项目所有的三级费项异常
                project_tertiary_exceptions = []
                for exception in all_tertiary_exceptions:
                    if exception.get('project_name') == project_name:
                        project_tertiary_exceptions.append(exception)
                combined_data[project_name]['tertiary_exceptions'] = project_tertiary_exceptions
        
        render_anomaly_section(all_anomalies, all_data=combined_data, month=month, section_type="multi")
    else:
        st.info("没有找到异常数据")
    
    # 添加项目选择下拉框和详细图表
    st.divider()
    st.subheader("📈 项目详细分析")
    
    # 项目选择下拉框
    project_names = list(all_data.keys())
    selected_project = st.selectbox(
        "选择要查看详细分析的项目:",
        project_names,
        index=0,
        key="project_selector_multi"
    )
    
    if selected_project:
        selected_data = all_data[selected_project]
        
        # 显示选中项目的详细分析 - 显示甜甜圈图
        st.markdown(f"#### 📊 {selected_project} 进度指标")
        
        # 显示3个独立的甜甜圈图
        three_donut_charts = create_three_donut_charts(selected_data, month)
        
        # 使用3列布局显示独立的甜甜圈图
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.plotly_chart(three_donut_charts[0], use_container_width=True, key=f"donut_chart_1_multi_{selected_project}")
        
        with col2:
            st.plotly_chart(three_donut_charts[1], use_container_width=True, key=f"donut_chart_2_multi_{selected_project}")
        
        with col3:
            st.plotly_chart(three_donut_charts[2], use_container_width=True, key=f"donut_chart_3_multi_{selected_project}")
        
        # 显示其他详细信息
        st.markdown(f"#### 📋 {selected_project} 详细信息")
        
        # 苹果风格卡片CSS - 与前面保持一致
        multi_detail_card_style = """
        <style>
        .multi-detail-card {
            background: linear-gradient(135deg, #E3F2FD 0%, #90CAF9 25%, #42A5F5 50%, #1E88E5 75%, #1565C0 100%);
            border-radius: 16px;
            padding: 20px;
            margin: 15px 0;
            box-shadow: 0 8px 25px rgba(66,165,245,0.3);
            color: white;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
            border: 2px solid rgba(255,255,255,0.1);
        }
        .multi-detail-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
            transition: left 0.6s ease;
        }
        .multi-detail-card:hover {
            transform: translateY(-6px) scale(1.02);
            box-shadow: 0 12px 30px rgba(66,165,245,0.4);
        }
        .multi-detail-card:hover::before {
            left: 100%;
        }
        .multi-detail-label {
            font-size: 14px;
            opacity: 0.9;
            font-weight: 500;
            margin-bottom: 8px;
            position: relative;
            z-index: 2;
        }
        .multi-detail-value {
            font-size: 24px;
            font-weight: bold;
            margin: 10px 0;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
            position: relative;
            z-index: 2;
        }
        .multi-detail-unit {
            font-size: 12px;
            opacity: 0.8;
            position: relative;
            z-index: 2;
        }
        </style>
        """
        st.markdown(multi_detail_card_style, unsafe_allow_html=True)
        
        # 使用3列布局显示金额指标
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # 年总目标
            target_html = f"""
            <div class="multi-detail-card">
                <div class="multi-detail-label">年总目标</div>
                <div class="multi-detail-value">{selected_data['year_cum_target_wy']:.2f}</div>
                <div class="multi-detail-unit">万元</div>
            </div>
            """
            st.markdown(target_html, unsafe_allow_html=True)
        
        with col2:
            # 月累目标
            month_target_html = f"""
            <div class="multi-detail-card">
                <div class="multi-detail-label">月累目标</div>
                <div class="multi-detail-value">{selected_data['cum_target'][month-1]/10000:.2f}</div>
                <div class="multi-detail-unit">万元</div>
            </div>
            """
            st.markdown(month_target_html, unsafe_allow_html=True)
        
        with col3:
            # 月累已发生
            month_actual_html = f"""
            <div class="multi-detail-card">
                <div class="multi-detail-label">月累已发生</div>
                <div class="multi-detail-value">{selected_data['cum_actual'][month-1]/10000:.2f}</div>
                <div class="multi-detail-unit">万元</div>
            </div>
            """
            st.markdown(month_actual_html, unsafe_allow_html=True)
        
        # 图表控制选项
        st.markdown("#### 二级费项图表展示")
        control_col1, control_col2, control_col3, control_col4 = st.columns(4)
        
        with control_col1:
            show_target_bars = st.checkbox("显示目标累计柱状图", value=True, key=f"show_target_bars_multi_{selected_project}")
        with control_col2:
            show_actual_bars = st.checkbox("显示已发生累计柱状图", value=True, key=f"show_actual_bars_multi_{selected_project}")
        with control_col3:
            show_target_line = st.checkbox("显示总目标累计折线", value=True, key=f"show_target_line_multi_{selected_project}")
        with control_col4:
            show_actual_line = st.checkbox("显示总已发生成本折线", value=True, key=f"show_actual_line_multi_{selected_project}")
        
        # 创建两张独立的图表
        bar_chart = create_bar_chart(selected_data)
        line_chart = create_line_chart(selected_data)
        
        if bar_chart and line_chart:
            # 控制柱状图显示
            if not show_target_bars and len(bar_chart.data) >= 1:
                bar_chart.data[0].visible = False
            if not show_actual_bars and len(bar_chart.data) >= 2:
                bar_chart.data[1].visible = False
            
            # 控制折线图显示
            if not show_target_line and len(line_chart.data) >= 1:
                line_chart.data[0].visible = False
            if not show_actual_line and len(line_chart.data) >= 2:
                line_chart.data[1].visible = False
            
            # 上下排列显示图表
            st.plotly_chart(bar_chart, use_container_width=True, height=300, key=f"bar_chart_multi_{selected_project}")
            st.plotly_chart(line_chart, use_container_width=True, height=300, key=f"line_chart_multi_{selected_project}")
        else:
            st.warning("没有找到二级费项数据")

def render_dashboard(all_data, all_main_dfs, all_tertiary_dfs, month):
    """渲染主仪表盘"""
    if not all_data:
        st.error("没有可用的数据")
        return
    
    # 显示KPI指标
    render_kpi_metrics(all_data, all_main_dfs, month)
    
    # 根据项目数量选择显示方式
    if len(all_data) == 1:
        # 单项目模式：显示详细指标
        project_name = list(all_data.keys())[0]
        data = all_data[project_name]
        
        # 显示3个独立的甜甜圈图
        st.markdown("#### 进度指标可视化")
        three_donut_charts = create_three_donut_charts(data, month)
        
        # 使用3列布局显示独立的甜甜圈图
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.plotly_chart(three_donut_charts[0], use_container_width=True, key=f"donut_chart_1_single_{project_name}")
        
        with col2:
            st.plotly_chart(three_donut_charts[1], use_container_width=True, key=f"donut_chart_2_single_{project_name}")
        
        with col3:
            st.plotly_chart(three_donut_charts[2], use_container_width=True, key=f"donut_chart_3_single_{project_name}")
        
        # 显示金额指标详情 - 与多项目模式保持一致
        st.markdown("#### 金额指标详情")
        
        # 渐变蓝色模块CSS - 统一蓝色渐变风格，减小尺寸
        card_style = """
        <style>
        .gradient-blue-card-compact {
            background: linear-gradient(135deg, #E3F2FD 0%, #90CAF9 25%, #42A5F5 50%, #1E88E5 75%, #1565C0 100%);
            border-radius: 12px;
            padding: 15px;
            margin: 8px 0;
            box-shadow: 0 6px 20px rgba(66,165,245,0.3);
            color: white;
            text-align: center;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
            border: 2px solid rgba(255,255,255,0.1);
        }
        .gradient-blue-card-compact::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
            transition: left 0.6s ease;
        }
        .gradient-blue-card-compact:hover {
            transform: translateY(-4px) scale(1.01);
            box-shadow: 0 10px 25px rgba(66,165,245,0.4);
        }
        .gradient-blue-card-compact:hover::before {
            left: 100%;
        }
        .metric-value-compact {
            font-size: 20px;
            font-weight: bold;
            margin: 8px 0;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
            position: relative;
            z-index: 2;
        }
        .metric-label-compact {
            font-size: 12px;
            opacity: 0.95;
            font-weight: 500;
            position: relative;
            z-index: 2;
        }
        .metric-unit-compact {
            font-size: 10px;
            opacity: 0.8;
            position: relative;
            z-index: 2;
        }
        </style>
        """
        st.markdown(card_style, unsafe_allow_html=True)
        
        # 使用3列布局显示金额指标
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # 年总目标模块
            target_html = f"""
            <div class="gradient-blue-card-compact">
                <div class="metric-label-compact">年总目标</div>
                <div class="metric-value-compact">{data.get('year_cum_target_wy', 0):.2f}</div>
                <div class="metric-unit-compact">万元</div>
            </div>
            """
            st.markdown(target_html, unsafe_allow_html=True)
        
        with col2:
            # 年累计已发生模块
            actual_html = f"""
            <div class="gradient-blue-card-compact">
                <div class="metric-label-compact">年累计已发生</div>
                <div class="metric-value-compact">{data.get('year_cum_actual_wy', 0):.2f}</div>
                <div class="metric-unit-compact">万元</div>
            </div>
            """
            st.markdown(actual_html, unsafe_allow_html=True)
        
        with col3:
            # 月累目标模块
            month_target_html = f"""
            <div class="gradient-blue-card-compact">
                <div class="metric-label-compact">月累目标</div>
                <div class="metric-value-compact">{data.get('cum_target', [0])[month-1]/10000:.2f}</div>
                <div class="metric-unit-compact">万元</div>
            </div>
            """
            st.markdown(month_target_html, unsafe_allow_html=True)
        
        # 显示异常项
        if 'exceptions' in data and data['exceptions']:
            render_anomaly_section(data['exceptions'], project_name, {project_name: data}, month, "single")
        else:
            st.info("没有找到异常数据")
        
        # 显示详细分析
        render_single_project_analysis(project_name, data, month)
    else:
        # 多项目模式：显示项目对比分析
        render_multi_project_analysis(all_data, all_main_dfs, all_tertiary_dfs, month)
    
    # 添加客户下载表格功能
    st.markdown("---")
    st.subheader("📥 项目数据导出")
    
    # 创建客户下载表格
    client_table = create_client_download_table(all_main_dfs, all_data)
    
    # 显示表格预览
    st.markdown("#### 📋 项目数据表格预览")
    st.dataframe(client_table, use_container_width=True)
    
    # 提供下载按钮
    st.markdown("#### 📥 下载数据")
    
    # 转换为Excel格式
    output = pd.ExcelWriter('temp_client_data.xlsx', engine='openpyxl')
    client_table.to_excel(output, sheet_name='客户数据', index=False)
    output.close()
    
    # 读取文件并创建下载按钮
    with open('temp_client_data.xlsx', 'rb') as file:
        st.download_button(
            label="📥 下载项目数据表格 (Excel)",
            data=file.read(),
            file_name=f"客户数据表格_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="下载包含所有项目12个月数据的Excel表格"
        )
    
    # 清理临时文件
    import os
    if os.path.exists('temp_client_data.xlsx'):
        os.remove('temp_client_data.xlsx')