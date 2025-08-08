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
    create_secondary_fee_combined_chart, # Added secondary fee combined chart
    # create_three_donut_charts_matplotlib # 暂时注释掉 Matplotlib 版本
)

from utils.data_processor import create_summary_excel, merge_project_data, create_client_download_table, create_monthly_fee_summary, create_secondary_fee_overall_data, extract_labor_service_breakdown, create_labor_service_summary

def render_kpi_metrics(all_data, all_dfs, month):
    """渲染KPI指标 - 苹果风格卡片设计"""
    # 确保month是整数类型
    month = int(month) if isinstance(month, str) else month
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
    # 确保month是整数类型
    month = int(month) if isinstance(month, str) else month
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
    
    # 创建标题和按钮的布局
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.subheader("主控费项异常监控")
    
    with col2:
        if all_data and section_type == "multi":  # 只在多项目模式下显示排行榜按钮
            # 初始化session state
            if 'show_exception_ranking' not in st.session_state:
                st.session_state.show_exception_ranking = False
            
            # 根据当前状态显示不同的按钮文本
            button_text = "❌ 收起排行榜" if st.session_state.show_exception_ranking else "📊 异常排行榜"
            
            if st.button(button_text, key="exception_ranking_btn"):
                st.session_state.show_exception_ranking = not st.session_state.show_exception_ranking
                st.rerun()  # 立即重新运行页面
    
    # 显示异常排行榜（如果被激活）
    if st.session_state.get('show_exception_ranking', False) and all_data and section_type == "multi":
        try:
            render_exception_ranking(all_data, month)
        except Exception as e:
            st.error(f"显示排行榜时出错: {e}")
            st.session_state.show_exception_ranking = False  # 出错时重置状态
    
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
    # 确保month是整数类型
    month = int(month) if isinstance(month, str) else month
    st.subheader(" 项目对比分析")
    
    # 项目使用率对比图
    st.plotly_chart(create_project_comparison_chart(all_data, month), use_container_width=True, key="project_comparison_chart")
    
    # 详细指标对比表
    st.subheader(" 详细指标对比")
    comparison_data = create_kpi_display(all_data)
    df_comparison = pd.DataFrame(comparison_data)
    
    # 定义样式函数
    def highlight_usage_rate(val):
        """根据年使用率添加颜色标记"""
        if val >= 100:
            # 超过100%标红色
            return 'background-color: #ffcdd2; color: #d32f2f;'  # 浅红背景，深红文字
        elif val >= 90:
            # 超过90%黄色
            return 'background-color: #fff9c4; color: #f57f17;'  # 浅黄背景，深黄文字
        else:
            return ''
    
    # 应用样式到年使用率列
    styled_df = df_comparison.style.map(highlight_usage_rate, subset=['年使用率(%)'])
    
    st.dataframe(styled_df, use_container_width=True)
    
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
                    label="📥 下载项目汇总表",
                    data=excel_data,
                    file_name="项目汇总表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel_multi",
                    help="下载所有项目的汇总数据表，而不是拼接表"
                )
    
    # 每月费项数据处理
    monthly_fee_df = create_monthly_fee_summary(all_main_dfs)
    if monthly_fee_df is not None:
        # 格式化数据，保留2位小数
        monthly_fee_df['目标成本'] = monthly_fee_df['目标成本'].round(2)
        monthly_fee_df['已发生成本'] = monthly_fee_df['已发生成本'].round(2)
        
        # 绘制柱状图
        st.subheader("每月费项成本柱状图")
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
    
    # 二级费项组合图表
    st.subheader("二级费项整体分析")
    
    # 获取二级费项整体数据
    secondary_fee_data = create_secondary_fee_overall_data(all_main_dfs)
    
    if secondary_fee_data:
        # 创建费项选择下拉框
        fee_names = [item['name'] for item in secondary_fee_data]
        selected_fee = st.selectbox(
            "选择要分析的费项:",
            fee_names,
            index=0,
            key="secondary_fee_selector"
        )
        
        if selected_fee:
            # 创建组合图表
            combined_chart = create_secondary_fee_combined_chart(secondary_fee_data, selected_fee)
            if combined_chart:
                st.plotly_chart(combined_chart, use_container_width=True, key="secondary_fee_combined_chart")
                
                # 添加下载按钮
                # 创建选中费项的数据表格
                selected_fee_data = None
                for item in secondary_fee_data:
                    if item['name'] == selected_fee:
                        selected_fee_data = item
                        break
                
                if selected_fee_data:
                    # 创建数据表格
                    chart_data = []
                    months = [f'{i}月' for i in range(1, 13)]
                    
                    for i, month in enumerate(months):
                        chart_data.append({
                            '月份': month,
                            '单月目标成本(万元)': round(selected_fee_data['target'][i] / 10000, 2),
                            '单月已发生成本(万元)': round(selected_fee_data['actual'][i] / 10000, 2),
                            '月累目标(万元)': round(selected_fee_data['cum_target'][i] / 10000, 2),
                            '月累已发生(万元)': round(selected_fee_data['cum_actual'][i] / 10000, 2)
                        })
                    
                    df_chart = pd.DataFrame(chart_data)
                    csv = df_chart.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label=f"📥 下载{selected_fee}整体数据",
                        data=csv,
                        file_name=f"{selected_fee}_整体分析数据.csv",
                        mime="text/csv",
                        key=f"download_secondary_fee_{selected_fee}"
                    )
            else:
                st.warning(f"无法创建{selected_fee}的图表")
    else:
        st.info("没有找到二级费项数据")
    
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
    
    # 添加人工服务拆分表格展示 - 移到这里
    st.markdown("---")
    st.subheader("人工服务拆分汇总数据")
    
    # 获取所有项目文件
    from pathlib import Path
    data_dir = Path("data")
    all_files = [f.name for f in data_dir.glob("*.xlsx") if f.is_file()]
    
    if all_files:
        # 创建人工服务拆分汇总表
        labor_summary = create_labor_service_summary(all_files, data_dir)
        
        if labor_summary is not None:
            # 显示汇总信息
            st.markdown("#### 人工服务拆分汇总表")
            st.info(f"已生成 {len(all_files)} 个项目的汇总表，所有数据已合并计算")
            
            # 创建人工服务拆分图表
            st.markdown("#### 📊 人工服务拆分图表分析")
            
            # 找到三个汇总费项
            summary_items = {}
            for idx, row in labor_summary.iterrows():
                fee_name = str(row.iloc[0]).strip()  # 第一列是费项名称
                
                # 提取数值数据（1-12月）
                monthly_data = []
                for month in range(1, 13):
                    month_col = f"{month}月"
                    if month_col in labor_summary.columns:
                        value = row[month_col]
                        monthly_data.append(float(value) if pd.notna(value) else 0.0)
                    else:
                        monthly_data.append(0.0)
                
                # 根据费项名称分类
                if "自有人员汇总" in fee_name:
                    summary_items['自有人员汇总'] = {
                        'name': fee_name,
                        'data': monthly_data,
                        'color': '#ff7f0e'  # 橙色
                    }
                elif "专业外包汇总" in fee_name:
                    summary_items['专业外包汇总'] = {
                        'name': fee_name,
                        'data': monthly_data,
                        'color': '#2ca02c'  # 绿色
                    }
                elif "劳务派遣汇总" in fee_name:
                    summary_items['劳务派遣汇总'] = {
                        'name': fee_name,
                        'data': monthly_data,
                        'color': '#1f77b4'  # 蓝色
                    }
            
            if summary_items:
                # 创建组合图表
                fig = go.Figure()
                
                # 准备数据
                months = [f"{i}月" for i in range(1, 13)]
                
                # 为每个汇总费项添加数据
                for item_name, item_data in summary_items.items():
                    color = item_data['color']
                    
                    # 添加已发生金额柱状图
                    fig.add_trace(go.Bar(
                        name=f"{item_data['name']} - 已发生金额",
                        x=months,
                        y=item_data['data'],
                        marker_color=color,
                        opacity=0.8,
                        hovertemplate='<b>%{x}</b><br>' +
                                    f'{item_data["name"]} - 已发生金额<br>' +
                                    '金额: %{y:,.2f}元<br>' +
                                    '<extra></extra>'
                    ))
                    
                    # 添加目标金额折线图
                    fig.add_trace(go.Scatter(
                        name=f"{item_data['name']} - 目标金额",
                        x=months,
                        y=item_data['data'],  # 这里应该使用目标金额数据
                        mode='lines+markers',
                        line=dict(color=color, width=3, dash='dash'),
                        marker=dict(size=8, symbol='diamond'),
                        hovertemplate='<b>%{x}</b><br>' +
                                    f'{item_data["name"]} - 目标金额<br>' +
                                    '金额: %{y:,.2f}元<br>' +
                                    '<extra></extra>'
                    ))
                
                # 更新布局
                fig.update_layout(
                    title={
                        'text': '人工服务拆分汇总图表',
                        'x': 0.5,
                        'xanchor': 'center',
                        'font': {'size': 18, 'color': '#1D1D1F'}
                    },
                    xaxis_title="月份",
                    yaxis_title="金额（元）",
                    barmode='group',
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif"),
                    height=500,
                    showlegend=True,
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )
                
                # 显示图表
                st.plotly_chart(fig, use_container_width=True, key="labor_summary_chart")
                
                # 显示汇总费项列表
                st.markdown("#### 📋 汇总费项列表")
                for item_name, item_data in summary_items.items():
                    total_amount = sum(item_data['data'])
                    st.info(f"**{item_data['name']}**: 年度总金额 {total_amount:,.2f} 元")
            else:
                st.info("未找到指定的汇总费项（自有人员汇总、专业外包汇总、劳务派遣汇总）")
            
            # 设置表格样式
            def style_labor_table(df):
                # 为数值列添加千分位分隔符和颜色
                styled_df = df.copy()
                for col in styled_df.columns:
                    if col != '项目名称' and str(col).isdigit() and 1 <= int(col) <= 12:
                        styled_df[col] = styled_df[col].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "0.00")
                
                return styled_df
            
            styled_labor_df = style_labor_table(labor_summary)
            st.dataframe(styled_labor_df, use_container_width=True, hide_index=True)
            
            # 提供下载按钮
            csv_data = labor_summary.to_csv(index=False, encoding='utf-8-sig')
            file_name = f"人工服务拆分汇总表_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv"
            
            st.download_button(
                label="📥 下载人工服务拆分汇总表",
                data=csv_data,
                file_name=file_name,
                mime="text/csv",
                key="download_labor_service_data"
            )
        else:
            st.info("未找到人工服务拆分数据")
    else:
        st.info("没有可用的项目文件")
    
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
        st.markdown(f"#### {selected_project} 进度指标")
        
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
        st.markdown(f"####  {selected_project} 详细信息")
        
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
                <div class="multi-detail-value">{selected_data['cum_target'][int(''.join(filter(str.isdigit, str(month))))-1]/10000:.2f}</div>
                <div class="multi-detail-unit">万元</div>
            </div>
            """
            st.markdown(month_target_html, unsafe_allow_html=True)
        
        with col3:
            # 月累已发生
            month_actual_html = f"""
            <div class="multi-detail-card">
                <div class="multi-detail-label">月累已发生</div>
                <div class="multi-detail-value">{selected_data['cum_actual'][int(''.join(filter(str.isdigit, str(month))))-1]/10000:.2f}</div>
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
    # 确保month是整数类型
    month = int(month) if isinstance(month, str) else month
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
        
        # 显示详细分析
        render_single_project_analysis(project_name, data, month)
    else:
        # 多项目模式：显示项目对比分析
        render_multi_project_analysis(all_data, all_main_dfs, all_tertiary_dfs, month)
    
    # 添加客户下载表格功能 - 移到最后
    st.markdown("---")
    st.subheader("📥 项目数据导出")
    
    # 创建客户下载表格
    client_table = create_client_download_table(all_main_dfs, all_data)
    
    # 显示表格预览
    st.markdown("#### 项目数据表格预览")
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

def render_exception_ranking(all_data, month):
    """渲染异常数量排行榜"""
    # 确保month是整数类型，处理各种可能的格式
    if isinstance(month, str):
        # 处理"12月"这样的字符串格式
        import re
        month_match = re.search(r'(\d+)', month)
        if month_match:
            month = int(month_match.group(1))
        else:
            st.error(f"无法解析月份格式: {month}")
            return
    else:
        month = int(month) if month is not None else 12
    
    if not all_data:
        st.warning("没有可用的项目数据")
        return
        
    st.markdown("---")
    st.subheader("📊 项目异常数量排行榜")
    
    # 统计每个项目的异常数量
    project_stats = {}
    
    try:
        for project_name, data in all_data.items():
            if not isinstance(data, dict):
                continue
                
            # 统计二级费项异常（主控费项异常）
            main_exceptions = data.get('exceptions', [])
            if not isinstance(main_exceptions, list):
                main_exceptions = []
                
            main_red_count = 0
            main_yellow_count = 0
            for e in main_exceptions:
                if isinstance(e, dict):
                    # 处理异常数据中的月份格式
                    exception_month = e.get('month', 0)
                    if isinstance(exception_month, str):
                        import re
                        month_match = re.search(r'(\d+)', exception_month)
                        if month_match:
                            exception_month = int(month_match.group(1))
                        else:
                            exception_month = 0
                    else:
                        exception_month = int(exception_month) if exception_month is not None else 0
                    
                    if e.get('exception_type') == 'red' and exception_month <= month:
                        main_red_count += 1
                    elif e.get('exception_type') == 'yellow' and exception_month <= month:
                        main_yellow_count += 1
            
            # 统计三级费项异常
            tertiary_exceptions = data.get('tertiary_exceptions', [])
            if not isinstance(tertiary_exceptions, list):
                tertiary_exceptions = []
                
            tertiary_red_count = 0
            tertiary_yellow_count = 0
            for e in tertiary_exceptions:
                if isinstance(e, dict):
                    # 处理异常数据中的月份格式
                    exception_month = e.get('month', 0)
                    if isinstance(exception_month, str):
                        import re
                        month_match = re.search(r'(\d+)', exception_month)
                        if month_match:
                            exception_month = int(month_match.group(1))
                        else:
                            exception_month = 0
                    else:
                        exception_month = int(exception_month) if exception_month is not None else 0
                    
                    if e.get('exception_type') == 'red' and exception_month <= month:
                        tertiary_red_count += 1
                    elif e.get('exception_type') == 'yellow' and exception_month <= month:
                        tertiary_yellow_count += 1
            
            project_stats[project_name] = {
                'main_red': main_red_count,
                'main_yellow': main_yellow_count,
                'main_total': main_red_count + main_yellow_count,
                'tertiary_red': tertiary_red_count,
                'tertiary_yellow': tertiary_yellow_count,
                'tertiary_total': tertiary_red_count + tertiary_yellow_count,
                'total_exceptions': main_red_count + main_yellow_count + tertiary_red_count + tertiary_yellow_count
            }
    except Exception as e:
        st.error(f"处理项目数据时出错: {e}")
        return
    
    # 创建两列布局显示排行榜
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 🔴 二级费项异常排行榜")
        
        # 按二级费项异常总数排序
        main_ranking = sorted(project_stats.items(), key=lambda x: x[1]['main_total'], reverse=True)
        
        # 创建表格数据
        main_table_data = []
        for rank, (project_name, stats) in enumerate(main_ranking, 1):
            main_table_data.append({
                '排名': rank,
                '项目名称': project_name,
                '红色异常': stats['main_red'],
                '黄色异常': stats['main_yellow'],
                '总异常数': stats['main_total']
            })
        
        # 显示表格
        if main_table_data:
            main_df = pd.DataFrame(main_table_data)
            
            # 添加颜色样式
            def highlight_main_ranking(row):
                if row['排名'] == 1 and row['总异常数'] > 0:
                    return ['background-color: #fee2e2; color: #dc2626;'] * len(row)  # 第一名红色
                elif row['排名'] == 2 and row['总异常数'] > 0:
                    return ['background-color: #fef3c7; color: #d97706;'] * len(row)  # 第二名黄色
                elif row['排名'] == 3 and row['总异常数'] > 0:
                    return ['background-color: #fef2dd; color: #ea580c;'] * len(row)  # 第三名橙色
                return [''] * len(row)
            
            styled_main_df = main_df.style.apply(highlight_main_ranking, axis=1)
            st.dataframe(styled_main_df, use_container_width=True, hide_index=True)
        else:
            st.info("暂无二级费项异常数据")
    
    with col2:
        st.markdown("#### 🔵 三级费项异常排行榜")
        
        # 按三级费项异常总数排序
        tertiary_ranking = sorted(project_stats.items(), key=lambda x: x[1]['tertiary_total'], reverse=True)
        
        # 创建表格数据
        tertiary_table_data = []
        for rank, (project_name, stats) in enumerate(tertiary_ranking, 1):
            tertiary_table_data.append({
                '排名': rank,
                '项目名称': project_name,
                '红色异常': stats['tertiary_red'],
                '黄色异常': stats['tertiary_yellow'],
                '总异常数': stats['tertiary_total']
            })
        
        # 显示表格
        if tertiary_table_data:
            tertiary_df = pd.DataFrame(tertiary_table_data)
            
            # 添加颜色样式
            def highlight_tertiary_ranking(row):
                if row['排名'] == 1 and row['总异常数'] > 0:
                    return ['background-color: #fee2e2; color: #dc2626;'] * len(row)  # 第一名红色
                elif row['排名'] == 2 and row['总异常数'] > 0:
                    return ['background-color: #fef3c7; color: #d97706;'] * len(row)  # 第二名黄色
                elif row['排名'] == 3 and row['总异常数'] > 0:
                    return ['background-color: #fef2dd; color: #ea580c;'] * len(row)  # 第三名橙色
                return [''] * len(row)
            
            styled_tertiary_df = tertiary_df.style.apply(highlight_tertiary_ranking, axis=1)
            st.dataframe(styled_tertiary_df, use_container_width=True, hide_index=True)
        else:
            st.info("暂无三级费项异常数据")
    
    # 综合排行榜
    st.markdown("#### 🏆 综合异常排行榜（总异常数）")
    
    # 按总异常数排序
    total_ranking = sorted(project_stats.items(), key=lambda x: x[1]['total_exceptions'], reverse=True)
    
    # 创建综合表格数据
    total_table_data = []
    for rank, (project_name, stats) in enumerate(total_ranking, 1):
        total_table_data.append({
            '排名': rank,
            '项目名称': project_name,
            '二级费项异常': stats['main_total'],
            '三级费项异常': stats['tertiary_total'],
            '总异常数': stats['total_exceptions']
        })
    
    # 显示综合表格
    if total_table_data:
        total_df = pd.DataFrame(total_table_data)
        
        # 添加颜色样式
        def highlight_total_ranking(row):
            if row['排名'] == 1 and row['总异常数'] > 0:
                return ['background-color: #fee2e2; color: #dc2626; font-weight: bold;'] * len(row)  # 第一名红色加粗
            elif row['排名'] == 2 and row['总异常数'] > 0:
                return ['background-color: #fef3c7; color: #d97706; font-weight: bold;'] * len(row)  # 第二名黄色加粗
            elif row['排名'] == 3 and row['总异常数'] > 0:
                return ['background-color: #fef2dd; color: #ea580c; font-weight: bold;'] * len(row)  # 第三名橙色加粗
            return [''] * len(row)
        
        styled_total_df = total_df.style.apply(highlight_total_ranking, axis=1)
        st.dataframe(styled_total_df, use_container_width=True, hide_index=True)
    else:
        st.info("暂无异常数据")
    
    # 下载按钮
    if total_table_data:
        csv_data = total_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 下载异常排行榜",
            data=csv_data,
            file_name=f"项目异常排行榜_{month}月.csv",
            mime="text/csv",
            key="download_exception_ranking"
        )