import streamlit as st
from pathlib import Path
from components.sidebar import render_sidebar
from components.dashboard import render_dashboard
from utils.data_processor import load_and_process_files, create_summary_excel, extract_table_from_excel, create_data_summary_tables
from utils.cache_manager import get_cache_manager
from components.cache_indicator import start_performance_timer, end_performance_timer, show_cache_benefit_message

# 页面配置
st.set_page_config(
    page_title="运营成本管理看板",
    page_icon="📊",  # 使用emoji图标替代本地图片
    layout="wide",
    initial_sidebar_state="expanded"
)

# 数据目录
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

# 注意：已移除output目录，现在使用内存缓存

def main():
    st.title("运营成本管理看板")
    
    # 注意：已移除文件管理功能，现在专注于数据分析

    # ====== 缓存管理功能 ======
    cache_manager = get_cache_manager()
    
    # 缓存管理按钮
    if st.button("⚡ 缓存管理", type="secondary"):
        st.session_state.show_cache_manager = True
    
    if st.session_state.get('show_cache_manager', False):
        st.subheader("⚡ 缓存管理")
        
        # 获取缓存统计信息
        cache_stats = cache_manager.get_cache_stats()
        
        if 'error' not in cache_stats:
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("缓存文件数", cache_stats['cache_count'])
            with col2:
                st.metric("缓存大小(MB)", cache_stats['total_size_mb'])
            with col3:
                st.metric("元数据数", cache_stats['metadata_count'])
            with col4:
                if st.button("🧹 清理过期缓存"):
                    cache_manager.cleanup_expired_cache()
                    st.rerun()
        
        # 缓存操作按钮
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🗑️ 清除所有缓存"):
                cache_manager.clear_all_cache()
                st.rerun()
        with col2:
            if st.button("🔄 刷新缓存统计"):
                st.rerun()
        with col3:
            if st.button("❌ 关闭缓存管理"):
                st.session_state.show_cache_manager = False
                st.rerun()
        
        st.markdown("---")
        st.info("💡 **缓存说明:** 系统使用内存缓存已处理的数据，相同文件再次选择时将直接从内存加载，大幅提升加载速度。缓存有效期为24小时，会话结束后自动清理。")
    
    # ====== 缓存管理功能 END ======

    # 渲染侧边栏并获取文件
    all_uploaded_files, extracted_files, selected_files, uploaded_main_dfs, uploaded_tertiary_dfs, include_self_owned_labor, selected_uploaded_projects = render_sidebar(DATA_DIR)
    # 与 session_state 同步，防止选择值未正确传递
    selected_uploaded_projects = st.session_state.get('selected_uploaded_projects', selected_uploaded_projects or [])
    
    # 月份选择
    month = st.slider("选择月份:", min_value=1, max_value=12, value=5)  # 默认5月

    # 动态收集所有项目数据
    all_main_dfs = {}
    all_tertiary_dfs = {}
    all_data = {}
    
    # 1. 添加“已选择”的上传项目（未选择则不展示）
    if uploaded_main_dfs:
        if selected_uploaded_projects:
            for project_name in selected_uploaded_projects:
                if project_name in uploaded_main_dfs:
                    all_main_dfs[project_name] = uploaded_main_dfs[project_name]
                if project_name in uploaded_tertiary_dfs:
                    all_tertiary_dfs[project_name] = uploaded_tertiary_dfs[project_name]
        else:
            # 未选择则不加入，提示在下方统一给出
            pass
    
    # 2. 添加“已选择”的示例数据项目（需明确选择）
    if selected_files:
        st.write(f"**正在处理选中的 {len(selected_files)} 个文件...**")
        # 开始性能计时
        start_performance_timer()
        
        for filename in selected_files:
            try:
                # 处理原始Excel文件
                file_path = DATA_DIR / filename
                main_df, tertiary_df = extract_table_from_excel(file_path, include_self_owned_labor)
                if main_df is not None and tertiary_df is not None:
                    project_name = filename.replace('.xlsx', '')
                    all_main_dfs[project_name] = main_df
                    all_tertiary_dfs[project_name] = tertiary_df

                else:
                    st.warning(f"无法从文件 {filename} 中提取有效数据")
            except Exception as e:
                st.error(f"处理文件 {filename} 时出错: {e}")
        
        # 结束性能计时并显示结果
        if selected_files:
            elapsed_time = end_performance_timer()
            if elapsed_time > 0:
                st.info(f"⏱️ 文件处理完成，耗时: {elapsed_time:.2f}秒")
    
    # 3. 分类型与模式分析
    if all_main_dfs:
        from utils.data_processor import process_excel_data, process_tertiary_fee_data

        # 解析类型前缀与项目显示名
        def split_type_and_display(project: str):
            parts = project.split('-', 1)
            if len(parts) == 2 and parts[0]:
                return parts[0], parts[1]
            return project, project

        type_to_projects = {}
        project_to_display = {}
        for project_name in list(all_main_dfs.keys()):
            t, display = split_type_and_display(project_name)
            type_to_projects.setdefault(t, []).append(project_name)
            project_to_display[project_name] = display

        all_types = list(type_to_projects.keys())

        export_file_prefix = None

        # 单一类型自动/选择
        if len(all_types) == 1:
            selected_mode = '单一类型分析'
            selected_type = all_types[0]
        else:
            selected_mode = st.radio(
                '选择分析模式',
                ['单一类型分析', '汇总分析', '区块分析'],
                horizontal=True,
                index=0
            )
            selected_type = None
            if selected_mode == '单一类型分析':
                selected_type = st.selectbox('选择要分析的类型', all_types)

        # 根据模式构建要分析的项目集与导出前缀
        if selected_mode == '单一类型分析':
            target_projects = type_to_projects[selected_type]
            # 过滤并改名（去掉类型前缀）
            filtered_main = {project_to_display[p]: all_main_dfs[p] for p in target_projects}
            filtered_ter = {project_to_display[p]: all_tertiary_dfs.get(p) for p in target_projects if p in all_tertiary_dfs}

            # 执行分析
            for project_name, main_df in filtered_main.items():
                data = process_excel_data(main_df, month, project_name, include_self_owned_labor)
                if data:
                    tertiary_df = filtered_ter.get(project_name)
                    if tertiary_df is not None:
                        tertiary_result = process_tertiary_fee_data(tertiary_df, month, project_name, include_self_owned_labor)
                        data['tertiary_fee_items'] = tertiary_result['tertiary_fee_items']
                        data['tertiary_exceptions'] = tertiary_result['exceptions']
                    all_data[project_name] = data

            export_file_prefix = f"{selected_type}汇总表格"
            total_projects = len(all_data)
            st.success(f"按类型“{selected_type}”分析：共 {total_projects} 个项目")

        elif selected_mode == '汇总分析':
            # 将所有项目视为同一类，同时去掉显示名的类型前缀
            renamed_main = {project_to_display[p]: df for p, df in all_main_dfs.items()}
            renamed_ter = {project_to_display[p]: df for p, df in all_tertiary_dfs.items() if p in all_tertiary_dfs}

            for project_name, main_df in renamed_main.items():
                data = process_excel_data(main_df, month, project_name, include_self_owned_labor)
                if data:
                    tertiary_df = renamed_ter.get(project_name)
                    if tertiary_df is not None:
                        tertiary_result = process_tertiary_fee_data(tertiary_df, month, project_name, include_self_owned_labor)
                        data['tertiary_fee_items'] = tertiary_result['tertiary_fee_items']
                        data['tertiary_exceptions'] = tertiary_result['exceptions']
                    all_data[project_name] = data

            export_file_prefix = "汇总表格"
            st.success(f"汇总分析：共 {len(all_data)} 个项目（已按同一类型处理）")

        else:  # 区块分析
            # 将每个类型的汇总表视为一个“虚拟项目”，统一进行一次多项目分析
            st.info("区块分析：按类型生成汇总单表，并将各类型汇总作为多项目统一分析与导出。")

            combined_main = {}
            combined_ter = {}
            combined_data = {}

            for t in all_types:
                # 取该类型下的原始项目数据字典
                original_main_subset = {project_to_display[p]: all_main_dfs[p] for p in type_to_projects[t]}
                original_ter_subset = {project_to_display[p]: all_tertiary_dfs.get(p) for p in type_to_projects[t] if p in all_tertiary_dfs}

                # 使用“数据汇总表格”逻辑进行合并
                summary_tables = create_data_summary_tables(original_main_subset, original_ter_subset, include_self_owned_labor)

                # 选择主要费项表（根据是否包含自有人工）
                main_sheet_4 = '4主要费项费项月累成本使用情况'
                main_sheet_4_1 = '4-1主要费项费项月累成本使用情况'
                tertiary_sheet = '三级费项月累表格'
                merged_main_df = summary_tables.get(main_sheet_4 if include_self_owned_labor else main_sheet_4_1)
                merged_ter_df = summary_tables.get(tertiary_sheet)

                virtual_project_name = f"{t}汇总"
                if merged_main_df is None:
                    st.warning(f"类型“{t}”未生成主要费项汇总表，已跳过")
                    continue

                combined_main[virtual_project_name] = merged_main_df
                if merged_ter_df is not None:
                    combined_ter[virtual_project_name] = merged_ter_df

                data = process_excel_data(merged_main_df, month, virtual_project_name, include_self_owned_labor)
                if data:
                    if merged_ter_df is not None:
                        tertiary_result = process_tertiary_fee_data(merged_ter_df, month, virtual_project_name, include_self_owned_labor)
                        data['tertiary_fee_items'] = tertiary_result['tertiary_fee_items']
                        data['tertiary_exceptions'] = tertiary_result['exceptions']
                    combined_data[virtual_project_name] = data

            if not combined_data:
                st.warning("区块分析没有可用的数据")
                return

            # 统一一次渲染，触发多项目分析视图
            render_dashboard(combined_data, combined_main, combined_ter, month, include_self_owned_labor, export_file_prefix="区块汇总表格")
            return

        # 显示分析结果摘要
        if all_data:
            total_projects = len(all_data)
            uploaded_count = len(selected_uploaded_projects) if selected_uploaded_projects else 0
            selected_count = len(selected_files) if selected_files else 0

            if uploaded_count > 0 and selected_count > 0:
                st.write(f"**分析结果:** 成功处理 {total_projects} 个项目 (上传 {uploaded_count} 个 + 选择 {selected_count} 个)")
            elif uploaded_count > 0:
                st.write(f"**分析结果:** 成功处理 {total_projects} 个刚上传的项目")
            elif selected_count > 0:
                st.write(f"**分析结果:** 成功处理 {total_projects} 个选中的项目")
            else:
                st.write(f"**分析结果:** 成功处理 {total_projects} 个项目")

            if total_projects > 1:
                st.info(f"📊 已生成 {total_projects} 个项目的汇总表，所有数据已合并计算")
        else:
            st.warning("未得到可分析的数据")
    else:
        if uploaded_main_dfs and not selected_uploaded_projects:
            st.warning("没有可分析的文件：已上传但未选择上传项目，请先在左侧选择要分析的项目")
        else:
            st.warning("没有可分析的文件，请先上传并选择项目，或在未上传时选择示例数据")
    
    # 渲染仪表盘
    render_dashboard(all_data, all_main_dfs, all_tertiary_dfs, month, include_self_owned_labor, export_file_prefix=locals().get('export_file_prefix'))

if __name__ == "__main__":
    main()