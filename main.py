import streamlit as st
from pathlib import Path
from components.sidebar import render_sidebar
from components.dashboard import render_dashboard
from utils.data_processor import load_and_process_files, create_summary_excel, extract_table_from_excel, create_data_summary_tables
from utils.cache_manager import get_cache_manager
from components.cache_indicator import start_performance_timer, end_performance_timer, show_cache_benefit_message

# 定义SVG图标
CHART_ICON = """
<svg t="1758768181220" class="icon" viewBox="0 0 1024 1024" version="1.1" xmlns="http://www.w3.org/2000/svg" p-id="1621" width="20" height="20" style="display: inline-block; vertical-align: middle; margin-right: 8px;">
    <path d="M532.15763 144.004741a75.851852 75.851852 0 0 1 75.851851 75.851852v440.301037a75.851852 75.851852 0 0 1-75.851851 75.851851H491.861333a75.851852 75.851852 0 0 1-75.851852-75.851851V219.856593a75.851852 75.851852 0 0 1 75.851852-75.851852h40.296297z m256 111.995259a75.851852 75.851852 0 0 1 75.851851 75.851852V660.176593a75.851852 75.851852 0 0 1-75.851851 75.851851H747.861333a75.851852 75.851852 0 0 1-75.851852-75.851851V331.851852a75.851852 75.851852 0 0 1 75.851852-75.851852h40.296297z m-512 112.014222a75.851852 75.851852 0 0 1 75.851851 75.851852v216.291556a75.851852 75.851852 0 0 1-75.851851 75.851851H235.861333a75.851852 75.851852 0 0 1-75.851852-75.851851V443.847111a75.851852 75.851852 0 0 1 75.851852-75.851852h40.296297z" fill="#279CFF" p-id="1622"></path>
    <path d="M160.009481 816.014222m32.009482 0l639.981037 0q32.009481 0 32.009481 32.009482l0-0.018963q0 32.009481-32.009481 32.009481l-639.981037 0q-32.009481 0-32.009482-32.009481l0 0.018963q0-32.009481 32.009482-32.009482Z" fill="#279CFF" fill-opacity=".5" p-id="1623"></path>
</svg>
"""

# 页面配置
st.set_page_config(
    page_title="运营成本管理看板",
    page_icon=CHART_ICON,  # 使用自定义SVG图标
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
    
    # 月份选择：延后到数据加载完毕后再显示（默认用于数据处理为6月）
    month = 6

    # 动态收集所有项目数据
    all_main_dfs = {}
    all_tertiary_dfs = {}
    all_data = {}
    
    # 处理选中的data目录中的文件（仅在点击开始分析后）
    if selected_files and st.session_state.get('start_analysis', False):
        st.write(f"**正在处理选中的 {len(selected_files)} 个文件...**")
        # 开始性能计时
        start_performance_timer()
        
        # 并行读取与解析，提升首次加载速度
        from concurrent.futures import ThreadPoolExecutor, as_completed

        def _parse_one(file_name: str):
            file_path = DATA_DIR / file_name
            try:
                main_df, tertiary_df = extract_table_from_excel(file_path, include_self_owned_labor)
                return (file_name, main_df, tertiary_df, None)
            except Exception as e:
                return (file_name, None, None, str(e))

        max_workers = min(8, len(selected_files)) if selected_files else 1
        if max_workers < 1:
            max_workers = 1

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_parse_one, fn) for fn in selected_files]
            for fut in as_completed(futures):
                filename = None
                try:
                    filename, main_df, tertiary_df, err = fut.result()
                except Exception as e:
                    err = str(e)
                if err:
                    st.warning(f"处理文件 {filename} 时出错: {err}")
                    continue
                if main_df is not None and tertiary_df is not None:
                    project_name = filename.replace('.xlsx', '')
                    all_main_dfs[project_name] = main_df
                    all_tertiary_dfs[project_name] = tertiary_df
                else:
                    st.warning(f"无法从文件 {filename} 中提取有效数据")
        
        # 结束性能计时并显示结果
        if selected_files:
            elapsed_time = end_performance_timer()
            if elapsed_time > 0:
                st.info(f"⏱️ 文件处理完成，耗时: {elapsed_time:.2f}秒")
        
        # 将处理后的数据保存到session state中，避免重复处理
        st.session_state['processed_main_dfs'] = all_main_dfs.copy()
        st.session_state['processed_tertiary_dfs'] = all_tertiary_dfs.copy()
        st.session_state['processed_include_self_owned_labor'] = include_self_owned_labor
        
        # 后台预计算各项目1..12月分析结果，加速月份切换
        try:
            from concurrent.futures import ThreadPoolExecutor
            from utils.data_processor import precompute_project_all_months
            with ThreadPoolExecutor(max_workers=min(6, len(all_main_dfs))) as ex:
                for project_name, main_df in all_main_dfs.items():
                    ex.submit(precompute_project_all_months, main_df, project_name, include_self_owned_labor)
        except Exception:
            pass

        # 重置分析标志
        st.session_state['start_analysis'] = False
    
    # 如果已经有处理过的数据，直接使用
    elif st.session_state.get('processed_main_dfs') and st.session_state.get('processed_include_self_owned_labor') == include_self_owned_labor:
        all_main_dfs = st.session_state['processed_main_dfs']
        all_tertiary_dfs = st.session_state['processed_tertiary_dfs']
        st.info("📊 使用已处理的数据进行分析")
    
    # 如果自有人工成本设置发生变化，但有已选择的文件，重新处理这些文件
    elif (st.session_state.get('processed_include_self_owned_labor') is not None and 
          st.session_state.get('processed_include_self_owned_labor') != include_self_owned_labor and
          selected_files):
        st.info("🔄 检测到自有人工成本设置变化，正在重新处理已选择的文件...")
        
        # 开始性能计时
        start_performance_timer()
        
        # 重新处理已选择的文件
        from concurrent.futures import ThreadPoolExecutor, as_completed
        
        def _parse_one(file_name: str):
            file_path = DATA_DIR / file_name
            try:
                main_df, tertiary_df = extract_table_from_excel(file_path, include_self_owned_labor)
                return (file_name, main_df, tertiary_df, None)
            except Exception as e:
                return (file_name, None, None, str(e))

        max_workers = min(8, len(selected_files)) if selected_files else 1
        if max_workers < 1:
            max_workers = 1

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_parse_one, fn) for fn in selected_files]
            for fut in as_completed(futures):
                filename = None
                try:
                    filename, main_df, tertiary_df, err = fut.result()
                except Exception as e:
                    err = str(e)
                if err:
                    st.warning(f"处理文件 {filename} 时出错: {err}")
                    continue
                if main_df is not None and tertiary_df is not None:
                    project_name = filename.replace('.xlsx', '')
                    all_main_dfs[project_name] = main_df
                    all_tertiary_dfs[project_name] = tertiary_df
                else:
                    st.warning(f"无法从文件 {filename} 中提取有效数据")
        
        # 结束性能计时并显示结果
        if selected_files:
            elapsed_time = end_performance_timer()
            if elapsed_time > 0:
                st.info(f"⏱️ 文件重新处理完成，耗时: {elapsed_time:.2f}秒")
        
        # 更新处理后的数据
        st.session_state['processed_main_dfs'] = all_main_dfs.copy()
        st.session_state['processed_tertiary_dfs'] = all_tertiary_dfs.copy()
        st.session_state['processed_include_self_owned_labor'] = include_self_owned_labor
        
        # 后台预计算各项目1..12月分析结果，加速月份切换
        try:
            from concurrent.futures import ThreadPoolExecutor
            from utils.data_processor import precompute_project_all_months
            with ThreadPoolExecutor(max_workers=min(6, len(all_main_dfs))) as ex:
                for project_name, main_df in all_main_dfs.items():
                    ex.submit(precompute_project_all_months, main_df, project_name, include_self_owned_labor)
        except Exception:
            pass
    
    # 如果文件选择或设置发生变化，清除缓存
    else:
        if 'processed_main_dfs' in st.session_state:
            del st.session_state['processed_main_dfs']
            del st.session_state['processed_tertiary_dfs']
            del st.session_state['processed_include_self_owned_labor']
            st.info("🔄 检测到设置变化，已清除缓存数据")
    
    # 3. 分类型与模式分析
    if all_main_dfs:
        from utils.data_processor import process_excel_data, process_tertiary_fee_data
        from utils.file_utils import split_type_and_display

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

            # 将渲染用的数据源限制为所选类型，避免“二级费项整体分析”统计到所有项目
            all_main_dfs = filtered_main
            all_tertiary_dfs = filtered_ter

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

                # 使用"数据汇总表格"逻辑进行合并（支持缓存）
                with st.spinner(f"正在生成 {t} 类型汇总表..."):
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
            selected_count = len(selected_files) if selected_files else 0

            st.write(f"**分析结果:** 成功处理 {total_projects} 个选中的项目")

            if total_projects > 1:
                st.info(f"📊 已生成 {total_projects} 个项目的汇总表，所有数据已合并计算")
        else:
            st.warning("未得到可分析的数据")
    else:
        if not selected_files:
            st.warning("没有可分析的文件，请先在左侧选择要分析的文件")
        else:
            st.warning("没有可分析的文件，请先选择data文件夹中的Excel文件")
    
    # 渲染仪表盘：仅在有数据时显示月份选择器和仪表盘
    if all_main_dfs:
        # 数据已就绪，再显示月份选择器
        month = st.slider("选择月份:", min_value=1, max_value=12, value=6)
        render_dashboard(all_data, all_main_dfs, all_tertiary_dfs, month, include_self_owned_labor, export_file_prefix=locals().get('export_file_prefix'))

if __name__ == "__main__":
    main()