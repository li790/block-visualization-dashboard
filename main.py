import streamlit as st
from pathlib import Path
from components.sidebar import render_sidebar
from components.dashboard import render_dashboard
from utils.data_processor import load_and_process_files, create_summary_excel, extract_table_from_excel, get_excel_files
from utils.cache_manager import get_cache_manager
from components.cache_indicator import start_performance_timer, end_performance_timer, show_cache_benefit_message

# 页面配置
st.set_page_config(
    page_title="运营成本管理看板",
    page_icon="icon icon1",  # 使用用户指定的本地图片路径
    layout="wide",
    initial_sidebar_state="expanded"
)

# 数据目录
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

# 输出目录
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

def main():
    st.title("运营成本管理看板")
    
    # ====== 文件管理主页面实现 ======
    existing_files = get_excel_files(DATA_DIR)
    if 'main_show_file_manager' not in st.session_state:
        st.session_state.main_show_file_manager = False
    # 初始化 main_selected_files，防止未初始化报错
    if 'main_selected_files' not in st.session_state:
        st.session_state.main_selected_files = []
    if st.button("📋 管理文件", type="primary"):
        st.session_state.main_show_file_manager = True
    if st.session_state.main_show_file_manager:
        st.subheader("📂 文件管理")
        st.write(f"**data目录中有 {len(existing_files)} 个文件:**")
        # 按钮排布优化：一行四列
        col_a, col_b, col_c, col_d = st.columns([1,1,1,1])
        with col_a:
            if st.button("✅ 全选", key="main_select_all"):
                st.session_state.main_selected_files = [f.name for f in existing_files]
        with col_b:
            if st.button("❌ 取消全选", key="main_deselect_all"):
                st.session_state.main_selected_files = []
        with col_c:
            if st.button("🗑️ 批量删除", key="main_batch_delete"):
                selected_files = st.session_state.main_selected_files
                if selected_files:
                    deleted_count = 0
                    for filename in selected_files:
                        file_path = DATA_DIR / filename
                        if file_path.exists():
                            try:
                                file_path.unlink()
                                deleted_count += 1
                            except Exception as e:
                                st.error(f"删除 {filename} 失败: {e}")
                    if deleted_count > 0:
                        st.session_state.main_selected_files = []
                        st.rerun()
                else:
                    st.warning("请先选择要删除的文件")
        with col_d:
            if st.button("❌ 关闭管理", key="main_close_manager"):
                st.session_state.main_show_file_manager = False
                st.rerun()
        st.markdown("---")
        # 文件列表 - 固定高度滚动条，展示所有文件
        st.write("**文件列表:**")
        st.markdown("""
            <style>
            .main-file-list-scroll {
                max-height: 250px;
                overflow-y: auto;
                border: 1px solid #ddd;
                border-radius: 8px;
                background: #fff;
                padding: 8px;
                margin: 8px 0;
            }
            </style>
        """, unsafe_allow_html=True)
        st.markdown('<div class="main-file-list-scroll">', unsafe_allow_html=True)
        for i, file in enumerate(existing_files):
            col1, col2, col3, col4 = st.columns([1, 3, 1, 1])
            with col1:
                is_selected = file.name in st.session_state.main_selected_files
                if st.checkbox("", value=is_selected, key=f"main_check_{i}"):
                    if file.name not in st.session_state.main_selected_files:
                        st.session_state.main_selected_files.append(file.name)
                else:
                    if file.name in st.session_state.main_selected_files:
                        st.session_state.main_selected_files.remove(file.name)
            with col2:
                st.write(f"{i+1}. {file.name}")
            with col3:
                if st.button("删除", key=f"main_del_{i}"):
                    try:
                        file.unlink()
                        if file.name in st.session_state.main_selected_files:
                            st.session_state.main_selected_files.remove(file.name)
                        st.rerun()
                    except Exception as e:
                        st.error(f"删除失败: {e}")
            with col4:
                if st.button("重命名", key=f"main_rename_{i}"):
                    st.session_state.main_renaming_file = file.name
                    st.session_state.main_renaming_index = i
        st.markdown('</div>', unsafe_allow_html=True)
        # 不再显示“还有X个文件未显示”提示
        # 重命名弹窗
        if st.session_state.get('main_renaming_file'):
            st.write(f"**重命名文件: {st.session_state.main_renaming_file}**")
            new_name = st.text_input("新文件名 (保留.xlsx扩展名):", value=st.session_state.main_renaming_file, key="main_new_filename")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("确认重命名", key="main_confirm_rename"):
                    try:
                        old_path = DATA_DIR / st.session_state.main_renaming_file
                        new_path = DATA_DIR / new_name
                        old_path.rename(new_path)
                        if st.session_state.main_renaming_file in st.session_state.main_selected_files:
                            st.session_state.main_selected_files.remove(st.session_state.main_renaming_file)
                            st.session_state.main_selected_files.append(new_name)
                        st.session_state.main_renaming_file = None
                        st.rerun()
                    except Exception as e:
                        st.error(f"重命名失败: {e}")
            with col2:
                if st.button("取消", key="main_cancel_rename"):
                    st.session_state.main_renaming_file = None
                    st.rerun()
        st.markdown("---")
    # ====== 文件管理主页面实现 END ======

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
        st.info("💡 **缓存说明:** 系统会自动缓存已处理的数据，相同文件再次选择时将直接从缓存加载，大幅提升加载速度。缓存有效期为24小时。")
    
    # ====== 缓存管理功能 END ======

    # 渲染侧边栏并获取文件
    all_uploaded_files, extracted_files, selected_files, uploaded_main_dfs, uploaded_tertiary_dfs, include_self_owned_labor = render_sidebar(DATA_DIR)
    
    # 月份选择
    month = st.slider("选择月份:", min_value=1, max_value=12, value=5)  # 默认5月

    # 动态收集所有项目数据
    all_main_dfs = {}
    all_tertiary_dfs = {}
    all_data = {}
    
    # 1. 添加上传的所有项目
    if uploaded_main_dfs:
        all_main_dfs.update(uploaded_main_dfs)
        all_tertiary_dfs.update(uploaded_tertiary_dfs)
    
    # 2. 添加选择的所有项目
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
    
    # 3. 统一分析所有项目
    if all_main_dfs:
        from utils.data_processor import process_excel_data, process_tertiary_fee_data
        for project_name, main_df in all_main_dfs.items():
            data = process_excel_data(main_df, month, project_name, include_self_owned_labor)
            if data:
                tertiary_df = all_tertiary_dfs.get(project_name)
                if tertiary_df is not None:
                    tertiary_result = process_tertiary_fee_data(tertiary_df, month, project_name, include_self_owned_labor)
                    data['tertiary_fee_items'] = tertiary_result['tertiary_fee_items']
                    data['tertiary_exceptions'] = tertiary_result['exceptions']
                all_data[project_name] = data
        
        # 显示分析结果
        total_projects = len(all_data)
        uploaded_count = len(uploaded_main_dfs) if uploaded_main_dfs else 0
        selected_count = len(selected_files) if selected_files else 0
        
        if uploaded_count > 0 and selected_count > 0:
            st.write(f"**分析结果:** 成功处理 {total_projects} 个项目 (上传 {uploaded_count} 个 + 选择 {selected_count} 个)")
        elif uploaded_count > 0:
            st.write(f"**分析结果:** 成功处理 {total_projects} 个刚上传的项目")
        elif selected_count > 0:
            st.write(f"**分析结果:** 成功处理 {total_projects} 个选中的项目")
        else:
            st.write(f"**分析结果:** 成功处理 {total_projects} 个项目")
        
        # 静默显示缓存优势信息
        
        # 如果有多项目，显示汇总表信息
        if total_projects > 1:
            st.info(f"📊 已生成 {total_projects} 个项目的汇总表，所有数据已合并计算")
    else:
        st.warning("没有可分析的文件")
    
    # 渲染仪表盘
    render_dashboard(all_data, all_main_dfs, all_tertiary_dfs, month, include_self_owned_labor)

if __name__ == "__main__":
    main()