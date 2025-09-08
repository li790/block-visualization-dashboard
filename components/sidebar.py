import streamlit as st
from pathlib import Path
import os
import tempfile
from utils.data_processor import get_excel_files, extract_table_from_excel
from utils.cache_manager import get_cache_manager

# 自定义CSS样式美化文件上传区域
def inject_custom_css():
    st.markdown("""
    <style>
    .file-upload-container {
        border: none;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        text-align: center;
    }
    .file-upload-header {
        color: #4CAF50;
        font-weight: bold;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

def render_sidebar(data_dir):
    """渲染侧边栏"""
    # 注入自定义CSS
    inject_custom_css()
    
    # 初始化时间戳
    import time
    if 'current_time' not in st.session_state:
        st.session_state.current_time = time.time()
    current_time = time.time()
    

    
    with st.sidebar:
        # 只保留上传、保存、选择分析等功能，不显示文件管理标题和文件数
        # 文件上传通道 - 需要提取表格的文件
        st.markdown('<div class="file-upload-container">', unsafe_allow_html=True)
        st.markdown('<div class="file-upload-header">需提取表格的文件上传</div>', unsafe_allow_html=True)
        extracted_files = st.file_uploader(
            "上传需要提取表格的Excel文件",
            type=['xlsx', 'xls'],
            help="支持.xlsx和.xls格式，可同时选择多个文件，系统将自动提取特定工作表的前39行",
            accept_multiple_files=True,
            key="extracted_files"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 合并上传的文件
        all_uploaded_files = []
        if extracted_files:
            all_uploaded_files.extend(extracted_files)
        
        # 取消上传后文件清单与提示的显示，避免干扰
        
        st.divider()
        
        # 主要费项表格选择
        st.markdown("**📊 主要费项表格选择**")
        
        # 创建两列布局：按钮和帮助图标
        col1, col2 = st.columns([3, 1])
        
        with col1:
            include_self_owned_labor = st.toggle(
                "包含自有人工成本",
                value=False,
                help="切换主要费项表格的数据源"
            )
        
        
        # 显示当前选择的表格类型
        if include_self_owned_labor:
            st.info("📋 当前选择：包含自有人工成本的主要费项表格")
        else:
            st.info("📋 当前选择：不包含自有人工成本的主要费项表格")
        
        st.divider()
        
        # 缓存状态显示
        cache_manager = get_cache_manager()
        cache_stats = cache_manager.get_cache_stats()
        if 'error' not in cache_stats and cache_stats['cache_count'] > 0:
            st.info(f"⚡ 缓存状态: {cache_stats['cache_count']} 个文件已缓存 ({cache_stats['total_size_mb']}MB)")
        
        # 选择数据来源与文件：
        # - 无上传文件：展示示例数据(data目录)供选择
        # - 有上传文件：仅展示本次上传项目供选择，不展示示例数据
        selected_files = []
        # 统一初始化：来自 session_state 的上传项目选择
        selected_uploaded_projects = st.session_state.get('selected_uploaded_projects', [])
        existing_files = get_excel_files(data_dir)
        has_uploaded = bool(extracted_files)
        if not has_uploaded:
            if existing_files:
                file_names = [f.name for f in existing_files]
                if 'selected_files' not in st.session_state:
                    st.session_state['selected_files'] = []
                selected_files = st.multiselect(
                    "选择要分析的文件:",
                    file_names,
                    key='selected_files'
                )
                def select_all():
                    st.session_state['selected_files'] = file_names
                def deselect_all():
                    st.session_state['selected_files'] = []
                if len(st.session_state['selected_files']) < len(file_names):
                    st.button('全选所有文件', on_click=select_all)
                else:
                    st.button('全不选', on_click=deselect_all)
        
        # 新增：用于主流程分析的DataFrame收集
        uploaded_main_dfs = {}
        uploaded_tertiary_dfs = {}
        rendered_uploaded_selector = False
        # 处理需提取表格的文件
        if extracted_files:
            processed_count = 0
            for file in extracted_files:
                try:
                    project_name = Path(file.name).stem.strip()
                    # 只在内存中处理，不再保存到output目录
                    main_df, tertiary_df = extract_table_from_excel(file, include_self_owned_labor)
                    if main_df is not None and tertiary_df is not None:
                        # 新增：收集DataFrame
                        uploaded_main_dfs[project_name] = main_df
                        uploaded_tertiary_dfs[project_name] = tertiary_df
                        processed_count += 1
                    else:
                        st.error(f"无法从文件 {file.name} 中提取有效数据")
                except Exception as e:
                    st.error(f"处理文件 {file.name} 时出错: {e}")
            if processed_count > 0:
                # 在成功提示上方提供上传项目选择与全选按钮
                uploaded_names = list(uploaded_main_dfs.keys())
                if 'selected_uploaded_projects' not in st.session_state:
                    st.session_state['selected_uploaded_projects'] = []
                st.markdown("**选择要分析的上传项目:**")
                def select_all_uploaded():
                    st.session_state['selected_uploaded_projects'] = uploaded_names
                def deselect_all_uploaded():
                    st.session_state['selected_uploaded_projects'] = []
                st.multiselect(
                    label="",
                    options=uploaded_names,
                    key='selected_uploaded_projects'
                )
                col_u1, col_u2 = st.columns([1,1])
                with col_u1:
                    st.button('全选上传项目', key='select_all_uploaded_top', on_click=select_all_uploaded)
                with col_u2:
                    st.button('全不选', key='deselect_all_uploaded_top', on_click=deselect_all_uploaded)
                rendered_uploaded_selector = True
                # 同步局部变量，确保返回值正确
                selected_uploaded_projects = st.session_state.get('selected_uploaded_projects', [])

        # 当存在上传项目时，提供上传项目的选择框（必须选择后才参与分析）
        if uploaded_main_dfs and not rendered_uploaded_selector:
            uploaded_names = list(uploaded_main_dfs.keys())
            if 'selected_uploaded_projects' not in st.session_state:
                st.session_state['selected_uploaded_projects'] = []
            st.markdown("**选择要分析的上传项目:**")
            def select_all_uploaded():
                st.session_state['selected_uploaded_projects'] = uploaded_names
            def deselect_all_uploaded():
                st.session_state['selected_uploaded_projects'] = []
            selected_uploaded_projects = st.multiselect(
                label="",
                options=uploaded_names,
                key='selected_uploaded_projects'
            )
            col_u1, col_u2 = st.columns([1,1])
            with col_u1:
                st.button('全选上传项目', key='select_all_uploaded_side', on_click=select_all_uploaded)
            with col_u2:
                st.button('全不选', key='deselect_all_uploaded_side', on_click=deselect_all_uploaded)
        else:
            # 若没有上传数据或已在上方渲染，确保局部变量与 session 同步
            selected_uploaded_projects = st.session_state.get('selected_uploaded_projects', [])
        # 返回新增内容（增加 selected_uploaded_projects）
        return all_uploaded_files, extracted_files, selected_files, uploaded_main_dfs, uploaded_tertiary_dfs, include_self_owned_labor, selected_uploaded_projects