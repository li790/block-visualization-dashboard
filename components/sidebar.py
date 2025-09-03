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
        
        if all_uploaded_files:
            st.write(f"**已选择 {len(all_uploaded_files)} 个文件:**")
            for i, file in enumerate(all_uploaded_files):
                file_type = "(需提取表格)" if file in extracted_files else "(普通文件)"
                st.write(f"{i+1}. {file.name} {file_type}")
            
            # 注意：已移除文件保存功能，文件只在内存中处理，不会保存到data目录
            st.info("💡 文件将在内存中处理，不会保存到磁盘，确保数据安全")
            
            # 移除自动落盘逻辑和保存选项
            # 文件只在内存中处理，避免数据污染和多用户冲突
        
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
        
        # 选择现有文件
        existing_files = get_excel_files(data_dir)
        if existing_files:
            file_names = [f.name for f in existing_files]
            # 初始化session_state，避免default和session_state冲突
            if 'selected_files' not in st.session_state:
                st.session_state['selected_files'] = [file_names[0]] if file_names else []
            # 多选框（不再传default参数）
            selected_files = st.multiselect(
                "选择要分析的文件:",
                file_names,
                key='selected_files'
            )
            # 全选/全不选按钮（用回调）
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
        # 处理需提取表格的文件
        if extracted_files:
            st.write("\n**正在处理需提取表格的文件...**")
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
                        st.success(f"已成功提取 {file.name} 的表格")
                        processed_count += 1
                    else:
                        st.error(f"无法从文件 {file.name} 中提取有效数据")
                except Exception as e:
                    st.error(f"处理文件 {file.name} 时出错: {e}")
            if processed_count > 0:
                st.success(f"成功处理 {processed_count} 个文件")
        # 返回新增内容
        return all_uploaded_files, extracted_files, selected_files, uploaded_main_dfs, uploaded_tertiary_dfs, include_self_owned_labor