import streamlit as st
from pathlib import Path
from typing import Dict, List, Tuple
import re
from utils.data_processor import get_excel_files

# 定义SVG图标
CHART_ICON = """
<svg t="1758768181220" class="icon" viewBox="0 0 1024 1024" version="1.1" xmlns="http://www.w3.org/2000/svg" p-id="1621" width="20" height="20" style="display: inline-block; vertical-align: middle; margin-right: 8px;">
    <path d="M532.15763 144.004741a75.851852 75.851852 0 0 1 75.851851 75.851852v440.301037a75.851852 75.851852 0 0 1-75.851851 75.851851H491.861333a75.851852 75.851852 0 0 1-75.851852-75.851851V219.856593a75.851852 75.851852 0 0 1 75.851852-75.851852h40.296297z m256 111.995259a75.851852 75.851852 0 0 1 75.851851 75.851852V660.176593a75.851852 75.851852 0 0 1-75.851851 75.851851H747.861333a75.851852 75.851852 0 0 1-75.851852-75.851851V331.851852a75.851852 75.851852 0 0 1 75.851852-75.851852h40.296297z m-512 112.014222a75.851852 75.851852 0 0 1 75.851851 75.851852v216.291556a75.851852 75.851852 0 0 1-75.851851 75.851851H235.861333a75.851852 75.851852 0 0 1-75.851852-75.851851V443.847111a75.851852 75.851852 0 0 1 75.851852-75.851852h40.296297z" fill="#279CFF" p-id="1622"></path>
    <path d="M160.009481 816.014222m32.009482 0l639.981037 0q32.009481 0 32.009481 32.009482l0-0.018963q0 32.009481-32.009481 32.009481l-639.981037 0q-32.009481 0-32.009482-32.009481l0 0.018963q0-32.009481 32.009482-32.009482Z" fill="#279CFF" fill-opacity=".5" p-id="1623"></path>
</svg>
"""

def split_type_and_display(project: str) -> Tuple[str, str]:
    """解析类型前缀与项目显示名"""
    parts = project.split('-', 1)
    if len(parts) == 2 and parts[0]:
        return parts[0], parts[1]
    return project, project

def categorize_files(file_names: List[str]) -> Dict[str, List[str]]:
    """将文件按类型分类"""
    type_to_files = {}
    for filename in file_names:
        # 去掉.xlsx扩展名
        project_name = filename.replace('.xlsx', '').replace('.xls', '')
        type_prefix, display_name = split_type_and_display(project_name)
        type_to_files.setdefault(type_prefix, []).append(filename)
    return type_to_files


def render_file_selector(data_dir: Path) -> List[str]:
    """渲染新的文件选择界面，支持按类型分类选择"""
    
    # 获取所有Excel文件
    existing_files = get_excel_files(data_dir)
    if not existing_files:
        st.warning("data文件夹中没有找到Excel文件")
        return []
    
    file_names = [f.name for f in existing_files]
    
    # 初始化session state
    if 'selected_files' not in st.session_state:
        st.session_state['selected_files'] = []
    if 'expanded_types' not in st.session_state:
        st.session_state['expanded_types'] = set()
    if 'type_selections' not in st.session_state:
        st.session_state['type_selections'] = {}
    
    # 按类型分类文件
    type_to_files = categorize_files(file_names)
    
    # 显示选择统计
    total_files = len(file_names)
    selected_count = len(st.session_state['selected_files'])
    
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown(f"<strong style='display: flex; align-items: center;'>{CHART_ICON} 文件选择 ({selected_count}/{total_files})</strong>", unsafe_allow_html=True)
    with col2:
        if st.button("全选", key="select_all_btn"):
            st.session_state['selected_files'] = file_names.copy()
            # 更新所有类型的选中状态
            for type_prefix in type_to_files.keys():
                st.session_state['type_selections'][type_prefix] = True
            st.rerun()
    with col3:
        if st.button("全不选", key="deselect_all_btn"):
            st.session_state['selected_files'] = []
            # 清空所有类型的选中状态
            for type_prefix in type_to_files.keys():
                st.session_state['type_selections'][type_prefix] = False
            st.rerun()
    
    st.divider()
    
    # 按类型显示文件选择
    for type_prefix, files in sorted(type_to_files.items()):
        type_display_name = type_prefix
        file_count = len(files)
        
        # 检查该类型是否全部选中
        type_files = set(files)
        selected_type_files = set(st.session_state['selected_files'])
        is_type_fully_selected = type_files.issubset(selected_type_files)
        is_type_partially_selected = bool(type_files.intersection(selected_type_files))
        
        # 类型选择状态
        if type_prefix not in st.session_state['type_selections']:
            st.session_state['type_selections'][type_prefix] = is_type_fully_selected
        
        # 创建类型选择行
        col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
        
        with col1:
            # 类型复选框
            type_selected = st.checkbox(
                f"**{type_display_name}** ({file_count}个文件)",
                value=st.session_state['type_selections'][type_prefix],
                key=f"type_checkbox_{type_prefix}"
            )
            st.session_state['type_selections'][type_prefix] = type_selected
        
        with col2:
            # 展开/收起按钮
            expand_key = f"expand_{type_prefix}"
            if st.button("详情", key=expand_key):
                if type_prefix in st.session_state['expanded_types']:
                    st.session_state['expanded_types'].remove(type_prefix)
                else:
                    st.session_state['expanded_types'].add(type_prefix)
                st.rerun()
        
        with col3:
            # 全选该类型
            if st.button("全选", key=f"select_type_{type_prefix}"):
                # 添加该类型的所有文件
                for file in files:
                    if file not in st.session_state['selected_files']:
                        st.session_state['selected_files'].append(file)
                st.session_state['type_selections'][type_prefix] = True
                st.rerun()
        
        with col4:
            # 取消选择该类型
            if st.button("取消", key=f"deselect_type_{type_prefix}"):
                # 移除该类型的所有文件
                st.session_state['selected_files'] = [
                    f for f in st.session_state['selected_files'] 
                    if f not in files
                ]
                st.session_state['type_selections'][type_prefix] = False
                st.rerun()
        
        # 处理类型选择变化
        if type_selected and not is_type_fully_selected:
            # 选中该类型的所有文件
            for file in files:
                if file not in st.session_state['selected_files']:
                    st.session_state['selected_files'].append(file)
        elif not type_selected and is_type_partially_selected:
            # 取消选择该类型的所有文件
            st.session_state['selected_files'] = [
                f for f in st.session_state['selected_files'] 
                if f not in files
            ]
        
        # 显示类型详情（展开时）
        if type_prefix in st.session_state['expanded_types']:
            with st.expander(f"📋 {type_display_name} 详情", expanded=True):
                # 显示该类型的所有文件
                for file in files:
                    col1, col2 = st.columns([4, 1])
                    
                    with col1:
                        # 文件选择复选框
                        file_selected = st.checkbox(
                            file.replace('.xlsx', ''),
                            value=file in st.session_state['selected_files'],
                            key=f"file_checkbox_{file}"
                        )
                    
                    with col2:
                        # 删除按钮
                        if st.button("❌", key=f"remove_{file}", help="从选择中移除"):
                            if file in st.session_state['selected_files']:
                                st.session_state['selected_files'].remove(file)
                                # 更新类型选择状态
                                remaining_type_files = [
                                    f for f in files 
                                    if f in st.session_state['selected_files']
                                ]
                                st.session_state['type_selections'][type_prefix] = len(remaining_type_files) == len(files)
                            st.rerun()
                    
                    # 处理单个文件选择变化
                    if file_selected and file not in st.session_state['selected_files']:
                        st.session_state['selected_files'].append(file)
                        # 检查是否该类型的所有文件都已选中
                        remaining_type_files = [
                            f for f in files 
                            if f in st.session_state['selected_files']
                        ]
                        st.session_state['type_selections'][type_prefix] = len(remaining_type_files) == len(files)
                    elif not file_selected and file in st.session_state['selected_files']:
                        st.session_state['selected_files'].remove(file)
                        st.session_state['type_selections'][type_prefix] = False
        
        st.divider()
    
    # 显示当前选择的文件
    if st.session_state['selected_files']:
        st.markdown("**📌 当前选择的文件:**")
        
        # 按类型分组显示选中的文件
        selected_by_type = {}
        for file in st.session_state['selected_files']:
            project_name = file.replace('.xlsx', '').replace('.xls', '')
            type_prefix, _ = split_type_and_display(project_name)
            selected_by_type.setdefault(type_prefix, []).append(file)
        
        for type_prefix, files in selected_by_type.items():
            type_display_name = type_prefix
            with st.expander(f"✅ {type_display_name} ({len(files)}个文件)", expanded=False):
                for file in files:
                    st.write(f"• {file}")
    
    return st.session_state['selected_files']
