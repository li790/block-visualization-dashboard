import streamlit as st
from pathlib import Path
import os
import tempfile
from utils.data_processor import get_excel_files, extract_table_from_excel

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
            
            # 自动落盘逻辑
            total_upload_size = sum([file.size for file in all_uploaded_files]) if all_uploaded_files else 0
            auto_save_triggered = False
            if total_upload_size > 500 * 1024 * 1024:  # 500MB
                st.warning("上传文件总大小超过500MB，已自动保存到data目录后再分析。")
                auto_save_triggered = True
                saved_count = 0
                for file in all_uploaded_files:
                    try:
                        file_path = data_dir / file.name
                        with open(file_path, "wb") as f:
                            f.write(file.getbuffer())
                        saved_count += 1
                    except Exception as e:
                        st.error(f"自动保存文件 {file.name} 失败: {e}")
                if saved_count > 0:
                    st.success(f"自动保存 {saved_count} 个文件到data目录")
            
            # 保存选项
            save_option = st.selectbox(
                "保存到data目录?",
                ["不保存", "保存所有文件", "选择保存"]
            )
            
            if save_option != "不保存":
                if save_option == "保存所有文件":
                    if st.button("保存所有文件"):
                        saved_count = 0
                        for file in all_uploaded_files:
                            try:
                                file_path = data_dir / file.name
                                with open(file_path, "wb") as f:
                                    f.write(file.getbuffer())
                                saved_count += 1
                            except Exception as e:
                                st.error(f"保存文件 {file.name} 失败: {e}")
                        
                        if saved_count > 0:
                            st.success(f"成功保存 {saved_count} 个文件")
                            st.rerun()
                
                elif save_option == "选择保存":
                    # 让用户选择要保存的文件
                    files_to_save = st.multiselect(
                        "选择要保存的文件:",
                        [f.name for f in all_uploaded_files],
                        default=[f.name for f in all_uploaded_files]
                    )
                    
                    if st.button("保存选中文件"):
                        saved_count = 0
                        for file in all_uploaded_files:
                            if file.name in files_to_save:
                                try:
                                    file_path = data_dir / file.name
                                    with open(file_path, "wb") as f:
                                        f.write(file.getbuffer())
                                    saved_count += 1
                                except Exception as e:
                                    st.error(f"保存文件 {file.name} 失败: {e}")
                        
                        if saved_count > 0:
                            st.success(f"成功保存 {saved_count} 个文件")
                            st.rerun()
        
        st.divider()
        
        # 选择现有文件
        existing_files = get_excel_files(data_dir)
        if existing_files:
            selected_files = st.multiselect(
                "选择要分析的文件:",
                [f.name for f in existing_files],
                default=[existing_files[0].name] if existing_files else None
            )
        else:
            selected_files = []
            st.warning("data目录中没有Excel文件")
        
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
                    main_df, tertiary_df = extract_table_from_excel(file)
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
        return all_uploaded_files, extracted_files, selected_files, uploaded_main_dfs, uploaded_tertiary_dfs