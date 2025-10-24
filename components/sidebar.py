import streamlit as st
from pathlib import Path
import os
import tempfile
from utils.data_processor import get_excel_files, extract_table_from_excel
from utils.cache_manager import get_cache_manager
from utils.file_utils import split_type_and_display, categorize_files

# 定义文件夹SVG图标
FOLDER_ICON = """
<svg t="1758769663608" class="icon" viewBox="0 0 1026 1024" version="1.1" xmlns="http://www.w3.org/2000/svg" p-id="15298" width="20" height="20" style="display: inline-block; vertical-align: middle; margin-right: 8px;">
    <path d="M923.904 140.544H550.4a156.416 156.416 0 0 1-87.296-25.6l-89.6-59.648a95.744 95.744 0 0 0-53.76-16.384h-215.04a102.4 102.4 0 0 0-102.4 102.4v742.656a102.4 102.4 0 0 0 102.4 102.4h819.2a102.4 102.4 0 0 0 102.4-102.4v-640a102.4 102.4 0 0 0-102.4-103.424z" fill="#0067E4" p-id="15299"></path>
    <path d="M2.56 243.456m102.4 0l819.2 0q102.4 0 102.4 102.4l0 538.368q0 102.4-102.4 102.4l-819.2 0q-102.4 0-102.4-102.4l0-538.368q0-102.4 102.4-102.4Z" fill="#0085FF" p-id="15300"></path>
</svg>
"""

# ProDAM3 图标（简洁方形徽标）
PRODAM_ICON = """
<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" width="18" height="18" style="display:inline-block;vertical-align:middle;margin-right:6px;">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#3b82f6"/>
      <stop offset="100%" stop-color="#2563eb"/>
    </linearGradient>
  </defs>
  <rect x="4" y="4" width="56" height="56" rx="10" fill="url(#g)"/>
  <text x="50%" y="52%" dominant-baseline="middle" text-anchor="middle" font-family="Segoe UI, Microsoft YaHei" font-size="20" fill="#ffffff" font-weight="700">P</text>
</svg>
"""


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
    
    /* 保留主要部分之间的分割线 */
    .stDivider {
        background: linear-gradient(90deg, transparent, #e5e7eb, transparent);
        height: 1px;
        margin: 4px 0 !important;
        border: none;
    }
    
    /* 保留类别之间的分割线 */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, #e5e7eb, transparent);
        margin: 2px 0 !important;
    }

    /* 文件选择区专用分隔线：与上方按钮拉开更大间距，防止重叠视觉 */
    .file-selector-hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, #e5e7eb, transparent);
        margin: 12px 0 !important; /* 基础：更大一点 */
    }

    /* 明确在侧边栏中提升优先级，覆盖 [data-testid="stSidebar"] hr 规则 */
    [data-testid="stSidebar"] .file-selector-hr {
        margin-top: 18px !important;
        margin-bottom: 12px !important;
        display: block !important;
    }

    /* 缩小侧边栏各模块之间的垂直间距 */
    [data-testid="stSidebar"] .stVerticalBlock {
        gap: 0.25rem !important; /* 默认约为 1rem */
    }

    /* 进一步限定侧边栏内分割线的外边距，确保覆盖 */
    [data-testid="stSidebar"] .stDivider {
        margin-top: 4px !important;
        margin-bottom: 4px !important;
    }
    [data-testid="stSidebar"] hr {
        margin-top: 2px !important;
        margin-bottom: 2px !important;
    }

    /* 文件选择区域上方容器：给按钮行增加额外下边距 */
    [data-testid="stSidebar"] .file-selector-controls {
        margin-top: 14px !important;   /* 往下移一点（+2px） */
        margin-bottom: 16px !important;
        display: block;
    }

    /* 当 Streamlit 的列不在自定义容器内时，用间隔元素强制拉开距离 */
    [data-testid="stSidebar"] .file-selector-spacer {
        display: block;
        height: 2px; /* 额外增加 2px 间距 */
    }

    /* 标题与按钮之间的专用占位（放在按钮容器之前） */
    [data-testid="stSidebar"] .file-selector-spacer-before {
        display: block;
        height: 2px; /* 明确在标题与按钮之间再加 2px */
    }

    /* 预留：蓝色按钮规则将被移动到样式块末尾以提升优先级 */
    
    /* 隐藏文件列表内部的分割线 */
    .streamlit-expanderContent hr {
        display: none;
    }
    
    /* 白色卡片包裹效果 */
    .stCheckbox > label {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 12px 16px;
        margin: 2px 0;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        transition: all 0.2s ease;
    }
    
    .stCheckbox > label:hover {
        background: #f8fafc;
        border-color: #d1d5db;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    .stCheckbox > label:has(input:checked) {
        background: #dbeafe;
        border-color: #3b82f6;
        color: #1d4ed8;
    }
    
    /* 勾选框选中状态 - 蓝色填充 */
    .stCheckbox > label:has(input:checked) .stCheckbox > div[data-testid="stCheckbox"] > div {
        background-color: #3b82f6;
        border-color: #3b82f6;
    }
    
    .stCheckbox > label:has(input:checked) .stCheckbox > div[data-testid="stCheckbox"] > div > div {
        background-color: #ffffff;
    }
    
    /* 按钮样式 - 还原为白色背景 */
    .stButton > button {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 6px;
        padding: 6px 12px;
        font-size: 12px;
        color: #6b7280;
        transition: all 0.2s ease;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
    }
    
    .stButton > button:hover {
        background: #f8fafc;
        border-color: #d1d5db;
        color: #374151;
    }
    
    /* 开始分析和重置按钮 - 蓝色填充 */
    button[data-testid="baseButton-secondary"]:contains("开始分析"),
    button[data-testid="baseButton-secondary"]:contains("重置") {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
        border: 1px solid #3b82f6 !important;
        color: #ffffff !important;
        font-weight: 500;
        box-shadow: 0 2px 4px rgba(59, 130, 246, 0.3);
    }
    
    button[data-testid="baseButton-secondary"]:contains("开始分析"):hover,
    button[data-testid="baseButton-secondary"]:contains("重置"):hover {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        border-color: #2563eb !important;
        color: #ffffff !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(59, 130, 246, 0.4);
    }
    
    /* 减少模块间距 */
    .element-container {
        margin: 1px 0;
    }
    
    .stCheckbox {
        margin-bottom: 0;
    }
    
    .stButton {
        margin-bottom: 0;
    }
    
    /* 展开器样式 */
    .streamlit-expanderHeader {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        margin: 1px 0;
        padding: 12px 16px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
    }
    
    .streamlit-expanderContent {
        background: #f8fafc;
        border: 1px solid #e5e7eb;
        border-radius: 0 0 8px 8px;
        margin: 0;
        padding: 8px 16px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        border-top: none;
    }
    
    /* 减少整体间距 */
    .main .block-container {
        padding: 0.5rem;
    }
    </style>
    <style>
    /* 仅针对“开始分析/重置”两个按钮的专属样式（末尾定义，优先级更高） */
    .analysis-actions [data-testid^="baseButton-"],
    .is-analysis {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
        color: #ffffff !important;
        border: 1px solid #3b82f6 !important;
        box-shadow: 0 2px 4px rgba(59, 130, 246, 0.25) !important;
    }
    .analysis-actions [data-testid^="baseButton-"]:hover,
    .is-analysis:hover {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        border-color: #2563eb !important;
        color: #ffffff !important;
    }
    .analysis-actions [data-testid^="baseButton-"]:active,
    .is-analysis:active {
        transform: translateY(1px);
        box-shadow: 0 1px 3px rgba(59, 130, 246, 0.2) !important;
    }
    </style>
    
    <script>
    // 确保开始分析和重置按钮有蓝色样式
    function styleBlueButtons() {
        // 遍历所有 Streamlit 按钮，按文本内容匹配“开始分析/重置”两枚按钮
        const buttons = document.querySelectorAll('button[data-testid^="baseButton-"]');
        buttons.forEach(button => {
            const text = (button.textContent || button.innerText || '').trim();
            if (text.includes('开始分析') || text.includes('重置')) {
                button.style.background = 'linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)';
                button.style.border = '1px solid #3b82f6';
                button.style.color = '#ffffff';
                button.style.fontWeight = '500';
                button.style.boxShadow = '0 2px 4px rgba(59, 130, 246, 0.3)';

                button.addEventListener('mouseenter', function() {
                    this.style.background = 'linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)';
                    this.style.borderColor = '#2563eb';
                    this.style.color = '#ffffff';
                    this.style.transform = 'translateY(-1px)';
                    this.style.boxShadow = '0 4px 8px rgba(59, 130, 246, 0.4)';
                });

                button.addEventListener('mouseleave', function() {
                    this.style.background = 'linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)';
                    this.style.borderColor = '#3b82f6';
                    this.style.color = '#ffffff';
                    this.style.transform = 'translateY(0)';
                    this.style.boxShadow = '0 2px 4px rgba(59, 130, 246, 0.3)';
                });
            }
        });
    }
    
    // 页面加载后执行
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', styleBlueButtons);
    } else {
        styleBlueButtons();
    }
    
    // 定期检查新添加的按钮
    setInterval(styleBlueButtons, 1000);
    </script>
    """, unsafe_allow_html=True)


def render_categorized_file_selector(data_dir):
    """渲染分类文件选择器"""
    # 强制清除可能的缓存问题
    if 'file_selector_version' not in st.session_state:
        st.session_state['file_selector_version'] = 'v2.0'
    
    # 添加调试信息
    
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
    if 'start_analysis' not in st.session_state:
        st.session_state['start_analysis'] = False
    # 记录需要强制取消勾选的文件（用于避免直接修改已实例化的checkbox）
    if 'files_to_uncheck' not in st.session_state:
        st.session_state['files_to_uncheck'] = set()
    
    # 按类型分类文件
    type_to_files = categorize_files(file_names)
    
    # 显示选择统计
    total_files = len(file_names)
    selected_count = len(st.session_state['selected_files'])
    
    # 文件选择标题
    st.markdown(f"<strong style='display: flex; align-items: center;'>{FOLDER_ICON} 文件选择 ({selected_count}/{total_files})</strong>", unsafe_allow_html=True)
    
    # 标题与按钮之间放置一个显式 2px 占位，确保再拉开距离
    st.markdown('<div class="file-selector-spacer-before"></div>', unsafe_allow_html=True)
    # 全局操作按钮（外层容器加类名，便于与下方区域拉开距离）
    st.markdown('<div class="file-selector-controls">', unsafe_allow_html=True)
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button(" 全选 ", key="select_all_btn", use_container_width=True):
            st.session_state['selected_files'] = file_names.copy()
            # 更新所有类型的选中状态
            for type_prefix in type_to_files.keys():
                st.session_state['type_selections'][type_prefix] = True
            st.rerun()
    with col2:
        if st.button(" 取消 ", key="deselect_all_btn", use_container_width=True):
            st.session_state['selected_files'] = []
            # 清空所有类型的选中状态
            for type_prefix in type_to_files.keys():
                st.session_state['type_selections'][type_prefix] = False
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    # 如果上面的margin-top未生效，使用一个显式的spacer元素再拉开2px
    st.markdown('<div class="file-selector-spacer"></div>', unsafe_allow_html=True)
    
    # 文件选择与按钮之间的专用分隔线（更大间距）
    st.markdown('<hr class="file-selector-hr" />', unsafe_allow_html=True)
    
    # 按类型显示文件选择
    for type_prefix, files in sorted(type_to_files.items()):
        type_display_name = type_prefix
        file_count = len(files)
        
        # 检查该类型是否全部/部分选中
        type_files = set(files)
        selected_type_files = set(st.session_state['selected_files'])
        is_type_fully_selected = type_files.issubset(selected_type_files)
        is_type_partially_selected = bool(type_files.intersection(selected_type_files))

        # 创建类型选择行 - 减少横向间距
        col1, col2 = st.columns([5, 1])  # 减少按钮列宽度
        
        with col1:
            # 类型复选框
            # 仅在“全部已选中”时勾选，部分选中则保持未勾选（避免误触发全清）
            type_selected = st.checkbox(
                f" {type_display_name} ({file_count}个文件)",
                value=is_type_fully_selected,
                key=f"type_checkbox_{type_prefix}"
            )
        
        with col2:
            # 详情按钮
            expand_key = f"expand_{type_prefix}"
            if st.button("详情", key=expand_key, use_container_width=True, help="展开/收起详情"):
                if type_prefix in st.session_state['expanded_types']:
                    st.session_state['expanded_types'].remove(type_prefix)
                else:
                    st.session_state['expanded_types'].add(type_prefix)
                st.rerun()
        
        # 处理类型选择变化（仅在用户切换“全选/全不选”时生效）
        if type_selected and not is_type_fully_selected:
            # 从“部分/全不选” → “全选”
            for file in files:
                if file not in st.session_state['selected_files']:
                    st.session_state['selected_files'].append(file)
        elif (not type_selected) and is_type_fully_selected:
            # 从“全选” → “全不选”，避免在“部分选中”时误清空
            st.session_state['selected_files'] = [
                f for f in st.session_state['selected_files']
                if f not in files
            ]
        
        # 显示类型详情（展开时）
        if type_prefix in st.session_state['expanded_types']:
            with st.expander(f" {type_display_name} 详情", expanded=True):
                # 显示该类型的所有文件（按名称排序，保证键稳定）
                for i, file in enumerate(sorted(files)):
                    # 在渲染checkbox之前，若该文件标记为需要取消勾选，则删除其状态key
                    cb_key = f"file_checkbox_{type_prefix}_{i}"
                    if file in st.session_state['files_to_uncheck']:
                        if cb_key in st.session_state:
                            del st.session_state[cb_key]
                        # 确保从集合移除，避免重复处理
                        st.session_state['files_to_uncheck'].discard(file)

                    # 文件选择复选框
                    file_selected = st.checkbox(
                        f" {file.replace('.xlsx', '')}",
                        value=file in st.session_state['selected_files'],
                        key=cb_key
                    )

                    # 处理单个文件选择变化
                    if file_selected and file not in st.session_state['selected_files']:
                        st.session_state['selected_files'].append(file)
                    elif (not file_selected) and file in st.session_state['selected_files']:
                        st.session_state['selected_files'].remove(file)
        
        st.divider()
    
    # 显示当前选择的文件
    if st.session_state['selected_files']:
        st.markdown("**当前选择的文件**")
        
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
    
    # 一键分析按钮（包裹专属样式容器，仅影响该区域按钮）
  
    if st.session_state['selected_files']:
        st.markdown('<div class="analysis-actions">', unsafe_allow_html=True)
        col1, col2 = st.columns([2, 1])
        with col1:
            if st.button("开始分析", key="start_analysis_btn", use_container_width=True, type="secondary"):
                # 设置分析标志
                st.session_state['start_analysis'] = True
                st.rerun()
        with col2:
            if st.button("重置", key="reset_selection_btn", use_container_width=True, type="secondary"):
                st.session_state['selected_files'] = []
                st.session_state['start_analysis'] = False
                # 清空所有类型的选中状态
                for type_prefix in type_to_files.keys():
                    st.session_state['type_selections'][type_prefix] = False
                # 清除处理过的数据缓存
                if 'processed_main_dfs' in st.session_state:
                    del st.session_state['processed_main_dfs']
                    del st.session_state['processed_tertiary_dfs']
                    del st.session_state['processed_include_self_owned_labor']
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info(" 请先选择要分析的文件")
    
    return st.session_state['selected_files']

def render_sidebar(data_dir):
    """渲染侧边栏"""
    # 注入自定义CSS
    inject_custom_css()
    
    # 初始化时间戳
    import time
    if 'current_time' not in st.session_state:
        st.session_state.current_time = time.time()
    current_time = time.time()
    
    # 强制清除可能的缓存问题
    if 'sidebar_version' not in st.session_state:
        st.session_state['sidebar_version'] = 'v2.0'
    
    with st.sidebar:
        
        # 主要费项表格选择
        st.markdown(f"**主要费项表格选择**", unsafe_allow_html=True)
        
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
            st.info(" 当前选择：包含自有人工成本的主要费项表格")
        else:
            st.info(" 当前选择：不包含自有人工成本的主要费项表格")
        
        st.divider()
        
        # 缓存状态显示
        cache_manager = get_cache_manager()
        cache_stats = cache_manager.get_cache_stats()
        if 'error' not in cache_stats and cache_stats['cache_count'] > 0:
            st.info(f"⚡ 缓存状态: {cache_stats['cache_count']} 个文件已缓存 ({cache_stats['total_size_mb']}MB)")
        
        # 新的分类文件选择器
        selected_files = render_categorized_file_selector(data_dir)

        # 侧边栏底部外链：物业成本各模块数据（使用icon文件夹中的图片作为图标）
        st.markdown("---")
        try:
            import base64
            possible_icons = [
                Path('icon/ProDAM3.png'),
                Path('icon/ProDAM3.jpg'),
                Path('icons/ProDAM3.png'),
                Path('icons/ProDAM3.jpg')
            ]
            icon_src = ''
            for p in possible_icons:
                if p.exists():
                    with open(p, 'rb') as f:
                        b64 = base64.b64encode(f.read()).decode('utf-8')
                        ext = 'png' if p.suffix.lower() == '.png' else 'jpeg'
                        icon_src = f'data:image/{ext};base64,{b64}'
                        break
            icon_img_html = f'<img src="{icon_src}" alt="ProDAM" width="64" height="64" style="vertical-align:middle;" />' if icon_src else ''
        except Exception:
            icon_img_html = ''

        st.markdown(
            '<div style="padding:6px 0;">'
            f'<a href="http://101.33.231.179:8800" target="_blank" '
            'style="display:inline-flex;align-items:center;gap:10px;text-decoration:none;color:#2563eb;font-weight:700;font-size:18px;">'
            f'{icon_img_html} 物业成本各模块数据分析</a>'
            '</div>',
            unsafe_allow_html=True
        )
        
        # 返回简化的内容（移除上传相关参数）
        return [], [], selected_files, {}, {}, include_self_owned_labor, []