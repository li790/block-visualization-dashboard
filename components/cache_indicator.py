import streamlit as st
import time
from utils.cache_manager import get_cache_manager


def render_cache_indicator():
    """渲染缓存状态指示器"""
    cache_manager = get_cache_manager()
    cache_stats = cache_manager.get_cache_stats()
    
    if 'error' not in cache_stats and cache_stats['cache_count'] > 0:
        # 创建缓存状态指示器
        st.markdown("""
        <style>
        .cache-indicator {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 8px;
            padding: 8px 12px;
            margin: 4px 0;
            color: white;
            font-size: 12px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .cache-indicator:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown(
            f'<div class="cache-indicator">⚡ 缓存: {cache_stats["cache_count"]} 文件 ({cache_stats["total_size_mb"]}MB)</div>',
            unsafe_allow_html=True
        )

def render_performance_info():
    """渲染性能信息"""
    if 'performance_start_time' not in st.session_state:
        st.session_state.performance_start_time = time.time()
    
    if 'last_operation_time' in st.session_state:
        elapsed_time = st.session_state.last_operation_time
        if elapsed_time < 1:
            status = "🟢 极快"
            color = "#4CAF50"
        elif elapsed_time < 3:
            status = "🟡 快速"
            color = "#FF9800"
        else:
            status = "🔴 较慢"
            color = "#F44336"
        
        st.markdown(f"""
        <style>
        .performance-info {{
            background: {color}20;
            border: 1px solid {color};
            border-radius: 6px;
            padding: 6px 10px;
            margin: 4px 0;
            color: {color};
            font-size: 11px;
            text-align: center;
        }}
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown(
            f'<div class="performance-info">{status} 加载时间: {elapsed_time:.1f}秒</div>',
            unsafe_allow_html=True
        )

def start_performance_timer():
    """开始性能计时"""
    st.session_state.performance_start_time = time.time()

def end_performance_timer():
    """结束性能计时并记录"""
    if 'performance_start_time' in st.session_state:
        elapsed_time = time.time() - st.session_state.performance_start_time
        st.session_state.last_operation_time = elapsed_time
        return elapsed_time
    return 0

def show_cache_benefit_message():
    """显示缓存优势信息"""
    cache_manager = get_cache_manager()
    cache_stats = cache_manager.get_cache_stats()
    
    if 'error' not in cache_stats and cache_stats['cache_count'] > 0:
        st.success(f"🚀 **性能提升:** 检测到 {cache_stats['cache_count']} 个缓存文件，重复选择相同文件时将大幅提升加载速度！") 