import streamlit as st
import pandas as pd
import hashlib
import time
from pathlib import Path
from typing import Dict, Any, Optional, Tuple
import pickle
import os

class MemoryCacheManager:
    """内存缓存管理器，使用Streamlit session_state进行缓存，避免文件系统冲突"""
    
    def __init__(self):
        # 初始化内存缓存
        if 'memory_cache' not in st.session_state:
            st.session_state.memory_cache = {}
        if 'cache_metadata' not in st.session_state:
            st.session_state.cache_metadata = {}
        if 'cache_timestamps' not in st.session_state:
            st.session_state.cache_timestamps = {}
        
    def _generate_cache_key(self, file_path: str, include_self_owned_labor: bool, 
                           file_size: int, file_mtime: float) -> str:
        """生成缓存键"""
        # 使用文件路径、参数、大小和修改时间生成唯一键
        key_data = f"{file_path}_{include_self_owned_labor}_{file_size}_{file_mtime}"
        return hashlib.md5(key_data.encode()).hexdigest()
    
    def _get_file_info(self, file_path: Path) -> Tuple[int, float]:
        """获取文件信息"""
        try:
            stat = file_path.stat()
            return stat.st_size, stat.st_mtime
        except:
            return 0, 0
    
    def get_cached_data(self, file_path: str, include_self_owned_labor: bool) -> Optional[Tuple[pd.DataFrame, pd.DataFrame]]:
        """从内存缓存获取数据"""
        try:
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                return None
                
            file_size, file_mtime = self._get_file_info(file_path_obj)
            cache_key = self._generate_cache_key(str(file_path), include_self_owned_labor, file_size, file_mtime)
            
            # 检查内存缓存是否存在且有效
            if cache_key in st.session_state.memory_cache:
                # 检查缓存是否过期（24小时）
                if cache_key in st.session_state.cache_timestamps:
                    if time.time() - st.session_state.cache_timestamps[cache_key] < 86400:
                        return st.session_state.memory_cache[cache_key]
                    else:
                        # 缓存过期，删除
                        self._remove_cache(cache_key)
            
            return None
            
        except Exception as e:
            return None
    
    def save_cached_data(self, file_path: str, include_self_owned_labor: bool, 
                        main_df: pd.DataFrame, tertiary_df: pd.DataFrame):
        """保存数据到内存缓存"""
        try:
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                return
                
            file_size, file_mtime = self._get_file_info(file_path_obj)
            cache_key = self._generate_cache_key(str(file_path), include_self_owned_labor, file_size, file_mtime)
            
            # 保存数据到内存缓存
            st.session_state.memory_cache[cache_key] = (main_df, tertiary_df)
            st.session_state.cache_timestamps[cache_key] = time.time()
            
            # 更新缓存元数据
            st.session_state.cache_metadata[cache_key] = {
                'file_path': str(file_path),
                'include_self_owned_labor': include_self_owned_labor,
                'file_size': file_size,
                'file_mtime': file_mtime,
                'timestamp': time.time()
            }
            
        except Exception as e:
            pass
    
    def get_analysis_cache(self, cache_type: str, project_name: str, month: int, 
                          include_self_owned_labor: bool = False) -> Optional[Dict[str, Any]]:
        """从内存缓存获取分析结果"""
        try:
            cache_key = f"{cache_type}_{project_name}_{month}_{include_self_owned_labor}"
            
            if cache_key in st.session_state.memory_cache:
                # 检查缓存是否过期（24小时）
                if cache_key in st.session_state.cache_timestamps:
                    if time.time() - st.session_state.cache_timestamps[cache_key] < 86400:
                        return st.session_state.memory_cache[cache_key]
                    else:
                        # 缓存过期，删除
                        self._remove_cache(cache_key)
            
            return None
            
        except Exception:
            return None
    
    def save_analysis_cache(self, cache_type: str, project_name: str, month: int, 
                           data: Dict[str, Any], include_self_owned_labor: bool = False):
        """保存分析结果到内存缓存"""
        try:
            cache_key = f"{cache_type}_{project_name}_{month}_{include_self_owned_labor}"
            st.session_state.memory_cache[cache_key] = data
            st.session_state.cache_timestamps[cache_key] = time.time()
                
        except Exception:
            pass
    
    def get_secondary_fee_cache(self, project_name: str, month: int, 
                               include_self_owned_labor: bool = False) -> Optional[pd.DataFrame]:
        """获取二级费项缓存"""
        return self.get_analysis_cache("secondary_fee", project_name, month, include_self_owned_labor)
    
    def save_secondary_fee_cache(self, project_name: str, month: int, 
                                data: pd.DataFrame, include_self_owned_labor: bool = False):
        """保存二级费项缓存"""
        self.save_analysis_cache("secondary_fee", project_name, month, data, include_self_owned_labor)
    
    def get_anomaly_cache(self, project_name: str, month: int, 
                         include_self_owned_labor: bool = False) -> Optional[Dict[str, Any]]:
        """获取异常数据缓存"""
        return self.get_analysis_cache("anomaly", project_name, month, include_self_owned_labor)
    
    def save_anomaly_cache(self, project_name: str, month: int, 
                          data: Dict[str, Any], include_self_owned_labor: bool = False):
        """保存异常数据缓存"""
        self.save_analysis_cache("anomaly", project_name, month, data, include_self_owned_labor)
    
    def get_project_analysis_cache(self, project_name: str, month: int, 
                                  include_self_owned_labor: bool = False) -> Optional[Dict[str, Any]]:
        """获取项目详细分析缓存"""
        return self.get_analysis_cache("project_analysis", project_name, month, include_self_owned_labor)
    
    def save_project_analysis_cache(self, project_name: str, month: int, 
                                   data: Dict[str, Any], include_self_owned_labor: bool = False):
        """保存项目详细分析缓存"""
        self.save_analysis_cache("project_analysis", project_name, month, data, include_self_owned_labor)
    
    def _remove_cache(self, cache_key: str):
        """删除指定的缓存"""
        try:
            if cache_key in st.session_state.memory_cache:
                del st.session_state.memory_cache[cache_key]
            if cache_key in st.session_state.cache_metadata:
                del st.session_state.cache_metadata[cache_key]
            if cache_key in st.session_state.cache_timestamps:
                del st.session_state.cache_timestamps[cache_key]
        except Exception as e:
            st.warning(f"删除缓存失败: {e}")
    
    def clear_all_cache(self):
        """清除所有缓存"""
        try:
            st.session_state.memory_cache.clear()
            st.session_state.cache_metadata.clear()
            st.session_state.cache_timestamps.clear()
            st.success("🗑️ 已清除所有内存缓存")
            
        except Exception as e:
            st.error(f"清除缓存失败: {e}")
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """获取缓存统计信息"""
        try:
            cache_count = len(st.session_state.memory_cache)
            metadata_count = len(st.session_state.cache_metadata)
            
            # 估算内存使用量（粗略计算）
            total_size_mb = round(cache_count * 0.1, 2)  # 假设每个缓存项约0.1MB
            
            return {
                'cache_count': cache_count,
                'total_size_mb': total_size_mb,
                'metadata_count': metadata_count
            }
        except Exception as e:
            return {'error': str(e)}
    
    def cleanup_expired_cache(self):
        """清理过期的缓存"""
        try:
            current_time = time.time()
            expired_keys = []
            
            for cache_key, timestamp in st.session_state.cache_timestamps.items():
                if current_time - timestamp > 86400:  # 24小时
                    expired_keys.append(cache_key)
            
            for cache_key in expired_keys:
                self._remove_cache(cache_key)
            
            # 静默清理过期缓存
                
        except Exception as e:
            st.warning(f"清理过期缓存失败: {e}")

# 全局内存缓存管理器实例
memory_cache_manager = MemoryCacheManager()

def get_cache_manager():
    """获取全局内存缓存管理器实例"""
    return memory_cache_manager 