import streamlit as st
import pandas as pd
import hashlib
import time
from pathlib import Path
from typing import Dict, Any, Optional, Tuple
import pickle
import os

class CacheManager:
    """缓存管理器，用于缓存处理过的数据，减少重复加载时间"""
    
    def __init__(self, cache_dir: str = "output/buffer"):
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.cache_metadata_file = self.cache_dir / "cache_metadata.pkl"
        self.cache_metadata = self._load_cache_metadata()
        
    def _load_cache_metadata(self) -> Dict[str, Any]:
        """加载缓存元数据"""
        if self.cache_metadata_file.exists():
            try:
                with open(self.cache_metadata_file, 'rb') as f:
                    return pickle.load(f)
            except:
                return {}
        return {}
    
    def _save_cache_metadata(self):
        """保存缓存元数据"""
        try:
            with open(self.cache_metadata_file, 'wb') as f:
                pickle.dump(self.cache_metadata, f)
        except Exception as e:
            st.warning(f"保存缓存元数据失败: {e}")
    
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
        """获取缓存的数据"""
        try:
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                return None
                
            file_size, file_mtime = self._get_file_info(file_path_obj)
            cache_key = self._generate_cache_key(str(file_path), include_self_owned_labor, file_size, file_mtime)
            
            # 检查缓存是否存在且有效
            if cache_key in self.cache_metadata:
                cache_info = self.cache_metadata[cache_key]
                cache_file = self.cache_dir / f"{cache_key}.pkl"
                
                # 检查缓存文件是否存在
                if cache_file.exists():
                    # 检查缓存是否过期（24小时）
                    if time.time() - cache_info['timestamp'] < 86400:
                        try:
                            with open(cache_file, 'rb') as f:
                                cached_data = pickle.load(f)
                                return cached_data['main_df'], cached_data['tertiary_df']
                        except Exception as e:
                            # 删除无效缓存
                            self._remove_cache(cache_key)
                    else:
                        self._remove_cache(cache_key)
            
            return None
            
        except Exception as e:
            return None
    
    def save_cached_data(self, file_path: str, include_self_owned_labor: bool, 
                        main_df: pd.DataFrame, tertiary_df: pd.DataFrame):
        """保存数据到缓存"""
        try:
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                return
                
            file_size, file_mtime = self._get_file_info(file_path_obj)
            cache_key = self._generate_cache_key(str(file_path), include_self_owned_labor, file_size, file_mtime)
            
            # 保存数据到缓存文件
            cache_file = self.cache_dir / f"{cache_key}.pkl"
            cache_data = {
                'main_df': main_df,
                'tertiary_df': tertiary_df,
                'file_path': str(file_path),
                'include_self_owned_labor': include_self_owned_labor
            }
            
            with open(cache_file, 'wb') as f:
                pickle.dump(cache_data, f)
            
            # 更新缓存元数据
            self.cache_metadata[cache_key] = {
                'file_path': str(file_path),
                'include_self_owned_labor': include_self_owned_labor,
                'file_size': file_size,
                'file_mtime': file_mtime,
                'timestamp': time.time(),
                'cache_file': str(cache_file)
            }
            
            self._save_cache_metadata()
            
        except Exception as e:
            pass
    
    def get_analysis_cache(self, cache_type: str, project_name: str, month: int, 
                          include_self_owned_labor: bool = False) -> Optional[Dict[str, Any]]:
        """获取分析结果缓存"""
        try:
            cache_key = f"{cache_type}_{project_name}_{month}_{include_self_owned_labor}"
            cache_file = self.cache_dir / f"analysis_{cache_key}.pkl"
            
            if cache_file.exists():
                # 检查缓存是否过期（24小时）
                file_mtime = cache_file.stat().st_mtime
                if time.time() - file_mtime < 86400:
                    try:
                        with open(cache_file, 'rb') as f:
                            return pickle.load(f)
                    except Exception:
                        cache_file.unlink()
            
            return None
            
        except Exception:
            return None
    
    def save_analysis_cache(self, cache_type: str, project_name: str, month: int, 
                           data: Dict[str, Any], include_self_owned_labor: bool = False):
        """保存分析结果缓存"""
        try:
            cache_key = f"{cache_type}_{project_name}_{month}_{include_self_owned_labor}"
            cache_file = self.cache_dir / f"analysis_{cache_key}.pkl"
            
            with open(cache_file, 'wb') as f:
                pickle.dump(data, f)
                
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
            if cache_key in self.cache_metadata:
                cache_file = Path(self.cache_metadata[cache_key]['cache_file'])
                if cache_file.exists():
                    cache_file.unlink()
                del self.cache_metadata[cache_key]
                self._save_cache_metadata()
        except Exception as e:
            st.warning(f"删除缓存失败: {e}")
    
    def clear_all_cache(self):
        """清除所有缓存"""
        try:
            # 删除所有缓存文件
            for cache_file in self.cache_dir.glob("*.pkl"):
                if cache_file.name != "cache_metadata.pkl":
                    cache_file.unlink()
            
            # 清空元数据
            self.cache_metadata = {}
            self._save_cache_metadata()
            st.success("🗑️ 已清除所有缓存")
            
        except Exception as e:
            st.error(f"清除缓存失败: {e}")
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """获取缓存统计信息"""
        try:
            cache_files = [f for f in self.cache_dir.glob("*.pkl") if f.name != "cache_metadata.pkl"]
            total_size = sum(f.stat().st_size for f in cache_files)
            
            return {
                'cache_count': len(cache_files),
                'total_size_mb': round(total_size / (1024 * 1024), 2),
                'metadata_count': len(self.cache_metadata)
            }
        except Exception as e:
            return {'error': str(e)}
    
    def cleanup_expired_cache(self):
        """清理过期的缓存"""
        try:
            current_time = time.time()
            expired_keys = []
            
            for cache_key, cache_info in self.cache_metadata.items():
                if current_time - cache_info['timestamp'] > 86400:  # 24小时
                    expired_keys.append(cache_key)
            
            for cache_key in expired_keys:
                self._remove_cache(cache_key)
            
            # 静默清理过期缓存
                
        except Exception as e:
            st.warning(f"清理过期缓存失败: {e}")

# 全局缓存管理器实例
cache_manager = CacheManager()

def get_cache_manager() -> CacheManager:
    """获取全局缓存管理器实例"""
    return cache_manager 