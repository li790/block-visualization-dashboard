import pandas as pd
import numpy as np
import streamlit as st
from pathlib import Path
import os
from io import BytesIO
import time
import hashlib
from utils.cache_manager import get_cache_manager

def extract_table_from_excel(file, include_self_owned_labor=False):
    """从Excel文件中提取工作表的数据:
    1. 主要费项费项月累成本使用情况(前39行) - 根据include_self_owned_labor参数选择4或4-1开头的工作表
    2. 三级费项月累表格(前153行)
    支持文件路径或文件对象，支持缓存功能
    
    Args:
        file: Excel文件路径或文件对象
        include_self_owned_labor: 是否包含自有人工成本
            - True: 使用"4主要费项费项月累成本使用情况" (包含自有人工成本)
            - False: 使用"4-1主要费项费项月累成本使用情况" (不包含自有人工成本)
    """
    # 获取缓存管理器
    cache_manager = get_cache_manager()
    
    # 如果是文件路径，先尝试从缓存获取
    if isinstance(file, str) or isinstance(file, Path):
        file_path = str(file)
        cached_data = cache_manager.get_cached_data(file_path, include_self_owned_labor)
        if cached_data:
            return cached_data
    
    try:
        # 支持文件路径或文件对象
        if isinstance(file, str) or isinstance(file, Path):
            xl = pd.ExcelFile(file)
        else:
            # 处理文件对象，包括memoryview类型
            if hasattr(file, 'getbuffer'):
                buffer = file.getbuffer()
                # 将memoryview转换为bytes，然后用BytesIO包装
                if hasattr(buffer, 'tobytes'):
                    file_bytes = buffer.tobytes()
                else:
                    file_bytes = bytes(buffer)
                # 使用BytesIO包装bytes，避免FutureWarning
                xl = pd.ExcelFile(BytesIO(file_bytes))
            else:
                st.error(f"不支持的文件对象类型: {type(file)}")
                return None, None
            
        # 根据参数选择主要费项工作表
        if include_self_owned_labor:
            main_sheet = '4主要费项费项月累成本使用情况'
            main_sheets_to_try = [main_sheet, '主要费项费项月累成本使用情况', '主要费项']
        else:
            main_sheet = '4-1主要费项费项月累成本使用情况'
            main_sheets_to_try = [main_sheet, '主要费项费项月累成本使用情况', '主要费项']
        
        tertiary_sheet = '三级费项月累表格'
        tertiary_sheets_to_try = [tertiary_sheet, '三级费项']
        
        main_df = None
        found_sheet_name = None
        
        # 首先尝试直接查找包含关键词的工作表
        for actual_sheet in xl.sheet_names:
            if include_self_owned_labor and '4主要费项' in actual_sheet and '4-1' not in actual_sheet:
                try:
                    main_df = xl.parse(actual_sheet).head(39)
                    found_sheet_name = actual_sheet
                    break
                except Exception as e:
                    continue
            elif not include_self_owned_labor and '4-1主要费项' in actual_sheet:
                try:
                    main_df = xl.parse(actual_sheet).head(39)
                    found_sheet_name = actual_sheet
                    break
                except Exception as e:
                    continue
        
        # 如果上面的方法失败，再尝试精确匹配
        if main_df is None:
            for sheet in main_sheets_to_try:
                # 检查是否有完全匹配
                exact_match = False
                for actual_sheet in xl.sheet_names:
                    if actual_sheet == sheet:
                        exact_match = True
                        found_sheet_name = actual_sheet
                        break
                    elif actual_sheet.strip() == sheet.strip():
                        exact_match = True
                        found_sheet_name = actual_sheet
                        break
                
                if exact_match and found_sheet_name:
                    try:
                        main_df = xl.parse(found_sheet_name).head(39)
                        break
                    except Exception as e:
                        continue
        
        if main_df is None:
            sheet_type = "包含自有人工成本" if include_self_owned_labor else "不包含自有人工成本"
            st.error(f"文件中未找到{sheet_type}的主要费项工作表，尝试了: {', '.join(main_sheets_to_try)}")
            st.error(f"文件中实际存在的工作表: {', '.join(xl.sheet_names)}")
            return None, None
        
        tertiary_df = None
        for sheet in tertiary_sheets_to_try:
            if sheet in xl.sheet_names:
                tertiary_df = xl.parse(sheet).head(153)
                break
        
        if tertiary_df is None:
            st.error(f"文件中未找到三级费项工作表，尝试了: {', '.join(tertiary_sheets_to_try)}")
            st.error(f"文件中实际存在的工作表: {', '.join(xl.sheet_names)}")
            return None, None
        
        # 如果是文件路径，保存到缓存
        if isinstance(file, str) or isinstance(file, Path):
            file_path = str(file)
            cache_manager.save_cached_data(file_path, include_self_owned_labor, main_df, tertiary_df)
        
        return main_df, tertiary_df
    except Exception as e:
        st.error(f"处理文件时出错: {str(e)}")
        return None, None

def get_excel_files(data_dir):
    """获取指定目录下的所有Excel文件，只返回原始文件"""
    files = []
    # 只遍历data目录，不递归子目录
    for file in data_dir.glob("*.xlsx"):
        if not file.name.startswith("~$"):
            # 排除已处理的文件（包含"主要费项"或"三级费项"的文件）
            if "_主要费项.xlsx" not in file.name and "_三级费项.xlsx" not in file.name:
                files.append(file)
    return files

# 费项类别映射表
FEE_CATEGORY_MAP = {
    '1.1.1': '人工服务',
    '1.1.2': '维护服务',
    '1.1.3': '检测服务',
    '1.2.1': '外观类',
    '1.2.2': '保洁类',
    '1.2.3': '绿化类',
    '1.2.4': '消防类',
    '1.2.5': '安防类',
    '1.2.6': '工程类',
    '1.3.1': '土建维修整改',
    '1.3.2': '电梯维修整改',
    '1.3.3': '消防维修整改',
    '1.3.4': '电气维修整改',
    '1.3.5': '暖通维修整改',
    '1.3.6': '弱电维修整改',
    '1.3.7': '给排水维修整改',
    '1.3.8': '景观维修整改',
    '1.3.9': '休闲设施整改',
    '1.3.10': '其他维修整改',
    '1.4.1': '公共水费及相关费',
    '1.4.2': '公共电费',
    '1.4.3': '公共燃气费',
    '1.4.4': '公共采暖费',
    '1.4.5': '公共其他能源费',
    '1.5.1': '打印机租赁',
    '1.5.2': '办公耗材-合计',
    '1.5.3': '房租',
    '1.5.4': '物管能源费',
    '1.5.5': '网费',
    '1.5.6': '通讯费',
    '1.5.7': '差旅费',
    '1.5.8': '交通费',
    '1.5.9': '业务招待费',
    '1.6.1': '开办物资购买',
    '1.6.2': '物业用房装修',
    '1.6.3': '物业用房开荒保洁',
    '1.6.4': '前期承诺整改',
    '1.6.5': '项目拓展费用'
}

def 补全费项类别(fee_code):
    """根据费项编码补全类别名称"""
    return FEE_CATEGORY_MAP.get(fee_code, '未知类别')

def process_excel_data(df, month, project_name=None, include_self_owned_labor=False):
    """处理Excel数据 - 适配用户表格格式，支持缓存"""
    # 确保month是整数类型
    month = int(month) if isinstance(month, str) else month
    
    # 尝试从缓存获取
    if project_name:
        cache_manager = get_cache_manager()
        cached_data = cache_manager.get_project_analysis_cache(project_name, month, include_self_owned_labor)
        if cached_data:
            return cached_data
    
    try:
        # 查找关键行
        total_target_row = None
        total_actual_row = None
        fee_items = []
        
        # 增强查找逻辑，适应合并后的数据格式
        for idx, row in df.iterrows():
            first_col = str(row.iloc[0]).strip()
            second_col = str(row.iloc[1]).strip()
            
            # 查找总成本行 - 增强版本，支持更多可能的标签
            if any(keyword in first_col for keyword in ["总成本", "年总成本", "汇总", "合计"]) or \
               any(keyword in second_col for keyword in ["总成本", "年总成本", "汇总", "合计"]):
                if any(keyword in second_col for keyword in ["月累总目标成本", "目标成本", "累计目标"]):
                    total_target_row = idx
                elif any(keyword in second_col for keyword in ["月累已发生成本", "已发生成本", "累计已发生"]):
                    total_actual_row = idx
            
            # 查找二级费项（排除累计行）
            elif ("已发生金额" in second_col or "目标金额" in second_col) and "累计" not in second_col:
                fee_name = first_col
                if fee_name not in [item['name'] for item in fee_items]:
                    fee_items.append({
                        'name': fee_name,
                        'target_row': None,
                        'actual_row': None
                    })
                
                # 找到对应的行
                for item in fee_items:
                    if item['name'] == fee_name:
                        if "目标金额" in second_col:
                            item['target_row'] = idx
                        elif "已发生金额" in second_col:
                            item['actual_row'] = idx
        
        # 安全地转换数据为数值类型
        def safe_convert_to_float(data):
            """安全地将数据转换为浮点数，处理字符串和无效值"""
            converted_data = []
            for item in data:
                if pd.isna(item):
                    converted_data.append(0.0)
                elif isinstance(item, str):
                    # 处理字符串，提取数字部分
                    import re
                    # 增强版本：支持更多格式的数字表示
                    item = item.replace(',', '').replace('，', '').strip()  # 移除千分位符号
                    number_match = re.search(r'([+-]?\d*\.?\d+)', str(item))
                    if number_match:
                        try:
                            converted_data.append(float(number_match.group(1)))
                        except ValueError:
                            converted_data.append(0.0)
                    else:
                        converted_data.append(0.0)
                else:
                    try:
                        converted_data.append(float(item))
                    except (ValueError, TypeError):
                        converted_data.append(0.0)
            return np.array(converted_data)
        
        # 直接读取月累总目标成本和月累已发生成本行的数据
        if total_target_row is not None:
            total_target_data = df.iloc[total_target_row, 2:14].values  # 1-12月数据
            total_target_data = safe_convert_to_float(total_target_data)
        else:
            # 增强版备选方案1: 使用固定行求和
            target_rows = [3, 7, 11, 15, 19, 23]
            def extract_sum_rows(rows):
                row_arrays = []
                for r in rows:
                    if 0 <= r < len(df):
                        row_values = df.iloc[r, 2:14].values
                        row_arrays.append(safe_convert_to_float(row_values))
                if not row_arrays:
                    return np.zeros(12)
                return np.sum(np.vstack(row_arrays), axis=0)
            total_target_data = extract_sum_rows(target_rows)
            
            # 增强版备选方案2: 如果备选方案1结果为全0，尝试通过费项汇总计算
            if np.all(total_target_data == 0):
                # 检查是否是合并后的数据（通常合并数据最后几行会有汇总信息）
                # 尝试从最后5行中查找可能的汇总行
                for r in range(max(0, len(df)-5), len(df)):
                    row_values = df.iloc[r, 2:14].values
                    converted = safe_convert_to_float(row_values)
                    if np.any(converted > 0):
                        total_target_data = converted
                        break

        if total_actual_row is not None:
            total_actual_data = df.iloc[total_actual_row, 2:14].values  # 1-12月数据
            total_actual_data = safe_convert_to_float(total_actual_data)
        else:
            # 增强版备选方案1: 使用固定行求和
            actual_rows = [1, 5, 9, 13, 17, 21]
            def extract_sum_rows(rows):
                row_arrays = []
                for r in rows:
                    if 0 <= r < len(df):
                        row_values = df.iloc[r, 2:14].values
                        row_arrays.append(safe_convert_to_float(row_values))
                if not row_arrays:
                    return np.zeros(12)
                return np.sum(np.vstack(row_arrays), axis=0)
            total_actual_data = extract_sum_rows(actual_rows)
            
            # 增强版备选方案2: 如果备选方案1结果为全0，尝试通过费项汇总计算
            if np.all(total_actual_data == 0):
                # 检查是否是合并后的数据（通常合并数据最后几行会有汇总信息）
                # 尝试从最后5行中查找可能的汇总行
                for r in range(max(0, len(df)-5), len(df)):
                    row_values = df.iloc[r, 2:14].values
                    converted = safe_convert_to_float(row_values)
                    if np.any(converted > 0):
                        total_actual_data = converted
                        break

        # 确保即使在极端情况下也有有效数据
        if np.all(total_target_data == 0) and np.all(total_actual_data == 0):
            # 作为最后的备选方案，尝试直接计算费项数据的总和
            for item in fee_items:
                if item['target_row'] is not None and item['actual_row'] is not None:
                    target_data = df.iloc[item['target_row'], 2:14].values
                    actual_data = df.iloc[item['actual_row'], 2:14].values
                    target_data = safe_convert_to_float(target_data)
                    actual_data = safe_convert_to_float(actual_data)
                    total_target_data += target_data
                    total_actual_data += actual_data
        
        # 根据新要求修改计算逻辑：
        # 1. 累计总目标成本 = 原始表格 总成本 月累总目标成本
        # 2. 累计总已发生成本 = 原始表格 总成本 月累已发生成本
        # 3. 年使用率 = 用户选择月份的累计总已发生成本 / 12月份的累计总目标成本
        # 4. 月累使用率 = 用户选择月份的累计总已发生成本 / 用户选择月份的累计总目标成本
        # 5. 年总目标 = 12月份的累计总目标成本
        # 6. 月累目标 = 用户选择月份的累计总目标成本
        # 7. 月累已发生 = 用户选择月份的累计总已发生成本
        
        # 计算1-12月累计值
        cum_target = total_target_data  # 累计总目标成本
        cum_actual = total_actual_data  # 累计总已发生成本

        # 针对部分项目12月未填数据（或尾部月份为空）的情况，
        # 对累计序列做向前填充：将连续尾部的0替换为前一个非零累计值。
        def _forward_fill_cum(arr):
            filled = arr.copy()
            last_val = 0.0
            for i in range(len(filled)):
                try:
                    val = float(filled[i])
                except Exception:
                    val = 0.0
                if val != 0.0:
                    last_val = val
                else:
                    # 仅当已有累计值大于0时进行填充
                    if last_val > 0.0:
                        filled[i] = last_val
            return filled

        cum_target = _forward_fill_cum(cum_target)
        cum_actual = _forward_fill_cum(cum_actual)
        
        # 增强年总目标计算逻辑，确保在多项目合并时也能正确计算
        # 年总目标 = 12月份的累计总目标成本
        # 额外检查：如果最后一个月数据为0，尝试使用所有月份的最大值或总和作为备选
        year_cum_target = cum_target[-1]  # 第12个月的数据
        
        # 如果第12个月数据为0，检查是否是合并数据的特殊情况
        if year_cum_target == 0:
            # 备选方案1: 使用所有月份的最大值（可能是数据格式问题）
            if np.max(cum_target) > 0:
                year_cum_target = np.max(cum_target)
            # 备选方案2: 如果所有月份都为0，尝试通过总目标计算
            elif total_target_data.sum() > 0:
                # 假设月度目标是均匀分布的，估算年总目标
                year_cum_target = total_target_data.sum() * (12 / np.count_nonzero(total_target_data)) if np.count_nonzero(total_target_data) > 0 else total_target_data.sum()
        
        year_cum_actual = cum_actual[-1]  # 第12个月的数据
        
        # 万元换算
        year_cum_target_wy = round(year_cum_target / 10000, 2)
        year_cum_actual_wy = round(year_cum_actual / 10000, 2)
        
        # 使用率计算（根据新要求）
        year_usage = round(100 * cum_actual[month-1] / year_cum_target, 2) if year_cum_target else 0
        month_usage = round(100 * cum_actual[month-1] / cum_target[month-1], 2) if cum_target[month-1] else 0
        # 确保month是整数类型
        month_int = int(month) if isinstance(month, str) else month
        time_progress = round(100 * month_int / 12, 2)
        
        # 调试信息已移除
        
        # 处理二级费项数据和异常项
        processed_fee_items = []
        exceptions = []
        for item in fee_items:
            if item['target_row'] is not None and item['actual_row'] is not None:
                target_data = df.iloc[item['target_row'], 2:14].values
                actual_data = df.iloc[item['actual_row'], 2:14].values
                
                # 安全转换数据
                target_data = safe_convert_to_float(target_data)
                actual_data = safe_convert_to_float(actual_data)
                
                # 计算累计值
                cum_target_item = np.cumsum(target_data)
                cum_actual_item = np.cumsum(actual_data)
                
                # 当前月份的累计值
                current_cum_target = float(cum_target_item[month-1])
                current_cum_actual = float(cum_actual_item[month-1])
                
                processed_fee_items.append({
                    'name': item['name'],
                    'cum_target': current_cum_target,
                    'cum_actual': current_cum_actual
                })
                
                # 异常项检测
                # 12月份的目标金额累计
                year_target_item = cum_target_item[-1]
                
                # 按月检查异常
                for m in range(1, 13):
                    month_cum_target = cum_target_item[m-1]
                    month_cum_actual = cum_actual_item[m-1]
                    
                    exception_type = None
                    if month_cum_actual > year_target_item:
                        exception_type = 'red'
                    elif month_cum_actual > month_cum_target:
                        exception_type = 'yellow'
                    
                    if exception_type and m <= month:  # 只检查到当前选择的月份
                        exceptions.append({
                            'fee_name': item['name'],
                            'month': m,
                            'exception_type': exception_type,
                            'cum_actual': float(month_cum_actual),
                            'cum_target': float(month_cum_target),
                            'year_target': float(year_target_item)
                        })
        
        result = {
            'total_target': float(total_target_data.sum()),
            'cum_target': [float(x) for x in cum_target],
            'cum_actual': [float(x) for x in cum_actual],
            'year_cum_target_wy': year_cum_target_wy,
            'year_cum_actual_wy': year_cum_actual_wy,
            'year_usage': year_usage,
            'month_usage': month_usage,
            'time_progress': time_progress,
            'fee_items': processed_fee_items,
            'exceptions': exceptions,
            # 添加月累目标和月累已发生字段（万元）
            'month_cum_target_wy': round(cum_target[month-1] / 10000, 2) if month-1 < len(cum_target) else 0,
            'month_cum_actual_wy': round(cum_actual[month-1] / 10000, 2) if month-1 < len(cum_actual) else 0
        }
        
        # 保存到缓存
        if project_name:
            cache_manager = get_cache_manager()
            cache_manager.save_project_analysis_cache(project_name, month, result, include_self_owned_labor)
        
        return result
        
    except Exception as e:
        st.error(f"处理Excel文件时出错: {e}")
        return None

def load_and_process_files(uploaded_files, selected_files, data_dir, month, include_self_owned_labor=False):
    """加载并处理多个文件，数据需拆解处理
    data_dir: 存放原始数据的目录
    include_self_owned_labor: 是否包含自有人工成本
    注意：已移除output_dir依赖，现在使用内存缓存
    """
    all_data = {}
    all_main_dfs = {}
    all_tertiary_dfs = {}
    
    # 注意：已移除文件系统缓存，现在使用内存缓存
    
    # 处理上传的文件（这些文件已经在sidebar中处理过了，这里只需要处理已提取的数据）
    if uploaded_files:
        for i, uploaded_file in enumerate(uploaded_files):
            try:
                # 使用上传文件的真实名称（去掉扩展名），加上时间戳保证唯一
                uploaded_filename = uploaded_file.name.replace('.xlsx', '').replace('.xls', '')
                unique_id = time.strftime("%Y%m%d%H%M%S") + f"_{i}"
                unique_filename = f"{uploaded_filename}_{unique_id}"

                # 直接从上传文件中提取数据
                main_df, tertiary_df = extract_table_from_excel(uploaded_file, include_self_owned_labor)
                if main_df is None or tertiary_df is None:
                    st.warning(f"无法从文件 {uploaded_file.name} 中提取有效数据")
                    continue

                # 注意：不再保存到文件系统，数据只在内存中处理
                # 处理数据
                all_main_dfs[unique_filename] = main_df
                all_tertiary_dfs[unique_filename] = tertiary_df

                # 处理主要费项数据
                data = process_excel_data(main_df, month, unique_filename, include_self_owned_labor)
                if data:
                    all_data[unique_filename] = data

                    # 处理三级费项数据
                    tertiary_result = process_tertiary_fee_data(tertiary_df, month, unique_filename, include_self_owned_labor)
                    data['tertiary_fee_items'] = tertiary_result['tertiary_fee_items']
                    data['tertiary_exceptions'] = tertiary_result['exceptions']

                # 注意：不再保存分析结果到文件系统，数据只在内存中处理
                st.success(f"成功处理文件: {uploaded_file.name}")
            except Exception as e:
                st.error(f"处理上传文件 {uploaded_file.name} 时出错: {e}")
    
    # 处理选中的文件（从data目录读取原始文件，然后处理）
    if selected_files:
        st.write(f"**正在处理选中的 {len(selected_files)} 个文件...**")
        for filename in selected_files:
            try:
                # 检查是否是已处理的文件（包含"主要费项"或"三级费项"的文件）
                if "_主要费项.xlsx" in filename or "_三级费项.xlsx" in filename:
                    # 跳过已处理的文件，只处理原始文件
                    continue
                    
                # 处理原始Excel文件
                file_path = data_dir / filename
                main_df, tertiary_df = extract_table_from_excel(file_path, include_self_owned_labor)
                if main_df is None or tertiary_df is None:
                    st.warning(f"无法从文件 {filename} 中提取有效数据")
                    continue
                
                project_name = filename.replace('.xlsx', '')
                all_main_dfs[project_name] = main_df
                all_tertiary_dfs[project_name] = tertiary_df
                
                # 处理主要费项数据
                data = process_excel_data(main_df, month, project_name, include_self_owned_labor)
                if data:
                    all_data[project_name] = data
                    
                    # 处理三级费项数据
                    tertiary_result = process_tertiary_fee_data(tertiary_df, month, project_name, include_self_owned_labor)
                    data['tertiary_fee_items'] = tertiary_result['tertiary_fee_items']
                    data['tertiary_exceptions'] = tertiary_result['exceptions']
                
                # 注意：不再保存分析结果到文件系统，数据只在内存中处理
            except Exception as e:
                st.error(f"处理文件 {filename} 时出错: {e}")
    
    return all_data, all_main_dfs, all_tertiary_dfs

def create_tertiary_exception_table(all_data):
    """生成三级费项异常统计表"""
    if not all_data:
        return None
    
    exception_list = []
    
    # 遍历所有项目的异常数据
    for project_name, data in all_data.items():
        if 'tertiary_exceptions' in data:
            for exception in data['tertiary_exceptions']:
                # 添加项目名称到异常记录
                exception_record = exception.copy()
                exception_record['project_name'] = project_name
                exception_list.append(exception_record)
    
    # 转换为DataFrame
    if exception_list:
        exception_df = pd.DataFrame(exception_list)
        # 选择并排序列
        columns_order = ['project_name', 'fee_code', 'fee_name', 'month', 'cum_actual', 'cum_target', 'year_target', 'exception_type']
        exception_df = exception_df[columns_order]
        # 格式化数值列，保留2位小数
        numeric_columns = ['cum_actual', 'cum_target', 'year_target']
        exception_df[numeric_columns] = exception_df[numeric_columns].map(lambda x: f"{x:.2f}")
        # 替换异常类型为中文
        exception_df['exception_type'] = exception_df['exception_type'].replace({'red': '超年度目标', 'yellow': '超月度目标'})
        return exception_df
    
    return None

def create_summary_excel(all_dfs):
    """创建真正的汇总Excel表 - 参考单项目处理方式，对固定行进行专门求和处理"""
    if not all_dfs:
        return None
    
    # 获取第一个DataFrame作为模板
    first_df = list(all_dfs.values())[0]
    
    # 创建新的汇总DataFrame，保持相同的结构
    summary_df = first_df.copy()
    
    # 初始化汇总数据为0
    for col in summary_df.columns:
        if col not in [summary_df.columns[0], summary_df.columns[1]]:  # 跳过前两列（费项名称和类型）
            summary_df[col] = 0
    
    # 创建一个字典用于存储每个项目的关键行索引
    project_key_rows = {}
    
    # 第一步：为每个项目识别关键行（参考process_excel_data中的查找逻辑）
    for project_name, df in all_dfs.items():
        target_rows = []  # 存储目标金额相关行索引
        actual_rows = []  # 存储已发生金额相关行索引
        
        for idx, row in df.iterrows():
            first_col = str(row.iloc[0]).strip()
            second_col = str(row.iloc[1]).strip()
            
            # 查找关键行 - 参考process_excel_data中的增强查找逻辑
            if any(keyword in first_col for keyword in ["总成本", "年总成本", "汇总", "合计"]) or \
               any(keyword in second_col for keyword in ["总成本", "年总成本", "汇总", "合计"]):
                if any(keyword in second_col for keyword in ["月累总目标成本", "目标成本", "累计目标", "目标金额"]):
                    target_rows.append(idx)
                elif any(keyword in second_col for keyword in ["月累已发生成本", "已发生成本", "累计已发生", "已发生金额"]):
                    actual_rows.append(idx)
            
            # 查找二级费项相关行
            elif ("已发生金额" in second_col or "目标金额" in second_col) and "累计" not in second_col:
                if "目标金额" in second_col:
                    target_rows.append(idx)
                elif "已发生金额" in second_col:
                    actual_rows.append(idx)
        
        # 存储当前项目的关键行信息
        project_key_rows[project_name] = {
            'target_rows': target_rows,
            'actual_rows': actual_rows
        }
    
    # 第二步：对关键行进行专门的求和处理，并对其他行进行正常合并
    for project_name, df in all_dfs.items():
        key_rows = project_key_rows[project_name]
        
        for col in df.columns:
            if col not in [df.columns[0], df.columns[1]]:  # 跳过前两列
                # 确保数据类型为数值型
                numeric_data = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
                # 对所有行（包括关键行）进行累加
                summary_df[col] += numeric_data
    
    # 第三步：仅在存在明确的项目名称列时才覆盖；否则新增一列而不影响月份数据
    name_like_cols = [c for c in summary_df.columns if any(k in str(c) for k in ["项目名称", "项目", "文件名", "Excel文件名", "Excel 文件名"]) ]
    if name_like_cols:
        for c in name_like_cols:
            summary_df[c] = "汇总"
    else:
        summary_df["项目"] = "汇总"
    
    return summary_df

def process_tertiary_fee_data(df, month, project_name=None, include_self_owned_labor=False):
    """处理三级费项数据并检测异常，支持缓存"""
    # 确保month是整数类型
    month = int(month) if isinstance(month, str) else month
    
    # 尝试从缓存获取
    if project_name:
        cache_manager = get_cache_manager()
        cached_data = cache_manager.get_anomaly_cache(project_name, month, include_self_owned_labor)
        if cached_data:
            return cached_data
    
    try:
        tertiary_fee_items = []
        exceptions = []
        
        # 动态检测列索引，避免硬编码导致的问题
        max_columns = df.shape[1]
        
        # 根据实际表格结构：数据列从第3列开始（索引2），对应1-12月
        min_required_columns = 14  # 至少需要14列（费项编码 + 数据类型 + 12个月数据）
        if max_columns < min_required_columns:
            st.error(f"三级费项表格列数不足，仅找到{max_columns}列，需要至少{min_required_columns}列")
            return {'tertiary_fee_items': [], 'exceptions': []}
        
        # 数据列范围：第3列到第14列（索引2-13），对应1-12月
        data_columns = list(range(2, min(14, max_columns)))
        
        # 用于存储每个费项的4行数据
        fee_data = {}
        
        # 第一遍扫描：收集每个费项的所有数据
        current_fee_code = None
        current_fee_data = {}
        
        for idx, row in df.iterrows():
            # 第一列是费项编码
            fee_code = str(row.iloc[0]).strip()
            
            # 检查是否是有效的三级费项编码
            if len(fee_code.split('.')) == 3 and fee_code.replace('.', '').isdigit():
                # 新的费项编码
                if current_fee_code and current_fee_code != fee_code:
                    # 保存前一个费项的数据
                    if len(current_fee_data) == 4:  # 确保有完整的4行数据
                        fee_data[current_fee_code] = current_fee_data
                    current_fee_data = {}
                
                current_fee_code = fee_code
                second_col = str(row.iloc[1]).strip()
                
                # 提取该行的数据
                row_data = []
                for col in data_columns:
                    try:
                        cell_value = row.iloc[col]
                        if isinstance(cell_value, str) and ('已发生金额' in cell_value or '目标金额' in cell_value):
                            val = 0
                        else:
                            val = float(cell_value) if pd.notna(cell_value) else 0
                        row_data.append(val)
                    except ValueError:
                        row_data.append(0)
                
                # 根据第二列内容确定数据类型
                if '已发生金额' in second_col and '累计' not in second_col:
                    current_fee_data['monthly_actual'] = row_data
                elif '已发生金额' in second_col and '累计' in second_col:
                    current_fee_data['cum_actual'] = row_data
                elif '目标金额' in second_col and '累计' not in second_col:
                    current_fee_data['monthly_target'] = row_data
                elif '目标金额' in second_col and '累计' in second_col:
                    current_fee_data['cum_target'] = row_data
        
        # 保存最后一个费项的数据
        if current_fee_code and len(current_fee_data) == 4:
            fee_data[current_fee_code] = current_fee_data
        
        # 第二遍扫描：处理数据并检测异常
        for fee_code, data in fee_data.items():
            try:
                # 补全费项类别名称
                fee_name = 补全费项类别(fee_code)
                
                # 获取各种数据
                monthly_target = data.get('monthly_target', [0] * 12)
                monthly_actual = data.get('monthly_actual', [0] * 12)
                cum_target = data.get('cum_target', [0] * 12)
                cum_actual = data.get('cum_actual', [0] * 12)
                
                # 确保数据长度一致
                min_length = min(len(monthly_target), len(monthly_actual), len(cum_target), len(cum_actual))
                monthly_target = monthly_target[:min_length]
                monthly_actual = monthly_actual[:min_length]
                cum_target = cum_target[:min_length]
                cum_actual = cum_actual[:min_length]
                
                # 年总目标（12月份累计目标）
                year_target = cum_target[-1] if len(cum_target) > 0 else 0
                
                # 当前月份数据
                current_target = cum_target[month-1] if month-1 < len(cum_target) else 0
                current_actual = cum_actual[month-1] if month-1 < len(cum_actual) else 0
                
                # 数据验证：即使年度目标为0也进行统计
                if year_target == 0:
                    # 年度目标为0时，仍然添加到统计中，但不单独添加到异常列表
                    tertiary_fee_items.append({
                        'code': fee_code,
                        'name': fee_name,
                        'cum_target': current_target,
                        'cum_actual': current_actual,
                        'monthly_data': {
                            'target': monthly_target,
                            'actual': monthly_actual,
                            'cum_target': cum_target,
                            'cum_actual': cum_actual
                        }
                    })
                    # 继续执行异常检测逻辑，如果触发红色或黄色异常才会添加到异常列表
                
                tertiary_fee_items.append({
                    'code': fee_code,
                    'name': fee_name,
                    'cum_target': current_target,
                    'cum_actual': current_actual,
                    'monthly_data': {
                        'target': monthly_target,
                        'actual': monthly_actual,
                        'cum_target': cum_target,
                        'cum_actual': cum_actual
                    }
                })
                
                # 异常检测 - 检查到当前选择的月份
                for m in range(1, month+1):
                    # 确保索引在有效范围内
                    if m-1 >= len(cum_target) or m-1 >= len(cum_actual):
                        continue
                        
                    m_target = cum_target[m-1]
                    m_actual = cum_actual[m-1]
                    
                    exception_type = None
                    # 红色异常：累计已发生金额超过年度总目标
                    if m_actual > year_target:
                        exception_type = 'red'
                    # 黄色异常：累计已发生金额超过该月累计目标
                    elif m_actual > m_target:
                        exception_type = 'yellow'
                    
                    if exception_type:
                        exceptions.append({
                            'fee_code': fee_code,
                            'fee_name': fee_name,
                            'month': m,
                            'exception_type': exception_type,
                            'cum_actual': m_actual,
                            'cum_target': m_target,
                            'year_target': year_target
                        })
                        
            except Exception as e:
                st.warning(f"处理三级费项 {fee_code} 时出错: {e}")
                continue
        
        # 添加异常信息到返回结果
        result = {
            'tertiary_fee_items': tertiary_fee_items,
            'exceptions': exceptions
        }
        
        # 保存到缓存
        if project_name:
            cache_manager = get_cache_manager()
            cache_manager.save_anomaly_cache(project_name, month, result, include_self_owned_labor)
        
        return result
    except Exception as e:
        st.error(f"处理三级费项数据时出错: {e}")
        return {'tertiary_fee_items': [], 'exceptions': []}

def merge_project_data(all_data, all_main_dfs, month, include_self_owned_labor=False):
    """合并多个项目的数据并计算合并后的关键指标，支持缓存"""
    if not all_data or not all_main_dfs:
        return None
    
    # 尝试从缓存获取
    cache_manager = get_cache_manager()
    project_hash = hashlib.md5(str(sorted(all_data.keys())).encode()).hexdigest()
    cached_data = cache_manager.get_project_analysis_cache(f"merged_{project_hash}", month, include_self_owned_labor)
    if cached_data:
        return cached_data
    
    # 方法1: 先尝试合并原始Excel数据
    merged_df = create_summary_excel(all_main_dfs)
    merged_data = None
    if merged_df is not None and not merged_df.empty:
        merged_data = process_excel_data(merged_df, month, f"merged_{project_hash}", include_self_owned_labor)
    
    # 方法2: 直接对各项目数据进行精确汇总（作为备用方案）
    # 这将确保即使合并Excel数据失败，也能通过汇总各项目数据得到准确结果
    if not merged_data or merged_data['year_cum_target_wy'] == 0:
        # 初始化合并结果
        merged_data = {
            'total_target': 0.0,
            'cum_target': [0.0] * 12,
            'cum_actual': [0.0] * 12,
            'year_cum_target_wy': 0.0,
            'year_cum_actual_wy': 0.0,
            'year_usage': 0.0,
            'month_usage': 0.0,
            'time_progress': 0.0,
            'fee_items': [],
            'exceptions': [],
            'merged_projects': list(all_data.keys())
        }
        
        # 汇总各项目的已发生金额、目标金额等数据
        valid_projects = 0
        for project_name, project_data in all_data.items():
            # 确保项目数据有效
            if not project_data or not isinstance(project_data, dict):
                continue
            
            # 汇总总目标成本和已发生成本
            if 'total_target' in project_data:
                merged_data['total_target'] += project_data['total_target']
            
            # 汇总累计目标成本（1-12月）
            if 'cum_target' in project_data and isinstance(project_data['cum_target'], list) and len(project_data['cum_target']) >= 12:
                for i in range(12):
                    merged_data['cum_target'][i] += project_data['cum_target'][i]
            
            # 汇总累计已发生成本（1-12月）
            if 'cum_actual' in project_data and isinstance(project_data['cum_actual'], list) and len(project_data['cum_actual']) >= 12:
                for i in range(12):
                    merged_data['cum_actual'][i] += project_data['cum_actual'][i]
            
            # 汇总年总目标（万元）
            if 'year_cum_target_wy' in project_data:
                merged_data['year_cum_target_wy'] += project_data['year_cum_target_wy']
            
            # 汇总年已发生（万元）
            if 'year_cum_actual_wy' in project_data:
                merged_data['year_cum_actual_wy'] += project_data['year_cum_actual_wy']
            
            valid_projects += 1
        
        # 计算使用率等指标（基于汇总数据）
        if valid_projects > 0:
            # 计算时间进度
            month_int = int(month) if isinstance(month, str) else month
            merged_data['time_progress'] = round(100 * month_int / 12, 2)
            
            # 基于累计目标数组动态计算年总目标（取最后一个非零累计值，若都为0则取0）
            def _last_nonzero_value(values):
                for v in reversed(values):
                    if v and float(v) != 0.0:
                        return float(v)
                return 0.0
            calculated_year_target = _last_nonzero_value(merged_data['cum_target'])
            if calculated_year_target > 0:
                merged_data['year_cum_target_wy'] = round(calculated_year_target / 10000, 2)

            # 计算年使用率和月累使用率（避免除零错误）
            current_month = month_int - 1  # 转为0-based索引
            if merged_data['year_cum_target_wy'] > 0 and current_month >= 0 and current_month < 12:
                # 年使用率 = 当前月累计已发生 / 年总目标
                merged_data['year_usage'] = round(100 * merged_data['cum_actual'][current_month] / (merged_data['year_cum_target_wy'] * 10000), 2)
                
                # 月累使用率 = 当前月累计已发生 / 当前月累计目标
                if merged_data['cum_target'][current_month] > 0:
                    merged_data['month_usage'] = round(100 * merged_data['cum_actual'][current_month] / merged_data['cum_target'][current_month], 2)
            
            # 添加月累目标和月累已发生字段（万元）
            if current_month >= 0 and current_month < 12:
                merged_data['month_cum_target_wy'] = round(merged_data['cum_target'][current_month] / 10000, 2)
                merged_data['month_cum_actual_wy'] = round(merged_data['cum_actual'][current_month] / 10000, 2)
            else:
                merged_data['month_cum_target_wy'] = 0
                merged_data['month_cum_actual_wy'] = 0
    
    # 确保项目列表信息被添加
    if merged_data and 'merged_projects' not in merged_data:
        merged_data['merged_projects'] = list(all_data.keys())
    
    # 保存到缓存
    if merged_data:
        cache_manager.save_project_analysis_cache(f"merged_{project_hash}", month, merged_data, include_self_owned_labor)
    
    return merged_data 

def create_monthly_fee_summary(all_dfs):
    """创建每月费项汇总表 - 统计所有项目每月各项费项合计的已发生成本和目标成本"""
    if not all_dfs:
        return None
    
    # 初始化月度汇总数据
    monthly_data = {}
    
    for project_name, df in all_dfs.items():
        # 查找二级费项数据
        for idx, row in df.iterrows():
            first_col = str(row.iloc[0]).strip()
            second_col = str(row.iloc[1]).strip()
            
            # 跳过带有序号的行（如1.1.1）和总成本行
            if ("1.1." in first_col or "总成本" in first_col or "年总成本" in first_col):
                continue
            
            # 跳过纯数字编号的行
            if first_col.replace('.', '').replace(' ', '').isdigit():
                continue
            
            # 处理目标金额
            if "目标金额" in second_col and "累计" not in second_col:
                fee_name = first_col
                target_data = df.iloc[idx, 2:14].values  # 1-12月数据
                
                # 查找对应的已发生金额行
                actual_data = None
                for sub_idx, sub_row in df.iterrows():
                    sub_fee_name = str(sub_row.iloc[0]).strip()
                    sub_second_col = str(sub_row.iloc[1]).strip()
                    
                    if (sub_fee_name == fee_name and 
                        "已发生金额" in sub_second_col and 
                        "累计" not in sub_second_col):
                        actual_data = df.iloc[sub_idx, 2:14].values
                        break
                
                if actual_data is not None:
                    # 按月份累加数据
                    for month in range(12):
                        month_key = month + 1  # 1-12月
                        if month_key not in monthly_data:
                            monthly_data[month_key] = {'target': 0, 'actual': 0}
                        
                        # 转换为数字并累加
                        try:
                            target_val = float(target_data[month]) if pd.notna(target_data[month]) else 0
                            actual_val = float(actual_data[month]) if pd.notna(actual_data[month]) else 0
                            monthly_data[month_key]['target'] += target_val
                            monthly_data[month_key]['actual'] += actual_val
                        except Exception as e:
                            continue
                else:
                    pass  # 静默处理，不显示警告
    
    # 转换为DataFrame
    months = sorted(monthly_data.keys())
    data = {
        '月份': [f'{m}月' for m in months],
        '目标成本': [monthly_data[m]['target']/10000 for m in months],  # 转换为万元
        '已发生成本': [monthly_data[m]['actual']/10000 for m in months]  # 转换为万元
    }
    
    # 移除调试信息
    
    return pd.DataFrame(data)

def create_client_download_table(all_dfs, all_data): 
    # 初始化所有项目的累计数据
    total_monthly_target_sum = np.zeros(12, dtype=float)
    total_monthly_actual_sum = np.zeros(12, dtype=float)
    total_total_target_array = np.zeros(12, dtype=float)
    total_total_actual_array = np.zeros(12, dtype=float)
    
    for project_name, df in all_dfs.items():
        # 获取项目数据
        project_data = all_data[project_name]
        
        # 提取原始数据
        total_target_data = None
        total_actual_data = None
        fee_items_data = []
        
        # 查找总成本行
        for idx, row in df.iterrows():
            first_col = str(row.iloc[0]).strip()
            second_col = str(row.iloc[1]).strip()
            
            if "总成本" in first_col or "年总成本" in first_col:
                if "月累总目标成本" in second_col:
                    total_target_data = df.iloc[idx, 2:14].values  # 1-12月
                elif "月累已发生成本" in second_col:
                    total_actual_data = df.iloc[idx, 2:14].values  # 1-12月
            
            # 查找二级费项 - 排除带有序号的行（如1.1.1）
            elif ("已发生金额" in second_col or "目标金额" in second_col) and "累计" not in second_col:
                fee_name = first_col.strip()
                
                # 跳过带有序号的行（如1.1.1、1.1.2等）
                if any(char.isdigit() and char != '0' for char in fee_name.split('.')[0] if '.' in fee_name):
                    continue
                
                # 跳过纯数字编号的行
                if fee_name.replace('.', '').replace(' ', '').isdigit():
                    continue
                
                if "目标金额" in second_col:
                    target_data = df.iloc[idx, 2:14].values
                    # 查找对应的已发生金额行
                    for sub_idx, sub_row in df.iterrows():
                        sub_fee_name = str(sub_row.iloc[0]).strip()
                        sub_second_col = str(sub_row.iloc[1]).strip()
                        
                        # 确保是同一个费项且是已发生金额
                        if (sub_fee_name == fee_name and \
                            "已发生金额" in sub_second_col and \
                            "累计" not in sub_second_col):
                            actual_data = df.iloc[sub_idx, 2:14].values
                            fee_items_data.append({
                                'name': fee_name,
                                'target': target_data,
                                'actual': actual_data
                            })
                            break
        
        # 计算每月各费项之和
        monthly_target_sum = np.zeros(12, dtype=float)
        monthly_actual_sum = np.zeros(12, dtype=float)
        
        for item in fee_items_data:
            # 确保数据类型转换
            target_data = np.array(item['target'], dtype=float)
            monthly_target_sum += target_data
            actual_data = np.array(item['actual'], dtype=float)
            monthly_actual_sum += actual_data
        
        # 检查是否找到了总成本数据
        if total_target_data is None or total_actual_data is None:
            continue
            
        # 计算累计数据
        total_target_array = np.array(total_target_data, dtype=float)
        total_actual_array = np.array(total_actual_data, dtype=float)
        
        # 累加到总数据
        total_monthly_target_sum += monthly_target_sum
        total_monthly_actual_sum += monthly_actual_sum
        total_total_target_array += total_target_array
        total_total_actual_array += total_actual_array
    
    # 重置result_data，确保只包含汇总数据
    result_data = []
    
    # 为合并后的数据添加4行数据，月份横向排布
    # 行1: 已发生金额（每月各费项成本已发生金额之和）
    row1_data = {
        '数据类型': '已发生金额'
    }
    # 添加1-12月数据
    for month in range(1, 13):
        row1_data[f'{month}月'] = float(total_monthly_actual_sum[month-1])
    result_data.append(row1_data)
    
    # 行2: 月累已发生金额（年总成本月累已发生金额）
    row2_data = {
        '数据类型': '月累已发生金额'
    }
    # 添加1-12月数据
    for month in range(1, 13):
        row2_data[f'{month}月'] = float(total_total_actual_array[month-1])
    result_data.append(row2_data)
    
    # 行3: 目标金额（每月各费项目标金额之和）
    row3_data = {
        '数据类型': '目标金额'
    }
    # 添加1-12月数据
    for month in range(1, 13):
        row3_data[f'{month}月'] = float(total_monthly_target_sum[month-1])
    result_data.append(row3_data)
    
    # 行4: 目标金额累计（年总成本月累目标金额）
    row4_data = {
        '数据类型': '目标金额累计'
    }
    # 添加1-12月数据
    for month in range(1, 13):
        row4_data[f'{month}月'] = float(total_total_target_array[month-1])
    result_data.append(row4_data)
    
    # 创建DataFrame
    result_df = pd.DataFrame(result_data)
    
    # 重新排列列顺序：数据类型、1月、2月、...、12月（移除累计金额列）
    columns = ['数据类型'] + [f'{i}月' for i in range(1, 13)]
    result_df = result_df[columns]
    
    return result_df

def create_secondary_fee_overall_data(all_main_dfs, month=None, include_self_owned_labor=False):
    """创建二级费项整体数据，用于组合图表展示，支持缓存
    
    Args:
        all_main_dfs: 所有项目的主要费项数据字典
        month: 月份，用于缓存键
        include_self_owned_labor: 是否包含自有人工成本
        
    Returns:
        list: 包含各二级费项整体数据的列表
    """
    # 尝试从缓存获取
    if month is not None:
        cache_manager = get_cache_manager()
        # 使用项目名称的哈希作为缓存键
        project_hash = hashlib.md5(str(sorted(all_main_dfs.keys())).encode()).hexdigest()
        cached_data = cache_manager.get_secondary_fee_cache(f"combined_{project_hash}", month, include_self_owned_labor)
        if cached_data is not None:
            print(f"二级费项数据从缓存加载成功，项目哈希: {project_hash}, 月份: {month}")
            return cached_data
        else:
            print(f"未找到二级费项缓存，项目哈希: {project_hash}, 月份: {month}")
    secondary_fee_data = {}

    import re

    def to_float_array(values):
        """将任意序列安全转换为float数组，容错逗号、中文逗号、空格、破折号等"""
        cleaned: list[float] = []
        for v in list(values):
            if pd.isna(v):
                cleaned.append(0.0)
                continue
            if isinstance(v, (int, float, np.number)):
                cleaned.append(float(v))
                continue
            s = str(v).strip()
            # 常见非数值表示视为0
            if s in {"-", "—", "--", "——", "N/A", "", "None"}:
                cleaned.append(0.0)
                continue
            # 去除千分位、中文逗号、空格
            s = s.replace(",", "").replace("，", "").replace(" ", "")
            # 去除末尾的非数字/小数点/负号字符（如单位等）
            s = re.sub(r"[^0-9.\-]+$", "", s)
            try:
                cleaned.append(float(s))
            except Exception:
                cleaned.append(0.0)
        return np.array(cleaned, dtype=float)
    
    for project_name, df in all_main_dfs.items():
        if df is None:
            continue
            
        # 查找二级费项数据
        fee_items_data = []
        
        # 首先收集所有费项名称
        fee_names = set()
        for idx, row in df.iterrows():
            first_col = str(row.iloc[0]).strip()
            second_col = str(row.iloc[1]).strip()
            
            # 查找二级费项（排除累计行）
            if (("已发生金额" in second_col or "已发生成本" in second_col or "已发生" in second_col) or
                ("目标金额" in second_col or "目标成本" in second_col or "目标" in second_col)) and ("累计" not in second_col):
                fee_name = first_col
                
                # 跳过纯数字编号的行
                if fee_name.replace('.', '').replace(' ', '').isdigit():
                    continue
                
                fee_names.add(fee_name)
        
        # 调试信息：打印找到的费项名称
        if "物耗成本" in fee_names:
            print(f"项目 {project_name} 中找到物耗成本费项")
        else:
            # 检查是否有类似的名称
            for name in fee_names:
                if "物耗" in name:
                    print(f"项目 {project_name} 中找到类似物耗成本的费项: {name}")
        
        # 对每个费项名称，查找其目标金额和已发生金额
        for fee_name in fee_names:
            target_data = None
            actual_data = None
            
            # 查找目标金额数据
            for idx, row in df.iterrows():
                first_col = str(row.iloc[0]).strip()
                second_col = str(row.iloc[1]).strip()
                
                # 更宽松的匹配条件
                if (first_col == fee_name or first_col.strip() == fee_name.strip()) and \
                   ("目标金额" in second_col or "目标成本" in second_col or "目标" in second_col) and \
                   "累计" not in second_col:
                    try:
                        raw = df.iloc[idx, 2:14].values  # 1-12月
                        # 确保数据类型正确（容错千分位/中文符号）
                        target_data = to_float_array(raw)
                        if fee_name == "物耗成本" or "物耗" in fee_name:
                            print(f"物耗成本目标数据: {target_data}")
                    except Exception as e:
                        print(f"处理目标金额数据时出错: {e}")
                        target_data = None
                    break
            
            # 查找已发生金额数据
            for idx, row in df.iterrows():
                first_col = str(row.iloc[0]).strip()
                second_col = str(row.iloc[1]).strip()
                
                # 更宽松的匹配条件
                if (first_col == fee_name or first_col.strip() == fee_name.strip()) and \
                   ("已发生金额" in second_col or "已发生成本" in second_col or "已发生" in second_col) and \
                   "累计" not in second_col:
                    try:
                        raw = df.iloc[idx, 2:14].values  # 1-12月
                        # 确保数据类型正确（容错千分位/中文符号）
                        actual_data = to_float_array(raw)
                        if fee_name == "物耗成本" or "物耗" in fee_name:
                            print(f"物耗成本已发生数据: {actual_data}")
                    except Exception as e:
                        print(f"处理已发生金额数据时出错: {e}")
                        actual_data = None
                    break
            
            # 放宽条件：任意一项存在即加入；缺失的用0填充
            if target_data is None:
                target_data = np.zeros(12, dtype=float)
            if actual_data is None:
                actual_data = np.zeros(12, dtype=float)

            # 计算累计数据
            cum_target = np.cumsum(target_data)
            cum_actual = np.cumsum(actual_data)

            fee_items_data.append({
                'name': fee_name,
                'target': target_data.tolist(),  # 单月目标
                'actual': actual_data.tolist(),  # 单月已发生
                'cum_target': cum_target.tolist(),  # 月累目标
                'cum_actual': cum_actual.tolist()   # 月累已发生
            })

            if fee_name == "物耗成本":
                print(f"物耗成本数据加入：target(首月)={target_data[0]}, actual(首月)={actual_data[0]}")
        
        # 将数据按费项名称分组
        for item in fee_items_data:
            fee_name = item['name']
            if fee_name not in secondary_fee_data:
                secondary_fee_data[fee_name] = {
                    'target': np.zeros(12),  # 单月目标
                    'actual': np.zeros(12),  # 单月已发生
                    'cum_target': np.zeros(12),  # 月累目标
                    'cum_actual': np.zeros(12)   # 月累已发生
                }
            
            # 累加各项目的数据
            secondary_fee_data[fee_name]['target'] += np.array(item['target'])
            secondary_fee_data[fee_name]['actual'] += np.array(item['actual'])
            secondary_fee_data[fee_name]['cum_target'] += np.array(item['cum_target'])
            secondary_fee_data[fee_name]['cum_actual'] += np.array(item['cum_actual'])
    
    # 转换为列表格式，便于图表使用
    result = []
    for fee_name, data in secondary_fee_data.items():
        result.append({
            'name': fee_name,
            'target': data['target'].tolist(),  # 单月目标
            'actual': data['actual'].tolist(),  # 单月已发生
            'cum_target': data['cum_target'].tolist(),  # 月累目标
            'cum_actual': data['cum_actual'].tolist()   # 月累已发生
        })
        
        if fee_name == "物耗成本":
            print(f"最终物耗成本数据 - target: {data['target']}, actual: {data['actual']}")
    
    # 保存到缓存
    if month is not None:
        cache_manager = get_cache_manager()
        project_hash = hashlib.md5(str(sorted(all_main_dfs.keys())).encode()).hexdigest()
        cache_manager.save_secondary_fee_cache(f"combined_{project_hash}", month, result, include_self_owned_labor)
    
    return result

def extract_labor_service_breakdown(file):
    """从Excel文件中提取人工服务拆分数据
    
    Args:
        file: Excel文件路径或文件对象
        
    Returns:
        DataFrame: 人工服务拆分数据（前130行）
    """
    try:
        # 支持文件路径或文件对象
        if isinstance(file, str) or isinstance(file, Path):
            xl = pd.ExcelFile(file)
        else:
            # 处理文件对象，包括memoryview类型
            if hasattr(file, 'getbuffer'):
                buffer = file.getbuffer()
                # 将memoryview转换为bytes，然后用BytesIO包装
                if hasattr(buffer, 'tobytes'):
                    file_bytes = buffer.tobytes()
                else:
                    file_bytes = bytes(buffer)
                # 使用BytesIO包装bytes，避免FutureWarning
                xl = pd.ExcelFile(BytesIO(file_bytes))
            else:
                st.error(f"不支持的文件对象类型: {type(file)}")
                return None
        
        # 查找人工服务拆分工作表
        labor_sheet = None
        for sheet_name in xl.sheet_names:
            if '人工服务拆分' in sheet_name:
                labor_sheet = sheet_name
                break
        
        if labor_sheet is None:
            # 如果没找到，尝试其他可能的名称
            for sheet_name in xl.sheet_names:
                if '人工' in sheet_name and ('拆分' in sheet_name or '服务' in sheet_name):
                    labor_sheet = sheet_name
                    break
        
        if labor_sheet is None:
            st.warning(f"文件中未找到人工服务拆分工作表")
            return None
        
        # 读取前130行数据
        df = xl.parse(labor_sheet, nrows=130)
        
        # 清理数据：移除空行和无效行
        # 检查第一列和第二列，如果都是空值或无效值，则移除该行
        cleaned_rows = []
        for idx, row in df.iterrows():
            first_col = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            second_col = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # 跳过空行、标题行等
            if (first_col == "" or first_col == "nan" or 
                second_col == "" or second_col == "nan" or
                "人工服务费项" in first_col or "费项" in first_col):
                continue
            
            # 检查是否有数值数据（第3列开始）
            has_numeric_data = False
            for col_idx in range(2, min(14, len(row))):
                try:
                    val = row.iloc[col_idx]
                    if pd.notna(val) and str(val).strip() != "" and str(val).strip() != "nan":
                        # 尝试转换为数字
                        float_val = float(str(val).replace(',', '').replace('，', ''))
                        if float_val != 0:
                            has_numeric_data = True
                            break
                except:
                    continue
            
            # 只保留有数值数据的行
            if has_numeric_data:
                cleaned_rows.append(row)
        
        if cleaned_rows:
            cleaned_df = pd.DataFrame(cleaned_rows)
            return cleaned_df
        else:
            return None
        
    except Exception as e:
        st.error(f"提取人工服务拆分数据时出错: {e}")
        return None

def create_labor_service_summary(all_files, data_dir):
    """创建多项目人工服务拆分汇总表
    
    Args:
        all_files: 所有项目文件名列表
        data_dir: 数据目录路径
        
    Returns:
        DataFrame: 汇总的人工服务拆分数据
    """
    all_labor_data = []
    
    for filename in all_files:
        try:
            file_path = data_dir / filename
            labor_df = extract_labor_service_breakdown(file_path)
            
            if labor_df is not None:
                # 添加项目名称列
                project_name = filename.replace('.xlsx', '')
                labor_df['项目名称'] = project_name
                all_labor_data.append(labor_df)
                
        except Exception as e:
            st.error(f"处理文件 {filename} 的人工服务拆分数据时出错: {e}")
    
    if all_labor_data:
        # 合并所有项目的数据
        combined_df = pd.concat(all_labor_data, ignore_index=True)
        
        # 调试信息
        print(f"合并后的数据形状: {combined_df.shape}")
        print(f"合并后的列名: {list(combined_df.columns)}")
        
        # 找到数值列（1-12月）- 支持多种格式
        numeric_columns = []
        for col in combined_df.columns:
            col_str = str(col).strip()
            # 支持纯数字、数字+月、数字+月+其他字符等格式
            if (col_str.isdigit() and 1 <= int(col_str) <= 12) or \
               (col_str.replace('月', '').isdigit() and 1 <= int(col_str.replace('月', '')) <= 12) or \
               (col_str.replace('月', '').replace(' ', '').isdigit() and 1 <= int(col_str.replace('月', '').replace(' ', '')) <= 12):
                numeric_columns.append(col)
        
        print(f"找到的数值列: {numeric_columns}")
        
        if numeric_columns and len(numeric_columns) > 0:
            # 确定分组列（前两列通常是费项名称和数据类型）
            group_columns = []
            if len(combined_df.columns) >= 2:
                # 排除项目名称列和数值列
                exclude_cols = numeric_columns + ['项目名称']
                group_columns = [col for col in combined_df.columns if col not in exclude_cols]
            
            print(f"分组列: {group_columns}")
            
            if group_columns:
                # 创建汇总表 - 按费项名称和数据类型分组，对数值列进行累加
                summary_df = combined_df.groupby(group_columns)[numeric_columns].sum().reset_index()
                
                print(f"汇总后的数据形状: {summary_df.shape}")
                print(f"汇总后的列名: {list(summary_df.columns)}")
                
                # 添加项目名称列，统一设置为"汇总"
                summary_df['项目名称'] = '汇总'
                
                # 重新排列列顺序，确保与原表一致
                # 首先获取原始列顺序（不包括项目名称）
                original_cols = [col for col in combined_df.columns if col != '项目名称']
                
                # 确保数值列按正确顺序排列（1月到12月）
                numeric_cols_ordered = []
                for i in range(1, 13):
                    month_col = f"{i}月"
                    if month_col in summary_df.columns:
                        numeric_cols_ordered.append(month_col)
                    else:
                        # 处理可能的空格问题
                        month_col_with_space = f"{i} 月"
                        if month_col_with_space in summary_df.columns:
                            numeric_cols_ordered.append(month_col_with_space)
                
                # 构建最终的列顺序：分组列 + 数值列 + 项目名称
                final_cols = group_columns + numeric_cols_ordered + ['项目名称']
                
                # 确保所有列都存在
                available_cols = [col for col in final_cols if col in summary_df.columns]
                summary_df = summary_df[available_cols]
                
                print(f"最终汇总表形状: {summary_df.shape}")
                print(f"最终列顺序: {list(summary_df.columns)}")
                return summary_df
            else:
                print("没有找到合适的分组列，返回原始数据")
                # 如果没有找到合适的分组列，返回原始数据
                return combined_df
        else:
            print("没有找到数值列，返回原始数据")
            # 如果没有找到数值列，返回原始数据
            return combined_df
    
    return None

def create_formatted_summary_table(all_dfs, all_data):
    """创建格式化的汇总表显示，包含项目列表和汇总数据"""
    if not all_dfs or not all_data:
        return None
    
    # 创建汇总表
    summary_df = create_summary_excel(all_dfs)
    if summary_df is None:
        return None
    
    # 添加项目信息
    project_list = list(all_data.keys())
    
    # 创建汇总信息
    summary_info = {
        '汇总类型': '多项目汇总',
        '参与项目数': len(project_list),
        '项目列表': ', '.join(project_list),
        '汇总时间': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    return {
        'summary_df': summary_df,
        'summary_info': summary_info,
        'project_list': project_list
    }

def recalculate_project_metrics_for_month(data, month, project_name=None, include_self_owned_labor=False):
    """
    根据指定月份重新计算项目的指标数据
    用于多项目分析中甜甜圈图的动态更新
    
    Args:
        data: 原始项目数据
        month: 目标月份
        project_name: 项目名称（用于缓存）
        include_self_owned_labor: 是否包含自有劳务
    
    Returns:
        更新后的项目数据
    """
    # 确保month是整数类型
    month = int(month) if isinstance(month, str) else month
    
    # 尝试从缓存获取
    if project_name:
        cache_manager = get_cache_manager()
        cached_data = cache_manager.get_project_analysis_cache(project_name, month, include_self_owned_labor)
        if cached_data:
            return cached_data
    
    # 如果没有缓存，重新计算数据
    # 这里我们需要重新处理原始数据，但由于我们没有原始的DataFrame，
    # 我们可以基于现有的累计数据重新计算指标
    
    # 获取累计数据
    cum_target = data.get('cum_target', [0] * 12)
    cum_actual = data.get('cum_actual', [0] * 12)
    
    # 确保数据长度足够
    while len(cum_target) < 12:
        cum_target.append(0)
    while len(cum_actual) < 12:
        cum_actual.append(0)
    
    # 计算年总目标（12月份的累计总目标成本）
    year_cum_target = cum_target[-1] if len(cum_target) > 0 else 0
    year_cum_actual = cum_actual[-1] if len(cum_actual) > 0 else 0
    
    # 万元换算
    year_cum_target_wy = round(year_cum_target / 10000, 2)
    year_cum_actual_wy = round(year_cum_actual / 10000, 2)
    
    # 根据新月份重新计算使用率
    month_index = month - 1
    if month_index < len(cum_actual) and month_index < len(cum_target):
        current_cum_actual = cum_actual[month_index]
        current_cum_target = cum_target[month_index]
        
        # 使用率计算
        year_usage = round(100 * current_cum_actual / year_cum_target, 2) if year_cum_target else 0
        month_usage = round(100 * current_cum_actual / current_cum_target, 2) if current_cum_target else 0
    else:
        year_usage = 0
        month_usage = 0
    
    # 时间进度
    time_progress = round(100 * month / 12, 2)
    
    # 创建更新后的数据
    updated_data = data.copy()
    updated_data.update({
        'year_usage': year_usage,
        'month_usage': month_usage,
        'time_progress': time_progress,
        'year_cum_target_wy': year_cum_target_wy,
        'year_cum_actual_wy': year_cum_actual_wy,
        'cum_target': cum_target,  # 确保保留累计数据
        'cum_actual': cum_actual   # 确保保留累计数据
    })
    
    # 确保保留其他重要数据字段
    if 'fee_items' in data:
        # 重新计算fee_items的累计值，根据选择的月份
        updated_fee_items = []
        for item in data['fee_items']:
            # 这里我们需要重新处理原始数据，但由于我们没有原始的DataFrame，
            # 我们只能使用现有的累计数据
            # 在实际应用中，这里应该重新调用process_excel_data来获取正确的数据
            updated_fee_items.append(item)
        updated_data['fee_items'] = updated_fee_items
    if 'exceptions' in data:
        updated_data['exceptions'] = data['exceptions']
    if 'tertiary_fee_items' in data:
        updated_data['tertiary_fee_items'] = data['tertiary_fee_items']
    if 'tertiary_exceptions' in data:
        updated_data['tertiary_exceptions'] = data['tertiary_exceptions']
    
    # 保存到缓存
    if project_name:
        cache_manager = get_cache_manager()
        cache_manager.save_project_analysis_cache(project_name, month, updated_data, include_self_owned_labor)
    
    return updated_data

def create_comprehensive_summary_table(all_data, all_main_dfs, all_tertiary_dfs, month, include_self_owned_labor=False):
    """创建综合汇总表格，包含所有图表显示的数据
    
    Args:
        all_data: 所有项目的数据字典
        all_main_dfs: 所有项目的主要费项DataFrame
        all_tertiary_dfs: 所有项目的三级费项DataFrame
        month: 选择的月份
        include_self_owned_labor: 是否包含自有人工成本
        
    Returns:
        DataFrame: 综合汇总表格
    """
    try:
        summary_data = []
        
        # 1. 项目基本信息汇总
        for project_name, data in all_data.items():
            if not isinstance(data, dict):
                continue
                
            # 获取主要费项数据
            fee_items = data.get('fee_items', [])
            if isinstance(fee_items, list):
                for item in fee_items:
                    if isinstance(item, dict):
                        # 使用正确的字段名称
                        fee_name = item.get('name', '')
                        cum_target = item.get('cum_target', 0)
                        cum_actual = item.get('cum_actual', 0)
                        
                        # 计算使用率
                        usage_rate = round(100 * cum_actual / cum_target, 2) if cum_target else 0
                        
                        summary_data.append({
                            '数据类型': '项目基本信息',
                            '项目名称': project_name,
                            '费项名称': fee_name,
                            '费项编码': f'FEE_{len(summary_data)}',  # 生成编码
                            '目标金额': cum_target,
                            '已发生金额': cum_actual,
                            '使用率': usage_rate,
                            '月份': month,
                            '数据来源': '主要费项分析'
                        })
            
            # 获取三级费项数据
            tertiary_items = data.get('tertiary_fee_items', [])
            if isinstance(tertiary_items, list):
                for item in tertiary_items:
                    if isinstance(item, dict):
                        # 使用正确的字段名称
                        fee_name = item.get('name', '')
                        cum_target = item.get('cum_target', 0)
                        cum_actual = item.get('cum_actual', 0)
                        
                        # 计算使用率
                        usage_rate = round(100 * cum_actual / cum_target, 2) if cum_target else 0
                        
                        summary_data.append({
                            '数据类型': '三级费项明细',
                            '项目名称': project_name,
                            '费项名称': fee_name,
                            '费项编码': f'TER_{len(summary_data)}',  # 生成编码
                            '目标金额': cum_target,
                            '已发生金额': cum_actual,
                            '使用率': usage_rate,
                            '月份': month,
                            '数据来源': '三级费项分析'
                        })
            
            # 获取异常数据
            exceptions = data.get('exceptions', [])
            if isinstance(exceptions, list):
                for item in exceptions:
                    if isinstance(item, dict):
                        # 使用正确的字段名称
                        fee_name = item.get('fee_name', '')
                        cum_actual = item.get('cum_actual', 0)
                        cum_target = item.get('cum_target', 0)
                        year_target = item.get('year_target', 0)
                        
                        # 计算使用率
                        usage_rate = round(100 * cum_actual / year_target, 2) if year_target else 0
                        
                        summary_data.append({
                            '数据类型': '异常数据',
                            '项目名称': project_name,
                            '费项名称': fee_name,
                            '费项编码': f'EXC_{len(summary_data)}',  # 生成编码
                            '异常类型': item.get('exception_type', ''),
                            '异常月份': item.get('month', ''),
                            '目标金额': cum_target,
                            '已发生金额': cum_actual,
                            '使用率': usage_rate,
                            '月份': month,
                            '数据来源': '异常检测'
                        })
            
            # 获取三级费项异常
            tertiary_exceptions = data.get('tertiary_exceptions', [])
            if isinstance(tertiary_exceptions, list):
                for item in tertiary_exceptions:
                    if isinstance(item, dict):
                        # 使用正确的字段名称
                        fee_name = item.get('fee_name', '')
                        cum_actual = item.get('cum_actual', 0)
                        cum_target = item.get('cum_target', 0)
                        year_target = item.get('year_target', 0)
                        
                        # 计算使用率
                        usage_rate = round(100 * cum_actual / year_target, 2) if year_target else 0
                        
                        summary_data.append({
                            '数据类型': '三级费项异常',
                            '项目名称': project_name,
                            '费项名称': fee_name,
                            '费项编码': f'TER_EXC_{len(summary_data)}',  # 生成编码
                            '异常类型': item.get('exception_type', ''),
                            '异常月份': item.get('month', ''),
                            '目标金额': cum_target,
                            '已发生金额': cum_actual,
                            '使用率': usage_rate,
                            '月份': month,
                            '数据来源': '三级费项异常检测'
                        })
        
        # 2. 多项目汇总数据（如果有多个项目）
        if len(all_data) > 1:
            merged_data = merge_project_data(all_data, all_main_dfs, month, include_self_owned_labor)
            if merged_data:
                # 添加合并后的汇总数据
                summary_data.append({
                    '数据类型': '多项目汇总',
                    '项目名称': '所有项目汇总',
                    '费项名称': '总成本',
                    '费项编码': 'TOTAL',
                    '目标金额': merged_data.get('total_target', 0),
                    '已发生金额': merged_data.get('total_actual', 0),
                    '使用率': merged_data.get('total_usage_rate', 0),
                    '月份': month,
                    '数据来源': '多项目合并分析'
                })
        
        # 3. 月度趋势数据
        for project_name, data in all_data.items():
            if not isinstance(data, dict):
                continue
                
            # 获取月度数据（如果有的话）
            monthly_data = data.get('monthly_data', {})
            if isinstance(monthly_data, dict):
                for month_num, month_info in monthly_data.items():
                    if isinstance(month_info, dict):
                        summary_data.append({
                            '数据类型': '月度趋势',
                            '项目名称': project_name,
                            '费项名称': '月度累计',
                            '费项编码': f'MONTH_{month_num}',
                            '目标金额': month_info.get('target', 0),
                            '已发生金额': month_info.get('actual', 0),
                            '使用率': month_info.get('usage_rate', 0),
                            '月份': month_num,
                            '数据来源': '月度趋势分析'
                        })
        
        # 创建DataFrame
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            
            # 重新排列列顺序
            column_order = [
                '数据类型', '项目名称', '费项名称', '费项编码', 
                '目标金额', '已发生金额', '使用率', '月份', '数据来源'
            ]
            
            # 确保所有列都存在
            for col in column_order:
                if col not in summary_df.columns:
                    summary_df[col] = ''
            
            # 重新排列列顺序
            summary_df = summary_df[column_order]
            
            # 格式化数值列
            if '目标金额' in summary_df.columns:
                summary_df['目标金额'] = pd.to_numeric(summary_df['目标金额'], errors='coerce').fillna(0)
            if '已发生金额' in summary_df.columns:
                summary_df['已发生金额'] = pd.to_numeric(summary_df['已发生金额'], errors='coerce').fillna(0)
            if '使用率' in summary_df.columns:
                summary_df['使用率'] = pd.to_numeric(summary_df['使用率'], errors='coerce').fillna(0)
            
            return summary_df
        else:
            return pd.DataFrame()
            
    except Exception as e:
        st.error(f"创建综合汇总表格时出错: {e}")
        return pd.DataFrame()

def create_multi_project_export_data(all_data, all_main_dfs, all_tertiary_dfs, month, include_self_owned_labor=False):
    """创建多项目汇总数据导出表格，包含用户当前选择月份的多项目合并指标数据
    
    Args:
        all_data: 所有项目的数据字典
        all_main_dfs: 所有项目的主要费项DataFrame
        all_tertiary_dfs: 所有项目的三级费项DataFrame
        month: 选择的月份
        include_self_owned_labor: 是否包含自有人工成本
        
    Returns:
        dict: 包含各种汇总数据的字典
    """
    try:
        export_data = {}
        
        # 1. 多项目合并指标数据（包含所有月份）
        merged_metrics = {}
        
        # 先计算所有项目的年度总数据（只需要计算一次）
        total_year_target = 0
        total_year_actual = 0
        
        for project_name, data in all_data.items():
            if not isinstance(data, dict):
                continue
            # 累加年度总目标
            year_target = data.get('year_cum_target_wy', 0)
            total_year_target += year_target
            # 累加年度总实际
            year_actual = data.get('year_cum_actual_wy', 0)
            total_year_actual += year_actual
        
        # 计算所有月份（1-12月）的合并指标
        for month_idx in range(1, 13):
            month_num = month_idx
            
            # 初始化该月的累加数据
            month_cum_target = 0
            month_cum_actual = 0
            
            # 累加所有项目在该月的累计数据
            for project_name, data in all_data.items():
                if not isinstance(data, dict):
                    continue
                    
                # 获取该月的累计目标数据
                cum_target = data.get('cum_target', [])
                if month_idx <= len(cum_target):
                    month_cum_target += cum_target[month_idx - 1]
                
                # 获取该月的累计实际数据
                cum_actual = data.get('cum_actual', [])
                if month_idx <= len(cum_actual):
                    month_cum_actual += cum_actual[month_idx - 1]
            
            # 计算该月的指标
            year_usage_rate = round((total_year_actual / total_year_target * 100), 2) if total_year_target > 0 else 0
            month_usage_rate = round((month_cum_actual / month_cum_target * 100), 2) if month_cum_target > 0 else 0
            time_progress = round((month_num / 12 * 100), 2)
            
            # 转换为万元
            total_year_target_wy = total_year_target  # 已经是万元，不需要转换
            total_year_actual_wy = total_year_actual  # 已经是万元，不需要转换
            month_cum_target_wy = round(month_cum_target / 10000, 2)
            month_cum_actual_wy = round(month_cum_actual / 10000, 2)
            
            # 存储该月的合并指标
            merged_metrics[f'{month_num}月'] = {
                '月份': month_num,
                '年总目标(万元)': total_year_target_wy,
                '年总已发生(万元)': total_year_actual_wy,
                '月累目标(万元)': month_cum_target_wy,
                '月累已发生(万元)': month_cum_actual_wy,
                '年使用率(%)': year_usage_rate,
                '月累使用率(%)': month_usage_rate,
                '时间进度(%)': time_progress
            }
        
        export_data['合并指标'] = merged_metrics
        
        # 2. 合并项目列表
        project_list = []
        for project_name in all_data.keys():
            project_list.append({
                '项目名称': project_name,
                '项目类型': '成本分析项目',
                '包含自有人工成本': include_self_owned_labor
            })
        export_data['项目列表'] = project_list
        
        # 3. 每月费项成本与月累趋势图表源数据
        monthly_trend_data = []
        for project_name, data in all_data.items():
            if isinstance(data, dict):
                cum_target = data.get('cum_target', [])
                cum_actual = data.get('cum_actual', [])
                
                for i, (target, actual) in enumerate(zip(cum_target, cum_actual)):
                    if i < len(cum_target) and i < len(cum_actual):
                        monthly_trend_data.append({
                            '项目名称': project_name,
                            '月份': i + 1,
                            '累计目标成本': target,
                            '累计已发生成本': actual,
                            '使用率': round(100 * actual / target, 2) if target else 0
                        })
        export_data['月度趋势数据'] = monthly_trend_data
        
        # 4. 项目对比分析数据表格
        project_comparison = []
        for project_name, data in all_data.items():
            if isinstance(data, dict):
                fee_items = data.get('fee_items', [])
                if isinstance(fee_items, list):
                    for item in fee_items:
                        if isinstance(item, dict):
                            fee_name = item.get('name', '')
                            cum_target = item.get('cum_target', 0)
                            cum_actual = item.get('cum_actual', 0)
                            usage_rate = round(100 * cum_actual / cum_target, 2) if cum_target else 0
                            
                            project_comparison.append({
                                '项目名称': project_name,
                                '费项名称': fee_name,
                                '费项编码': f'FEE_{len(project_comparison)}',
                                '目标金额': cum_target,
                                '已发生金额': cum_actual,
                                '使用率': usage_rate,
                                '月份': month
                            })
        export_data['项目对比分析'] = project_comparison
        
        # 5. 详细指标对比汇总
        detailed_comparison = []
        for project_name, data in all_data.items():
            if isinstance(data, dict):
                detailed_comparison.append({
                    '项目名称': project_name,
                    '年总目标成本': data.get('year_cum_target_wy', 0),
                    '年总已发生成本': data.get('year_cum_actual_wy', 0),
                    '月累目标成本': data.get('month_cum_target_wy', 0),
                    '月累已发生成本': data.get('month_cum_actual_wy', 0),
                    '年使用率': data.get('year_usage', 0),
                    '月累使用率': data.get('month_usage', 0),
                    '时间进度': data.get('time_progress', 0)
                })
        export_data['详细指标对比'] = detailed_comparison
        
        # 6. 各类二级费项整体分析表格
        secondary_fee_analysis = []
        for project_name, data in all_data.items():
            if isinstance(data, dict):
                fee_items = data.get('fee_items', [])
                if isinstance(fee_items, list):
                    for item in fee_items:
                        if isinstance(item, dict):
                            fee_name = item.get('name', '')
                            cum_target = item.get('cum_target', 0)
                            cum_actual = item.get('cum_actual', 0)
                            usage_rate = round(100 * cum_actual / cum_target, 2) if cum_target else 0
                            
                            secondary_fee_analysis.append({
                                '项目名称': project_name,
                                '二级费项名称': fee_name,
                                '累计目标成本': cum_target,
                                '累计已发生成本': cum_actual,
                                '使用率': usage_rate,
                                '分析月份': month
                            })
        export_data['二级费项分析'] = secondary_fee_analysis
        
        # 7. 主控费项异常详情
        main_exceptions = []
        for project_name, data in all_data.items():
            if isinstance(data, dict):
                exceptions = data.get('exceptions', [])
                if isinstance(exceptions, list):
                    for item in exceptions:
                        if isinstance(item, dict):
                            main_exceptions.append({
                                '项目名称': project_name,
                                '费项名称': item.get('fee_name', ''),
                                '异常月份': item.get('month', ''),
                                '异常类型': item.get('exception_type', ''),
                                '累计目标成本': item.get('cum_target', 0),
                                '累计已发生成本': item.get('cum_actual', 0),
                                '年总目标成本': item.get('year_target', 0),
                                '异常说明': f"{item.get('exception_type', '')}异常：已发生成本超过目标"
                            })
        export_data['主控费项异常'] = main_exceptions
        
        # 8. 三级费项异常详情明细
        tertiary_exceptions = []
        for project_name, data in all_data.items():
            if isinstance(data, dict):
                tertiary_exceptions_data = data.get('tertiary_exceptions', [])
                if isinstance(tertiary_exceptions_data, list):
                    for item in tertiary_exceptions_data:
                        if isinstance(item, dict):
                            tertiary_exceptions.append({
                                '项目名称': project_name,
                                '费项名称': item.get('fee_name', ''),
                                '异常月份': item.get('month', ''),
                                '异常类型': item.get('exception_type', ''),
                                '累计目标成本': item.get('cum_target', 0),
                                '累计已发生成本': item.get('cum_actual', 0),
                                '年总目标成本': item.get('year_target', 0),
                                '异常说明': f"{item.get('exception_type', '')}异常：已发生成本超过目标"
                            })
        export_data['三级费项异常'] = tertiary_exceptions
        
        # 9. 异常排行榜数据
        # 统计每个项目的异常数量
        project_stats = {}
        for project_name, data in all_data.items():
            if not isinstance(data, dict):
                continue
                
            # 统计二级费项异常（主控费项异常）
            main_exceptions = data.get('exceptions', [])
            if not isinstance(main_exceptions, list):
                main_exceptions = []
                
            main_red_count = 0
            main_yellow_count = 0
            for e in main_exceptions:
                if isinstance(e, dict):
                    # 处理异常数据中的月份格式
                    exception_month = e.get('month', 0)
                    if isinstance(exception_month, str):
                        import re
                        month_match = re.search(r'(\d+)', exception_month)
                        if month_match:
                            exception_month = int(month_match.group(1))
                        else:
                            exception_month = 0
                    else:
                        exception_month = int(exception_month) if exception_month is not None else 0
                    
                    if e.get('exception_type') == 'red' and exception_month <= month:
                        main_red_count += 1
                    elif e.get('exception_type') == 'yellow' and exception_month <= month:
                        main_yellow_count += 1
            
            # 统计三级费项异常
            tertiary_exceptions_data = data.get('tertiary_exceptions', [])
            if not isinstance(tertiary_exceptions_data, list):
                tertiary_exceptions_data = []
                
            tertiary_red_count = 0
            tertiary_yellow_count = 0
            for e in tertiary_exceptions_data:
                if isinstance(e, dict):
                    # 处理异常数据中的月份格式
                    exception_month = e.get('month', 0)
                    if isinstance(exception_month, str):
                        import re
                        month_match = re.search(r'(\d+)', exception_month)
                        if month_match:
                            exception_month = int(month_match.group(1))
                        else:
                            exception_month = 0
                    else:
                        exception_month = int(exception_month) if exception_month is not None else 0
                    
                    if e.get('exception_type') == 'red' and exception_month <= month:
                        tertiary_red_count += 1
                    elif e.get('exception_type') == 'yellow' and exception_month <= month:
                        tertiary_yellow_count += 1
            
            # 计算总数
            main_total = main_red_count + main_yellow_count
            tertiary_total = tertiary_red_count + tertiary_yellow_count
            total_exceptions = main_total + tertiary_total
            
            project_stats[project_name] = {
                'main_red': main_red_count,
                'main_yellow': main_yellow_count,
                'main_total': main_total,
                'tertiary_red': tertiary_red_count,
                'tertiary_yellow': tertiary_yellow_count,
                'tertiary_total': tertiary_total,
                'total_exceptions': total_exceptions
            }
        
        # 创建异常排行榜数据
        if project_stats:
            # 二级费项异常排行榜
            main_ranking = sorted(project_stats.items(), key=lambda x: x[1]['main_total'], reverse=True)
            main_ranking_data = []
            for rank, (project_name, stats) in enumerate(main_ranking, 1):
                main_ranking_data.append({
                    '排名': rank,
                    '项目名称': project_name,
                    '红色异常': stats['main_red'],
                    '黄色异常': stats['main_yellow'],
                    '总异常数': stats['main_total']
                })
            export_data['二级费项异常排行榜'] = main_ranking_data
            
            # 三级费项异常排行榜
            tertiary_ranking = sorted(project_stats.items(), key=lambda x: x[1]['tertiary_total'], reverse=True)
            tertiary_ranking_data = []
            for rank, (project_name, stats) in enumerate(tertiary_ranking, 1):
                tertiary_ranking_data.append({
                    '排名': rank,
                    '项目名称': project_name,
                    '红色异常': stats['tertiary_red'],
                    '黄色异常': stats['tertiary_yellow'],
                    '总异常数': stats['tertiary_total']
                })
            export_data['三级费项异常排行榜'] = tertiary_ranking_data
            
            # 综合异常排行榜
            total_ranking = sorted(project_stats.items(), key=lambda x: x[1]['total_exceptions'], reverse=True)
            total_ranking_data = []
            for rank, (project_name, stats) in enumerate(total_ranking, 1):
                total_ranking_data.append({
                    '排名': rank,
                    '项目名称': project_name,
                    '二级费项异常': stats['main_total'],
                    '三级费项异常': stats['tertiary_total'],
                    '总异常数': stats['total_exceptions']
                })
            export_data['综合异常排行榜'] = total_ranking_data
        
        # 10. 项目数据表格预览数据（关键数据合并）
        # 参考图二格式：数据类型行 + 12个月列
        project_data_preview = []
        
        # 初始化12个月的数据累加
        monthly_actual_sum = [0] * 12      # 月度已发生金额累加
        monthly_target_sum = [0] * 12      # 月度目标金额累加
        cum_actual_sum = [0] * 12          # 月累已发生金额累加
        cum_target_sum = [0] * 12          # 月累目标金额累加
        
        # 累加所有项目的数据
        for project_name, data in all_data.items():
            if not isinstance(data, dict):
                continue
                
            # 调试：打印数据结构
            if project_name == list(all_data.keys())[0]:  # 只打印第一个项目的数据结构
                print(f"Debug - Project {project_name} data keys: {list(data.keys())}")
                if 'monthly_data' in data:
                    print(f"Debug - monthly_data structure: {data['monthly_data']}")
                else:
                    print(f"Debug - No monthly_data field found")
                
            # 获取月度数据
            monthly_data = data.get('monthly_data', {})
            if not isinstance(monthly_data, dict):
                continue
                
            # 获取累计目标数据
            cum_target = data.get('cum_target', [])
            if not isinstance(cum_target, list):
                cum_target = []
                
            # 获取累计实际数据
            cum_actual = data.get('cum_actual', [])
            if not isinstance(cum_actual, list):
                cum_actual = []
                
            # 获取月度目标数据
            monthly_target = data.get('monthly_target', [])
            if not isinstance(monthly_target, list):
                monthly_target = []
                
            # 获取月度实际数据
            monthly_actual = data.get('monthly_actual', [])
            if not isinstance(monthly_actual, list):
                monthly_actual = []
            
            # 累加12个月的数据
            for month_idx in range(12):
                month_key = f'{month_idx + 1}月'
                
                # 累加月度数据 - 修复：从累计数据中计算月度数据
                month_actual = 0
                if month_idx < len(cum_actual):
                    if month_idx == 0:
                        month_actual = cum_actual[month_idx]  # 第1月 = 第1月累计
                    else:
                        month_actual = cum_actual[month_idx] - cum_actual[month_idx - 1]  # 其他月 = 当月累计 - 上月累计
                monthly_actual_sum[month_idx] += month_actual
                
                # 累加月度目标 - 修复：从累计数据中计算月度数据
                month_target = 0
                if month_idx < len(cum_target):
                    if month_idx == 0:
                        month_target = cum_target[month_idx]  # 第1月 = 第1月累计
                    else:
                        month_target = cum_target[month_idx] - cum_target[month_idx - 1]  # 其他月 = 当月累计 - 上月累计
                monthly_target_sum[month_idx] += month_target
                
                # 累加累计数据
                cum_target_value = cum_target[month_idx] if month_idx < len(cum_target) else 0
                cum_target_sum[month_idx] += cum_target_value
                
                cum_actual_value = cum_actual[month_idx] if month_idx < len(cum_actual) else 0
                cum_actual_sum[month_idx] += cum_actual_value
        
        # 转换为元并创建图二格式的数据（保持原始精度）
        # 数据类型0：已发生金额
        actual_row = {
            '数据类型': '已发生金额',
            '1月': monthly_actual_sum[0],
            '2月': monthly_actual_sum[1],
            '3月': monthly_actual_sum[2],
            '4月': monthly_actual_sum[3],
            '5月': monthly_actual_sum[4],
            '6月': monthly_actual_sum[5],
            '7月': monthly_actual_sum[6],
            '8月': monthly_actual_sum[7],
            '9月': monthly_actual_sum[8],
            '10月': monthly_actual_sum[9],
            '11月': monthly_actual_sum[10],
            '12月': monthly_actual_sum[11]
        }
        
        # 数据类型1：月累已发生金额
        cum_actual_row = {
            '数据类型': '月累已发生金额',
            '1月': cum_actual_sum[0],
            '2月': cum_actual_sum[1],
            '3月': cum_actual_sum[2],
            '4月': cum_actual_sum[3],
            '5月': cum_actual_sum[4],
            '6月': cum_actual_sum[5],
            '7月': cum_actual_sum[6],
            '8月': cum_actual_sum[7],
            '9月': cum_actual_sum[8],
            '10月': cum_actual_sum[9],
            '11月': cum_actual_sum[10],
            '12月': cum_actual_sum[11]
        }
        
        # 数据类型2：目标金额
        target_row = {
            '数据类型': '目标金额',
            '1月': monthly_target_sum[0],
            '2月': monthly_target_sum[1],
            '3月': monthly_target_sum[2],
            '4月': monthly_target_sum[3],
            '5月': monthly_target_sum[4],
            '6月': monthly_target_sum[5],
            '7月': monthly_target_sum[6],
            '8月': monthly_target_sum[7],
            '9月': monthly_target_sum[8],
            '10月': monthly_target_sum[9],
            '11月': monthly_target_sum[10],
            '12月': monthly_target_sum[11]
        }
        
        # 数据类型3：目标金额累计
        cum_target_row = {
            '数据类型': '目标金额累计',
            '1月': cum_target_sum[0],
            '2月': cum_target_sum[1],
            '3月': cum_target_sum[2],
            '4月': cum_target_sum[3],
            '5月': cum_target_sum[4],
            '6月': cum_target_sum[5],
            '7月': cum_target_sum[6],
            '8月': cum_target_sum[7],
            '9月': cum_target_sum[8],
            '10月': cum_target_sum[9],
            '11月': cum_target_sum[10],
            '12月': cum_target_sum[11]
        }
        
        # 按图二格式添加数据行
        project_data_preview.extend([actual_row, cum_actual_row, target_row, cum_target_row])
        
        if project_data_preview:
            export_data['项目数据表格预览'] = project_data_preview
        
        return export_data
        
    except Exception as e:
        st.error(f"创建多项目汇总数据导出表格时出错: {e}")
        return {}

def create_data_summary_tables(all_main_dfs, all_tertiary_dfs, include_self_owned_labor=False):
    """
    创建多项目数据汇总表格（可供二次分析区块项目）
    包含4主要费项费项月累成本使用情况、4-1主要费项费项月累成本使用情况和三级费项月累表格的累加数据
    
    Args:
        all_main_dfs: 所有项目的主要费项DataFrame字典
        all_tertiary_dfs: 所有项目的三级费项DataFrame字典
        include_self_owned_labor: 是否包含自有人工成本
    
    Returns:
        dict: 包含三个汇总表格的字典
    """
    try:
        if not all_main_dfs or not all_tertiary_dfs:
            return {}
        
        # 确定主要费项表格名称
        main_sheet_name_4 = '4主要费项费项月累成本使用情况'
        main_sheet_name_4_1 = '4-1主要费项费项月累成本使用情况'
        tertiary_sheet_name = '三级费项月累表格'
        
        # 初始化累加数据
        main_merged_df_4 = None
        main_merged_df_4_1 = None
        tertiary_merged_df = None
        
        # 累加主要费项数据（4表格）
        for project_name, main_df in all_main_dfs.items():
            if main_df is not None and not main_df.empty:
                if main_merged_df_4 is None:
                    # 第一次，直接复制整个DataFrame以保持原始格式
                    main_merged_df_4 = main_df.copy()
                    # 识别数值列（用于累加）和非数值列（保持原始值）
                    numeric_columns = main_merged_df_4.select_dtypes(include=[np.number]).columns
                    # 对于非数值列，保持原始值（如费项名称、单位等）
                    # 对于数值列，保持原始值（第一次复制）
                else:
                    # 后续项目，只累加数值列，保持非数值列不变
                    numeric_columns = main_df.select_dtypes(include=[np.number]).columns
                    for col in numeric_columns:
                        if col in main_merged_df_4.columns:
                            # 安全地累加数值，处理NaN值
                            current_values = main_merged_df_4[col].fillna(0)
                            new_values = main_df[col].fillna(0)
                            main_merged_df_4[col] = current_values + new_values
        
        # 累加主要费项数据（4-1表格）
        for project_name, main_df in all_main_dfs.items():
            if main_df is not None and not main_df.empty:
                if main_merged_df_4_1 is None:
                    # 第一次，直接复制整个DataFrame以保持原始格式
                    main_merged_df_4_1 = main_df.copy()
                    # 识别数值列（用于累加）和非数值列（保持原始值）
                    numeric_columns = main_merged_df_4_1.select_dtypes(include=[np.number]).columns
                    # 对于非数值列，保持原始值（如费项名称、单位等）
                    # 对于数值列，保持原始值（第一次复制）
                else:
                    # 后续项目，只累加数值列，保持非数值列不变
                    numeric_columns = main_df.select_dtypes(include=[np.number]).columns
                    for col in numeric_columns:
                        if col in main_merged_df_4_1.columns:
                            # 安全地累加数值，处理NaN值
                            current_values = main_merged_df_4_1[col].fillna(0)
                            new_values = main_df[col].fillna(0)
                            main_merged_df_4_1[col] = current_values + new_values
        
        # 累加三级费项数据
        for project_name, tertiary_df in all_tertiary_dfs.items():
            if tertiary_df is not None and not tertiary_df.empty:
                if tertiary_merged_df is None:
                    # 第一次，直接复制整个DataFrame以保持原始格式
                    tertiary_merged_df = tertiary_df.copy()
                    # 识别数值列（用于累加）和非数值列（保持原始值）
                    numeric_columns = tertiary_merged_df.select_dtypes(include=[np.number]).columns
                    # 对于非数值列，保持原始值（如费项名称、单位等）
                    # 对于数值列，保持原始值（第一次复制）
                else:
                    # 后续项目，只累加数值列，保持非数值列不变
                    numeric_columns = tertiary_df.select_dtypes(include=[np.number]).columns
                    for col in numeric_columns:
                        if col in tertiary_merged_df.columns:
                            # 安全地累加数值，处理NaN值
                            current_values = tertiary_merged_df[col].fillna(0)
                            new_values = tertiary_df[col].fillna(0)
                            tertiary_merged_df[col] = current_values + new_values
        
        # 准备返回数据
        summary_tables = {}
        
        # 同时返回4和4-1两张主要费项表格
        if main_merged_df_4 is not None:
            summary_tables[main_sheet_name_4] = main_merged_df_4
        
        if main_merged_df_4_1 is not None:
            summary_tables[main_sheet_name_4_1] = main_merged_df_4_1
        
        if tertiary_merged_df is not None:
            summary_tables[tertiary_sheet_name] = tertiary_merged_df
        
        return summary_tables
        
    except Exception as e:
        st.error(f"创建数据汇总表格时出错: {e}")
        return {}