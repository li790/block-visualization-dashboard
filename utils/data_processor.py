import pandas as pd
import numpy as np
import streamlit as st
from pathlib import Path
import os
from io import BytesIO
import time

def extract_table_from_excel(file):
    """从Excel文件中提取两个工作表的数据:
    1. 主要费项费项月累成本使用情况(前39行)
    2. 三级费项月累表格(前153行)
    支持文件路径或文件对象
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
                return None, None
            
        main_sheet = '4主要费项费项月累成本使用情况'
        tertiary_sheet = '三级费项月累表格'
        
        # 尝试不同的工作表名称格式
        main_sheets_to_try = [main_sheet, '主要费项费项月累成本使用情况', '主要费项']
        tertiary_sheets_to_try = [tertiary_sheet, '三级费项']
        
        main_df = None
        for sheet in main_sheets_to_try:
            if sheet in xl.sheet_names:
                main_df = xl.parse(sheet).head(39)
                break
        
        if main_df is None:
            st.error(f"文件中未找到主要费项工作表，尝试了: {', '.join(main_sheets_to_try)}")
            return None, None
        
        tertiary_df = None
        for sheet in tertiary_sheets_to_try:
            if sheet in xl.sheet_names:
                tertiary_df = xl.parse(sheet).head(153)
                break
        
        if tertiary_df is None:
            st.error(f"文件中未找到三级费项工作表，尝试了: {', '.join(tertiary_sheets_to_try)}")
            return None, None
        
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

def process_excel_data(df, month):
    """处理Excel数据 - 适配用户表格格式"""
    try:
        # 查找关键行
        total_target_row = None
        total_actual_row = None
        fee_items = []
        
        for idx, row in df.iterrows():
            first_col = str(row.iloc[0]).strip()
            second_col = str(row.iloc[1]).strip()
            
            # 查找总成本行
            if "总成本" in first_col or "年总成本" in first_col:
                if "月累总目标成本" in second_col:
                    total_target_row = idx
                elif "月累已发生成本" in second_col:
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
        
        # 提取总成本数据
        if total_target_row is not None and total_actual_row is not None:
            total_target_data = df.iloc[total_target_row, 2:14].values  # 1-12月
            total_actual_data = df.iloc[total_actual_row, 2:14].values  # 1-12月
        else:
            st.error("未找到总成本数据行")
            return None
        
        # 根据新要求修改计算逻辑：
        # 1. 累计总目标成本 = 原始表格 总成本 月累总目标成本
        # 2. 累计总已发生成本 = 原始表格 总成本 月累已发生成本
        # 3. 年使用率 = 用户选择月份的累计总已发生成本 / 12月份的累计总目标成本
        # 4. 月累使用率 = 用户选择月份的累计总已发生成本 / 用户选择月份的累计总目标成本
        # 5. 年总目标 = 12月份的累计总目标成本
        # 6. 月累目标 = 用户选择月份的累计总目标成本
        # 7. 月累已发生 = 用户选择月份的累计总已发生成本
        
        # 直接使用原始表格的月累数据（不重新计算累计）
        cum_target = total_target_data  # 原始表格 总成本 月累总目标成本
        cum_actual = total_actual_data  # 原始表格 总成本 月累已发生成本
        
        # 年总目标 = 12月份的累计总目标成本
        year_cum_target = cum_target[-1]  # 第12个月的数据
        year_cum_actual = cum_actual[-1]  # 第12个月的数据
        
        # 万元换算
        year_cum_target_wy = round(year_cum_target / 10000, 2)
        year_cum_actual_wy = round(year_cum_actual / 10000, 2)
        
        # 使用率计算（根据新要求）
        year_usage = round(100 * cum_actual[month-1] / year_cum_target, 2) if year_cum_target else 0
        month_usage = round(100 * cum_actual[month-1] / cum_target[month-1], 2) if cum_target[month-1] else 0
        time_progress = round(100 * month / 12, 2)
        
        # 调试信息已移除
        
        # 处理二级费项数据和异常项
        processed_fee_items = []
        exceptions = []
        for item in fee_items:
            if item['target_row'] is not None and item['actual_row'] is not None:
                target_data = df.iloc[item['target_row'], 2:14].values
                actual_data = df.iloc[item['actual_row'], 2:14].values
                
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
        
        return {
            'total_target': float(total_target_data.sum()),
            'cum_target': [float(x) for x in cum_target],
            'cum_actual': [float(x) for x in cum_actual],
            'year_cum_target_wy': year_cum_target_wy,
            'year_cum_actual_wy': year_cum_actual_wy,
            'year_usage': year_usage,
            'month_usage': month_usage,
            'time_progress': time_progress,
            'fee_items': processed_fee_items,
            'exceptions': exceptions
        }
        
    except Exception as e:
        st.error(f"处理Excel文件时出错: {e}")
        return None

def load_and_process_files(uploaded_files, selected_files, data_dir, output_dir, month):
    """加载并处理多个文件，数据需拆解处理
    data_dir: 存放原始数据的目录
    output_dir: 存放分析数据表和缓冲数据的目录
    """
    all_data = {}
    all_main_dfs = {}
    all_tertiary_dfs = {}
    
    # 创建缓冲数据目录
    buffer_dir = output_dir / 'buffer'
    buffer_dir.mkdir(exist_ok=True)
    
    # 创建分析结果目录
    analysis_dir = output_dir / 'analysis_results'
    analysis_dir.mkdir(exist_ok=True)
    
    # 创建处理后数据目录
    processed_dir = output_dir / 'processed_data'
    processed_dir.mkdir(exist_ok=True)
    
    # 处理上传的文件（这些文件已经在sidebar中处理过了，这里只需要处理已提取的数据）
    if uploaded_files:
        for i, uploaded_file in enumerate(uploaded_files):
            try:
                # 使用上传文件的真实名称（去掉扩展名），加上时间戳保证唯一
                uploaded_filename = uploaded_file.name.replace('.xlsx', '').replace('.xls', '')
                unique_id = time.strftime("%Y%m%d%H%M%S") + f"_{i}"
                unique_filename = f"{uploaded_filename}_{unique_id}"

                # 直接从上传文件中提取数据
                main_df, tertiary_df = extract_table_from_excel(uploaded_file)
                if main_df is None or tertiary_df is None:
                    st.warning(f"无法从文件 {uploaded_file.name} 中提取有效数据")
                    continue

                # 保存提取后的数据到唯一文件名
                main_file_path = processed_dir / f"{unique_filename}_主要费项.xlsx"
                tertiary_file_path = processed_dir / f"{unique_filename}_三级费项.xlsx"
                main_df.to_excel(main_file_path, index=False)
                tertiary_df.to_excel(tertiary_file_path, index=False)

                # 处理数据
                all_main_dfs[unique_filename] = main_df
                all_tertiary_dfs[unique_filename] = tertiary_df

                # 处理主要费项数据
                data = process_excel_data(main_df, month)
                if data:
                    all_data[unique_filename] = data

                    # 处理三级费项数据
                    tertiary_result = process_tertiary_fee_data(tertiary_df, month)
                    data['tertiary_fee_items'] = tertiary_result['tertiary_fee_items']
                    data['tertiary_exceptions'] = tertiary_result['exceptions']

                # 保存分析结果到output目录
                analysis_file_path = analysis_dir / f"{unique_filename}_分析结果.xlsx"
                with pd.ExcelWriter(analysis_file_path) as writer:
                    if data and 'fee_items' in data:
                        pd.DataFrame(data['fee_items']).to_excel(writer, sheet_name='二级费项分析', index=False)
                    if data and 'tertiary_fee_items' in data:
                        pd.DataFrame(data['tertiary_fee_items']).to_excel(writer, sheet_name='三级费项分析', index=False)

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
                main_df, tertiary_df = extract_table_from_excel(file_path)
                if main_df is None or tertiary_df is None:
                    st.warning(f"无法从文件 {filename} 中提取有效数据")
                    continue
                
                project_name = filename.replace('.xlsx', '')
                all_main_dfs[project_name] = main_df
                all_tertiary_dfs[project_name] = tertiary_df
                
                # 处理主要费项数据
                data = process_excel_data(main_df, month)
                if data:
                    all_data[project_name] = data
                    
                    # 处理三级费项数据
                    tertiary_result = process_tertiary_fee_data(tertiary_df, month)
                    data['tertiary_fee_items'] = tertiary_result['tertiary_fee_items']
                    data['tertiary_exceptions'] = tertiary_result['exceptions']
                    
                # 保存分析结果到output目录
                analysis_file_path = analysis_dir / f"{project_name}_分析结果.xlsx"
                with pd.ExcelWriter(analysis_file_path) as writer:
                    if data and 'fee_items' in data:
                        pd.DataFrame(data['fee_items']).to_excel(writer, sheet_name='二级费项分析', index=False)
                    if data and 'tertiary_fee_items' in data:
                        pd.DataFrame(data['tertiary_fee_items']).to_excel(writer, sheet_name='三级费项分析', index=False)
                
                st.success(f"成功处理文件: {filename}")
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
        # 格式化数值列
        numeric_columns = ['cum_actual', 'cum_target', 'year_target']
        exception_df[numeric_columns] = exception_df[numeric_columns].applymap(lambda x: f"{x:.2f}")
        # 替换异常类型为中文
        exception_df['exception_type'] = exception_df['exception_type'].replace({'red': '超年度目标', 'yellow': '超月度目标'})
        return exception_df
    
    return None

def create_summary_excel(all_dfs):
    """创建汇总Excel表 - 将所有项目的对应行列数据相加"""
    if not all_dfs:
        return None
    
    # 获取第一个DataFrame作为模板
    first_df = list(all_dfs.values())[0]
    summary_df = first_df.copy()
    
    # 初始化汇总数据为0
    for col in summary_df.columns:
        if col not in [summary_df.columns[0], summary_df.columns[1]]:  # 跳过前两列（费项名称和类型）
            summary_df[col] = 0
    
    # 将所有DataFrame的数值列相加
    for project_name, df in all_dfs.items():
        for col in df.columns:
            if col not in [df.columns[0], df.columns[1]]:  # 跳过前两列
                # 确保数据类型为数值型
                numeric_data = pd.to_numeric(df[col], errors='coerce').fillna(0)
                summary_df[col] += numeric_data
    
    return summary_df

def process_tertiary_fee_data(df, month):
    """处理三级费项数据并检测异常"""
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
        return {
            'tertiary_fee_items': tertiary_fee_items,
            'exceptions': exceptions
        }
    except Exception as e:
        st.error(f"处理三级费项数据时出错: {e}")
        return {'tertiary_fee_items': [], 'exceptions': []}

def merge_project_data(all_data, all_main_dfs, month):
    """合并多个项目的数据并计算合并后的关键指标"""
    if not all_data or not all_main_dfs:
        return None
    
    # 先合并原始Excel数据
    merged_df = create_summary_excel(all_main_dfs)
    if merged_df is None:
        return None
    
    # 使用合并后的原始数据重新计算关键指标
    merged_data = process_excel_data(merged_df, month)
    
    if merged_data:
        # 添加项目列表信息
        merged_data['merged_projects'] = list(all_data.keys())
    
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
                            monthly_data[month_key]['target'] += float(target_data[month])
                            monthly_data[month_key]['actual'] += float(actual_data[month])
                        except:
                            continue
    
    # 转换为DataFrame
    months = sorted(monthly_data.keys())
    data = {
        '月份': [f'{m}月' for m in months],
        '目标成本': [monthly_data[m]['target']/10000 for m in months],  # 转换为万元
        '已发生成本': [monthly_data[m]['actual']/10000 for m in months]  # 转换为万元
    }
    
    return pd.DataFrame(data)

def create_client_download_table(all_dfs, all_data): 
    # 创建结果DataFrame
    result_data = []
    
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
                        if (sub_fee_name == fee_name and 
                            "已发生金额" in sub_second_col and 
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
            actual_data = np.array(item['actual'], dtype=float)
            monthly_target_sum += target_data
            monthly_actual_sum += actual_data
        
        # 检查是否找到了总成本数据
        if total_target_data is None or total_actual_data is None:
            continue
            
        # 计算累计数据
        total_target_array = np.array(total_target_data, dtype=float)
        total_actual_array = np.array(total_actual_data, dtype=float)
        
        # 累计总目标成本 = 原始表格 总成本 月累总目标成本
        # 累计总已发生成本 = 原始表格 总成本 月累已发生成本
        cum_target = total_target_array  # 直接使用月累总目标成本
        cum_actual = total_actual_array  # 直接使用月累已发生成本
        
        # 为每个项目添加4行数据，月份横向排布
        # 行1: 已发生金额（每月各费项成本已发生金额之和）
        row1_data = {
            '项目': project_name,
            '数据类型': '已发生金额',
            '累计金额': float(np.sum(monthly_actual_sum))  # 年累计
        }
        # 添加1-12月数据
        for month in range(1, 13):
            row1_data[f'{month}月'] = float(monthly_actual_sum[month-1])
        result_data.append(row1_data)
        
        # 行2: 月累已发生金额（年总成本月累已发生金额）
        row2_data = {
            '项目': project_name,
            '数据类型': '月累已发生金额',
            '累计金额': float(np.sum(total_actual_array))  # 年累计
        }
        # 添加1-12月数据
        for month in range(1, 13):
            row2_data[f'{month}月'] = float(total_actual_array[month-1])
        result_data.append(row2_data)
        
        # 行3: 目标金额（每月各费项目标金额之和）
        row3_data = {
            '项目': project_name,
            '数据类型': '目标金额',
            '累计金额': float(np.sum(monthly_target_sum))  # 年累计
        }
        # 添加1-12月数据
        for month in range(1, 13):
            row3_data[f'{month}月'] = float(monthly_target_sum[month-1])
        result_data.append(row3_data)
        
        # 行4: 目标金额累计（年总成本月累目标金额）
        row4_data = {
            '项目': project_name,
            '数据类型': '目标金额累计',
            '累计金额': float(np.sum(total_target_array))  # 年累计
        }
        # 添加1-12月数据
        for month in range(1, 13):
            row4_data[f'{month}月'] = float(total_target_array[month-1])
        result_data.append(row4_data)
    
    # 创建DataFrame
    result_df = pd.DataFrame(result_data)
    
    # 重新排列列顺序：项目、数据类型、1月、2月、...、12月、累计金额
    columns = ['项目', '数据类型'] + [f'{i}月' for i in range(1, 13)] + ['累计金额']
    result_df = result_df[columns]
    
    return result_df