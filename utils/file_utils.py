from typing import Tuple, Dict, List

def split_type_and_display(project: str) -> Tuple[str, str]:
    """解析类型前缀与项目显示名 - 在第一个-前面截取类型前缀"""
    # 在第一个-前面截取类型前缀
    if '-' in project:
        type_prefix = project.split('-')[0]
        return type_prefix, project
    return project, project


def categorize_files(file_names: List[str]) -> Dict[str, List[str]]:
    """将文件按类型分类 - 基于第一个-前面的前缀"""
    type_to_files = {}
    for filename in file_names:
        # 去掉.xlsx扩展名
        project_name = filename.replace('.xlsx', '').replace('.xls', '')
        type_prefix, _ = split_type_and_display(project_name)
        type_to_files.setdefault(type_prefix, []).append(filename)
    return type_to_files