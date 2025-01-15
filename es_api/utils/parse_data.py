# utils/parse_data.py
from typing import List, Dict
from es_api.model.keyanexport import ProjectModel, FileInfo


def parse_keyanexport_data(raw_data: dict) -> List[Dict]:
    """
    将 ProjectInfo 和 Files 合并为宽表格式，同时保留 Files 为列表结构
    :param raw_data: 从接口返回的原始数据
    :return: 合并后的宽表格式数据列表
    """
    # 校验并解析 ProjectInfo
    project_info = ProjectModel(**raw_data["Result"]["ProjectInfo"])

    # 解析 Files 列表
    files = [FileInfo(**file).dict() for file in raw_data["Result"]["Files"]]

    # 将 Files 嵌套到 ProjectInfo 中
    project_data = project_info.dict()
    project_data["Files"] = files  # 保留 Files 为列表

    return [project_data]  # 返回一个包含单条记录的列表