# utils/parse_data.py
from typing import List, Dict
from geofileparse_api.model.keyanexport import ProjectModel, FileInfo
from langchain_core.documents import Document


def parse_keyanexport_data(raw_data: dict) -> List[Dict]:
    """
    将 ProjectInfo 和 Files 合并为宽表格式，同时保留 Files 的分割数据
    :param raw_data: 从接口返回的原始数据
    :return: 合并后的宽表格式数据列表
    """
    project_info = raw_data["Result"]["ProjectInfo"]
    files = raw_data["Result"]["Files"]

    project_data = {**project_info, "Files": files}
    return [project_data]  # 返回包含项目和文件信息的数据
