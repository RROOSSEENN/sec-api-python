# utils/parse_data.py
from typing import List, Dict
from geofileparse_api.model.keyanexport import ProjectModel, FileInfo
from langchain_core.documents import Document


def parse_keyanexport_data(raw_data: dict, documents: List[Document]) -> List[Dict]:
    """
    将 ProjectInfo 和 Files 合并为宽表格式，同时保留 Files 为列表结构
    :param raw_data: 从接口返回的原始数据
    :param documents: 分割后的文档列表，包含 page_content 和 metadata
    :return: 合并后的宽表格式数据列表
    """
    # 校验并解析 ProjectInfo
    project_info = ProjectModel(**raw_data["Result"]["ProjectInfo"])

    # 解析 Files 列表
    files = []
    for file, doc in zip(raw_data["Result"]["Files"], documents):
        file_info = FileInfo(
            **file,  # 原始文件信息
            page_content=doc.page_content,  # 从 Document 中提取 page_content
            metadata=doc.metadata  # 从 Document 中提取 metadata
        )
        files.append(file_info.dict())  # 转换为字典并添加到文件列表

    # 将 Files 嵌套到 ProjectInfo 中
    project_data = project_info.dict()
    project_data["Files"] = files  # 保留 Files 为列表

    return [project_data]  # 返回一个包含单条记录的列表