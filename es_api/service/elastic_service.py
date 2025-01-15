from elasticsearch import Elasticsearch, helpers
from typing import List, Dict, Optional


class ElasticService:
    def __init__(self, uri: str = "http://192.168.154.26:9200"):
        """初始化 Elasticsearch 客户端"""
        self.client = Elasticsearch(hosts=[uri])

    def check_and_create_index(self, index: str, mappings: Optional[Dict] = None, settings: Optional[Dict] = None):
        """
        检查索引是否存在，如果不存在则创建
        :param index: 索引名称
        :param mappings: 索引的字段映射
        :param settings: 索引的设置
        """
        if not self.client.indices.exists(index=index):
            print(f"Index '{index}' does not exist. Creating...")
            body = {}
            if mappings:
                body["mappings"] = mappings
            if settings:
                body["settings"] = settings
            self.client.indices.create(index=index, body=body)
            print(f"Index '{index}' created successfully.")
        else:
            print(f"Index '{index}' already exists.")

    def write_bulk(self, data: List[Dict], index: str, mappings: Optional[Dict] = None, settings: Optional[Dict] = None):
        """
        批量写入数据到 Elasticsearch
        :param data: 宽表格式数据，每条记录是一个字典
        :param index: 写入的 Elasticsearch 索引名称
        :param mappings: 索引的字段映射（用于创建索引时）
        :param settings: 索引的设置（用于创建索引时）
        """
        # 检查并创建索引
        self.check_and_create_index(index=index, mappings=mappings, settings=settings)

        # 准备批量写入的数据
        actions = [
            {"_index": index, "_source": record} for record in data
        ]
        try:
            print(f"Writing {len(actions)} documents to index '{index}'...")
            response = helpers.bulk(self.client, actions)
            print(f"Successfully wrote {response[0]} documents to index '{index}'.")
            return response
        except Exception as e:
            print(f"Failed to write to Elasticsearch: {str(e)}")
            raise RuntimeError(f"Failed to write to Elasticsearch: {str(e)}")