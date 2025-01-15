# routers/research_router.py
from fastapi import APIRouter, HTTPException
import requests
from es_api.service.elastic_service import ElasticService
from es_api.utils.parse_data import parse_keyanexport_data

router = APIRouter()
es_service = ElasticService(uri="http://192.168.154.26:9200")

@router.get("/ingest/keyanexport")
async def ingest_research_data(projectNumber: str):
    """
    拉取研究项目数据并写入 Elasticsearch
    :param projectNumber: 项目编号
    """
    print(f"Received request with projectNumber: {projectNumber}")  # 打印收到的请求参数
    try:
        # 从外部接口获取数据
        url = f"http://192.168.2.188:6001/api/keyanexport/keyanprojectinfo?projectNumber={projectNumber}"
        print(f"Fetching data from external API: {url}")
        response = requests.get(url, timeout=10)  # 设置超时时间，避免长时间无响应

        if response.status_code != 200:
            print(f"Failed to fetch data, status code: {response.status_code}")
            raise HTTPException(status_code=500, detail="Failed to fetch data from source API")

        raw_data = response.json()
        print(f"Successfully fetched data: {raw_data}")

        # 解析数据
        print("Parsing data...")
        parsed_data = parse_keyanexport_data(raw_data)
        print(f"Parsed data: {parsed_data}")

        # 自定义索引配置
        mappings = {
            "properties": {
                "ProjectNumber": {"type": "keyword"},
                "ProjectName": {"type": "keyword"},
                "Files": {
                    "type": "nested",
                    "properties": {
                        "Filehead": {"type": "keyword"},
                        "Filetype": {"type": "keyword"}
                    }
                }
            }
        }
        settings = {"number_of_shards": 1, "number_of_replicas": 1}

        # 写入 Elasticsearch
        print(f"Writing data to Elasticsearch index 'keyanexport'...")
        es_response = es_service.write_bulk(data=parsed_data, index="keyanexport", mappings=mappings, settings=settings)
        print(f"Elasticsearch response: {es_response}")

        return {"status": "success", "response": es_response}
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
