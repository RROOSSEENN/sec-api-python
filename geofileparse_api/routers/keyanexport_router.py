from langchain_core.documents import Document
from fastapi import APIRouter, HTTPException
import requests
import logging
from geofileparse_api.service.elastic_service import ElasticService
from geofileparse_api.utils.docs import GeoReportSplitter
from geofileparse_api.utils.parse_data import parse_keyanexport_data
from geofileparse_api.utils.process_file import FileProcessor

router = APIRouter()
es_service = ElasticService(uri="http://192.168.154.26:9200")

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("E:\R\workplace\es-api-python\geofileparse_api\log\ingest_keyanexport.log", encoding="utf-8")
    ]
)

@router.get("/ingest/keyanexport")
async def ingest_research_data(projectNumber: str):
    """
    拉取研究项目数据并写入 Elasticsearch
    :param projectNumber: 项目编号
    """
    logging.info(f"Received request with projectNumber: {projectNumber}")
    try:
        # 从外部接口获取科研项目相关信息
        project_info_url = f"http://192.168.2.188:6001/api/keyanexport/keyanprojectinfo?projectNumber={projectNumber}"
        logging.info(f"Fetching data from external API: {project_info_url}")
        project_info_response = requests.get(project_info_url, timeout=10)

        if project_info_response.status_code != 200:
            logging.error(f"Failed to fetch data, status code: {project_info_response.status_code}")
            raise HTTPException(status_code=500, detail="Failed to fetch data from source API")

        raw_data = project_info_response.json()
        logging.info(f"Successfully fetched data: {raw_data}")

        # 获取文件ID列表
        SysFileIds = [file_info['SysFileId'] for file_info in raw_data['Result']['Files']]

        project_file_url = f"http://192.168.2.188:6001/api/keyanexport/download"

        for SysFileId in SysFileIds:
            processor = FileProcessor(project_file_url, SysFileId, projectNumber)
            final_file = processor.download_and_process_file()

            # 处理文件分割
            if final_file.endswith(".docx"):
                documents = GeoReportSplitter().split_document(final_file)
            elif final_file.endswith(".pdf"):
                logging.warning(f"PDF processing not implemented: {final_file}")
                documents = generate_placeholder_documents(final_file)
            else:
                logging.error(f"Invalid file type: {final_file}")
                continue

        logging.info("Parsing data...")
        parsed_data = parse_keyanexport_data(raw_data, documents)
        logging.info(f"Parsed data: {parsed_data}")

        mappings = {
            "properties": {
                "ProjectNumber": {"type": "keyword"},
                "ProjectName": {"type": "keyword"},
                "Files": {
                    "type": "nested",
                    "properties": {
                        "Filehead": {"type": "keyword"},
                        "Filetype": {"type": "keyword"},
                        "page_content": {"type": "text"},
                        "metadata": {"type": "object", "enabled": True}
                    }
                }
            }
        }
        settings = {"number_of_shards": 1, "number_of_replicas": 1}

        logging.info("Writing data to Elasticsearch...")
        es_response = es_service.write_bulk(data=parsed_data, index="keyanexport", mappings=mappings, settings=settings)
        logging.info(f"Elasticsearch response: {es_response}")

        query_uri = f"http://192.168.154.26:9200/keyanexport/_search?q=ProjectNumber:{projectNumber}"
        return {"status": "success", "messages": {"datas": query_uri}}
    except Exception as e:
        logging.error(f"Error occurred: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

# PDF 占位符生成函数
def generate_placeholder_documents(file_path: str) -> list:
    """
    为 PDF 文件生成占位符文档
    :param file_path: PDF 文件路径
    :return: 占位符文档列表
    """
    return [
        Document(
            page_content="This is a placeholder for PDF content.",
            metadata={"file_path": file_path, "type": "PDF", "notes": "Placeholder data for debugging."}
        )
    ]