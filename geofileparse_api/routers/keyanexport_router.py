from langchain_core.documents import Document
from fastapi import APIRouter, HTTPException
import requests
import logging
from geofileparse_api.service.elastic_service import ElasticService
from geofileparse_api.utils.docs import GeoReportSplitter
from geofileparse_api.utils.parse_data import parse_keyanexport_data
from geofileparse_api.utils.pdfs import PDFLoader
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
        project_info_url = f"http://192.168.2.188:6001/api/keyanexport/keyanprojectinfo?projectNumber={projectNumber}"
        logging.info(f"Fetching data from external API: {project_info_url}")
        project_info_response = requests.get(project_info_url, timeout=10)

        if project_info_response.status_code != 200:
            logging.error(f"Failed to fetch data, status code: {project_info_response.status_code}")
            raise HTTPException(status_code=500, detail="Failed to fetch data from source API")

        raw_data = project_info_response.json()
        # logging.info(f"Successfully fetched data: {raw_data}")

        project_file_url = f"http://192.168.2.188:6001/api/keyanexport/download"

        # 更新 Files 数据，处理每个文件的分割结果
        for file_info in raw_data['Result']['Files']:
            file_id = file_info["SysFileId"]
            file_name = file_info["FileOldName"]
            processor = FileProcessor(project_file_url, file_id, projectNumber)
            final_file = processor.download_and_process_file()

            if final_file:
                if final_file.endswith(".docx"):
                    documents = GeoReportSplitter().split_document(final_file)
                    file_info["segments"] = [
                        {"page_content": doc.page_content, "metadata": doc.metadata}
                        for doc in documents
                    ]
                elif final_file.lower().endswith(".pdf"):
                    loader = PDFLoader(file_path=final_file, need_ocr=True)
                    documents = loader.load()
                    file_info["segments"] = [
                        {"page_content": doc.page_content, "metadata": doc.metadata}
                        for doc in documents
                    ]

                else:
                    logging.warning(f"跳过不支持的文件类型: {final_file}")
                    file_info["segments"] = []  # 空分割结果
            else:
                logging.warning(f"文件处理失败，跳过: {file_name}")
                file_info["segments"] = []  # 文件处理失败的情况

        logging.info("Parsing data...")
        parsed_data = parse_keyanexport_data(raw_data)
        # logging.info(f"Parsed data: {parsed_data}")

        mappings = {
            "properties": {
                "ProjectNumber": {"type": "keyword"},
                "ProjectName": {"type": "keyword"},
                "Files": {
                    "type": "nested",
                    "properties": {
                        "Filehead": {"type": "keyword"},
                        "Filetype": {"type": "keyword"},
                        "segments": {
                            "type": "nested",
                            "properties": {
                                "page_content": {"type": "text"},
                                "metadata": {"type": "object", "enabled": True}
                            }
                        }
                    }
                }
            }
        }
        settings = {"number_of_shards": 1, "number_of_replicas": 1}

        logging.info("Writing data to Elasticsearch...")
        es_response = es_service.write_bulk(data=parsed_data, index="keyanketi", mappings=mappings, settings=settings)
        logging.info(f"Elasticsearch response: {es_response}")

        query_uri = f"http://192.168.154.26:9200/keyanexport/_search?q=ProjectNumber:{projectNumber}"
        return {"status": "success", "messages": {"datas": query_uri}}

    except Exception as e:
        logging.error(f"Error occurred: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))