import os
import sys
sys.path.append(os.getcwd())

from geofileparse_api.utils.process_file import FileProcessor
from geofileparse_api.utils.docs import GeoReportSplitter


# 使用示例
if __name__ == "__main__":
    # 配置 API 接口和文件 ID
    api_url = "http://192.168.2.188:6001/api/keyanexport/download"
    file_id = "16636711373382"

    # 创建文件处理对象并开始处理
    processor = FileProcessor(api_url, file_id, "2023-KY-031")
    final_file_path = processor.download_and_process_file()

    print(final_file_path)

    # documents =  GeoReportSplitter().split_document(final_file_path)

    # for doc in documents:
    #     print(type(doc))