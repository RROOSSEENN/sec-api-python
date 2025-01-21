import os
import requests
import urllib.parse
import magic
import win32com.client  # 只在 Windows 环境使用


class FileProcessor:
    def __init__(self, api_url, file_id, project_number):
        """
        初始化文件处理类
        :param api_url: 文件下载的 API 接口地址
        :param file_id: 文件的唯一标识
        :param project_number: 项目编号，用于生成保存目录
        """
        self.api_url = api_url
        self.file_id = file_id
        self.project_number = project_number

    def extract_filename_from_headers(self, headers):
        """
        从响应头中提取并解码文件名
        :param headers: HTTP 响应头
        :return: 解码后的文件名
        """
        content_disposition = headers.get("Content-Disposition", "")
        if "filename*" in content_disposition:
            filename_part = content_disposition.split("filename*=")[1]
            filename = urllib.parse.unquote(filename_part.split("''")[-1])  # 解码并获取文件名
        elif "filename=" in content_disposition:
            filename = content_disposition.split("filename=")[1].strip('"')
        else:
            filename = "downloaded_file"
        
        return filename

    def save_file(self, file_content, file_name):
        """
        保存文件到本地
        :param file_content: 二进制内容
        :param file_name: 保存的文件名
        :return: 保存的文件路径
        """
        # 根据 project_number 创建目录
        target_dir = os.path.join(r"E:\R\routine\2025\1\DocxCache", str(self.project_number))
        
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)  # 如果目录不存在，创建目录
        
        # 构建文件完整路径
        file_path = os.path.join(target_dir, file_name)  # 合并文件名
        
        # 如果文件已存在，则跳过保存
        if os.path.exists(file_path):
            print(f"文件 {file_name} 已经存在，跳过保存。")
            return file_path  # 返回现有的文件路径
        
        # 如果文件不存在，则保存文件
        with open(file_path, "wb") as f:
            f.write(file_content)
        
        return file_path

    def is_docx(self, file_path):
        """
        判断文件是否为 .docx 格式
        :param file_path: 文件路径
        :return: 是否为 .docx 格式 (True/False)
        """
        mime = magic.Magic(mime=True)
        mime_type = mime.from_file(file_path)
        return mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    def is_pdf(self, file_path):
        """
        判断文件是否为 .pdf 格式
        :param file_path: 文件路径
        :return: 是否为 .pdf 格式 (True/False)
        """
        mime = magic.Magic(mime=True)
        mime_type = mime.from_file(file_path)
        return mime_type == "application/pdf"

    def convert_doc_to_docx(self, input_path):
        """
        将 .doc 文件转换为 .docx 文件
        :param input_path: 输入文件路径
        :return: 转换后的文件路径
        """
        if not input_path.lower().endswith(".doc"):
            raise ValueError("输入文件不是 .doc 格式")

        output_path = input_path + "x"  # 转换后的文件名
        if os.path.exists(output_path):
            print(f"文件 {output_path} 已经存在，跳过转换。")
            return output_path  # 如果转换后的文件已经存在，直接返回原路径
        
        word = win32com.client.Dispatch("Word.Application")
        try:
            print(f"尝试打开文件: {input_path}")
            doc = word.Documents.Open(os.path.abspath(input_path))  # 使用绝对路径
            doc.SaveAs(output_path, FileFormat=16)  # 16 表示 .docx 格式
            doc.Close()
        finally:
            word.Quit()
        
        return output_path

    def process_file(self, file_content, file_name):
        """
        处理文件，根据文件类型转换或直接使用
        :param file_content: 文件二进制内容
        :param file_name: 原始文件名
        :return: 处理后的文件路径
        """
        # 保存文件到本地
        file_path = self.save_file(file_content, file_name)

        # 判断文件类型并进行相应处理
        if file_path.lower().endswith(".doc"):
            print(f"{file_name} 不是 docx 文件，正在转换...")
            new_file_path = self.convert_doc_to_docx(file_path)
            os.remove(file_path)  # 删除原始 .doc 文件
            print(f"转换完成，新文件: {new_file_path}")
            return new_file_path  # 返回新的文件路径
        elif file_path.lower().endswith(".pdf"):
            print(f"{file_name} 是 PDF 文件，无需转换。")
            return file_path
        elif not self.is_docx(file_path):
            print(f"{file_name} 不支持的文件类型，跳过处理。")
            return None
        else:
            print(f"{file_name} 已经是 docx 文件，无需转换。")
            return file_path

    def download_and_process_file(self):
        """
        从接口拉取文件并处理
        :return: 处理后的文件路径
        """
        # 构建下载 URL
        download_url = f"{self.api_url}?id={self.file_id}"
        response = requests.get(download_url, stream=True)

        if response.status_code == 200:
            # 从响应头提取正确的文件名
            original_file_name = self.extract_filename_from_headers(response.headers)

            # 处理文件
            final_file_path = self.process_file(response.content, original_file_name)
            print(f"最终处理后的文件路径: {final_file_path}")
            return final_file_path
        else:
            print(f"文件下载失败，状态码: {response.status_code}")
            return None

# 使用示例
if __name__ == "__main__":
    # 配置 API 接口和文件 ID
    api_url = "http://192.168.2.188:6001/api/keyanexport/download"
    file_id = "1872438665989459968"
    project_number = "2023-KY-031"

    # 创建文件处理对象并开始处理
    processor = FileProcessor(api_url, file_id, project_number)
    final_file_path = processor.download_and_process_file()

    if final_file_path:
        print(f"文件处理成功，文件路径: {final_file_path}")
    else:
        print("文件处理失败。")