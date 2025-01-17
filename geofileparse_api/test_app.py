import os
import requests
import urllib.parse
import magic
import win32com.client  # 只在 Windows 环境使用

def save_file(file_content, file_name):
    """
    保存文件到本地
    :param file_content: 二进制内容
    :param file_name: 保存的文件名
    """
    file_path = os.path.join(os.getcwd(), file_name)  # 保存到当前工作目录
    with open(file_path, "wb") as f:
        f.write(file_content)
    return file_path

def is_docx(file_path):
    """
    判断文件是否为 .docx 格式
    :param file_path: 文件路径
    :return: 是否为 .docx 格式 (True/False)
    """
    mime = magic.Magic(mime=True)
    mime_type = mime.from_file(file_path)
    return mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

def convert_doc_to_docx(input_path):
    """
    将 .doc 文件转换为 .docx 文件
    :param input_path: 输入文件路径
    :return: 转换后的文件路径
    """
    if not input_path.lower().endswith(".doc"):
        raise ValueError("输入文件不是 .doc 格式")

    output_path = input_path + "x"  # 转换后的文件名
    word = win32com.client.Dispatch("Word.Application")
    try:
        # 确保文件路径没有问题
        print(f"尝试打开文件: {input_path}")
        doc = word.Documents.Open(os.path.abspath(input_path))  # 使用绝对路径
        doc.SaveAs(output_path, FileFormat=16)  # 16 表示 .docx 格式
        doc.Close()
    finally:
        word.Quit()
    
    return output_path

def process_file(file_content, file_name):
    """
    处理文件，根据文件类型转换或直接使用
    :param file_content: 文件二进制内容
    :param file_name: 原始文件名
    """
    # 保存文件到本地
    file_path = save_file(file_content, file_name)

    # 判断文件类型
    if not is_docx(file_path):
        print(f"{file_name} 不是 docx 文件，正在转换...")
        new_file_path = convert_doc_to_docx(file_path)
        os.remove(file_path)  # 删除原始 .doc 文件
        print(f"转换完成，新文件: {new_file_path}")
        return new_file_path  # 返回新的文件路径
    else:
        print(f"{file_name} 已经是 docx 文件，无需转换。")
        return file_path

def extract_filename_from_headers(headers):
    """
    从响应头中提取并解码文件名
    :param headers: HTTP 响应头
    :return: 解码后的文件名
    """
    content_disposition = headers.get("Content-Disposition", "")
    if "filename*" in content_disposition:
        # 提取并解码文件名
        filename_part = content_disposition.split("filename*=")[1]
        filename = urllib.parse.unquote(filename_part.split("''")[-1])  # 解码并获取文件名
    elif "filename=" in content_disposition:
        # 如果没有 filename*，则使用 filename 部分
        filename = content_disposition.split("filename=")[1].strip('"')
    else:
        filename = "downloaded_file"
    
    return filename

# 示例：处理从接口获取的文件
if __name__ == "__main__":
    # 模拟从接口拉取文件
    file_id = "16636858969286"
    api_url = f"http://192.168.2.188:6001/api/keyanexport/download?id={file_id}"
    response = requests.get(api_url, stream=True)

    if response.status_code == 200:
        # 从响应头提取正确的文件名
        original_file_name = extract_filename_from_headers(response.headers)

        # 处理文件
        final_file_path = process_file(response.content, original_file_name)
        print(f"最终处理后的文件路径: {final_file_path}")
    else:
        print(f"文件下载失败，状态码: {response.status_code}")
