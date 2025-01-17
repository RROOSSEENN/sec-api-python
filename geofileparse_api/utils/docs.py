
import os
import re
import yaml
import docx
from docx import Document as DocxDocument

from collections import Counter
from typing import Any, List, Optional

from langchain.docstore.document import Document

neg_pattern_file = os.path.join(os.path.dirname(__file__), 'neg_pattern.yaml')
# 加载YAML文件
with open(neg_pattern_file, 'r', encoding="utf-8") as file:
    data = yaml.safe_load(file)
    NEG_PATTERNS = data['neg_patterns']
    NEG_KEYWORD_PATTERNS = data['keyword_patterns']

class GeoReportSplitter:
    def __init__(self, separators: Optional[List[str]] = None,):
        self.separators = separators

    def read_word_document(self, file_path):
        """从Word文档读取内容并转换为纯文本，同时保留标题信息，处理标题自动编号无法读取的情况"""

        doc = DocxDocument(file_path)
        full_text = []
        for para in doc.paragraphs:
            if para.style.name.startswith('Heading'):
                full_text.append(f'\n{para.style.name} {para.text}\n')
            else:
                full_text.append(para.text)
        return '\n'.join(full_text)
    
    def _convert_table_to_markdown(self, table: docx.table.Table) -> str:
        """Helper function to convert a docx table to a Markdown format string."""
        markdown = ""
        if table.rows:
            markdown += "| " + " | ".join(cell.text.strip() for cell in table.rows[0].cells) + " |\n"
            markdown += "| " + " | ".join(['---']*len(table.rows[0].cells)) + " |\n"
            for row in table.rows[1:]:
                row_text = "| " + " | ".join(cell.text.strip() for cell in row.cells) + " |\n"
                markdown += row_text
        return markdown
    
    def _get_body_font_size(self, doc):
        body_font_sizes = []

        for para in doc.paragraphs:
            # 排除标题文本
            if para.style.name.startswith('Heading'):  
                continue
            for run in para.runs:
                # 使用抽样率确定是否处理该文本运行
                if not isinstance(run, docx.text.run.Run):
                    continue
                font_size = run.font.size
                if font_size is None:
                    continue
                body_font_sizes.append(font_size.pt)

        if body_font_sizes:  # 如果存在正文文本
            average_font_size = sum(body_font_sizes) / len(body_font_sizes)
            if average_font_size > 12:
                return None
            return average_font_size
        else:
            return None  # 如果没有正文文本，则返回None
    
    def _get_paragraph_font_size(self, para: docx.text.paragraph.Paragraph):
        # total_runs = len(para.runs)
        font_sizes = [run.font.size.pt for run in para.runs if isinstance(run, docx.text.run.Run) and run.font.size is not None]
        # 统计每个字体大小在文本段中出现的次数
        if len(font_sizes) == 0:
            return None
        font_size_counts = Counter(font_sizes)
        
        # 找到出现次数最多的字体大小及其出现次数
        most_common_font_size, _ = font_size_counts.most_common(1)[0]
        
        return most_common_font_size
        
    def _get_table_title(self, paragraph: docx.text.paragraph.Paragraph) -> bool:
        table_title_pattern = r"(表\d+(\.\d+)?\s+.*|.*\s+表\d+(\.\d+)?)"
        # 首先检查文本长度是否适中，过长的文本可能不是标题
        if len(paragraph.text) > 50:
            return None
        
        # 检查文本是否包含表标题的常见模式
        if re.search(table_title_pattern, paragraph.text):
            return paragraph.text
    
    def _is_title(self, paragraph: docx.text.paragraph.Paragraph) -> bool:
        title_pattern = r'^\d'
        if paragraph.style.name.startswith('Heading'):
            if re.match(title_pattern, paragraph.text):
                return True
        return False
    
    def _get_title(self, paragraph: docx.text.paragraph.Paragraph, body_font_size: float = 12.0) -> Optional[str]:
        
        text = paragraph.text.strip()
        is_short_text = len(text) >= 30
        if is_short_text:
            return None, None
        
        para_font_size = self._get_paragraph_font_size(paragraph)
        #如果字体小于正文就直接返回
        if para_font_size is not None and para_font_size < body_font_size:
            return None, None
        
        title_pattern = r'^\d'
        # header_level_pattern = r'^(\d+)[\.\。\·、]'
        header_level_pattern = r'^(\d+(?:[\.、]\d+)*)[\.、]?'
        toc_pattern = r'^\d+(\.\d+)*[\s\S]*?(\d+)$'  # 用于匹配目录条目
        neg_pattern = r'^(\d+[)）>#].*|\d+.*?(\d+层|\d+F|' + '|'.join(NEG_PATTERNS) + r')|\d+(\.\d+)?-\d+(\.\d+)?)'
        neg_keyword_pattern = r"\b(" + "|".join(map(re.escape, NEG_KEYWORD_PATTERNS)) + r")\b"


        is_heading = paragraph.style.name.startswith('Heading')
        is_not_toc = not re.match(toc_pattern, text)
        is_not_neg = not re.match(neg_pattern, text)
        is_not_neg_keyword = not re.search(neg_keyword_pattern, text)
        other_situation = is_not_toc and is_not_neg and re.match(title_pattern, paragraph.text) and is_not_neg_keyword

        # if is_heading:
            # header_level = int(paragraph.style.name.split()[1])
            # if re.match(title_pattern, paragraph.text):
                # return paragraph.text, header_level
        if is_heading or other_situation:
            match = re.search(header_level_pattern, paragraph.text)
            if match:
                title_number = match.group(1)
                header_level = title_number.count('.') + title_number.count('、') + 1
                return paragraph.text, header_level
            elif int(paragraph.style.name.split()[1]):
                header_level = int(paragraph.style.name.split()[1])
                return paragraph.text, header_level
        return None, None
    
    def get_table_contents(self, file_path):
        doc = DocxDocument(file_path)
        toc_pattern = r'^\d+(\.\d+)*[\s\S]*?(\d+)$'  # 用于匹配目录条目
        table_contents = []
        for element in doc.element.body:
            if element.tag.endswith('p'):  # 如果是段落
                para = docx.text.paragraph.Paragraph(element, doc)
                if re.match(toc_pattern, para.text):
                    table_contents.append(para.text)
        
        return table_contents
    
    def split_document(self, file_path, max_content_length=256):
        """根据Word中的标题样式来分割文档"""
        doc = DocxDocument(file_path)
        documents = []
        current_headers = []
        content = ""
        current_level = 0  # 当前标题级别
        prev_paragraph_text = None  # 前一个段落文本，可能是表的标题

        body_font_size = self._get_body_font_size(doc)
        if not body_font_size:
            body_font_size = 12

        def commit_content():
            nonlocal content
            if content:
                documents.append(Document(page_content=content, metadata={'title': current_headers.copy()}))
                content = ""
            return content

        for element in doc.element.body:
            if element.tag.endswith('p'):  # 如果是段落
                para = docx.text.paragraph.Paragraph(element, doc)
                _, header_level = self._get_title(para, body_font_size=body_font_size)
                if header_level:
                    # header_level = int(para.style.name.split()[1])
                    # 当遇到新标题时, 先提交当前正文
                    content = commit_content() 
                    if header_level > current_level:
                        current_headers.append(para.text)
                    elif header_level == current_level:
                        current_headers[-1] = para.text
                    else:
                        current_headers = current_headers[:header_level - 1]
                        current_headers.append(para.text)
                    current_level = header_level
                else:
                    content += para.text + "\n"
                    if len(content) >= max_content_length:
                        content = commit_content()
                # 更新表格标题的潜在文本
                prev_paragraph_text = self._get_table_title(para)

            elif element.tag.endswith('tbl'):  # 如果是表格
                table = docx.table.Table(element, doc)
                # print(table)
                table_title = prev_paragraph_text if prev_paragraph_text else None
                # print(table_title)
                # 将表格转换为 Markdown 格式
                table_content = self._convert_table_to_markdown(table)
                documents.append(Document(page_content=table_content, metadata={'title': current_headers.copy(), 'table_title': table_title}))

        # 提交最后的内容
        if content:
            documents.append(Document(page_content=content, metadata={'title': current_headers.copy()}))

        # for idx, doc in enumerate(documents):
        #     print(f"Document {idx + 1} Title:", doc.metadata.get('title'))
        return documents

    def split_text_from_docs(self, docs):
        documents = []
        for doc in docs:
            documents.append(self.split_text_from_doc(doc))
        
        return documents


if __name__ == "__main__":
    # file_path = r"/media/geodataset/report/2009/2009-G-001_11880/勘察报告【2009-G-001】“馨亭小区”A区（三期）商品住宅项目.docx"
    file_path = r"E:\R\workplace\es-api-python\上海市科学技术委员会“扬帆计划”项目申请书-雷丹V5.docx"
    # file_path = r"E:\R\routine\勘察报告【2022-G-132】普陀区桃浦科技智慧城W06-1401单元026-01地块项目.docx"
    # convert_docx_to_markdown(file_path)
    # file_path = r"/media/geodataset/report/2015/2015-G-031-13_22613/勘察报告【2015-G-031-13】滨海新区轨道交通B1线一期工程第一标段——国祥西道站岩土工程勘察详勘报告.docx"
    # file_path = r"/media/geodataset/report/2020/2020-G-006_32062/勘察报告【2020-G-006】G228公路（海丹路-瓦洪公路）给水管排管工程（海丹路～两港西大道）.docx"
    # file_path = r"/media/geodataset/report/2020/2020-G-124_33110/勘察报告【2020-G-124】上海集成电路设计产业园3-4项目.docx"
    # file_path = r"/media/geodataset/report/2020/2020-G-124_33110/勘察报告【2020-G-124】上海集成电路设计产业园3-4项目水文地质勘察报告.docx"
    # file_path = r"/media/geodataset/report/2015/2015-G-177_23464/勘察报告【2015-G-177】上海临港普洛斯国际物流园区G地块物流仓库项目（二期）岩土工程勘察纲要.docx"
    # file_path = r"/media/geodataset/report/2015/2015-G-135_23081/勘察报告【2015-G-135】外运发展浦东空港物流基地.docx"
    # file_path = r"/media/geodataset/report/2015/2015-G-002_22182/勘察报告【2015-G-002-2】上海市轨道交通14号线工程详勘5标武定路站～静安寺站区间.docx"
    # file_path = r"/media/geodataset/report/2015/2015-G-045_22466/勘察报告【2015-G-045】普凯（武汉东西湖）国际物流园项目.docx"
    # file_path = r"Y:\dataset\report\2022\2022-G-185\勘察报告【2022-G-185】南翔镇JDC2-0201单元14-02、14-03地块项目.docx"
    # file_path = r"Y:\dataset\report\2023\2023-G-030\勘察报告【2023-G-030-4】上海浦东国际机场四期扩建工程市政配套工程-综合管廊、出水箱涵.docx"
    # file_path = r"Y:\dataset\report\2022\2022-G-129\勘察报告【2022-G-129】强华股份集成电路核心装备关键新材料生产基地项目.docx"
    # file_path = r"Y:\dataset\report\2023\2023-G-030\勘察报告【2023-G-030-3】上海浦东国际机场四期扩建工程市政配套工程（不含二级排水、能源中心）项目（出租车蓄车场及员工停车库）.docx"
    documents =  GeoReportSplitter().split_document(file_path)
    print(type(documents))