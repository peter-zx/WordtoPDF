"""
文档格式转换服务
支持将RTF、ODT、HTML等格式转换为DOCX
"""

import os
import shutil
from typing import List, Dict, Any, Optional
from pathlib import Path


class FormatConverterService:
    """文档格式转换服务"""

    # 支持的输入格式
    SUPPORTED_INPUT_FORMATS = {
        'rtf': 'Rich Text Format',
        'odt': 'OpenDocument Text',
        'html': 'HTML Document',
        'htm': 'HTML Document',
        'txt': 'Plain Text'
    }

    # 输出格式
    OUTPUT_FORMAT = 'docx'

    @staticmethod
    def convert_file(input_path: str, output_dir: str) -> Dict[str, Any]:
        """
        转换单个文件为DOCX格式

        Args:
            input_path: 输入文件路径
            output_dir: 输出目录

        Returns:
            转换结果
        """
        result = {
            'input_file': input_path,
            'success': False,
            'output_file': None,
            'error': None
        }

        try:
            # 获取文件信息
            file_name = os.path.basename(input_path)
            name, ext = os.path.splitext(file_name)
            ext = ext.lower().replace('.', '')

            # 检查是否支持该格式
            if ext not in FormatConverterService.SUPPORTED_INPUT_FORMATS:
                result['error'] = f"不支持的文件格式: {ext}"
                return result

            # 创建输出目录
            os.makedirs(output_dir, exist_ok=True)

            # 输出文件路径
            output_path = os.path.join(output_dir, f"{name}.docx")

            # 根据不同格式使用不同的转换方法
            if ext == 'txt':
                # TXT转DOCX
                FormatConverterService._convert_txt_to_docx(input_path, output_path)
            elif ext in ['html', 'htm']:
                # HTML转DOCX
                FormatConverterService._convert_html_to_docx(input_path, output_path)
            elif ext == 'rtf':
                # RTF转DOCX
                FormatConverterService._convert_rtf_to_docx(input_path, output_path)
            elif ext == 'odt':
                # ODT转DOCX
                FormatConverterService._convert_odt_to_docx(input_path, output_path)

            result['success'] = True
            result['output_file'] = output_path

        except Exception as e:
            result['error'] = str(e)

        return result

    @staticmethod
    def _convert_txt_to_docx(input_path: str, output_path: str):
        """TXT转DOCX"""
        try:
            from docx import Document

            # 读取TXT文件
            with open(input_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 创建DOCX文档
            doc = Document()
            doc.add_paragraph(content)
            doc.save(output_path)

        except ImportError:
            raise Exception("需要安装python-docx: pip install python-docx")

    @staticmethod
    def _convert_html_to_docx(input_path: str, output_path: str):
        """HTML转DOCX"""
        try:
            from docx import Document
            from bs4 import BeautifulSoup

            # 读取HTML文件
            with open(input_path, 'r', encoding='utf-8') as f:
                html_content = f.read()

            # 解析HTML
            soup = BeautifulSoup(html_content, 'html.parser')

            # 创建DOCX文档
            doc = Document()

            # 提取文本内容
            for paragraph in soup.find_all(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                text = paragraph.get_text(strip=True)
                if text:
                    doc.add_paragraph(text)

            doc.save(output_path)

        except ImportError:
            raise Exception("需要安装python-docx和beautifulsoup4: pip install python-docx beautifulsoup4")

    @staticmethod
    def _convert_rtf_to_docx(input_path: str, output_path: str):
        """RTF转DOCX"""
        try:
            # 方法1: 使用pypandoc转换(推荐)
            try:
                import pypandoc
                pypandoc.convert_file(input_path, 'docx', outputfile=output_path)
                return
            except Exception:
                pass

            # 方法2: 使用pywin32转换(Windows)
            try:
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False

                # 打开RTF文件
                doc = word.Documents.Open(input_path)
                # 保存为DOCX
                doc.SaveAs(output_path, FileFormat=16)  # 16 = wdFormatXMLDocument
                doc.Close()
                word.Quit()
                return
            except Exception:
                pass

            # 方法3: 使用striprtf提取文本
            try:
                from striprtf.striprtf import rtf_to_text

                # 读取RTF文件
                with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
                    rtf_content = f.read()

                # 提取文本
                text = rtf_to_text(rtf_content)

                # 创建DOCX文档
                from docx import Document
                doc = Document()
                doc.add_paragraph(text)
                doc.save(output_path)
                return
            except Exception:
                pass

            # 如果所有方法都失败,抛出异常
            raise Exception("RTF转换失败: 请安装pypandoc或pywin32或striprtf")

        except Exception as e:
            raise Exception(f"RTF转换失败: {str(e)}")

    @staticmethod
    def _convert_odt_to_docx(input_path: str, output_path: str):
        """ODT转DOCX"""
        try:
            import pypandoc

            # 使用pypandoc转换
            pypandoc.convert_file(input_path, 'docx', outputfile=output_path)

        except ImportError:
            raise Exception("需要安装pypandoc: pip install pypandoc")
        except Exception as e:
            raise Exception(f"ODT转换失败: {str(e)}")

    @staticmethod
    def batch_convert(input_files: List[str], output_dir: str) -> List[Dict[str, Any]]:
        """
        批量转换文件

        Args:
            input_files: 输入文件列表
            output_dir: 输出目录

        Returns:
            转换结果列表
        """
        results = []

        for input_file in input_files:
            result = FormatConverterService.convert_file(input_file, output_dir)
            results.append(result)

        return results

    @staticmethod
    def get_supported_formats() -> Dict[str, str]:
        """
        获取支持的格式列表

        Returns:
            支持的格式字典
        """
        return FormatConverterService.SUPPORTED_INPUT_FORMATS
