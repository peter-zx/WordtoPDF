"""
文档格式转换服务
支持将RTF、ODT、HTML等格式转换为DOCX
"""

import os
from typing import List, Dict, Any


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

            # 尝试不同编码读取TXT文件
            encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030']
            content = None

            for encoding in encodings:
                try:
                    with open(input_path, 'r', encoding=encoding) as f:
                        content = f.read()
                    break
                except UnicodeDecodeError:
                    continue

            if content is None:
                raise Exception("无法识别文件编码")

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

            # 尝试不同编码读取HTML文件
            encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030']
            html_content = None

            for encoding in encodings:
                try:
                    with open(input_path, 'r', encoding=encoding) as f:
                        html_content = f.read()
                    break
                except UnicodeDecodeError:
                    continue

            if html_content is None:
                raise Exception("无法识别文件编码")

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
        """RTF转DOCX - 使用Word或WPS打开并另存为DOCX"""
        try:
            import win32com.client

            # 转换为绝对路径
            input_path = os.path.abspath(input_path)
            output_path = os.path.abspath(output_path)

            # 尝试使用WPS（因为系统优先调用WPS）
            try:
                # WPS的COM对象名称
                wps = win32com.client.Dispatch("Kwps.Application")
                wps.Visible = False
                wps.DisplayAlerts = False

                # 打开RTF文件，指定编码参数
                doc = wps.Documents.Open(
                    input_path,
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                    Encoding=936  # 936 = GBK编码
                )

                # 另存为DOCX格式
                doc.SaveAs(output_path, FileFormat=16)

                # 关闭文档和WPS
                doc.Close()
                wps.Quit()
                return

            except Exception as wps_error:
                # WPS失败，尝试Microsoft Word
                try:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = False

                    # 打开RTF文件（Word会自动检测编码）
                    doc = word.Documents.Open(
                        input_path,
                        ConfirmConversions=False,
                        ReadOnly=False,
                        AddToRecentFiles=False,
                        Visible=False,
                        Encoding=936  # 指定GBK编码
                    )

                    # 另存为DOCX格式
                    doc.SaveAs(output_path, FileFormat=16)  # 16 = wdFormatXMLDocument (DOCX)

                    # 关闭文档和Word
                    doc.Close()
                    word.Quit()
                    return

                except Exception as word_error:
                    raise Exception(f"RTF转换失败: WPS错误-{str(wps_error)}, Word错误-{str(word_error)}")

        except Exception as e:
            raise Exception(f"RTF转换失败: 请确保已安装Microsoft Word或WPS Office。错误信息: {str(e)}")

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
