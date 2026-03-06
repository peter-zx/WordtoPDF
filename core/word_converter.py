# -*- coding: utf-8 -*-
"""Word转PDF模块"""
import os


class WordConverter:
    """Word转PDF转换器"""

    def __init__(self):
        self.comtypes_available = False
        self._check_dependencies()

    def _check_dependencies(self):
        """检查依赖是否可用"""
        try:
            import comtypes.client
            import pythoncom
            self.comtypes_available = True
        except ImportError:
            self.comtypes_available = False

    def is_available(self):
        """检查是否可用"""
        return self.comtypes_available

    def convert_folder(self, source_folder, output_folder, keep_structure=True, auto_wrap_folder=True):
        """转换文件夹中的Word文档"""
        if not self.comtypes_available:
            raise ImportError("需要安装 pywin32 和 comtypes 库")

        if not os.path.exists(source_folder):
            raise FileNotFoundError(f"源文件夹不存在: {source_folder}")

        import comtypes.client
        import pythoncom
        pythoncom.CoInitialize()

        results = []
        success = 0
        fail = 0
        
        # 自动创建顶层文件夹
        if auto_wrap_folder:
            source_folder_name = os.path.basename(source_folder.rstrip(os.sep))
            output_folder = os.path.join(output_folder, f"{source_folder_name}_PDF输出")
            results.append(f"📁 输出文件夹: {output_folder}")

        try:
            for root, dirs, files in os.walk(source_folder):
                for file in files:
                    ext = os.path.splitext(file)[1].lower()
                    if ext in ['.docx', '.doc']:
                        try:
                            source_path = os.path.join(root, file)

                            if keep_structure:
                                rel_path = os.path.relpath(root, source_folder)
                                dest_folder = os.path.join(output_folder, rel_path)
                            else:
                                dest_folder = output_folder

                            os.makedirs(dest_folder, exist_ok=True)

                            pdf_name = os.path.splitext(file)[0] + ".pdf"
                            dest_path = os.path.join(dest_folder, pdf_name)

                            # 转换
                            word = comtypes.client.CreateObject('Word.Application')
                            word.Visible = False
                            doc = word.Documents.Open(source_path)
                            doc.SaveAs(dest_path, FileFormat=17)
                            doc.Close()
                            word.Quit()

                            success += 1
                            
                            # 显示相对路径，更清晰
                            if keep_structure:
                                rel_dest = os.path.relpath(dest_path, output_folder)
                                results.append(f"✓ {file} -> {rel_dest}")
                            else:
                                results.append(f"✓ {file} -> {pdf_name}")

                        except Exception as e:
                            fail += 1
                            results.append(f"✗ {file} 失败: {str(e)}")

        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

        # 添加总结信息
        if auto_wrap_folder:
            results.append(f"\n📊 转换统计:")
            results.append(f"   输出位置: {output_folder}")
            results.append(f"   成功: {success} 个文件")
            results.append(f"   失败: {fail} 个文件")
            results.append(f"   总计: {success + fail} 个文件")

        return results, success, fail
