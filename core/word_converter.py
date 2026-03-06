# -*- coding: utf-8 -*-
"""Word转PDF模块 - 优化版本"""
import os
import sys
import subprocess
import tempfile
import time
import threading


class WordConverter:
    """Word转PDF转换器 - 优化版本"""

    def __init__(self):
        self.win32com_available = False
        self._check_dependencies()

    def _check_dependencies(self):
        """检查依赖是否可用"""
        try:
            import win32com.client
            self.win32com_available = True
        except ImportError:
            self.win32com_available = False

    def is_available(self):
        """检查是否可用"""
        return self.win32com_available
    
    def get_converter_info(self):
        """获取转换器信息"""
        info = {
            "win32com_available": self.win32com_available,
            "recommended_method": "Microsoft Word (win32com)" if self.win32com_available else "无可用转换器"
        }
        
        if self.win32com_available:
            try:
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                word.Quit()
                info["word_available"] = True
            except:
                info["word_available"] = False
        
        return info

    def _convert_with_subprocess(self, input_path, output_path):
        """使用子进程转换单个文件 - 避免进程冲突"""
        # 创建转换脚本
        script_content = f'''
import sys
import os
import win32com.client
import time

try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    input_file = r"{input_path}"
    output_file = r"{output_path}"

    doc = word.Documents.Open(input_file, ReadOnly=True, Visible=False)
    doc.SaveAs(output_file, FileFormat=17)
    doc.Close()
    word.Quit()
    
    # 确保进程完全退出
    time.sleep(1)
    print("SUCCESS")
    sys.exit(0)

except Exception as e:
    print(f"ERROR: {{str(e)}}")
    try:
        if 'word' in locals():
            word.Quit()
    except:
        pass
    sys.exit(1)
'''
        
        # 写入临时脚本
        with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False, encoding='utf-8') as f:
            f.write(script_content)
            script_path = f.name

        try:
            # 运行子进程
            result = subprocess.run(
                [sys.executable, script_path],
                capture_output=True,
                text=True,
                timeout=60,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # 检查输出文件
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                return True, ""
            else:
                return False, result.stderr or "转换失败"

        except subprocess.TimeoutExpired:
            return False, "转换超时"
        except Exception as e:
            return False, str(e)
        finally:
            # 清理临时文件
            try:
                os.remove(script_path)
            except:
                pass

    def convert_folder(self, source_folder, output_folder, keep_structure=True, auto_wrap_folder=True, progress_callback=None):
        """转换文件夹中的Word文档 - 优化版本"""
        if not self.win32com_available:
            raise ImportError("需要安装 pywin32 库")

        if not os.path.exists(source_folder):
            raise FileNotFoundError(f"源文件夹不存在: {source_folder}")

        results = []
        success = 0
        fail = 0
        
        # 自动创建顶层文件夹
        if auto_wrap_folder:
            source_folder_name = os.path.basename(source_folder.rstrip(os.sep))
            output_folder = os.path.join(output_folder, f"{source_folder_name}_PDF输出")
            results.append(f"📁 输出文件夹: {output_folder}")

        # 收集所有Word文件
        word_files = []
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in ['.docx', '.doc']:
                    word_files.append((root, file))

        total_files = len(word_files)
        
        if total_files == 0:
            results.append("⚠️ 没有找到Word文档")
            return results, 0, 0

        # 开始转换
        for index, (root, file) in enumerate(word_files, 1):
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

                # 使用子进程转换
                conversion_success, error_msg = self._convert_with_subprocess(source_path, dest_path)
                
                if conversion_success:
                    success += 1
                    # 显示相对路径，更清晰
                    if keep_structure:
                        rel_dest = os.path.relpath(dest_path, output_folder)
                        results.append(f"✓ [{index}/{total_files}] {file} -> {rel_dest}")
                    else:
                        results.append(f"✓ [{index}/{total_files}] {file} -> {pdf_name}")
                else:
                    fail += 1
                    results.append(f"✗ [{index}/{total_files}] {file} 失败: {error_msg}")

                # 进度回调
                if progress_callback:
                    progress_callback(index, total_files, file, conversion_success, error_msg)

                # 文件间延迟，确保进程完全退出
                time.sleep(0.5)

            except Exception as e:
                fail += 1
                results.append(f"✗ [{index}/{total_files}] {file} 失败: {str(e)}")
                if progress_callback:
                    progress_callback(index, total_files, file, False, str(e))

        # 添加总结信息
        results.append(f"\n📊 转换统计:")
        results.append(f"   输出位置: {output_folder}")
        results.append(f"   成功: {success} 个文件")
        results.append(f"   失败: {fail} 个文件")
        results.append(f"   总计: {success + fail} 个文件")
        results.append(f"   成功率: {(success/total_files*100):.1f}%")

        return results, success, fail
