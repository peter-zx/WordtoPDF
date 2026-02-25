"""
批量DOCX转PDF服务 - 多方案支持版本
支持：LibreOffice、CloudConvert API
"""

import os
import sys
import time
import subprocess
import shutil
from typing import List, Tuple, Dict
from pathlib import Path
from multiprocessing import Pool, cpu_count


class BatchPDFService:
    """批量DOCX转PDF服务 - 多方案支持"""

    # 转换模式
    MODE_LIBREOFFICE = "libreoffice"
    MODE_CLOUDCONVERT = "cloudconvert"
    MODE_AUTO = "auto"  # 自动选择

    @staticmethod
    def scan_folder_structure(folder_path: str) -> Dict:
        """扫描文件夹结构,识别所有DOCX文件"""
        structure = {
            "path": folder_path,
            "name": os.path.basename(folder_path),
            "type": "folder",
            "children": [],
            "docx_files": []
        }

        try:
            items = sorted(os.listdir(folder_path))

            for item in items:
                item_path = os.path.join(folder_path, item)

                if os.path.isdir(item_path):
                    child_structure = BatchPDFService.scan_folder_structure(item_path)
                    structure["children"].append(child_structure)

                elif item.lower().endswith('.docx'):
                    structure["docx_files"].append({
                        "name": item,
                        "path": item_path,
                        "relative_path": os.path.relpath(item_path, folder_path)
                    })

        except Exception as e:
            print(f"扫描文件夹失败: {str(e)}")

        return structure

    @staticmethod
    def get_all_docx_files(structure: Dict) -> List[Dict]:
        """从文件夹结构中获取所有DOCX文件"""
        files = []
        files.extend(structure.get("docx_files", []))
        for child in structure.get("children", []):
            files.extend(BatchPDFService.get_all_docx_files(child))
        return files

    @staticmethod
    def check_libreoffice():
        """检查LibreOffice是否可用"""
        try:
            result = subprocess.run(
                ['soffice', '--version'],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            return result.returncode == 0
        except:
            return False

    @staticmethod
    def convert_with_libreoffice(input_path: str, output_path: str, max_retries: int = 3) -> Tuple[bool, str]:
        """使用LibreOffice转换文件"""
        for attempt in range(max_retries):
            try:
                cmd = [
                    'soffice',
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', os.path.dirname(output_path),
                    input_path
                ]

                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=60,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )

                if result.returncode == 0:
                    if os.path.exists(output_path):
                        return (True, "")

                    # LibreOffice可能使用了不同的文件名
                    pdf_files = [f for f in os.listdir(os.path.dirname(output_path))
                               if f.endswith('.pdf')]
                    if pdf_files:
                        generated_pdf = os.path.join(os.path.dirname(output_path), pdf_files[0])
                        if generated_pdf != output_path:
                            shutil.move(generated_pdf, output_path)
                        return (True, "")

                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                else:
                    return (False, f"LibreOffice转换失败: {result.stderr}")

            except subprocess.TimeoutExpired:
                if attempt < max_retries - 1:
                    time.sleep(3)
                    continue
                else:
                    return (False, "转换超时")

            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                else:
                    return (False, str(e))

        return (False, "未知错误")

    @staticmethod
    def convert_with_cloudconvert(input_path: str, output_path: str, api_key: str) -> Tuple[bool, str]:
        """使用CloudConvert API转换文件（云端方案）"""
        try:
            import requests

            # 创建转换任务
            response = requests.post(
                'https://api.cloudconvert.com/v2/convert',
                headers={
                    'Authorization': f'Bearer {api_key}',
                    'Content-Type': 'application/json'
                },
                json={
                    'operation': 'convert',
                    'input': 'upload',
                    'input_format': 'docx',
                    'output_format': 'pdf',
                    'file': open(input_path, 'rb')
                },
                timeout=300
            )

            if response.status_code != 202:
                return (False, f"CloudConvert API错误: {response.text}")

            # 等待转换完成
            task_id = response.json().get('id')
            while True:
                status_response = requests.get(
                    f'https://api.cloudconvert.com/v2/tasks/{task_id}',
                    headers={'Authorization': f'Bearer {api_key}'},
                    timeout=60
                )
                status_data = status_response.json()

                if status_data.get('status') == 'finished':
                    # 下载结果
                    download_url = status_data.get('output', {}).get('url')
                    if download_url:
                        file_response = requests.get(download_url, timeout=60)
                        with open(output_path, 'wb') as f:
                            f.write(file_response.content)
                        return (True, "")
                    else:
                        return (False, "无法获取下载链接")

                elif status_data.get('status') == 'failed':
                    return (False, f"转换失败: {status_data.get('message')}")

                elif status_data.get('status') == 'error':
                    return (False, f"转换错误: {status_data.get('message')}")

                time.sleep(2)

        except ImportError:
            return (False, "需要安装requests库: pip install requests")
        except Exception as e:
            return (False, f"CloudConvert API错误: {str(e)}")

    @staticmethod
    def batch_convert_with_structure(
        structure: Dict,
        output_base_dir: str,
        progress_callback=None,
        mode: str = MODE_AUTO,
        api_key: str = None
    ) -> List[dict]:
        """
        批量转换DOCX文件,保持文件夹结构

        Args:
            structure: 文件夹结构
            output_base_dir: 输出基础目录
            progress_callback: 进度回调函数
            mode: 转换模式 (libreoffice/cloudconvert/auto)
            api_key: CloudConvert API密钥（仅在cloudconvert模式下需要）

        Returns:
            转换结果列表
        """
        # 自动选择模式
        if mode == BatchPDFService.MODE_AUTO:
            if BatchPDFService.check_libreoffice():
                mode = BatchPDFService.MODE_LIBREOFFICE
                print("使用LibreOffice模式")
            else:
                mode = BatchPDFService.MODE_CLOUDCONVERT
                print("LibreOffice不可用，使用CloudConvert模式")

        # 检查CloudConvert API密钥
        if mode == BatchPDFService.MODE_CLOUDCONVERT and not api_key:
            raise Exception("CloudConvert模式需要提供API密钥")

        # 获取所有DOCX文件
        all_files = BatchPDFService.get_all_docx_files(structure)
        total = len(all_files)

        if total == 0:
            return []

        results = []
        top_folder_name = structure["name"]
        top_output_dir = os.path.join(output_base_dir, top_folder_name)
        os.makedirs(top_output_dir, exist_ok=True)

        # 根据模式选择转换方法
        if mode == BatchPDFService.MODE_LIBREOFFICE:
            # LibreOffice模式 - 使用多进程
            tasks = []
            for file_info in all_files:
                relative_path = file_info["relative_path"]
                relative_dir = os.path.dirname(relative_path)
                output_dir = os.path.join(top_output_dir, relative_dir)
                os.makedirs(output_dir, exist_ok=True)
                input_name = Path(file_info["name"]).stem
                output_file = os.path.join(output_dir, f"{input_name}.pdf")
                tasks.append((file_info["path"], output_file, 3))

            for batch_start in range(0, total, 10):
                batch_end = min(batch_start + 10, total)
                batch_tasks = tasks[batch_start:batch_end]

                try:
                    with Pool(processes=2) as pool:
                        batch_results = pool.map(
                            lambda args: BatchPDFService.convert_with_libreoffice(*args),
                            batch_tasks
                        )

                    for i, (success, error) in enumerate(batch_results):
                        index = batch_start + i + 1
                        file_info = all_files[batch_start + i]

                        result = {
                            "input_file": file_info["path"],
                            "output_file": os.path.join(
                                os.path.dirname(batch_tasks[i][1]),
                                Path(batch_tasks[i][1]).stem + ".pdf"
                            ) if success else None,
                            "relative_path": file_info.get("relative_path", ""),
                            "success": success,
                            "error": error if not success else None
                        }
                        results.append(result)

                        if progress_callback:
                            progress_callback(index, total, result)

                except Exception as e:
                    # 降级为单进程
                    for i, task in enumerate(batch_tasks):
                        index = batch_start + i + 1
                        success, error = BatchPDFService.convert_with_libreoffice(*task)
                        file_info = all_files[batch_start + i]

                        result = {
                            "input_file": file_info["path"],
                            "output_file": task[1] if success else None,
                            "relative_path": file_info.get("relative_path", ""),
                            "success": success,
                            "error": error if not success else None
                        }
                        results.append(result)

                        if progress_callback:
                            progress_callback(index, total, result)

                if batch_end < total:
                    time.sleep(1)

        elif mode == BatchPDFService.MODE_CLOUDCONVERT:
            # CloudConvert模式 - 串行处理（API限制）
            for index, file_info in enumerate(all_files, 1):
                try:
                    relative_path = file_info["relative_path"]
                    relative_dir = os.path.dirname(relative_path)
                    output_dir = os.path.join(top_output_dir, relative_dir)
                    os.makedirs(output_dir, exist_ok=True)
                    input_name = Path(file_info["name"]).stem
                    output_file = os.path.join(output_dir, f"{input_name}.pdf")

                    success, error = BatchPDFService.convert_with_cloudconvert(
                        file_info["path"],
                        output_file,
                        api_key
                    )

                    result = {
                        "input_file": file_info["path"],
                        "output_file": output_file if success else None,
                        "relative_path": relative_path,
                        "success": success,
                        "error": error if not success else None
                    }
                    results.append(result)

                    if progress_callback:
                        progress_callback(index, total, result)

                    # API限制，稍作延迟
                    time.sleep(1)

                except Exception