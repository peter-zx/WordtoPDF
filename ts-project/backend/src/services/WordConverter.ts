import { spawn } from 'child_process';
import fs from 'fs-extra';
import path from 'path';
import { ConversionResult, ProgressInfo } from '../../shared/types.js';

export class WordConverter {
  private maxConcurrent: number;
  
  constructor(maxConcurrent = 2) {
    this.maxConcurrent = maxConcurrent;
  }

  /**
   * 使用LibreOffice转换Word到PDF
   */
  private async convertWithLibreOffice(inputPath: string, outputPath: string): Promise<ConversionResult> {
    return new Promise((resolve) => {
      try {
        // 检查LibreOffice是否安装
        const command = process.platform === 'win32' ? 'soffice' : 'libreoffice';
        
        const args = [
          '--headless',
          '--convert-to', 'pdf',
          '--outdir', path.dirname(outputPath),
          inputPath
        ];

        const process = spawn(command, args);
        
        let stdout = '';
        let stderr = '';
        
        process.stdout.on('data', (data) => {
          stdout += data.toString();
        });
        
        process.stderr.on('data', (data) => {
          stderr += data.toString();
        });
        
        process.on('close', async (code) => {
          if (code === 0) {
            // 检查输出文件
            const expectedPdf = path.join(
              path.dirname(outputPath), 
              path.basename(inputPath, path.extname(inputPath)) + '.pdf'
            );
            
            if (await fs.pathExists(expectedPdf)) {
              // 重命名到目标路径
              await fs.move(expectedPdf, outputPath, { overwrite: true });
              resolve({ success: true, message: '转换成功', inputPath, outputPath });
            } else {
              resolve({ success: false, message: '转换失败，输出文件不存在', inputPath, outputPath });
            }
          } else {
            resolve({ success: false, message: `LibreOffice错误: ${stderr}`, inputPath, outputPath });
          }
        });
        
        process.on('error', (error) => {
          resolve({ success: false, message: `进程启动失败: ${error.message}`, inputPath, outputPath });
        });
        
        // 超时处理
        setTimeout(() => {
          if (!process.killed) {
            process.kill();
            resolve({ success: false, message: '转换超时', inputPath, outputPath });
          }
        }, 30000); // 30秒超时
        
      } catch (error: any) {
        resolve({ success: false, message: `转换异常: ${error.message}`, inputPath, outputPath });
      }
    });
  }

  /**
   * 使用云服务API转换（备用方案）
   */
  private async convertWithCloudService(inputPath: string, outputPath: string): Promise<ConversionResult> {
    // 这里可以实现使用云服务API的转换
    // 例如：Microsoft Graph API、Google Docs API等
    
    return {
      success: false,
      message: '云服务转换暂未实现，请安装LibreOffice',
      inputPath,
      outputPath
    };
  }

  /**
   * 检查转换器可用性
   */
  async checkAvailability(): Promise<{ available: boolean; method: string; message: string }> {
    try {
      const command = process.platform === 'win32' ? 'soffice' : 'libreoffice';
      const process = spawn(command, ['--version']);
      
      return new Promise((resolve) => {
        process.on('close', (code) => {
          if (code === 0) {
            resolve({ 
              available: true, 
              method: 'LibreOffice', 
              message: 'LibreOffice已安装，可以使用本地转换'
            });
          } else {
            resolve({ 
              available: false, 
              method: '无', 
              message: 'LibreOffice未安装，请安装LibreOffice或使用云服务'
            });
          }
        });
        
        process.on('error', () => {
          resolve({ 
            available: false, 
            method: '无', 
            message: 'LibreOffice未安装，请安装LibreOffice或使用云服务'
          });
        });
      });
    } catch (error) {
      return { 
        available: false, 
        method: '无', 
        message: '检查失败，请安装LibreOffice'
      };
    }
  }

  /**
   * 转换单个文件
   */
  async convertFile(inputPath: string, outputPath: string): Promise<ConversionResult> {
    // 确保输出目录存在
    await fs.ensureDir(path.dirname(outputPath));
    
    // 首先尝试LibreOffice
    const result = await this.convertWithLibreOffice(inputPath, outputPath);
    
    if (!result.success) {
      // 如果LibreOffice失败，尝试云服务
      return await this.convertWithCloudService(inputPath, outputPath);
    }
    
    return result;
  }

  /**
   * 批量转换文件夹
   */
  async convertFolder(
    sourceFolder: string, 
    outputFolder: string, 
    keepStructure = true,
    progressCallback?: (progress: ProgressInfo) => void
  ): Promise<{ results: ConversionResult[]; success: number; fail: number }> {
    const results: ConversionResult[] = [];
    let success = 0;
    let fail = 0;

    // 收集所有Word文件
    const wordFiles: Array<{ inputPath: string; outputPath: string; relativePath: string }> = [];
    
    await this.collectWordFiles(sourceFolder, outputFolder, keepStructure, wordFiles);
    
    const totalFiles = wordFiles.length;
    
    if (totalFiles === 0) {
      return { results, success, fail };
    }

    // 使用Promise.all限制并发数
    const batches = [];
    for (let i = 0; i < wordFiles.length; i += this.maxConcurrent) {
      batches.push(wordFiles.slice(i, i + this.maxConcurrent));
    }

    for (const batch of batches) {
      const batchPromises = batch.map(async (file, index) => {
        const currentIndex = wordFiles.indexOf(file) + 1;
        
        // 更新进度
        if (progressCallback) {
          progressCallback({
            current: currentIndex,
            total: totalFiles,
            filename: path.basename(file.inputPath),
            success: false
          });
        }

        const result = await this.convertFile(file.inputPath, file.outputPath);
        
        if (result.success) {
          success++;
        } else {
          fail++;
        }

        // 更新完成进度
        if (progressCallback) {
          progressCallback({
            current: currentIndex,
            total: totalFiles,
            filename: path.basename(file.inputPath),
            success: result.success,
            error: result.message
          });
        }

        results.push(result);
        return result;
      });

      await Promise.all(batchPromises);
    }

    return { results, success, fail };
  }

  /**
   * 递归收集Word文件
   */
  private async collectWordFiles(
    currentFolder: string,
    outputBase: string,
    keepStructure: boolean,
    files: Array<{ inputPath: string; outputPath: string; relativePath: string }>,
    relativePath = ''
  ): Promise<void> {
    try {
      const items = await fs.readdir(currentFolder);
      
      for (const item of items) {
        const fullPath = path.join(currentFolder, item);
        const stats = await fs.stat(fullPath);
        
        if (stats.isDirectory()) {
          const newRelativePath = keepStructure ? path.join(relativePath, item) : '';
          await this.collectWordFiles(fullPath, outputBase, keepStructure, files, newRelativePath);
        } else if (stats.isFile()) {
          const ext = path.extname(item).toLowerCase();
          if (['.doc', '.docx'].includes(ext)) {
            const outputPath = keepStructure 
              ? path.join(outputBase, relativePath, path.basename(item, ext) + '.pdf')
              : path.join(outputBase, path.basename(item, ext) + '.pdf');
            
            files.push({
              inputPath: fullPath,
              outputPath,
              relativePath: path.join(relativePath, item)
            });
          }
        }
      }
    } catch (error) {
      console.error(`扫描文件夹错误: ${currentFolder}`, error);
    }
  }
}

export default WordConverter;