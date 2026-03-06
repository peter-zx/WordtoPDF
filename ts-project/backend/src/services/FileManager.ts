import fs from 'fs-extra';
import path from 'path';
import { FileInfo, CopyResult } from '../../shared/types.js';

export class FileManager {
  public scannedFiles: FileInfo[] = [];
  public checkStates: boolean[] = [];

  /**
   * 扫描文件夹中的文件
   */
  async scanFolder(folderPath: string, fileTypes: string[] = ['.docx', '.doc', '.pdf']): Promise<FileInfo[]> {
    this.scannedFiles = [];
    this.checkStates = [];
    
    await this.scanFolderRecursive(folderPath, folderPath, fileTypes);
    return this.scannedFiles;
  }

  /**
   * 递归扫描文件夹
   */
  private async scanFolderRecursive(
    currentPath: string,
    basePath: string,
    fileTypes: string[],
    relativePath = ''
  ): Promise<void> {
    try {
      const items = await fs.readdir(currentPath);
      
      for (const item of items) {
        const fullPath = path.join(currentPath, item);
        const stats = await fs.stat(fullPath);
        
        if (stats.isDirectory()) {
          // 递归扫描子文件夹
          const newRelativePath = path.join(relativePath, item);
          await this.scanFolderRecursive(fullPath, basePath, fileTypes, newRelativePath);
        } else if (stats.isFile()) {
          // 检查文件类型
          const ext = path.extname(item).toLowerCase();
          if (fileTypes.includes(ext)) {
            const fileInfo: FileInfo = {
              name: item,
              path: fullPath,
              size: stats.size,
              sizeStr: this.formatFileSize(stats.size),
              type: ext.substring(1).toUpperCase(), // 移除点并大写
              selected: false
            };
            
            this.scannedFiles.push(fileInfo);
            this.checkStates.push(false);
          }
        }
      }
    } catch (error) {
      console.error(`扫描文件夹错误: ${currentPath}`, error);
    }
  }

  /**
   * 格式化文件大小
   */
  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 B';
    
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  /**
   * 获取选中的文件
   */
  getSelectedFiles(): FileInfo[] {
    return this.scannedFiles.filter((file, index) => this.checkStates[index]);
  }

  /**
   * 获取选中文件数量
   */
  getSelectedCount(): number {
    return this.checkStates.filter(state => state).length;
  }

  /**
   * 全选
   */
  selectAll(): void {
    this.checkStates = this.checkStates.map(() => true);
    this.scannedFiles.forEach((file, index) => {
      file.selected = this.checkStates[index];
    });
  }

  /**
   * 取消全选
   */
  deselectAll(): void {
    this.checkStates = this.checkStates.map(() => false);
    this.scannedFiles.forEach((file, index) => {
      file.selected = this.checkStates[index];
    });
  }

  /**
   * 复制文件到指定文件夹
   */
  async copyFiles(files: FileInfo[], targetFolder: string): Promise<CopyResult> {
    const results: string[] = [];
    let success = 0;
    let fail = 0;

    // 确保目标文件夹存在
    await fs.ensureDir(targetFolder);

    for (const file of files) {
      try {
        const targetPath = path.join(targetFolder, file.name);
        
        // 检查目标文件是否已存在
        if (await fs.pathExists(targetPath)) {
          // 如果已存在，添加序号
          const nameWithoutExt = path.basename(file.name, path.extname(file.name));
          const ext = path.extname(file.name);
          let counter = 1;
          let newTargetPath = targetPath;
          
          while (await fs.pathExists(newTargetPath)) {
            newTargetPath = path.join(targetFolder, `${nameWithoutExt}_${counter}${ext}`);
            counter++;
          }
          
          await fs.copy(file.path, newTargetPath);
          results.push(`✓ ${file.name} -> ${path.basename(newTargetPath)} (重命名)`);
        } else {
          await fs.copy(file.path, targetPath);
          results.push(`✓ ${file.name}`);
        }
        
        success++;
      } catch (error: any) {
        results.push(`✗ ${file.name} 失败: ${error.message}`);
        fail++;
      }
    }

    return { success, fail, results };
  }

  /**
   * 复制到多个叶子文件夹
   */
  async copyToSelectedLeafFolders(
    files: FileInfo[], 
    selectedFolder: string
  ): Promise<CopyResult> {
    // 获取所有叶子文件夹
    const leafFolders = await this.getLeafFolders(selectedFolder);
    
    if (leafFolders.length === 0) {
      // 如果没有子文件夹，直接复制到当前文件夹
      return await this.copyFiles(files, selectedFolder);
    }

    const results: string[] = [];
    let success = 0;
    let fail = 0;

    // 为每个叶子文件夹复制文件
    for (const leafFolder of leafFolders) {
      const folderResult = await this.copyFiles(files, leafFolder);
      
      success += folderResult.success;
      fail += folderResult.fail;
      results.push(`📁 ${path.basename(leafFolder)}:`);
      results.push(...folderResult.results.map(r => `  ${r}`));
    }

    return { success, fail, results };
  }

  /**
   * 获取所有叶子文件夹
   */
  private async getLeafFolders(folderPath: string): Promise<string[]> {
    const leafFolders: string[] = [];
    
    const scanForLeafFolders = async (currentPath: string): Promise<void> => {
      try {
        const items = await fs.readdir(currentPath);
        let hasSubfolders = false;
        
        for (const item of items) {
          const fullPath = path.join(currentPath, item);
          const stats = await fs.stat(fullPath);
          
          if (stats.isDirectory()) {
            hasSubfolders = true;
            await scanForLeafFolders(fullPath);
          }
        }
        
        // 如果没有子文件夹，则当前文件夹是叶子文件夹
        if (!hasSubfolders) {
          leafFolders.push(currentPath);
        }
      } catch (error) {
        console.error(`扫描叶子文件夹错误: ${currentPath}`, error);
      }
    };
    
    await scanForLeafFolders(folderPath);
    return leafFolders;
  }

  /**
   * 切换文件选择状态
   */
  toggleFileSelection(index: number): void {
    if (index >= 0 && index < this.checkStates.length) {
      this.checkStates[index] = !this.checkStates[index];
      this.scannedFiles[index].selected = this.checkStates[index];
    }
  }

  /**
   * 反选
   */
  invertSelection(): void {
    this.checkStates = this.checkStates.map(state => !state);
    this.scannedFiles.forEach((file, index) => {
      file.selected = this.checkStates[index];
    });
  }
}

export default FileManager;