import ExcelJS from 'exceljs';
import fs from 'fs-extra';
import path from 'path';
import { ExcelParseResult, FolderStructure } from '../../shared/types.js';

export class ExcelParser {
  /**
   * 从Excel文件解析文件夹结构
   */
  async parseFromFile(filePath: string): Promise<ExcelParseResult> {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      
      // 默认使用第一个工作表
      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        throw new Error('Excel文件中没有工作表');
      }
      
      return this.parseWorksheet(worksheet);
    } catch (error: any) {
      throw new Error(`Excel解析失败: ${error.message}`);
    }
  }

  /**
   * 解析工作表数据
   */
  private parseWorksheet(worksheet: ExcelJS.Worksheet): ExcelParseResult {
    const structure: FolderStructure[] = [];
    let totalFolders = 0;
    
    // 假设第一列是文件夹名称，支持层级结构
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 跳过标题行
      
      const folderName = row.getCell(1).text?.trim();
      if (!folderName) return;
      
      // 解析层级（通过缩进或斜杠）
      const level = this.detectFolderLevel(folderName);
      const cleanName = this.cleanFolderName(folderName);
      
      if (cleanName) {
        this.addToStructure(structure, cleanName, level);
        totalFolders++;
      }
    });
    
    return { structure, totalFolders };
  }

  /**
   * 检测文件夹层级
   */
  private detectFolderLevel(folderName: string): number {
    // 方法1: 通过缩进空格判断层级
    const leadingSpaces = folderName.match(/^\s*/)?.[0].length || 0;
    if (leadingSpaces > 0) {
      return Math.floor(leadingSpaces / 2); // 假设每级缩进2个空格
    }
    
    // 方法2: 通过斜杠判断层级
    const slashCount = (folderName.match(/\//g) || []).length;
    if (slashCount > 0) {
      return slashCount;
    }
    
    // 方法3: 通过连字符判断层级
    const dashCount = (folderName.match(/-/g) || []).length;
    if (dashCount > 0) {
      return dashCount;
    }
    
    return 0; // 默认顶级
  }

  /**
   * 清理文件夹名称
   */
  private cleanFolderName(folderName: string): string {
    // 移除前导空格
    let clean = folderName.trim();
    
    // 移除层级标记（如果有）
    clean = clean.replace(/^[\s\-\/]+/, '');
    
    // 移除Windows非法字符
    const invalidChars = /[<>:"|?*]/g;
    clean = clean.replace(invalidChars, '_');
    
    // 限制长度
    if (clean.length > 255) {
      clean = clean.substring(0, 255);
    }
    
    return clean;
  }

  /**
   * 将文件夹添加到结构树中
   */
  private addToStructure(
    structure: FolderStructure[], 
    folderName: string, 
    level: number
  ): void {
    if (level === 0) {
      // 顶级文件夹
      structure.push({ name: folderName });
    } else {
      // 子文件夹，需要找到父级
      let currentLevel = structure;
      
      for (let i = 0; i < level; i++) {
        if (currentLevel.length === 0) {
          // 如果当前层级为空，创建占位文件夹
          const placeholder: FolderStructure = { name: `未命名${i + 1}` };
          currentLevel.push(placeholder);
        }
        
        const lastFolder = currentLevel[currentLevel.length - 1];
        if (!lastFolder.children) {
          lastFolder.children = [];
        }
        
        currentLevel = lastFolder.children;
      }
      
      currentLevel.push({ name: folderName });
    }
  }

  /**
   * 创建文件夹结构
   */
  async createStructure(basePath: string, structure: FolderStructure[]): Promise<{ created: number; errors: string[] }> {
    let created = 0;
    const errors: string[] = [];
    
    const createFolderRecursive = async (
      parentPath: string, 
      folders: FolderStructure[]
    ): Promise<void> => {
      for (const folder of folders) {
        const folderPath = path.join(parentPath, folder.name);
        
        try {
          await fs.ensureDir(folderPath);
          created++;
          
          // 递归创建子文件夹
          if (folder.children && folder.children.length > 0) {
            await createFolderRecursive(folderPath, folder.children);
          }
        } catch (error: any) {
          errors.push(`创建文件夹失败: ${folderPath} - ${error.message}`);
        }
      }
    };
    
    await createFolderRecursive(basePath, structure);
    return { created, errors };
  }

  /**
   * 验证Excel文件格式
   */
  async validateExcel(filePath: string): Promise<{ valid: boolean; message: string }> {
    try {
      if (!await fs.pathExists(filePath)) {
        return { valid: false, message: '文件不存在' };
      }
      
      const stats = await fs.stat(filePath);
      if (stats.size === 0) {
        return { valid: false, message: '文件为空' };
      }
      
      // 检查文件扩展名
      const ext = path.extname(filePath).toLowerCase();
      if (!['.xlsx', '.xls'].includes(ext)) {
        return { valid: false, message: '不支持的文件格式，请使用.xlsx或.xls文件' };
      }
      
      // 尝试读取文件
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      
      if (workbook.worksheets.length === 0) {
        return { valid: false, message: 'Excel文件中没有工作表' };
      }
      
      return { valid: true, message: '文件格式正确' };
    } catch (error: any) {
      return { valid: false, message: `文件验证失败: ${error.message}` };
    }
  }
}

export default ExcelParser;