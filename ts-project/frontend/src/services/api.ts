import axios from 'axios';

const api = axios.create({
  baseURL: '/api',
  timeout: 30000,
});

export interface ConversionResult {
  success: boolean;
  message: string;
  filePath?: string;
}

export interface FolderConversionResult {
  total: number;
  success: number;
  failed: number;
  results: Array<{
    fileName: string;
    success: boolean;
    message: string;
    outputPath?: string;
  }>;
}

export interface ExcelParseResult {
  sourceFiles: string[];
  targetFolders: Array<{
    folderName: string;
    files: string[];
  }>;
}

export const apiService = {
  // 健康检查
  async healthCheck(): Promise<boolean> {
    try {
      const response = await api.get('/health');
      return response.status === 200;
    } catch (error) {
      console.error('Health check failed:', error);
      return false;
    }
  },

  // 上传并转换单个Word文件
  async convertWordFile(file: File): Promise<ConversionResult> {
    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await api.post('/word/convert', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });
      return response.data;
    } catch (error: any) {
      return {
        success: false,
        message: error.response?.data?.message || '转换失败',
      };
    }
  },

  // 转换Word文件夹
  async convertWordFolder(folderPath: string, keepStructure: boolean = true): Promise<FolderConversionResult> {
    try {
      const response = await api.post('/word/convert-folder', {
        folderPath,
        keepStructure,
      });
      return response.data;
    } catch (error: any) {
      return {
        total: 0,
        success: 0,
        failed: 0,
        results: [],
      };
    }
  },

  // 解析Excel文件
  async parseExcelFile(file: File): Promise<ExcelParseResult> {
    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await api.post('/excel/parse', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });
      return response.data;
    } catch (error: any) {
      return {
        sourceFiles: [],
        targetFolders: [],
      };
    }
  },

  // 复制文件到目标文件夹
  async copyFiles(sourceFiles: string[], targetFolder: string): Promise<{ success: boolean; message: string }> {
    try {
      const response = await api.post('/files/copy', {
        sourceFiles,
        targetFolder,
      });
      return response.data;
    } catch (error: any) {
      return {
        success: false,
        message: error.response?.data?.message || '文件复制失败',
      };
    }
  },

  // 获取服务器文件列表
  async getFileList(path?: string): Promise<string[]> {
    try {
      const response = await api.get('/files/list', {
        params: { path },
      });
      return response.data.files || [];
    } catch (error) {
      return [];
    }
  },
};

export default apiService;