// 共享类型定义

export interface FileInfo {
  name: string;
  path: string;
  size: number;
  sizeStr: string;
  type: string;
  selected: boolean;
}

export interface FolderStructure {
  name: string;
  children?: FolderStructure[];
}

export interface ConversionResult {
  success: boolean;
  message: string;
  inputPath: string;
  outputPath: string;
}

export interface ExcelParseResult {
  structure: FolderStructure[];
  totalFolders: number;
}

export interface CopyResult {
  success: number;
  fail: number;
  results: string[];
}

export interface ProgressInfo {
  current: number;
  total: number;
  filename: string;
  success: boolean;
  error?: string;
}

// API响应类型
export interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  message?: string;
  error?: string;
}

// 配置类型
export interface AppConfig {
  maxConcurrentConversions: number;
  supportedFileTypes: string[];
  defaultOutputPath: string;
}