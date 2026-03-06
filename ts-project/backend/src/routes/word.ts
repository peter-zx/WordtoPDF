import express from 'express';
import WordConverter from '../services/WordConverter.js';

const router = express.Router();
const wordConverter = new WordConverter();

/**
 * 检查Word转换器状态
 */
router.get('/status', async (req, res) => {
  try {
    const status = await wordConverter.checkAvailability();
    res.json({ success: true, data: status });
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * 转换单个Word文件
 */
router.post('/convert', async (req, res) => {
  try {
    const { inputPath, outputPath } = req.body;
    
    if (!inputPath || !outputPath) {
      return res.status(400).json({ 
        success: false, 
        error: '缺少输入路径或输出路径参数' 
      });
    }
    
    const result = await wordConverter.convertFile(inputPath, outputPath);
    res.json({ success: true, data: result });
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * 批量转换文件夹中的Word文件
 */
router.post('/convert-folder', async (req, res) => {
  try {
    const { sourceFolder, outputFolder, keepStructure = true } = req.body;
    
    if (!sourceFolder || !outputFolder) {
      return res.status(400).json({ 
        success: false, 
        error: '缺少源文件夹或输出文件夹参数' 
      });
    }
    
    // 设置进度回调（如果前端需要实时进度）
    const progressCallback = (progress: any) => {
      // 这里可以实现WebSocket或Server-Sent Events来推送进度
      console.log(`进度: ${progress.current}/${progress.total} - ${progress.filename}`);
    };
    
    const result = await wordConverter.convertFolder(
      sourceFolder, 
      outputFolder, 
      keepStructure,
      progressCallback
    );
    
    res.json({ success: true, data: result });
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * 上传并转换Word文件
 */
import multer from 'multer';
import path from 'path';
import fs from 'fs-extra';

// 配置multer
const storage = multer.diskStorage({
  destination: async (req, file, cb) => {
    const uploadDir = path.join(process.cwd(), 'uploads');
    await fs.ensureDir(uploadDir);
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueName = Date.now() + '-' + Math.round(Math.random() * 1E9) + path.extname(file.originalname);
    cb(null, uniqueName);
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['.doc', '.docx'];
    const ext = path.extname(file.originalname).toLowerCase();
    
    if (allowedTypes.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('只支持.doc和.docx文件'));
    }
  },
  limits: {
    fileSize: 50 * 1024 * 1024 // 50MB
  }
});

router.post('/upload-convert', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: '没有上传文件' });
    }
    
    const inputPath = req.file.path;
    const outputDir = path.join(process.cwd(), 'output');
    await fs.ensureDir(outputDir);
    
    const outputFileName = path.basename(req.file.originalname, path.extname(req.file.originalname)) + '.pdf';
    const outputPath = path.join(outputDir, outputFileName);
    
    const result = await wordConverter.convertFile(inputPath, outputPath);
    
    // 清理上传的文件
    await fs.remove(inputPath);
    
    if (result.success) {
      // 返回下载链接
      res.json({
        success: true,
        data: {
          ...result,
          downloadUrl: `/output/${outputFileName}`
        }
      });
    } else {
      res.json({ success: false, error: result.message });
    }
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

export default router;