import React, { useState } from 'react';
import { Upload, Button, Progress, List, message, Card, Switch, Space, Typography, Alert } from 'antd';
import { UploadOutlined, FolderOpenOutlined, CheckCircleOutlined, CloseCircleOutlined } from '@ant-design/icons';
import apiService, { FolderConversionResult } from '../services/api';

const { Title, Text } = Typography;

interface WordConverterProps {
  serverStatus: 'checking' | 'online' | 'offline';
}

const WordConverter: React.FC<WordConverterProps> = ({ serverStatus }) => {
  const [converting, setConverting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [results, setResults] = useState<FolderConversionResult | null>(null);
  const [keepStructure, setKeepStructure] = useState(true);

  const handleFileUpload = async (file: File) => {
    try {
      setConverting(true);
      setProgress(0);
      
      const result = await apiService.convertWordFile(file);
      
      if (result.success) {
        message.success(`文件转换成功: ${file.name}`);
        setResults({
          total: 1,
          success: 1,
          failed: 0,
          results: [{
            fileName: file.name,
            success: true,
            message: '转换成功',
            outputPath: result.filePath
          }]
        });
      } else {
        message.error(`转换失败: ${result.message}`);
        setResults({
          total: 1,
          success: 0,
          failed: 1,
          results: [{
            fileName: file.name,
            success: false,
            message: result.message
          }]
        });
      }
    } catch (error) {
      message.error('文件上传失败');
    } finally {
      setConverting(false);
      setProgress(100);
    }
  };

  const handleFolderConversion = async () => {
    const folderPath = prompt('请输入要转换的文件夹路径:');
    if (!folderPath) return;
    
    try {
      setConverting(true);
      setProgress(0);
      setResults(null);
      
      // 模拟进度更新
      const progressInterval = setInterval(() => {
        setProgress(prev => {
          if (prev >= 90) {
            clearInterval(progressInterval);
            return prev;
          }
          return prev + 10;
        });
      }, 1000);
      
      const result = await apiService.convertWordFolder(folderPath, keepStructure);
      
      clearInterval(progressInterval);
      setProgress(100);
      
      if (result.total > 0) {
        message.success(`转换完成: 成功 ${result.success} 个，失败 ${result.failed} 个`);
        setResults(result);
      } else {
        message.warning('未找到可转换的文件');
      }
    } catch (error) {
      message.error('文件夹转换失败');
    } finally {
      setConverting(false);
    }
  };

  const uploadProps = {
    beforeUpload: (file: File) => {
      if (!file.name.endsWith('.doc') && !file.name.endsWith('.docx')) {
        message.error('请上传Word文档(.doc或.docx)');
        return false;
      }
      handleFileUpload(file);
      return false;
    },
    showUploadList: false,
  };

  return (
    <div className="container">
      <div className="page-header">
        <Title level={2}>Word转PDF工具</Title>
        <Text type="secondary">支持单个文件上传和批量文件夹转换</Text>
      </div>

      {serverStatus === 'offline' && (
        <Alert
          message="服务器未连接"
          description="请确保后端服务正在运行，否则无法进行文件转换。"
          type="warning"
          showIcon
          style={{ marginBottom: 20 }}
        />
      )}

      <Space direction="vertical" size="large" style={{ width: '100%' }}>
        <Card title="单个文件转换" style={{ width: '100%' }}>
          <Upload {...uploadProps}>
            <Button icon={<UploadOutlined />} size="large" disabled={serverStatus !== 'online'}>
              上传Word文件
            </Button>
          </Upload>
          <Text type="secondary" style={{ display: 'block', marginTop: 8 }}>
            支持 .doc 和 .docx 格式
          </Text>
        </Card>

        <Card title="批量文件夹转换" style={{ width: '100%' }}>
          <Space direction="vertical" style={{ width: '100%' }}>
            <div>
              <Text>保留文件夹结构: </Text>
              <Switch 
                checked={keepStructure} 
                onChange={setKeepStructure}
                disabled={serverStatus !== 'online'}
              />
              <Text type="secondary" style={{ marginLeft: 8 }}>
                {keepStructure ? '保持原文件夹结构' : '所有文件输出到同一目录'}
              </Text>
            </div>
            
            <Button 
              icon={<FolderOpenOutlined />} 
              size="large" 
              onClick={handleFolderConversion}
              disabled={serverStatus !== 'online' || converting}
              loading={converting}
            >
              选择文件夹进行批量转换
            </Button>
          </Space>
        </Card>

        {converting && (
          <Card title="转换进度" style={{ width: '100%' }}>
            <Progress percent={progress} status="active" />
            <Text type="secondary" style={{ display: 'block', marginTop: 8 }}>
              正在处理文件，请稍候...
            </Text>
          </Card>
        )}

        {results && (
          <Card title="转换结果" style={{ width: '100%' }}>
            <div style={{ marginBottom: 16 }}>
              <Text strong>总计: {results.total} 个文件</Text>
              <Text type="success" style={{ marginLeft: 16 }}>
                成功: {results.success}
              </Text>
              <Text type="danger" style={{ marginLeft: 16 }}>
                失败: {results.failed}
              </Text>
            </div>
            
            <List
              size="small"
              bordered
              dataSource={results.results}
              renderItem={(item) => (
                <List.Item>
                  <Space>
                    {item.success ? (
                      <CheckCircleOutlined style={{ color: '#52c41a' }} />
                    ) : (
                      <CloseCircleOutlined style={{ color: '#ff4d4f' }} />
                    )}
                    <Text style={{ color: item.success ? '#52c41a' : '#ff4d4f' }}>
                      {item.fileName}
                    </Text>
                    <Text type="secondary">{item.message}</Text>
                  </Space>
                </List.Item>
              )}
              style={{ maxHeight: 300, overflowY: 'auto' }}
            />
          </Card>
        )}
      </Space>
    </div>
  );
};

export default WordConverter;