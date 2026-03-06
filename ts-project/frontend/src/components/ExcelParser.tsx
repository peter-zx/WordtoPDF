import React, { useState } from 'react';
import { Upload, Button, Card, List, message, Typography, Space, Alert, Tag } from 'antd';
import { UploadOutlined, FileExcelOutlined, FolderOutlined } from '@ant-design/icons';
import apiService, { ExcelParseResult } from '../services/api';

const { Title, Text } = Typography;

interface ExcelParserProps {
  serverStatus: 'checking' | 'online' | 'offline';
}

const ExcelParser: React.FC<ExcelParserProps> = ({ serverStatus }) => {
  const [parsing, setParsing] = useState(false);
  const [parseResult, setParseResult] = useState<ExcelParseResult | null>(null);

  const handleExcelUpload = async (file: File) => {
    try {
      setParsing(true);
      
      const result = await apiService.parseExcelFile(file);
      
      if (result.sourceFiles.length > 0) {
        message.success(`Excel解析成功，找到 ${result.sourceFiles.length} 个源文件`);
        setParseResult(result);
      } else {
        message.warning('Excel文件中未找到有效的文件路径信息');
        setParseResult(null);
      }
    } catch (error) {
      message.error('Excel文件解析失败');
    } finally {
      setParsing(false);
    }
  };

  const handleCopyFiles = async (targetFolder: string) => {
    if (!parseResult) return;
    
    try {
      const result = await apiService.copyFiles(parseResult.sourceFiles, targetFolder);
      
      if (result.success) {
        message.success('文件复制成功');
      } else {
        message.error(`文件复制失败: ${result.message}`);
      }
    } catch (error) {
      message.error('文件复制操作失败');
    }
  };

  const uploadProps = {
    beforeUpload: (file: File) => {
      if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        message.error('请上传Excel文件(.xlsx或.xls)');
        return false;
      }
      handleExcelUpload(file);
      return false;
    },
    showUploadList: false,
  };

  return (
    <div className="container">
      <div className="page-header">
        <Title level={2}>Excel文件解析工具</Title>
        <Text type="secondary">解析Excel中的文件路径信息，自动整理文件结构</Text>
      </div>

      {serverStatus === 'offline' && (
        <Alert
          message="服务器未连接"
          description="请确保后端服务正在运行，否则无法进行Excel解析。"
          type="warning"
          showIcon
          style={{ marginBottom: 20 }}
        />
      )}

      <Space direction="vertical" size="large" style={{ width: '100%' }}>
        <Card title="上传Excel文件" style={{ width: '100%' }}>
          <Upload {...uploadProps}>
            <Button 
              icon={<UploadOutlined />} 
              size="large" 
              disabled={serverStatus !== 'online'}
              loading={parsing}
            >
              上传Excel文件
            </Button>
          </Upload>
          <Text type="secondary" style={{ display: 'block', marginTop: 8 }}>
            支持 .xlsx 和 .xls 格式，Excel中应包含源文件路径和目标文件夹信息
          </Text>
        </Card>

        {parseResult && (
          <>
            <Card title="解析结果" style={{ width: '100%' }}>
              <Space direction="vertical" style={{ width: '100%' }}>
                <div>
                  <Text strong>找到的源文件: </Text>
                  <Tag color="blue">{parseResult.sourceFiles.length} 个文件</Tag>
                </div>
                
                <List
                  size="small"
                  bordered
                  dataSource={parseResult.sourceFiles}
                  renderItem={(filePath) => (
                    <List.Item>
                      <FileExcelOutlined style={{ color: '#52c41a', marginRight: 8 }} />
                      <Text code>{filePath}</Text>
                    </List.Item>
                  )}
                  style={{ maxHeight: 200, overflowY: 'auto' }}
                />

                {parseResult.targetFolders.length > 0 && (
                  <div>
                    <Text strong>目标文件夹结构: </Text>
                    <Tag color="green">{parseResult.targetFolders.length} 个文件夹</Tag>
                  </div>
                )}

                {parseResult.targetFolders.map((folder, index) => (
                  <Card key={index} size="small" title={folder.folderName}>
                    <List
                      size="small"
                      dataSource={folder.files}
                      renderItem={(file) => (
                        <List.Item>
                          <Text>{file}</Text>
                        </List.Item>
                      )}
                    />
                  </Card>
                ))}
              </Space>
            </Card>

            <Card title="文件操作" style={{ width: '100%' }}>
              <Space>
                <Button 
                  icon={<FolderOutlined />}
                  onClick={() => {
                    const targetFolder = prompt('请输入目标文件夹路径:');
                    if (targetFolder) {
                      handleCopyFiles(targetFolder);
                    }
                  }}
                  disabled={serverStatus !== 'online'}
                >
                  复制文件到目标文件夹
                </Button>
                
                <Button 
                  onClick={() => {
                    // 导出文件列表功能
                    const fileList = parseResult.sourceFiles.join('\n');
                    const blob = new Blob([fileList], { type: 'text/plain' });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'file-list.txt';
                    a.click();
                    URL.revokeObjectURL(url);
                    message.success('文件列表导出成功');
                  }}
                >
                  导出文件列表
                </Button>
              </Space>
            </Card>
          </>
        )}
      </Space>
    </div>
  );
};

export default ExcelParser;