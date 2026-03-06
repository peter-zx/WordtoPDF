import React, { useState, useEffect } from 'react';
import { Card, Button, List, Input, Space, Typography, message, Alert, Tag } from 'antd';
import { FolderOpenOutlined, UploadOutlined, DownloadOutlined, SearchOutlined } from '@ant-design/icons';
import apiService from '../services/api';

const { Title, Text } = Typography;
const { Search } = Input;

interface FileManagerProps {
  serverStatus: 'checking' | 'online' | 'offline';
}

const FileManager: React.FC<FileManagerProps> = ({ serverStatus }) => {
  const [files, setFiles] = useState<string[]>([]);
  const [currentPath, setCurrentPath] = useState<string>('');
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (serverStatus === 'online') {
      loadFiles();
    }
  }, [serverStatus]);

  const loadFiles = async (path?: string) => {
    try {
      setLoading(true);
      const fileList = await apiService.getFileList(path);
      setFiles(fileList);
      setCurrentPath(path || '');
    } catch (error) {
      message.error('获取文件列表失败');
    } finally {
      setLoading(false);
    }
  };

  const handleSearch = (value: string) => {
    if (value.trim()) {
      loadFiles(value.trim());
    } else {
      loadFiles();
    }
  };

  const handleNavigate = (path: string) => {
    if (path.endsWith('/')) {
      loadFiles(path);
    }
  };

  const handleUpload = () => {
    message.info('文件上传功能开发中...');
  };

  const handleDownload = (fileName: string) => {
    message.info(`下载文件: ${fileName}`);
    // 实际实现需要后端提供下载接口
  };

  return (
    <div className="container">
      <div className="page-header">
        <Title level={2}>文件管理系统</Title>
        <Text type="secondary">浏览和管理服务器上的文件</Text>
      </div>

      {serverStatus === 'offline' && (
        <Alert
          message="服务器未连接"
          description="请确保后端服务正在运行，否则无法访问文件系统。"
          type="warning"
          showIcon
          style={{ marginBottom: 20 }}
        />
      )}

      <Space direction="vertical" size="large" style={{ width: '100%' }}>
        <Card title="文件浏览器" style={{ width: '100%' }}>
          <Space direction="vertical" style={{ width: '100%' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Search
                placeholder="输入文件夹路径或文件名进行搜索"
                enterButton={<SearchOutlined />}
                size="large"
                onSearch={handleSearch}
                style={{ width: '70%' }}
                disabled={serverStatus !== 'online'}
              />
              
              <Space>
                <Button 
                  icon={<UploadOutlined />} 
                  onClick={handleUpload}
                  disabled={serverStatus !== 'online'}
                >
                  上传文件
                </Button>
                <Button 
                  icon={<FolderOpenOutlined />} 
                  onClick={() => loadFiles()}
                  disabled={serverStatus !== 'online'}
                >
                  刷新
                </Button>
              </Space>
            </div>

            {currentPath && (
              <div>
                <Text type="secondary">当前路径: </Text>
                <Tag color="blue">{currentPath || '根目录'}</Tag>
              </div>
            )}

            <List
              loading={loading}
              bordered
              dataSource={files}
              renderItem={(file) => {
                const isDirectory = file.endsWith('/');
                const displayName = isDirectory ? file.slice(0, -1) : file;
                
                return (
                  <List.Item
                    actions={[
                      <Button 
                        key="download" 
                        type="link" 
                        icon={<DownloadOutlined />}
                        onClick={() => handleDownload(file)}
                        disabled={isDirectory || serverStatus !== 'online'}
                      >
                        下载
                      </Button>
                    ]}
                  >
                    <List.Item.Meta
                      avatar={<FolderOpenOutlined style={{ color: isDirectory ? '#1890ff' : '#52c41a' }} />}
                      title={
                        <Button 
                          type="link" 
                          onClick={() => isDirectory && handleNavigate(file)}
                          style={{ padding: 0, height: 'auto' }}
                        >
                          {displayName}
                        </Button>
                      }
                      description={
                        <Tag color={isDirectory ? 'blue' : 'green'}>
                          {isDirectory ? '文件夹' : '文件'}
                        </Tag>
                      }
                    />
                  </List.Item>
                );
              }}
              style={{ maxHeight: 500, overflowY: 'auto' }}
              locale={{ emptyText: '暂无文件' }}
            />

            <div style={{ textAlign: 'center', marginTop: 16 }}>
              <Text type="secondary">
                共找到 {files.length} 个文件/文件夹
              </Text>
            </div>
          </Space>
        </Card>

        <Card title="快速操作" style={{ width: '100%' }}>
          <Space wrap>
            <Button 
              onClick={() => loadFiles('./output')}
              disabled={serverStatus !== 'online'}
            >
              查看输出文件夹
            </Button>
            <Button 
              onClick={() => loadFiles('./temp')}
              disabled={serverStatus !== 'online'}
            >
              查看临时文件夹
            </Button>
            <Button 
              onClick={() => loadFiles('./logs')}
              disabled={serverStatus !== 'online'}
            >
              查看日志文件夹
            </Button>
          </Space>
        </Card>
      </Space>
    </div>
  );
};

export default FileManager;