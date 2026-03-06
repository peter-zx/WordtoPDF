import React, { useState } from 'react';
import { Card, Button, Space, Typography, Switch, Divider, Alert, List, Tag } from 'antd';
import { ReloadOutlined, InfoCircleOutlined } from '@ant-design/icons';

const { Title, Text } = Typography;

interface SettingsProps {
  serverStatus: 'checking' | 'online' | 'offline';
  onServerStatusCheck: () => void;
}

const Settings: React.FC<SettingsProps> = ({ serverStatus, onServerStatusCheck }) => {
  const [keepStructure, setKeepStructure] = useState(true);
  const [autoDeleteTemp, setAutoDeleteTemp] = useState(true);
  const [concurrencyLimit, setConcurrencyLimit] = useState(2);

  const systemInfo = [
    { label: '前端版本', value: '1.0.0' },
    { label: '后端API版本', value: '1.0.0' },
    { label: '服务器状态', value: serverStatus === 'online' ? '在线' : serverStatus === 'offline' ? '离线' : '检查中' },
    { label: '并发限制', value: `${concurrencyLimit} 个进程` },
    { label: '临时文件清理', value: autoDeleteTemp ? '启用' : '禁用' },
  ];

  const getStatusColor = () => {
    switch (serverStatus) {
      case 'online':
        return '#52c41a';
      case 'offline':
        return '#ff4d4f';
      default:
        return '#faad14';
    }
  };

  return (
    <div className="container">
      <div className="page-header">
        <Title level={2}>系统设置</Title>
        <Text type="secondary">管理系统配置和查看系统信息</Text>
      </div>

      <Space direction="vertical" size="large" style={{ width: '100%' }}>
        <Card title="服务器状态" style={{ width: '100%' }}>
          <Space direction="vertical" style={{ width: '100%' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center' }}>
                <div style={{
                  width: '12px',
                  height: '12px',
                  borderRadius: '50%',
                  backgroundColor: getStatusColor(),
                  marginRight: '12px'
                }} />
                <Text strong style={{ fontSize: '16px', color: getStatusColor() }}>
                  {serverStatus === 'online' ? '服务器在线' : 
                   serverStatus === 'offline' ? '服务器离线' : '检查服务器状态...'}
                </Text>
              </div>
              
              <Button 
                icon={<ReloadOutlined />} 
                onClick={onServerStatusCheck}
                loading={serverStatus === 'checking'}
              >
                刷新状态
              </Button>
            </div>

            {serverStatus === 'offline' && (
              <Alert
                message="连接问题解决方案"
                description={
                  <div>
                    <p>1. 确保后端服务正在端口3001上运行</p>
                    <p>2. 检查防火墙设置是否阻止了连接</p>
                    <p>3. 确认网络连接正常</p>
                    <p>4. 重启后端服务后刷新页面</p>
                  </div>
                }
                type="error"
                showIcon
              />
            )}
          </Space>
        </Card>

        <Card title="转换设置" style={{ width: '100%' }}>
          <Space direction="vertical" style={{ width: '100%' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <Text strong>保留文件夹结构</Text>
                <br />
                <Text type="secondary">转换时保持原文件夹层级关系</Text>
              </div>
              <Switch 
                checked={keepStructure} 
                onChange={setKeepStructure}
              />
            </div>

            <Divider style={{ margin: '16px 0' }} />

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <Text strong>自动清理临时文件</Text>
                <br />
                <Text type="secondary">转换完成后自动删除临时文件</Text>
              </div>
              <Switch 
                checked={autoDeleteTemp} 
                onChange={setAutoDeleteTemp}
              />
            </div>

            <Divider style={{ margin: '16px 0' }} />

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <Text strong>并发处理限制</Text>
                <br />
                <Text type="secondary">同时处理的文件数量 (1-4)</Text>
              </div>
              <Space>
                <Button 
                  size="small" 
                  onClick={() => setConcurrencyLimit(Math.max(1, concurrencyLimit - 1))}
                  disabled={concurrencyLimit <= 1}
                >
                  -
                </Button>
                <Tag>{concurrencyLimit}</Tag>
                <Button 
                  size="small" 
                  onClick={() => setConcurrencyLimit(Math.min(4, concurrencyLimit + 1))}
                  disabled={concurrencyLimit >= 4}
                >
                  +
                </Button>
              </Space>
            </div>
          </Space>
        </Card>

        <Card title="系统信息" style={{ width: '100%' }}>
          <List
            dataSource={systemInfo}
            renderItem={(item) => (
              <List.Item>
                <List.Item.Meta
                  avatar={<InfoCircleOutlined />}
                  title={item.label}
                  description={
                    <Tag 
                      color={item.label === '服务器状态' ? 
                        (serverStatus === 'online' ? 'green' : serverStatus === 'offline' ? 'red' : 'orange') : 'blue'
                      }
                    >
                      {item.value}
                    </Tag>
                  }
                />
              </List.Item>
            )}
          />
        </Card>

        <Card title="使用说明" style={{ width: '100%' }}>
          <Space direction="vertical" style={{ width: '100%' }}>
            <Text strong>Word转PDF功能:</Text>
            <Text type="secondary">• 支持单个文件上传和批量文件夹转换</Text>
            <Text type="secondary">• 转换完成后文件保存在output文件夹</Text>
            <Text type="secondary">• 支持保留原文件夹结构</Text>
            
            <Text strong style={{ marginTop: 16 }}>Excel解析功能:</Text>
            <Text type="secondary">• 解析Excel中的文件路径信息</Text>
            <Text type="secondary">• 自动整理文件到指定文件夹</Text>
            <Text type="secondary">• 支持文件列表导出</Text>
            
            <Text strong style={{ marginTop: 16 }}>文件管理:</Text>
            <Text type="secondary">• 浏览服务器文件系统</Text>
            <Text type="secondary">• 支持文件搜索和导航</Text>
            <Text type="secondary">• 文件上传下载功能</Text>
          </Space>
        </Card>
      </Space>
    </div>
  );
};

export default Settings;