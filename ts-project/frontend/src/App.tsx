import React, { useState, useEffect } from 'react';
import { Layout, Menu, Spin, Alert } from 'antd';
import { FileWordOutlined, FileExcelOutlined, FolderOutlined, SettingOutlined } from '@ant-design/icons';
import WordConverter from './components/WordConverter';
import ExcelParser from './components/ExcelParser';
import FileManager from './components/FileManager';
import Settings from './components/Settings';
import apiService from './services/api';
import './App.css';

const { Header, Sider, Content } = Layout;

const menuItems = [
  {
    key: 'word',
    icon: <FileWordOutlined />,
    label: 'Word转PDF',
  },
  {
    key: 'excel',
    icon: <FileExcelOutlined />,
    label: 'Excel解析',
  },
  {
    key: 'files',
    icon: <FolderOutlined />,
    label: '文件管理',
  },
  {
    key: 'settings',
    icon: <SettingOutlined />,
    label: '设置',
  },
];

function App() {
  const [selectedKey, setSelectedKey] = useState('word');
  const [serverStatus, setServerStatus] = useState<'checking' | 'online' | 'offline'>('checking');

  useEffect(() => {
    checkServerStatus();
  }, []);

  const checkServerStatus = async () => {
    setServerStatus('checking');
    const isOnline = await apiService.healthCheck();
    setServerStatus(isOnline ? 'online' : 'offline');
  };

  const renderContent = () => {
    switch (selectedKey) {
      case 'word':
        return <WordConverter serverStatus={serverStatus} />;
      case 'excel':
        return <ExcelParser serverStatus={serverStatus} />;
      case 'files':
        return <FileManager serverStatus={serverStatus} />;
      case 'settings':
        return <Settings onServerStatusCheck={checkServerStatus} serverStatus={serverStatus} />;
      default:
        return <WordConverter serverStatus={serverStatus} />;
    }
  };

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

  const getStatusText = () => {
    switch (serverStatus) {
      case 'online':
        return '服务器在线';
      case 'offline':
        return '服务器离线';
      default:
        return '检查服务器状态...';
    }
  };

  return (
    <Layout style={{ minHeight: '100vh' }}>
      <Header style={{ 
        background: '#001529', 
        color: 'white', 
        display: 'flex', 
        justifyContent: 'space-between', 
        alignItems: 'center' 
      }}>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <h1 style={{ color: 'white', margin: 0, marginRight: '20px' }}>文档整理工具</h1>
        </div>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <div style={{ 
            display: 'flex', 
            alignItems: 'center', 
            marginRight: '20px',
            color: getStatusColor() 
          }}>
            <div style={{
              width: '8px',
              height: '8px',
              borderRadius: '50%',
              backgroundColor: getStatusColor(),
              marginRight: '8px'
            }} />
            {getStatusText()}
          </div>
        </div>
      </Header>
      
      <Layout>
        <Sider width={200} style={{ background: '#fff' }}>
          <Menu
            mode="inline"
            selectedKeys={[selectedKey]}
            items={menuItems}
            onClick={({ key }) => setSelectedKey(key)}
            style={{ height: '100%', borderRight: 0 }}
          />
        </Sider>
        
        <Layout style={{ padding: '24px' }}>
          <Content
            style={{
              background: '#fff',
              padding: '24px',
              margin: 0,
              minHeight: 280,
            }}
          >
            {serverStatus === 'checking' ? (
              <div style={{ textAlign: 'center', padding: '50px' }}>
                <Spin size="large" />
                <div style={{ marginTop: '20px' }}>正在检查服务器状态...</div>
              </div>
            ) : serverStatus === 'offline' ? (
              <Alert
                message="服务器连接失败"
                description="请确保后端服务正在运行在端口3001上，然后刷新页面重试。"
                type="error"
                showIcon
              />
            ) : (
              renderContent()
            )}
          </Content>
        </Layout>
      </Layout>
    </Layout>
  );
}

export default App;