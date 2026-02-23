"""
仪表盘+卡片+分栏 组合布局样式
基于平面设计四大原则: 对齐、对比、重复、亲密性
"""

STYLES_DASHBOARD = """
<style>
/* ========================================
   全局设置
   ======================================== */
:root {
    /* 间距系统 (8px 基准) */
    --spacing-xs: 8px;
    --spacing-sm: 12px;
    --spacing-md: 16px;
    --spacing-lg: 24px;
    --spacing-xl: 32px;
    --spacing-xxl: 48px;

    /* 颜色系统 */
    --color-primary: #4299e1;
    --color-primary-dark: #3182ce;
    --color-primary-light: #63b3ed;
    --color-secondary: #718096;
    --color-success: #48bb78;
    --color-warning: #ed8936;
    --color-error: #f56565;
    --color-bg: #f7fafc;
    --color-card: #ffffff;
    --color-border: #e2e8f0;
    --color-text: #2d3748;
    --color-text-light: #718096;

    /* 圆角系统 */
    --radius-sm: 6px;
    --radius-md: 8px;
    --radius-lg: 12px;
    --radius-xl: 16px;

    /* 阴影系统 */
    --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
    --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    --shadow-xl: 0 20px 25px -5px rgba(0, 0, 0, 0.1);
}

.stApp {
    background: var(--color-bg) !important;
}

/* ========================================
   顶部导航栏
   ======================================== */
.top-nav {
    background: linear-gradient(135deg, var(--color-primary) 0%, var(--color-primary-dark) 100%);
    padding: var(--spacing-md) var(--spacing-xl);
    box-shadow: var(--shadow-lg);
    margin-bottom: var(--spacing-lg);
    border-radius: 0 0 var(--radius-lg) var(--radius-lg);
}

.top-nav h1 {
    color: white !important;
    font-size: 1.8rem !important;
    font-weight: 700 !important;
    margin: 0 !important;
}

.top-nav-subtitle {
    color: rgba(255, 255, 255, 0.9);
    font-size: 0.9rem;
    margin-top: var(--spacing-xs);
}

/* ========================================
   侧边导航栏
   ======================================== */
.sidebar-nav {
    padding: 0;
    min-height: calc(100vh - 120px);
}

.nav-item {
    display: flex;
    align-items: center;
    padding: var(--spacing-md);
    margin-bottom: var(--spacing-sm);
    border-radius: var(--radius-md);
    cursor: pointer;
    transition: all 0.3s ease;
    border: 2px solid transparent;
}

.nav-item:hover {
    background: var(--color-bg);
    border-color: var(--color-primary-light);
    transform: translateX(4px);
}

.nav-item.active {
    background: linear-gradient(135deg, var(--color-primary-light) 0%, var(--color-primary) 100%);
    color: white;
    border-color: var(--color-primary-dark);
    box-shadow: var(--shadow-md);
}

.nav-item-icon {
    font-size: 1.5rem;
    margin-right: var(--spacing-md);
}

.nav-item-text {
    font-weight: 600;
    font-size: 0.95rem;
}

.nav-item-description {
    font-size: 0.8rem;
    color: #718096;
    padding-left: 48px;
    margin-bottom: 16px;
    margin-top: 4px;
}

/* ========================================
   主内容区 - 卡片容器
   ======================================== */
.main-content {
    padding: var(--spacing-lg);
}

/* 功能卡片 */
.function-card {
    background: var(--color-card);
    border: 2px solid var(--color-border);
    border-radius: var(--radius-lg);
    padding: var(--spacing-lg);
    margin-bottom: var(--spacing-lg);
    box-shadow: var(--shadow-md);
    transition: all 0.3s ease;
}

.function-card:hover {
    box-shadow: var(--shadow-lg);
    border-color: var(--color-primary-light);
}

.function-card-header {
    display: flex;
    align-items: center;
    margin-bottom: var(--spacing-lg);
    padding-bottom: var(--spacing-md);
    border-bottom: 2px solid var(--color-bg);
}

.function-card-icon {
    font-size: 2rem;
    margin-right: var(--spacing-md);
}

.function-card-title {
    font-size: 1.3rem;
    font-weight: 700;
    color: var(--color-text);
    margin: 0;
}

.function-card-description {
    font-size: 0.85rem;
    color: var(--color-text-light);
    margin-top: var(--spacing-xs);
}

/* ========================================
   步骤卡片 - Google Cloud风格
   ======================================== */
.step-container {
    display: flex;
    flex-direction: column;
    gap: var(--spacing-lg);
    margin-bottom: var(--spacing-lg);
}

.step-card {
    background: white;
    border: 1px solid var(--color-border);
    border-radius: var(--radius-md);
    overflow: hidden;
    transition: all 0.3s ease;
}

.step-card:hover {
    border-color: var(--color-primary-light);
    box-shadow: var(--shadow-md);
}

.step-card-header {
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    padding: var(--spacing-md) var(--spacing-lg);
    border-bottom: 1px solid var(--color-border);
    display: flex;
    align-items: center;
}

.step-number {
    background: linear-gradient(135deg, var(--color-primary) 0%, var(--color-primary-dark) 100%);
    color: white;
    width: 28px;
    height: 28px;
    border-radius: 6px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 0.85rem;
    margin-right: var(--spacing-sm);
    box-shadow: var(--shadow-sm);
}

.step-title {
    font-size: 0.95rem;
    font-weight: 600;
    color: var(--color-text);
}

.step-content {
    padding: var(--spacing-lg);
    background: white;
}

.step-hints {
    background: #f8f9fa;
    border-left: 3px solid var(--color-primary);
    padding: var(--spacing-md);
    margin-bottom: var(--spacing-lg);
    border-radius: var(--radius-sm);
}

.step-hints p {
    margin: 0;
    padding: var(--spacing-xs) 0;
    color: var(--color-text);
    font-size: 0.9rem;
}

/* ========================================
   按钮样式 - Google Cloud风格
   ======================================== */
.action-button {
    background: linear-gradient(135deg, var(--color-primary) 0%, var(--color-primary-dark) 100%);
    color: white;
    border: none;
    border-radius: var(--radius-sm);
    padding: var(--spacing-md) var(--spacing-xl);
    font-size: 0.9rem;
    font-weight: 500;
    cursor: pointer;
    transition: all 0.2s ease;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    width: 100%;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: var(--spacing-sm);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}

.action-button:hover {
    background: linear-gradient(135deg, var(--color-primary-light) 0%, var(--color-primary) 100%);
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
    transform: translateY(-1px);
}

.action-button:active {
    transform: translateY(0);
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
}

.action-button:disabled {
    background: #e2e8f0;
    color: #a0aec0;
    cursor: not-allowed;
    transform: none;
    box-shadow: none;
}

/* Streamlit原生按钮样式 */
.stButton > button {
    background: linear-gradient(135deg, var(--color-primary) 0%, var(--color-primary-dark) 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: var(--radius-sm) !important;
    padding: var(--spacing-md) var(--spacing-xl) !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    cursor: pointer !important;
    transition: all 0.2s ease !important;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1) !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
}

.stButton > button:hover {
    background: linear-gradient(135deg, var(--color-primary-light) 0%, var(--color-primary) 100%) !important;
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15) !important;
    transform: translateY(-1px) !important;
}

.stButton > button:disabled {
    background: #e2e8f0 !important;
    color: #a0aec0 !important;
    cursor: not-allowed !important;
    transform: none !important;
    box-shadow: none !important;
}

/* ========================================
   输入框样式 - Google Cloud风格
   ======================================== */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stSelectbox > div > div > div {
    border: 1px solid var(--color-border) !important;
    border-radius: var(--radius-md) !important;
    padding: var(--spacing-sm) var(--spacing-md) !important;
    font-size: 0.9rem !important;
    transition: all 0.2s ease !important;
    background: #f8f9fa !important;
    font-family: 'Monaco', 'Menlo', 'Ubuntu Mono', monospace !important;
}

.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: var(--color-primary) !important;
    background: white !important;
    box-shadow: 0 0 0 3px rgba(66, 153, 225, 0.1) !important;
    outline: none !important;
}

.stTextArea > div > div > textarea {
    line-height: 1.6 !important;
    min-height: 120px !important;
}

/* ========================================
   文件上传器样式
   ======================================== */
.stFileUploader {
    border: 2px dashed var(--color-border) !important;
    border-radius: var(--radius-md) !important;
    padding: var(--spacing-xl) !important;
    background: var(--color-bg) !important;
    transition: all 0.3s ease !important;
}

.stFileUploader:hover {
    border-color: var(--color-primary) !important;
    background: white !important;
}

/* ========================================
   信息框样式
   ======================================== */
.info-box {
    background: #ebf8ff;
    border-left: 4px solid var(--color-primary);
    border-radius: var(--radius-md);
    padding: var(--spacing-md);
    margin: var(--spacing-md) 0;
    font-size: 0.85rem;
    color: #2c5282;
    display: flex;
    align-items: center;
    gap: var(--spacing-sm);
}

.success-box {
    background: #f0fff4;
    border-left: 4px solid var(--color-success);
    border-radius: var(--radius-md);
    padding: var(--spacing-md);
    margin: var(--spacing-md) 0;
    font-size: 0.85rem;
    color: #276749;
    display: flex;
    align-items: center;
    gap: var(--spacing-sm);
}

.warning-box {
    background: #fffaf0;
    border-left: 4px solid var(--color-warning);
    border-radius: var(--radius-md);
    padding: var(--spacing-md);
    margin: var(--spacing-md) 0;
    font-size: 0.85rem;
    color: #9c4221;
    display: flex;
    align-items: center;
    gap: var(--spacing-sm);
}

.error-box {
    background: #fff5f5;
    border-left: 4px solid var(--color-error);
    border-radius: var(--radius-md);
    padding: var(--spacing-md);
    margin: var(--spacing-md) 0;
    font-size: 0.85rem;
    color: #c53030;
    display: flex;
    align-items: center;
    gap: var(--spacing-sm);
}

/* ========================================
   分割线样式
   ======================================== */
hr {
    border: none;
    height: 2px;
    background: linear-gradient(90deg, transparent 0%, var(--color-border) 50%, transparent 100%);
    margin: var(--spacing-lg) 0;
}

/* ========================================
   响应式设计
   ======================================== */
@media (max-width: 768px) {
    :root {
        --spacing-lg: 16px;
        --spacing-xl: 24px;
    }

    .step-container {
        grid-template-columns: 1fr;
    }

    .function-card-header {
        flex-direction: column;
        text-align: center;
    }

    .function-card-icon {
        margin-right: 0;
        margin-bottom: var(--spacing-sm);
    }
}

/* ========================================
   动画效果
   ======================================== */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(10px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.function-card {
    animation: fadeIn 0.5s ease;
}

/* ========================================
   滚动条样式
   ======================================== */
::-webkit-scrollbar {
    width: 8px;
}

::-webkit-scrollbar-track {
    background: var(--color-bg);
}

::-webkit-scrollbar-thumb {
    background: var(--color-border);
    border-radius: 4px;
}

::-webkit-scrollbar-thumb:hover {
    background: var(--color-secondary);
}
</style>
"""
