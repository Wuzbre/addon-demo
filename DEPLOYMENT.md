# 部署指南

## 域名部署配置

你的应用已经配置为部署到 `dingdocs-plugin.yuce-tech.cn` 域名。以下是部署步骤和配置说明。

## 已完成的配置修改

### 1. WebSocket配置
- 修改了 `config/webpackDevServer.config.js`，设置默认WebSocket主机为你的域名
- 配置了WSS协议用于安全连接
- 启用了跨域访问支持

### 2. 启动脚本配置
- 修改了 `scripts/start.js`，强制使用HTTPS
- 设置了正确的端口（443）
- 配置了WebSocket连接参数

### 3. Manifest配置
- `manifest.json` 中已正确配置了域名URL
- Script服务URL: `https://dingdocs-plugin.yuce-tech.cn/script.html`
- UI服务URL: `https://dingdocs-plugin.yuce-tech.cn/ui.html`

## 部署步骤

### 1. 服务器要求
- Node.js 环境
- HTTPS证书（用于WSS连接）
- 域名解析到服务器IP

### 2. 安装依赖
```bash
npm install
```

### 3. 启动开发服务器
```bash
npm start
```

### 4. 生产环境构建
```bash
npm run build
```

## 环境变量配置

如果需要自定义配置，可以设置以下环境变量：

```bash
# 强制使用HTTPS
HTTPS=true

# 设置端口
PORT=443

# WebSocket配置
WDS_SOCKET_HOST=dingdocs-plugin.yuce-tech.cn
WDS_SOCKET_PORT=443
WDS_SOCKET_PATH=/ws

# 允许所有主机访问（用于域名部署）
DANGEROUSLY_DISABLE_HOST_CHECK=true
```

## 注意事项

1. **HTTPS证书**：确保服务器有有效的SSL证书
2. **防火墙**：确保443端口对外开放
3. **域名解析**：确保域名正确解析到服务器IP
4. **WebSocket支持**：确保服务器支持WebSocket连接

## 故障排除

### WebSocket连接失败
- 检查HTTPS证书是否有效
- 确认防火墙设置允许443端口
- 验证域名解析是否正确

### 热重载不工作
- 检查WebSocket配置是否正确
- 确认浏览器控制台是否有WebSocket连接错误
- 验证WSS协议是否正常工作

## 测试部署

部署完成后，访问以下URL测试：
- UI界面: `https://dingdocs-plugin.yuce-tech.cn/ui.html`
- Script服务: `https://dingdocs-plugin.yuce-tech.cn/script.html`

如果页面正常加载且热重载功能正常，说明部署成功。
