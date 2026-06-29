# M365 Copilot OpenAI Proxy

将 Microsoft 365 Copilot 暴露为 OpenAI 兼容 API 的 Docker 代理服务。

基于 [m365-copilot-openai-proxy](https://github.com/nicolai-anyhoa/m365-copilot-openai-proxy) 项目，封装为 Docker 镜像，支持：
- Chromium headless 自动刷新 Token
- 多架构镜像 (amd64 + arm64)
- GitHub Actions 自动构建发布到 GHCR
- API Key 认证保护
- CORS 跨域支持

## API 端点

| 端点                          | 说明                                |
| ----------------------------- | ----------------------------------- |
| `GET /healthz`              | 健康检查 + Token 状态               |
| `GET /token/status`         | Token 有效性与过期时间              |
| `GET /v1/models`            | 模型列表                            |
| `POST /v1/chat/completions` | OpenAI Chat Completions（支持流式） |
| `POST /v1/responses`        | OpenAI Responses API（支持流式）    |
| `POST /v1/messages`         | Anthropic Messages API（支持流式）  |

## 快速部署

### 1. 创建 .env 文件

```bash
cp .env.example .env
# 首次可留空 M365_ACCESS_TOKEN，容器启动后由 Chromium headless 自动捕获
```

### 2. 启动服务

```bash
docker compose up -d
```

服务在 `http://localhost:8000` 启动。

### 3. 首次登录 M365

容器内运行 Chromium headless，首次需要登录 M365 账号以获取 session cookie。登录状态保存在 Docker volume `chrome-profile` 中，后续重启无需重新登录。

登录方法（选一种）：

- **VNC 远程桌面**：通过 noVNC 连接容器完成登录
- **手动 Token**：用浏览器油猴脚本获取 Token，写入 `.env`

```bash
# 查看日志确认服务状态
docker compose logs -f
```

## 环境变量

| 变量                       | 必需 | 默认值           | 说明                                        |
| -------------------------- | ---- | ---------------- | ------------------------------------------- |
| `M365_ACCESS_TOKEN`      | 否*  | —               | Substrate Token，留空则由 Chromium 自动捕获 |
| `M365_TIME_ZONE`         | 否   | `Asia/Tokyo`   | 发送给 Copilot 的时区                       |
| `M365_MODEL_ALIAS`       | 否   | `m365-copilot` | 模型名称                                    |
| `API_KEY`                | 否   | —               | API Key 认证密钥，留空则不启用认证          |
| `AUTO_REFRESH`           | 否   | `true`         | 是否自动刷新 Token                          |
| `REFRESH_BEFORE_SECONDS` | 否   | `300`          | Token 过期前多少秒开始刷新                  |
| `CHROME_CDP_PORT`        | 否   | `9222`         | Chromium CDP 端口                           |

## 客户端配置

| 设置             | 值                             |
| ---------------- | ------------------------------ |
| Base URL         | `http://your-server:8000/v1` |
| API Key          | 你设置的 `API_KEY` 值（如未设置则填 `dummy`） |
| Model            | `m365-copilot`               |
| Persistent model | `m365-copilot:persist`       |

### Claude Code

```bash
export ANTHROPIC_BASE_URL=http://your-server:8000
export ANTHROPIC_API_KEY=your-api-key   # 如未设置 API_KEY 则填 dummy
claude
```

### Cherry Studio / OpenCode

```
Base URL: http://your-server:8000/v1
API Key: your-api-key    # 如未设置 API_KEY 则填 dummy
Model: m365-copilot
```

## API Key 认证

在 `.env` 中设置 `API_KEY=your-secret-key` 即可启用认证。所有请求需携带 `Authorization: Bearer your-secret-key` 头。

留空 `API_KEY=` 则不启用认证，任何人都可以直接使用。

```bash
# 带认证的请求示例
curl -H "Authorization: Bearer your-secret-key" http://localhost:8000/v1/models
```

## 持久会话

- **Header 模式**：请求头 `X-M365-Session-Id: my-session`
- **模型后缀模式**：使用模型名 `m365-copilot:persist`

两种方式都会在同一 Copilot 对话中保留上下文。

## 手动获取 Token（推荐）

项目提供了 Tampermonkey 油猴脚本 `get_token.js`，可在浏览器中一键提取 Substrate Token：

### 步骤

1. 安装 [Tampermonkey](https://www.tampermonkey.net/) 浏览器扩展
2. 点击 Tampermonkey 图标 → **添加新脚本**
3. 将 `get_token.js` 的内容粘贴进去，保存
4. 打开 [M365 Copilot](https://m365.cloud.microsoft/chat) 并登录你的 M365 账号
5. 在 Copilot 对话框中**输入任意字符**（触发 WebSocket 连接）
6. 页面右上角会弹出 Token 提取面板
7. 点击 **Copy Token** 复制 Token 值
8. 将 Token 写入 `.env` 文件：

```bash
M365_ACCESS_TOKEN=粘贴的token值
```

9. 重启容器使其生效：

```bash
docker compose restart
```

> **Token 约 1 小时过期**。容器内的 Chromium headless 会自动刷新（如果启用了 `AUTO_REFRESH=true`），无需手动更新。
> 如果禁用了自动刷新，Token 过期后需要重复上述步骤手动获取。

### 也支持 Chromium headless 自动捕获

如果不想手动获取 Token，可以留空 `M365_ACCESS_TOKEN`，容器启动后由 Chromium headless 自动捕获。但首次需要在容器中登录 M365 账号（通过 VNC 或手动注入 Cookie）。

推荐使用油猴脚本方式，更简单直观。

## 架构

```
容器启动
  ├─ Chromium headless → m365.cloud.microsoft/chat (CDP 端口 9222)
  │   └─ 登录状态持久化于 /chrome-profile volume
  │
  └─ copilot-openai-proxy serve
      ├─ 通过 CDP 自动捕获 Substrate WebSocket Token
      ├─ Token 过期前 5 分钟自动刷新
      └─ 提供 OpenAI 兼容 API (端口 8000)
```

## License

Apache License 2.0
