## 关键组件与交互
### `app/main.py`
- FastAPI 入口。
- 启动时确保模板存在。
- 提供 `POST /api/v1/docx/meeting-agenda`，接收 JSON 并返回下载地址。
- 提供 `GET /api/v1/docx/download/{file_id}`，按文件ID返回 docx 文件流。
- 抽取 `create_meeting_agenda_file` 作为可复用生成能力。

### `app/mcp_sse_server.py`
- MCP SSE 服务入口。
- 提供工具 `generate_meeting_agenda_docx`。
- 复用 `app/main.py` 中的 `create_meeting_agenda_file`，保证与 HTTP 接口行为一致。

### `app/schemas.py`
- 定义请求体：标题、元信息、议程、样式。
- 通过字段约束保证接口输入质量。

### `app/template_builder.py`
- 在模板缺失时自动创建 `templates/meeting_agenda.docx`。
- 生成占位符：`title`、`meta_subdoc`、`agenda_subdoc`。

### `app/render_service.py`
- 负责 docxtpl 渲染主流程。
- 构建 `meta_subdoc` 与 `agenda_subdoc`，支持行距、缩进、行首加粗等样式能力。

## 数据流
- 外部系统发送 JSON 请求 -> FastAPI 路由解析为 Pydantic 模型 -> render_service 渲染并落盘 -> 返回 `download_url` -> 客户端访问下载接口获取 docx 字节流。
- MCP 客户端通过 SSE 调用工具 -> 参数校验为 `AgendaDocRequest` -> 复用同一生成逻辑 -> 返回与 HTTP 接口一致的结构化结果。

## 外部依赖
- fastapi / uvicorn
- docxtpl
- python-docx
- pydantic

## 最近重要变更
- 新增完整最小可运行后端结构。
- 新增模板自动创建能力，避免初始仓库缺少模板导致不可运行。
- 接口设计改为纯 JSON 输入，满足外部调用场景。
- 文档生成接口改为两步下载模式，支持先获取下载URL再下载文件。
- 新增 SSE 协议 MCP 工具封装，保留原 HTTP 接口不变。

## 用户反馈整合
- 用户强调“由外部提供 JSON，接口负责返回生成文件”。
- 已按该要求实现外部 JSON 输入；并按最新反馈改为返回下载URL的两步下载模式。

## 附加文档
- 当前暂无附加文档。
