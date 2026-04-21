## 当前目标
- 完成可运行的 `docxtpl + FastAPI` 工具接口，接收外部 JSON 后返回下载 URL，并通过下载接口获取 DOCX 文件。
- 将会议议程生成功能封装为 SSE 协议的 MCP 工具，同时保留原 HTTP 接口。

## 对应路线图任务
- 关联 `projectRoadmap.md`：目标一、目标二、目标三。

## 当前上下文
- 仓库初始仅包含样式参考文件：`preview.html` 与 `专家会议程.pdf`。
- 已完成最小后端实现，模板由服务启动时自动创建到 `templates/meeting_agenda.docx`。
- 已实现两步下载流程：`POST /api/v1/docx/meeting-agenda` 返回 `download_url`，`GET /api/v1/docx/download/{file_id}` 下载文件。
- 已新增 MCP SSE 服务入口 `app/mcp_sse_server.py`，工具名为 `generate_meeting_agenda_docx`。

## 下一步
- 运行服务并用真实 JSON 请求验证下载结果。
- 对照参考样式细调模板（对齐、缩进、间距、字体细节）。
- 根据业务需要增加更多模板和接口路由。
