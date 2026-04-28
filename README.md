# llm2word

基于 `FastAPI + docxtpl` 的 DOCX 生成服务。

## 1. 安装依赖

```bash
pip install -r requirements.txt
```

## 2. 启动服务

```bash
uvicorn app.main:app --reload --port 8000
```

## 2.1 启动 MCP SSE 工具服务

```bash
python -m app.mcp_sse_server
```

- 该服务通过 SSE 提供 MCP 工具：`generate_meeting_agenda_docx`
- 工具入参与 `AgendaDocRequest` 一致
- 工具返回与 HTTP 接口一致的结构：`code`、`success`、`message`、`data`

### 2.2 MCP 工具参数说明与请求示例

- MCP SSE 地址（默认）：`http://127.0.0.1:8000/sse`
- 工具名：`generate_meeting_agenda_docx`

参数说明：

- `title`：必填，会议标题。
- `meeting_type`：可选，会议类型，默认 `通知`，当前支持 `通知`、`工作专题会议`。
- `meta`：必填，会议信息数组，每项 `{"label": "时间", "value": "..."}`。
- `attendees`：可选，出席人员数组，默认 `[]`，不能为 `null`。仅 `工作专题会议` 使用。
- `agenda`：必填，议程数组，每项 `{"text": "...", "level": 1, "leading_bold": "..."}`。
- `filename`：可选，导出文件名，默认 `meeting_agenda.docx`。

说明：

- MCP 工具不再对外暴露样式字段。
- 服务端会根据 `meeting_type` 自动选择内置样式配置。
- `通知` 类型对应现有段落式样式。
- `工作专题会议` 类型会生成带表格的通知单样式，建议在 `meta` 中传入 `时间`、`地点`、`主持人`，并通过 `attendees` 传入出席人员。
- 所有字段都不允许传 `null`；无内容的列表字段请传 `[]`。

调用示例（平铺式）：

```json
{
  "title": "上海社科院智算服务平台建设专家座谈会",
  "meeting_type": "通知",
  "meta": [
    {"label": "时　间", "value": "2026年4月3日(周五)13:30"}
  ],
  "attendees": [],
  "agenda": [
    {"text": "一、专家发言", "level": 1}
  ],
  "filename": "专家会议程.docx"
}
```

`工作专题会议` 示例：

```json
{
  "title": "智算平台建设工作专题会",
  "meeting_type": "工作专题会议",
  "meta": [
    {"label": "时间", "value": "2026年4月3日 13:30"},
    {"label": "地点", "value": "105会议室"},
    {"label": "主持人", "value": "张三"}
  ],
  "attendees": ["相关部门负责人", "项目组成员"],
  "agenda": [
    {"text": "一、汇报项目总体进展", "level": 1},
    {"text": "二、讨论当前问题与解决方案", "level": 1}
  ],
  "filename": "工作专题会议.docx"
}
```

返回结构示例：

```json
{
  "code": 200,
  "success": true,
  "message": "docx generated successfully",
  "data": {
    "file_id": "专家会议程-a1b2c3d4",
    "filename": "专家会议程.docx",
    "download_url": "http://59.110.150.224:7800/api/v1/docx/download/%E4%B8%93%E5%AE%B6%E4%BC%9A%E8%AE%AE%E7%A8%8B-a1b2c3d4"
  }
}
```

## 3. 调用接口

- URL: `POST /api/v1/docx/meeting-agenda`
- Content-Type: `application/json`
- 返回：`application/json`（包含 `code`、`success`、`message`、`data`）

- 下载URL：`GET /api/v1/docx/download/{file_id}`
- 下载返回：`application/vnd.openxmlformats-officedocument.wordprocessingml.document`

示例请求体：

```json
{
  "title": "上海社科院智算服务平台建设专家座谈会",
  "meta": [
    {"label": "时　间", "value": "2026年4月3日(周五)13:30"},
    {"label": "地　点", "value": "上海社科院总部105 会议室"},
    {"label": "主　持", "value": "吴雪明,上海社科院副院长"}
  ],
  "agenda": [
    {"text": "一、中国移动技术代表汇报尚社智算服务平台建设情况", "level": 1},
    {"text": "二、院信息办主任赵虹汇报“数智哲社”场景应用示例和智算服务平台安全策略", "level": 1},
    {"text": "三、专家发言", "level": 1},
    {"text": "刘炜,信息研究所所长", "level": 2},
    {"text": "张雪魁,新闻研究所副所长", "level": 2},
    {"text": "五、交流与讨论", "level": 1}
  ],
  "style": {
    "title_font": "黑体",
    "body_font": "仿宋",
    "title_size_pt": 21.5,
    "body_size_pt": 15.5,
    "line_spacing": 1.0,
    "title_bold": true,
    "label_bold": true,
    "indent_level1_chars": 2,
    "indent_level2_chars": 4,
    "indent_level3_chars": 6
  },
  "filename": "专家会议程.docx"
}
```

curl 示例（第一步：生成并获取下载URL）：

```bash
curl -X POST "http://127.0.0.1:8000/api/v1/docx/meeting-agenda" \
  -H "Content-Type: application/json" \
  -d @sample.json
```

第一步返回示例：

```json
{
  "code": 200,
  "success": true,
  "message": "docx generated successfully",
  "data": {
    "file_id": "专家会议程-a1b2c3d4",
    "filename": "专家会议程.docx",
    "download_url": "http://127.0.0.1:8000/api/v1/docx/download/专家会议程-a1b2c3d4"
  }
}
```

curl 示例（第二步：下载文件）：

```bash
curl -L "<第一步返回的download_url>" \
  --output 专家会议程.docx
```
