import os
from pathlib import Path
from typing import Any, Literal
from urllib.parse import quote

from mcp.server.fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import FileResponse, JSONResponse
import uvicorn

try:
    from mcp.server.transport_security import TransportSecuritySettings
except ModuleNotFoundError:  # Compatibility fallback for MCP builds without this module.
    TransportSecuritySettings = None

from app.docx_service import (
    GENERATED_DIR,
    TEMPLATE_PATH,
    build_content_disposition,
    create_meeting_agenda_file,
    resolve_download_record,
)
from app.schemas import AgendaDocRequest
from app.template_builder import ensure_template_exists

_transport_security = None
if TransportSecuritySettings is not None:
    _transport_security = TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=[
            "127.0.0.1:8000",
            "localhost:8000",
            "59.110.150.224:8000",
        ],
        allowed_origins=[
            "http://127.0.0.1:8000",
            "http://localhost:8000",
            "http://59.110.150.224:8000",
        ],
    )

mcp = FastMCP("llm2word-mcp", transport_security=_transport_security)

MCP_DOWNLOAD_BASE_URL = os.getenv(
    "MCP_DOWNLOAD_BASE_URL", "http://127.0.0.1:8000")

MEETING_TYPE_STYLES: dict[str, dict[str, Any]] = {
    "通知": {
        "title_font": "黑体",
        "body_font": "仿宋",
        "title_size_pt": 21.5,
        "body_size_pt": 15.5,
        "line_spacing": 1.0,
        "title_bold": True,
        "label_bold": True,
        "indent_level1_chars": 2.0,
        "indent_level2_chars": 4.0,
        "indent_level3_chars": 6.0,
    },
    "工作专题会议": {
        "title_font": "黑体",
        "body_font": "仿宋",
        "title_size_pt": 22.0,
        "body_size_pt": 15.5,
        "line_spacing": 1.0,
        "title_bold": True,
        "label_bold": True,
        "indent_level1_chars": 2.0,
        "indent_level2_chars": 4.0,
        "indent_level3_chars": 6.0,
    },
}


def _bootstrap_runtime() -> None:
    ensure_template_exists(TEMPLATE_PATH)
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)


def _build_style(meeting_type: str) -> dict[str, Any]:
    style = MEETING_TYPE_STYLES.get(meeting_type)
    if style is None:
        supported = "、".join(MEETING_TYPE_STYLES)
        raise ValueError(f"不支持的 meeting_type: {meeting_type}，当前支持：{supported}")
    return dict(style)


@mcp.tool(
    name="generate_meeting_agenda_docx",
    description=(
        "根据会议通知信息生成 docx 文件。"
        "请求体必须是顶层平铺参数对象，只允许：title、meta、attendees、agenda、meeting_type、filename；禁止使用 params 包裹层。"
        "字段类型：title(str, 必填)、meta(list[object], 必填)、agenda(list[object], 必填)、attendees(list[str], 可选，默认空数组)、meeting_type(str, 可选，默认“通知”，支持“通知”“工作专题会议”)、filename(str, 可选)。"
        "所有字段都不能为 null；列表字段无内容时传空数组 []。"
        "当 meeting_type=\"工作专题会议\" 时，attendees 用于渲染“出席范围”人员列表。若原文中出现“出席范围”“参会人员”“参会名单”“与会人员”“参加人员”等信息，应优先抽取为 attendees，每人一项，不要把多人拼成一个长字符串，也不要只写到 meta。"
        "当 meeting_type=\"通知\" 时，attendees 通常传空数组 []."
        "meta 适合承载单值字段，例如时间、地点、主持人；agenda 适合承载议程条目；attendees 只承载人员列表。"
        "示例1（通知）：{\"title\":\"4月16日处务例会\",\"meeting_type\":\"通知\",\"meta\":[{\"label\":\"时间\",\"value\":\"2023-04-16 14:00\"}],\"attendees\":[],\"agenda\":[{\"text\":\"一、尚网办项目\",\"level\":1}],\"filename\":\"4月16日处务例会议程.docx\"}。"
        "示例2（工作专题会议）：{\"title\":\"智算平台建设工作专题会\",\"meeting_type\":\"工作专题会议\",\"meta\":[{\"label\":\"时间\",\"value\":\"2026年4月3日 13:30\"},{\"label\":\"地点\",\"value\":\"105会议室\"},{\"label\":\"主持人\",\"value\":\"张三\"}],\"attendees\":[\"李四\",\"王五\",\"赵六\"],\"agenda\":[{\"text\":\"一、汇报项目总体进展\",\"level\":1}],\"filename\":\"工作专题会议.docx\"}。"
        "返回 code/success/message/data，其中 data 包含 file_id、filename、download_url。"
    ),
)
def generate_meeting_agenda_docx(
    title: str,
    meta: list[dict[str, str]],
    agenda: list[dict[str, Any]],
    attendees: list[str] = [],
    meeting_type: Literal["通知", "工作专题会议"] = "通知",
    filename: str = "meeting_agenda.docx",
) -> dict[str, Any]:
    """生成会议议程 DOCX。

    强约束说明：
    - 请求体必须是顶层平铺参数对象，不允许 `{"params": {...}}` 包裹层。
    - 顶层仅使用：`title`、`meta`、`attendees`、`agenda`、`meeting_type`、`filename`。
    - 所有字段均不允许为 `null`；列表字段无内容时请传空数组 `[]`。
    - `meta` 为必填，类型：`list[dict[str, str]]`，每项至少包含 `label` 和 `value`，用于单值字段。
    - `attendees` 为可选，类型：`list[str]`，默认空数组；用于出席人员列表。
    - `agenda` 为必填，类型：`list[dict[str, Any]]`，每项至少包含 `text`，可选 `level`、`leading_bold`。
    - 样式字段不再对外暴露，服务端会根据 `meeting_type` 自动选择内置样式。
    - 当前支持 `meeting_type="通知"` 与 `meeting_type="工作专题会议"`，默认值为 `通知`。
    - 当 `meeting_type="工作专题会议"` 时，应优先从原文中抽取人员名单填入 `attendees`。只要原文出现“出席范围”“参会人员”“参会名单”“与会人员”“参加人员”等信息，就应尽量生成 `attendees`。
    - `attendees` 的每个元素只放一个人名或一个参会对象；不要把多个人拼在同一个字符串里。
    - `meta` 不要再承担人员列表，人员列表优先放到 `attendees`。
    - 当 `meeting_type="通知"` 时，`attendees` 一般传 `[]`。
    - 不建议使用未声明字段（如 `item`、`responsible`、`topic`、`speaker`），可能导致最终业务校验失败。

    参数说明：
    - title: 必填，会议标题。
    - meta: 必填，类型 `list[dict[str, str]]`，标准数组：[ {"label": "时间", "value": "2026-04-20"} ]。
    - attendees: 可选，类型 `list[str]`，默认 `[]`，不允许传 `null`。工作专题会议时，如果原文有人员信息，应优先填充该字段。
    - agenda: 必填，类型 `list[dict[str, Any]]`，标准数组项：{"text": "一、开场", "level": 1, "leading_bold": "一、"}。
    - meeting_type: 可选，会议类型，默认 `通知`，当前支持：`通知`、`工作专题会议`。
    - filename: 可选，导出文件名，默认 `meeting_agenda.docx`。

    请求示例（平铺式）：
    {
      "title": "上海社科院智算服务平台建设专家座谈会",
      "meeting_type": "通知",
      "meta": [{"label": "时　间", "value": "2026年4月3日(周五)13:30"}, {"label": "主　持", "value": "吴雪明"}],
      "attendees": [],
      "agenda": [{"text": "一、专家发言（刘炜）：分享建设建议", "level": 1}],
      "filename": "专家会议程.docx"
    }

    表格式示例（工作专题会议）：
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

    attendees 抽取示例：
    - 原文出现“出席范围：张三、李四、王五”
      则应传：`"attendees": ["张三", "李四", "王五"]`
    - 原文出现“参会人员：办公室、财务处、信息中心”
      则应传：`"attendees": ["办公室", "财务处", "信息中心"]`
    - 不要传：`"attendees": ["张三、李四、王五"]`
    - 不要只写：`{"label": "出席范围", "value": "张三、李四、王五"}` 而省略 `attendees`

    错误示例1（错误的 params 包裹层）：
    {
      "params": {
        "title": "4月16日处务例会"
      }
    }

    错误示例2（不支持的 agenda 字段名）：
    {
      "title": "4月16日处务例会",
      "agenda": [{"item": "尚网办项目", "responsible": "黄雪莲", "details": "需细分"}]
    }
    """
    _bootstrap_runtime()

    request_data = {
        "title": title,
        "meeting_type": meeting_type,
        "meta": meta,
        "attendees": attendees,
        "agenda": agenda,
        "style": _build_style(meeting_type),
        "filename": filename,
    }

    request = AgendaDocRequest.model_validate(request_data)
    data = create_meeting_agenda_file(request)
    data["download_url"] = f"{MCP_DOWNLOAD_BASE_URL}/api/v1/docx/download/{quote(data['file_id'], safe='')}"
    return {
        "code": 200,
        "success": True,
        "message": "docx generated successfully",
        "data": data,
    }


async def download_generated_docx(request: Request) -> FileResponse | JSONResponse:
    file_id = request.path_params["file_id"]
    record = resolve_download_record(file_id)
    if not record:
        return JSONResponse(status_code=404, content={"detail": "file_id not found"})

    file_path = Path(record["path"])
    if not file_path.exists():
        return JSONResponse(status_code=404, content={"detail": "generated file not found"})

    headers = {
        "Content-Disposition": build_content_disposition(record["filename"])}
    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )


if __name__ == "__main__":
    app = mcp.sse_app()
    app.add_route(
        "/api/v1/docx/download/{file_id}", download_generated_docx, methods=["GET"])
    uvicorn.run(app, host="0.0.0.0", port=8000)
