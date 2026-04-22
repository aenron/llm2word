import os
from pathlib import Path
from typing import Any
from urllib.parse import quote

from mcp.server.fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import FileResponse, JSONResponse
import uvicorn
from mcp.server.transport_security import TransportSecuritySettings

from app.main import (
    GENERATED_DIR,
    TEMPLATE_PATH,
    _build_content_disposition,
    _resolve_download_record,
    create_meeting_agenda_file,
)
from app.schemas import AgendaDocRequest
from app.template_builder import ensure_template_exists

mcp = FastMCP(
    "llm2word-mcp",
    transport_security=TransportSecuritySettings(
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
    ),
)

MCP_DOWNLOAD_BASE_URL = os.getenv(
    "MCP_DOWNLOAD_BASE_URL", "http://127.0.0.1:8000")


def _bootstrap_runtime() -> None:
    ensure_template_exists(TEMPLATE_PATH)
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)


def _build_style(
    title_font: str,
    body_font: str,
    title_size_pt: float,
    body_size_pt: float,
    line_spacing: float,
    title_bold: bool,
    label_bold: bool,
    indent_level1_chars: float,
    indent_level2_chars: float,
    indent_level3_chars: float,
) -> dict[str, Any]:
    return {
        "title_font": title_font,
        "body_font": body_font,
        "title_size_pt": title_size_pt,
        "body_size_pt": body_size_pt,
        "line_spacing": line_spacing,
        "title_bold": title_bold,
        "label_bold": label_bold,
        "indent_level1_chars": indent_level1_chars,
        "indent_level2_chars": indent_level2_chars,
        "indent_level3_chars": indent_level3_chars,
    }


@mcp.tool(
    name="generate_meeting_agenda_docx",
    description=(
        "根据会议议程参数生成 docx 文件。"
        "仅接受顶层平铺参数：title/meta/agenda/filename 与 style 相关可选字段；禁止使用 params 包裹层。"
        "参数类型：title(str, 必填)、meta(list[object], 必填)、"
        "agenda(list[object], 必填)、"
        "filename(str, 可选)、title_font(str, 默认黑体)、body_font(str, 默认仿宋)、title_size_pt(number, 默认21.5)、body_size_pt(number, 默认15.5)、line_spacing(number, 默认1.0)、title_bold(bool, 默认true)、label_bold(bool, 默认true)、indent_level1_chars(number, 默认2.0)、indent_level2_chars(number, 默认4.0)、indent_level3_chars(number, 默认6.0)。"
        "示例请求：{\"title\":\"4月16日处务例会\",\"meta\":[{\"label\":\"时间\",\"value\":\"2023-04-16 14:00\"}],\"agenda\":[{\"text\":\"一、尚网办项目\",\"level\":1}],\"title_font\":\"黑体\",\"body_font\":\"仿宋\",\"line_spacing\":1.0,\"filename\":\"4月16日处务例会议程.docx\"}。"
        "返回 code/success/message/data，其中 data 包含 file_id、filename、download_url。"
    ),
)
def generate_meeting_agenda_docx(
    title: str,
    meta: list[dict[str, str]],
    agenda: list[dict[str, Any]],
    filename: str = "meeting_agenda.docx",
    title_font: str = "黑体",
    body_font: str = "仿宋",
    title_size_pt: float = 21.5,
    body_size_pt: float = 15.5,
    line_spacing: float = 1.0,
    title_bold: bool = True,
    label_bold: bool = True,
    indent_level1_chars: float = 2.0,
    indent_level2_chars: float = 4.0,
    indent_level3_chars: float = 6.0,
) -> dict[str, Any]:
    """生成会议议程 DOCX。

    强约束说明：
    - 请求体必须是顶层平铺参数对象，不允许 `{"params": {...}}` 包裹层。
    - 顶层仅使用：`title`、`meta`、`agenda`、`filename` 与各个样式可选字段。
    - `meta` 为必填，类型：`list[dict[str, str]]`，每项至少包含 `label` 和 `value`。
    - `agenda` 为必填，类型：`list[dict[str, Any]]`，每项至少包含 `text`，可选 `level`、`leading_bold`。
    - 样式字段均提供默认值，支持：`title_font`、`body_font`、`title_size_pt`、`body_size_pt`、`line_spacing`、`title_bold`、`label_bold`、`indent_level1_chars`、`indent_level2_chars`、`indent_level3_chars`。
    - 不建议使用未声明字段（如 `item`、`responsible`、`topic`、`speaker`），可能导致最终业务校验失败。

    参数说明：
    - title: 必填，会议标题。
    - meta: 必填，类型 `list[dict[str, str]]`，标准数组：[ {"label": "时间", "value": "2026-04-20"} ]。
    - agenda: 必填，类型 `list[dict[str, Any]]`，标准数组项：{"text": "一、开场", "level": 1, "leading_bold": "一、"}。
    - filename: 可选，导出文件名，默认 `meeting_agenda.docx`。
    - title_font/body_font: 字体名称，默认分别为 `黑体`、`仿宋`。
    - title_size_pt/body_size_pt/line_spacing: 数字类型，默认分别为 `21.5`、`15.5`、`1.0`。
    - title_bold/label_bold: 布尔类型，默认均为 `true`。
    - indent_level1_chars/indent_level2_chars/indent_level3_chars: 数字类型，默认分别为 `2.0`、`4.0`、`6.0`。

    请求示例（平铺式）：
    {
      "title": "上海社科院智算服务平台建设专家座谈会",
      "meta": [{"label": "时　间", "value": "2026年4月3日(周五)13:30"}, {"label": "主　持", "value": "吴雪明"}],
      "agenda": [{"text": "一、专家发言（刘炜）：分享建设建议", "level": 1}],
      "title_font": "黑体",
      "body_font": "仿宋",
      "line_spacing": 1.0,
      "filename": "专家会议程.docx"
    }

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
        "meta": meta,
        "agenda": agenda,
        "style": _build_style(
            title_font=title_font,
            body_font=body_font,
            title_size_pt=title_size_pt,
            body_size_pt=body_size_pt,
            line_spacing=line_spacing,
            title_bold=title_bold,
            label_bold=label_bold,
            indent_level1_chars=indent_level1_chars,
            indent_level2_chars=indent_level2_chars,
            indent_level3_chars=indent_level3_chars,
        ),
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
    record = _resolve_download_record(file_id)
    if not record:
        return JSONResponse(status_code=404, content={"detail": "file_id not found"})

    file_path = Path(record["path"])
    if not file_path.exists():
        return JSONResponse(status_code=404, content={"detail": "generated file not found"})

    headers = {
        "Content-Disposition": _build_content_disposition(record["filename"])}
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
