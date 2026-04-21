import os
from pathlib import Path
from typing import Any
from urllib.parse import quote

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field
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


class MCPMetaItem(BaseModel):
    label: str
    value: str


class MCPAgendaItem(BaseModel):
    text: str
    level: int = Field(default=1, ge=1, le=3)
    leading_bold: str | None = None


class MCPStyleInput(BaseModel):
    title_font: str = "黑体"
    body_font: str = "仿宋"

    title_size_pt: float = Field(21.5, ge=10, le=72)
    body_size_pt: float = Field(15.5, ge=8, le=48)

    line_spacing: float = Field(1.0, ge=1.0, le=4.0)
    title_bold: bool = True
    label_bold: bool = True

    indent_level1_chars: float = Field(2.0, ge=0.0, le=10.0)
    indent_level2_chars: float = Field(4.0, ge=0.0, le=12.0)
    indent_level3_chars: float = Field(6.0, ge=0.0, le=16.0)


def _normalize_style(style: MCPStyleInput | None) -> dict[str, Any]:
    if style is None:
        return {}
    return style.model_dump()


def _normalize_meta(meta: list[MCPMetaItem]) -> list[dict[str, str]]:
    return [
        {
            "label": item.label,
            "value": item.value,
        }
        for item in meta
    ]


def _normalize_agenda(agenda: list[MCPAgendaItem]) -> list[dict[str, Any]]:
    return [
        {
            "text": item.text,
            "level": item.level,
            "leading_bold": item.leading_bold,
        }
        for item in agenda
    ]


@mcp.tool(
    name="generate_meeting_agenda_docx",
    description=(
        "根据会议议程参数生成 docx 文件。"
        "仅接受顶层平铺参数：title/meta/agenda/style/filename；禁止使用 params 包裹层。"
        "参数类型：title(str, 必填)、meta(list[{label:str,value:str}], 必填)、"
        "agenda(list[{text:str,level:int,leading_bold?:str}], 必填)、"
        "style(object|null, 可选)、filename(str, 可选)。"
        "示例请求：{\"title\":\"4月16日处务例会\",\"meta\":[{\"label\":\"时间\",\"value\":\"2023-04-16 14:00\"}],\"agenda\":[{\"text\":\"一、尚网办项目\",\"level\":1}],\"style\":null,\"filename\":\"4月16日处务例会议程.docx\"}。"
        "返回 code/success/message/data，其中 data 包含 file_id、filename、download_url。"
    ),
)
def generate_meeting_agenda_docx(
    title: str,
    meta: list[MCPMetaItem],
    agenda: list[MCPAgendaItem],
    style: MCPStyleInput | None = None,
    filename: str = "meeting_agenda.docx",
) -> dict[str, Any]:
    """生成会议议程 DOCX。

    强约束说明：
    - 请求体必须是顶层平铺参数对象，不允许 `{"params": {...}}` 包裹层。
    - 顶层仅使用：`title`、`meta`、`agenda`、`style`、`filename`。
    - `meta` 为必填，类型：`list[MCPMetaItem]`，元素结构：`{"label": str, "value": str}`。
    - `agenda` 为必填，类型：`list[MCPAgendaItem]`，元素结构：`{"text": str, "level": int, "leading_bold": str | null}`。
    - 不建议使用未声明字段（如 `item`、`responsible`、`topic`、`speaker`），可能导致参数校验失败。

    参数说明：
    - title: 必填，会议标题。
    - meta: 必填，类型 `list[MCPMetaItem]`，标准数组：[ {"label": "时间", "value": "2026-04-20"} ]。
    - agenda: 必填，类型 `list[MCPAgendaItem]`，标准数组项：{"text": "一、开场", "level": 1, "leading_bold": "一、"}。
    - style: 可选，样式对象；也可传 `null`，表示使用默认样式。
    - filename: 可选，导出文件名，默认 `meeting_agenda.docx`。

    请求示例（平铺式）：
    {
      "title": "上海社科院智算服务平台建设专家座谈会",
      "meta": [{"label": "时　间", "value": "2026年4月3日(周五)13:30"}, {"label": "主　持", "value": "吴雪明"}],
      "agenda": [{"text": "一、专家发言（刘炜）：分享建设建议", "level": 1}],
      "style": {"title_font": "黑体", "body_font": "仿宋", "line_spacing": 1.0},
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
        "meta": _normalize_meta(meta),
        "agenda": _normalize_agenda(agenda),
        "style": _normalize_style(style),
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
