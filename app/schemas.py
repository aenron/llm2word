import json
from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator


class MetaItem(BaseModel):
    label: str = Field(..., description="元信息标签，如：时间、地点")
    value: str = Field(..., description="元信息内容")


class AgendaItem(BaseModel):
    text: str = Field(..., description="议程文本")
    level: int = Field(1, ge=1, le=3, description="层级，1 为主层级")
    leading_bold: str | None = Field(
        default=None,
        description="可选：行首需要加粗的前缀文本",
    )


class RenderStyle(BaseModel):
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


class AgendaDocRequest(BaseModel):
    title: str
    meeting_type: Literal["通知", "工作专题会议"] = "通知"
    meta: list[MetaItem] = Field(default_factory=list)
    attendees: list[str] = Field(default_factory=list)
    agenda: list[AgendaItem] = Field(default_factory=list)
    style: RenderStyle = Field(default_factory=RenderStyle)
    filename: str = Field("meeting_agenda.docx", description="返回文件名")

    @field_validator("meta", mode="before")
    @classmethod
    def parse_meta(cls, value: Any) -> Any:
        if value is None:
            raise ValueError("meta 不能为 null")
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return []
            try:
                return json.loads(text)
            except json.JSONDecodeError as exc:
                raise ValueError("meta 不是合法的 JSON 字符串") from exc
        return value

    @field_validator("agenda", mode="before")
    @classmethod
    def parse_agenda(cls, value: Any) -> Any:
        if value is None:
            raise ValueError("agenda 不能为 null")
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return []
            try:
                return json.loads(text)
            except json.JSONDecodeError as exc:
                raise ValueError("agenda 不是合法的 JSON 字符串") from exc
        return value

    @field_validator("attendees", mode="before")
    @classmethod
    def parse_attendees(cls, value: Any) -> Any:
        if value is None:
            raise ValueError("attendees 不能为 null")
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return []
            try:
                return json.loads(text)
            except json.JSONDecodeError as exc:
                raise ValueError("attendees 不是合法的 JSON 字符串") from exc
        return value

    @field_validator("style", mode="before")
    @classmethod
    def parse_style(cls, value: Any) -> Any:
        if value is None:
            raise ValueError("style 不能为 null")
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return {}
            try:
                return json.loads(text)
            except json.JSONDecodeError as exc:
                raise ValueError("style 不是合法的 JSON 字符串") from exc
        return value
