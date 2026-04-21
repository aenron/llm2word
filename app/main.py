from pathlib import Path
import re
from urllib.parse import quote
from uuid import uuid4

from fastapi import FastAPI, HTTPException, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse
from pypinyin import lazy_pinyin

from app.render_service import render_agenda_docx
from app.schemas import AgendaDocRequest
from app.template_builder import ensure_template_exists, write_high_fidelity_template

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "meeting_agenda.docx"
GENERATED_DIR = PROJECT_ROOT / "generated_files"
DOWNLOAD_BASE_URL = "http://59.110.150.224:7800"

generated_file_index: dict[str, dict[str, str]] = {}
FILE_ID_PATTERN = re.compile(r"^(?P<stem>.+)-(?P<short_id>[0-9a-f]{8})$")

app = FastAPI(
    title="LLM DOCX Tool API",
    version="0.1.0",
    description="接收外部 JSON，生成并返回 DOCX 文件",
)


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    raw_body = await request.body()
    print("RAW REQUEST BODY (422):", raw_body.decode("utf-8", errors="ignore"))
    print("VALIDATION ERRORS:", exc.errors())
    return JSONResponse(status_code=422, content={"detail": exc.errors()})


def _build_ascii_fallback(filename: str) -> str:
    stem = Path(filename).stem
    ext = Path(filename).suffix or ".docx"

    pinyin_stem = "".join(lazy_pinyin(stem))
    slug = re.sub(r"[^a-zA-Z0-9._-]+", "-", pinyin_stem).strip("-._")

    if not slug:
        slug = "download"

    return f"{slug}{ext}"


def _build_content_disposition(filename: str) -> str:
    ascii_fallback = _build_ascii_fallback(filename)
    try:
        filename.encode("latin-1")
        ascii_fallback = filename
    except UnicodeEncodeError:
        pass

    encoded = quote(filename, safe="")
    return f"attachment; filename=\"{ascii_fallback}\"; filename*=UTF-8''{encoded}"


def _normalize_filename(filename: str) -> str:
    normalized = filename.strip() or "meeting_agenda.docx"
    normalized = normalized.replace("/", "_").replace("\\", "_")
    if not normalized.lower().endswith(".docx"):
        normalized = f"{normalized}.docx"
    return normalized


def create_meeting_agenda_file(payload: AgendaDocRequest) -> dict[str, str]:
    file_bytes = render_agenda_docx(payload, template_path=TEMPLATE_PATH)

    safe_name = _normalize_filename(payload.filename)

    short_id = uuid4().hex[:8]
    file_id = f"{Path(safe_name).stem}-{short_id}"
    stored_path = GENERATED_DIR / f"{short_id}.docx"
    stored_path.write_bytes(file_bytes)

    generated_file_index[file_id] = {
        "path": str(stored_path),
        "filename": safe_name,
    }

    download_url = f"{DOWNLOAD_BASE_URL}/api/v1/docx/download/{quote(file_id, safe='')}"
    return {
        "file_id": file_id,
        "filename": safe_name,
        "download_url": download_url,
    }


def _resolve_download_record(file_id: str) -> dict[str, str] | None:
    record = generated_file_index.get(file_id)
    if record:
        return record

    match = FILE_ID_PATTERN.match(file_id)
    if not match:
        return None

    short_id = match.group("short_id")
    stem = match.group("stem").strip()
    if not stem:
        return None

    file_path = GENERATED_DIR / f"{short_id}.docx"
    if not file_path.exists():
        return None

    filename = _normalize_filename(f"{stem}.docx")
    return {
        "path": str(file_path),
        "filename": filename,
    }


@app.on_event("startup")
def bootstrap_template() -> None:
    ensure_template_exists(TEMPLATE_PATH)
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/v1/templates/meeting-agenda/rebuild")
def rebuild_meeting_agenda_template() -> dict[str, str]:
    write_high_fidelity_template(TEMPLATE_PATH, overwrite=True)
    return {
        "status": "ok",
        "template": str(TEMPLATE_PATH),
        "message": "high-fidelity template rebuilt",
    }


@app.post("/api/v1/docx/meeting-agenda")
async def generate_meeting_agenda(payload: AgendaDocRequest, request: Request) -> dict[str, object]:
    raw_body = await request.body()
    print("RAW REQUEST BODY:", raw_body.decode("utf-8", errors="ignore"))
    data = create_meeting_agenda_file(payload)
    return {
        "code": 200,
        "success": True,
        "message": "docx generated successfully",
        "data": data,
    }


@app.get("/api/v1/docx/download/{file_id}", name="download_generated_docx")
def download_generated_docx(file_id: str) -> FileResponse:
    record = _resolve_download_record(file_id)
    if not record:
        raise HTTPException(status_code=404, detail="file_id not found")

    file_path = Path(record["path"])
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="generated file not found")

    headers = {
        "Content-Disposition": _build_content_disposition(record["filename"])}
    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
