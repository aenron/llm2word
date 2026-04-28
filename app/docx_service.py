from pathlib import Path
import re
from urllib.parse import quote
from uuid import uuid4

from pypinyin import lazy_pinyin

from app.render_service import render_agenda_docx
from app.schemas import AgendaDocRequest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "meeting_agenda.docx"
GENERATED_DIR = PROJECT_ROOT / "generated_files"
DOWNLOAD_BASE_URL = "http://59.110.150.224:7800"

generated_file_index: dict[str, dict[str, str]] = {}
FILE_ID_PATTERN = re.compile(r"^(?P<stem>.+)-(?P<short_id>[0-9a-f]{8})$")


def build_ascii_fallback(filename: str) -> str:
    stem = Path(filename).stem
    ext = Path(filename).suffix or ".docx"

    pinyin_stem = "".join(lazy_pinyin(stem))
    slug = re.sub(r"[^a-zA-Z0-9._-]+", "-", pinyin_stem).strip("-._")

    if not slug:
        slug = "download"

    return f"{slug}{ext}"


def build_content_disposition(filename: str) -> str:
    ascii_fallback = build_ascii_fallback(filename)
    try:
        filename.encode("latin-1")
        ascii_fallback = filename
    except UnicodeEncodeError:
        pass

    encoded = quote(filename, safe="")
    return f"attachment; filename=\"{ascii_fallback}\"; filename*=UTF-8''{encoded}"


def normalize_filename(filename: str) -> str:
    normalized = filename.strip() or "meeting_agenda.docx"
    normalized = normalized.replace("/", "_").replace("\\", "_")
    if not normalized.lower().endswith(".docx"):
        normalized = f"{normalized}.docx"
    return normalized


def create_meeting_agenda_file(payload: AgendaDocRequest) -> dict[str, str]:
    file_bytes = render_agenda_docx(payload, template_path=TEMPLATE_PATH)

    safe_name = normalize_filename(payload.filename)

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


def resolve_download_record(file_id: str) -> dict[str, str] | None:
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

    filename = normalize_filename(f"{stem}.docx")
    return {
        "path": str(file_path),
        "filename": filename,
    }
