from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse

from app.docx_service import (
    GENERATED_DIR,
    TEMPLATE_PATH,
    build_content_disposition,
    create_meeting_agenda_file,
    resolve_download_record,
)
from app.schemas import AgendaDocRequest
from app.template_builder import ensure_template_exists, write_high_fidelity_template

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
    record = resolve_download_record(file_id)
    if not record:
        raise HTTPException(status_code=404, detail="file_id not found")

    file_path = Path(record["path"])
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="generated file not found")

    headers = {
        "Content-Disposition": build_content_disposition(record["filename"])}
    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
