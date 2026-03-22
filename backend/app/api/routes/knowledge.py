from pathlib import Path

from fastapi import APIRouter, File, UploadFile
from fastapi.concurrency import run_in_threadpool

from app.rag.ingest import ingest_file_to_store

router = APIRouter()


@router.post("/upload")
async def upload_knowledge(file: UploadFile = File(...)) -> dict:
    suffix = Path(file.filename or "document.pdf").suffix or ".pdf"
    tmp_path = Path("./data/uploads") / (Path(file.filename or "document").stem + suffix)
    tmp_path.parent.mkdir(parents=True, exist_ok=True)
    content = await file.read()
    tmp_path.write_bytes(content)
    added = await run_in_threadpool(ingest_file_to_store, tmp_path)
    return {"status": "ok", "chunks_added": added, "filename": file.filename}
