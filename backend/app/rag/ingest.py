from pathlib import Path
from typing import List

from pypdf import PdfReader

from app.rag.store import knowledge_store


def chunk_text(text: str, source: str, chunk_size: int = 1200, overlap: int = 150) -> List[dict]:
    cleaned = " ".join(text.split())
    chunks: List[dict] = []
    start = 0
    while start < len(cleaned):
        end = min(start + chunk_size, len(cleaned))
        content = cleaned[start:end]
        chunks.append({"page_content": content, "metadata": {"source": source}})
        if end == len(cleaned):
            break
        start = max(0, end - overlap)
    return chunks


def extract_text(file_path: Path) -> str:
    if file_path.suffix.lower() == ".pdf":
        reader = PdfReader(str(file_path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    return file_path.read_text(encoding="utf-8", errors="ignore")


def ingest_file_to_store(file_path: Path) -> int:
    text = extract_text(file_path)
    chunks = chunk_text(text, source=file_path.name)
    return knowledge_store.add_chunks(chunks)
