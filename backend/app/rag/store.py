from __future__ import annotations

from pathlib import Path
from typing import List

import chromadb
from sentence_transformers import SentenceTransformer

from app.core.config import settings
from app.schemas.presentation import KnowledgeChunk


class KnowledgeStore:
    def __init__(self) -> None:
        self.client = None
        self.collection = None
        self.embedder = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")
        try:
            db_path = Path(settings.chroma_dir)
            db_path.mkdir(parents=True, exist_ok=True)
            self.client = chromadb.PersistentClient(path=str(db_path))
            self.collection = self.client.get_or_create_collection(name="text_knowledge")
        except Exception as exc:  # noqa: BLE001
            print(f"[KnowledgeStore] Chroma init failed: {exc}")
            self.client = None
            self.collection = None

    def add_chunks(self, chunks: List[dict]) -> int:
        if not chunks or self.collection is None:
            return 0
        texts = [chunk["page_content"] for chunk in chunks]
        metas = [chunk["metadata"] for chunk in chunks]
        ids = [f"{meta.get('source','doc')}-{i}" for i, meta in enumerate(metas)]
        vectors = self.embedder.encode(texts, normalize_embeddings=True).tolist()
        self.collection.upsert(ids=ids, documents=texts, metadatas=metas, embeddings=vectors)
        return len(chunks)

    def search(self, query: str, k: int = 6) -> List[KnowledgeChunk]:
        if self.collection is None or self.collection.count() == 0:
            return []
        vector = self.embedder.encode([query], normalize_embeddings=True).tolist()
        result = self.collection.query(query_embeddings=vector, n_results=k)
        docs = result.get("documents", [[]])[0]
        metas = result.get("metadatas", [[]])[0]
        distances = result.get("distances", [[]])[0]
        output: List[KnowledgeChunk] = []
        for doc, meta, dist in zip(docs, metas, distances):
            output.append(
                KnowledgeChunk(
                    source=(meta or {}).get("source", "unknown"),
                    content=doc,
                    score=float(1 - dist) if dist is not None else 0.0,
                )
            )
        return output


knowledge_store = KnowledgeStore()
