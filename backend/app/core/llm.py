import os

from langchain_openai import ChatOpenAI

from app.core.config import settings


def get_model(model: str | None = None) -> ChatOpenAI:
    model_name = model or settings.openai_model
    return ChatOpenAI(
        model=model_name,
        temperature=0.35,
        api_key=settings.openai_api_key,
        base_url=settings.openai_base_url,
        default_headers={
            "X-Title": "AI Presentation Generator",
            "HTTP-Referer": "http://localhost:5173",
        },
    )
