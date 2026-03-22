from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse

from app.agents.nodes import run_presentation_graph
from app.schemas.presentation import GeneratePresentationRequest
from app.services.pptx_exporter import export_presentation

router = APIRouter()


@router.post("/generate")
async def generate_presentation(payload: GeneratePresentationRequest):
    try:
        result = await run_presentation_graph(payload)
        pptx_path = export_presentation(result)
        return FileResponse(
            path=pptx_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=pptx_path.name,
        )
    except Exception as exc:  # noqa: BLE001
        message = str(exc)
        if "insufficient_quota" in message or "429" in message:
            raise HTTPException(status_code=402, detail="Недостаточно средств/квоты OpenAI API. Пополни billing и повтори.")
        raise HTTPException(status_code=500, detail=message)
