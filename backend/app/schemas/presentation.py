from typing import List

from pydantic import BaseModel, Field


class KnowledgeChunk(BaseModel):
    source: str
    content: str
    score: float = 0.0


class GeneratePresentationRequest(BaseModel):
    topic: str = Field(..., description="Тема или идея презентации")
    audience: str = Field(default="Широкая аудитория")
    purpose: str = Field(default="Информировать")
    slide_count: int = Field(default=8, ge=4, le=15)
    tone: str = Field(default="professional")


class SlidePlan(BaseModel):
    slide_number: int
    title: str
    goal: str
    key_points: List[str]


class SlideContent(BaseModel):
    slide_number: int
    title: str
    bullets: List[str]
    speaker_notes: str = ""


class SlideDesign(BaseModel):
    slide_number: int
    layout: str
    visual_idea: str = ""
    icon_hint: str = ""
    emphasis: str = "content"
    background_style: str = "light"
    image_type: str = "none"
    image_path: str | None = None

class ThemePalette(BaseModel):
    theme_name: str = "modern_corporate"
    primary_color: str = "#1F4E79"
    secondary_color: str = "#DCE6F1"
    accent_color: str = "#5B9BD5"
    background_color: str = "#F7F9FC"
    text_color: str = "#1F2937"
    font_family: str = "Aptos"
    style_notes: str = "Минималистичный деловой стиль с умеренными акцентами."


class PresentationPlan(BaseModel):
    title: str
    subtitle: str = ""
    slides: List[SlidePlan]


class PresentationDraft(BaseModel):
    title: str
    subtitle: str = ""
    slides: List[SlideContent]


class PresentationDesign(BaseModel):
    slides: List[SlideDesign]


class FinalSlide(BaseModel):
    slide_number: int
    title: str
    bullets: list[str]
    speaker_notes: str = ""
    layout: str = "title_and_bullets"
    visual_idea: str = ""
    icon_hint: str = ""
    emphasis: str = "content"
    background_style: str = "light"
    image_type: str = "none"
    image_path: str | None = None


class FinalPresentation(BaseModel):
    title: str
    subtitle: str = ""
    audience: str = ""
    purpose: str = ""
    slides: List[FinalSlide]
    sources: List[KnowledgeChunk] = []
    theme: ThemePalette = Field(default_factory=ThemePalette)
