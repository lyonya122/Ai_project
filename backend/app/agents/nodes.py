from __future__ import annotations

from typing import List, TypedDict

from langgraph.graph import END, StateGraph
from pydantic import BaseModel, Field

from app.core.llm import get_model
from app.rag.store import knowledge_store
from app.schemas.presentation import (
    FinalPresentation,
    FinalSlide,
    GeneratePresentationRequest,
    PresentationDesign,
    PresentationDraft,
    PresentationPlan,
    SlideContent,
    SlideDesign,
    SlidePlan,
    ThemePalette,
)


class PlannerOut(BaseModel):
    title: str
    subtitle: str = ""
    slides: List[SlidePlan] = Field(default_factory=list)


class WriterOut(BaseModel):
    slides: List[SlideContent] = Field(default_factory=list)


class DesignerOut(BaseModel):
    slides: List[SlideDesign] = Field(default_factory=list)


class ArtDirectorOut(BaseModel):
    theme: ThemePalette = Field(default_factory=ThemePalette)
    slides: List[SlideDesign] = Field(default_factory=list)


class AgentState(TypedDict, total=False):
    request: GeneratePresentationRequest
    knowledge: list
    plan: PresentationPlan
    draft: PresentationDraft
    design: PresentationDesign
    theme: ThemePalette
    result: FinalPresentation


ALLOWED_LAYOUTS = {"title", "title_and_bullets", "two_columns", "comparison", "timeline", "closing"}


def _knowledge_to_text(chunks: list) -> str:
    if not chunks:
        return "База знаний не загружена."

    prepared = []
    for c in chunks[:6]:
        snippet = (c.content or "").strip()
        if len(snippet) > 700:
            snippet = snippet[:700] + "..."
        prepared.append(f"Источник: {c.source}\nФрагмент: {snippet}")
    return "\n\n".join(prepared)


def _normalize_plan(plan: PresentationPlan, slide_count: int) -> PresentationPlan:
    slides = sorted(plan.slides, key=lambda s: s.slide_number)

    if len(slides) > slide_count:
        slides = slides[:slide_count]

    if len(slides) < slide_count:
        existing = len(slides)
        for i in range(existing + 1, slide_count + 1):
            slides.append(
                SlidePlan(
                    slide_number=i,
                    title=f"Слайд {i}",
                    goal="Раскрыть следующий логический аспект темы",
                    key_points=[
                        "Ключевая идея",
                        "Факт или аргумент",
                        "Практический вывод",
                    ],
                )
            )

    for idx, slide in enumerate(slides, start=1):
        slide.slide_number = idx
        if not slide.key_points:
            slide.key_points = ["Ключевая идея", "Подтверждение", "Вывод"]

    return PresentationPlan(title=plan.title, subtitle=plan.subtitle, slides=slides)


def _normalize_draft(plan: PresentationPlan, draft: PresentationDraft) -> PresentationDraft:
    draft_map = {s.slide_number: s for s in draft.slides}
    normalized: List[SlideContent] = []

    for plan_slide in plan.slides:
        existing = draft_map.get(plan_slide.slide_number)
        if existing:
            bullets = [b.strip() for b in existing.bullets if b.strip()]
            if len(bullets) < 3:
                bullets = bullets + plan_slide.key_points[: max(0, 3 - len(bullets))]
            bullets = bullets[:5]
            normalized.append(
                SlideContent(
                    slide_number=plan_slide.slide_number,
                    title=existing.title or plan_slide.title,
                    bullets=bullets,
                    speaker_notes=existing.speaker_notes or "",
                )
            )
        else:
            normalized.append(
                SlideContent(
                    slide_number=plan_slide.slide_number,
                    title=plan_slide.title,
                    bullets=plan_slide.key_points[:5],
                    speaker_notes=plan_slide.goal,
                )
            )

    return PresentationDraft(title=plan.title, subtitle=plan.subtitle, slides=normalized)


def _normalize_design(draft: PresentationDraft, design: PresentationDesign) -> PresentationDesign:
    design_map = {s.slide_number: s for s in design.slides}
    normalized: List[SlideDesign] = []
    previous_layout = None

    fallback_layouts = ["title_and_bullets", "two_columns", "comparison", "timeline", "closing"]

    for idx, slide in enumerate(draft.slides, start=1):
        existing = design_map.get(slide.slide_number)
        layout = existing.layout if existing and existing.layout in ALLOWED_LAYOUTS else fallback_layouts[min(idx - 1, len(fallback_layouts) - 1)]
        if layout == previous_layout and layout == "title_and_bullets":
            layout = "two_columns"
        previous_layout = layout

        normalized.append(
            SlideDesign(
                slide_number=slide.slide_number,
                layout=layout,
                visual_idea=(existing.visual_idea if existing else "") or "Минималистичная деловая иллюстрация по теме слайда",
                icon_hint=(existing.icon_hint if existing else "") or "presentation",
                emphasis=(existing.emphasis if existing else "balanced") or "balanced",
                background_style=(existing.background_style if existing else "light") or "light",
                image_type=(existing.image_type if existing else "icon") or "icon",
            )
        )

    return PresentationDesign(slides=normalized)


def _merge_art_direction(base: PresentationDesign, art: ArtDirectorOut) -> tuple[PresentationDesign, ThemePalette]:
    base_map = {s.slide_number: s for s in base.slides}
    art_map = {s.slide_number: s for s in art.slides}
    merged: List[SlideDesign] = []

    for slide_number, slide in sorted(base_map.items()):
        styled = art_map.get(slide_number)
        merged.append(
            SlideDesign(
                slide_number=slide_number,
                layout=styled.layout if styled and styled.layout in ALLOWED_LAYOUTS else slide.layout,
                visual_idea=(styled.visual_idea if styled and styled.visual_idea else slide.visual_idea),
                icon_hint=(styled.icon_hint if styled and styled.icon_hint else slide.icon_hint),
                emphasis=(styled.emphasis if styled and styled.emphasis else slide.emphasis),
                background_style=(styled.background_style if styled and styled.background_style else slide.background_style),
                image_type=(styled.image_type if styled and styled.image_type else slide.image_type),
            )
        )

    theme = art.theme or ThemePalette()
    return PresentationDesign(slides=merged), theme


async def retrieve_node(state: AgentState) -> AgentState:
    request = state["request"]
    chunks = knowledge_store.search(request.topic, k=6)
    return {"knowledge": chunks}


async def planner_node(state: AgentState) -> AgentState:
    request = state["request"]
    model = get_model().with_structured_output(PlannerOut)

    prompt = f"""
Ты агент-структуризатор презентаций.

Построй логичную структуру презентации РОВНО ИЗ {request.slide_count} КОНТЕНТНЫХ СЛАЙДОВ.
Не создавай меньше и не создавай больше.
Нумерация строго от 1 до {request.slide_count}.

Тема: {request.topic}
Целевая аудитория: {request.audience}
Цель: {request.purpose}
Тон: {request.tone}

Релевантные знания:
{_knowledge_to_text(state.get('knowledge', []))}

Верни:
- title
- subtitle
- slides

Каждый slide должен содержать:
- slide_number
- title
- goal
- key_points

Каждый слайд должен раскрывать отдельную часть темы.
Структура должна идти от введения и контекста к основной части, выводам и рекомендациям.
"""

    response = await model.ainvoke(prompt)
    plan = PresentationPlan.model_validate(response.model_dump())
    plan = _normalize_plan(plan, request.slide_count)
    return {"plan": plan}


async def writer_node(state: AgentState) -> AgentState:
    request = state["request"]
    plan = state["plan"]
    model = get_model().with_structured_output(WriterOut)

    prompt = f"""
Ты агент-копирайтер презентаций.

Нужно написать РОВНО {len(plan.slides)} слайдов по заданному плану.
Не пропускай ни один слайд.
Для каждого номера слайда верни один объект.

Тема: {request.topic}
Аудитория: {request.audience}
Цель: {request.purpose}
Тон: {request.tone}

План презентации:
{plan.model_dump_json(indent=2)}

Релевантные знания:
{_knowledge_to_text(state.get('knowledge', []))}

Верни список slides.
Для каждого слайда нужны поля:
- slide_number
- title
- bullets
- speaker_notes

Требования:
- bullets: 3-5 коротких, понятных, деловых пунктов
- title должен соответствовать плану
- speaker_notes: 2-4 предложения
- язык: русский
"""

    response = await model.ainvoke(prompt)
    draft = PresentationDraft(title=plan.title, subtitle=plan.subtitle, slides=response.slides)
    draft = _normalize_draft(plan, draft)
    return {"draft": draft}


async def designer_node(state: AgentState) -> AgentState:
    draft = state["draft"]
    model = get_model().with_structured_output(DesignerOut)

    prompt = f"""
Ты агент-дизайнер презентаций.

Нужно подготовить дизайн РОВНО ДЛЯ {len(draft.slides)} слайдов.
Не пропускай слайды.

Черновик презентации:
{draft.model_dump_json(indent=2)}

Верни slides, где у каждого элемента есть:
- slide_number
- layout
- visual_idea
- icon_hint
- emphasis
- background_style
- image_type

Допустимые layout:
- title
- title_and_bullets
- two_columns
- comparison
- timeline
- closing
"""

    response = await model.ainvoke(prompt)
    design = PresentationDesign.model_validate(response.model_dump())
    design = _normalize_design(draft, design)
    return {"design": design}


async def art_director_node(state: AgentState) -> AgentState:
    request = state["request"]
    plan = state["plan"]
    draft = state["draft"]
    design = state["design"]
    model = get_model().with_structured_output(ArtDirectorOut)

    prompt = f"""
Ты арт-директор презентации.

Твоя задача — сделать дек визуально цельным, современным и разнообразным,
не меняя основную логику содержания.

Тема: {request.topic}
Аудитория: {request.audience}
Цель: {request.purpose}
Тон: {request.tone}

План:
{plan.model_dump_json(indent=2)}

Текущий контент:
{draft.model_dump_json(indent=2)}

Текущий дизайн:
{design.model_dump_json(indent=2)}

Верни:
1. theme — объект с полями:
- theme_name
- primary_color
- secondary_color
- accent_color
- background_color
- text_color
- font_family
- style_notes

2. slides — список стилевых решений для каждого слайда с полями:
- slide_number
- layout
- visual_idea
- icon_hint
- emphasis
- background_style
- image_type

Требования:
- не делай подряд больше двух одинаковых layout
- сравнение используй для сравнений и альтернатив
- timeline используй для этапов/процесса
- closing используй для выводов/рекомендаций/CTA
- стиль: современный, деловой, чистый
- цвета возвращай в HEX формате, например #1F4E79
"""

    response = await model.ainvoke(prompt)
    merged_design, theme = _merge_art_direction(design, response)
    return {"design": merged_design, "theme": theme}


async def compose_node(state: AgentState) -> AgentState:
    request = state["request"]
    draft = state["draft"]
    design_map = {item.slide_number: item for item in state["design"].slides}

    slides: List[FinalSlide] = []
    for slide in sorted(draft.slides, key=lambda s: s.slide_number):
        d = design_map.get(slide.slide_number)
        slides.append(
            FinalSlide(
                slide_number=slide.slide_number,
                title=slide.title,
                bullets=slide.bullets,
                speaker_notes=slide.speaker_notes,
                layout=d.layout if d else "title_and_bullets",
                visual_idea=d.visual_idea if d else "",
                icon_hint=d.icon_hint if d else "",
                emphasis=d.emphasis if d else "balanced",
                background_style=d.background_style if d else "light",
                image_type=d.image_type if d else "icon",
            )
        )

    result = FinalPresentation(
        title=draft.title,
        subtitle=draft.subtitle,
        audience=request.audience,
        purpose=request.purpose,
        slides=slides,
        sources=state.get("knowledge", []),
        theme=state.get("theme", ThemePalette()),
    )
    return {"result": result}


def build_graph():
    graph = StateGraph(AgentState)
    graph.add_node("retrieve", retrieve_node)
    graph.add_node("planner", planner_node)
    graph.add_node("writer", writer_node)
    graph.add_node("designer", designer_node)
    graph.add_node("art_director", art_director_node)
    graph.add_node("compose", compose_node)

    graph.set_entry_point("retrieve")
    graph.add_edge("retrieve", "planner")
    graph.add_edge("planner", "writer")
    graph.add_edge("writer", "designer")
    graph.add_edge("designer", "art_director")
    graph.add_edge("art_director", "compose")
    graph.add_edge("compose", END)
    return graph.compile()


presentation_graph = build_graph()


async def run_presentation_graph(request: GeneratePresentationRequest) -> FinalPresentation:
    state = await presentation_graph.ainvoke({"request": request})
    return state["result"]
