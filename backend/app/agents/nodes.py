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
)
from app.services.image_service import generate_slide_image


class PlannerOut(BaseModel):
    title: str
    subtitle: str = ""
    slides: List[SlidePlan] = Field(default_factory=list)


class WriterOut(BaseModel):
    slides: List[SlideContent] = Field(default_factory=list)


class DesignerOut(BaseModel):
    slides: List[SlideDesign] = Field(default_factory=list)


class AgentState(TypedDict, total=False):
    request: GeneratePresentationRequest
    knowledge: list
    plan: PresentationPlan
    draft: PresentationDraft
    design: PresentationDesign
    result: FinalPresentation


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

    return PresentationPlan(
        title=plan.title,
        subtitle=plan.subtitle,
        slides=slides,
    )


def _normalize_draft(plan: PresentationPlan, draft: PresentationDraft) -> PresentationDraft:
    draft_map = {s.slide_number: s for s in draft.slides}
    normalized: List[SlideContent] = []

    for plan_slide in plan.slides:
        existing = draft_map.get(plan_slide.slide_number)
        if existing:
            bullets = [b.strip() for b in existing.bullets if b.strip()]
            if len(bullets) < 3:
                bullets = bullets + plan_slide.key_points[: max(0, 3 - len(bullets))]
            bullets = bullets[:6]

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
                    bullets=plan_slide.key_points[:6],
                    speaker_notes=plan_slide.goal,
                )
            )

    return PresentationDraft(
        title=plan.title,
        subtitle=plan.subtitle,
        slides=normalized,
    )


def _normalize_design(draft: PresentationDraft, design: PresentationDesign) -> PresentationDesign:
    allowed_layouts = {
        "title_and_bullets",
        "two_columns",
        "comparison",
        "timeline",
        "metrics",
        "section",
        "image_focus",
        "visual_split",
        "minimal_statement",
        "closing",
    }

    allowed_backgrounds = {"light", "accent", "muted", "white", "dark"}
    allowed_image_types = {"none", "illustration"}

    design_map = {s.slide_number: s for s in design.slides}
    normalized: List[SlideDesign] = []

    for idx, slide in enumerate(draft.slides, start=1):
        existing = design_map.get(slide.slide_number)

        layout = "title_and_bullets"
        visual_idea = "Минималистичная органичная иллюстрация по теме слайда"
        icon_hint = "presentation"
        emphasis = "content"
        background_style = "light"
        image_type = "none"

        if existing:
            if getattr(existing, "layout", None) in allowed_layouts:
                layout = existing.layout
            if getattr(existing, "background_style", None) in allowed_backgrounds:
                background_style = existing.background_style
            if getattr(existing, "image_type", None) in allowed_image_types:
                image_type = existing.image_type

            visual_idea = getattr(existing, "visual_idea", None) or visual_idea
            icon_hint = getattr(existing, "icon_hint", None) or icon_hint
            emphasis = getattr(existing, "emphasis", None) or emphasis

        if idx == len(draft.slides):
            layout = "closing"
            image_type = "none"

        normalized.append(
            SlideDesign(
                slide_number=slide.slide_number,
                layout=layout,
                visual_idea=visual_idea,
                icon_hint=icon_hint,
                emphasis=emphasis,
                background_style=background_style,
                image_type=image_type,
                image_path=None,
            )
        )

    # Убираем длинные серии одинаковых layout подряд
    for i in range(2, len(normalized)):
        if (
            normalized[i].layout == normalized[i - 1].layout
            and normalized[i - 1].layout == normalized[i - 2].layout
        ):
            if normalized[i].layout == "title_and_bullets":
                normalized[i].layout = "two_columns"
            elif normalized[i].layout == "two_columns":
                normalized[i].layout = "visual_split"
            else:
                normalized[i].layout = "title_and_bullets"

    # Если выбран визуальный layout, но картинка не запланирована — включаем illustration
    for slide in normalized:
        if slide.layout in {"image_focus", "visual_split"}:
            slide.image_type = "illustration"

    return PresentationDesign(slides=normalized)


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
- bullets: 4-6 коротких, понятных, деловых пунктов
- title должен соответствовать плану
- speaker_notes: 2-4 предложения
- язык: русский
- делай упор на информативность и связность
"""

    response = await model.ainvoke(prompt)
    draft = PresentationDraft(
        title=plan.title,
        subtitle=plan.subtitle,
        slides=response.slides,
    )
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
- title_and_bullets
- two_columns
- comparison
- timeline
- metrics
- section
- image_focus
- visual_split
- minimal_statement
- closing

background_style выбирай из:
- light
- accent
- muted
- white
- dark

image_type выбирай из:
- none
- illustration

Требования:
- делай упор на понятную структуру и современное минималистичное оформление
- избегай лишнего дробления на множество мелких карточек
- если слайд требует иллюстрации, укажи image_type=illustration
- visual_idea пиши как описание будущей картинки
"""

    response = await model.ainvoke(prompt)
    design = PresentationDesign.model_validate(response.model_dump())
    design = _normalize_design(draft, design)
    return {"design": design}


async def art_director_node(state: AgentState) -> AgentState:
    draft = state["draft"]
    design = state["design"]
    model = get_model().with_structured_output(DesignerOut)

    prompt = f"""
Ты арт-директор презентаций.

Сделай презентацию визуально современной, минималистичной и информационно насыщенной.
Не используй layout quote.
Избегай слишком мелкого дробления на карточки.
Сохраняй мягкие пастельные стили и деловую подачу.

Черновик текста:
{draft.model_dump_json(indent=2)}

Текущий дизайн:
{design.model_dump_json(indent=2)}

Верни slides, где у каждого элемента есть:
- slide_number
- layout
- visual_idea
- icon_hint
- emphasis
- background_style
- image_type

Допустимые layout:
- title_and_bullets
- two_columns
- comparison
- timeline
- metrics
- section
- image_focus
- visual_split
- minimal_statement
- closing

background_style выбирай из:
- light
- accent
- muted
- white
- dark

image_type выбирай из:
- none
- illustration

Правила:
- не делай все слайды одинаковыми
- не используй quote
- если нужен акцент на одной идее + визуале, используй image_focus
- если нужен текст + иллюстрация, используй visual_split
- если нужен сильный смысловой слайд без блоков, используй minimal_statement
- для two_columns подбирай осмысленное смысловое деление
- metrics используй для факторов, категорий и показателей
- timeline используй только для этапов или roadmap
- последний слайд должен быть closing
- visual_idea пиши так, чтобы по нему можно было сгенерировать минималистичную органичную иллюстрацию
- используй мягкие пастельные стили, современную деловую подачу и акцент на читаемости
- не перегружай слайды декоративными элементами
- делай упор на информацию: лучше 1-2 крупных смысловых блока, чем много мелких карточек
- для title_and_bullets и comparison допускается немного больше текста, если это улучшает понятность
- используй в основном стили: light, accent, muted, white
- dark используй редко и только как мягкий акцентный стиль, без настоящего темного фона
"""

    response = await model.ainvoke(prompt)
    design = PresentationDesign.model_validate(response.model_dump())
    design = _normalize_design(draft, design)
    return {"design": design}


async def compose_node(state: AgentState) -> AgentState:
    request = state["request"]
    draft = state["draft"]
    design_map = {item.slide_number: item for item in state["design"].slides}

    slides: List[FinalSlide] = []
    for slide in sorted(draft.slides, key=lambda s: s.slide_number):
        d = design_map.get(slide.slide_number)

        image_path = None
        if (
            d
            and getattr(d, "image_type", "none") == "illustration"
            and d.visual_idea
            and d.layout in {"image_focus", "visual_split"}
            ):
            try:
                image_path = generate_slide_image(
                    title=slide.title,
                    visual_idea=d.visual_idea,
                    slide_number=slide.slide_number,
                )
            except Exception as e:
                print(f"[image generation] failed for slide {slide.slide_number}: {e}")
                image_path = None

        slides.append(
            FinalSlide(
                slide_number=slide.slide_number,
                title=slide.title,
                bullets=slide.bullets,
                speaker_notes=slide.speaker_notes,
                layout=d.layout if d else "title_and_bullets",
                visual_idea=d.visual_idea if d else "",
                icon_hint=d.icon_hint if d else "",
                emphasis=d.emphasis if d else "content",
                background_style=d.background_style if d else "light",
                image_type=d.image_type if d else "none",
                image_path=image_path,
            )
        )

    result = FinalPresentation(
        title=draft.title,
        subtitle=draft.subtitle,
        audience=request.audience,
        purpose=request.purpose,
        slides=slides,
        sources=state.get("knowledge", []),
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