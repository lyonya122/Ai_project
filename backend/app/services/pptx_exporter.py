from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from app.core.config import settings
from app.schemas.presentation import FinalPresentation, FinalSlide, ThemePalette


SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def _hex_to_rgb(hex_value: str, fallback: str = "#1F4E79") -> RGBColor:
    value = (hex_value or fallback).strip().lstrip("#")
    if len(value) != 6:
        value = fallback.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def _apply_background(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_title(slide, text: str, theme: ThemePalette, color: RGBColor | None = None) -> None:
    tx = slide.shapes.add_textbox(Inches(0.7), Inches(0.45), Inches(11.5), Inches(0.7))
    tf = tx.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.name = theme.font_family
    p.font.color.rgb = color or _hex_to_rgb(theme.text_color)


def _add_subtitle(slide, text: str, theme: ThemePalette, top: float = 1.3) -> None:
    tx = slide.shapes.add_textbox(Inches(0.8), Inches(top), Inches(11.0), Inches(0.5))
    p = tx.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(14)
    p.font.name = theme.font_family
    p.font.color.rgb = _hex_to_rgb(theme.text_color)


def _add_bullets_box(slide, bullets: list[str], theme: ThemePalette, left: float, top: float, width: float, height: float, text_color: RGBColor | None = None) -> None:
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.clear()

    items = bullets[:5] if bullets else ["Ключевая мысль по теме слайда"]
    for idx, bullet in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = bullet
        p.level = 0
        p.space_after = Pt(8)
        p.font.size = Pt(20 if len(items) <= 4 else 18)
        p.font.name = theme.font_family
        p.font.color.rgb = text_color or _hex_to_rgb(theme.text_color)


def _add_visual_card(slide, slide_data: FinalSlide, theme: ThemePalette, left: float, top: float, width: float, height: float) -> None:
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _hex_to_rgb(theme.secondary_color)
    shape.line.color.rgb = _hex_to_rgb(theme.accent_color)
    shape.line.width = Pt(1.2)

    tf = shape.text_frame
    tf.clear()

    p0 = tf.paragraphs[0]
    p0.text = slide_data.icon_hint or "visual"
    p0.font.bold = True
    p0.font.size = Pt(16)
    p0.font.name = theme.font_family
    p0.font.color.rgb = _hex_to_rgb(theme.primary_color)

    p1 = tf.add_paragraph()
    p1.text = slide_data.visual_idea or "Подходящее визуальное решение для усиления идеи слайда"
    p1.font.size = Pt(14)
    p1.font.name = theme.font_family
    p1.font.color.rgb = _hex_to_rgb(theme.text_color)


def _render_title_and_bullets(slide, slide_data: FinalSlide, theme: ThemePalette) -> None:
    _apply_background(slide, _hex_to_rgb(theme.background_color))
    accent = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0), Inches(0), Inches(0.35), SLIDE_H)
    accent.fill.solid()
    accent.fill.fore_color.rgb = _hex_to_rgb(theme.primary_color)
    accent.line.fill.background()

    _add_title(slide, slide_data.title, theme)
    _add_bullets_box(slide, slide_data.bullets, theme, 1.0, 1.45, 6.8, 4.9)
    _add_visual_card(slide, slide_data, theme, 8.2, 1.5, 4.2, 4.5)


def _render_two_columns(slide, slide_data: FinalSlide, theme: ThemePalette) -> None:
    _apply_background(slide, _hex_to_rgb(theme.background_color))
    _add_title(slide, slide_data.title, theme)

    left_card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(4.9))
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = RGBColor(255, 255, 255)
    left_card.line.color.rgb = _hex_to_rgb(theme.secondary_color)

    right_card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.9))
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = _hex_to_rgb(theme.secondary_color)
    right_card.line.color.rgb = _hex_to_rgb(theme.accent_color)

    _add_bullets_box(slide, slide_data.bullets, theme, 1.1, 1.8, 5.0, 4.2)
    _add_visual_card(slide, slide_data, theme, 7.1, 1.9, 5.1, 3.9)


def _render_comparison(slide, slide_data: FinalSlide, theme: ThemePalette) -> None:
    _apply_background(slide, _hex_to_rgb(theme.background_color))
    _add_title(slide, slide_data.title, theme)

    bullets = slide_data.bullets[:4]
    left_items = bullets[:2] or ["Сторона А", "Ключевой аргумент"]
    right_items = bullets[2:4] or ["Сторона Б", "Ключевой аргумент"]

    for left, label, items in [(0.8, "Вариант A", left_items), (6.8, "Вариант B", right_items)]:
        card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(left), Inches(1.7), Inches(5.1), Inches(4.7))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = _hex_to_rgb(theme.accent_color)

        hdr = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(left), Inches(1.7), Inches(5.1), Inches(0.6))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = _hex_to_rgb(theme.primary_color)
        hdr.line.fill.background()
        tf = hdr.text_frame
        p = tf.paragraphs[0]
        p.text = label
        p.font.bold = True
        p.font.size = Pt(16)
        p.font.name = theme.font_family
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        _add_bullets_box(slide, items, theme, left + 0.25, 2.55, 4.6, 3.4)


def _render_timeline(slide, slide_data: FinalSlide, theme: ThemePalette) -> None:
    _apply_background(slide, _hex_to_rgb(theme.background_color))
    _add_title(slide, slide_data.title, theme)

    line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(1.1), Inches(3.45), Inches(10.8), Inches(0.08))
    line.fill.solid()
    line.fill.fore_color.rgb = _hex_to_rgb(theme.accent_color)
    line.line.fill.background()

    bullets = slide_data.bullets[:4] or ["Шаг 1", "Шаг 2", "Шаг 3", "Шаг 4"]
    positions = [1.2, 4.0, 6.8, 9.6]
    for idx, item in enumerate(bullets):
        x = positions[idx]
        circle = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x), Inches(3.0), Inches(0.55), Inches(0.55))
        circle.fill.solid()
        circle.fill.fore_color.rgb = _hex_to_rgb(theme.primary_color)
        circle.line.fill.background()

        num_box = slide.shapes.add_textbox(Inches(x), Inches(3.06), Inches(0.55), Inches(0.3))
        p = num_box.text_frame.paragraphs[0]
        p.text = str(idx + 1)
        p.alignment = PP_ALIGN.CENTER
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.name = theme.font_family
        p.font.color.rgb = RGBColor(255, 255, 255)

        text_box = slide.shapes.add_textbox(Inches(x - 0.25), Inches(3.75), Inches(1.4), Inches(1.3))
        p2 = text_box.text_frame.paragraphs[0]
        p2.text = item
        p2.alignment = PP_ALIGN.CENTER
        p2.font.size = Pt(14)
        p2.font.name = theme.font_family
        p2.font.color.rgb = _hex_to_rgb(theme.text_color)


def _render_closing(slide, slide_data: FinalSlide, theme: ThemePalette) -> None:
    _apply_background(slide, _hex_to_rgb(theme.primary_color))
    _add_title(slide, slide_data.title, theme, RGBColor(255, 255, 255))
    _add_subtitle(slide, slide_data.visual_idea or "Ключевой вывод и рекомендуемое действие", theme, top=1.35)

    quote = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(2.0), Inches(11.2), Inches(2.7))
    quote.fill.solid()
    quote.fill.fore_color.rgb = _hex_to_rgb(theme.accent_color)
    quote.line.fill.background()

    tf = quote.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = slide_data.bullets[0] if slide_data.bullets else "Главный итог презентации"
    p.font.bold = True
    p.font.size = Pt(24)
    p.font.name = theme.font_family
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    for bullet in slide_data.bullets[1:4]:
        p2 = tf.add_paragraph()
        p2.text = bullet
        p2.font.size = Pt(18)
        p2.font.name = theme.font_family
        p2.font.color.rgb = RGBColor(255, 255, 255)
        p2.alignment = PP_ALIGN.CENTER


def _render_slide(prs: Presentation, slide_data: FinalSlide, theme: ThemePalette) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    layout = slide_data.layout or "title_and_bullets"

    if layout == "two_columns":
        _render_two_columns(slide, slide_data, theme)
    elif layout == "comparison":
        _render_comparison(slide, slide_data, theme)
    elif layout == "timeline":
        _render_timeline(slide, slide_data, theme)
    elif layout == "closing":
        _render_closing(slide, slide_data, theme)
    else:
        _render_title_and_bullets(slide, slide_data, theme)

    if slide_data.speaker_notes:
        slide.notes_slide.notes_text_frame.text = slide_data.speaker_notes


def export_presentation(data: FinalPresentation) -> Path:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    theme = data.theme

    cover = prs.slides.add_slide(prs.slide_layouts[6])
    _apply_background(cover, _hex_to_rgb(theme.primary_color))

    band = cover.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.1), Inches(11.5), Inches(4.8))
    band.fill.solid()
    band.fill.fore_color.rgb = _hex_to_rgb(theme.secondary_color)
    band.line.fill.background()

    title_box = cover.shapes.add_textbox(Inches(1.25), Inches(1.7), Inches(10.6), Inches(1.6))
    p = title_box.text_frame.paragraphs[0]
    p.text = data.title
    p.font.bold = True
    p.font.size = Pt(28)
    p.font.name = theme.font_family
    p.font.color.rgb = _hex_to_rgb(theme.primary_color)

    subtitle = data.subtitle or f"Аудитория: {data.audience} • Цель: {data.purpose}"
    sub_box = cover.shapes.add_textbox(Inches(1.3), Inches(3.25), Inches(10.2), Inches(0.8))
    p2 = sub_box.text_frame.paragraphs[0]
    p2.text = subtitle
    p2.font.size = Pt(16)
    p2.font.name = theme.font_family
    p2.font.color.rgb = _hex_to_rgb(theme.text_color)

    theme_box = cover.shapes.add_textbox(Inches(1.3), Inches(4.55), Inches(5.0), Inches(0.5))
    p3 = theme_box.text_frame.paragraphs[0]
    p3.text = f"Style: {theme.theme_name}"
    p3.font.size = Pt(13)
    p3.font.name = theme.font_family
    p3.font.color.rgb = _hex_to_rgb(theme.accent_color)

    for slide in data.slides:
        _render_slide(prs, slide, theme)

    out_dir = Path(settings.generated_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    filename = f"presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
    path = out_dir / filename
    prs.save(str(path))
    return path
