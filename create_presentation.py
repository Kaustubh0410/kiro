"""
Create Smart Food Management PowerPoint Presentation
Generates a visually stunning 10-slide presentation using python-pptx.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# =============================================================================
# Global Constants
# =============================================================================
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Colors
DARK_BG = RGBColor(0x1B, 0x28, 0x38)
CARD_BG = RGBColor(0x2C, 0x3E, 0x50)
GREEN = RGBColor(0x2E, 0xCC, 0x71)
ORANGE = RGBColor(0xF3, 0x9C, 0x12)
RED = RGBColor(0xE7, 0x4C, 0x3C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xEC, 0xF0, 0xF1)
FOOTER_GRAY = RGBColor(0x95, 0xA5, 0xA6)
BLUE = RGBColor(0x34, 0x98, 0xDB)
PURPLE = RGBColor(0x9B, 0x59, 0xB6)

FOOTER_TEXT = "(c) JTP Co., Ltd. All Rights Reserved."


def add_dark_background(slide):
    """Add a full-slide dark background rectangle as the first shape."""
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = DARK_BG
    bg.line.fill.background()
    return bg


def add_footer(slide):
    """Add the JTP copyright footer centered at the bottom."""
    footer = slide.shapes.add_textbox(
        Inches(0), Inches(7.0), SLIDE_WIDTH, Inches(0.4)
    )
    tf = footer.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = FOOTER_TEXT
    run.font.size = Pt(10)
    run.font.color.rgb = FOOTER_GRAY


def set_text_props(text_frame, text, font_size, color, bold=False, alignment=PP_ALIGN.LEFT):
    """Helper to set text properties on a text frame's first paragraph."""
    tf = text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    return tf


def add_paragraph(text_frame, text, font_size, color, bold=False, alignment=PP_ALIGN.LEFT):
    """Add a new paragraph to an existing text frame."""
    p = text_frame.add_paragraph()
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    return p


def add_card(slide, left, top, width, height, fill_color=None, accent_color=None, accent_position="top"):
    """Add a rounded rectangle card with optional colored accent."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    card.fill.solid()
    card.fill.fore_color.rgb = fill_color or CARD_BG
    card.line.fill.background()

    # Add accent line
    if accent_color:
        if accent_position == "top":
            accent = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.08)
            )
        elif accent_position == "left":
            accent = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, left, top, Inches(0.08), height
            )
        accent.fill.solid()
        accent.fill.fore_color.rgb = accent_color
        accent.line.fill.background()

    return card


def create_presentation():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # Use blank layout
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1 - Title
    # =========================================================================
    slide = prs.slides.add_slide(blank_layout)
    add_dark_background(slide)

    # Green accent rectangle on left side
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.4), SLIDE_HEIGHT
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = GREEN
    accent.line.fill.background()

    # Large food emoji
    emoji_box = slide.shapes.add_textbox(Inches(9.5), Inches(1.0), Inches(3), Inches(2))
    set_text_props(emoji_box.text_frame, "\U0001F37D\uFE0F", 72, WHITE, alignment=PP_ALIGN.CENTER)

    # Title
    title_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(8), Inches(1.2))
    set_text_props(title_box.text_frame, "Smart Food Management", 44, WHITE, bold=True, alignment=PP_ALIGN.LEFT)

    # Subtitle
    sub_box = slide.shapes.add_textbox(Inches(1.5), Inches(3.2), Inches(8), Inches(0.8))
    set_text_props(sub_box.text_frame, "Intelligent Platform for Zero Food Waste", 24, LIGHT_GRAY, alignment=PP_ALIGN.LEFT)

    # Stat callout
    stat_box = slide.shapes.add_textbox(Inches(1.5), Inches(4.5), Inches(9), Inches(0.8))
    set_text_props(stat_box.text_frame, "1.3 Billion Tonnes of food wasted annually", 28, ORANGE, bold=True, alignment=PP_ALIGN.LEFT)

    add_footer(slide)

    # =========================================================================
    # SLIDE 2 - The Crisis
    # =========================================================================
    slide = prs.slides.add_slide(blank_layout)
    add_dark_background(slide)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1.0))
    set_text_props(title_box.text_frame, "A Global Food Waste Crisis", 36, WHITE, bold=True, alignment=PP_ALIGN.CENTER)

    # 3 stat cards
    cards_data = [
        ("\U0001F5D1\uFE0F", "1.3B Tonnes", "Food wasted yearly", ORANGE),
        ("\U0001F4B0", "$1 Trillion", "Economic loss annually", ORANGE),
        ("\U0001F622", "828 Million", "People go hungry", RED),
    ]
    card_width = Inches(3.5)
    card_height = Inches(3.5)
    start_left = Inches(1.2)
    spacing = Inches(0.5)

    for i, (emoji, stat, desc, accent_color) in enumerate(cards_data):
        left = start_left + i * (card_width + spacing)
        top = Inches(2.2)

        card = add_card(slide, left, top, card_width, card_height, accent_color=accent_color, accent_position="top")

        # Emoji
        emoji_tb = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(0.4), card_width - Inches(0.6), Inches(1.0))
        set_text_props(emoji_tb.text_frame, emoji, 48, WHITE, alignment=PP_ALIGN.CENTER)

        # Stat number
        stat_tb = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(1.5), card_width - Inches(0.6), Inches(0.8))
        set_text_props(stat_tb.text_frame, stat, 32, accent_color, bold=True, alignment=PP_ALIGN.CENTER)

        # Description
        desc_tb = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(2.4), card_width - Inches(0.6), Inches(0.8))
        set_text_props(desc_tb.text_frame, desc, 18, LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

    add_footer(slide)

    # =========================================================================
    # SLIDES 3, 4, 5 - Problem + Solution pairs
    # =========================================================================
    problem_solution_slides = [
        {
            "header": "The Challenge & Our Answer",
            "problem_emoji": "\U0001F6AB",
            "problem_title": "Forgotten Food",
            "problem_desc": "Households throw away food because they forget what is in their fridge",
            "solution_emoji": "\u2705",
            "solution_title": "Smart Inventory Tracking",
            "solution_desc": "Automatic reminders before food expires. Never waste forgotten items again.",
        },
        {
            "header": "The Challenge & Our Answer",
            "problem_emoji": "\U0001F3EA",
            "problem_title": "Wasted Surplus",
            "problem_desc": "Restaurants & stores discard surplus with no way to redistribute",
            "solution_emoji": "\U0001F91D",
            "solution_title": "Surplus Food Marketplace",
            "solution_desc": "Connects donors with food banks & NGOs instantly. Every extra meal finds a home.",
        },
        {
            "header": "The Challenge & Our Answer",
            "problem_emoji": "\U0001F937",
            "problem_title": "Unused Ingredients",
            "problem_desc": "People do not know how to use leftover ingredients creatively",
            "solution_emoji": "\U0001F468\u200D\U0001F373",
            "solution_title": "AI Recipe Suggestions",
            "solution_desc": "Creates delicious meals from what you already have. Turn leftovers into feasts.",
        },
    ]

    for ps_data in problem_solution_slides:
        slide = prs.slides.add_slide(blank_layout)
        add_dark_background(slide)

        # Header
        header_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12), Inches(0.6))
        set_text_props(header_box.text_frame, ps_data["header"], 18, LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

        # Problem card (left half)
        p_left = Inches(0.8)
        p_top = Inches(1.5)
        p_width = Inches(5.2)
        p_height = Inches(4.8)

        add_card(slide, p_left, p_top, p_width, p_height, accent_color=RED, accent_position="top")

        # Problem emoji
        pe_box = slide.shapes.add_textbox(p_left + Inches(0.5), p_top + Inches(0.5), p_width - Inches(1.0), Inches(1.2))
        set_text_props(pe_box.text_frame, ps_data["problem_emoji"], 48, WHITE, alignment=PP_ALIGN.CENTER)

        # Problem title
        pt_box = slide.shapes.add_textbox(p_left + Inches(0.5), p_top + Inches(1.8), p_width - Inches(1.0), Inches(0.7))
        set_text_props(pt_box.text_frame, ps_data["problem_title"], 24, WHITE, bold=True, alignment=PP_ALIGN.CENTER)

        # Problem description
        pd_box = slide.shapes.add_textbox(p_left + Inches(0.5), p_top + Inches(2.7), p_width - Inches(1.0), Inches(1.5))
        set_text_props(pd_box.text_frame, ps_data["problem_desc"], 16, LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

        # Arrow between
        arrow_box = slide.shapes.add_textbox(Inches(6.1), Inches(3.5), Inches(1.0), Inches(0.8))
        set_text_props(arrow_box.text_frame, "\u2192", 36, GREEN, bold=True, alignment=PP_ALIGN.CENTER)

        # Solution card (right half)
        s_left = Inches(7.2)
        s_top = Inches(1.5)
        s_width = Inches(5.2)
        s_height = Inches(4.8)

        add_card(slide, s_left, s_top, s_width, s_height, accent_color=GREEN, accent_position="top")

        # Solution emoji
        se_box = slide.shapes.add_textbox(s_left + Inches(0.5), s_top + Inches(0.5), s_width - Inches(1.0), Inches(1.2))
        set_text_props(se_box.text_frame, ps_data["solution_emoji"], 48, WHITE, alignment=PP_ALIGN.CENTER)

        # Solution title
        st_box = slide.shapes.add_textbox(s_left + Inches(0.5), s_top + Inches(1.8), s_width - Inches(1.0), Inches(0.7))
        set_text_props(st_box.text_frame, ps_data["solution_title"], 24, GREEN, bold=True, alignment=PP_ALIGN.CENTER)

        # Solution description
        sd_box = slide.shapes.add_textbox(s_left + Inches(0.5), s_top + Inches(2.7), s_width - Inches(1.0), Inches(1.5))
        set_text_props(sd_box.text_frame, ps_data["solution_desc"], 16, LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

        add_footer(slide)

    # =========================================================================
    # SLIDE 6 - Platform Features
    # =========================================================================
    slide = prs.slides.add_slide(blank_layout)
    add_dark_background(slide)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12), Inches(1.0))
    set_text_props(title_box.text_frame, "Six Superpowers in One Platform", 36, WHITE, bold=True, alignment=PP_ALIGN.CENTER)

    # 6 feature cards in 2x3 grid
    features = [
        ("\U0001F514", "Smart Expiry Alerts", "Never let food expire unnoticed", GREEN),
        ("\U0001F9D1\u200D\U0001F373", "AI Recipe Suggestions", "Cook smart with what you have", ORANGE),
        ("\U0001FAC2", "Community Food Sharing", "Share surplus with neighbors", GREEN),
        ("\U0001F4CA", "Waste Analytics Dashboard", "Track and reduce your waste", ORANGE),
        ("\u2764\uFE0F", "Donation Network", "Connect with food banks easily", GREEN),
        ("\U0001F4F1", "Multi-platform Access", "Use on any device, anywhere", ORANGE),
    ]

    card_w = Inches(3.6)
    card_h = Inches(2.4)
    start_x = Inches(0.9)
    start_y = Inches(1.8)
    x_gap = Inches(0.4)
    y_gap = Inches(0.4)

    for i, (emoji, name, desc, accent) in enumerate(features):
        col = i % 3
        row = i // 3
        left = start_x + col * (card_w + x_gap)
        top = start_y + row * (card_h + y_gap)

        add_card(slide, left, top, card_w, card_h, accent_color=accent, accent_position="top")

        # Emoji
        e_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.3), card_w - Inches(0.4), Inches(0.8))
        set_text_props(e_box.text_frame, emoji, 36, WHITE, alignment=PP_ALIGN.CENTER)

        # Feature name
        n_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(1.1), card_w - Inches(0.4), Inches(0.6))
        set_text_props(n_box.text_frame, name, 16, WHITE, bold=True, alignment=PP_ALIGN.CENTER)

        # Description
        d_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(1.7), card_w - Inches(0.4), Inches(0.5))
        set_text_props(d_box.text_frame, desc, 12, LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

    add_footer(slide)

    # =========================================================================
    # SLIDE 7 - How It Works
    # =========================================================================
    slide = prs.slides.add_slide(blank_layout)
    add_dark_background(slide)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12), Inches(1.0))
    set_text_props(title_box.text_frame, "How It Works For You", 36, WHITE, bold=True, alignment=PP_ALIGN.CENTER)

    layers = [
        ("\U0001F4F1", "Works on any device", "Seamless experience on phone, tablet, or computer", BLUE),
        ("\u26A1", "Lightning-fast & beautiful", "Instant responses with an intuitive, clean design", ORANGE),
        ("\U0001F512", "Your data is always safe", "Bank-level security protecting your information", GREEN),
        ("\U0001F9E0", "Smart AI that learns your habits", "Gets better and more personalized the more you use it", PURPLE),
    ]

    bar_width = Inches(11.0)
    bar_height = Inches(1.2)
    bar_start_x = Inches(1.1)
    bar_start_y = Inches(1.8)
    bar_gap = Inches(0.3)

    for i, (emoji, title, desc, accent) in enumerate(layers):
        top = bar_start_y + i * (bar_height + bar_gap)

        # Bar background
        bar = add_card(slide, bar_start_x, top, bar_width, bar_height, accent_color=accent, accent_position="left")

        # Emoji
        e_box = slide.shapes.add_textbox(bar_start_x + Inches(0.4), top + Inches(0.15), Inches(0.8), Inches(0.9))
        set_text_props(e_box.text_frame, emoji, 28, WHITE, alignment=PP_ALIGN.CENTER)

        # Title text
        t_box = slide.shapes.add_textbox(bar_start_x + Inches(1.4), top + Inches(0.1), Inches(9.0), Inches(0.6))
        set_text_props(t_box.text_frame, title, 20, WHITE, bold=True, alignment=PP_ALIGN.LEFT)

        # Description
        d_box = slide.shapes.add_textbox(bar_start_x + Inches(1.4), top + Inches(0.65), Inches(9.0), Inches(0.5))
        set_text_props(d_box.text_frame, desc, 14, LIGHT_GRAY, alignment=PP_ALIGN.LEFT)

    add_footer(slide)

    # =========================================================================
    # SLIDE 8 - Impact
    # =========================================================================
    slide = prs.slides.add_slide(blank_layout)
    add_dark_background(slide)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1.0))
    set_text_props(title_box.text_frame, "The Impact We Will Create", 36, WHITE, bold=True, alignment=PP_ALIGN.CENTER)

    metrics = [
        ("50%", "Less household food waste", GREEN),
        ("10M+", "Meals redistributed", ORANGE),
        ("ZERO", "Food expires forgotten", GREEN),
    ]

    metric_width = Inches(3.5)
    metric_height = Inches(3.5)
    metric_start_x = Inches(1.2)
    metric_gap = Inches(0.5)

    for i, (number, desc, color) in enumerate(metrics):
        left = metric_start_x + i * (metric_width + metric_gap)
        top = Inches(2.2)

        # Card background
        add_card(slide, left, top, metric_width, metric_height, accent_color=color, accent_position="top")

        # Large number
        n_box = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(0.8), metric_width - Inches(0.6), Inches(1.5))
        set_text_props(n_box.text_frame, number, 54, color, bold=True, alignment=PP_ALIGN.CENTER)

        # Description
        d_box = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(2.3), metric_width - Inches(0.6), Inches(0.8))
        set_text_props(d_box.text_frame, desc, 18, WHITE, alignment=PP_ALIGN.CENTER)

    add_footer(slide)

    # =========================================================================
    # SLIDE 9 - Roadmap
    # =========================================================================
    slide = prs.slides.add_slide(blank_layout)
    add_dark_background(slide)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12), Inches(1.0))
    set_text_props(title_box.text_frame, "Our Journey Forward", 36, WHITE, bold=True, alignment=PP_ALIGN.CENTER)

    phases = [
        ("1", "\U0001F4CB Smart Inventory & Alerts", GREEN),
        ("2", "\U0001F9D1\u200D\U0001F373 AI Recipes & Community Sharing", ORANGE),
        ("3", "\U0001F3EA Marketplace & Donation Network", BLUE),
        ("4", "\U0001F30D Global Expansion & Partnerships", PURPLE),
    ]

    phase_width = Inches(2.7)
    phase_height = Inches(3.2)
    phase_start_x = Inches(0.7)
    phase_gap = Inches(0.4)
    phase_top = Inches(2.2)

    # Connecting line
    line_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        phase_start_x + Inches(1.35),
        phase_top + Inches(0.55),
        Inches(11.0),
        Inches(0.06),
    )
    line_shape.fill.solid()
    line_shape.fill.fore_color.rgb = LIGHT_GRAY
    line_shape.line.fill.background()

    for i, (num, title, color) in enumerate(phases):
        left = phase_start_x + i * (phase_width + phase_gap)

        # Number circle
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, left + Inches(0.9), phase_top, Inches(0.9), Inches(0.9)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = color
        circle.line.fill.background()

        # Number text in circle
        tf = circle.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = num
        run.font.size = Pt(24)
        run.font.color.rgb = WHITE
        run.font.bold = True

        # Phase card below
        card = add_card(slide, left, phase_top + Inches(1.2), phase_width, Inches(2.0), accent_color=color, accent_position="top")

        # Phase title
        t_box = slide.shapes.add_textbox(left + Inches(0.2), phase_top + Inches(1.5), phase_width - Inches(0.4), Inches(1.5))
        tf = t_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = title
        run.font.size = Pt(14)
        run.font.color.rgb = WHITE
        run.font.bold = True

    add_footer(slide)

    # =========================================================================
    # SLIDE 10 - Closing
    # =========================================================================
    slide = prs.slides.add_slide(blank_layout)
    add_dark_background(slide)

    # Decorative shapes
    # Green rectangle top-right
    deco1 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(11.5), Inches(0), Inches(1.833), Inches(0.4)
    )
    deco1.fill.solid()
    deco1.fill.fore_color.rgb = GREEN
    deco1.line.fill.background()

    # Orange rectangle bottom-left
    deco2 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.1), Inches(2.0), Inches(0.4)
    )
    deco2.fill.solid()
    deco2.fill.fore_color.rgb = ORANGE
    deco2.line.fill.background()

    # Green circle decorative
    deco3 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(0.5), Inches(0.5), Inches(0.6), Inches(0.6)
    )
    deco3.fill.solid()
    deco3.fill.fore_color.rgb = GREEN
    deco3.line.fill.background()

    # Orange circle decorative
    deco4 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(12.0), Inches(6.0), Inches(0.5), Inches(0.5)
    )
    deco4.fill.solid()
    deco4.fill.fore_color.rgb = ORANGE
    deco4.line.fill.background()

    # Large emoji
    emoji_box = slide.shapes.add_textbox(Inches(5.5), Inches(1.0), Inches(2.5), Inches(1.5))
    set_text_props(emoji_box.text_frame, "\U0001F331", 72, GREEN, alignment=PP_ALIGN.CENTER)

    # Main title
    title_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.8), Inches(10.333), Inches(1.2))
    set_text_props(title_box.text_frame, "Smart Food Management", 44, WHITE, bold=True, alignment=PP_ALIGN.CENTER)

    # Tagline
    tag_box = slide.shapes.add_textbox(Inches(1.5), Inches(4.0), Inches(10.333), Inches(0.8))
    set_text_props(tag_box.text_frame, "Every Meal Matters. Zero Waste. Full Plates.", 28, GREEN, bold=True, alignment=PP_ALIGN.CENTER)

    # JTP branding
    brand_box = slide.shapes.add_textbox(Inches(1.5), Inches(5.2), Inches(10.333), Inches(0.6))
    set_text_props(brand_box.text_frame, "JTP Co., Ltd.", 20, WHITE, alignment=PP_ALIGN.CENTER)

    add_footer(slide)

    # =========================================================================
    # Save
    # =========================================================================
    output_path = "/projects/sandbox/kiro/SmartFoodManagement.pptx"
    prs.save(output_path)
    print(f"Presentation saved to: {output_path}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    create_presentation()
