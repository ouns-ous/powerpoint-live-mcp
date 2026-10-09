"""Live Demo for PowerPoint MCP Controller
Demonstrates live visual creation of modern slides in Microsoft PowerPoint.
"""

import time
from ppt_controller import controller


def run_demo():
    print(">>> 1. Creating new presentation in PowerPoint...")
    res = controller.new_presentation(width=960, height=540)
    print("Result:", res)
    time.sleep(1)

    # ----------------------------------------------------
    # SLIDE 1: Cover Slide
    # ----------------------------------------------------
    print(">>> 2. Building Slide 1 (Cover)...")
    s1 = controller.add_slide(layout="blank")
    # Modern dark blue / indigo background
    controller.set_slide_background(1, "#0F172A", gradient_color_hex="#1E293B")
    time.sleep(0.5)

    # Decorative top badge
    controller.add_shape(
        slide_index=1,
        shape_type="rounded_rectangle",
        left=80,
        top=80,
        width=170,
        height=32,
        fill_color_hex="#3B82F6",
        text="✨ LIVE MCP POWER",
        text_color_hex="#FFFFFF",
        font_size=11,
        bold=True,
    )
    time.sleep(0.5)

    # Main Title
    controller.add_textbox(
        slide_index=1,
        text="AI-Powered PowerPoint Automation",
        left=80,
        top=130,
        width=800,
        height=90,
        font_size=38,
        bold=True,
        color_hex="#F8FAFC",
        font_name="Segoe UI",
    )
    time.sleep(0.5)

    # Subtitle
    controller.add_textbox(
        slide_index=1,
        text="Direct live control from Antigravity, Codex & LLM Agents.\nWatch your slides get designed and animated in real-time!",
        left=80,
        top=230,
        width=780,
        height=70,
        font_size=18,
        bold=False,
        color_hex="#94A3B8",
        font_name="Segoe UI",
    )
    time.sleep(0.5)

    # Add subtle animation to title
    controller.add_animation(slide_index=1, shape_index_or_id=2, effect="fade")
    controller.set_slide_transition(slide_index=1, transition="fade", speed="medium")
    time.sleep(1)

    # ----------------------------------------------------
    # SLIDE 2: 3 Modern Feature Cards
    # ----------------------------------------------------
    print(">>> 3. Building Slide 2 (Feature Cards)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(2, "#F8FAFC")
    time.sleep(0.5)

    # Header
    controller.add_textbox(
        slide_index=2,
        text="Key Capabilities & Features",
        left=80,
        top=50,
        width=800,
        height=50,
        font_size=28,
        bold=True,
        color_hex="#0F172A",
    )
    controller.add_textbox(
        slide_index=2,
        text="Everything an AI agent needs to produce professional presentations.",
        left=80,
        top=100,
        width=800,
        height=30,
        font_size=15,
        color_hex="#64748B",
    )
    time.sleep(0.5)

    # 3 Cards side by side
    card_width = 245
    card_height = 280
    gap = 25
    start_x = 80
    start_y = 150

    cards_data = [
        {
            "title": "👀 Live Visuals",
            "body": "PowerPoint stays open and active on your desktop.\nYou see every text box, shape, color, and slide appear live as the agent works.",
            "accent": "#2563EB",
            "bg": "#FFFFFF",
        },
        {
            "title": "🎨 Rich Styling",
            "body": "Custom color palettes, dark mode gradients, cards, auto-shapes, bullet lists, and tables with full font controls.",
            "accent": "#7C3AED",
            "bg": "#FFFFFF",
        },
        {
            "title": "🎬 Motion & Control",
            "body": "Add entrance animations (fade, zoom, bounce), slide transitions, and even launch full-screen presentations remotely.",
            "accent": "#059669",
            "bg": "#FFFFFF",
        },
    ]

    for i, c in enumerate(cards_data):
        cx = start_x + i * (card_width + gap)
        controller.add_card(
            slide_index=2,
            title=c["title"],
            body=c["body"],
            left=cx,
            top=start_y,
            width=card_width,
            height=card_height,
            bg_color_hex=c["bg"],
            border_color_hex="#E2E8F0",
            title_color_hex="#0F172A",
            body_color_hex="#475569",
            accent_bar_color_hex=c["accent"],
        )
        time.sleep(0.4)

    controller.set_slide_transition(slide_index=2, transition="push", speed="fast")
    time.sleep(1)

    # ----------------------------------------------------
    # SLIDE 3: Comparison Table
    # ----------------------------------------------------
    print(">>> 4. Building Slide 3 (Interactive Data Table)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(3, "#F8FAFC")
    time.sleep(0.5)

    controller.add_textbox(
        slide_index=3,
        text="Traditional File Generation vs Live MCP",
        left=80,
        top=50,
        width=800,
        height=50,
        font_size=28,
        bold=True,
        color_hex="#0F172A",
    )

    table_data = [
        ["Feature", "Static Library (python-pptx)", "PowerPoint MCP Live Controller"],
        ["Real-time UI", "❌ Blind (file saved only)", "✅ Live visible window on screen"],
        ["Animation Support", "❌ Very limited / complex", "✅ Native COM animations (fade, zoom...)"],
        ["Slide Transitions", "❌ Not rendered live", "✅ Native PowerPoint transitions"],
        ["Remote Slideshow", "❌ Impossible", "✅ Live start, next, previous, exit"],
        ["Visual Verification", "❌ Need to reopen file", "✅ Instant real-time feedback"],
    ]

    controller.add_table(
        slide_index=3,
        rows=6,
        cols=3,
        data=table_data,
        left=80,
        top=130,
        width=800,
        height=320,
        header_bg_hex="#1E293B",
        header_text_hex="#F8FAFC",
        row_bg_hex="#F1F5F9",
        alt_row_bg_hex="#FFFFFF",
        text_color_hex="#334155",
    )
    time.sleep(0.5)

    controller.set_slide_notes(
        slide_index=3,
        notes_text="Demonstrates how MCP enables live bidirectional interaction with Microsoft Office."
    )

    # Return to slide 1
    controller.goto_slide(1)
    print(">>> Demo completed successfully! PowerPoint window is active on your screen.")


if __name__ == "__main__":
    run_demo()
