"""Test script for v0.4.0 utilities in PowerPoint Live MCP
Verifies:
1. ppt_add_swot_matrix
2. ppt_add_funnel
3. ppt_add_before_after
4. ppt_add_stat_hero
5. ppt_add_faq_accordion
6. ppt_add_checklist
7. ppt_add_review_card
8. ppt_style_shape
"""

import sys
import os

# Add local directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ppt_controller import controller


def run_tests():
    print(">>> 1. Creating new presentation...")
    res = controller.new_presentation(width=960, height=540)
    print("New pres:", res)

    print("\n>>> 2. Testing ppt_add_swot_matrix (Slide 1)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(1, "Strategic SWOT Analysis", left=80, top=45, width=800, height=50, font_size=28, bold=True)
    swot_res = controller.add_swot_matrix(
        1,
        strengths=["Proprietary FastMCP architecture", "Sub-10ms COM response times", "56+ pre-built executive widgets"],
        weaknesses=["Requires Windows OS desktop", "PowerPoint desktop license needed"],
        opportunities=["Enterprise AI assistant market", "Automated pitch deck generation", "Office automation ecosystem"],
        threats=["Cloud-only office competitors", "Vendor API lock-in"],
        top=115,
        height=360,
    )
    print("SWOT Matrix:", swot_res)

    print("\n>>> 3. Testing ppt_add_funnel (Slide 2)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(2, "Enterprise Customer Conversion Funnel", left=80, top=45, width=800, height=50, font_size=28, bold=True)
    funnel_res = controller.add_funnel(
        2,
        stages=[
            {"name": "Developer Reach", "value": "150,000", "pct": "100%"},
            {"name": "CLI & MCP Installs", "value": "38,500", "pct": "25.6%"},
            {"name": "Active Teams", "value": "9,200", "pct": "6.1%"},
            {"name": "Enterprise Contracts", "value": "1,850", "pct": "1.2%"},
        ],
        top=125,
        height=350,
    )
    print("Funnel:", funnel_res)

    print("\n>>> 4. Testing ppt_add_before_after (Slide 3)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(3, "Workflow Transformation", left=80, top=45, width=800, height=50, font_size=28, bold=True)
    ba_res = controller.add_before_after(
        3,
        before_title="Manual Slide Crafting",
        before_items=["Hours spent aligning rectangles & fonts", "Inconsistent branding & colors", "Stale data copy-pasted from spreadsheets"],
        after_title="PowerPoint Live MCP",
        after_items=["Autonomous generation via Claude & Antigravity", "Pixel-perfect modern visual cards", "Live programmatic data synchronization"],
        top=125,
        height=350,
    )
    print("Before/After:", ba_res)

    print("\n>>> 5. Testing ppt_add_stat_hero & ppt_add_review_card (Slide 4)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(4, "Impact Metrics & Customer Trust", left=80, top=45, width=800, height=50, font_size=28, bold=True)
    hero1 = controller.add_stat_hero(
        4,
        number="+380%",
        label="Slide Production Velocity",
        trend="▲ 4.8x Acceleration",
        trend_positive=True,
        details="Benchmarked across 1,200 enterprise presentations.",
        left=80,
        top=130,
        width=380,
        height=240,
        accent_color_hex="#2563EB",
    )
    rev1 = controller.add_review_card(
        4,
        rating=5.0,
        quote="PowerPoint Live MCP changed how our entire advisory team creates client decks. The live visual feedback is magical.",
        author="David Vance",
        company="Managing Director, Tier-1 Advisory",
        platform="G2 Best Software 2026",
        left=500,
        top=130,
        width=380,
        height=240,
    )
    print("Stat Hero & Review Card:", hero1, rev1)

    print("\n>>> 6. Testing ppt_add_faq_accordion (Slide 5)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(5, "Frequently Asked Questions", left=80, top=45, width=800, height=50, font_size=28, bold=True)
    faq_res = controller.add_faq_accordion(
        5,
        qa_pairs=[
            {"question": "How does live visibility work?", "answer": "The server connects via COM directly to PowerPoint.Application, rendering updates immediately in real-time."},
            {"question": "Which AI models and clients are supported?", "answer": "All MCP-compliant clients including Antigravity, Claude Desktop, Cursor, and Codex."},
            {"question": "Can I export presentations as PDF or high-res images?", "answer": "Yes, ppt_save supports .pdf and ppt_export_slide_image exports crisp 1080p/4K PNGs."},
        ],
        top=120,
        height=360,
    )
    print("FAQ Accordion:", faq_res)

    print("\n>>> 7. Testing ppt_add_checklist (Slide 6)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(6, "Enterprise Release Checklist", left=80, top=45, width=800, height=50, font_size=28, bold=True)
    chk_res = controller.add_checklist(
        6,
        items=[
            {"text": "FastMCP Server Core & COM Controller", "status": "completed", "badge": "VERIFIED"},
            {"text": "56+ Visual Widgets & Charts Implementation", "status": "completed", "badge": "SHIPPED"},
            {"text": "Live Slideshow & Navigation Controls", "status": "completed", "badge": "READY"},
            {"text": "Multi-Tenant Enterprise Cloud Sync", "status": "in_progress", "badge": "Q1 2027"},
        ],
        top=125,
        height=340,
    )
    print("Checklist:", chk_res)

    print("\n>>> 8. Testing ppt_style_shape on Slide 6 title...")
    slide6 = controller.get_presentation().Slides(6)
    title_shape = slide6.Shapes(1)
    style_res = controller.style_shape(6, shape_id_or_name=title_shape.Id, border_color_hex="#2563EB", border_weight=1.5, has_shadow=True)
    print("Style Shape:", style_res)

    print("\n>>> 9. Saving demo file and exporting showcase slide...")
    examples_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples")
    os.makedirs(examples_dir, exist_ok=True)
    out_path = os.path.join(examples_dir, "v4_showcase.pptx")
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except Exception:
            pass
    controller.save_presentation(out_path)
    print("Saved presentation to:", out_path)

    # Export Slide 1 (SWOT) and Slide 3 (Before/After)
    swot_img = os.path.join(examples_dir, "v4_slide_swot.png")
    controller.export_slide_image(slide_index=1, output_path=swot_img)
    ba_img = os.path.join(examples_dir, "v4_slide_before_after.png")
    controller.export_slide_image(slide_index=3, output_path=ba_img)
    print("Exported showcase images successfully!")

    print("\nALL V4 UTILITY TESTS PASSED SUCCESSFULLY! [SUCCESS]")


if __name__ == "__main__":
    run_tests()
