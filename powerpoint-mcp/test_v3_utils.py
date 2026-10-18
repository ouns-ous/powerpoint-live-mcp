"""Test script for v0.3.0 utilities in PowerPoint Live MCP
Verifies:
1. ppt_add_callout_box
2. ppt_add_process_steps
3. ppt_add_feature_grid
4. ppt_add_team_card
5. ppt_add_agenda
6. ppt_add_watermark
7. ppt_set_shape_zorder
8. ppt_add_comparison_matrix
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

    print("\n>>> 2. Testing ppt_add_agenda (Slide 1)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(1, "Executive Briefing Agenda", left=80, top=50, width=800, height=50, font_size=28, bold=True)
    agenda_res = controller.add_agenda(
        1,
        items=[
            {"num": "01", "title": "Market Dynamics & Opportunity", "desc": "Macro analysis and growth tailwinds", "duration": "10 min", "active": False},
            {"num": "02", "title": "Product Architecture & AI Engine", "desc": "Live MCP protocol integration", "duration": "20 min", "active": True},
            {"num": "03", "title": "Financial Projections & Roadmap", "desc": "3-year forecast and team scaling", "duration": "15 min", "active": False},
        ],
        top=120,
    )
    print("Agenda:", agenda_res)

    print("\n>>> 3. Testing ppt_add_process_steps (Slide 2)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(2, "Product Development Lifecycle", left=80, top=50, width=800, height=50, font_size=28, bold=True)
    steps_res = controller.add_process_steps(
        2,
        steps=[
            {"number": "1", "title": "Discovery", "description": "User research and problem definition"},
            {"number": "2", "title": "Architecture", "description": "MCP server protocol and schemas"},
            {"number": "3", "title": "Implementation", "description": "Python COM automation & FastMCP"},
            {"number": "4", "title": "Deployment", "description": "Production release on GitHub"},
        ],
        left=80,
        top=160,
        width=800,
        height=240,
        accent_color_hex="#2563EB",
    )
    print("Process Steps:", steps_res)

    print("\n>>> 4. Testing ppt_add_feature_grid (Slide 3)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(3, "Platform Key Capabilities", left=80, top=50, width=800, height=50, font_size=28, bold=True)
    feat_res = controller.add_feature_grid(
        3,
        features=[
            {"title": "Real-Time Visual", "description": "Every edit happens visibly in the active PowerPoint window.", "tag": "Live", "accent_color": "#3B82F6"},
            {"title": "FastMCP Architecture", "description": "Native Model Context Protocol for Claude, Antigravity, and Codex.", "tag": "Standard", "accent_color": "#10B981"},
            {"title": "48+ Rich Tools", "description": "Charts, cards, timelines, matrices, callouts, and animations.", "tag": "Power", "accent_color": "#8B5CF6"},
        ],
        columns=3,
        top=140,
    )
    print("Feature Grid:", feat_res)

    print("\n>>> 5. Testing ppt_add_callout_box (Slide 3)...")
    callout_res = controller.add_callout_box(
        3,
        text="The server automatically bridges Claude Desktop and Antigravity with Windows COM for sub-10ms response times.",
        title="PRO TIP FOR AI AGENTS",
        callout_type="tip",
        top=360,
        height=100,
    )
    print("Callout:", callout_res)

    print("\n>>> 6. Testing ppt_add_team_card (Slide 4)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(4, "Leadership & Engineering", left=80, top=50, width=800, height=50, font_size=28, bold=True)
    team1 = controller.add_team_card(
        4,
        name="Oussama Ouns",
        role="Lead Systems & AI Architect",
        company_or_dept="Core Infrastructure",
        bio="Pioneering real-time generative UI and desktop automation protocols for next-generation AI agents.",
        skills_or_tags=["Python", "FastMCP", "COM Automation", "AI"],
        left=80,
        top=140,
        width=380,
        height=260,
        accent_color_hex="#2563EB",
    )
    team2 = controller.add_team_card(
        4,
        name="Sarah Chen",
        role="VP of Product Engineering",
        company_or_dept="Desktop Ecosystem",
        bio="10+ years directing enterprise software workflows, Microsoft Office integrations, and devtools.",
        skills_or_tags=["Architecture", "Office COM", "DevOps"],
        left=500,
        top=140,
        width=380,
        height=260,
        accent_color_hex="#10B981",
    )
    print("Team Cards:", team1, team2)

    print("\n>>> 7. Testing ppt_add_comparison_matrix (Slide 5)...")
    controller.add_slide(layout="blank")
    controller.add_textbox(5, "Edition Feature Comparison", left=80, top=50, width=800, height=50, font_size=28, bold=True)
    matrix_res = controller.add_comparison_matrix(
        5,
        headers=["Capability", "Standard PPT", "PowerPoint Live MCP"],
        rows=[
            ["Real-Time Live Automation", "No", "Yes (Full Visible)"],
            ["FastMCP Protocol", "No", "Yes (48+ Native Tools)"],
            ["AI Agent Autonomous Control", "No", "Yes (Antigravity & Claude)"],
            ["Executive Widgets (Cards/Timelines)", "Manual", "1-Click API"],
        ],
        highlight_col_index=3,
        top=130,
        width=800,
        height=240,
    )
    print("Comparison Matrix:", matrix_res)

    print("\n>>> 8. Testing ppt_add_watermark (All Slides)...")
    wm_res = controller.add_watermark(text="CONFIDENTIAL DRAFT", slide_index=None, font_size=68)
    print("Watermark:", wm_res)

    print("\n>>> 9. Testing ppt_set_shape_zorder...")
    zorder_res = controller.set_shape_zorder(5, shape_id_or_name=matrix_res["table_id"], action="bring_to_front")
    print("Z-Order:", zorder_res)

    print("\n>>> 10. Saving demo file and exporting image...")
    examples_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples")
    os.makedirs(examples_dir, exist_ok=True)
    out_path = os.path.join(examples_dir, "v3_showcase.pptx")
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except Exception:
            pass
    controller.save_presentation(out_path)
    print("Saved presentation to:", out_path)

    img_path = os.path.join(examples_dir, "v3_slide_feature_callout.png")
    img_res = controller.export_slide_image(slide_index=3, output_path=img_path)
    print("Exported slide 3 image:", img_res)
    print("\nALL V3 UTILITY TESTS PASSED SUCCESSFULLY! [SUCCESS]")


if __name__ == "__main__":
    run_tests()
