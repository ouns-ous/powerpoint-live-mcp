"""Test script for PowerPoint Live MCP v0.6.0 utilities.
Tests all 8 new strategic & analytical tools:
1. ppt_add_venn_diagram
2. ppt_add_pyramid_hierarchy
3. ppt_add_matrix_2x2
4. ppt_add_roadmap_milestones
5. ppt_add_user_persona
6. ppt_add_roi_card
7. ppt_add_dashboard_strip
8. ppt_batch_format_deck
"""

import os
import sys
import time

# Ensure powerpoint-mcp is in sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ppt_controller import controller


def run_tests():
    print(">>> [1/9] Creating new presentation...")
    res = controller.new_presentation(width=960.0, height=540.0)
    print("New presentation created:", res)

    # 1. Slide 1: Venn Diagram
    print(">>> [2/9] Testing add_venn_diagram...")
    s1 = controller.add_slide(layout="blank")["slide_index"]
    controller.add_textbox(s1, "STRATEGIC SYNERGY: AI + DESKTOP AUTOMATION", 80, 40, 800, 40, font_size=20, bold=True, color_hex="#0F172A")
    controller.add_venn_diagram(
        slide_index=s1,
        title_1="Generative AI",
        title_2="Native COM APIs",
        overlap_title="Live MCP Control",
        desc_1="Autonomous agents, natural reasoning, and creative layout generation.",
        desc_2="100% fidelity, zero rasterization, instant desktop feedback.",
        overlap_desc="Autonomous desktop presentation copilot.",
        color_1_hex="#3B82F6",
        color_2_hex="#10B981",
        left=180,
        top=130,
        circle_size=320,
    )
    print("Slide 1 Venn diagram added.")

    # 2. Slide 2: Pyramid Hierarchy
    print(">>> [3/9] Testing add_pyramid_hierarchy...")
    s2 = controller.add_slide(layout="blank")["slide_index"]
    controller.add_textbox(s2, "ENTERPRISE AUTOMATION MATURITY PYRAMID", 80, 40, 800, 40, font_size=20, bold=True, color_hex="#0F172A")
    controller.add_pyramid_hierarchy(
        slide_index=s2,
        levels=[
            {"title": "Autonomous Ecosystem", "desc": "Self-healing multi-agent pipelines"},
            {"title": "Interactive Feedback", "desc": "Live COM bridge and instant preview"},
            {"title": "Reusable Templates", "desc": "Standardized executive design systems"},
            {"title": "Scripted Automation", "desc": "Basic one-off slide generators"},
            {"title": "Manual Presentation Design", "desc": "Slow, error-prone manual formatting"},
        ],
        left=120,
        top=110,
        width=720,
        height=380,
    )
    print("Slide 2 Pyramid hierarchy added.")

    # 3. Slide 3: 2x2 Prioritization Matrix
    print(">>> [4/9] Testing add_matrix_2x2...")
    s3 = controller.add_slide(layout="blank")["slide_index"]
    controller.add_textbox(s3, "STRATEGIC FEATURE PRIORITIZATION (2x2 MATRIX)", 80, 40, 800, 40, font_size=20, bold=True, color_hex="#0F172A")
    controller.add_matrix_2x2(
        slide_index=s3,
        x_label_low="Low Implementation Effort",
        x_label_high="High Implementation Effort",
        y_label_low="Low Strategic Value",
        y_label_high="High Strategic Value",
        top_left_title="QUICK WINS (Deploy Immediately)",
        top_left_items=["Progressive slide builds", "Pre-built executive cards", "Venn & Pyramid diagrams"],
        top_right_title="STRATEGIC BETS (Q3 Roadmap)",
        top_right_items=["Full multi-agent presentation orchestra", "Live Excel-to-PowerPoint chart sync"],
        bottom_left_title="TACTICAL FILL-INS (Low Priority)",
        bottom_left_items=["Niche animation transitions", "Legacy format exports (.ppt)"],
        bottom_right_title="TIME SINKS (Avoid / Deprecate)",
        bottom_right_items=["Manual pixel adjustment scripts", "Third-party cloud converters"],
        left=100,
        top=110,
        width=760,
        height=360,
    )
    print("Slide 3 Matrix 2x2 added.")

    # 4. Slide 4: Roadmap Milestones
    print(">>> [5/9] Testing add_roadmap_milestones...")
    s4 = controller.add_slide(layout="blank")["slide_index"]
    controller.add_textbox(s4, "2026 PRODUCT ROADMAP & MILESTONES", 80, 40, 800, 40, font_size=20, bold=True, color_hex="#0F172A")
    controller.add_roadmap_milestones(
        slide_index=s4,
        phases=[
            {
                "period": "Q1 2026",
                "title": "Foundation",
                "deliverables": ["FastMCP Engine", "Windows COM Bridge", "Real-time Live Sync"],
                "status": "completed",
            },
            {
                "period": "Q2 2026",
                "title": "Executive Suite",
                "deliverables": ["KPI Hero Cards", "SWOT & Funnel Visuals", "Sequential Stepper"],
                "status": "completed",
            },
            {
                "period": "Q3 2026",
                "title": "Strategy & Analytics",
                "deliverables": ["2x2 Matrix & Venn Diagrams", "ROI Calculators", "Batch Formatting"],
                "status": "in_progress",
            },
            {
                "period": "Q4 2026",
                "title": "Autonomous Agents",
                "deliverables": ["Autonomous Deck Generation", "Cross-App MCP Orchestration"],
                "status": "planned",
            },
        ],
        left=80,
        top=130,
        width=800,
        height=320,
    )
    print("Slide 4 Roadmap milestones added.")

    # 5. Slide 5: User Persona
    print(">>> [6/9] Testing add_user_persona...")
    s5 = controller.add_slide(layout="blank")["slide_index"]
    controller.add_textbox(s5, "TARGET BUYER & USER PERSONA", 80, 40, 800, 40, font_size=20, bold=True, color_hex="#0F172A")
    controller.add_user_persona(
        slide_index=s5,
        name="Elena Rostova",
        role="VP of Strategic Operations",
        company="Global Enterprise Tech",
        quote="I want my team creating insights, not wasting hours hand-aligning PowerPoint shapes.",
        goals=[
            "Accelerate executive board decks from 3 days to 15 minutes",
            "Maintain immaculate corporate branding across 400+ slides",
            "Give AI autonomous control over native desktop applications",
        ],
        pain_points=[
            "Design friction and inconsistent typography across business units",
            "Tools that output static images rather than native editable slides",
            "Fragmented workflow between Python analysis and executive decks",
        ],
        tech_stack=["PowerPoint Live MCP", "Python", "FastMCP", "Antigravity", "Office 365"],
        left=80,
        top=120,
        width=800,
        height=360,
        accent_color_hex="#2563EB",
    )
    print("Slide 5 User persona added.")

    # 6. Slide 6: ROI Card
    print(">>> [7/9] Testing add_roi_card...")
    s6 = controller.add_slide(layout="blank")["slide_index"]
    controller.add_textbox(s6, "EXECUTIVE BUSINESS CASE & ROI ANALYSIS", 80, 40, 800, 40, font_size=20, bold=True, color_hex="#0F172A")
    controller.add_roi_card(
        slide_index=s6,
        before_cost="$320,000 / yr (Manual Agency & Hours)",
        after_cost="$48,000 / yr (PowerPoint Live MCP)",
        savings_amount="$272,000",
        savings_percentage="85% Annual Savings",
        payback_period="1.4 Months",
        metrics_summary="6,800 executive presentation hours saved per enterprise department annually.",
        left=80,
        top=130,
        width=800,
        height=300,
    )
    print("Slide 6 ROI card added.")

    # 7. Slide 7: Dashboard Strip
    print(">>> [8/9] Testing add_dashboard_strip...")
    s7 = controller.add_slide(layout="blank")["slide_index"]
    controller.add_textbox(s7, "REAL-TIME PLATFORM PERFORMANCE DASHBOARD", 80, 40, 800, 40, font_size=20, bold=True, color_hex="#0F172A")
    controller.add_dashboard_strip(
        slide_index=s7,
        kpis=[
            {"value": "99.99%", "label": "COM Bridge Reliability", "change": "+0.05% vs SLA", "positive": True},
            {"value": "18.4k", "label": "Automated Shapes / Day", "change": "+42% MoM", "positive": True},
            {"value": "85ms", "label": "Avg Command Latency", "change": "-18ms faster", "positive": True},
            {"value": "72", "label": "Active MCP Tools", "change": "v0.6.0 Suite", "positive": True},
        ],
        left=80,
        top=160,
        width=800,
        height=140,
    )
    print("Slide 7 Dashboard strip added.")

    # 8. Slide 8: Batch Format Deck & Footer
    print(">>> [9/9] Testing batch_format_deck...")
    fmt_res = controller.batch_format_deck(
        font_name="Segoe UI",
        title_color_hex="#0F172A",
        footer_text="PowerPoint Live MCP v0.6.0  |  Confidential & Proprietary",
    )
    print("Batch format result:", fmt_res)

    # Save presentation
    os.makedirs(os.path.join(os.path.dirname(__file__), "examples"), exist_ok=True)
    out_pptx = os.path.join(os.path.dirname(__file__), "examples", "v060_demo.pptx")
    controller.save_presentation(out_pptx)
    print(f"Presentation saved to: {out_pptx}")

    # Export slide 1, 3, 5, 7 as sample screenshots
    for s_idx in [1, 3, 5, 6]:
        img_out = os.path.join(os.path.dirname(__file__), "examples", f"slide_{s_idx}_preview.png")
        controller.export_slide_image(slide_index=s_idx, output_path=img_out)
        print(f"Slide {s_idx} preview exported to: {img_out}")

    print("\n>>> ALL V0.6.0 UTILITY TESTS PASSED SUCCESSFULLY! <<<")


if __name__ == "__main__":
    run_tests()
