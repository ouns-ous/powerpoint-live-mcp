"""Test script for the new PowerPoint MCP utilities:
- KPI Metric Cards
- Interactive Process Flow / Timeline
- Developer Code Editor mockup
- Vector Bar Chart
- Status Badges / Pills
- Slide Image Export (PNG)
- Slide Content Reader / Inspector
"""

import os
import time
from ppt_controller import controller


def test_new_features():
    print(">>> 1. Creating presentation for new utilities demo...")
    controller.new_presentation(width=960, height=540)
    time.sleep(1)

    # ----------------------------------------------------
    # SLIDE 1: KPI Metric Dashboard + Badges
    # ----------------------------------------------------
    print(">>> 2. Building Slide 1 (Metrics & Badges)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(1, "#F8FAFC")

    controller.add_badge(
        slide_index=1,
        text="Q3 2026 BUSINESS METRICS",
        left=80,
        top=40,
        bg_color_hex="#EFF6FF",
        text_color_hex="#1D4ED8",
        border_color_hex="#BFDBFE",
    )

    controller.add_textbox(
        slide_index=1,
        text="Indicateurs Clés de Performance (KPIs)",
        left=80,
        top=75,
        width=800,
        height=45,
        font_size=26,
        bold=True,
        color_hex="#0F172A",
    )

    # 3 Metric Cards side by side
    controller.add_metric_card(
        slide_index=1,
        value="+148%",
        label="Croissance Annuelle",
        subtitle="↑ 28% vs prévisions",
        left=80,
        top=140,
        width=250,
        height=150,
        value_color_hex="#2563EB",
        accent_bar_color_hex="#2563EB",
    )

    controller.add_metric_card(
        slide_index=1,
        value="$4.85M",
        label="Revenu Récurrent (ARR)",
        subtitle="Objectif annuel atteint à 92%",
        left=355,
        top=140,
        width=250,
        height=150,
        value_color_hex="#059669",
        accent_bar_color_hex="#059669",
    )

    controller.add_metric_card(
        slide_index=1,
        value="99.98%",
        label="Disponibilité Service (SLA)",
        subtitle="Zéro incident critique",
        left=630,
        top=140,
        width=250,
        height=150,
        value_color_hex="#7C3AED",
        accent_bar_color_hex="#7C3AED",
    )

    # ----------------------------------------------------
    # SLIDE 2: Roadmap / Process Timeline
    # ----------------------------------------------------
    print(">>> 3. Building Slide 2 (Process Flow Timeline)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(2, "#F8FAFC")

    controller.add_badge(
        slide_index=2,
        text="ROADMAP & PROCESS",
        left=80,
        top=40,
        bg_color_hex="#FEF3C7",
        text_color_hex="#D97706",
    )

    controller.add_textbox(
        slide_index=2,
        text="Plan de Déploiement en 4 Phases",
        left=80,
        top=75,
        width=800,
        height=45,
        font_size=26,
        bold=True,
        color_hex="#0F172A",
    )

    timeline_steps = [
        {"title": "1. Découverte", "desc": "Analyse des besoins et audit de l'infrastructure existante."},
        {"title": "2. Conception", "desc": "Architecture MCP, modélisation des bases SQLite et maquettes."},
        {"title": "3. Intégration", "desc": "Connexion temps réel Windows COM avec PowerPoint et Claude."},
        {"title": "4. Déploiement", "desc": "Tests complets, automatisation et mise en production."},
    ]

    controller.add_timeline(
        slide_index=2,
        steps=timeline_steps,
        left=80,
        top=150,
        width=800,
        height=260,
        accent_color_hex="#2563EB",
    )

    # ----------------------------------------------------
    # SLIDE 3: Developer Code Window
    # ----------------------------------------------------
    print(">>> 4. Building Slide 3 (Code Block Mockup)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(3, "#0F172A")

    controller.add_textbox(
        slide_index=3,
        text="Exemple d'Intégration FastMCP en Python",
        left=80,
        top=45,
        width=800,
        height=45,
        font_size=24,
        bold=True,
        color_hex="#F8FAFC",
    )

    sample_code = """from mcp.server.fastmcp import FastMCP
from ppt_controller import controller

mcp = FastMCP("PowerPoint-Live-Controller")

@mcp.tool()
def ppt_add_card(slide_index: int, title: str, body: str):
    \"\"\"Creates a modern UI card with title and body.\"\"\"
    return controller.add_card(slide_index, title, body, left=80, top=140, width=380, height=220)

if __name__ == "__main__":
    mcp.run()"""

    controller.add_code_block(
        slide_index=3,
        code=sample_code,
        language="python",
        left=80,
        top=105,
        width=800,
        height=360,
        font_size=13,
    )

    # ----------------------------------------------------
    # SLIDE 4: Vector Bar Chart
    # ----------------------------------------------------
    print(">>> 5. Building Slide 4 (Vector Bar Chart)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(4, "#F8FAFC")

    controller.add_textbox(
        slide_index=4,
        text="Performances Comparatives des Modèles IA",
        left=80,
        top=50,
        width=800,
        height=45,
        font_size=26,
        bold=True,
        color_hex="#0F172A",
    )

    chart_data = [
        {"label": "Antigravity 2.0 (High)", "value": 96, "display_value": "96%", "color": "#2563EB"},
        {"label": "Codex Architecture", "value": 91, "display_value": "91%", "color": "#059669"},
        {"label": "Claude 3.7 Sonnet", "value": 94, "display_value": "94%", "color": "#7C3AED"},
        {"label": "Base LLM Baseline", "value": 68, "display_value": "68%", "color": "#94A3B8"},
    ]

    controller.add_bar_chart(
        slide_index=4,
        data=chart_data,
        left=80,
        top=130,
        width=800,
        height=320,
    )

    # ----------------------------------------------------
    # 6. Test Slide Image Export
    # ----------------------------------------------------
    img_out = os.path.join(os.path.dirname(__file__), "slide_1_export.png")
    print(f">>> 6. Exporting Slide 1 as image to {img_out}...")
    export_res = controller.export_slide_image(slide_index=1, output_path=img_out)
    print("Export result:", export_res)

    # ----------------------------------------------------
    # 7. Test Inspect Slide Content
    # ----------------------------------------------------
    print(">>> 7. Inspecting Slide 1 structure...")
    inspect_res = controller.read_slide_content(slide_index=1)
    print(f"Slide 1 has {inspect_res['shapes_count']} shapes detected.")

    controller.goto_slide(1)
    print(">>> All new utility features executed successfully live!")


if __name__ == "__main__":
    test_new_features()
