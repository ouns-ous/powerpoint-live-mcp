"""Demo script to verify newly added tools:
Cover Page, Page Setup, Quote Card, KPI Metric Cards, Process Steps, Highlighting, Endnotes, Markdown Export.
"""

import os
import sys
import time

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

print("=" * 65)
print("[*] TESTING NEW WORD MCP ADVANCED TOOLS (64 Tools Suite)")
print("=" * 65)

from word_controller import controller
from docx_engine import DocxEngine

try:
    print("\n[1/6] Testing Live Cover Page & Page Setup...")
    controller.new_document()
    controller.set_page_setup(orientation="portrait", margins="normal", paper_size="a4")
    cov_res = controller.add_cover_page(
        title="Annual Technology & AI Strategy Report",
        subtitle="Transforming Enterprise Operations with Autonomous Agents",
        author="Enterprise AI Architecture Team",
        organization="GlobalTech Group Ltd.",
        date_str="Fiscal Year 2026/2027",
        theme_color="#1F4E79",
    )
    print("   -> Cover page added:", cov_res["title"])
    time.sleep(1)

    print("\n[2/6] Adding Executive Quote Card...")
    controller.add_heading("1. Strategic Vision", level=1)
    controller.add_quote_card(
        quote="Empowering our workforce with autonomous agentic pairs has accelerated software delivery by 3.4x while maintaining strict compliance.",
        author="Dr. Sarah Jenkins",
        title="Chief Technology Officer",
        border_color="#2563EB",
    )
    time.sleep(0.5)

    print("\n[3/6] Adding KPI Metric Cards Grid...")
    controller.add_heading("2. Key Performance Indicators", level=1)
    kpi_res = controller.add_kpi_cards([
        {"value": "$28.4M", "label": "ARR Generated", "change": "+34% YoY"},
        {"value": "99.98%", "label": "System Uptime", "change": "+0.4%"},
        {"value": "4.8 / 5", "label": "Customer CSAT", "change": "+12% NPS"},
        {"value": "< 12 ms", "label": "Latency P99", "change": "-45% ms"},
    ])
    print(f"   -> Added {kpi_res['cards_count']} KPI cards.")
    time.sleep(0.5)

    print("\n[4/6] Adding Process Roadmap Steps...")
    controller.add_heading("3. Implementation Roadmap", level=1)
    step_res = controller.add_process_steps([
        {"step": 1, "title": "Infrastructure Modernization", "description": "Migrate core microservices to cloud-native orchestrator with zero downtime."},
        {"step": 2, "title": "AI Agent Integration", "description": "Deploy MCP-enabled autonomous coding and review agents across developer teams."},
        {"step": 3, "title": "Enterprise Rollout & Governance", "description": "Complete SOC2 Type II audit and full team enablement sessions."},
    ])
    print(f"   -> Added {step_res['steps_count']} roadmap steps.")
    time.sleep(0.5)

    print("\n[5/6] Testing Keyword Highlighting & Endnote...")
    controller.add_paragraph("Please note that this document is STRICTLY CONFIDENTIAL and subject to NDA.")
    hl_res = controller.highlight_text("STRICTLY CONFIDENTIAL", highlight_color="yellow")
    print(f"   -> Highlighted {hl_res['highlighted_count']} occurrences.")

    controller.add_endnote("Refer to Appendix C for detailed audit methodology and source datasets.", target_text="SOC2 Type II")
    print("   -> Endnote added.")
    time.sleep(0.5)

    print("\n[6/6] Saving & Testing Markdown Export...")
    saved_doc = os.path.abspath("demo_advanced_showcase.docx")
    controller.save_document(saved_doc)
    print("   -> Document saved:", saved_doc)

    export_md = os.path.abspath("demo_extracted.md")
    md_res = DocxEngine.export_markdown(saved_doc, export_md)
    print("   -> Exported to Markdown:", export_md)
    print(f"   -> Markdown preview (first 120 chars): {md_res['markdown'][:120]}...")

    print("\n[+] Exporting final PDF showcase...")
    pdf_res = controller.export_pdf(os.path.abspath("demo_advanced_showcase.pdf"))
    print("   -> PDF exported to:", pdf_res["pdf_path"])

    print("\n[SUCCESS] ALL ADVANCED TOOLS VERIFIED AND PASSED!")

except Exception as e:
    print(f"\n[!] Error during test: {e}")
