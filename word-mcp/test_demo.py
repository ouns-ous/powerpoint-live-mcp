"""Interactive Demo & Verification for Word MCP Server.
Tests both headless DOCX/OOXML engine and live Windows Word COM desktop automation.
"""

import os
import sys
import time

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

print("=" * 60)
print("[*] TESTING WORD MCP SERVER (Live COM & Headless OOXML)")
print("=" * 60)

# 1. Test Headless Markdown to DOCX
print("\n[1/3] Testing Headless Markdown to DOCX conversion...")
from docx_engine import DocxEngine

md_sample = """# Executive Project Proposal

Welcome to the automated document generated via **Word MCP**.

## Key Objectives
- Accelerate document drafting by **85%**
- Enforce unified typography and enterprise branding
- Support live collaboration and change management

> **Note**: This document was generated programmatically from Markdown.

## Milestone Schedule

| Phase | Milestone | Owner | Status |
| :--- | :--- | :--- | :--- |
| 1 | Requirements & Discovery | Engineering | Completed |
| 2 | Live Automation Engine | AI Team | In Progress |
| 3 | Final Executive Review | Leadership | Scheduled |

```python
# Sample code snippet
def generate_report():
    print("Document successfully created!")
```
"""

test_md_docx = os.path.abspath("demo_markdown_output.docx")
res_md = DocxEngine.create_from_markdown(
    markdown_text=md_sample,
    output_path=test_md_docx,
    title="Executive Project Proposal",
    author="Antigravity AI",
)
print("   -> Generated:", res_md["output_path"])

# Test Audit
audit_res = DocxEngine.audit_document(test_md_docx)
print(f"   -> Audit: {audit_res['paragraphs_count']} paragraphs, {audit_res['tables_count']} tables, status={audit_res['status']}")

# 2. Test Template Filling
print("\n[2/3] Testing Template Filling...")
template_docx = os.path.abspath("demo_template.docx")
DocxEngine.create_from_markdown(
    "# Client Contract\n\nClient Name: {{client_name}}\nEffective Date: {{date}}\nTotal Value: {{amount}}",
    output_path=template_docx,
)

filled_docx = os.path.abspath("demo_filled.docx")
fill_res = DocxEngine.fill_template(
    template_path=template_docx,
    data={"client_name": "ACME Corporation Ltd", "date": "October 2026", "amount": "$150,000 USD"},
    output_path=filled_docx,
)
print(f"   -> Template filled ({fill_res['replacements_made']} replacements): {filled_docx}")

# 3. Test Live Windows COM Automation
print("\n[3/3] Testing Live Microsoft Word Desktop Automation...")
from word_controller import controller

try:
    print("   -> Launching / connecting to Word Desktop...")
    doc_res = controller.new_document(title="Live Word Automation Demo")
    print("   -> Created live document:", doc_res["name"])
    time.sleep(1)

    print("   -> Adding Heading 1...")
    controller.add_heading("1. Executive Summary", level=1, color="#1F4E79")
    time.sleep(0.5)

    print("   -> Adding styled paragraph...")
    controller.add_paragraph(
        "This document was created and formatted live in Microsoft Word using the Word MCP server. "
        "Watch your text, tables, callouts, and styles being injected live on screen!",
        font_name="Calibri",
        font_size=11,
        color="#334155",
        space_after=8,
    )
    time.sleep(0.5)

    print("   -> Adding Callout Box...")
    controller.add_callout_box(
        title="Pro Tip",
        text="Every single tool call is recorded in Word's Undo stack. You can press Ctrl+Z in Word anytime!",
        callout_type="tip",
    )
    time.sleep(0.5)

    print("   -> Adding Bulleted List...")
    controller.add_bullet_list([
        "Live Word COM Automation on Windows",
        "Over 40 high-performance tools for document manipulation",
        "Native export to PDF and live tracked changes",
        "Clean integration with Claude, Cursor, and Antigravity",
    ], list_type="bullet")
    time.sleep(0.5)

    print("   -> Adding Data Table...")
    controller.add_table(
        headers=["Feature", "Word MCP", "Legacy Tools"],
        rows=[
            ["Live Visible Desktop", "Yes (Real-time)", "No (File only)"],
            ["Per-Action Undo (Ctrl+Z)", "Supported", "Not Supported"],
            ["Callout / Alert Cards", "Supported", "Manual"],
            ["PDF Native Export", "1-Click Native", "Requires conversion"],
        ],
        header_bg="#1E293B",
        alternate_shading=True,
    )
    time.sleep(0.5)

    print("   -> Adding Header & Footer with dynamic page numbers...")
    controller.set_header_footer(
        header_text="CONFIDENTIAL & PROPRIETARY",
        footer_text="Word MCP Enterprise Suite",
        include_page_number=True,
    )
    time.sleep(0.5)

    print("   -> Adding Editorial Comment...")
    controller.add_comment("Verified ready for client delivery.", target_text="Executive Summary", author="Lead Reviewer")
    time.sleep(0.5)

    print("   -> Saving Document...")
    saved_live_doc = os.path.abspath("demo_live_showcase.docx")
    controller.save_document(saved_live_doc)
    print("   -> Saved to:", saved_live_doc)

    print("   -> Exporting to PDF...")
    pdf_res = controller.export_pdf()
    print("   -> PDF Exported to:", pdf_res["pdf_path"])

    print("\n✅ ALL TESTS & LIVE DEMOS COMPLETED SUCCESSFULLY!")

except Exception as e:
    print(f"\n⚠️ Live COM error: {e}")
