---
name: word-master
description: "Expert guidance and best practices for creating, editing, and formatting Microsoft Word (.docx) documents using the Word MCP Server. Triggers: creating documents, drafting reports, converting markdown to word, live Word desktop editing, tracked changes, contracts review, tables, callouts, headers/footers, and PDF export."
---

# Microsoft Word Automation & Document Engineering (Word MCP)

This skill guides the AI assistant on how to create, format, edit, and review Microsoft Word documents using the `word-mcp` tools.

---

## 🎯 Architecture: Two Powerful Engines

The Word MCP provides two complementary engines:

1. **Live Windows Word Automation (Desktop COM)**:
   - Controls Microsoft Word in real-time right before the user's eyes.
   - Every single operation is isolated inside Word's native `UndoRecord` (`Ctrl+Z`).
   - Use when the user wants to see their document created live, inspect comments, or view formatting in Word desktop.

2. **Headless OOXML / python-docx Engine**:
   - Ultra-fast programmatic document processing without launching Word GUI.
   - High-fidelity Markdown-to-DOCX conversion with executive typography.
   - DOCX-to-Markdown extraction (`word_export_markdown`).
   - Batch template filling (`{{placeholders}}`), diffing two documents, and structural audit.

---

## 🎨 Enterprise Document Style Guide

When creating or formatting Word documents, adhere to these professional design guidelines:

### Typography & Spacing
- **Body Font**: `Calibri` or `Aptos` (11 pt, `#262626` dark charcoal, 1.15 line spacing, 6 pt space after).
- **Headings**:
  - **Document Title**: 24–30 pt, Bold, `#1F4E79` (Navy Deep), Centered or Left, 18 pt space after.
  - **Heading 1**: 18–20 pt, Bold, `#1F4E79`, 12 pt space before, 6 pt space after.
  - **Heading 2**: 14–16 pt, Bold, `#2563EB` (Royal Blue), 10 pt space before, 4 pt space after.
  - **Heading 3**: 12–13 pt, Semi-bold, `#374151` (Slate), 8 pt space before, 4 pt space after.

### Color Palette
- **Primary Navy**: `#1F4E79` (Headings, titles, primary accents)
- **Accent Blue**: `#2563EB` (Subheadings, hyperlinks, key callouts)
- **Table Header**: `#1E293B` (Dark slate with `#FFFFFF` bold text)
- **Zebra Row Shading**: `#F8FAFC` (Subtle off-white / light slate)
- **Callout Backgrounds**:
  - Info: `#EFF6FF` (border `#2563EB`)
  - Warning: `#FFFBEB` (border `#D97706`)
  - Tip: `#ECFDF5` (border `#059669`)
  - Danger: `#FEF2F2` (border `#DC2626`)

---

## 📋 Common Workflows

### 1. Executive Report with Cover Page & Cards
```python
# 1. Start clean live document
word_new_document()

# 2. Add Cover Page
word_add_cover_page(
    title="Annual Strategy & Transformation",
    subtitle="Accelerating Value Delivery with AI Pairs",
    author="Architecture Office",
    organization="Acme Corporation",
    theme_color="#1F4E79"
)

# 3. Add Sections, Quote & KPI Cards
word_add_heading("1. Strategic Overview", level=1)
word_add_quote_card(
    quote="Agility and precision defined our quarterly milestones.",
    author="Jane Doe",
    title="VP Engineering"
)

word_add_kpi_cards([
    {"value": "$32.4M", "label": "ARR Generated", "change": "+28% YoY"},
    {"value": "99.98%", "label": "System SLA", "change": "+0.4%"},
    {"value": "4.9 / 5", "label": "CSAT Score", "change": "+15%"},
])

# 4. Add Roadmap Process Steps
word_add_process_steps([
    {"step": 1, "title": "Discovery", "description": "System architecture audit"},
    {"step": 2, "title": "Execution", "description": "Continuous deployment phase"},
    {"step": 3, "title": "Scale", "description": "Enterprise rollout"},
])

# 5. Configure Header and Footer
word_set_header_footer(
    header_text="CONFIDENTIAL - STRATEGIC REPORT",
    footer_text="Acme Corp Enterprise Suite",
    include_page_number=True,
    different_first_page=True
)

# 6. Save & Export to PDF
word_save("C:\\path\\to\\report.docx")
word_export_pdf("C:\\path\\to\\report.pdf")
```

### 2. Converting Markdown to Word & Back
```python
# Markdown -> DOCX
word_create_from_markdown(
    markdown_content="# Proposal\n\n**Bold text**, lists, and tables...",
    output_path="C:\\path\\to\\output.docx",
    title="Enterprise Proposal"
)

# DOCX -> Markdown
word_export_markdown(
    file_path="C:\\path\\to\\output.docx",
    output_path="C:\\path\\to\\output.md"
)
```

### 3. Contract Review & Tracked Changes
```python
# 1. Open document
word_open("C:\\contracts\\MSA_v1.docx")

# 2. Turn on Track Changes
word_toggle_track_changes(enabled=True)

# 3. Find and replace clauses with tracked revisions
word_search_and_replace("Net 30 days", "Net 60 days")

# 4. Add editorial review comment
word_add_comment(
    text="Updated per Amendment #4 signed on Oct 10.",
    target_text="Net 60 days",
    author="Legal Counsel"
)

# 5. Protect document to allow only revisions
word_set_protection(protection_type="revisions", password="secure")

# 6. Save revised contract
word_save("C:\\contracts\\MSA_v2_redline.docx")
```

---

## 🛠️ Complete Tool Reference (64 Tools)

### Document Lifecycle
- `word_status`: Check Word app status, active doc, page/word/table counts.
- `word_new_document`: Create new document and focus window.
- `word_open`: Open existing `.docx` file in Word desktop.
- `word_save`: Save active document.
- `word_export_pdf`: 1-click export to high-quality PDF via Word native engine.
- `word_close`: Close document with optional save.
- `word_focus_window`: Bring Word window to front and maximize.

### Executive Cards & Visuals
- `word_add_cover_page`: Full executive cover page with metadata.
- `word_add_quote_card`: Pull-quote testimonial card.
- `word_add_kpi_cards`: Grid of KPI stat metric cards.
- `word_add_process_steps`: Numbered roadmap/timeline steps.
- `word_add_callout_box`: Styled alert card (info, warning, tip, note, danger).

### Content & Typography
- `word_add_heading`: Heading 1 to 6 with font sizing, color, spacing.
- `word_add_paragraph`: Full styling (font, bold, italic, color, alignment, spacing).
- `word_add_bullet_list`: Bulleted or numbered lists.
- `word_add_image`: Insert image with sizing in cm, caption, alignment.
- `word_add_equation`: UnicodeMath / LaTeX mathematical formula.
- `word_add_footnote`: Insert footnote.
- `word_add_endnote`: Insert endnote at document end.
- `word_list_endnotes`: List all endnotes.
- `word_add_hyperlink`: Clickable web link.
- `word_add_bookmark`: Named bookmark.
- `word_search_and_replace`: Document-wide find & replace.
- `word_highlight_text`: Keyword highlighting across document.
- `word_format_text`: Format occurrences of specific text.

### Tables & Charts
- `word_add_table`: Styled data table with dark headers, borders, zebra rows.
- `word_modify_table`: Set cell text/color, add/delete rows & columns, autofit.
- `word_add_chart`: Insert native Office chart (column, bar, line, pie).

### Layout & Page Setup
- `word_set_page_setup`: Configure margins, orientation, paper size.
- `word_set_columns`: Multi-column layout (1, 2, or 3 columns).
- `word_add_page_break`: Insert page break.
- `word_add_section_break`: Insert section break.
- `word_set_header_footer`: Running header & footer with dynamic page numbers.
- `word_add_watermark`: Translucent watermark (DRAFT, CONFIDENTIAL).
- `word_remove_watermark`: Remove existing watermarks.
- `word_add_toc`: Insert Table of Contents field (H1–H3).
- `word_update_fields`: Update TOC, page numbers, and cross-references.
- `word_add_bates_numbering`: Legal Bates stamping in footer.

### Review, Comments & Track Changes
- `word_add_comment`: Insert comment anchored to text.
- `word_list_comments`: List all comments, authors, timestamps.
- `word_reply_comment`: Add threaded reply to comment.
- `word_resolve_comment`: Mark comment as resolved.
- `word_delete_comment`: Delete comment.
- `word_toggle_track_changes`: Enable/disable track changes recording.
- `word_list_revisions`: List tracked insertions and deletions.
- `word_accept_revisions`: Accept all or author-filtered changes.
- `word_reject_revisions`: Reject all or author-filtered changes.

### Protection, Forms & Security
- `word_set_protection`: Enforce document protection (revisions, comments, readonly).
- `word_remove_protection`: Unlock document.
- `word_redact_text`: Black-out sensitive phrases.
- `word_add_content_control`: Insert form field (rich text, date picker).
- `word_list_content_controls`: List all form fields.

### Inspection, Safety & Merging
- `word_read_document_text`: Extract paragraphs and styles.
- `word_read_outline`: Extract heading hierarchy.
- `word_undo`: Undo last N actions via Word undo stack.
- `word_take_screenshot`: Capture visual snapshot of the Word window.
- `word_merge_documents`: Merge multiple Word documents into one.

### Headless & OOXML Engine
- `word_create_from_markdown`: Convert Markdown to professional .docx document.
- `word_export_markdown`: Convert .docx back to Markdown.
- `word_fill_template`: Replace `{{variables}}` across body, tables, headers.
- `word_compare_documents`: Generate redline comparison `.docx`.
- `word_generate_change_summary`: Generate human-readable changelog `.txt`.
- `word_audit_document`: Scan structure, check paragraph health and integrity.
- `word_scrub_metadata`: Remove author, company, comments, and revision traces.
