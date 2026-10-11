# Microsoft Word Live Automation & OOXML MCP Server 📄🚀

A comprehensive, state-of-the-art **Model Context Protocol (MCP)** server that enables AI coding assistants (**Antigravity**, **Claude Desktop**, **Cursor**, **Codex**) to create, edit, format, and audit **Microsoft Word** documents with live real-time desktop automation on Windows, plus fast headless OOXML / python-docx capabilities.

---

## 🌟 Highlights

- **Dual-Engine Architecture**:
  - **Live Windows COM Automation**: Controls Microsoft Word desktop in real-time. Watch your documents, paragraphs, headings, tables, callouts, and watermarks rendered live on your screen.
  - **Headless OOXML / python-docx Engine**: Fast offline processing for Markdown-to-DOCX conversion, batch template filling, document comparison, and integrity audits.
- **Per-Action Undo (`Ctrl+Z`)**: Every tool call is isolated within Word's native `UndoRecord` stack. Pressing `Ctrl+Z` in Word cleanly reverts the AI's last action.
- **Modern Executive Styling**: Professional palettes (`#1F4E79` Navy, `#2563EB` Royal Blue, `#1E293B` Dark Slate tables), typography rules, callout cards (info, tip, warning, danger), zebra-striped tables, quote cards, KPI metric grids, and process roadmaps.
- **Tracked Changes & Review**: Enable/disable track changes, list revisions, accept/reject revisions, insert threaded comments, and resolve comments.
- **Direct Native PDF Export**: 1-click export to high-resolution print-ready PDF using Word's built-in `ExportAsFixedFormat`.
- **64 High-Performance Tools**: Comprehensive toolset covering lifecycle, content, typography, tables, page layout, multi-column setups, headers/footers with dynamic page numbers, watermarks, formulas, bookmarks, content controls, charts, and audits.

---

## 📁 Repository Structure

```
├── word-mcp/
│   ├── server.py              # FastMCP Server (64 Tools)
│   ├── word_controller.py     # Windows COM Automation Controller
│   ├── docx_engine.py         # Headless OOXML & python-docx Engine
│   ├── test_demo.py           # Verification & Live Interactive Demo
│   ├── test_new_utils.py      # Demo for KPI cards, Roadmaps, Quotes, Cover Page
│   ├── SKILL.md               # Companion Agent Skill & Workflows
│   ├── pyproject.toml         # Dependencies (mcp<2, pywin32, python-docx, etc.)
│   └── README.md              # Documentation
```

---

## 🚀 Quick Setup

### 1. In Antigravity / Claude Desktop

Add this configuration to your `mcp_config.json` (or `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "word": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "C:\\path\\to\\word-mcp",
        "python",
        "server.py"
      ]
    }
  }
}
```

Or using the virtual environment python directly:

```json
{
  "mcpServers": {
    "word": {
      "command": "C:\\path\\to\\word-mcp\\.venv\\Scripts\\python.exe",
      "args": [
        "C:\\path\\to\\word-mcp\\server.py"
      ]
    }
  }
}
```

---

## 🛠️ Complete Tool Reference (64 Tools)

### 1. Document Lifecycle & Application
| Tool | Description |
| :--- | :--- |
| `word_status` | Check Word running status, active document, page/word/char count |
| `word_new_document` | Create a new blank document and bring Word to foreground |
| `word_open` | Open existing `.docx` file in Word desktop |
| `word_save` | Save active document in place or to a new destination |
| `word_export_pdf` | 1-click native PDF export using Word's engine |
| `word_close` | Close active document with optional save |
| `word_focus_window` | Maximize Word window and bring to foreground |

### 2. Executive Components & Visual Cards
| Tool | Description |
| :--- | :--- |
| `word_add_cover_page` | Insert executive cover / title page (Title, Subtitle, Author, Date, Org) |
| `word_add_quote_card` | Insert pull-quote or testimonial card with accent border |
| `word_add_kpi_cards` | Insert horizontal grid of stat / KPI metric cards with trends |
| `word_add_process_steps` | Insert sequential step-by-step roadmap or timeline flow |
| `word_add_callout_box` | Insert alert callout box (info, warning, tip, note, danger) |

### 3. Content & Typography
| Tool | Description |
| :--- | :--- |
| `word_add_heading` | Insert styled heading (H1 to H6) with custom font, color, spacing |
| `word_add_paragraph` | Add styled paragraph (font family, size, bold, italic, color, line spacing) |
| `word_add_bullet_list` | Add bulleted or numbered list |
| `word_add_image` | Insert image with scaling in cm, caption, alignment |
| `word_add_equation` | Insert mathematical formula using UnicodeMath / LaTeX syntax |
| `word_add_footnote` | Insert footnote anchored to target text |
| `word_add_endnote` | Insert endnote at document end |
| `word_list_endnotes` | List all endnotes in document |
| `word_add_hyperlink` | Insert clickable web link |
| `word_add_bookmark` | Insert named bookmark for cross-referencing |
| `word_search_and_replace` | Find and replace text throughout document |
| `word_highlight_text` | Highlight all occurrences of target keyword (yellow, green, blue, pink) |
| `word_format_text` | Format all occurrences of target text (bold, italic, color, size) |

### 4. Tables & Data Graphics
| Tool | Description |
| :--- | :--- |
| `word_add_table` | Insert styled table with dark headers, subtle borders, zebra rows |
| `word_modify_table` | Modify table (set cell text/color, add/delete rows & columns, autofit) |
| `word_add_chart` | Insert native Microsoft Office chart (column, bar, line, pie) |

### 5. Layout, Page Setup & Fields
| Tool | Description |
| :--- | :--- |
| `word_set_page_setup` | Configure margins (normal, narrow, wide), orientation, paper size (A4, Letter) |
| `word_set_columns` | Set multi-column layout (1, 2, or 3 columns) with custom spacing |
| `word_add_page_break` | Insert page break |
| `word_add_section_break` | Insert section break (next page or continuous) |
| `word_set_header_footer` | Set running header & footer with dynamic page numbering ("Page X of Y") |
| `word_add_watermark` | Add diagonal or horizontal translucent watermark (DRAFT, CONFIDENTIAL) |
| `word_remove_watermark` | Remove watermarks from active document |
| `word_add_toc` | Insert Table of Contents field (H1–H3) |
| `word_update_fields` | Refresh all dynamic fields (TOC, page numbers, cross-references) |
| `word_add_bates_numbering` | Add legal Bates stamping numbering to footer across all pages |

### 6. Review, Comments & Track Changes
| Tool | Description |
| :--- | :--- |
| `word_add_comment` | Insert editorial review comment anchored to text |
| `word_list_comments` | List all comments, authors, timestamps, scope |
| `word_reply_comment` | Add threaded reply to existing comment |
| `word_resolve_comment` | Mark comment as resolved or open |
| `word_delete_comment` | Delete comment by index |
| `word_toggle_track_changes` | Toggle or set Word's Track Changes recording |
| `word_list_revisions` | List tracked insertions, deletions, formatting changes |
| `word_accept_revisions` | Accept all or author-filtered tracked changes |
| `word_reject_revisions` | Reject all or author-filtered tracked changes |

### 7. Document Protection & Legal Controls
| Tool | Description |
| :--- | :--- |
| `word_set_protection` | Enforce document protection (revisions-only, comments-only, readonly) |
| `word_remove_protection` | Remove protection with password |
| `word_redact_text` | Redact sensitive phrases with black highlight and replacement mark |
| `word_add_content_control` | Insert form field widget (rich text, plain text, date picker) |
| `word_list_content_controls` | List all form fields and tags in document |

### 8. Inspection, Safety & Merging
| Tool | Description |
| :--- | :--- |
| `word_read_document_text` | Extract paragraphs and styles from active document |
| `word_read_outline` | Extract heading outline hierarchy |
| `word_undo` | Undo last N actions via Word undo stack |
| `word_take_screenshot` | Capture visual snapshot of the Word window |
| `word_merge_documents` | Merge multiple Word files into a single master document |

### 9. Headless OOXML & python-docx Engine
| Tool | Description |
| :--- | :--- |
| `word_create_from_markdown` | Convert full Markdown document to professional .docx |
| `word_export_markdown` | Convert any existing .docx document back into clean Markdown text |
| `word_fill_template` | Replace `{{variables}}` across body, tables, headers |
| `word_compare_documents` | Compare two documents and generate redline diff `.docx` |
| `word_generate_change_summary` | Generate email-ready text changelog comparing two files |
| `word_audit_document` | Scan structure, check paragraph health and integrity |
| `word_scrub_metadata` | Sanitize author, company, comments, revision traces |

---

## 🧪 Live Demos

Run general demo:
```bash
cd word-mcp
uv run python test_demo.py
```

Run advanced tools demo (Cover page, KPI cards, Roadmap, Quotes, Markdown export):
```bash
cd word-mcp
uv run python test_new_utils.py
```

---

## 📄 License
MIT License
