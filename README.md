# PowerPoint Live Automation MCP Server 🎬🚀

A powerful **Model Context Protocol (MCP)** server that enables AI coding assistants (**Antigravity**, **Codex**, **Claude Desktop**, **Cursor**) to control **Microsoft PowerPoint** on Windows in real-time with live visible updates and animations.

Built with **Windows COM Automation (`pywin32`)** and **FastMCP**.

---

## 🌟 Highlights

- **Live Real-Time Automation**: Connects directly to Microsoft PowerPoint desktop. Watch your slides being drawn, shaped, colored, and animated live on your desktop!
- **Modern Slide Components**: Sleek 16:9 widescreen slides, dark mode gradients, modern UI cards, KPI metrics, timelines/roadmaps, data tables, and developer code windows.
- **Animations & Transitions**: Entrance animations (fade, zoom, bounce, fly-in), slide transitions, and full presenter controls.
- **Image Export**: Export any slide directly to 1080p high-resolution PNG or JPG image snapshots.
- **Bi-directional Slide Inspection**: Read back text, shapes, tables, and notes from any existing slide.

---

## 📁 Repository Structure

```
├── powerpoint-mcp/
│   ├── server.py              # FastMCP Server (23 Live Tools)
│   ├── ppt_controller.py      # Windows COM Automation Controller
│   ├── test_demo.py           # Live 3-Slide Interactive Demo
│   ├── test_new_utils.py      # Demo for KPI cards, Timelines, Charts, Code editor
│   ├── pyproject.toml         # Dependencies (mcp<2, pywin32, pydantic)
│   └── README.md              # Detailed Documentation
└── README.md
```

---

## 🚀 Quick Setup

### 1. In Antigravity / Claude Desktop

Add this configuration to your `mcp_config.json` (or `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "C:\\path\\to\\powerpoint-mcp",
        "python",
        "server.py"
      ]
    }
  }
}
```

Or using direct Python executable:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": [
        "C:\\path\\to\\powerpoint-mcp\\server.py"
      ]
    }
  }
}
```

---

## 🛠️ Available MCP Tools

| Tool | Description |
| :--- | :--- |
| `ppt_status` | Check PowerPoint running state, active presentations, slide count |
| `ppt_new_presentation` | Create a new 16:9 widescreen presentation and focus window |
| `ppt_open` | Open existing `.pptx` presentation |
| `ppt_save` | Save presentation as `.pptx` or export to `.pdf` |
| `ppt_close` | Close active presentation |
| `ppt_add_slide` | Add new slide and automatically jump view to it |
| `ppt_goto_slide` | Navigate editor view to a specific slide |
| `ppt_set_background` | Set solid color or smooth 2-color gradient background |
| `ppt_add_textbox` | Add formatted rich text (size, font, color, alignment) |
| `ppt_add_card` | Add a modern UI container card with title, body, and accent stripe |
| `ppt_add_shape` | Add auto-shapes (rectangles, rounded rectangles, stars, arrows, etc.) |
| `ppt_add_bullet_list` | Add a clean bulleted list |
| `ppt_add_table` | Add styled data tables with custom headers and alternating rows |
| `ppt_add_image` | Insert images from local disk |
| `ppt_add_animation` | Add entrance animations (fade, zoom, bounce, fly-in...) |
| `ppt_set_transition` | Add visual slide transitions (fade, push, wipe...) |
| `ppt_set_speaker_notes` | Add speaker presenter notes |
| `ppt_add_metric_card` | Add KPI metric widget (large stat, label, and trend/subtext) |
| `ppt_add_timeline` | Add sleek process flow / roadmap / timeline with numbered steps |
| `ppt_add_code_block` | Add dark developer code editor mockup with macOS buttons |
| `ppt_add_bar_chart` | Add clean horizontal vector bar chart with custom colors |
| `ppt_add_badge` | Add rounded tag / pill badge component |
| `ppt_export_slide_image` | Export slide to high-res PNG or JPG image (1080p) |
| `ppt_duplicate_slide` | Duplicate an existing slide |
| `ppt_move_slide` | Reorder / move slides |
| `ppt_read_slide_content` | Inspect and read all text, tables, and shapes from slide |
| `ppt_search_and_replace_text` | Find and replace text or template variables (`{{VAR}}`) across slides |
| `ppt_add_quote_card` | Add testimonial/quote card with quote mark and author details |
| `ppt_add_pros_cons` | Add 2-column side-by-side Pros (green) vs Cons (red) cards |
| `ppt_add_pricing_table` | Add multi-tier SaaS pricing comparison cards with popular badge |
| `ppt_add_donut_chart` | Add circular percentage progress metric widget (donut) |
| `ppt_delete_shape` | Delete a specific shape by ID or name |
| `ppt_clear_slide` | Clear all shapes from a slide to reset to blank |
| `ppt_add_footer` | Add professional footer bar with confidentiality and slide number |
| `ppt_run_slideshow` | Launch full-screen presentation mode |
| `ppt_slideshow_next` | Advance slideshow to next step/slide |
| `ppt_slideshow_previous` | Go back to previous slideshow slide |
| `ppt_slideshow_exit` | Exit slideshow mode back to editor |

---

## 🧪 Live Demos

Run the general demo:
```bash
cd powerpoint-mcp
uv run python test_demo.py
```

Run the new utilities demo (KPIs, timeline, code editor, bar charts, image export):
```bash
cd powerpoint-mcp
uv run python test_new_utils.py
```

---

## 📄 License
MIT License
