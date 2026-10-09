# PowerPoint Live Automation MCP Suite 🎬🚀

A complete Model Context Protocol (MCP) suite enabling AI coding assistants (**Antigravity**, **Codex**, **Claude**, **Cursor**) to control **Microsoft PowerPoint** on Windows in real-time with visible live updates, animations, and product database integrations.

---

## 🌟 Highlights

- **Live PowerPoint Automation**: Connects directly to Microsoft PowerPoint desktop via Windows COM (`pywin32`). Watch slides being drawn, shaped, colored, and animated live on your screen.
- **23 Comprehensive MCP Tools**: Create presentations, add slides, styled text boxes, modern UI cards, custom shapes, data tables, bullet points, animations (fade, zoom, bounce, fly-in), transitions, and remote slideshow control.
- **Turnkey Database Integrations**: Includes ready-to-use SQLite & JSON product catalogs with automatic PowerPoint generation scripts.

---

## 📁 Repository Structure

```
├── powerpoint-mcp/                # The Core PowerPoint MCP Server
│   ├── server.py                  # FastMCP Server (23 Live Tools)
│   ├── ppt_controller.py          # Windows COM Automation Controller
│   ├── test_demo.py               # Live Interactive Demo
│   ├── pyproject.toml             # Dependencies (mcp<2, pywin32, pydantic)
│   └── README.md                  # Server Documentation & Tool Specs
│
├── database_1_produit/            # Single Product Database
│   ├── products.db                # SQLite Database (Sony WH-1000XM5)
│   ├── products.json              # Clean JSON Export
│   ├── schema.sql                 # SQL Schema & Seed Data
│   └── init_db.py                 # DB Initializer & Queries
│
├── database_4_produits/           # 4-Products Catalog Database
│   ├── products.db                # SQLite Database (4 Diverse Products)
│   ├── products.json              # Clean JSON Export
│   ├── schema.sql                 # SQL Schema & Seed Data
│   ├── init_db.py                 # DB Initializer & Queries
│   └── generate_slides_from_db.py # Automated PowerPoint Catalog Generator
│
└── README.md
```

---

## 🚀 Quick Start

### 1. Configure the MCP Server in Antigravity / Claude Desktop

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

### 2. Run the Live Demo

Watch PowerPoint open and create a 3-slide animated presentation on your screen:
```bash
cd powerpoint-mcp
uv run python test_demo.py
```

### 3. Generate a 5-Slide Presentation from SQLite Database

```bash
cd database_4_produits
python generate_slides_from_db.py
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
| `ppt_add_card` | Add a modern UI container card with title, body, and accent bar |
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
| `ppt_run_slideshow` | Launch full-screen presentation mode |
| `ppt_slideshow_next` | Advance slideshow to next step/slide |
| `ppt_slideshow_previous` | Go back to previous slideshow slide |
| `ppt_slideshow_exit` | Exit slideshow mode back to editor |

---

## 📄 License
MIT License
