# Microsoft Office Live Automation MCP Suite 🚀💼

A state-of-the-art **Model Context Protocol (MCP)** suite for Windows enabling AI coding assistants (**Antigravity**, **Claude Desktop**, **Cursor**, **Codex**) to automate **Microsoft PowerPoint** and **Microsoft Word** in real-time desktop sessions with live visual feedback, alongside high-performance headless OOXML processing engines.

Built with **Windows COM Automation (`pywin32`)**, **`python-docx`**, and **FastMCP**.

---

## 📦 Included MCP Servers

### 1. [PowerPoint Live MCP (`powerpoint-mcp/`)](./powerpoint-mcp/README.md) 🎬
- **Live 16:9 Automation**: Real-time slide creation and shapes rendering.
- **Modern UI Components**: KPI metric cards, roadmap timelines, code editors, quote cards, pricing tables, donut charts, and data tables.
- **Animations & Transitions**: Entrance animations (fade, zoom, bounce, fly-in) and slide transitions.
- **Image Export**: Export slides directly to 1080p high-resolution images.

### 2. [Word Live Automation & OOXML MCP (`word-mcp/`)](./word-mcp/README.md) 📄 (64 Tools)
- **Live Desktop COM Automation**: Real-time document creation, executive cover pages, styled headings (H1–H6), executive typography, quote cards, KPI metric grids, process roadmaps, and tables.
- **Per-Action Undo (`Ctrl+Z`)**: Every tool call is isolated inside Word's native `UndoRecord` stack.
- **Tracked Changes & Editorial Review**: Toggle track changes, list revisions, accept/reject revisions, threaded comments, and resolution.
- **Legal & Compliance Tools**: Document protection modes, sensitive text redaction, legal Bates numbering, and content control form fields.
- **1-Click Native PDF Export**: Fast high-quality PDF output using Word's built-in engine.
- **Headless OOXML / python-docx Engine**: Markdown-to-DOCX conversion, DOCX-to-Markdown extraction, batch template filling (`{{variable}}`), redline document comparison, and structural integrity audit.
- **Companion Agent Skill**: Pre-configured skill teaching AI models best practices, typography rules, and common enterprise document workflows.

---

## 🚀 Quick Setup

In your `~/.gemini/config/mcp_config.json` (or `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "C:\\Users\\HP\\Desktop\\tkhrbi9\\mcps\\powerpoint-mcp\\.venv\\Scripts\\python.exe",
      "args": [
        "C:\\Users\\HP\\Desktop\\tkhrbi9\\mcps\\powerpoint-mcp\\server.py"
      ]
    },
    "word": {
      "command": "C:\\Users\\HP\\Desktop\\tkhrbi9\\mcps\\word-mcp\\.venv\\Scripts\\python.exe",
      "args": [
        "C:\\Users\\HP\\Desktop\\tkhrbi9\\mcps\\word-mcp\\server.py"
      ]
    }
  }
}
```

---

## 🧪 Live Demos

### Run PowerPoint Demo
```bash
cd powerpoint-mcp
uv run python test_demo.py
```

### Run Word Demo
```bash
cd word-mcp
uv run python test_demo.py
```

### Run Word Advanced Showcase (Cover Page, KPIs, Roadmap, Quotes, Markdown Export)
```bash
cd word-mcp
uv run python test_new_utils.py
```

---

## 📄 License
MIT License
