"""PowerPoint Live Automation MCP Server
Enables AI assistants (Antigravity, Codex, Claude, Cursor) to control Microsoft PowerPoint
in real-time on Windows with live visual feedback.
"""

import json
from typing import Any, Dict, List, Optional
from mcp.server.fastmcp import FastMCP
from ppt_controller import controller

# Initialize FastMCP Server
mcp = FastMCP("PowerPoint-Live-Controller")


@mcp.tool()
def ppt_status() -> Dict[str, Any]:
    """Get the current status of Microsoft PowerPoint and the active presentation.
    Returns whether PowerPoint is running, number of open presentations, slide count, and current slide.
    """
    return controller.get_info()


@mcp.tool()
def ppt_new_presentation(width: float = 960.0, height: float = 540.0) -> Dict[str, Any]:
    """Create a new blank presentation with 16:9 widescreen layout and bring PowerPoint to foreground.
    Default width=960.0, height=540.0 (standard 16:9 widescreen).
    """
    return controller.new_presentation(width=width, height=height)


@mcp.tool()
def ppt_open(file_path: str) -> Dict[str, Any]:
    """Open an existing PowerPoint presentation file (.pptx, .ppt) and bring it to front.
    Args:
        file_path: Absolute or relative path to the PowerPoint file.
    """
    return controller.open_presentation(file_path)


@mcp.tool()
def ppt_save(file_path: Optional[str] = None) -> Dict[str, Any]:
    """Save the active presentation. If file_path ends with .pdf, exports to PDF.
    Args:
        file_path: Optional destination path (.pptx or .pdf). If omitted, saves to existing location.
    """
    return controller.save_presentation(file_path)


@mcp.tool()
def ppt_close(save: bool = True) -> Dict[str, Any]:
    """Close the active presentation.
    Args:
        save: Whether to save changes before closing (default True).
    """
    return controller.close_presentation(save=save)


@mcp.tool()
def ppt_list_slides() -> Dict[str, Any]:
    """List all slides in the active presentation, including their shape counts and title previews."""
    return controller.get_info()


@mcp.tool()
def ppt_goto_slide(slide_index: int) -> Dict[str, Any]:
    """Navigate PowerPoint's view to a specific slide (1-indexed) so the user sees it immediately.
    Args:
        slide_index: 1-indexed slide number.
    """
    return controller.goto_slide(slide_index)


@mcp.tool()
def ppt_add_slide(layout: str = "blank", position: Optional[int] = None) -> Dict[str, Any]:
    """Add a new slide to the presentation and automatically jump the active view to it.
    Args:
        layout: Layout name: "blank", "title", "title_and_content", "two_content", "title_only".
        position: Optional 1-indexed position. If omitted, appends to the end.
    """
    return controller.add_slide(layout=layout, position=position)


@mcp.tool()
def ppt_delete_slide(slide_index: int) -> Dict[str, Any]:
    """Delete a slide by its 1-indexed number.
    Args:
        slide_index: 1-indexed slide number.
    """
    return controller.delete_slide(slide_index)


@mcp.tool()
def ppt_set_background(
    slide_index: int,
    color_hex: str = "#0F172A",
    gradient_color_hex: Optional[str] = None,
) -> Dict[str, Any]:
    """Set solid or gradient background color for a slide.
    Args:
        slide_index: 1-indexed slide number.
        color_hex: Main hex color (e.g. '#0F172A' for dark navy, '#FFFFFF' for white).
        gradient_color_hex: Optional second hex color for a smooth 2-color gradient.
    """
    return controller.set_slide_background(
        slide_index=slide_index,
        color_hex=color_hex,
        gradient_color_hex=gradient_color_hex,
    )


@mcp.tool()
def ppt_add_textbox(
    slide_index: int,
    text: str,
    left: float = 80,
    top: float = 80,
    width: float = 800,
    height: float = 60,
    font_size: float = 22,
    bold: bool = False,
    italic: bool = False,
    font_name: str = "Segoe UI",
    color_hex: str = "#0F172A",
    bg_color_hex: Optional[str] = None,
    alignment: str = "left",
) -> Dict[str, Any]:
    """Add a styled text box to a slide.
    Args:
        slide_index: 1-indexed slide number.
        text: Text content (can include newlines).
        left: X position in points from top-left.
        top: Y position in points from top-left.
        width: Box width in points.
        height: Box height in points.
        font_size: Font size in pt (e.g. 36 for slide title, 20 for body).
        bold: Whether text is bold.
        italic: Whether text is italic.
        font_name: Font family name (e.g. 'Segoe UI', 'Arial', 'Calibri').
        color_hex: Font color in hex (e.g. '#FFFFFF', '#0F172A').
        bg_color_hex: Optional background fill color for the box.
        alignment: Text alignment: 'left', 'center', 'right', 'justify'.
    """
    return controller.add_textbox(
        slide_index=slide_index,
        text=text,
        left=left,
        top=top,
        width=width,
        height=height,
        font_size=font_size,
        bold=bold,
        italic=italic,
        font_name=font_name,
        color_hex=color_hex,
        bg_color_hex=bg_color_hex,
        alignment=alignment,
    )


@mcp.tool()
def ppt_add_card(
    slide_index: int,
    title: str,
    body: str,
    left: float,
    top: float,
    width: float,
    height: float,
    bg_color_hex: str = "#F8FAFC",
    border_color_hex: str = "#CBD5E1",
    title_color_hex: str = "#0F172A",
    body_color_hex: str = "#475569",
    accent_bar_color_hex: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a modern card component (rounded rectangle with title, description, and optional accent bar).
    Ideal for modern multi-column UI cards, feature lists, metrics, or comparisons.
    Args:
        slide_index: 1-indexed slide number.
        title: Card title.
        body: Card description / body text.
        left: X coordinate in points.
        top: Y coordinate in points.
        width: Width in points.
        height: Height in points.
        bg_color_hex: Background color (default soft light gray '#F8FAFC').
        border_color_hex: Border color (default '#CBD5E1').
        title_color_hex: Title text color (default '#0F172A').
        body_color_hex: Body text color (default '#475569').
        accent_bar_color_hex: Optional colorful accent bar on top (e.g. '#3B82F6').
    """
    return controller.add_card(
        slide_index=slide_index,
        title=title,
        body=body,
        left=left,
        top=top,
        width=width,
        height=height,
        bg_color_hex=bg_color_hex,
        border_color_hex=border_color_hex,
        title_color_hex=title_color_hex,
        body_color_hex=body_color_hex,
        accent_bar_color_hex=accent_bar_color_hex,
    )


@mcp.tool()
def ppt_add_shape(
    slide_index: int,
    shape_type: str = "rectangle",
    left: float = 100,
    top: float = 100,
    width: float = 200,
    height: float = 100,
    fill_color_hex: str = "#3B82F6",
    line_color_hex: Optional[str] = None,
    line_width: float = 1.5,
    text: str = "",
    text_color_hex: str = "#FFFFFF",
    font_size: float = 16,
    bold: bool = False,
) -> Dict[str, Any]:
    """Add an auto-shape (rectangle, rounded_rectangle, circle/oval, diamond, right_arrow, star, hexagon, heart, cloud).
    Args:
        slide_index: 1-indexed slide number.
        shape_type: 'rectangle', 'rounded_rectangle', 'circle', 'oval', 'diamond', 'right_arrow', 'star', etc.
        left: X position in points.
        top: Y position in points.
        width: Width in points.
        height: Height in points.
        fill_color_hex: Background fill color hex (e.g. '#3B82F6').
        line_color_hex: Optional border color hex.
        line_width: Border width in points.
        text: Optional text inside shape.
        text_color_hex: Text color hex.
        font_size: Font size in pt.
        bold: Whether text is bold.
    """
    return controller.add_shape(
        slide_index=slide_index,
        shape_type=shape_type,
        left=left,
        top=top,
        width=width,
        height=height,
        fill_color_hex=fill_color_hex,
        line_color_hex=line_color_hex,
        line_width=line_width,
        text=text,
        text_color_hex=text_color_hex,
        font_size=font_size,
        bold=bold,
    )


@mcp.tool()
def ppt_add_bullet_list(
    slide_index: int,
    items: List[str],
    left: float = 80,
    top: float = 150,
    width: float = 800,
    height: float = 320,
    font_size: float = 18,
    color_hex: str = "#334155",
    font_name: str = "Segoe UI",
) -> Dict[str, Any]:
    """Add a bullet point list to a slide.
    Args:
        slide_index: 1-indexed slide number.
        items: List of strings (one per bullet point).
        left: X position in points.
        top: Y position in points.
        width: Width in points.
        height: Height in points.
        font_size: Font size in pt.
        color_hex: Text color hex.
        font_name: Font family name.
    """
    return controller.add_bullet_list(
        slide_index=slide_index,
        items=items,
        left=left,
        top=top,
        width=width,
        height=height,
        font_size=font_size,
        color_hex=color_hex,
        font_name=font_name,
    )


@mcp.tool()
def ppt_add_table(
    slide_index: int,
    rows: int,
    cols: int,
    data: List[List[str]],
    left: float = 80,
    top: float = 140,
    width: float = 800,
    height: float = 300,
    header_bg_hex: str = "#2563EB",
    header_text_hex: str = "#FFFFFF",
    row_bg_hex: str = "#F8FAFC",
    alt_row_bg_hex: str = "#FFFFFF",
    text_color_hex: str = "#1E293B",
) -> Dict[str, Any]:
    """Add a styled data table with headers and alternating rows to a slide.
    Args:
        slide_index: 1-indexed slide number.
        rows: Number of rows.
        cols: Number of columns.
        data: 2D array of strings matching rows x cols.
        left: X position in points.
        top: Y position in points.
        width: Table width in points.
        height: Table height in points.
        header_bg_hex: Header row background color.
        header_text_hex: Header text color.
        row_bg_hex: Odd row background color.
        alt_row_bg_hex: Even row background color.
        text_color_hex: Body text color.
    """
    return controller.add_table(
        slide_index=slide_index,
        rows=rows,
        cols=cols,
        data=data,
        left=left,
        top=top,
        width=width,
        height=height,
        header_bg_hex=header_bg_hex,
        header_text_hex=header_text_hex,
        row_bg_hex=row_bg_hex,
        alt_row_bg_hex=alt_row_bg_hex,
        text_color_hex=text_color_hex,
    )


@mcp.tool()
def ppt_add_image(
    slide_index: int,
    image_path: str,
    left: float = 100,
    top: float = 100,
    width: Optional[float] = None,
    height: Optional[float] = None,
) -> Dict[str, Any]:
    """Insert an image file (.png, .jpg, .svg) from disk onto a slide.
    Args:
        slide_index: 1-indexed slide number.
        image_path: Path to the image file on local disk.
        left: X position in points.
        top: Y position in points.
        width: Optional width in points (keeps aspect ratio if omitted).
        height: Optional height in points.
    """
    return controller.add_image(
        slide_index=slide_index,
        image_path=image_path,
        left=left,
        top=top,
        width=width,
        height=height,
    )


@mcp.tool()
def ppt_add_animation(
    slide_index: int,
    shape_index_or_id: Any = 1,
    effect: str = "fade",
    trigger: str = "on_click",
) -> Dict[str, Any]:
    """Add visual entrance animation to a shape so elements animate live on screen!
    Args:
        slide_index: 1-indexed slide number.
        shape_index_or_id: 1-indexed shape number or shape ID. Defaults to last shape if omitted.
        effect: 'fade', 'fly_in', 'zoom', 'bounce', 'wipe', 'spin', 'rise_up', 'appear'.
        trigger: 'on_click' or 'after_previous'.
    """
    return controller.add_animation(
        slide_index=slide_index,
        shape_index_or_id=shape_index_or_id,
        effect=effect,
        trigger=trigger,
    )


@mcp.tool()
def ppt_set_transition(
    slide_index: int,
    transition: str = "fade",
    speed: str = "medium",
) -> Dict[str, Any]:
    """Set slide transition effect.
    Args:
        slide_index: 1-indexed slide number.
        transition: 'fade', 'push', 'wipe', 'cut', 'dissolve', 'checkerboard', 'random'.
        speed: 'slow', 'medium', 'fast'.
    """
    return controller.set_slide_transition(
        slide_index=slide_index,
        transition=transition,
        speed=speed,
    )


@mcp.tool()
def ppt_set_speaker_notes(slide_index: int, notes_text: str) -> Dict[str, Any]:
    """Add or update presenter speaker notes for a slide.
    Args:
        slide_index: 1-indexed slide number.
        notes_text: Presenter notes string.
    """
    return controller.set_speaker_notes(slide_index=slide_index, notes_text=notes_text)


@mcp.tool()
def ppt_run_slideshow(start_slide: int = 1) -> Dict[str, Any]:
    """Launch full-screen PowerPoint slideshow presentation live on screen.
    Args:
        start_slide: 1-indexed slide number to start presenting from (default 1).
    """
    return controller.run_slideshow(start_slide=start_slide)


@mcp.tool()
def ppt_slideshow_next() -> Dict[str, Any]:
    """Advance to the next slide or animation step during active slideshow."""
    return controller.slideshow_next()


@mcp.tool()
def ppt_slideshow_previous() -> Dict[str, Any]:
    """Go back to the previous slide or animation step during active slideshow."""
    return controller.slideshow_previous()


@mcp.tool()
def ppt_slideshow_exit() -> Dict[str, Any]:
    """Exit active slideshow and return to the normal PowerPoint editor view."""
    return controller.slideshow_exit()


@mcp.tool()
def ppt_export_slide_image(
    slide_index: int,
    output_path: str,
    width: int = 1920,
    height: int = 1080,
) -> Dict[str, Any]:
    """Export a slide to high-resolution PNG or JPG image file on disk.
    Args:
        slide_index: 1-indexed slide number.
        output_path: Local destination file path (e.g. 'slide_1.png' or 'slide_1.jpg').
        width: Image width in pixels (default 1920).
        height: Image height in pixels (default 1080).
    """
    return controller.export_slide_image(
        slide_index=slide_index,
        output_path=output_path,
        width=width,
        height=height,
    )


@mcp.tool()
def ppt_duplicate_slide(slide_index: int) -> Dict[str, Any]:
    """Duplicate an existing slide and view the new copy immediately.
    Args:
        slide_index: 1-indexed slide number to duplicate.
    """
    return controller.duplicate_slide(slide_index=slide_index)


@mcp.tool()
def ppt_move_slide(from_index: int, to_index: int) -> Dict[str, Any]:
    """Reorder/move a slide to a new position.
    Args:
        from_index: Current 1-indexed slide number.
        to_index: Target 1-indexed position.
    """
    return controller.move_slide(from_index=from_index, to_index=to_index)


@mcp.tool()
def ppt_read_slide_content(slide_index: int) -> Dict[str, Any]:
    """Inspect and extract all text, shapes, tables, coordinates, and notes from a slide.
    Allows LLM to read existing slides and understand their layout.
    Args:
        slide_index: 1-indexed slide number.
    """
    return controller.read_slide_content(slide_index=slide_index)


@mcp.tool()
def ppt_add_metric_card(
    slide_index: int,
    value: str,
    label: str,
    left: float,
    top: float,
    width: float = 240,
    height: float = 140,
    subtitle: Optional[str] = None,
    value_color_hex: str = "#2563EB",
    bg_color_hex: str = "#FFFFFF",
    border_color_hex: str = "#E2E8F0",
    accent_bar_color_hex: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a prominent KPI metric widget (large bold stat, label, and trend/subtext).
    Args:
        slide_index: 1-indexed slide number.
        value: Large stat string (e.g. '+142%', '$2.4M', '99.9%').
        label: Metric title (e.g. 'Annual Recurring Revenue').
        left: X position in points.
        top: Y position in points.
        width: Card width in points (default 240).
        height: Card height in points (default 140).
        subtitle: Optional trend or secondary text (e.g. '↑ 18% vs last quarter').
        value_color_hex: Color of the big number (default '#2563EB').
        bg_color_hex: Card background color.
        border_color_hex: Card border color.
        accent_bar_color_hex: Optional top colorful accent stripe.
    """
    return controller.add_metric_card(
        slide_index=slide_index,
        value=value,
        label=label,
        left=left,
        top=top,
        width=width,
        height=height,
        subtitle=subtitle,
        value_color_hex=value_color_hex,
        bg_color_hex=bg_color_hex,
        border_color_hex=border_color_hex,
        accent_bar_color_hex=accent_bar_color_hex,
    )


@mcp.tool()
def ppt_add_timeline(
    slide_index: int,
    steps: List[Dict[str, str]],
    left: float = 80,
    top: float = 160,
    width: float = 800,
    height: float = 240,
    accent_color_hex: str = "#3B82F6",
    bg_color_hex: str = "#FFFFFF",
) -> Dict[str, Any]:
    """Create a sleek horizontal roadmap / process flow / timeline with connected numbered nodes and cards.
    Args:
        slide_index: 1-indexed slide number.
        steps: List of dicts with 'title' and 'desc' keys, e.g. [{'title': 'Phase 1', 'desc': 'Discovery'}]
        left: X position in points.
        top: Y position in points.
        width: Total width across which steps will be distributed.
        height: Total height of the timeline component.
        accent_color_hex: Color for node circles and accent bars.
        bg_color_hex: Background color for step cards.
    """
    return controller.add_timeline(
        slide_index=slide_index,
        steps=steps,
        left=left,
        top=top,
        width=width,
        height=height,
        accent_color_hex=accent_color_hex,
        bg_color_hex=bg_color_hex,
    )


@mcp.tool()
def ppt_add_code_block(
    slide_index: int,
    code: str,
    language: str = "python",
    left: float = 80,
    top: float = 120,
    width: float = 800,
    height: float = 340,
    bg_color_hex: str = "#18181B",
    font_size: float = 13,
) -> Dict[str, Any]:
    """Create a dark developer code editor mockup with macOS window buttons and monospaced font.
    Args:
        slide_index: 1-indexed slide number.
        code: Source code text (can include multiple lines).
        language: Programming language name shown in header (e.g. 'python', 'javascript', 'sql').
        left: X position in points.
        top: Y position in points.
        width: Window width in points.
        height: Window height in points.
        bg_color_hex: Editor background color (default '#18181B').
        font_size: Font size in pt (default 13).
    """
    return controller.add_code_block(
        slide_index=slide_index,
        code=code,
        language=language,
        left=left,
        top=top,
        width=width,
        height=height,
        bg_color_hex=bg_color_hex,
        font_size=font_size,
    )


@mcp.tool()
def ppt_add_bar_chart(
    slide_index: int,
    data: List[Dict[str, Any]],
    left: float = 80,
    top: float = 140,
    width: float = 800,
    height: float = 280,
    bar_color_hex: str = "#3B82F6",
    bg_color_hex: str = "#F8FAFC",
) -> Dict[str, Any]:
    """Create a clean horizontal vector bar chart with labels and progress bars.
    Args:
        slide_index: 1-indexed slide number.
        data: List of dicts: [{'label': 'Product A', 'value': 85, 'display_value': '$85k', 'color': '#3B82F6'}, ...]
        left: X position in points.
        top: Y position in points.
        width: Chart width in points.
        height: Chart height in points.
        bar_color_hex: Default bar fill color.
        bg_color_hex: Container background color.
    """
    return controller.add_bar_chart(
        slide_index=slide_index,
        data=data,
        left=left,
        top=top,
        width=width,
        height=height,
        bar_color_hex=bar_color_hex,
        bg_color_hex=bg_color_hex,
    )


@mcp.tool()
def ppt_add_badge(
    slide_index: int,
    text: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: float = 26,
    bg_color_hex: str = "#EFF6FF",
    text_color_hex: str = "#1D4ED8",
    border_color_hex: Optional[str] = None,
    font_size: float = 11,
    bold: bool = True,
) -> Dict[str, Any]:
    """Create a rounded tag / pill badge component for categories or statuses.
    Args:
        slide_index: 1-indexed slide number.
        text: Badge text (e.g. 'FEATURED', 'Q3 RELEASE', 'ACTIVE').
        left: X position in points.
        top: Y position in points.
        width: Optional width in points (auto-computed from text length if omitted).
        height: Height in points (default 26).
        bg_color_hex: Background fill color.
        text_color_hex: Text font color.
        border_color_hex: Optional border color.
        font_size: Font size in pt (default 11).
        bold: Whether text is bold.
    """
    return controller.add_badge(
        slide_index=slide_index,
        text=text,
        left=left,
        top=top,
        width=width,
        height=height,
        bg_color_hex=bg_color_hex,
        text_color_hex=text_color_hex,
        border_color_hex=border_color_hex,
        font_size=font_size,
        bold=bold,
    )


@mcp.tool()
def ppt_search_and_replace_text(
    find_text: str,
    replace_text: str,
    slide_index: Optional[int] = None,
) -> Dict[str, Any]:
    """Find and replace text across all slides or within a specific slide (templates/variables).
    Args:
        find_text: Text string to search for (case-insensitive search).
        replace_text: Replacement text string.
        slide_index: Optional 1-indexed slide number. If omitted, replaces across all slides.
    """
    return controller.search_and_replace_text(
        find_text=find_text,
        replace_text=replace_text,
        slide_index=slide_index,
    )


@mcp.tool()
def ppt_add_quote_card(
    slide_index: int,
    quote: str,
    author: str,
    role: str = "",
    left: float = 80,
    top: float = 140,
    width: float = 800,
    height: float = 240,
    bg_color_hex: str = "#FFFFFF",
    border_color_hex: str = "#E2E8F0",
    quote_color_hex: str = "#0F172A",
    accent_color_hex: str = "#3B82F6",
) -> Dict[str, Any]:
    """Create a testimonial or quotation card with decorative quote mark and author details.
    Args:
        slide_index: 1-indexed slide number.
        quote: Quotation text.
        author: Author name (e.g. 'Steve Jobs', 'Satya Nadella').
        role: Author title or company (e.g. 'CEO, Microsoft').
        left: X position in points.
        top: Y position in points.
        width: Card width in points.
        height: Card height in points.
        bg_color_hex: Background card color.
        border_color_hex: Border color.
        quote_color_hex: Quote text font color.
        accent_color_hex: Stripe and quote icon color.
    """
    return controller.add_quote_card(
        slide_index=slide_index,
        quote=quote,
        author=author,
        role=role,
        left=left,
        top=top,
        width=width,
        height=height,
        bg_color_hex=bg_color_hex,
        border_color_hex=border_color_hex,
        quote_color_hex=quote_color_hex,
        accent_color_hex=accent_color_hex,
    )


@mcp.tool()
def ppt_add_pros_cons(
    slide_index: int,
    pros: List[str],
    cons: List[str],
    left: float = 80,
    top: float = 140,
    width: float = 800,
    height: float = 330,
) -> Dict[str, Any]:
    """Create side-by-side comparison cards for Pros (green) and Cons (red).
    Args:
        slide_index: 1-indexed slide number.
        pros: List of positive points / strengths.
        cons: List of negative points / risks.
        left: X position in points.
        top: Y position in points.
        width: Total width of both columns.
        height: Height of the comparison cards.
    """
    return controller.add_pros_cons(
        slide_index=slide_index,
        pros=pros,
        cons=cons,
        left=left,
        top=top,
        width=width,
        height=height,
    )


@mcp.tool()
def ppt_add_pricing_table(
    slide_index: int,
    tiers: List[Dict[str, Any]],
    left: float = 80,
    top: float = 130,
    width: float = 800,
    height: float = 360,
) -> Dict[str, Any]:
    """Create a SaaS pricing comparison cards layout (e.g. Free, Pro, Enterprise).
    Args:
        slide_index: 1-indexed slide number.
        tiers: List of tier dicts: [{'name': 'Pro', 'price': '$29', 'period': '/mo', 'highlighted': True, 'features': ['10 Projects', 'API Access']}]
        left: X position in points.
        top: Y position in points.
        width: Total width of the table.
        height: Total height of the cards.
    """
    return controller.add_pricing_table(
        slide_index=slide_index,
        tiers=tiers,
        left=left,
        top=top,
        width=width,
        height=height,
    )


@mcp.tool()
def ppt_add_donut_chart(
    slide_index: int,
    percentage: float,
    label: str,
    left: float = 380,
    top: float = 160,
    size: float = 180,
    track_color_hex: str = "#E2E8F0",
    fill_color_hex: str = "#2563EB",
    bg_color_hex: str = "#FFFFFF",
) -> Dict[str, Any]:
    """Create a circular percentage progress metric widget (donut).
    Args:
        slide_index: 1-indexed slide number.
        percentage: Number between 0 and 100.
        label: Label text underneath the chart.
        left: X position in points.
        top: Y position in points.
        size: Diameter in points (default 180).
        track_color_hex: Inactive track ring color.
        fill_color_hex: Active fill ring color.
        bg_color_hex: Inner cutout color (should match slide background).
    """
    return controller.add_donut_chart(
        slide_index=slide_index,
        percentage=percentage,
        label=label,
        left=left,
        top=top,
        size=size,
        track_color_hex=track_color_hex,
        fill_color_hex=fill_color_hex,
        bg_color_hex=bg_color_hex,
    )


@mcp.tool()
def ppt_delete_shape(
    slide_index: int,
    shape_id_or_name: Any,
) -> Dict[str, Any]:
    """Delete a specific shape by its ID or name from a slide.
    Args:
        slide_index: 1-indexed slide number.
        shape_id_or_name: ID integer/string or shape name.
    """
    return controller.delete_shape(
        slide_index=slide_index,
        shape_id_or_name=shape_id_or_name,
    )


@mcp.tool()
def ppt_clear_slide(slide_index: int) -> Dict[str, Any]:
    """Remove all shapes from a slide to reset it to completely blank.
    Args:
        slide_index: 1-indexed slide number.
    """
    return controller.clear_slide(slide_index=slide_index)


@mcp.tool()
def ppt_add_footer(
    slide_index: Optional[int] = None,
    text: str = "Confidential & Proprietary",
    show_slide_number: bool = True,
) -> Dict[str, Any]:
    """Add a professional bottom footer bar with notice and slide number.
    Args:
        slide_index: Optional 1-indexed slide number. If omitted, adds footer to ALL slides.
        text: Footer notice text (default 'Confidential & Proprietary').
        show_slide_number: Whether to display 'Slide X' on the bottom right (default True).
    """
    return controller.add_footer(
        slide_index=slide_index,
        text=text,
        show_slide_number=show_slide_number,
    )


def main():
    mcp.run()


if __name__ == "__main__":
    main()


