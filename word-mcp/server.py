"""Microsoft Word Live Automation & OOXML MCP Server.
Empowers AI coding assistants (Antigravity, Claude Desktop, Cursor) with full,
real-time control over Microsoft Word on Windows, plus fast headless .docx processing.
"""

from typing import Any, Dict, List, Optional
from mcp.server.fastmcp import FastMCP
from word_controller import controller
from docx_engine import DocxEngine

# Initialize FastMCP server
mcp = FastMCP("Word-Controller")


# ── LIVE WINDOWS WORD DESKTOP AUTOMATION ─────────────────────────────────────────

@mcp.tool()
def word_status() -> Dict[str, Any]:
    """Get the current running status of Microsoft Word and the active document.
    Returns whether Word is running, open document count, page count, word count,
    paragraph count, comments count, and active track-changes state.
    """
    return controller.get_info()


@mcp.tool()
def word_new_document(title: Optional[str] = None, template_path: Optional[str] = None) -> Dict[str, Any]:
    """Create a new blank document in Microsoft Word and bring it to foreground.
    Args:
        title: Optional document title to automatically insert and format at the top.
        template_path: Optional path to a .dotx / .docx template file.
    """
    return controller.new_document(title=title, template_path=template_path)


@mcp.tool()
def word_open(file_path: str) -> Dict[str, Any]:
    """Open an existing Word document (.docx, .doc) in Microsoft Word desktop and bring to front.
    Args:
        file_path: Absolute or relative path to the document.
    """
    return controller.open_document(file_path)


@mcp.tool()
def word_save(file_path: Optional[str] = None) -> Dict[str, Any]:
    """Save the active document in Word.
    Args:
        file_path: Optional new destination path (.docx). If omitted, saves in place.
    """
    return controller.save_document(file_path)


@mcp.tool()
def word_export_pdf(output_path: Optional[str] = None) -> Dict[str, Any]:
    """Export the active Word document directly to a high-quality PDF using Word's native engine.
    Args:
        output_path: Optional destination path (.pdf). If omitted, saves alongside the active file.
    """
    return controller.export_pdf(output_path)


@mcp.tool()
def word_close(save: bool = True) -> Dict[str, Any]:
    """Close the active document in Microsoft Word.
    Args:
        save: Whether to save changes before closing (default True).
    """
    return controller.close_document(save=save)


@mcp.tool()
def word_focus_window() -> Dict[str, Any]:
    """Bring Microsoft Word window to foreground and maximize it so the user can watch live changes."""
    controller.focus_window()
    return {"status": "Microsoft Word window focused."}


@mcp.tool()
def word_add_heading(
    text: str,
    level: int = 1,
    align: str = "left",
    color: str = "#1F4E79",
    space_before: float = 12.0,
    space_after: float = 6.0,
) -> Dict[str, Any]:
    """Insert a styled heading (level 1-6) into the active document.
    Args:
        text: Heading text.
        level: Heading level (1 to 6).
        align: 'left', 'center', or 'right'.
        color: Hex color string (e.g., '#1F4E79', '#2563EB').
        space_before: Space before heading in points.
        space_after: Space after heading in points.
    """
    return controller.add_heading(
        text=text,
        level=level,
        align=align,
        color=color,
        space_before=space_before,
        space_after=space_after,
    )


@mcp.tool()
def word_add_paragraph(
    text: str,
    font_name: str = "Calibri",
    font_size: float = 11.0,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    color: str = "#262626",
    highlight: Optional[str] = None,
    align: str = "left",
    space_after: float = 6.0,
    line_spacing: float = 1.15,
) -> Dict[str, Any]:
    """Add a styled paragraph to the active Word document.
    Args:
        text: Paragraph text.
        font_name: Typography font family (e.g. 'Calibri', 'Arial', 'Georgia').
        font_size: Font size in points.
        bold: Whether text is bold.
        italic: Whether text is italicized.
        underline: Whether text is underlined.
        color: Hex color (e.g. '#262626', '#000000').
        highlight: Optional highlight color ('yellow', 'green', 'blue', 'pink').
        align: 'left', 'center', 'right', or 'justify'.
        space_after: Spacing after paragraph in points.
        line_spacing: Multiplier (e.g. 1.0, 1.15, 1.5).
    """
    return controller.add_paragraph(
        text=text,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color,
        highlight=highlight,
        align=align,
        space_after=space_after,
        line_spacing=line_spacing,
    )


@mcp.tool()
def word_add_bullet_list(items: List[str], list_type: str = "bullet") -> Dict[str, Any]:
    """Add a bulleted or numbered list into the active document.
    Args:
        items: List of string items.
        list_type: 'bullet' or 'numbered'.
    """
    return controller.add_bullet_list(items=items, list_type=list_type)


@mcp.tool()
def word_add_callout_box(title: str, text: str, callout_type: str = "info") -> Dict[str, Any]:
    """Add a modern styled callout box with thick left colored stripe and soft background fill.
    Args:
        title: Callout heading (e.g. 'Important Note', 'Key Takeaway', 'Caution').
        text: Body text of the callout.
        callout_type: 'info' (blue), 'warning' (amber), 'tip' (green), 'note' (gray), 'danger' (red).
    """
    return controller.add_callout_box(title=title, text=text, callout_type=callout_type)


@mcp.tool()
def word_add_table(
    headers: List[str],
    rows: List[List[Any]],
    header_bg: str = "#1E293B",
    header_text_color: str = "#FFFFFF",
    alternate_shading: bool = True,
    alt_bg: str = "#F8FAFC",
    autofit: bool = True,
) -> Dict[str, Any]:
    """Insert a styled data table with customized headers, cell borders, and alternating rows.
    Args:
        headers: List of column header names.
        rows: List of data rows (each row is a list of values).
        header_bg: Hex color for header background (default '#1E293B').
        header_text_color: Hex color for header text (default '#FFFFFF').
        alternate_shading: Whether to alternate row shading for readability.
        alt_bg: Hex color for even rows (default '#F8FAFC').
        autofit: Whether to fit table width to window margins.
    """
    return controller.add_table(
        headers=headers,
        rows=rows,
        header_bg=header_bg,
        header_text_color=header_text_color,
        alternate_shading=alternate_shading,
        alt_bg=alt_bg,
        autofit=autofit,
    )


@mcp.tool()
def word_modify_table(
    table_index: int = 1,
    action: str = "set_cell",
    row: int = 1,
    col: int = 1,
    text: str = "",
    bg_color: Optional[str] = None,
    bold: Optional[bool] = None,
) -> Dict[str, Any]:
    """Modify an existing table in the document.
    Args:
        table_index: 1-based index of table in document.
        action: 'set_cell', 'add_row', 'add_column', 'delete_row', 'delete_column', 'autofit'.
        row: 1-based row index (for set_cell, delete_row).
        col: 1-based col index (for set_cell, delete_column).
        text: New text to set (for set_cell).
        bg_color: Hex color for cell shading.
        bold: Whether to bold the cell text.
    """
    return controller.modify_table(
        table_index=table_index,
        action=action,
        row=row,
        col=col,
        text=text,
        bg_color=bg_color,
        bold=bold,
    )


@mcp.tool()
def word_add_image(
    image_path: str,
    width_cm: Optional[float] = None,
    height_cm: Optional[float] = None,
    align: str = "center",
    caption: Optional[str] = None,
) -> Dict[str, Any]:
    """Insert an image with optional dimensions in cm, alignment, and caption.
    Args:
        image_path: Path to PNG, JPG, or WEBP image file.
        width_cm: Optional width in centimeters.
        height_cm: Optional height in centimeters.
        align: 'left', 'center', or 'right'.
        caption: Optional caption text printed below the figure.
    """
    return controller.add_image(
        image_path=image_path,
        width_cm=width_cm,
        height_cm=height_cm,
        align=align,
        caption=caption,
    )


@mcp.tool()
def word_add_page_break() -> Dict[str, Any]:
    """Insert a page break at current document position."""
    return controller.add_page_break()


@mcp.tool()
def word_add_section_break(break_type: str = "section_next") -> Dict[str, Any]:
    """Insert a section break into the document.
    Args:
        break_type: 'section_next' (new page) or 'section_continuous'.
    """
    return controller.add_section_break(break_type=break_type)


@mcp.tool()
def word_set_header_footer(
    header_text: Optional[str] = None,
    footer_text: Optional[str] = None,
    include_page_number: bool = True,
    different_first_page: bool = False,
) -> Dict[str, Any]:
    """Configure running headers and footers with dynamic page numbers ('Page X of Y').
    Args:
        header_text: Text placed in header (aligned right).
        footer_text: Text placed in footer (e.g. confidentiality, title).
        include_page_number: Whether to append dynamic page numbering.
        different_first_page: Whether cover page has no header/footer.
    """
    return controller.set_header_footer(
        header_text=header_text,
        footer_text=footer_text,
        include_page_number=include_page_number,
        different_first_page=different_first_page,
    )


@mcp.tool()
def word_add_watermark(
    text: str = "CONFIDENTIAL",
    font_size: float = 72.0,
    color: str = "#E2E8F0",
    orientation: str = "diagonal",
) -> Dict[str, Any]:
    """Add a translucent watermark across all pages (e.g., 'DRAFT', 'CONFIDENTIAL').
    Args:
        text: Watermark string.
        font_size: Font size in points (default 72).
        color: Hex color (default subtle gray '#E2E8F0').
        orientation: 'diagonal' or 'horizontal'.
    """
    return controller.add_watermark(
        text=text,
        font_size=font_size,
        color=color,
        orientation=orientation,
    )


@mcp.tool()
def word_remove_watermark() -> Dict[str, Any]:
    """Remove watermarks from active document."""
    return controller.remove_watermark()


@mcp.tool()
def word_add_comment(
    text: str,
    target_text: Optional[str] = None,
    author: str = "AI Assistant",
) -> Dict[str, Any]:
    """Add an editorial review comment anchored to target text or current selection.
    Args:
        text: Comment message.
        target_text: Exact text in document to attach the comment to.
        author: Author label for the comment.
    """
    return controller.add_comment(text=text, target_text=target_text, author=author)


@mcp.tool()
def word_list_comments() -> List[Dict[str, Any]]:
    """List all comments in active document, including author, date, text, and referenced phrase."""
    return controller.list_comments()


@mcp.tool()
def word_reply_comment(comment_index: int, reply_text: str, author: str = "AI Assistant") -> Dict[str, Any]:
    """Add a threaded reply to an existing comment.
    Args:
        comment_index: 1-based index of comment.
        reply_text: Text of reply.
        author: Author name.
    """
    return controller.reply_comment(comment_index=comment_index, reply_text=reply_text, author=author)


@mcp.tool()
def word_resolve_comment(comment_index: int, resolved: bool = True) -> Dict[str, Any]:
    """Mark a comment as resolved or reopen it.
    Args:
        comment_index: 1-based index of comment.
        resolved: True to resolve, False to reopen.
    """
    return controller.resolve_comment(comment_index=comment_index, resolved=resolved)


@mcp.tool()
def word_delete_comment(comment_index: int) -> Dict[str, Any]:
    """Delete a specific comment by index.
    Args:
        comment_index: 1-based index of comment.
    """
    return controller.delete_comment(comment_index=comment_index)


@mcp.tool()
def word_toggle_track_changes(enabled: Optional[bool] = None) -> Dict[str, Any]:
    """Toggle or explicitly set Word's Track Changes recording mode.
    Args:
        enabled: True to turn on, False to turn off, None to toggle.
    """
    return controller.toggle_track_changes(enabled=enabled)


@mcp.tool()
def word_list_revisions() -> List[Dict[str, Any]]:
    """List all tracked changes / revisions (insertions, deletions) in active document."""
    return controller.list_revisions()


@mcp.tool()
def word_accept_revisions(author: Optional[str] = None) -> Dict[str, Any]:
    """Accept tracked changes across the document.
    Args:
        author: Optional filter by author name. If omitted, accepts all revisions.
    """
    return controller.accept_revisions(author=author)


@mcp.tool()
def word_reject_revisions(author: Optional[str] = None) -> Dict[str, Any]:
    """Reject tracked changes across the document.
    Args:
        author: Optional filter by author name. If omitted, rejects all revisions.
    """
    return controller.reject_revisions(author=author)


@mcp.tool()
def word_search_and_replace(
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_whole_word: bool = False,
) -> Dict[str, Any]:
    """Find and replace text throughout the document.
    Args:
        find_text: Text string to find.
        replace_text: Replacement text.
        match_case: Whether to match case.
        match_whole_word: Whether to match whole words only.
    """
    return controller.search_and_replace(
        find_text=find_text,
        replace_text=replace_text,
        match_case=match_case,
        match_whole_word=match_whole_word,
    )


@mcp.tool()
def word_add_toc(title: str = "Table of Contents") -> Dict[str, Any]:
    """Insert a dynamic Table of Contents (TOC) based on heading styles (H1-H3).
    Args:
        title: Header text above TOC.
    """
    return controller.add_toc(title=title)


@mcp.tool()
def word_update_fields() -> Dict[str, Any]:
    """Update all dynamic document fields (Table of Contents, page numbers, cross-references)."""
    return controller.update_fields()


@mcp.tool()
def word_add_equation(equation_text: str) -> Dict[str, Any]:
    """Insert a mathematical equation using UnicodeMath / LaTeX syntax.
    Args:
        equation_text: Formula string (e.g. 'E = mc^2', 'x = (-b +- sqrt(b^2 - 4ac))/(2a)').
    """
    return controller.add_equation(equation_text=equation_text)


@mcp.tool()
def word_add_footnote(note_text: str, target_text: Optional[str] = None) -> Dict[str, Any]:
    """Add a footnote to the document.
    Args:
        note_text: Text of the footnote.
        target_text: Optional text anchor in document to attach footnote mark to.
    """
    return controller.add_footnote(note_text=note_text, target_text=target_text)


@mcp.tool()
def word_add_hyperlink(text: str, url: str) -> Dict[str, Any]:
    """Insert a clickable hyperlink with display text and URL destination.
    Args:
        text: Visible anchor text.
        url: Web URL destination (e.g. 'https://github.com').
    """
    return controller.add_hyperlink(text=text, url=url)


@mcp.tool()
def word_add_bookmark(name: str, target_text: Optional[str] = None) -> Dict[str, Any]:
    """Add a named bookmark for cross-referencing.
    Args:
        name: Unique bookmark identifier (letters/numbers/underscores).
        target_text: Optional text anchor to bookmark.
    """
    return controller.add_bookmark(name=name, target_text=target_text)


@mcp.tool()
def word_read_document_text(max_paragraphs: int = 200) -> Dict[str, Any]:
    """Extract paragraphs and styles from the active document in Word.
    Args:
        max_paragraphs: Maximum number of paragraphs to return.
    """
    return controller.read_document_text(max_paragraphs=max_paragraphs)


@mcp.tool()
def word_read_outline() -> List[Dict[str, Any]]:
    """Extract document heading outline (hierarchy of Headings 1-6)."""
    return controller.read_outline()


@mcp.tool()
def word_undo(count: int = 1) -> Dict[str, Any]:
    """Revert the last N operations in Word using Word's Undo stack.
    Args:
        count: Number of actions to undo (default 1).
    """
    return controller.undo(count=count)


@mcp.tool()
def word_take_screenshot(output_path: Optional[str] = None) -> Dict[str, Any]:
    """Take a screenshot of the Microsoft Word window for visual verification.
    Args:
        output_path: Optional path to save PNG screenshot.
    """
    return controller.take_screenshot(output_path=output_path)


# ── HEADLESS & PROGRAMMATIC OOXML / python-docx ENGINE ─────────────────────────

@mcp.tool()
def word_create_from_markdown(
    markdown_content: str,
    output_path: str,
    title: Optional[str] = None,
    author: Optional[str] = None,
) -> Dict[str, Any]:
    """Convert Markdown into a beautifully styled Word document (.docx).
    Supports headings, bold/italic, code blocks, callouts/quotes, bullet/numbered lists,
    tables with headers and alternating shading.
    Args:
        markdown_content: Full Markdown text.
        output_path: Destination path for the generated .docx file.
        title: Optional document title.
        author: Optional document author metadata.
    """
    return DocxEngine.create_from_markdown(
        markdown_text=markdown_content,
        output_path=output_path,
        title=title,
        author=author,
    )


@mcp.tool()
def word_fill_template(
    template_path: str,
    data: Dict[str, Any],
    output_path: str,
) -> Dict[str, Any]:
    """Fill placeholder variables like {{variable_name}} across an existing Word document.
    Replaces placeholders inside body paragraphs, tables, headers, and footers.
    Args:
        template_path: Path to template .docx file.
        data: Key-value dictionary of replacements (e.g. {'client_name': 'Acme Corp', 'date': '2026-10-10'}).
        output_path: Destination path for output .docx.
    """
    return DocxEngine.fill_template(
        template_path=template_path,
        data=data,
        output_path=output_path,
    )


@mcp.tool()
def word_compare_documents(
    original_path: str,
    revised_path: str,
    diff_output_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Compare two Word documents and generate a diff document with redlines (green insertions, red deletions).
    Args:
        original_path: Path to original baseline .docx.
        revised_path: Path to revised .docx.
        diff_output_path: Destination path for redline diff .docx.
    """
    return DocxEngine.compare_documents(
        original_path=original_path,
        revised_path=revised_path,
        diff_output_path=diff_output_path,
    )


@mcp.tool()
def word_generate_change_summary(
    original_path: str,
    revised_path: str,
    output_txt_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Generate an email-ready human-readable text changelog comparing two documents.
    Args:
        original_path: Path to base document.
        revised_path: Path to revised document.
        output_txt_path: Optional destination .txt file.
    """
    return DocxEngine.generate_change_summary(
        original_path=original_path,
        revised_path=revised_path,
        output_txt_path=output_txt_path,
    )


@mcp.tool()
def word_audit_document(file_path: str) -> Dict[str, Any]:
    """Audit a Word document's structure, statistics, paragraph health, and potential issues.
    Args:
        file_path: Path to .docx file.
    """
    return DocxEngine.audit_document(file_path=file_path)


@mcp.tool()
def word_scrub_metadata(input_path: str, output_path: Optional[str] = None) -> Dict[str, Any]:
    """Sanitize author, company, comments, and personal metadata from a Word document before delivery.
    Args:
        input_path: Path to input .docx.
        output_path: Destination path (if omitted, overwrites in place).
    """
    return DocxEngine.scrub_metadata(input_path=input_path, output_path=output_path)


@mcp.tool()
def word_add_cover_page(
    title: str,
    subtitle: Optional[str] = None,
    author: Optional[str] = None,
    organization: Optional[str] = None,
    date_str: Optional[str] = None,
    theme_color: str = "#1F4E79",
) -> Dict[str, Any]:
    """Insert a professional executive cover / title page at the start of the document.
    Configures a separate cover page layout without headers/footers.
    Args:
        title: Main document title.
        subtitle: Optional subtitle or project name.
        author: Author name.
        organization: Company or organization.
        date_str: Date string (defaults to current date).
        theme_color: Primary brand accent color (default '#1F4E79').
    """
    return controller.add_cover_page(
        title=title,
        subtitle=subtitle,
        author=author,
        organization=organization,
        date_str=date_str,
        theme_color=theme_color,
    )


@mcp.tool()
def word_set_page_setup(
    orientation: str = "portrait",
    margins: str = "normal",
    paper_size: str = "a4",
) -> Dict[str, Any]:
    """Configure page margins, orientation, and paper size for the document.
    Args:
        orientation: 'portrait' or 'landscape'.
        margins: 'normal' (1 in), 'narrow' (0.5 in), 'moderate' (0.75 in), or 'wide' (2 in).
        paper_size: 'a4', 'letter', or 'legal'.
    """
    return controller.set_page_setup(
        orientation=orientation,
        margins=margins,
        paper_size=paper_size,
    )


@mcp.tool()
def word_set_columns(
    column_count: int = 2,
    spacing_cm: float = 1.25,
    line_between: bool = False,
) -> Dict[str, Any]:
    """Set multi-column layout for newsletters, research papers, or executive briefs.
    Args:
        column_count: Number of columns (1, 2, or 3).
        spacing_cm: Gap spacing between columns in centimeters.
        line_between: Whether to draw a vertical line between columns.
    """
    return controller.set_columns(
        column_count=column_count,
        spacing_cm=spacing_cm,
        line_between=line_between,
    )


@mcp.tool()
def word_add_quote_card(
    quote: str,
    author: Optional[str] = None,
    title: Optional[str] = None,
    border_color: str = "#2563EB",
    bg_color: str = "#F8FAFC",
) -> Dict[str, Any]:
    """Insert a prominent pull-quote or customer testimonial card with thick accent border.
    Args:
        quote: Quote statement.
        author: Attribution author.
        title: Author's title or affiliation.
        border_color: Left accent border color (default '#2563EB').
        bg_color: Card background fill (default '#F8FAFC').
    """
    return controller.add_quote_card(
        quote=quote,
        author=author,
        title=title,
        border_color=border_color,
        bg_color=bg_color,
    )


@mcp.tool()
def word_add_kpi_cards(
    metrics: List[Dict[str, str]],
    header_color: str = "#1F4E79",
    bg_color: str = "#F8FAFC",
) -> Dict[str, Any]:
    """Insert a horizontal grid of modern KPI / stat metric cards side-by-side.
    Args:
        metrics: List of metrics dicts, e.g. [{'value': '$12.5M', 'label': 'ARR', 'change': '+28% YoY'}].
        header_color: Accent color for big numbers (default '#1F4E79').
        bg_color: Card background color (default '#F8FAFC').
    """
    return controller.add_kpi_cards(
        metrics=metrics,
        header_color=header_color,
        bg_color=bg_color,
    )


@mcp.tool()
def word_add_process_steps(
    steps: List[Dict[str, str]],
    accent_color: str = "#2563EB",
) -> Dict[str, Any]:
    """Insert a sequential step-by-step roadmap or timeline flow with numbered badges.
    Args:
        steps: List of step dicts, e.g. [{'step': 1, 'title': 'Discovery', 'description': 'Initial audits'}].
        accent_color: Badge color (default '#2563EB').
    """
    return controller.add_process_steps(
        steps=steps,
        accent_color=accent_color,
    )


@mcp.tool()
def word_add_endnote(note_text: str, target_text: Optional[str] = None) -> Dict[str, Any]:
    """Insert an endnote at the end of the document.
    Args:
        note_text: Text of the endnote.
        target_text: Optional text anchor to attach endnote mark to.
    """
    return controller.add_endnote(note_text=note_text, target_text=target_text)


@mcp.tool()
def word_list_endnotes() -> List[Dict[str, Any]]:
    """List all endnotes present in the document."""
    return controller.list_endnotes()


@mcp.tool()
def word_set_protection(
    protection_type: str = "revisions",
    password: str = "",
) -> Dict[str, Any]:
    """Enforce document protection against unintended edits.
    Args:
        protection_type: 'revisions' (allow only tracked changes), 'comments' (comments only), or 'readonly'.
        password: Optional protection password.
    """
    return controller.set_protection(protection_type=protection_type, password=password)


@mcp.tool()
def word_remove_protection(password: str = "") -> Dict[str, Any]:
    """Remove protection from the active document.
    Args:
        password: Password if document was locked with one.
    """
    return controller.remove_protection(password=password)


@mcp.tool()
def word_highlight_text(find_text: str, highlight_color: str = "yellow") -> Dict[str, Any]:
    """Find and highlight every occurrence of target text throughout the document.
    Args:
        find_text: Keyword or phrase to find.
        highlight_color: 'yellow', 'green', 'blue', or 'pink'.
    """
    return controller.highlight_text(find_text=find_text, highlight_color=highlight_color)


@mcp.tool()
def word_format_text(
    find_text: str,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[str] = None,
    font_size: Optional[float] = None,
) -> Dict[str, Any]:
    """Find and format occurrences of target text (change bold, italic, color, size).
    Args:
        find_text: Search phrase.
        bold: Set bold state.
        italic: Set italic state.
        color: Hex color string.
        font_size: Font size in points.
    """
    return controller.format_text(
        find_text=find_text,
        bold=bold,
        italic=italic,
        color=color,
        font_size=font_size,
    )


@mcp.tool()
def word_add_content_control(
    title: str,
    tag: Optional[str] = None,
    control_type: str = "rich_text",
    placeholder: str = "Click here to enter text.",
) -> Dict[str, Any]:
    """Insert a formal content control / form field widget.
    Args:
        title: Field title label.
        tag: Optional identifier tag.
        control_type: 'rich_text', 'plain_text', or 'date'.
        placeholder: Helper text shown before user types.
    """
    return controller.add_content_control(
        title=title,
        tag=tag,
        control_type=control_type,
        placeholder=placeholder,
    )


@mcp.tool()
def word_list_content_controls() -> List[Dict[str, Any]]:
    """List all content controls / form fields in the document."""
    return controller.list_content_controls()


@mcp.tool()
def word_add_chart(
    chart_type: str = "column",
    title: str = "Performance Overview",
) -> Dict[str, Any]:
    """Insert a native Microsoft Office chart (column, bar, line, pie).
    Args:
        chart_type: 'column', 'bar', 'line', or 'pie'.
        title: Chart title text.
    """
    return controller.add_chart(chart_type=chart_type, title=title)


@mcp.tool()
def word_merge_documents(file_paths: List[str], output_path: Optional[str] = None) -> Dict[str, Any]:
    """Merge multiple Word (.docx) files into a single master document.
    Args:
        file_paths: List of file paths in order of insertion.
        output_path: Optional destination file for the combined document.
    """
    return controller.merge_documents(file_paths=file_paths, output_path=output_path)


@mcp.tool()
def word_redact_text(target_text: str, replacement: str = "[REDACTED]") -> Dict[str, Any]:
    """Redact sensitive information across document with black highlight and replacement mark.
    Args:
        target_text: Sensitive phrase or pattern to redact.
        replacement: Label to replace with (default '[REDACTED]').
    """
    return controller.redact_text(target_text=target_text, replacement=replacement)


@mcp.tool()
def word_add_bates_numbering(
    prefix: str = "CONF-",
    start_number: int = 1,
    digits: int = 6,
) -> Dict[str, Any]:
    """Add legal Bates stamping numbering to the footer across all document pages.
    Args:
        prefix: Bates prefix (e.g. 'CONF-', 'PLAINTIFF-').
        start_number: Initial number sequence (default 1).
        digits: Zero-padding width (default 6 digits -> 000001).
    """
    return controller.add_bates_numbering(
        prefix=prefix,
        start_number=start_number,
        digits=digits,
    )


@mcp.tool()
def word_export_markdown(file_path: str, output_path: Optional[str] = None) -> Dict[str, Any]:
    """Convert any existing .docx document into clean Markdown formatted text.
    Args:
        file_path: Path to .docx file to extract.
        output_path: Optional destination .md file.
    """
    return DocxEngine.export_markdown(file_path=file_path, output_path=output_path)


def main():
    """Run FastMCP server."""
    mcp.run()


if __name__ == "__main__":
    main()

