"""Headless OOXML & python-docx Engine for Word MCP.
Provides fast offline manipulation: Markdown to DOCX conversion, template filling,
document comparison, integrity auditing, and metadata scrubbing.
"""

import difflib
import os
import re
from typing import Any, Dict, List, Optional
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import mistune


def set_cell_background(cell, hex_color: str):
    """Set background color of a table cell in python-docx."""
    hex_color = hex_color.lstrip("#")
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set padding/margins for a table cell in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


class DocxEngine:
    """Headless document processing engine."""

    @staticmethod
    def create_from_markdown(
        markdown_text: str,
        output_path: str,
        title: Optional[str] = None,
        author: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Convert standard Markdown into a styled, professional Word document."""
        abs_output = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_output), exist_ok=True)

        doc = docx.Document()

        # Page margins (1 inch)
        for section in doc.sections:
            section.top_margin = Inches(1.0)
            section.bottom_margin = Inches(1.0)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)

        # Configure base Normal style
        normal_style = doc.styles['Normal']
        normal_font = normal_style.font
        normal_font.name = 'Calibri'
        normal_font.size = Pt(11)
        normal_font.color.rgb = RGBColor(0x26, 0x26, 0x26)

        if title:
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run(title)
            r_title.bold = True
            r_title.font.name = 'Calibri'
            r_title.font.size = Pt(26)
            r_title.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
            p_title.paragraph_format.space_after = Pt(18)

        # Parse markdown AST using mistune
        md = mistune.create_markdown(renderer='ast', plugins=['table', 'strikethrough'])
        tokens = md(markdown_text)

        for token in tokens:
            t_type = token.get('type')

            if t_type == 'heading':
                level = min(token.get('attrs', {}).get('level', 1), 6)
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(12 if level > 1 else 18)
                p.paragraph_format.space_after = Pt(6)

                colors = {
                    1: RGBColor(0x1F, 0x4E, 0x79),
                    2: RGBColor(0x25, 0x63, 0xEB),
                    3: RGBColor(0x37, 0x41, 0x51),
                }
                sizes = {1: Pt(18), 2: Pt(14), 3: Pt(12)}

                DocxEngine._render_inline_runs(p, token.get('children', []), base_font_size=sizes.get(level, Pt(11.5)), base_bold=True, base_color=colors.get(level, RGBColor(0x1F, 0x4E, 0x79)))

            elif t_type == 'paragraph':
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.15
                DocxEngine._render_inline_runs(p, token.get('children', []))

            elif t_type == 'block_code':
                raw_code = token.get('raw', '')
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.left_indent = Inches(0.2)
                r = p.add_run(raw_code.strip())
                r.font.name = 'Consolas'
                r.font.size = Pt(9.5)
                r.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

            elif t_type == 'block_quote':
                # Callout blockquote
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                for child in token.get('children', []):
                    if child.get('type') == 'paragraph':
                        DocxEngine._render_inline_runs(p, child.get('children', []), base_italic=True, base_color=RGBColor(0x4B, 0x55, 0x63))

            elif t_type == 'list':
                is_ordered = token.get('attrs', {}).get('ordered', False)
                for item in token.get('children', []):
                    p = doc.add_paragraph(style='List Number' if is_ordered else 'List Bullet')
                    p.paragraph_format.space_after = Pt(3)
                    for item_child in item.get('children', []):
                        if item_child.get('type') == 'paragraph':
                            DocxEngine._render_inline_runs(p, item_child.get('children', []))
                        elif item_child.get('type') == 'text':
                            p.add_run(item_child.get('raw', ''))

            elif t_type == 'table':
                DocxEngine._render_table(doc, token)

            elif t_type == 'thematic_break':
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(12)
                r = p.add_run("―" * 40)
                r.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)

        if author:
            core_props = doc.core_properties
            core_props.author = author

        doc.save(abs_output)
        return {
            "output_path": abs_output,
            "status": "Document created from Markdown successfully.",
        }

    @staticmethod
    def _render_inline_runs(
        paragraph,
        children: List[Dict[str, Any]],
        base_font_size=Pt(11),
        base_bold=False,
        base_italic=False,
        base_color=RGBColor(0x26, 0x26, 0x26),
    ):
        for child in children:
            c_type = child.get('type')
            if c_type == 'text':
                r = paragraph.add_run(child.get('raw', ''))
                r.font.size = base_font_size
                r.bold = base_bold
                r.italic = base_italic
                r.font.color.rgb = base_color
            elif c_type == 'strong':
                for sub in child.get('children', []):
                    r = paragraph.add_run(sub.get('raw', ''))
                    r.font.size = base_font_size
                    r.bold = True
                    r.italic = base_italic
                    r.font.color.rgb = base_color
            elif c_type == 'emphasis':
                for sub in child.get('children', []):
                    r = paragraph.add_run(sub.get('raw', ''))
                    r.font.size = base_font_size
                    r.bold = base_bold
                    r.italic = True
                    r.font.color.rgb = base_color
            elif c_type == 'codespan':
                r = paragraph.add_run(child.get('raw', ''))
                r.font.name = 'Consolas'
                r.font.size = Pt(10)
                r.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
            elif c_type == 'link':
                url = child.get('attrs', {}).get('url', '')
                for sub in child.get('children', []):
                    r = paragraph.add_run(sub.get('raw', url))
                    r.font.size = base_font_size
                    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
                    r.underline = True

    @staticmethod
    def _render_table(doc, table_token: Dict[str, Any]):
        children = table_token.get('children', [])
        head_rows = []
        body_rows = []

        for ch in children:
            if ch.get('type') == 'table_head':
                head_rows.extend(ch.get('children', []))
            elif ch.get('type') == 'table_body':
                body_rows.extend(ch.get('children', []))

        all_rows = head_rows + body_rows
        if not all_rows:
            return

        cols_count = max(len(r.get('children', [])) for r in all_rows)
        tbl = doc.add_table(rows=len(all_rows), cols=cols_count)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Format header
        for r_idx, row_tok in enumerate(all_rows):
            is_head = (r_idx < len(head_rows))
            row_cells = row_tok.get('children', [])
            for c_idx, cell_tok in enumerate(row_cells):
                cell = tbl.cell(r_idx, c_idx)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(2)

                text_content = ""
                for t in cell_tok.get('children', []):
                    if t.get('type') == 'text':
                        text_content += t.get('raw', '')
                    elif t.get('type') == 'strong':
                        for st in t.get('children', []):
                            text_content += st.get('raw', '')

                r = p.add_run(text_content.strip())
                r.font.name = 'Calibri'
                r.font.size = Pt(10)

                if is_head:
                    r.bold = True
                    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    set_cell_background(cell, "1E293B")
                else:
                    r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
                    if r_idx % 2 == 1:
                        set_cell_background(cell, "F8FAFC")

                set_cell_margins(cell, top=120, bottom=120, left=180, right=180)

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    @staticmethod
    def fill_template(template_path: str, data: Dict[str, Any], output_path: str) -> Dict[str, Any]:
        """Fill placeholders like {{key}} in paragraphs, tables, and headers."""
        abs_in = os.path.abspath(template_path)
        abs_out = os.path.abspath(output_path)
        if not os.path.isfile(abs_in):
            raise FileNotFoundError(f"Template not found: {abs_in}")

        doc = docx.Document(abs_in)
        replacements_count = 0

        def replace_in_text(text: str) -> Tuple[str, int]:
            cnt = 0
            for k, v in data.items():
                pattern = r"\{\{\s*" + re.escape(k) + r"\s*\}\}"
                if re.search(pattern, text):
                    text, n = re.subn(pattern, str(v), text)
                    cnt += n
            return text, cnt

        # 1. Body paragraphs
        for p in doc.paragraphs:
            if "{{" in p.text:
                new_text, n = replace_in_text(p.text)
                if n > 0:
                    p.text = new_text
                    replacements_count += n

        # 2. Tables
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if "{{" in p.text:
                            new_text, n = replace_in_text(p.text)
                            if n > 0:
                                p.text = new_text
                                replacements_count += n

        # 3. Headers and footers
        for sec in doc.sections:
            for p in sec.header.paragraphs:
                if "{{" in p.text:
                    new_text, n = replace_in_text(p.text)
                    if n > 0:
                        p.text = new_text
                        replacements_count += n
            for p in sec.footer.paragraphs:
                if "{{" in p.text:
                    new_text, n = replace_in_text(p.text)
                    if n > 0:
                        p.text = new_text
                        replacements_count += n

        os.makedirs(os.path.dirname(abs_out), exist_ok=True)
        doc.save(abs_out)

        return {
            "template": abs_in,
            "output_path": abs_out,
            "replacements_made": replacements_count,
            "status": f"Filled template with {replacements_count} substitutions.",
        }

    @staticmethod
    def compare_documents(
        original_path: str,
        revised_path: str,
        diff_output_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Compare two Word documents and generate a diff document with redlines."""
        abs_orig = os.path.abspath(original_path)
        abs_rev = os.path.abspath(revised_path)

        doc1 = docx.Document(abs_orig)
        doc2 = docx.Document(abs_rev)

        t1 = [p.text.strip() for p in doc1.paragraphs if p.text.strip()]
        t2 = [p.text.strip() for p in doc2.paragraphs if p.text.strip()]

        differ = difflib.Differ()
        diff = list(differ.compare(t1, t2))

        diff_doc = docx.Document()
        p_hdr = diff_doc.add_paragraph()
        r_hdr = p_hdr.add_run(f"Document Comparison: {os.path.basename(abs_orig)} vs {os.path.basename(abs_rev)}\n")
        r_hdr.bold = True
        r_hdr.font.size = Pt(16)
        r_hdr.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

        insertions = 0
        deletions = 0
        unchanged = 0

        for line in diff:
            code = line[:2]
            text = line[2:]

            if code == '+ ':
                insertions += 1
                p = diff_doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.2)
                r = p.add_run(f"[INSERTION] {text}")
                r.font.color.rgb = RGBColor(0x05, 0x96, 0x69)
                r.underline = True
            elif code == '- ':
                deletions += 1
                p = diff_doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.2)
                r = p.add_run(f"[DELETION] {text}")
                r.font.color.rgb = RGBColor(0xDC, 0x26, 0x26)
                r.font.strike = True
            elif code == '  ':
                unchanged += 1
                p = diff_doc.add_paragraph()
                r = p.add_run(text)
                r.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

        if not diff_output_path:
            base, _ = os.path.splitext(abs_rev)
            diff_output_path = base + "_diff.docx"

        abs_diff = os.path.abspath(diff_output_path)
        diff_doc.save(abs_diff)

        return {
            "diff_path": abs_diff,
            "insertions": insertions,
            "deletions": deletions,
            "unchanged_paragraphs": unchanged,
            "status": "Documents compared successfully.",
        }

    @staticmethod
    def generate_change_summary(
        original_path: str,
        revised_path: str,
        output_txt_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Generate a human-readable email-ready plain text changelog."""
        abs_orig = os.path.abspath(original_path)
        abs_rev = os.path.abspath(revised_path)

        doc1 = docx.Document(abs_orig)
        doc2 = docx.Document(abs_rev)

        t1 = [p.text.strip() for p in doc1.paragraphs if p.text.strip()]
        t2 = [p.text.strip() for p in doc2.paragraphs if p.text.strip()]

        differ = difflib.Differ()
        diff = list(differ.compare(t1, t2))

        lines = [
            f"=== DOCUMENT CHANGE LOG ===",
            f"Original: {os.path.basename(abs_orig)}",
            f"Revised:  {os.path.basename(abs_rev)}",
            f"Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}",
            "--------------------------------------------------\n",
        ]

        item_num = 1
        for d in diff:
            code = d[:2]
            txt = d[2:]
            if code == '+ ':
                lines.append(f"[{item_num}] INSERTION:\n    \"{txt}\"\n")
                item_num += 1
            elif code == '- ':
                lines.append(f"[{item_num}] DELETION:\n    \"{txt}\"\n")
                item_num += 1

        summary_text = "\n".join(lines)

        if output_txt_path:
            abs_out = os.path.abspath(output_txt_path)
            with open(abs_out, "w", encoding="utf-8") as f:
                f.write(summary_text)

        return {
            "change_count": item_num - 1,
            "summary": summary_text,
            "saved_to": output_txt_path,
        }

    @staticmethod
    def audit_document(file_path: str) -> Dict[str, Any]:
        """Scan a .docx document for structural integrity, broken elements, and statistics."""
        abs_path = os.path.abspath(file_path)
        if not os.path.isfile(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        doc = docx.Document(abs_path)
        total_p = len(doc.paragraphs)
        empty_p = sum(1 for p in doc.paragraphs if not p.text.strip())
        headings = sum(1 for p in doc.paragraphs if p.style.name.startswith("Heading"))
        tables = len(doc.tables)

        words = sum(len(p.text.split()) for p in doc.paragraphs)
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    words += sum(len(p.text.split()) for p in cell.paragraphs)

        warnings = []
        if empty_p > total_p * 0.4 and total_p > 10:
            warnings.append("High number of blank paragraphs detected (>40%). Consider using paragraph spacing.")
        if headings == 0 and total_p > 5:
            warnings.append("No headings found in document. Structure may lack hierarchy.")

        return {
            "file": abs_path,
            "paragraphs_count": total_p,
            "empty_paragraphs": empty_p,
            "headings_count": headings,
            "tables_count": tables,
            "estimated_words": words,
            "warnings": warnings,
            "status": "healthy" if not warnings else "warnings_found",
        }

    @staticmethod
    def scrub_metadata(input_path: str, output_path: Optional[str] = None) -> Dict[str, Any]:
        """Sanitize personal identifiable information (author, company, comments) from .docx."""
        abs_in = os.path.abspath(input_path)
        if not output_path:
            output_path = abs_in
        abs_out = os.path.abspath(output_path)

        doc = docx.Document(abs_in)
        core = doc.core_properties
        core.author = ""
        core.last_modified_by = ""
        core.comments = ""
        core.keywords = ""
        core.category = ""
        core.subject = ""
        core.title = ""

        doc.save(abs_out)
        return {"scrubbed_file": abs_out, "status": "Document metadata scrubbed successfully."}

    @staticmethod
    def export_markdown(file_path: str, output_path: Optional[str] = None) -> Dict[str, Any]:
        """Convert a .docx document back into clean Markdown formatted text."""
        abs_in = os.path.abspath(file_path)
        if not os.path.isfile(abs_in):
            raise FileNotFoundError(f"File not found: {abs_in}")

        doc = docx.Document(abs_in)
        md_lines = []

        # Iterate body elements in order
        for child in doc.element.body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag == "p":
                # Find corresponding python-docx paragraph
                p_text = ""
                p_style = ""
                for p in doc.paragraphs:
                    if p._p == child:
                        p_style = p.style.name if p.style else ""
                        run_parts = []
                        for r in p.runs:
                            t = r.text
                            if not t:
                                continue
                            if r.bold and r.italic:
                                t = f"***{t}***"
                            elif r.bold:
                                t = f"**{t}**"
                            elif r.italic:
                                t = f"*{t}*"
                            run_parts.append(t)
                        p_text = "".join(run_parts)
                        break

                if not p_text.strip():
                    continue

                if "Heading 1" in p_style or "Titre 1" in p_style:
                    md_lines.append(f"\n# {p_text}\n")
                elif "Heading 2" in p_style or "Titre 2" in p_style:
                    md_lines.append(f"\n## {p_text}\n")
                elif "Heading 3" in p_style or "Titre 3" in p_style:
                    md_lines.append(f"\n### {p_text}\n")
                elif "List Bullet" in p_style:
                    md_lines.append(f"- {p_text}")
                elif "List Number" in p_style:
                    md_lines.append(f"1. {p_text}")
                else:
                    md_lines.append(f"{p_text}\n")

            elif tag == "tbl":
                for tbl in doc.tables:
                    if tbl._tbl == child:
                        rows = []
                        for row in tbl.rows:
                            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                            rows.append(cells)
                        if rows:
                            # Markdown table
                            header = "| " + " | ".join(rows[0]) + " |"
                            sep = "| " + " | ".join(["---"] * len(rows[0])) + " |"
                            table_md = [header, sep]
                            for r in rows[1:]:
                                table_md.append("| " + " | ".join(r) + " |")
                            md_lines.append("\n" + "\n".join(table_md) + "\n")
                        break

        markdown_result = "\n".join(md_lines).strip() + "\n"

        if output_path:
            abs_out = os.path.abspath(output_path)
            with open(abs_out, "w", encoding="utf-8") as f:
                f.write(markdown_result)
            return {"markdown": markdown_result, "output_path": abs_out, "status": "Markdown exported."}

        return {"markdown": markdown_result, "status": "Markdown exported."}

