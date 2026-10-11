"""Microsoft Word COM Automation Controller for Windows.
Provides real-time, live visual control of Microsoft Word desktop via Windows COM.
"""

import os
import sys
import time
from contextlib import contextmanager
from typing import Any, Dict, List, Optional, Tuple

if sys.platform == "win32":
    import pythoncom
    import win32com.client
    import win32gui
else:
    win32com = None
    win32gui = None

# Alignment types
WD_ALIGN = {
    "left": 0,       # wdAlignParagraphLeft
    "center": 1,     # wdAlignParagraphCenter
    "right": 2,      # wdAlignParagraphRight
    "justify": 3,    # wdAlignParagraphJustify
}

# Break types
WD_BREAKS = {
    "page": 7,              # wdPageBreak
    "column": 8,            # wdColumnBreak
    "section_next": 2,      # wdSectionBreakNextPage
    "section_continuous": 3,# wdSectionBreakContinuous
    "section_even": 4,      # wdSectionBreakEvenPage
    "section_odd": 5,       # wdSectionBreakOddPage
}

# Color helpers
def hex_to_bgr_int(hex_str: str) -> int:
    """Converts a hex color string (#RRGGBB or #RGB) to COM BGR/RGB integer."""
    hex_str = hex_str.strip().lstrip("#")
    if len(hex_str) == 3:
        hex_str = "".join([c * 2 for c in hex_str])
    if len(hex_str) != 6:
        return 0
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    # Word COM color is BGR integer: R + (G * 256) + (B * 65536)
    return r + (g * 256) + (b * 65536)


def cm_to_points(cm: float) -> float:
    """Convert centimeters to typographic points (1 cm = ~28.3465 pt)."""
    return cm * 28.3465


class WordController:
    """Manages live visual interaction with Microsoft Word desktop on Windows."""

    def __init__(self):
        self._app = None
        self._doc = None

    def get_app(self, bring_to_front: bool = True):
        """Connect to running Word or launch a new visible instance."""
        if sys.platform != "win32":
            raise RuntimeError("Word COM controller is only supported on Windows.")

        pythoncom.CoInitialize()

        try:
            self._app = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            self._app = win32com.client.Dispatch("Word.Application")

        self._app.Visible = True
        try:
            self._app.WindowState = 1  # wdWindowStateMaximize
            self._app.Activate()
        except Exception:
            pass

        if bring_to_front:
            self.focus_window()

        return self._app

    def focus_window(self):
        """Bring Microsoft Word window to foreground."""
        try:
            if win32gui:
                hwnd = win32gui.FindWindow("OpusApp", None)
                if hwnd:
                    win32gui.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                    win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass

    @contextmanager
    def undo_block(self, name: str = "AI Action"):
        """Wrap operations into a single Word UndoRecord (one Ctrl+Z)."""
        app = self.get_app(bring_to_front=False)
        rec = None
        try:
            rec = app.UndoRecord
            if rec.IsRecordingCustomRecord:
                try:
                    rec.EndCustomRecord()
                except Exception:
                    pass
            rec.StartCustomRecord(name[:64])
        except Exception:
            rec = None

        try:
            yield
        finally:
            if rec is not None:
                try:
                    rec.EndCustomRecord()
                except Exception:
                    pass

    def get_document(self, create_if_none: bool = True):
        """Get the current active document, or create one if none is open."""
        app = self.get_app()
        try:
            if app.Documents.Count > 0:
                self._doc = app.ActiveDocument
                return self._doc
        except Exception:
            pass

        if create_if_none:
            self._doc = app.Documents.Add()
            return self._doc
        raise RuntimeError("No document is open in Microsoft Word.")

    def get_info(self) -> Dict[str, Any]:
        """Get current status of Word and active document."""
        try:
            app = self.get_app(bring_to_front=False)
            doc_count = app.Documents.Count
            if doc_count == 0:
                return {
                    "running": True,
                    "documents_count": 0,
                    "active_document": None,
                    "status": "Word is running with no open documents.",
                }

            doc = app.ActiveDocument
            words = doc.Words.Count
            chars = doc.Characters.Count
            paras = doc.Paragraphs.Count
            tables = doc.Tables.Count
            pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
            comments = doc.Comments.Count
            revisions = doc.Revisions.Count

            return {
                "running": True,
                "documents_count": doc_count,
                "active_document": {
                    "name": doc.Name,
                    "path": doc.FullName,
                    "saved": doc.Saved,
                    "pages": pages,
                    "words": words,
                    "characters": chars,
                    "paragraphs": paras,
                    "tables": tables,
                    "comments": comments,
                    "revisions": revisions,
                    "track_changes": doc.TrackRevisions,
                },
                "status": "success",
            }
        except Exception as e:
            return {
                "running": False,
                "error": str(e),
                "status": "Word is not running or not accessible.",
            }

    def new_document(self, title: Optional[str] = None, template_path: Optional[str] = None) -> Dict[str, Any]:
        """Create a new document and bring Word to foreground."""
        app = self.get_app()
        if template_path and os.path.isfile(template_path):
            self._doc = app.Documents.Add(Template=os.path.abspath(template_path))
        else:
            self._doc = app.Documents.Add()

        self.focus_window()

        if title:
            with self.undo_block("Add Title"):
                rng = self._doc.Range(0, 0)
                rng.Text = title + "\n"
                rng.Font.Name = "Calibri"
                rng.Font.Size = 26
                rng.Font.Bold = True
                rng.Font.Color = hex_to_bgr_int("#1F4E79")
                rng.ParagraphFormat.Alignment = WD_ALIGN["center"]
                rng.ParagraphFormat.SpaceAfter = 18

        return {
            "name": self._doc.Name,
            "status": "New document created successfully.",
        }

    def open_document(self, file_path: str) -> Dict[str, Any]:
        """Open an existing Word document."""
        abs_path = os.path.abspath(file_path)
        if not os.path.isfile(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        app = self.get_app()
        self._doc = app.Documents.Open(abs_path)
        self.focus_window()

        return {
            "name": self._doc.Name,
            "path": self._doc.FullName,
            "status": "Document opened successfully.",
        }

    def save_document(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Save the active document."""
        doc = self.get_document()
        if file_path:
            abs_path = os.path.abspath(file_path)
            # Create parent dirs if necessary
            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
            doc.SaveAs2(abs_path)
            return {"saved_to": abs_path, "status": "Document saved successfully."}
        else:
            doc.Save()
            return {"saved_to": doc.FullName, "status": "Document saved successfully."}

    def export_pdf(self, output_path: Optional[str] = None) -> Dict[str, Any]:
        """Export the active document to PDF via Word COM ExportAsFixedFormat."""
        doc = self.get_document()
        if not output_path:
            base, _ = os.path.splitext(doc.FullName)
            output_path = base + ".pdf"
        abs_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)

        # wdExportFormatPDF = 17
        doc.ExportAsFixedFormat(
            OutputFileName=abs_path,
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,  # wdExportOptimizeForPrint
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )

        return {
            "pdf_path": abs_path,
            "status": "Document successfully exported to PDF.",
        }

    def close_document(self, save: bool = True) -> Dict[str, Any]:
        """Close active document."""
        doc = self.get_document()
        name = doc.Name
        # wdSaveChanges = -1, wdDoNotSaveChanges = 0
        save_const = -1 if save else 0
        doc.Close(SaveChanges=save_const)
        self._doc = None
        return {"closed": name, "saved": save, "status": "Document closed."}

    def add_heading(
        self,
        text: str,
        level: int = 1,
        align: str = "left",
        color: str = "#1F4E79",
        space_before: float = 12,
        space_after: float = 6,
    ) -> Dict[str, Any]:
        """Add a heading (level 1-6) with styling."""
        doc = self.get_document()
        with self.undo_block(f"Add Heading {level}"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)  # wdStory = 6 (jump to end)

            # Insert heading text with newline
            sel.TypeText(text + "\n")
            # Select the newly typed paragraph
            p = doc.Paragraphs(doc.Paragraphs.Count - 1)
            rng = p.Range

            # Apply Word built-in Heading style if available
            try:
                p.Style = f"Heading {level}"
            except Exception:
                pass

            # Fine-tune typography
            rng.Font.Color = hex_to_bgr_int(color)
            rng.Font.Bold = True
            if level == 1:
                rng.Font.Size = 20
            elif level == 2:
                rng.Font.Size = 16
            elif level == 3:
                rng.Font.Size = 13
            else:
                rng.Font.Size = 12

            p.Alignment = WD_ALIGN.get(align.lower(), 0)
            p.SpaceBefore = space_before
            p.SpaceAfter = space_after

        return {
            "text": text,
            "level": level,
            "status": f"Heading {level} added successfully.",
        }

    def add_paragraph(
        self,
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
        """Add a formatted paragraph."""
        doc = self.get_document()
        with self.undo_block("Add Paragraph"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)  # wdStory

            sel.TypeText(text + "\n")
            p = doc.Paragraphs(doc.Paragraphs.Count - 1)
            rng = p.Range

            rng.Font.Name = font_name
            rng.Font.Size = font_size
            rng.Font.Bold = bold
            rng.Font.Italic = italic
            rng.Font.Underline = 1 if underline else 0
            rng.Font.Color = hex_to_bgr_int(color)

            if highlight:
                # 7=wdYellow, 4=wdBrightGreen, 3=wdTurquoise, 6=wdPink
                hl_map = {"yellow": 7, "green": 4, "blue": 3, "pink": 6}
                rng.HighlightColorIndex = hl_map.get(highlight.lower(), 7)

            p.Alignment = WD_ALIGN.get(align.lower(), 0)
            p.SpaceAfter = space_after
            p.LineSpacingRule = 5  # wdLineSpaceMultiple
            p.LineSpacing = line_spacing * 12

        return {"status": "Paragraph added successfully.", "length": len(text)}

    def add_bullet_list(
        self,
        items: List[str],
        list_type: str = "bullet",
        space_after: float = 3.0,
    ) -> Dict[str, Any]:
        """Add a bulleted or numbered list."""
        doc = self.get_document()
        with self.undo_block(f"Add {list_type} List"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            start_p_idx = doc.Paragraphs.Count
            for item in items:
                sel.TypeText(item + "\n")

            end_p_idx = doc.Paragraphs.Count - 1
            # Select range of all items
            r_start = doc.Paragraphs(start_p_idx).Range.Start
            r_end = doc.Paragraphs(end_p_idx).Range.End
            rng = doc.Range(r_start, r_end)

            if list_type.lower() == "numbered":
                rng.ListFormat.ApplyNumberDefault()
            else:
                rng.ListFormat.ApplyBulletDefault()

            for idx in range(start_p_idx, end_p_idx + 1):
                p = doc.Paragraphs(idx)
                p.SpaceAfter = space_after

        return {"count": len(items), "type": list_type, "status": "List added."}

    def add_callout_box(
        self,
        title: str,
        text: str,
        callout_type: str = "info",
    ) -> Dict[str, Any]:
        """Add modern callout alert box with left accent stripe and background tint."""
        doc = self.get_document()

        palettes = {
            "info": {"border": "#2563EB", "bg": "#EFF6FF", "icon": "ℹ️"},
            "warning": {"border": "#D97706", "bg": "#FFFBEB", "icon": "⚠️"},
            "tip": {"border": "#059669", "bg": "#ECFDF5", "icon": "💡"},
            "note": {"border": "#4B5563", "bg": "#F3F4F6", "icon": "📝"},
            "danger": {"border": "#DC2626", "bg": "#FEF2F2", "icon": "🚨"},
        }
        palette = palettes.get(callout_type.lower(), palettes["info"])

        with self.undo_block(f"Add Callout Box ({callout_type})"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            # Insert a 1x1 table for modern callout card
            rng = sel.Range
            tbl = doc.Tables.Add(Range=rng, NumRows=1, NumColumns=1)
            tbl.Borders.Enable = False

            cell = tbl.Cell(1, 1)
            # Left border thick colored accent
            left_border = cell.Borders(-2)  # wdBorderLeft = -2
            left_border.LineStyle = 1        # wdLineStyleSingle
            left_border.LineWidth = 36       # 4.5 pt (36/8)
            left_border.Color = hex_to_bgr_int(palette["border"])

            # Background tint
            cell.Shading.BackgroundPatternColor = hex_to_bgr_int(palette["bg"])

            # Margins/padding inside callout
            tbl.TopPadding = 8
            tbl.BottomPadding = 8
            tbl.LeftPadding = 12
            tbl.RightPadding = 12

            # Populate text
            c_rng = cell.Range
            c_rng.Text = f"{palette['icon']}  {title}\n{text}"
            
            # Format title line
            p_title = c_rng.Paragraphs(1)
            p_title.Range.Font.Bold = True
            p_title.Range.Font.Size = 11.5
            p_title.Range.Font.Color = hex_to_bgr_int(palette["border"])
            p_title.SpaceAfter = 4

            # Format body line
            if c_rng.Paragraphs.Count > 1:
                p_body = c_rng.Paragraphs(2)
                p_body.Range.Font.Bold = False
                p_body.Range.Font.Size = 10.5
                p_body.Range.Font.Color = hex_to_bgr_int("#374151")

            sel.EndKey(6)
            sel.TypeText("\n")

        return {"title": title, "type": callout_type, "status": "Callout box added."}

    def add_table(
        self,
        headers: List[str],
        rows: List[List[Any]],
        header_bg: str = "#1E293B",
        header_text_color: str = "#FFFFFF",
        alternate_shading: bool = True,
        alt_bg: str = "#F8FAFC",
        autofit: bool = True,
    ) -> Dict[str, Any]:
        """Add styled data table with customized headers, borders, and alternating rows."""
        doc = self.get_document()
        num_rows = len(rows) + 1
        num_cols = len(headers)

        with self.undo_block("Add Data Table"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            tbl = doc.Tables.Add(Range=sel.Range, NumRows=num_rows, NumColumns=num_cols)
            tbl.Borders.Enable = True
            # Subtle border color
            tbl.Borders.InsideColor = hex_to_bgr_int("#E2E8F0")
            tbl.Borders.OutsideColor = hex_to_bgr_int("#CBD5E1")

            # Header formatting
            for col_idx, header_text in enumerate(headers, start=1):
                cell = tbl.Cell(1, col_idx)
                cell.Range.Text = str(header_text)
                cell.Range.Font.Bold = True
                cell.Range.Font.Color = hex_to_bgr_int(header_text_color)
                cell.Range.Font.Size = 10.5
                cell.Shading.BackgroundPatternColor = hex_to_bgr_int(header_bg)
                cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
                cell.Range.ParagraphFormat.Alignment = WD_ALIGN["center"]

            # Repeat header row across pages
            tbl.Rows(1).HeadingFormat = True

            # Data rows
            for r_idx, row_data in enumerate(rows, start=2):
                is_even = (r_idx % 2 == 0)
                for c_idx, val in enumerate(row_data, start=1):
                    if c_idx > num_cols:
                        break
                    cell = tbl.Cell(r_idx, c_idx)
                    cell.Range.Text = str(val)
                    cell.Range.Font.Size = 10.0
                    cell.Range.Font.Color = hex_to_bgr_int("#1E293B")
                    cell.VerticalAlignment = 1

                    if alternate_shading and is_even:
                        cell.Shading.BackgroundPatternColor = hex_to_bgr_int(alt_bg)

            if autofit:
                tbl.AutoFitBehavior(2)  # wdAutoFitWindow = 2

            sel.EndKey(6)
            sel.TypeText("\n")

        return {
            "columns": num_cols,
            "rows": num_rows,
            "status": "Table added and styled successfully.",
        }

    def modify_table(
        self,
        table_index: int = 1,
        action: str = "set_cell",
        row: int = 1,
        col: int = 1,
        text: str = "",
        bg_color: Optional[str] = None,
        bold: Optional[bool] = None,
    ) -> Dict[str, Any]:
        """Modify an existing table (action: set_cell, add_row, add_column, delete_row, delete_column)."""
        doc = self.get_document()
        if table_index < 1 or table_index > doc.Tables.Count:
            raise IndexError(f"Table index {table_index} out of range (Total: {doc.Tables.Count})")

        tbl = doc.Tables(table_index)

        with self.undo_block(f"Table {action}"):
            if action == "set_cell":
                cell = tbl.Cell(row, col)
                if text:
                    cell.Range.Text = text
                if bg_color:
                    cell.Shading.BackgroundPatternColor = hex_to_bgr_int(bg_color)
                if bold is not None:
                    cell.Range.Font.Bold = bold

            elif action == "add_row":
                tbl.Rows.Add()

            elif action == "add_column":
                tbl.Columns.Add()

            elif action == "delete_row":
                tbl.Rows(row).Delete()

            elif action == "delete_column":
                tbl.Columns(col).Delete()

            elif action == "autofit":
                tbl.AutoFitBehavior(2)

        return {"table_index": table_index, "action": action, "status": "Table modified."}

    def add_image(
        self,
        image_path: str,
        width_cm: Optional[float] = None,
        height_cm: Optional[float] = None,
        align: str = "center",
        caption: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Add image with optional sizing, alignment, and caption."""
        abs_path = os.path.abspath(image_path)
        if not os.path.isfile(abs_path):
            raise FileNotFoundError(f"Image not found: {abs_path}")

        doc = self.get_document()
        with self.undo_block("Insert Image"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            shape = doc.InlineShapes.AddPicture(FileName=abs_path, Range=sel.Range)

            if width_cm and height_cm:
                shape.LockAspectRatio = False
                shape.Width = cm_to_points(width_cm)
                shape.Height = cm_to_points(height_cm)
            elif width_cm:
                shape.LockAspectRatio = True
                shape.Width = cm_to_points(width_cm)
            elif height_cm:
                shape.LockAspectRatio = True
                shape.Height = cm_to_points(height_cm)

            # Center alignment
            p = shape.Range.ParagraphFormat
            p.Alignment = WD_ALIGN.get(align.lower(), 1)
            p.SpaceAfter = 6

            sel.EndKey(6)
            sel.TypeText("\n")

            if caption:
                sel.TypeText(f"Figure: {caption}\n")
                p_cap = doc.Paragraphs(doc.Paragraphs.Count - 1)
                p_cap.Range.Font.Italic = True
                p_cap.Range.Font.Size = 9.5
                p_cap.Range.Font.Color = hex_to_bgr_int("#64748B")
                p_cap.Alignment = WD_ALIGN.get(align.lower(), 1)
                p_cap.SpaceAfter = 12

        return {"image": abs_path, "status": "Image inserted successfully."}

    def add_page_break(self) -> Dict[str, Any]:
        """Insert a page break."""
        self.get_document()
        with self.undo_block("Page Break"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)
            sel.InsertBreak(WD_BREAKS["page"])
        return {"status": "Page break inserted."}

    def add_section_break(self, break_type: str = "section_next") -> Dict[str, Any]:
        """Insert section break (section_next, section_continuous)."""
        self.get_document()
        b_type = WD_BREAKS.get(break_type.lower(), WD_BREAKS["section_next"])
        with self.undo_block("Section Break"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)
            sel.InsertBreak(b_type)
        return {"type": break_type, "status": "Section break inserted."}

    def set_header_footer(
        self,
        header_text: Optional[str] = None,
        footer_text: Optional[str] = None,
        include_page_number: bool = True,
        different_first_page: bool = False,
    ) -> Dict[str, Any]:
        """Set document header and footer with dynamic page numbers."""
        doc = self.get_document()
        with self.undo_block("Configure Header and Footer"):
            sec = doc.Sections(1)
            sec.PageSetup.DifferentFirstPageHeaderFooter = different_first_page

            # wdHeaderFooterPrimary = 1
            if header_text is not None:
                hdr = sec.Headers(1)
                hdr.Range.Text = header_text
                hdr.Range.Font.Name = "Calibri"
                hdr.Range.Font.Size = 9.0
                hdr.Range.Font.Color = hex_to_bgr_int("#64748B")
                hdr.Range.ParagraphFormat.Alignment = WD_ALIGN["right"]

            ftr = sec.Footers(1)
            ftr.Range.Text = ""
            ftr_rng = ftr.Range
            ftr_rng.Font.Name = "Calibri"
            ftr_rng.Font.Size = 9.0
            ftr_rng.Font.Color = hex_to_bgr_int("#64748B")

            if footer_text:
                ftr_rng.Text = footer_text + ("  |  " if include_page_number else "")

            if include_page_number:
                ftr_rng.Collapse(0)  # wdCollapseEnd
                ftr_rng.InsertAfter("Page ")
                ftr_rng.Collapse(0)
                # wdFieldPage = 33
                doc.Fields.Add(Range=ftr_rng, Type=33)
                ftr_rng = ftr.Range
                ftr_rng.Collapse(0)
                ftr_rng.InsertAfter(" of ")
                ftr_rng.Collapse(0)
                # wdFieldNumPages = 26
                doc.Fields.Add(Range=ftr_rng, Type=26)

            ftr.Range.ParagraphFormat.Alignment = WD_ALIGN["center"]

        return {"status": "Header and footer configured successfully."}

    def add_watermark(
        self,
        text: str = "CONFIDENTIAL",
        font_size: float = 72.0,
        color: str = "#E2E8F0",
        orientation: str = "diagonal",
    ) -> Dict[str, Any]:
        """Add diagonal or horizontal text watermark across document pages."""
        doc = self.get_document()
        with self.undo_block("Add Watermark"):
            sec = doc.Sections(1)
            hdr = sec.Headers(1)
            
            # Remove existing watermark shapes if any
            for i in range(hdr.Shapes.Count, 0, -1):
                s = hdr.Shapes(i)
                if "WaterMark" in s.Name:
                    s.Delete()

            # Add WordArt watermark shape
            # msoTextEffect1 = 0
            shape = hdr.Shapes.AddTextEffect(
                0, text, "Calibri", font_size, False, False, 0, 0
            )
            shape.Name = "PowerWaterMark"
            shape.TextEffect.NormalizedHeight = False
            shape.Line.Visible = False
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = hex_to_bgr_int(color)
            shape.Fill.Transparency = 0.6  # 60% transparent

            if orientation.lower() == "diagonal":
                shape.Rotation = 315
            else:
                shape.Rotation = 0

            # Center shape on page
            shape.RelativeHorizontalPosition = 1  # wdRelativeHorizontalPositionPage
            shape.RelativeVerticalPosition = 1    # wdRelativeVerticalPositionPage
            shape.Left = -999995                  # wdShapeCenter
            shape.Top = -999995                   # wdShapeCenter
            shape.WrapFormat.Type = 3             # wdWrapNone

        return {"text": text, "orientation": orientation, "status": "Watermark added."}

    def remove_watermark(self) -> Dict[str, Any]:
        """Remove watermarks from the document."""
        doc = self.get_document()
        removed = 0
        with self.undo_block("Remove Watermark"):
            for sec in doc.Sections:
                for hdr in sec.Headers:
                    for i in range(hdr.Shapes.Count, 0, -1):
                        s = hdr.Shapes(i)
                        if "WaterMark" in s.Name or "PowerWaterMark" in s.Name:
                            s.Delete()
                            removed += 1
        return {"removed_count": removed, "status": "Watermarks removed."}

    def add_comment(
        self,
        text: str,
        target_text: Optional[str] = None,
        author: str = "AI Assistant",
    ) -> Dict[str, Any]:
        """Add a comment anchored to specific text or current selection."""
        doc = self.get_document()
        with self.undo_block("Add Comment"):
            app = self.get_app(bring_to_front=False)
            rng = None

            if target_text:
                # Search for target text
                find = doc.Content.Find
                find.ClearFormatting()
                find.Text = target_text
                if find.Execute():
                    rng = find.Parent

            if rng is None:
                rng = app.Selection.Range

            comment = doc.Comments.Add(Range=rng, Text=text)
            comment.Author = author

        return {"text": text, "author": author, "status": "Comment added."}

    def list_comments(self) -> List[Dict[str, Any]]:
        """List all comments in active document."""
        doc = self.get_document()
        comments = []
        for i in range(1, doc.Comments.Count + 1):
            c = doc.Comments(i)
            comments.append({
                "index": i,
                "author": c.Author,
                "date": str(c.Date),
                "text": c.Range.Text.strip(),
                "scope_text": c.Scope.Text.strip() if hasattr(c, "Scope") else "",
            })
        return comments

    def reply_comment(self, comment_index: int, reply_text: str, author: str = "AI Assistant") -> Dict[str, Any]:
        """Add a threaded reply to an existing comment (Word 2016+)."""
        doc = self.get_document()
        if comment_index < 1 or comment_index > doc.Comments.Count:
            raise IndexError(f"Comment index {comment_index} not found.")

        with self.undo_block("Reply to Comment"):
            parent = doc.Comments(comment_index)
            if hasattr(parent, "Replies"):
                reply = parent.Replies.Add(Range=parent.Range, Text=reply_text)
                reply.Author = author
            else:
                # Fallback for older Word versions: append text
                parent.Range.InsertAfter(f"\n[{author}]: {reply_text}")

        return {"parent_index": comment_index, "status": "Reply added."}

    def resolve_comment(self, comment_index: int, resolved: bool = True) -> Dict[str, Any]:
        """Mark comment as resolved or open (Word 2016+)."""
        doc = self.get_document()
        if comment_index < 1 or comment_index > doc.Comments.Count:
            raise IndexError(f"Comment index {comment_index} not found.")

        with self.undo_block("Resolve Comment"):
            c = doc.Comments(comment_index)
            if hasattr(c, "Done"):
                c.Done = resolved
            else:
                c.Range.InsertBefore("[RESOLVED] " if resolved else "")

        return {"index": comment_index, "resolved": resolved, "status": "Comment status updated."}

    def delete_comment(self, comment_index: int) -> Dict[str, Any]:
        """Delete specific comment."""
        doc = self.get_document()
        if comment_index < 1 or comment_index > doc.Comments.Count:
            raise IndexError(f"Comment index {comment_index} not found.")

        with self.undo_block("Delete Comment"):
            doc.Comments(comment_index).Delete()

        return {"deleted_index": comment_index, "status": "Comment deleted."}

    def toggle_track_changes(self, enabled: Optional[bool] = None) -> Dict[str, Any]:
        """Enable, disable, or toggle Track Changes mode."""
        doc = self.get_document()
        with self.undo_block("Toggle Track Changes"):
            if enabled is not None:
                doc.TrackRevisions = enabled
            else:
                doc.TrackRevisions = not doc.TrackRevisions

        return {"track_changes": doc.TrackRevisions, "status": "Track changes setting updated."}

    def list_revisions(self) -> List[Dict[str, Any]]:
        """List all tracked changes in document."""
        doc = self.get_document()
        revs = []
        # wdRevisionInsert = 1, wdRevisionDelete = 2, wdRevisionProperty = 3
        rev_types = {1: "insert", 2: "delete", 3: "property_change", 4: "paragraph_number", 5: "display_field"}
        for i in range(1, doc.Revisions.Count + 1):
            r = doc.Revisions(i)
            revs.append({
                "index": i,
                "author": r.Author,
                "date": str(r.Date),
                "type": rev_types.get(r.Type, f"type_{r.Type}"),
                "text": r.Range.Text.strip() if hasattr(r, "Range") else "",
            })
        return revs

    def accept_revisions(self, author: Optional[str] = None) -> Dict[str, Any]:
        """Accept all or author-filtered tracked changes."""
        doc = self.get_document()
        with self.undo_block("Accept Revisions"):
            if author:
                for i in range(doc.Revisions.Count, 0, -1):
                    r = doc.Revisions(i)
                    if r.Author.lower() == author.lower():
                        r.Accept()
            else:
                doc.Revisions.AcceptAll()

        return {"status": "Revisions accepted."}

    def reject_revisions(self, author: Optional[str] = None) -> Dict[str, Any]:
        """Reject all or author-filtered tracked changes."""
        doc = self.get_document()
        with self.undo_block("Reject Revisions"):
            if author:
                for i in range(doc.Revisions.Count, 0, -1):
                    r = doc.Revisions(i)
                    if r.Author.lower() == author.lower():
                        r.Reject()
            else:
                doc.Revisions.RejectAll()

        return {"status": "Revisions rejected."}

    def search_and_replace(
        self,
        find_text: str,
        replace_text: str,
        match_case: bool = False,
        match_whole_word: bool = False,
    ) -> Dict[str, Any]:
        """Search and replace text throughout entire document."""
        doc = self.get_document()
        with self.undo_block(f"Replace '{find_text}'"):
            find = doc.Content.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            # wdReplaceAll = 2
            find.Execute(
                FindText=find_text,
                MatchCase=match_case,
                MatchWholeWord=match_whole_word,
                ReplaceWith=replace_text,
                Replace=2,
            )

        return {"find": find_text, "replace": replace_text, "status": "Search and replace executed."}

    def add_toc(self, title: str = "Table of Contents") -> Dict[str, Any]:
        """Add dynamic Table of Contents (TOC)."""
        doc = self.get_document()
        with self.undo_block("Insert Table of Contents"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            if title:
                sel.TypeText(title + "\n")
                p = doc.Paragraphs(doc.Paragraphs.Count - 1)
                p.Range.Font.Bold = True
                p.Range.Font.Size = 16
                p.Range.Font.Color = hex_to_bgr_int("#1F4E79")
                p.SpaceAfter = 12

            # Add Table of Contents field
            doc.TablesOfContents.Add(
                Range=sel.Range,
                RightAlignPageNumbers=True,
                UseHeadingStyles=True,
                UpperHeadingLevel=1,
                LowerHeadingLevel=3,
                IncludePageNumbers=True,
            )

            sel.EndKey(6)
            sel.TypeText("\n")

        return {"status": "Table of Contents inserted."}

    def update_fields(self) -> Dict[str, Any]:
        """Update all fields in document (TOC, page numbers, cross-refs)."""
        doc = self.get_document()
        with self.undo_block("Update Fields"):
            doc.Fields.Update()
            for toc in doc.TablesOfContents:
                toc.Update()
        return {"status": "All document fields updated."}

    def add_equation(self, equation_text: str) -> Dict[str, Any]:
        """Insert mathematical equation using UnicodeMath / LaTeX style."""
        doc = self.get_document()
        with self.undo_block("Insert Equation"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            rng = sel.Range
            rng.Text = equation_text
            # Convert text range to OMath equation
            omaths = doc.OMaths.Add(rng)
            omaths(1).BuildUp()

            sel.EndKey(6)
            sel.TypeText("\n")

        return {"equation": equation_text, "status": "Equation inserted."}

    def add_footnote(self, note_text: str, target_text: Optional[str] = None) -> Dict[str, Any]:
        """Insert footnote anchored to target text or end of document."""
        doc = self.get_document()
        with self.undo_block("Insert Footnote"):
            app = self.get_app(bring_to_front=False)
            rng = None
            if target_text:
                find = doc.Content.Find
                find.ClearFormatting()
                find.Text = target_text
                if find.Execute():
                    rng = find.Parent

            if rng is None:
                sel = app.Selection
                sel.EndKey(6)
                rng = sel.Range

            fn = doc.Footnotes.Add(Range=rng, Text=note_text)

        return {"footnote_text": note_text, "status": "Footnote added."}

    def add_hyperlink(self, text: str, url: str) -> Dict[str, Any]:
        """Insert clickable hyperlink with URL."""
        doc = self.get_document()
        with self.undo_block("Insert Hyperlink"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            doc.Hyperlinks.Add(Anchor=sel.Range, Address=url, TextToDisplay=text)

            sel.EndKey(6)
            sel.TypeText(" ")

        return {"text": text, "url": url, "status": "Hyperlink added."}

    def add_bookmark(self, name: str, target_text: Optional[str] = None) -> Dict[str, Any]:
        """Insert a named bookmark."""
        doc = self.get_document()
        with self.undo_block(f"Add Bookmark '{name}'"):
            app = self.get_app(bring_to_front=False)
            rng = None
            if target_text:
                find = doc.Content.Find
                find.ClearFormatting()
                find.Text = target_text
                if find.Execute():
                    rng = find.Parent

            if rng is None:
                sel = app.Selection
                sel.EndKey(6)
                rng = sel.Range

            doc.Bookmarks.Add(Name=name, Range=rng)

        return {"name": name, "status": "Bookmark added."}

    def read_document_text(self, max_paragraphs: int = 200) -> Dict[str, Any]:
        """Read text content and structure from active document."""
        doc = self.get_document()
        total_p = doc.Paragraphs.Count
        limit = min(total_p, max_paragraphs)
        paras = []

        for i in range(1, limit + 1):
            p = doc.Paragraphs(i)
            txt = p.Range.Text.strip()
            if txt:
                paras.append({
                    "index": i,
                    "text": txt,
                    "style": str(p.Style.NameLocal) if hasattr(p, "Style") else "Normal",
                })

        return {
            "document_name": doc.Name,
            "total_paragraphs": total_p,
            "read_count": len(paras),
            "paragraphs": paras,
        }

    def read_outline(self) -> List[Dict[str, Any]]:
        """Extract heading outline hierarchy."""
        doc = self.get_document()
        outline = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            style_name = str(p.Style.NameLocal) if hasattr(p, "Style") else ""
            if "Heading" in style_name or "Titre" in style_name:
                txt = p.Range.Text.strip()
                if txt:
                    outline.append({
                        "index": i,
                        "style": style_name,
                        "text": txt,
                    })
        return outline

    def undo(self, count: int = 1) -> Dict[str, Any]:
        """Undo last operations in Word."""
        doc = self.get_document()
        doc.Undo(count)
        return {"undo_count": count, "status": f"Undid last {count} operation(s)."}

    def take_screenshot(self, output_path: Optional[str] = None) -> Dict[str, Any]:
        """Capture screenshot of the Word window."""
        try:
            from PIL import ImageGrab
            if not output_path:
                output_path = os.path.abspath(f"word_screenshot_{int(time.time())}.png")
            else:
                output_path = os.path.abspath(output_path)

            self.focus_window()
            time.sleep(0.3)

            # Get Word window rect
            if win32gui:
                hwnd = win32gui.FindWindow("OpusApp", None)
                if hwnd:
                    rect = win32gui.GetWindowRect(hwnd)
                    img = ImageGrab.grab(bbox=rect)
                else:
                    img = ImageGrab.grab()
            else:
                img = ImageGrab.grab()

            img.save(output_path)
            return {"screenshot_path": output_path, "status": "Screenshot captured successfully."}
        except Exception as e:
            return {"error": str(e), "status": "Failed to capture screenshot."}


    def add_cover_page(
        self,
        title: str,
        subtitle: Optional[str] = None,
        author: Optional[str] = None,
        organization: Optional[str] = None,
        date_str: Optional[str] = None,
        theme_color: str = "#1F4E79",
    ) -> Dict[str, Any]:
        """Insert an executive cover/title page at the start of document."""
        doc = self.get_document()
        with self.undo_block("Add Cover Page"):
            app = self.get_app(bring_to_front=False)
            sec = doc.Sections(1)
            sec.PageSetup.DifferentFirstPageHeaderFooter = True

            sel = app.Selection
            sel.HomeKey(6)  # Jump to start of document (wdStory)

            # Top accent bar spacing
            sel.TypeText("\n\n\n")

            # Document Title
            sel.TypeText(title + "\n")
            p_title = doc.Paragraphs(doc.Paragraphs.Count - 1 if doc.Paragraphs.Count > 1 else 1)
            # Find the title paragraph by range
            rng_title = sel.Range
            rng_title.MoveStart(4, -1)  # wdParagraph = 4
            p_title.Range.Font.Name = "Calibri"
            p_title.Range.Font.Size = 30.0
            p_title.Range.Font.Bold = True
            p_title.Range.Font.Color = hex_to_bgr_int(theme_color)
            p_title.SpaceAfter = 12

            if subtitle:
                sel.TypeText(subtitle + "\n")
                p_sub = doc.Paragraphs(p_title.Range.Paragraphs(1).Range.Paragraphs.Count + 1)
                p_sub.Range.Font.Name = "Calibri"
                p_sub.Range.Font.Size = 14.0
                p_sub.Range.Font.Color = hex_to_bgr_int("#64748B")
                p_sub.SpaceAfter = 200  # Push metadata to bottom

            sel.TypeText("\n\n")

            # Metadata block
            meta_lines = []
            if author:
                meta_lines.append(f"Author: {author}")
            if organization:
                meta_lines.append(f"Organization: {organization}")
            if date_str:
                meta_lines.append(f"Date: {date_str}")
            else:
                meta_lines.append(f"Date: {time.strftime('%B %d, %Y')}")

            if meta_lines:
                meta_text = "\n".join(meta_lines) + "\n"
                sel.TypeText(meta_text)
                for _ in range(len(meta_lines)):
                    p_meta = doc.Paragraphs(doc.Paragraphs.Count - 1)
                    p_meta.Range.Font.Name = "Calibri"
                    p_meta.Range.Font.Size = 10.5
                    p_meta.Range.Font.Color = hex_to_bgr_int("#475569")
                    p_meta.SpaceAfter = 4

            # Page break into body
            sel.InsertBreak(WD_BREAKS["page"])

        return {"title": title, "status": "Cover page created successfully."}

    def set_page_setup(
        self,
        orientation: str = "portrait",
        margins: str = "normal",
        paper_size: str = "a4",
    ) -> Dict[str, Any]:
        """Configure page margins, orientation, and size for active section."""
        doc = self.get_document()
        with self.undo_block("Configure Page Setup"):
            sec = doc.Sections(1)
            ps = sec.PageSetup

            # Orientation
            if orientation.lower() == "landscape":
                ps.Orientation = 1  # wdOrientLandscape
            else:
                ps.Orientation = 0  # wdOrientPortrait

            # Margins
            margin_presets = {
                "normal": (72.0, 72.0, 72.0, 72.0),
                "narrow": (36.0, 36.0, 36.0, 36.0),
                "moderate": (72.0, 72.0, 54.0, 54.0),
                "wide": (72.0, 72.0, 144.0, 144.0),
            }
            top, bottom, left, right = margin_presets.get(margins.lower(), margin_presets["normal"])
            ps.TopMargin = top
            ps.BottomMargin = bottom
            ps.LeftMargin = left
            ps.RightMargin = right

            # Paper Size
            paper_map = {
                "a4": (595.3, 841.9),
                "letter": (612.0, 792.0),
                "legal": (612.0, 1008.0),
            }
            w, h = paper_map.get(paper_size.lower(), paper_map["a4"])
            if orientation.lower() == "landscape":
                ps.PageWidth = h
                ps.PageHeight = w
            else:
                ps.PageWidth = w
                ps.PageHeight = h

        return {
            "orientation": orientation,
            "margins": margins,
            "paper_size": paper_size,
            "status": "Page setup updated.",
        }

    def set_columns(
        self,
        column_count: int = 2,
        spacing_cm: float = 1.25,
        line_between: bool = False,
    ) -> Dict[str, Any]:
        """Configure multi-column layout for the current document section."""
        doc = self.get_document()
        with self.undo_block("Configure Columns"):
            sec = doc.Sections(1)
            cols = sec.PageSetup.TextColumns
            cols.SetCount(column_count)
            cols.Spacing = cm_to_points(spacing_cm)
            cols.LineBetween = line_between

        return {"columns": column_count, "spacing_cm": spacing_cm, "status": "Columns configured."}

    def add_quote_card(
        self,
        quote: str,
        author: Optional[str] = None,
        title: Optional[str] = None,
        border_color: str = "#2563EB",
        bg_color: str = "#F8FAFC",
    ) -> Dict[str, Any]:
        """Insert a stylish pull-quote / testimonial card."""
        doc = self.get_document()
        with self.undo_block("Add Quote Card"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            tbl = doc.Tables.Add(Range=sel.Range, NumRows=1, NumColumns=1)
            tbl.Borders.Enable = False

            cell = tbl.Cell(1, 1)
            left_border = cell.Borders(-2)  # wdBorderLeft
            left_border.LineStyle = 1
            left_border.LineWidth = 36  # 4.5 pt
            left_border.Color = hex_to_bgr_int(border_color)

            cell.Shading.BackgroundPatternColor = hex_to_bgr_int(bg_color)
            tbl.TopPadding = 10
            tbl.BottomPadding = 10
            tbl.LeftPadding = 14
            tbl.RightPadding = 14

            c_rng = cell.Range
            attribution = f"\n\n― {author}" + (f", {title}" if title else "") if author else ""
            c_rng.Text = f"“{quote}”{attribution}"

            # Format quote text
            p_q = c_rng.Paragraphs(1)
            p_q.Range.Font.Italic = True
            p_q.Range.Font.Size = 11.5
            p_q.Range.Font.Color = hex_to_bgr_int("#1E293B")

            if author and c_rng.Paragraphs.Count > 1:
                p_a = c_rng.Paragraphs(c_rng.Paragraphs.Count)
                p_a.Range.Font.Italic = False
                p_a.Range.Font.Bold = True
                p_a.Range.Font.Size = 10.0
                p_a.Range.Font.Color = hex_to_bgr_int("#64748B")

            sel.EndKey(6)
            sel.TypeText("\n")

        return {"quote": quote, "author": author, "status": "Quote card added."}

    def add_kpi_cards(
        self,
        metrics: List[Dict[str, str]],
        header_color: str = "#1F4E79",
        bg_color: str = "#F8FAFC",
    ) -> Dict[str, Any]:
        """Insert a horizontal grid of modern KPI / stat metric cards."""
        doc = self.get_document()
        num_cards = len(metrics)
        if num_cards == 0:
            return {"status": "No metrics provided."}

        with self.undo_block("Add KPI Cards Grid"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            tbl = doc.Tables.Add(Range=sel.Range, NumRows=1, NumColumns=num_cards)
            tbl.Borders.Enable = True
            tbl.Borders.InsideColor = hex_to_bgr_int("#E2E8F0")
            tbl.Borders.OutsideColor = hex_to_bgr_int("#CBD5E1")
            tbl.AutoFitBehavior(2)  # wdAutoFitWindow

            for col_idx, item in enumerate(metrics, start=1):
                cell = tbl.Cell(1, col_idx)
                cell.Shading.BackgroundPatternColor = hex_to_bgr_int(bg_color)
                tbl.TopPadding = 8
                tbl.BottomPadding = 8
                tbl.LeftPadding = 8
                tbl.RightPadding = 8

                val = item.get("value", "")
                lbl = item.get("label", "").upper()
                chg = item.get("change", "")

                text_content = f"{val}\n{lbl}" + (f"\n{chg}" if chg else "")
                cell.Range.Text = text_content
                cell.VerticalAlignment = 1  # Center

                # Stat number line
                p_val = cell.Range.Paragraphs(1)
                p_val.Range.Font.Bold = True
                p_val.Range.Font.Size = 18.0
                p_val.Range.Font.Color = hex_to_bgr_int(header_color)
                p_val.Alignment = WD_ALIGN["center"]
                p_val.SpaceAfter = 2

                # Label line
                if cell.Range.Paragraphs.Count >= 2:
                    p_lbl = cell.Range.Paragraphs(2)
                    p_lbl.Range.Font.Bold = True
                    p_lbl.Range.Font.Size = 8.5
                    p_lbl.Range.Font.Color = hex_to_bgr_int("#64748B")
                    p_lbl.Alignment = WD_ALIGN["center"]
                    p_lbl.SpaceAfter = 2

                # Change line
                if cell.Range.Paragraphs.Count >= 3 and chg:
                    p_chg = cell.Range.Paragraphs(3)
                    p_chg.Range.Font.Size = 9.0
                    p_chg.Range.Font.Bold = True
                    p_chg.Range.Font.Color = hex_to_bgr_int("#059669" if "+" in chg else "#DC2626")
                    p_chg.Alignment = WD_ALIGN["center"]

            sel.EndKey(6)
            sel.TypeText("\n")

        return {"cards_count": num_cards, "status": "KPI cards added successfully."}

    def add_process_steps(
        self,
        steps: List[Dict[str, str]],
        accent_color: str = "#2563EB",
    ) -> Dict[str, Any]:
        """Insert a sequential step-by-step roadmap / process flow."""
        doc = self.get_document()
        with self.undo_block("Add Process Steps"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            tbl = doc.Tables.Add(Range=sel.Range, NumRows=len(steps), NumColumns=2)
            tbl.Borders.Enable = False
            tbl.AutoFitBehavior(2)
            # Column 1 narrow (badge), Column 2 wide (description)
            tbl.Columns(1).Width = cm_to_points(1.8)
            tbl.Columns(2).Width = cm_to_points(14.5)

            for idx, step_info in enumerate(steps, start=1):
                num = step_info.get("step", idx)
                title = step_info.get("title", f"Step {idx}")
                desc = step_info.get("description", "")

                # Badge cell
                c1 = tbl.Cell(idx, 1)
                c1.Range.Text = f"0{num}" if int(num) < 10 else f"{num}"
                c1.Range.Font.Bold = True
                c1.Range.Font.Size = 14.0
                c1.Range.Font.Color = hex_to_bgr_int(accent_color)
                c1.VerticalAlignment = 1
                c1.Range.ParagraphFormat.Alignment = WD_ALIGN["center"]

                # Text cell
                c2 = tbl.Cell(idx, 2)
                c2.Range.Text = f"{title}\n{desc}"
                p_t = c2.Range.Paragraphs(1)
                p_t.Range.Font.Bold = True
                p_t.Range.Font.Size = 11.0
                p_t.Range.Font.Color = hex_to_bgr_int("#1E293B")
                p_t.SpaceAfter = 2

                if c2.Range.Paragraphs.Count > 1:
                    p_d = c2.Range.Paragraphs(2)
                    p_d.Range.Font.Size = 10.0
                    p_d.Range.Font.Color = hex_to_bgr_int("#475569")
                    p_d.SpaceAfter = 8

            sel.EndKey(6)
            sel.TypeText("\n")

        return {"steps_count": len(steps), "status": "Process steps inserted."}

    def add_endnote(self, note_text: str, target_text: Optional[str] = None) -> Dict[str, Any]:
        """Insert an endnote into the document."""
        doc = self.get_document()
        with self.undo_block("Insert Endnote"):
            app = self.get_app(bring_to_front=False)
            rng = None
            if target_text:
                find = doc.Content.Find
                find.ClearFormatting()
                find.Text = target_text
                if find.Execute():
                    rng = find.Parent

            if rng is None:
                sel = app.Selection
                sel.EndKey(6)
                rng = sel.Range

            doc.Endnotes.Add(Range=rng, Text=note_text)

        return {"endnote": note_text, "status": "Endnote added successfully."}

    def list_endnotes(self) -> List[Dict[str, Any]]:
        """List all endnotes in active document."""
        doc = self.get_document()
        endnotes = []
        for i in range(1, doc.Endnotes.Count + 1):
            en = doc.Endnotes(i)
            endnotes.append({
                "index": i,
                "text": en.Range.Text.strip(),
            })
        return endnotes

    def set_protection(
        self,
        protection_type: str = "revisions",
        password: str = "",
    ) -> Dict[str, Any]:
        """Protect document against unauthorized edits (revisions, comments, readonly)."""
        doc = self.get_document()
        with self.undo_block("Protect Document"):
            # wdAllowOnlyRevisions = 0, wdAllowOnlyComments = 1, wdAllowOnlyReading = 3
            p_map = {
                "revisions": 0,
                "comments": 1,
                "readonly": 3,
            }
            p_const = p_map.get(protection_type.lower(), 0)
            doc.Protect(Type=p_const, NoReset=False, Password=password)

        return {"protection": protection_type, "status": "Document protected successfully."}

    def remove_protection(self, password: str = "") -> Dict[str, Any]:
        """Remove document protection."""
        doc = self.get_document()
        doc.Unprotect(Password=password)
        return {"status": "Document protection removed."}

    def highlight_text(self, find_text: str, highlight_color: str = "yellow") -> Dict[str, Any]:
        """Find and highlight all occurrences of target text throughout document."""
        doc = self.get_document()
        hl_map = {"yellow": 7, "green": 4, "blue": 3, "pink": 6}
        hl_code = hl_map.get(highlight_color.lower(), 7)

        count = 0
        with self.undo_block(f"Highlight '{find_text}'"):
            find = doc.Content.Find
            find.ClearFormatting()
            find.Text = find_text
            while find.Execute():
                find.Parent.HighlightColorIndex = hl_code
                count += 1

        return {"text": find_text, "color": highlight_color, "highlighted_count": count, "status": "Highlighting applied."}

    def format_text(
        self,
        find_text: str,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[str] = None,
        font_size: Optional[float] = None,
    ) -> Dict[str, Any]:
        """Find and format occurrences of target text (bold, italic, color, size)."""
        doc = self.get_document()
        count = 0
        with self.undo_block(f"Format '{find_text}'"):
            find = doc.Content.Find
            find.ClearFormatting()
            find.Text = find_text
            while find.Execute():
                rng = find.Parent
                if bold is not None:
                    rng.Font.Bold = bold
                if italic is not None:
                    rng.Font.Italic = italic
                if color:
                    rng.Font.Color = hex_to_bgr_int(color)
                if font_size:
                    rng.Font.Size = font_size
                count += 1

        return {"text": find_text, "formatted_count": count, "status": "Formatting applied."}

    def add_content_control(
        self,
        title: str,
        tag: Optional[str] = None,
        control_type: str = "rich_text",
        placeholder: str = "Click here to enter text.",
    ) -> Dict[str, Any]:
        """Insert a content control form field (rich_text, plain_text, date)."""
        doc = self.get_document()
        with self.undo_block(f"Add Content Control '{title}'"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            # wdContentControlRichText = 0, wdContentControlText = 1, wdContentControlDate = 6
            t_map = {"rich_text": 0, "plain_text": 1, "date": 6}
            cc_type = t_map.get(control_type.lower(), 0)

            cc = doc.ContentControls.Add(Type=cc_type, Range=sel.Range)
            cc.Title = title
            cc.Tag = tag or title
            cc.SetPlaceholderText(Text=placeholder)

            sel.EndKey(6)
            sel.TypeText("\n")

        return {"title": title, "type": control_type, "status": "Content control added."}

    def list_content_controls(self) -> List[Dict[str, Any]]:
        """List all content controls in active document."""
        doc = self.get_document()
        controls = []
        for i in range(1, doc.ContentControls.Count + 1):
            cc = doc.ContentControls(i)
            controls.append({
                "index": i,
                "title": cc.Title,
                "tag": cc.Tag,
                "text": cc.Range.Text.strip(),
            })
        return controls

    def add_chart(
        self,
        chart_type: str = "column",
        title: str = "Performance Overview",
    ) -> Dict[str, Any]:
        """Insert a native Microsoft Office chart (column, bar, line, pie)."""
        doc = self.get_document()
        with self.undo_block("Insert Chart"):
            app = self.get_app(bring_to_front=False)
            sel = app.Selection
            sel.EndKey(6)

            # xlColumnClustered = 51, xlBarClustered = 57, xlLine = 4, xlPie = 5
            c_map = {"column": 51, "bar": 57, "line": 4, "pie": 5}
            c_type = c_map.get(chart_type.lower(), 51)

            # Insert chart via Word COM AddChart2
            shape = doc.InlineShapes.AddChart2(Style=-1, Type=c_type, Range=sel.Range)
            if hasattr(shape, "Chart"):
                chart = shape.Chart
                chart.HasTitle = True
                chart.ChartTitle.Text = title

            sel.EndKey(6)
            sel.TypeText("\n")

        return {"chart_type": chart_type, "title": title, "status": "Chart inserted successfully."}

    def merge_documents(self, file_paths: List[str], output_path: Optional[str] = None) -> Dict[str, Any]:
        """Merge multiple Word documents into one combined document."""
        if not file_paths:
            raise ValueError("No file paths provided.")

        app = self.get_app()
        # Open first document as base
        first_doc = os.path.abspath(file_paths[0])
        doc = app.Documents.Open(first_doc)

        sel = app.Selection
        for next_file in file_paths[1:]:
            abs_next = os.path.abspath(next_file)
            if os.path.isfile(abs_next):
                sel.EndKey(6)
                sel.InsertBreak(WD_BREAKS["section_next"])
                sel.InsertFile(FileName=abs_next)

        if output_path:
            abs_out = os.path.abspath(output_path)
            doc.SaveAs2(abs_out)
            return {"merged_count": len(file_paths), "output_path": abs_out, "status": "Documents merged successfully."}

        return {"merged_count": len(file_paths), "status": "Documents merged in active window."}

    def redact_text(self, target_text: str, replacement: str = "[REDACTED]") -> Dict[str, Any]:
        """Redact sensitive text across document with black highlight / replacement."""
        doc = self.get_document()
        count = 0
        with self.undo_block(f"Redact '{target_text}'"):
            find = doc.Content.Find
            find.ClearFormatting()
            find.Text = target_text
            while find.Execute():
                rng = find.Parent
                rng.Text = replacement
                # Black highlight / bold red
                rng.Font.Bold = True
                rng.Font.Color = hex_to_bgr_int("#FFFFFF")
                rng.HighlightColorIndex = 1  # wdBlack = 1
                count += 1

        return {"target": target_text, "redacted_count": count, "status": "Text redacted successfully."}

    def add_bates_numbering(
        self,
        prefix: str = "CONF-",
        start_number: int = 1,
        digits: int = 6,
    ) -> Dict[str, Any]:
        """Add legal Bates stamping numbering to document footer."""
        doc = self.get_document()
        with self.undo_block("Add Bates Numbering"):
            sec = doc.Sections(1)
            ftr = sec.Footers(1)
            ftr_rng = ftr.Range
            ftr_rng.Collapse(0)  # wdCollapseEnd
            pad_start = str(start_number).zfill(digits)
            ftr_rng.InsertAfter(f"  |  Bates: {prefix}{pad_start}-")
            ftr_rng.Collapse(0)
            doc.Fields.Add(Range=ftr_rng, Type=33)  # wdFieldPage
            ftr.Range.Font.Size = 8.5
            ftr.Range.Font.Color = hex_to_bgr_int("#475569")

        return {"prefix": prefix, "start_number": start_number, "status": "Bates numbering applied."}


# Global singleton instance
controller = WordController()

