"""PowerPoint COM Automation Controller
Allows real-time, visible control of Microsoft PowerPoint desktop on Windows.
"""

import os
import time
from typing import Any, Dict, List, Optional
import win32com.client
from win32com.client import constants

# COM Constants mapping
SHAPE_TYPES = {
    "rectangle": 1,          # msoShapeRectangle
    "rounded_rectangle": 5,  # msoShapeRoundedRectangle
    "oval": 9,               # msoShapeOval
    "circle": 9,             # msoShapeOval
    "diamond": 4,            # msoShapeDiamond
    "right_arrow": 13,       # msoShapeRightArrow
    "left_arrow": 14,        # msoShapeLeftArrow
    "up_arrow": 15,          # msoShapeUpArrow
    "down_arrow": 16,        # msoShapeDownArrow
    "star": 92,              # msoShape5pointStar
    "hexagon": 10,           # msoShapeHexagon
    "heart": 21,             # msoShapeHeart
    "cloud": 139,            # msoShapeCloud
}

LAYOUT_TYPES = {
    "title": 1,              # ppLayoutTitle
    "title_and_content": 2,  # ppLayoutText
    "two_content": 3,        # ppLayoutTwoColumnText
    "table": 4,              # ppLayoutTable
    "title_only": 11,        # ppLayoutTitleOnly
    "blank": 12,             # ppLayoutBlank
}

ALIGN_TYPES = {
    "left": 1,               # ppAlignLeft
    "center": 2,             # ppAlignCenter
    "right": 3,              # ppAlignRight
    "justify": 4,            # ppAlignJustify
}

ANIMATION_EFFECTS = {
    "appear": 1,             # msoAnimEffectAppear
    "fly_in": 2,             # msoAnimEffectFly
    "dissolve": 8,           # msoAnimEffectDissolve
    "fade": 10,              # msoAnimEffectFade
    "wipe": 22,              # msoAnimEffectWipe
    "zoom": 23,              # msoAnimEffectZoom
    "spin": 28,              # msoAnimEffectSpin
    "bounce": 36,            # msoAnimEffectBounce
    "rise_up": 41,           # msoAnimEffectRiseUp
}

TRANSITION_EFFECTS = {
    "cut": 257,
    "random": 513,
    "checkerboard": 1025,
    "cover": 1284,
    "dissolve": 1537,
    "fade": 1793,
    "uncover": 2052,
    "wipe": 2561,
    "push": 3844,
}


def hex_to_rgb_int(hex_str: str) -> int:
    """Converts a hex color (#RRGGBB or #RGB) to COM BGR/RGB integer."""
    hex_str = hex_str.strip().lstrip("#")
    if len(hex_str) == 3:
        hex_str = "".join([c * 2 for c in hex_str])
    if len(hex_str) != 6:
        return 0  # fallback to black
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    # PowerPoint COM expects RGB as integer: R + (G * 256) + (B * 65536)
    return r + (g * 256) + (b * 65536)


class PowerPointController:
    """Manages live interaction with Microsoft PowerPoint via COM."""

    def __init__(self):
        self._app = None
        self._pres = None

    def get_app(self, bring_to_front: bool = True):
        """Connect to running PowerPoint or launch a new visible instance."""
        try:
            self._app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            self._app = win32com.client.Dispatch("PowerPoint.Application")

        self._app.Visible = 1  # msoTrue
        try:
            self._app.WindowState = 3  # ppWindowMaximized
            self._app.Activate()
        except Exception:
            pass

        if bring_to_front:
            self.focus_window()

        return self._app

    def focus_window(self):
        """Bring PowerPoint window to foreground so the user watches live changes."""
        try:
            import win32gui
            hwnd = win32gui.FindWindow("PPTFrameClass", None)
            if hwnd:
                win32gui.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def get_presentation(self):
        """Get the current active presentation, or create one if none exists."""
        app = self.get_app()
        try:
            if app.Presentations.Count > 0:
                self._pres = app.ActivePresentation
                return self._pres
        except Exception:
            pass

        # If no presentation open, create a new one
        self._pres = app.Presentations.Add(1)  # WithWindow = 1
        return self._pres

    def new_presentation(self, width: float = 960.0, height: float = 540.0) -> Dict[str, Any]:
        """Create a new blank presentation with 16:9 widescreen layout."""
        app = self.get_app()
        self._pres = app.Presentations.Add(1)
        self._pres.PageSetup.SlideWidth = width
        self._pres.PageSetup.SlideHeight = height
        self.focus_window()
        return {
            "status": "success",
            "message": "New presentation created",
            "slide_width": width,
            "slide_height": height,
            "slide_count": self._pres.Slides.Count,
        }

    def open_presentation(self, file_path: str) -> Dict[str, Any]:
        """Open an existing PowerPoint presentation file."""
        app = self.get_app()
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        self._pres = app.Presentations.Open(abs_path, 0, 0, 1)  # ReadOnly=0, Untitled=0, WithWindow=1
        self.focus_window()
        return {
            "status": "success",
            "message": f"Opened presentation: {os.path.basename(abs_path)}",
            "path": abs_path,
            "slide_count": self._pres.Slides.Count,
        }

    def save_presentation(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Save active presentation as .pptx or export to .pdf."""
        pres = self.get_presentation()
        if file_path:
            abs_path = os.path.abspath(file_path)
            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
            if abs_path.lower().endswith(".pdf"):
                # ppSaveAsPDF = 32
                pres.SaveAs(abs_path, 32)
                return {"status": "success", "message": f"Exported to PDF: {abs_path}", "path": abs_path}
            else:
                pres.SaveAs(abs_path)
                return {"status": "success", "message": f"Saved presentation: {abs_path}", "path": abs_path}
        else:
            pres.Save()
            return {"status": "success", "message": "Presentation saved", "path": pres.FullName}

    def close_presentation(self, save: bool = True) -> Dict[str, Any]:
        """Close active presentation cleanly."""
        pres = self.get_presentation()
        if not save:
            pres.Saved = 1  # Avoid dirty prompt
        else:
            try:
                pres.Save()
            except Exception:
                pass
        pres.Close()
        self._pres = None
        return {"status": "success", "message": "Presentation closed"}

    def get_info(self) -> Dict[str, Any]:
        """Get status of PowerPoint and active presentation."""
        app = self.get_app(bring_to_front=False)
        pres_count = app.Presentations.Count
        if pres_count == 0:
            return {
                "powerpoint_running": True,
                "presentations_open": 0,
                "active_presentation": None,
            }

        pres = app.ActivePresentation
        slides_info = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            title = ""
            for s in slide.Shapes:
                if s.HasTextFrame and s.TextFrame.HasText:
                    text = s.TextFrame.TextRange.Text.strip()
                    if text:
                        title = text.split("\n")[0][:60]
                        break
            slides_info.append({
                "slide_index": i,
                "shape_count": slide.Shapes.Count,
                "preview_title": title,
            })

        active_slide_idx = None
        try:
            active_slide_idx = app.ActiveWindow.View.Slide.SlideIndex
        except Exception:
            pass

        return {
            "powerpoint_running": True,
            "presentations_open": pres_count,
            "active_presentation": pres.Name,
            "slide_width": pres.PageSetup.SlideWidth,
            "slide_height": pres.PageSetup.SlideHeight,
            "slide_count": pres.Slides.Count,
            "active_slide_index": active_slide_idx,
            "slides": slides_info,
        }

    def goto_slide(self, slide_index: int) -> Dict[str, Any]:
        """Navigate view to a specific slide (1-indexed)."""
        pres = self.get_presentation()
        if slide_index < 1 or slide_index > pres.Slides.Count:
            raise ValueError(f"Slide index {slide_index} out of range (1..{pres.Slides.Count})")

        self._app.ActiveWindow.View.GotoSlide(slide_index)
        self.focus_window()
        return {"status": "success", "active_slide_index": slide_index}

    def add_slide(self, layout: str = "blank", position: Optional[int] = None) -> Dict[str, Any]:
        """Add a new slide and automatically jump view to it."""
        pres = self.get_presentation()
        layout_code = LAYOUT_TYPES.get(layout.lower(), 12)  # default blank

        if position is None or position > pres.Slides.Count:
            position = pres.Slides.Count + 1

        new_slide = pres.Slides.Add(position, layout_code)
        # Visually navigate to the new slide
        self._app.ActiveWindow.View.GotoSlide(new_slide.SlideIndex)
        self.focus_window()

        return {
            "status": "success",
            "message": f"Added slide at index {new_slide.SlideIndex}",
            "slide_index": new_slide.SlideIndex,
            "total_slides": pres.Slides.Count,
            "layout": layout,
        }

    def delete_slide(self, slide_index: int) -> Dict[str, Any]:
        """Delete a slide by index."""
        pres = self.get_presentation()
        if slide_index < 1 or slide_index > pres.Slides.Count:
            raise ValueError(f"Slide index {slide_index} out of range (1..{pres.Slides.Count})")

        pres.Slides(slide_index).Delete()
        return {
            "status": "success",
            "message": f"Deleted slide {slide_index}",
            "remaining_slides": pres.Slides.Count,
        }

    def set_slide_background(
        self,
        slide_index: int,
        color_hex: str,
        gradient_color_hex: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Set solid or gradient background color for a slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        rgb1 = hex_to_rgb_int(color_hex)
        slide.FollowMasterBackground = 0  # msoFalse

        if gradient_color_hex:
            rgb2 = hex_to_rgb_int(gradient_color_hex)
            # msoGradientHorizontal = 1, Style 1, Variant 1
            slide.Background.Fill.TwoColorGradient(1, 1)
            slide.Background.Fill.ForeColor.RGB = rgb1
            slide.Background.Fill.BackColor.RGB = rgb2
        else:
            slide.Background.Fill.Solid()
            slide.Background.Fill.ForeColor.RGB = rgb1

        return {
            "status": "success",
            "slide_index": slide_index,
            "color": color_hex,
            "gradient": gradient_color_hex,
        }

    def add_textbox(
        self,
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
        """Add a formatted text box to a slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        shape = slide.Shapes.AddTextbox(1, left, top, width, height)  # 1 = msoTextOrientationHorizontal
        tf = shape.TextFrame
        tf.WordWrap = -1  # msoTrue
        tf.MarginLeft = 10
        tf.MarginRight = 10
        tf.MarginTop = 8
        tf.MarginBottom = 8

        tr = tf.TextRange
        tr.Text = text
        tr.Font.Name = font_name
        tr.Font.Size = font_size
        tr.Font.Bold = 1 if bold else 0
        tr.Font.Italic = 1 if italic else 0
        tr.Font.Color.RGB = hex_to_rgb_int(color_hex)

        align_code = ALIGN_TYPES.get(alignment.lower(), 1)
        tr.ParagraphFormat.Alignment = align_code

        if bg_color_hex:
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)
        else:
            shape.Fill.Background()

        return {
            "status": "success",
            "slide_index": slide_index,
            "shape_name": shape.Name,
            "shape_id": shape.Id,
        }

    def add_card(
        self,
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
        """Creates a modern UI Card (container + header + description) on a slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # 1. Background Card shape (rounded rectangle = 5)
        card = slide.Shapes.AddShape(5, left, top, width, height)
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)
        if border_color_hex:
            card.Line.Visible = -1
            card.Line.ForeColor.RGB = hex_to_rgb_int(border_color_hex)
            card.Line.Weight = 1.5
        else:
            card.Line.Visible = 0

        # Optional accent top bar
        if accent_bar_color_hex:
            accent = slide.Shapes.AddShape(1, left + 4, top + 2, width - 8, 5)
            accent.Fill.Solid()
            accent.Fill.ForeColor.RGB = hex_to_rgb_int(accent_bar_color_hex)
            accent.Line.Visible = 0

        # 2. Text Frame inside card
        tf = card.TextFrame
        tf.WordWrap = -1
        tf.MarginLeft = 16
        tf.MarginRight = 16
        tf.MarginTop = 18
        tf.MarginBottom = 16

        tr = tf.TextRange
        tr.Text = f"{title}\n{body}"
        tr.ParagraphFormat.Alignment = 1  # Left

        # Format title line
        p1 = tr.Paragraphs(1)
        p1.Font.Name = "Segoe UI"
        p1.Font.Size = 18
        p1.Font.Bold = 1
        p1.Font.Color.RGB = hex_to_rgb_int(title_color_hex)

        # Format body line(s)
        if tr.Paragraphs().Count > 1:
            for p_idx in range(2, tr.Paragraphs().Count + 1):
                p = tr.Paragraphs(p_idx)
                p.Font.Name = "Segoe UI"
                p.Font.Size = 13
                p.Font.Bold = 0
                p.Font.Color.RGB = hex_to_rgb_int(body_color_hex)

        return {
            "status": "success",
            "slide_index": slide_index,
            "title": title,
            "card_id": card.Id,
        }

    def add_shape(
        self,
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
        """Add any auto-shape (rectangle, oval, arrow, star, etc.) with fill and text."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        shape_code = SHAPE_TYPES.get(shape_type.lower(), 1)
        shape = slide.Shapes.AddShape(shape_code, left, top, width, height)

        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = hex_to_rgb_int(fill_color_hex)

        if line_color_hex:
            shape.Line.Visible = -1
            shape.Line.ForeColor.RGB = hex_to_rgb_int(line_color_hex)
            shape.Line.Weight = line_width
        else:
            shape.Line.Visible = 0

        if text:
            tf = shape.TextFrame
            tf.WordWrap = -1
            tr = tf.TextRange
            tr.Text = text
            tr.Font.Name = "Segoe UI"
            tr.Font.Size = font_size
            tr.Font.Bold = 1 if bold else 0
            tr.Font.Color.RGB = hex_to_rgb_int(text_color_hex)
            tr.ParagraphFormat.Alignment = 2  # Center

        return {
            "status": "success",
            "slide_index": slide_index,
            "shape_name": shape.Name,
            "shape_id": shape.Id,
            "type": shape_type,
        }

    def add_bullet_list(
        self,
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
        """Add a sleek bullet point list to a slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        shape = slide.Shapes.AddTextbox(1, left, top, width, height)
        tf = shape.TextFrame
        tf.WordWrap = -1
        tr = tf.TextRange

        full_text = "\n".join(items)
        tr.Text = full_text
        tr.Font.Name = font_name
        tr.Font.Size = font_size
        tr.Font.Color.RGB = hex_to_rgb_int(color_hex)

        for p_idx in range(1, tr.Paragraphs().Count + 1):
            p = tr.Paragraphs(p_idx)
            p.ParagraphFormat.Bullet.Type = 1  # ppBulletUnnumbered
            p.ParagraphFormat.SpaceBefore = 8
            p.ParagraphFormat.SpaceAfter = 8

        return {
            "status": "success",
            "slide_index": slide_index,
            "item_count": len(items),
            "shape_id": shape.Id,
        }

    def add_table(
        self,
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
        """Add a cleanly styled data table to a slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        table_shape = slide.Shapes.AddTable(rows, cols, left, top, width, height)
        tbl = table_shape.Table

        for r_idx in range(rows):
            is_header = (r_idx == 0)
            row_bg = hex_to_rgb_int(header_bg_hex if is_header else (row_bg_hex if r_idx % 2 == 1 else alt_row_bg_hex))

            for c_idx in range(cols):
                cell = tbl.Cell(r_idx + 1, c_idx + 1)
                cell.Shape.Fill.Solid()
                cell.Shape.Fill.ForeColor.RGB = row_bg

                cell_text = ""
                if r_idx < len(data) and c_idx < len(data[r_idx]):
                    cell_text = str(data[r_idx][c_idx])

                tr = cell.Shape.TextFrame.TextRange
                tr.Text = cell_text
                tr.Font.Name = "Segoe UI"
                tr.Font.Size = 14 if not is_header else 15
                tr.Font.Bold = 1 if is_header else 0
                tr.Font.Color.RGB = hex_to_rgb_int(header_text_hex if is_header else text_color_hex)
                tr.ParagraphFormat.Alignment = 2 if is_header else 1

        return {
            "status": "success",
            "slide_index": slide_index,
            "rows": rows,
            "cols": cols,
            "shape_id": table_shape.Id,
        }

    def add_image(
        self,
        slide_index: int,
        image_path: str,
        left: float = 100,
        top: float = 100,
        width: Optional[float] = None,
        height: Optional[float] = None,
    ) -> Dict[str, Any]:
        """Insert an image file from disk into a slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        abs_path = os.path.abspath(image_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Image not found: {abs_path}")

        self._app.ActiveWindow.View.GotoSlide(slide_index)
        pic = slide.Shapes.AddPicture(
            abs_path,
            0,  # LinkToFile = False
            -1, # SaveWithDocument = True
            left,
            top,
            width if width else -1,
            height if height else -1,
        )

        return {
            "status": "success",
            "slide_index": slide_index,
            "shape_name": pic.Name,
            "shape_id": pic.Id,
            "path": abs_path,
        }

    def add_animation(
        self,
        slide_index: int,
        shape_index_or_id: Any,
        effect: str = "fade",
        trigger: str = "on_click",  # on_click or after_previous
    ) -> Dict[str, Any]:
        """Add an entrance animation to a shape (makes elements visually appear/animate!)."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Find target shape
        target_shape = None
        if isinstance(shape_index_or_id, int) and shape_index_or_id <= slide.Shapes.Count:
            target_shape = slide.Shapes(shape_index_or_id)
        else:
            for s in slide.Shapes:
                if str(s.Id) == str(shape_index_or_id) or s.Name.lower() == str(shape_index_or_id).lower():
                    target_shape = s
                    break

        if not target_shape:
            # Fallback to last shape on slide
            if slide.Shapes.Count > 0:
                target_shape = slide.Shapes(slide.Shapes.Count)
            else:
                raise ValueError("No shapes found on slide to animate")

        effect_code = ANIMATION_EFFECTS.get(effect.lower(), 10)  # default fade
        eff = slide.TimeLine.MainSequence.AddEffect(target_shape, effect_code)

        if trigger == "after_previous":
            eff.Timing.TriggerType = 3  # msoAnimTriggerAfterPrevious
        else:
            eff.Timing.TriggerType = 1  # msoAnimTriggerOnPageClick

        return {
            "status": "success",
            "slide_index": slide_index,
            "shape_name": target_shape.Name,
            "effect": effect,
        }

    def set_slide_transition(
        self,
        slide_index: int,
        transition: str = "fade",
        speed: str = "medium",  # slow, medium, fast
    ) -> Dict[str, Any]:
        """Set visual slide transition effect."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        eff_code = TRANSITION_EFFECTS.get(transition.lower(), 1793)  # default fade
        slide.SlideShowTransition.EntryEffect = eff_code

        speed_map = {"slow": 1, "medium": 2, "fast": 3}
        slide.SlideShowTransition.Speed = speed_map.get(speed.lower(), 2)

        return {
            "status": "success",
            "slide_index": slide_index,
            "transition": transition,
            "speed": speed,
        }

    def set_speaker_notes(self, slide_index: int, notes_text: str) -> Dict[str, Any]:
        """Add or update speaker notes for a slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        try:
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes_text
        except Exception:
            # If placeholder 2 not present, add a note shape
            slide.NotesPage.Shapes.AddTextbox(1, 50, 50, 500, 200).TextFrame.TextRange.Text = notes_text

        return {
            "status": "success",
            "slide_index": slide_index,
            "notes": notes_text,
        }

    # Alias for convenience
    set_slide_notes = set_speaker_notes

    def run_slideshow(self, start_slide: int = 1) -> Dict[str, Any]:
        """Start full-screen presentation mode live on screen!"""
        pres = self.get_presentation()
        pres.SlideShowSettings.StartingSlide = start_slide
        pres.SlideShowSettings.EndingSlide = pres.Slides.Count
        ss_window = pres.SlideShowSettings.Run()
        return {
            "status": "success",
            "message": "Slideshow started",
            "current_slide": start_slide,
        }

    def slideshow_next(self) -> Dict[str, Any]:
        """Advance to next slide or animation in active slideshow."""
        app = self.get_app()
        if app.SlideShowWindows.Count > 0:
            view = app.SlideShowWindows(1).View
            view.Next()
            curr = view.Slide.SlideIndex
            return {"status": "success", "message": "Advanced slideshow", "current_slide": curr}
        return {"status": "error", "message": "No active slideshow running"}

    def slideshow_previous(self) -> Dict[str, Any]:
        """Go back to previous slide in active slideshow."""
        app = self.get_app()
        if app.SlideShowWindows.Count > 0:
            view = app.SlideShowWindows(1).View
            view.Previous()
            curr = view.Slide.SlideIndex
            return {"status": "success", "message": "Previous slideshow slide", "current_slide": curr}
        return {"status": "error", "message": "No active slideshow running"}

    def slideshow_exit(self) -> Dict[str, Any]:
        """Exit running slideshow and return to PowerPoint editing view."""
        app = self.get_app()
        if app.SlideShowWindows.Count > 0:
            app.SlideShowWindows(1).View.Exit()
            return {"status": "success", "message": "Slideshow exited"}
        return {"status": "info", "message": "No slideshow was running"}

    def export_slide_image(
        self,
        slide_index: int,
        output_path: str,
        width: int = 1920,
        height: int = 1080,
    ) -> Dict[str, Any]:
        """Export a slide to high-resolution PNG or JPG image file."""
        pres = self.get_presentation()
        if slide_index < 1 or slide_index > pres.Slides.Count:
            raise ValueError(f"Slide index {slide_index} out of range (1..{pres.Slides.Count})")

        slide = pres.Slides(slide_index)
        abs_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        ext = "PNG" if abs_path.lower().endswith(".png") else "JPG"
        slide.Export(abs_path, ext, width, height)
        return {
            "status": "success",
            "message": f"Exported slide {slide_index} to image",
            "slide_index": slide_index,
            "image_path": abs_path,
            "resolution": f"{width}x{height}",
        }

    def duplicate_slide(self, slide_index: int) -> Dict[str, Any]:
        """Duplicate an existing slide and jump view to the new copy."""
        pres = self.get_presentation()
        if slide_index < 1 or slide_index > pres.Slides.Count:
            raise ValueError(f"Slide index {slide_index} out of range (1..{pres.Slides.Count})")

        slide = pres.Slides(slide_index)
        new_slide = slide.Duplicate()
        new_idx = new_slide.SlideIndex
        self._app.ActiveWindow.View.GotoSlide(new_idx)
        self.focus_window()
        return {
            "status": "success",
            "message": f"Duplicated slide {slide_index}",
            "original_index": slide_index,
            "new_slide_index": new_idx,
            "total_slides": pres.Slides.Count,
        }

    def move_slide(self, from_index: int, to_index: int) -> Dict[str, Any]:
        """Move / reorder a slide to a new position."""
        pres = self.get_presentation()
        if from_index < 1 or from_index > pres.Slides.Count:
            raise ValueError(f"from_index {from_index} out of range")
        if to_index < 1 or to_index > pres.Slides.Count:
            raise ValueError(f"to_index {to_index} out of range")

        slide = pres.Slides(from_index)
        slide.MoveTo(to_index)
        self._app.ActiveWindow.View.GotoSlide(to_index)
        self.focus_window()
        return {
            "status": "success",
            "message": f"Moved slide from {from_index} to {to_index}",
            "current_index": to_index,
        }

    def read_slide_content(self, slide_index: int) -> Dict[str, Any]:
        """Extract and inspect all shapes, text boxes, tables, and notes from a slide."""
        pres = self.get_presentation()
        if slide_index < 1 or slide_index > pres.Slides.Count:
            raise ValueError(f"Slide index {slide_index} out of range")

        slide = pres.Slides(slide_index)
        shapes_data = []

        for s in slide.Shapes:
            item = {
                "id": s.Id,
                "name": s.Name,
                "left": round(s.Left, 1),
                "top": round(s.Top, 1),
                "width": round(s.Width, 1),
                "height": round(s.Height, 1),
            }
            if s.HasTextFrame and s.TextFrame.HasText:
                item["text"] = s.TextFrame.TextRange.Text.strip()
            if s.HasTable:
                table_rows = []
                for r in range(1, s.Table.Rows.Count + 1):
                    row_data = []
                    for c in range(1, s.Table.Columns.Count + 1):
                        row_data.append(s.Table.Cell(r, c).Shape.TextFrame.TextRange.Text.strip())
                    table_rows.append(row_data)
                item["table_data"] = table_rows

            shapes_data.append(item)

        notes = ""
        try:
            notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text.strip()
        except Exception:
            pass

        return {
            "slide_index": slide_index,
            "shapes_count": len(shapes_data),
            "shapes": shapes_data,
            "speaker_notes": notes,
        }

    def add_metric_card(
        self,
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
        """Create a prominent KPI / Metric card widget with large numbers and subtext."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Card container
        card = slide.Shapes.AddShape(5, left, top, width, height)  # rounded rectangle
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)
        if border_color_hex:
            card.Line.Visible = -1
            card.Line.ForeColor.RGB = hex_to_rgb_int(border_color_hex)
            card.Line.Weight = 1.2
        else:
            card.Line.Visible = 0

        # Optional top accent bar
        if accent_bar_color_hex:
            accent = slide.Shapes.AddShape(1, left + 4, top + 2, width - 8, 4)
            accent.Fill.Solid()
            accent.Fill.ForeColor.RGB = hex_to_rgb_int(accent_bar_color_hex)
            accent.Line.Visible = 0

        # Value (large bold metric)
        val_box = slide.Shapes.AddTextbox(1, left + 14, top + 14, width - 28, 48)
        val_box.Fill.Background()
        val_tf = val_box.TextFrame
        val_tf.MarginLeft = 0
        val_tf.MarginRight = 0
        val_tf.MarginTop = 0
        val_tf.MarginBottom = 0
        tr_val = val_tf.TextRange
        tr_val.Text = value
        tr_val.Font.Name = "Segoe UI"
        tr_val.Font.Size = 34
        tr_val.Font.Bold = 1
        tr_val.Font.Color.RGB = hex_to_rgb_int(value_color_hex)

        # Label & Subtitle
        lbl_box = slide.Shapes.AddTextbox(1, left + 14, top + 64, width - 28, 60)
        lbl_box.Fill.Background()
        lbl_tf = lbl_box.TextFrame
        lbl_tf.MarginLeft = 0
        lbl_tf.MarginRight = 0
        lbl_tf.MarginTop = 0
        lbl_tf.MarginBottom = 0
        tr_lbl = lbl_tf.TextRange
        lbl_text = f"{label}\n{subtitle}" if subtitle else label
        tr_lbl.Text = lbl_text
        tr_lbl.Paragraphs(1).Font.Name = "Segoe UI"
        tr_lbl.Paragraphs(1).Font.Size = 14
        tr_lbl.Paragraphs(1).Font.Bold = 1
        tr_lbl.Paragraphs(1).Font.Color.RGB = hex_to_rgb_int("#0F172A")

        if subtitle and tr_lbl.Paragraphs().Count > 1:
            p2 = tr_lbl.Paragraphs(2)
            p2.Font.Name = "Segoe UI"
            p2.Font.Size = 12
            p2.Font.Bold = 0
            p2.Font.Color.RGB = hex_to_rgb_int("#64748B")

        return {
            "status": "success",
            "slide_index": slide_index,
            "value": value,
            "label": label,
            "card_id": card.Id,
        }

    def add_timeline(
        self,
        slide_index: int,
        steps: List[Dict[str, str]],
        left: float = 80,
        top: float = 160,
        width: float = 800,
        height: float = 240,
        accent_color_hex: str = "#3B82F6",
        bg_color_hex: str = "#FFFFFF",
    ) -> Dict[str, Any]:
        """Create a clean horizontal process flow / roadmap / timeline."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_steps = len(steps)
        if num_steps == 0:
            return {"status": "error", "message": "steps list is empty"}

        gap = 20
        item_w = (width - ((num_steps - 1) * gap)) / num_steps

        # Horizontal connecting line
        line_y = top + 22
        conn_line = slide.Shapes.AddShape(1, left + (item_w / 2), line_y, width - item_w, 3)
        conn_line.Fill.Solid()
        conn_line.Fill.ForeColor.RGB = hex_to_rgb_int("#CBD5E1")
        conn_line.Line.Visible = 0

        for i, step in enumerate(steps):
            cx = left + i * (item_w + gap)

            # Circular node
            node_size = 44
            node_x = cx + (item_w / 2) - (node_size / 2)
            node = slide.Shapes.AddShape(9, node_x, top, node_size, node_size)  # circle
            node.Fill.Solid()
            node.Fill.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
            node.Line.Visible = 0
            node_tf = node.TextFrame
            node_tr = node_tf.TextRange
            node_tr.Text = str(i + 1)
            node_tr.Font.Name = "Segoe UI"
            node_tr.Font.Size = 18
            node_tr.Font.Bold = 1
            node_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
            node_tr.ParagraphFormat.Alignment = 2  # Center

            # Step card underneath
            card_y = top + 56
            card_h = height - 56
            title_text = step.get("title", f"Step {i+1}")
            desc_text = step.get("desc", step.get("description", ""))

            self.add_card(
                slide_index=slide_index,
                title=title_text,
                body=desc_text,
                left=cx,
                top=card_y,
                width=item_w,
                height=card_h,
                bg_color_hex=bg_color_hex,
                border_color_hex="#E2E8F0",
                title_color_hex="#0F172A",
                body_color_hex="#475569",
                accent_bar_color_hex=accent_color_hex,
            )

        return {
            "status": "success",
            "slide_index": slide_index,
            "steps_count": num_steps,
        }

    def add_code_block(
        self,
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
        """Create a developer code editor mockup with window buttons and syntax look."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Editor Window background
        bg = slide.Shapes.AddShape(5, left, top, width, height)  # rounded rectangle
        bg.Fill.Solid()
        bg.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)
        bg.Line.Visible = -1
        bg.Line.ForeColor.RGB = hex_to_rgb_int("#27272A")
        bg.Line.Weight = 1.0

        # Window controls: 3 dots (red, yellow, green)
        dot_colors = ["#EF4444", "#F59E0B", "#10B981"]
        dot_y = top + 12
        dot_size = 11
        for i, c in enumerate(dot_colors):
            dx = left + 16 + (i * 18)
            dot = slide.Shapes.AddShape(9, dx, dot_y, dot_size, dot_size)
            dot.Fill.Solid()
            dot.Fill.ForeColor.RGB = hex_to_rgb_int(c)
            dot.Line.Visible = 0

        # Language Pill
        lang_badge = slide.Shapes.AddTextbox(1, left + width - 110, top + 8, 90, 20)
        lang_tr = lang_badge.TextFrame.TextRange
        lang_tr.Text = language.upper()
        lang_tr.Font.Name = "Consolas"
        lang_tr.Font.Size = 10
        lang_tr.Font.Bold = 1
        lang_tr.Font.Color.RGB = hex_to_rgb_int("#71717A")
        lang_tr.ParagraphFormat.Alignment = 3  # Right

        # Divider line
        div = slide.Shapes.AddShape(1, left, top + 34, width, 1)
        div.Fill.Solid()
        div.Fill.ForeColor.RGB = hex_to_rgb_int("#27272A")
        div.Line.Visible = 0

        # Code content
        code_box = slide.Shapes.AddTextbox(1, left + 14, top + 42, width - 28, height - 52)
        code_tf = code_box.TextFrame
        code_tf.WordWrap = -1
        code_tr = code_tf.TextRange
        code_tr.Text = code
        code_tr.Font.Name = "Consolas"
        code_tr.Font.Size = font_size
        code_tr.Font.Color.RGB = hex_to_rgb_int("#F4F4F5")

        return {
            "status": "success",
            "slide_index": slide_index,
            "language": language,
            "shape_id": bg.Id,
        }

    def add_bar_chart(
        self,
        slide_index: int,
        data: List[Dict[str, Any]],
        left: float = 80,
        top: float = 140,
        width: float = 800,
        height: float = 280,
        bar_color_hex: str = "#3B82F6",
        bg_color_hex: str = "#F8FAFC",
    ) -> Dict[str, Any]:
        """Create a clean horizontal vector bar chart with labels and progress bars."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Container
        box = slide.Shapes.AddShape(5, left, top, width, height)
        box.Fill.Solid()
        box.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)
        box.Line.Visible = -1
        box.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")

        num_items = len(data)
        if num_items == 0:
            return {"status": "error", "message": "data is empty"}

        # Find max value for scaling
        max_val = max([float(d.get("value", 1)) for d in data]) or 1.0

        row_h = (height - 30) / num_items
        max_bar_w = width * 0.55
        bar_start_x = left + (width * 0.32)

        for i, item in enumerate(data):
            label = str(item.get("label", ""))
            val = float(item.get("value", 0))
            display_val = str(item.get("display_value", f"{val:g}"))
            item_color = item.get("color", bar_color_hex)

            ry = top + 15 + (i * row_h)
            bar_y = ry + (row_h * 0.25)
            bar_h = row_h * 0.5

            # Label on left
            lbl_box = slide.Shapes.AddTextbox(1, left + 15, ry, (width * 0.28), row_h)
            lbl_tr = lbl_box.TextFrame.TextRange
            lbl_tr.Text = label
            lbl_tr.Font.Name = "Segoe UI"
            lbl_tr.Font.Size = 13
            lbl_tr.Font.Bold = 1
            lbl_tr.Font.Color.RGB = hex_to_rgb_int("#1E293B")
            lbl_tr.ParagraphFormat.Alignment = 3  # Right align towards bar

            # Background bar (track)
            track = slide.Shapes.AddShape(5, bar_start_x, bar_y, max_bar_w, bar_h)
            track.Fill.Solid()
            track.Fill.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
            track.Line.Visible = 0

            # Active bar fill
            ratio = min(max(val / max_val, 0.02), 1.0)
            fill_w = max_bar_w * ratio
            fill_bar = slide.Shapes.AddShape(5, bar_start_x, bar_y, fill_w, bar_h)
            fill_bar.Fill.Solid()
            fill_bar.Fill.ForeColor.RGB = hex_to_rgb_int(item_color)
            fill_bar.Line.Visible = 0

            # Value label at end of bar
            val_box = slide.Shapes.AddTextbox(1, bar_start_x + fill_w + 8, ry, 70, row_h)
            val_tr = val_box.TextFrame.TextRange
            val_tr.Text = display_val
            val_tr.Font.Name = "Segoe UI"
            val_tr.Font.Size = 12
            val_tr.Font.Bold = 1
            val_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

        return {
            "status": "success",
            "slide_index": slide_index,
            "items_count": num_items,
        }

    def add_badge(
        self,
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
        """Create a rounded tag / pill badge component."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        w = width if width else (max(len(text) * 8 + 24, 70))
        shape = slide.Shapes.AddShape(5, left, top, w, height)  # rounded rectangle
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)

        if border_color_hex:
            shape.Line.Visible = -1
            shape.Line.ForeColor.RGB = hex_to_rgb_int(border_color_hex)
            shape.Line.Weight = 1.0
        else:
            shape.Line.Visible = 0

        tf = shape.TextFrame
        tr = tf.TextRange
        tr.Text = text
        tr.Font.Name = "Segoe UI"
        tr.Font.Size = font_size
        tr.Font.Bold = 1 if bold else 0
        tr.Font.Color.RGB = hex_to_rgb_int(text_color_hex)
        tr.ParagraphFormat.Alignment = 2  # Center

        return {
            "status": "success",
            "slide_index": slide_index,
            "text": text,
            "shape_id": shape.Id,
        }


# Global singleton controller instance
controller = PowerPointController()

