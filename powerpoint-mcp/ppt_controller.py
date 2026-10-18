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

    def search_and_replace_text(
        self,
        find_text: str,
        replace_text: str,
        slide_index: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Find and replace text across slides or within a specific slide."""
        pres = self.get_presentation()
        total_replacements = 0

        target_slides = [pres.Slides(slide_index)] if slide_index else [pres.Slides(i) for i in range(1, pres.Slides.Count + 1)]

        for s in target_slides:
            for shp in s.Shapes:
                if shp.HasTextFrame and shp.TextFrame.HasText:
                    tr = shp.TextFrame.TextRange
                    # Count occurrences
                    if find_text.lower() in tr.Text.lower():
                        while True:
                            found = tr.Replace(find_text, replace_text)
                            if found is None or found.Length == 0:
                                break
                            total_replacements += 1

                # Check tables
                if shp.HasTable:
                    for r in range(1, shp.Table.Rows.Count + 1):
                        for c in range(1, shp.Table.Columns.Count + 1):
                            cell_tr = shp.Table.Cell(r, c).Shape.TextFrame.TextRange
                            if find_text.lower() in cell_tr.Text.lower():
                                while True:
                                    found = cell_tr.Replace(find_text, replace_text)
                                    if found is None or found.Length == 0:
                                        break
                                    total_replacements += 1

        return {
            "status": "success",
            "find_text": find_text,
            "replace_text": replace_text,
            "replacements_count": total_replacements,
            "scope": f"slide {slide_index}" if slide_index else "all slides",
        }

    def add_quote_card(
        self,
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
        """Create an elegant testimonial or quotation card with author details."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Card container
        card = slide.Shapes.AddShape(5, left, top, width, height)  # rounded rectangle
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)
        card.Line.Visible = -1
        card.Line.ForeColor.RGB = hex_to_rgb_int(border_color_hex)
        card.Line.Weight = 1.2

        # Left vertical accent stripe
        stripe = slide.Shapes.AddShape(1, left + 2, top + 6, 6, height - 12)
        stripe.Fill.Solid()
        stripe.Fill.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
        stripe.Line.Visible = 0

        # Big decorative quote mark
        q_mark = slide.Shapes.AddTextbox(1, left + 24, top + 10, 60, 45)
        q_mark.Fill.Background()
        q_tr = q_mark.TextFrame.TextRange
        q_tr.Text = "“"
        q_tr.Font.Name = "Georgia"
        q_tr.Font.Size = 48
        q_tr.Font.Bold = 1
        q_tr.Font.Color.RGB = hex_to_rgb_int(accent_color_hex)

        # Quote body
        body_box = slide.Shapes.AddTextbox(1, left + 40, top + 55, width - 80, height - 110)
        body_box.Fill.Background()
        b_tf = body_box.TextFrame
        b_tf.WordWrap = -1
        b_tr = b_tf.TextRange
        b_tr.Text = f'"{quote}"'
        b_tr.Font.Name = "Georgia"
        b_tr.Font.Size = 18
        b_tr.Font.Italic = 1
        b_tr.Font.Color.RGB = hex_to_rgb_int(quote_color_hex)

        # Author & Role on bottom right / left
        auth_box = slide.Shapes.AddTextbox(1, left + 40, top + height - 50, width - 80, 40)
        auth_box.Fill.Background()
        a_tr = auth_box.TextFrame.TextRange
        a_tr.Text = f"— {author}" + (f", {role}" if role else "")
        a_tr.Font.Name = "Segoe UI"
        a_tr.Font.Size = 13
        a_tr.Font.Bold = 1
        a_tr.Font.Color.RGB = hex_to_rgb_int("#475569")

        return {
            "status": "success",
            "slide_index": slide_index,
            "author": author,
            "card_id": card.Id,
        }

    def add_pros_cons(
        self,
        slide_index: int,
        pros: List[str],
        cons: List[str],
        left: float = 80,
        top: float = 140,
        width: float = 800,
        height: float = 330,
    ) -> Dict[str, Any]:
        """Create side-by-side comparison cards for Pros (green) and Cons (red)."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        gap = 24
        col_w = (width - gap) / 2

        # 1. Left Card: PROS (Green theme)
        card_pros = slide.Shapes.AddShape(5, left, top, col_w, height)
        card_pros.Fill.Solid()
        card_pros.Fill.ForeColor.RGB = hex_to_rgb_int("#F0FDF4")  # light green
        card_pros.Line.Visible = -1
        card_pros.Line.ForeColor.RGB = hex_to_rgb_int("#86EFAC")

        # Top green bar
        bar_p = slide.Shapes.AddShape(1, left + 4, top + 2, col_w - 8, 4)
        bar_p.Fill.Solid()
        bar_p.Fill.ForeColor.RGB = hex_to_rgb_int("#16A34A")
        bar_p.Line.Visible = 0

        # Pros Title
        t_pros = slide.Shapes.AddTextbox(1, left + 18, top + 15, col_w - 36, 35)
        t_pros.Fill.Background()
        t_pros_tr = t_pros.TextFrame.TextRange
        t_pros_tr.Text = "✅ AVANTAGES / FORCES"
        t_pros_tr.Font.Name = "Segoe UI"
        t_pros_tr.Font.Size = 16
        t_pros_tr.Font.Bold = 1
        t_pros_tr.Font.Color.RGB = hex_to_rgb_int("#15803D")

        # Pros list
        p_list = slide.Shapes.AddTextbox(1, left + 18, top + 55, col_w - 36, height - 70)
        p_list.Fill.Background()
        p_list_tf = p_list.TextFrame
        p_list_tf.WordWrap = -1
        p_list_tr = p_list_tf.TextRange
        p_list_tr.Text = "\n".join([f"•  {item}" for item in pros])
        p_list_tr.Font.Name = "Segoe UI"
        p_list_tr.Font.Size = 14
        p_list_tr.Font.Color.RGB = hex_to_rgb_int("#14532D")
        for i in range(1, p_list_tr.Paragraphs().Count + 1):
            p_list_tr.Paragraphs(i).ParagraphFormat.SpaceBefore = 6

        # 2. Right Card: CONS (Red/Rose theme)
        rx = left + col_w + gap
        card_cons = slide.Shapes.AddShape(5, rx, top, col_w, height)
        card_cons.Fill.Solid()
        card_cons.Fill.ForeColor.RGB = hex_to_rgb_int("#FEF2F2")  # light red
        card_cons.Line.Visible = -1
        card_cons.Line.ForeColor.RGB = hex_to_rgb_int("#FCA5A5")

        # Top red bar
        bar_c = slide.Shapes.AddShape(1, rx + 4, top + 2, col_w - 8, 4)
        bar_c.Fill.Solid()
        bar_c.Fill.ForeColor.RGB = hex_to_rgb_int("#DC2626")
        bar_c.Line.Visible = 0

        # Cons Title
        t_cons = slide.Shapes.AddTextbox(1, rx + 18, top + 15, col_w - 36, 35)
        t_cons.Fill.Background()
        t_cons_tr = t_cons.TextFrame.TextRange
        t_cons_tr.Text = "⚠️ INCONVÉNIENTS / RISQUES"
        t_cons_tr.Font.Name = "Segoe UI"
        t_cons_tr.Font.Size = 16
        t_cons_tr.Font.Bold = 1
        t_cons_tr.Font.Color.RGB = hex_to_rgb_int("#B91C1C")

        # Cons list
        c_list = slide.Shapes.AddTextbox(1, rx + 18, top + 55, col_w - 36, height - 70)
        c_list.Fill.Background()
        c_list_tf = c_list.TextFrame
        c_list_tf.WordWrap = -1
        c_list_tr = c_list_tf.TextRange
        c_list_tr.Text = "\n".join([f"•  {item}" for item in cons])
        c_list_tr.Font.Name = "Segoe UI"
        c_list_tr.Font.Size = 14
        c_list_tr.Font.Color.RGB = hex_to_rgb_int("#7F1D1D")
        for i in range(1, c_list_tr.Paragraphs().Count + 1):
            c_list_tr.Paragraphs(i).ParagraphFormat.SpaceBefore = 6

        return {
            "status": "success",
            "slide_index": slide_index,
            "pros_count": len(pros),
            "cons_count": len(cons),
        }

    def add_pricing_table(
        self,
        slide_index: int,
        tiers: List[Dict[str, Any]],
        left: float = 80,
        top: float = 130,
        width: float = 800,
        height: float = 360,
    ) -> Dict[str, Any]:
        """Create a multi-tier SaaS pricing comparison cards layout."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_tiers = len(tiers)
        if num_tiers == 0:
            return {"status": "error", "message": "tiers list is empty"}

        gap = 20
        col_w = (width - ((num_tiers - 1) * gap)) / num_tiers

        for i, t in enumerate(tiers):
            cx = left + i * (col_w + gap)
            is_popular = t.get("highlighted", False)

            # Container
            bg_col = "#FFFFFF" if not is_popular else "#F8FAFC"
            border_col = "#3B82F6" if is_popular else "#E2E8F0"
            border_weight = 2.0 if is_popular else 1.0

            card = slide.Shapes.AddShape(5, cx, top, col_w, height)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = hex_to_rgb_int(bg_col)
            card.Line.Visible = -1
            card.Line.ForeColor.RGB = hex_to_rgb_int(border_col)
            card.Line.Weight = border_weight

            # Optional POPULAR tag
            if is_popular:
                pop_badge = slide.Shapes.AddShape(5, cx + (col_w / 2) - 50, top - 12, 100, 24)
                pop_badge.Fill.Solid()
                pop_badge.Fill.ForeColor.RGB = hex_to_rgb_int("#2563EB")
                pop_badge.Line.Visible = 0
                b_tr = pop_badge.TextFrame.TextRange
                b_tr.Text = "POPULAIRE"
                b_tr.Font.Name = "Segoe UI"
                b_tr.Font.Size = 10
                b_tr.Font.Bold = 1
                b_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
                b_tr.ParagraphFormat.Alignment = 2

            # Plan Name
            name_box = slide.Shapes.AddTextbox(1, cx + 14, top + 18, col_w - 28, 28)
            name_box.Fill.Background()
            n_tr = name_box.TextFrame.TextRange
            n_tr.Text = t.get("name", "Plan").upper()
            n_tr.Font.Name = "Segoe UI"
            n_tr.Font.Size = 13
            n_tr.Font.Bold = 1
            n_tr.Font.Color.RGB = hex_to_rgb_int("#2563EB" if is_popular else "#64748B")

            # Price
            price_box = slide.Shapes.AddTextbox(1, cx + 14, top + 46, col_w - 28, 48)
            price_box.Fill.Background()
            p_tr = price_box.TextFrame.TextRange
            p_tr.Text = t.get("price", "$0")
            p_tr.Font.Name = "Segoe UI"
            p_tr.Font.Size = 30
            p_tr.Font.Bold = 1
            p_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

            # Period
            period_box = slide.Shapes.AddTextbox(1, cx + 14, top + 92, col_w - 28, 22)
            period_box.Fill.Background()
            prd_tr = period_box.TextFrame.TextRange
            prd_tr.Text = t.get("period", "/ mois")
            prd_tr.Font.Name = "Segoe UI"
            prd_tr.Font.Size = 11
            prd_tr.Font.Color.RGB = hex_to_rgb_int("#94A3B8")

            # Divider
            div = slide.Shapes.AddShape(1, cx + 14, top + 118, col_w - 28, 1)
            div.Fill.Solid()
            div.Fill.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
            div.Line.Visible = 0

            # Features list
            feat_box = slide.Shapes.AddTextbox(1, cx + 14, top + 128, col_w - 28, height - 145)
            feat_box.Fill.Background()
            f_tf = feat_box.TextFrame
            f_tf.WordWrap = -1
            f_tr = f_tf.TextRange
            features = t.get("features", [])
            f_tr.Text = "\n".join([f"✓ {f}" for f in features])
            f_tr.Font.Name = "Segoe UI"
            f_tr.Font.Size = 12
            f_tr.Font.Color.RGB = hex_to_rgb_int("#334155")
            for p_i in range(1, f_tr.Paragraphs().Count + 1):
                f_tr.Paragraphs(p_i).ParagraphFormat.SpaceBefore = 5

        return {
            "status": "success",
            "slide_index": slide_index,
            "tiers_count": num_tiers,
        }

    def add_donut_chart(
        self,
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
        """Create a circular percentage metric widget with central stat number."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Outer background track circle
        outer = slide.Shapes.AddShape(9, left, top, size, size)  # circle
        outer.Fill.Solid()
        outer.Fill.ForeColor.RGB = hex_to_rgb_int(fill_color_hex)
        outer.Line.Visible = 0

        # Inner cutout circle (creating donut ring)
        inner_size = size * 0.72
        inner_offset = (size - inner_size) / 2
        inner = slide.Shapes.AddShape(9, left + inner_offset, top + inner_offset, inner_size, inner_size)
        inner.Fill.Solid()
        inner.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)
        inner.Line.Visible = 0

        # Center percentage text
        pct_box = slide.Shapes.AddTextbox(1, left + inner_offset, top + inner_offset + (inner_size * 0.2), inner_size, 45)
        pct_box.Fill.Background()
        pct_tr = pct_box.TextFrame.TextRange
        pct_tr.Text = f"{int(percentage)}%"
        pct_tr.Font.Name = "Segoe UI"
        pct_tr.Font.Size = 28
        pct_tr.Font.Bold = 1
        pct_tr.Font.Color.RGB = hex_to_rgb_int(fill_color_hex)
        pct_tr.ParagraphFormat.Alignment = 2  # Center

        # Label underneath donut
        lbl_box = slide.Shapes.AddTextbox(1, left - 20, top + size + 12, size + 40, 40)
        lbl_box.Fill.Background()
        lbl_tr = lbl_box.TextFrame.TextRange
        lbl_tr.Text = label
        lbl_tr.Font.Name = "Segoe UI"
        lbl_tr.Font.Size = 14
        lbl_tr.Font.Bold = 1
        lbl_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")
        lbl_tr.ParagraphFormat.Alignment = 2

        return {
            "status": "success",
            "slide_index": slide_index,
            "percentage": percentage,
            "label": label,
        }

    def delete_shape(self, slide_index: int, shape_id_or_name: Any) -> Dict[str, Any]:
        """Delete a specific shape by its ID or name from a slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)

        found = False
        target_name = ""
        for s in slide.Shapes:
            if str(s.Id) == str(shape_id_or_name) or s.Name.lower() == str(shape_id_or_name).lower():
                target_name = s.Name
                s.Delete()
                found = True
                break

        if not found:
            raise ValueError(f"Shape '{shape_id_or_name}' not found on slide {slide_index}")

        return {
            "status": "success",
            "message": f"Deleted shape '{target_name}' from slide {slide_index}",
            "remaining_shapes": slide.Shapes.Count,
        }

    def clear_slide(self, slide_index: int) -> Dict[str, Any]:
        """Remove all shapes from a slide to reset it to blank."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        count = slide.Shapes.Count

        for i in range(count, 0, -1):
            slide.Shapes(i).Delete()

        return {
            "status": "success",
            "message": f"Cleared {count} shapes from slide {slide_index}",
            "slide_index": slide_index,
        }

    def add_footer(
        self,
        slide_index: Optional[int] = None,
        text: str = "Confidential & Proprietary",
        show_slide_number: bool = True,
    ) -> Dict[str, Any]:
        """Add a sleek professional footer with notice and slide number."""
        pres = self.get_presentation()
        target_slides = [pres.Slides(slide_index)] if slide_index else [pres.Slides(i) for i in range(1, pres.Slides.Count + 1)]

        for s in target_slides:
            w = pres.PageSetup.SlideWidth
            h = pres.PageSetup.SlideHeight

            # Thin divider line
            line = s.Shapes.AddShape(1, 40, h - 35, w - 80, 1)
            line.Fill.Solid()
            line.Fill.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
            line.Line.Visible = 0

            # Footer text left
            box_l = s.Shapes.AddTextbox(1, 40, h - 30, w - 160, 22)
            box_l.Fill.Background()
            tl = box_l.TextFrame.TextRange
            tl.Text = text
            tl.Font.Name = "Segoe UI"
            tl.Font.Size = 10
            tl.Font.Color.RGB = hex_to_rgb_int("#94A3B8")

            # Slide number right
            if show_slide_number:
                box_r = s.Shapes.AddTextbox(1, w - 100, h - 30, 60, 22)
                box_r.Fill.Background()
                tr = box_r.TextFrame.TextRange
                tr.Text = f"Slide {s.SlideIndex}"
                tr.Font.Name = "Segoe UI"
                tr.Font.Size = 10
                tr.Font.Color.RGB = hex_to_rgb_int("#94A3B8")
                tr.ParagraphFormat.Alignment = 3  # Right

        return {
            "status": "success",
            "scope": f"slide {slide_index}" if slide_index else "all slides",
        }

    def add_callout_box(
        self,
        slide_index: int,
        text: str,
        title: str = "KEY TAKEAWAY",
        callout_type: str = "tip",
        left: float = 80,
        top: float = 140,
        width: float = 800,
        height: float = 110,
    ) -> Dict[str, Any]:
        """Create an alert, tip, or callout banner box with accent strip and icon badge."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        themes = {
            "info": {"accent": "#3B82F6", "bg": "#EFF6FF", "border": "#BFDBFE", "tag": "[i] INFO"},
            "warning": {"accent": "#F59E0B", "bg": "#FFFBEB", "border": "#FDE68A", "tag": "[!] NOTE"},
            "danger": {"accent": "#EF4444", "bg": "#FEF2F2", "border": "#FECACA", "tag": "[!] CAUTION"},
            "success": {"accent": "#10B981", "bg": "#ECFDF5", "border": "#A7F3D0", "tag": "[OK] SUCCESS"},
            "tip": {"accent": "#8B5CF6", "bg": "#F5F3FF", "border": "#DDD6FE", "tag": "[*] TIP"},
        }
        theme = themes.get(callout_type.lower(), themes["tip"])

        # Main box
        box = slide.Shapes.AddShape(5, left, top, width, height)  # rounded rectangle
        box.Fill.Solid()
        box.Fill.ForeColor.RGB = hex_to_rgb_int(theme["bg"])
        box.Line.Visible = -1
        box.Line.ForeColor.RGB = hex_to_rgb_int(theme["border"])
        box.Line.Weight = 1.0

        # Left accent stripe
        stripe = slide.Shapes.AddShape(5, left, top, 6, height)
        stripe.Fill.Solid()
        stripe.Fill.ForeColor.RGB = hex_to_rgb_int(theme["accent"])
        stripe.Line.Visible = 0

        # Badge pill
        tag_w = min(len(theme["tag"]) * 8 + 14, 110)
        badge = slide.Shapes.AddShape(5, left + 20, top + 14, tag_w, 22)
        badge.Fill.Solid()
        badge.Fill.ForeColor.RGB = hex_to_rgb_int(theme["accent"])
        badge.Line.Visible = 0
        b_tr = badge.TextFrame.TextRange
        b_tr.Text = theme["tag"]
        b_tr.Font.Name = "Segoe UI"
        b_tr.Font.Size = 9
        b_tr.Font.Bold = 1
        b_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
        b_tr.ParagraphFormat.Alignment = 2

        # Title
        title_box = slide.Shapes.AddTextbox(1, left + 26 + tag_w, top + 12, width - tag_w - 40, 26)
        title_box.Fill.Background()
        t_tr = title_box.TextFrame.TextRange
        t_tr.Text = title
        t_tr.Font.Name = "Segoe UI"
        t_tr.Font.Size = 14
        t_tr.Font.Bold = 1
        t_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

        # Body text
        body_box = slide.Shapes.AddTextbox(1, left + 20, top + 42, width - 40, height - 50)
        body_box.Fill.Background()
        b_tf = body_box.TextFrame
        b_tf.WordWrap = -1
        bd_tr = b_tf.TextRange
        bd_tr.Text = text
        bd_tr.Font.Name = "Segoe UI"
        bd_tr.Font.Size = 12
        bd_tr.Font.Color.RGB = hex_to_rgb_int("#334155")

        return {
            "status": "success",
            "slide_index": slide_index,
            "callout_type": callout_type,
            "title": title,
        }

    def add_process_steps(
        self,
        slide_index: int,
        steps: List[Dict[str, str]],
        left: float = 80,
        top: float = 160,
        width: float = 800,
        height: float = 240,
        accent_color_hex: str = "#2563EB",
    ) -> Dict[str, Any]:
        """Create a modern numbered horizontal process workflow with connecting chevrons."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_steps = len(steps)
        if num_steps == 0:
            return {"status": "error", "message": "steps list is empty"}

        spacing = 24
        card_w = (width - (spacing * (num_steps - 1))) / num_steps
        card_h = height

        for i, s in enumerate(steps):
            cx = left + (i * (card_w + spacing))
            step_num = str(s.get("number", f"{i + 1}"))
            title = str(s.get("title", f"Step {i + 1}"))
            desc = str(s.get("description", ""))

            # Step Card
            card = slide.Shapes.AddShape(5, cx, top + 25, card_w, card_h - 25)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
            card.Line.Visible = -1
            card.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
            card.Line.Weight = 1.0

            # Number badge circle (sitting on top border)
            circle_size = 46
            circle_x = cx + (card_w - circle_size) / 2
            circle_y = top
            circle = slide.Shapes.AddShape(9, circle_x, circle_y, circle_size, circle_size)
            circle.Fill.Solid()
            circle.Fill.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
            circle.Line.Visible = -1
            circle.Line.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
            circle.Line.Weight = 3.0

            c_tr = circle.TextFrame.TextRange
            c_tr.Text = step_num
            c_tr.Font.Name = "Segoe UI"
            c_tr.Font.Size = 16
            c_tr.Font.Bold = 1
            c_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
            c_tr.ParagraphFormat.Alignment = 2

            # Title
            t_box = slide.Shapes.AddTextbox(1, cx + 10, top + 55, card_w - 20, 30)
            t_box.Fill.Background()
            t_tr = t_box.TextFrame.TextRange
            t_tr.Text = title
            t_tr.Font.Name = "Segoe UI"
            t_tr.Font.Size = 15
            t_tr.Font.Bold = 1
            t_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")
            t_tr.ParagraphFormat.Alignment = 2

            # Description
            d_box = slide.Shapes.AddTextbox(1, cx + 14, top + 90, card_w - 28, card_h - 100)
            d_box.Fill.Background()
            d_tf = d_box.TextFrame
            d_tf.WordWrap = -1
            d_tr = d_tf.TextRange
            d_tr.Text = desc
            d_tr.Font.Name = "Segoe UI"
            d_tr.Font.Size = 11
            d_tr.Font.Color.RGB = hex_to_rgb_int("#64748B")
            d_tr.ParagraphFormat.Alignment = 2

            # Connecting arrow to next step
            if i < num_steps - 1:
                arrow_w = spacing - 6
                arrow_h = 14
                ax = cx + card_w + 3
                ay = top + 15
                arrow = slide.Shapes.AddShape(13, ax, ay, arrow_w, arrow_h)  # right arrow
                arrow.Fill.Solid()
                arrow.Fill.ForeColor.RGB = hex_to_rgb_int("#94A3B8")
                arrow.Line.Visible = 0

        return {
            "status": "success",
            "slide_index": slide_index,
            "steps_count": num_steps,
        }

    def add_feature_grid(
        self,
        slide_index: int,
        features: List[Dict[str, str]],
        columns: int = 3,
        left: float = 80,
        top: float = 140,
        width: float = 800,
        height: float = 300,
    ) -> Dict[str, Any]:
        """Create a multi-column grid of feature/benefit cards."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_items = len(features)
        if num_items == 0:
            return {"status": "error", "message": "features list is empty"}

        cols = max(1, min(columns, num_items))
        rows = (num_items + cols - 1) // cols

        col_gap = 20
        row_gap = 20
        card_w = (width - (col_gap * (cols - 1))) / cols
        card_h = (height - (row_gap * (rows - 1))) / rows

        for i, feat in enumerate(features):
            c = i % cols
            r = i // cols
            cx = left + (c * (card_w + col_gap))
            cy = top + (r * (card_h + row_gap))

            title = str(feat.get("title", f"Feature {i + 1}"))
            desc = str(feat.get("description", ""))
            tag = feat.get("tag", "")
            accent = feat.get("accent_color", "#3B82F6")

            # Card background
            card = slide.Shapes.AddShape(5, cx, cy, card_w, card_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
            card.Line.Visible = -1
            card.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
            card.Line.Weight = 1.0

            # Top accent line
            top_bar = slide.Shapes.AddShape(5, cx + 4, cy + 2, card_w - 8, 4)
            top_bar.Fill.Solid()
            top_bar.Fill.ForeColor.RGB = hex_to_rgb_int(accent)
            top_bar.Line.Visible = 0

            cur_y = cy + 14
            # Optional tag badge
            if tag:
                badge = slide.Shapes.AddShape(5, cx + 16, cur_y, min(len(tag) * 8 + 14, 100), 20)
                badge.Fill.Solid()
                badge.Fill.ForeColor.RGB = hex_to_rgb_int(accent)
                badge.Line.Visible = 0
                bg_tr = badge.TextFrame.TextRange
                bg_tr.Text = tag.upper()
                bg_tr.Font.Name = "Segoe UI"
                bg_tr.Font.Size = 9
                bg_tr.Font.Bold = 1
                bg_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
                bg_tr.ParagraphFormat.Alignment = 2
                cur_y += 28

            # Title
            t_box = slide.Shapes.AddTextbox(1, cx + 16, cur_y, card_w - 32, 28)
            t_box.Fill.Background()
            t_tr = t_box.TextFrame.TextRange
            t_tr.Text = title
            t_tr.Font.Name = "Segoe UI"
            t_tr.Font.Size = 14
            t_tr.Font.Bold = 1
            t_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")
            cur_y += 32

            # Description
            d_box = slide.Shapes.AddTextbox(1, cx + 16, cur_y, card_w - 32, card_h - (cur_y - cy) - 10)
            d_box.Fill.Background()
            d_tf = d_box.TextFrame
            d_tf.WordWrap = -1
            d_tr = d_tf.TextRange
            d_tr.Text = desc
            d_tr.Font.Name = "Segoe UI"
            d_tr.Font.Size = 11
            d_tr.Font.Color.RGB = hex_to_rgb_int("#64748B")

        return {
            "status": "success",
            "slide_index": slide_index,
            "feature_count": num_items,
        }

    def add_team_card(
        self,
        slide_index: int,
        name: str,
        role: str,
        company_or_dept: str = "",
        bio: str = "",
        skills_or_tags: Optional[List[str]] = None,
        avatar_initials: str = "",
        left: float = 80,
        top: float = 140,
        width: float = 380,
        height: float = 280,
        accent_color_hex: str = "#2563EB",
    ) -> Dict[str, Any]:
        """Create an executive team member or speaker profile card."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Background card
        card = slide.Shapes.AddShape(5, left, top, width, height)
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
        card.Line.Visible = -1
        card.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
        card.Line.Weight = 1.0

        # Avatar circle
        avatar_size = 56
        avatar = slide.Shapes.AddShape(9, left + 20, top + 20, avatar_size, avatar_size)
        avatar.Fill.Solid()
        avatar.Fill.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
        avatar.Line.Visible = 0

        # Initials inside avatar
        initials = avatar_initials or "".join([w[0] for w in name.split()[:2]]).upper()
        av_tr = avatar.TextFrame.TextRange
        av_tr.Text = initials
        av_tr.Font.Name = "Segoe UI"
        av_tr.Font.Size = 18
        av_tr.Font.Bold = 1
        av_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
        av_tr.ParagraphFormat.Alignment = 2

        # Name & Role next to avatar
        text_x = left + 20 + avatar_size + 16
        text_w = width - (avatar_size + 52)
        n_box = slide.Shapes.AddTextbox(1, text_x, top + 18, text_w, 28)
        n_box.Fill.Background()
        n_tr = n_box.TextFrame.TextRange
        n_tr.Text = name
        n_tr.Font.Name = "Segoe UI"
        n_tr.Font.Size = 16
        n_tr.Font.Bold = 1
        n_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

        r_box = slide.Shapes.AddTextbox(1, text_x, top + 44, text_w, 22)
        r_box.Fill.Background()
        r_tr = r_box.TextFrame.TextRange
        r_tr.Text = role
        r_tr.Font.Name = "Segoe UI"
        r_tr.Font.Size = 12
        r_tr.Font.Bold = 1
        r_tr.Font.Color.RGB = hex_to_rgb_int(accent_color_hex)

        if company_or_dept:
            c_box = slide.Shapes.AddTextbox(1, text_x, top + 64, text_w, 20)
            c_box.Fill.Background()
            c_tr = c_box.TextFrame.TextRange
            c_tr.Text = company_or_dept
            c_tr.Font.Name = "Segoe UI"
            c_tr.Font.Size = 10
            c_tr.Font.Color.RGB = hex_to_rgb_int("#64748B")

        # Divider
        div = slide.Shapes.AddShape(1, left + 20, top + 92, width - 40, 1)
        div.Fill.Solid()
        div.Fill.ForeColor.RGB = hex_to_rgb_int("#F1F5F9")
        div.Line.Visible = 0

        # Bio
        if bio:
            bio_box = slide.Shapes.AddTextbox(1, left + 20, top + 102, width - 40, 100)
            bio_box.Fill.Background()
            bio_tf = bio_box.TextFrame
            bio_tf.WordWrap = -1
            bio_tr = bio_tf.TextRange
            bio_tr.Text = bio
            bio_tr.Font.Name = "Segoe UI"
            bio_tr.Font.Size = 11
            bio_tr.Font.Color.RGB = hex_to_rgb_int("#475569")

        # Skills / tags pills at bottom
        if skills_or_tags:
            tag_x = left + 20
            tag_y = top + height - 36
            for tag in skills_or_tags[:4]:
                tw = min(len(tag) * 8 + 14, 85)
                pill = slide.Shapes.AddShape(5, tag_x, tag_y, tw, 20)
                pill.Fill.Solid()
                pill.Fill.ForeColor.RGB = hex_to_rgb_int("#F1F5F9")
                pill.Line.Visible = 0
                p_tr = pill.TextFrame.TextRange
                p_tr.Text = tag
                p_tr.Font.Name = "Segoe UI"
                p_tr.Font.Size = 9
                p_tr.Font.Color.RGB = hex_to_rgb_int("#475569")
                p_tr.ParagraphFormat.Alignment = 2
                tag_x += tw + 8

        return {
            "status": "success",
            "slide_index": slide_index,
            "name": name,
            "role": role,
        }

    def add_agenda(
        self,
        slide_index: int,
        items: List[Dict[str, Any]],
        left: float = 80,
        top: float = 130,
        width: float = 800,
        height: float = 340,
        active_color_hex: str = "#2563EB",
    ) -> Dict[str, Any]:
        """Create a presentation agenda / table of contents with active topic highlighting."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_items = len(items)
        if num_items == 0:
            return {"status": "error", "message": "items list is empty"}

        row_h = height / num_items

        for i, item in enumerate(items):
            iy = top + (i * row_h)
            num = str(item.get("num", f"{i + 1:02d}"))
            title = str(item.get("title", f"Agenda Item {i + 1}"))
            desc = str(item.get("desc", ""))
            duration = str(item.get("duration", ""))
            is_active = bool(item.get("active", False))

            # Row background
            row_box = slide.Shapes.AddShape(5, left, iy + 4, width, row_h - 8)
            row_box.Fill.Solid()
            row_bg = "#EFF6FF" if is_active else "#FFFFFF"
            border_c = active_color_hex if is_active else "#E2E8F0"
            row_box.Fill.ForeColor.RGB = hex_to_rgb_int(row_bg)
            row_box.Line.Visible = -1
            row_box.Line.ForeColor.RGB = hex_to_rgb_int(border_c)
            row_box.Line.Weight = 1.5 if is_active else 1.0

            # Number badge
            num_pill = slide.Shapes.AddShape(5, left + 14, iy + (row_h - 32) / 2, 36, 32)
            num_pill.Fill.Solid()
            num_pill.Fill.ForeColor.RGB = hex_to_rgb_int(active_color_hex if is_active else "#F1F5F9")
            num_pill.Line.Visible = 0
            n_tr = num_pill.TextFrame.TextRange
            n_tr.Text = num
            n_tr.Font.Name = "Segoe UI"
            n_tr.Font.Size = 13
            n_tr.Font.Bold = 1
            n_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF" if is_active else "#475569")
            n_tr.ParagraphFormat.Alignment = 2

            # Title & Description
            text_x = left + 62
            text_w = width - 180
            t_box = slide.Shapes.AddTextbox(1, text_x, iy + 8, text_w, row_h - 16)
            t_box.Fill.Background()
            t_tf = t_box.TextFrame
            t_tf.WordWrap = -1
            t_tr = t_tf.TextRange
            if desc:
                t_tr.Text = f"{title}\n{desc}"
                p1 = t_tr.Paragraphs(1)
                p1.Font.Name = "Segoe UI"
                p1.Font.Size = 14
                p1.Font.Bold = 1
                p1.Font.Color.RGB = hex_to_rgb_int("#0F172A")

                p2 = t_tr.Paragraphs(2)
                p2.Font.Name = "Segoe UI"
                p2.Font.Size = 11
                p2.Font.Color.RGB = hex_to_rgb_int("#64748B")
            else:
                t_tr.Text = title
                t_tr.Font.Name = "Segoe UI"
                t_tr.Font.Size = 14
                t_tr.Font.Bold = 1
                t_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

            # Duration badge on right
            if duration:
                dur_box = slide.Shapes.AddShape(5, left + width - 110, iy + (row_h - 26) / 2, 90, 26)
                dur_box.Fill.Solid()
                dur_box.Fill.ForeColor.RGB = hex_to_rgb_int("#F8FAFC")
                dur_box.Line.Visible = -1
                dur_box.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
                d_tr = dur_box.TextFrame.TextRange
                d_tr.Text = duration
                d_tr.Font.Name = "Segoe UI"
                d_tr.Font.Size = 10
                d_tr.Font.Color.RGB = hex_to_rgb_int("#64748B")
                d_tr.ParagraphFormat.Alignment = 2

        return {
            "status": "success",
            "slide_index": slide_index,
            "items_count": num_items,
        }

    def add_watermark(
        self,
        text: str = "CONFIDENTIAL",
        slide_index: Optional[int] = None,
        font_size: float = 72,
        color_hex: str = "#E2E8F0",
        rotation: float = -30.0,
    ) -> Dict[str, Any]:
        """Add a semi-transparent diagonal watermark across background of slide(s)."""
        pres = self.get_presentation()
        target_slides = [pres.Slides(slide_index)] if slide_index else [pres.Slides(i) for i in range(1, pres.Slides.Count + 1)]

        for s in target_slides:
            w = pres.PageSetup.SlideWidth
            h = pres.PageSetup.SlideHeight

            box = s.Shapes.AddTextbox(1, 40, h / 3, w - 80, 160)
            box.Fill.Background()
            box.Line.Visible = 0
            box.Rotation = rotation

            tr = box.TextFrame.TextRange
            tr.Text = text.upper()
            tr.Font.Name = "Segoe UI"
            tr.Font.Size = font_size
            tr.Font.Bold = 1
            tr.Font.Color.RGB = hex_to_rgb_int(color_hex)
            tr.ParagraphFormat.Alignment = 2  # Center

            # Send watermark shape to back
            box.ZOrder(1)  # msoSendToBack

        return {
            "status": "success",
            "scope": f"slide {slide_index}" if slide_index else "all slides",
            "text": text,
        }

    def set_shape_zorder(
        self,
        slide_index: int,
        shape_id_or_name: Any,
        action: str = "bring_to_front",
    ) -> Dict[str, Any]:
        """Adjust the z-order layering of a shape ('bring_to_front', 'send_to_back', 'bring_forward', 'send_backward')."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)

        target_shape = None
        for s in slide.Shapes:
            if str(s.Id) == str(shape_id_or_name) or s.Name.lower() == str(shape_id_or_name).lower():
                target_shape = s
                break

        if not target_shape:
            raise ValueError(f"Shape '{shape_id_or_name}' not found on slide {slide_index}")

        zorder_map = {
            "bring_to_front": 0,   # msoBringToFront
            "send_to_back": 1,     # msoSendToBack
            "bring_forward": 2,    # msoBringForward
            "send_backward": 3,    # msoSendBackward
        }
        cmd = zorder_map.get(action.lower())
        if cmd is None:
            raise ValueError(f"Invalid zorder action '{action}'. Use: {list(zorder_map.keys())}")

        target_shape.ZOrder(cmd)
        return {
            "status": "success",
            "shape_name": target_shape.Name,
            "shape_id": target_shape.Id,
            "action": action,
            "slide_index": slide_index,
        }

    def add_comparison_matrix(
        self,
        slide_index: int,
        headers: List[str],
        rows: List[List[str]],
        left: float = 80,
        top: float = 130,
        width: float = 800,
        height: float = 320,
        header_bg_hex: str = "#0F172A",
        highlight_col_index: Optional[int] = 2,
    ) -> Dict[str, Any]:
        """Create a corporate feature comparison matrix / table with zebra stripes and column highlight."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_rows = len(rows) + 1
        num_cols = len(headers)

        table_shape = slide.Shapes.AddTable(num_rows, num_cols, left, top, width, height)
        table = table_shape.Table

        # Style Header
        for c_idx, h_text in enumerate(headers):
            cell = table.Cell(1, c_idx + 1)
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = hex_to_rgb_int(header_bg_hex)
            tr = cell.Shape.TextFrame.TextRange
            tr.Text = str(h_text)
            tr.Font.Name = "Segoe UI"
            tr.Font.Size = 13
            tr.Font.Bold = 1
            tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
            tr.ParagraphFormat.Alignment = 1 if c_idx == 0 else 2

        # Style Rows
        for r_idx, row_data in enumerate(rows):
            is_even = (r_idx % 2 == 0)
            row_bg = "#FFFFFF" if is_even else "#F8FAFC"

            for c_idx in range(num_cols):
                val = row_data[c_idx] if c_idx < len(row_data) else ""
                cell = table.Cell(r_idx + 2, c_idx + 1)
                cell.Shape.Fill.Solid()

                # Highlight column check (1-indexed)
                if highlight_col_index and (c_idx + 1 == highlight_col_index):
                    cell.Shape.Fill.ForeColor.RGB = hex_to_rgb_int("#EFF6FF")
                else:
                    cell.Shape.Fill.ForeColor.RGB = hex_to_rgb_int(row_bg)

                tr = cell.Shape.TextFrame.TextRange
                tr.Text = str(val)
                tr.Font.Name = "Segoe UI"
                tr.Font.Size = 11
                tr.Font.Color.RGB = hex_to_rgb_int("#0F172A" if (highlight_col_index and (c_idx + 1 == highlight_col_index)) else "#334155")
                tr.Font.Bold = 1 if (highlight_col_index and (c_idx + 1 == highlight_col_index)) else 0
                tr.ParagraphFormat.Alignment = 1 if c_idx == 0 else 2

        return {
            "status": "success",
            "slide_index": slide_index,
            "rows": num_rows,
            "cols": num_cols,
            "table_id": table_shape.Id,
        }

    def add_swot_matrix(
        self,
        slide_index: int,
        strengths: List[str],
        weaknesses: List[str],
        opportunities: List[str],
        threats: List[str],
        left: float = 80,
        top: float = 130,
        width: float = 800,
        height: float = 340,
    ) -> Dict[str, Any]:
        """Create a strategic 2x2 SWOT Matrix (Strengths, Weaknesses, Opportunities, Threats)."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        gap = 16
        quad_w = (width - gap) / 2
        quad_h = (height - gap) / 2

        quadrants = [
            {"title": "STRENGTHS", "letter": "S", "items": strengths, "color": "#10B981", "bg": "#F0FDF4", "border": "#BBF7D0", "col": 0, "row": 0},
            {"title": "WEAKNESSES", "letter": "W", "items": weaknesses, "color": "#F59E0B", "bg": "#FFFBEB", "border": "#FDE68A", "col": 1, "row": 0},
            {"title": "OPPORTUNITIES", "letter": "O", "items": opportunities, "color": "#3B82F6", "bg": "#EFF6FF", "border": "#BFDBFE", "col": 0, "row": 1},
            {"title": "THREATS", "letter": "T", "items": threats, "color": "#F43F5E", "bg": "#FFF1F2", "border": "#FECDD3", "col": 1, "row": 1},
        ]

        for q in quadrants:
            qx = left + (q["col"] * (quad_w + gap))
            qy = top + (q["row"] * (quad_h + gap))

            # Quadrant card
            card = slide.Shapes.AddShape(5, qx, qy, quad_w, quad_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = hex_to_rgb_int(q["bg"])
            card.Line.Visible = -1
            card.Line.ForeColor.RGB = hex_to_rgb_int(q["border"])
            card.Line.Weight = 1.0

            # Letter badge pill
            badge = slide.Shapes.AddShape(5, qx + 12, qy + 10, 24, 22)
            badge.Fill.Solid()
            badge.Fill.ForeColor.RGB = hex_to_rgb_int(q["color"])
            badge.Line.Visible = 0
            b_tr = badge.TextFrame.TextRange
            b_tr.Text = q["letter"]
            b_tr.Font.Name = "Segoe UI"
            b_tr.Font.Size = 11
            b_tr.Font.Bold = 1
            b_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
            b_tr.ParagraphFormat.Alignment = 2

            # Title
            t_box = slide.Shapes.AddTextbox(1, qx + 42, qy + 8, quad_w - 54, 24)
            t_box.Fill.Background()
            t_tr = t_box.TextFrame.TextRange
            t_tr.Text = q["title"]
            t_tr.Font.Name = "Segoe UI"
            t_tr.Font.Size = 13
            t_tr.Font.Bold = 1
            t_tr.Font.Color.RGB = hex_to_rgb_int(q["color"])

            # Items bullet list
            items_text = "\n".join([f"- {it}" for it in q["items"]]) if q["items"] else "None specified"
            i_box = slide.Shapes.AddTextbox(1, qx + 12, qy + 36, quad_w - 24, quad_h - 44)
            i_box.Fill.Background()
            i_tf = i_box.TextFrame
            i_tf.WordWrap = -1
            i_tr = i_tf.TextRange
            i_tr.Text = items_text
            i_tr.Font.Name = "Segoe UI"
            i_tr.Font.Size = 11
            i_tr.Font.Color.RGB = hex_to_rgb_int("#334155")

        return {
            "status": "success",
            "slide_index": slide_index,
            "matrix": "SWOT",
        }

    def add_funnel(
        self,
        slide_index: int,
        stages: List[Dict[str, Any]],
        left: float = 120,
        top: float = 130,
        width: float = 720,
        height: float = 340,
        primary_color_hex: str = "#2563EB",
    ) -> Dict[str, Any]:
        """Create a tapered horizontal funnel / pipeline visualization."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_stages = len(stages)
        if num_stages == 0:
            return {"status": "error", "message": "stages list is empty"}

        gap = 10
        row_h = (height - (gap * (num_stages - 1))) / num_stages

        for i, s in enumerate(stages):
            name = str(s.get("name", f"Stage {i + 1}"))
            val = str(s.get("value", ""))
            pct = str(s.get("pct", ""))

            # Taper width from 100% at top down to 55% at bottom
            taper_ratio = 1.0 - (0.42 * i / max(1, num_stages - 1))
            st_w = width * taper_ratio
            st_x = left + (width - st_w) / 2
            st_y = top + (i * (row_h + gap))

            # Funnel tier container
            card = slide.Shapes.AddShape(5, st_x, st_y, st_w, row_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
            card.Line.Visible = -1
            card.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
            card.Line.Weight = 1.0

            # Left accent strip
            stripe = slide.Shapes.AddShape(5, st_x, st_y, 6, row_h)
            stripe.Fill.Solid()
            stripe.Fill.ForeColor.RGB = hex_to_rgb_int(primary_color_hex)
            stripe.Line.Visible = 0

            # Stage Name
            n_box = slide.Shapes.AddTextbox(1, st_x + 16, st_y + (row_h - 24) / 2, st_w * 0.45, 24)
            n_box.Fill.Background()
            n_tr = n_box.TextFrame.TextRange
            n_tr.Text = name
            n_tr.Font.Name = "Segoe UI"
            n_tr.Font.Size = 13
            n_tr.Font.Bold = 1
            n_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

            # Value & Percentage
            if val:
                v_box = slide.Shapes.AddTextbox(1, st_x + (st_w * 0.48), st_y + (row_h - 24) / 2, st_w * 0.28, 24)
                v_box.Fill.Background()
                v_tr = v_box.TextFrame.TextRange
                v_tr.Text = val
                v_tr.Font.Name = "Segoe UI"
                v_tr.Font.Size = 13
                v_tr.Font.Bold = 1
                v_tr.Font.Color.RGB = hex_to_rgb_int(primary_color_hex)
                v_tr.ParagraphFormat.Alignment = 2

            if pct:
                pct_pill = slide.Shapes.AddShape(5, st_x + st_w - 74, st_y + (row_h - 22) / 2, 60, 22)
                pct_pill.Fill.Solid()
                pct_pill.Fill.ForeColor.RGB = hex_to_rgb_int("#EFF6FF")
                pct_pill.Line.Visible = -1
                pct_pill.Line.ForeColor.RGB = hex_to_rgb_int("#BFDBFE")
                p_tr = pct_pill.TextFrame.TextRange
                p_tr.Text = pct
                p_tr.Font.Name = "Segoe UI"
                p_tr.Font.Size = 10
                p_tr.Font.Bold = 1
                p_tr.Font.Color.RGB = hex_to_rgb_int(primary_color_hex)
                p_tr.ParagraphFormat.Alignment = 2

        return {
            "status": "success",
            "slide_index": slide_index,
            "stages_count": num_stages,
        }

    def add_before_after(
        self,
        slide_index: int,
        before_title: str = "Without Automation",
        before_items: Optional[List[str]] = None,
        after_title: str = "With PowerPoint Live MCP",
        after_items: Optional[List[str]] = None,
        left: float = 80,
        top: float = 130,
        width: float = 800,
        height: float = 330,
    ) -> Dict[str, Any]:
        """Create a side-by-side Before vs After transformation comparison."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        before_items = before_items or ["Manual copy-pasting", "Inconsistent formatting", "Slow slide production", "Repetitive design overhead"]
        after_items = after_items or ["Automated 1-click generation", "Pixel-perfect modern design", "Sub-second live updates", "Autonomous AI control"]

        col_w = (width - 60) / 2

        # 1. Before Column (Left)
        c1 = slide.Shapes.AddShape(5, left, top, col_w, height)
        c1.Fill.Solid()
        c1.Fill.ForeColor.RGB = hex_to_rgb_int("#FEF2F2")
        c1.Line.Visible = -1
        c1.Line.ForeColor.RGB = hex_to_rgb_int("#FECACA")
        c1.Line.Weight = 1.0

        b1_tag = slide.Shapes.AddShape(5, left + 18, top + 16, 75, 20)
        b1_tag.Fill.Solid()
        b1_tag.Fill.ForeColor.RGB = hex_to_rgb_int("#EF4444")
        b1_tag.Line.Visible = 0
        bt_tr = b1_tag.TextFrame.TextRange
        bt_tr.Text = "BEFORE"
        bt_tr.Font.Name = "Segoe UI"
        bt_tr.Font.Size = 9
        bt_tr.Font.Bold = 1
        bt_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
        bt_tr.ParagraphFormat.Alignment = 2

        t1 = slide.Shapes.AddTextbox(1, left + 18, top + 42, col_w - 36, 30)
        t1.Fill.Background()
        t1_tr = t1.TextFrame.TextRange
        t1_tr.Text = before_title
        t1_tr.Font.Name = "Segoe UI"
        t1_tr.Font.Size = 16
        t1_tr.Font.Bold = 1
        t1_tr.Font.Color.RGB = hex_to_rgb_int("#991B1B")

        b_text = "\n\n".join([f"- {it}" for it in before_items])
        tx1 = slide.Shapes.AddTextbox(1, left + 18, top + 80, col_w - 36, height - 90)
        tx1.Fill.Background()
        tx1_tf = tx1.TextFrame
        tx1_tf.WordWrap = -1
        tx1_tr = tx1_tf.TextRange
        tx1_tr.Text = b_text
        tx1_tr.Font.Name = "Segoe UI"
        tx1_tr.Font.Size = 12
        tx1_tr.Font.Color.RGB = hex_to_rgb_int("#7F1D1D")

        # 2. Middle Arrow
        mid_x = left + col_w + 12
        mid_y = top + (height / 2) - 15
        arrow = slide.Shapes.AddShape(13, mid_x, mid_y, 36, 30)  # right arrow
        arrow.Fill.Solid()
        arrow.Fill.ForeColor.RGB = hex_to_rgb_int("#2563EB")
        arrow.Line.Visible = 0

        # 3. After Column (Right)
        right_x = left + col_w + 60
        c2 = slide.Shapes.AddShape(5, right_x, top, col_w, height)
        c2.Fill.Solid()
        c2.Fill.ForeColor.RGB = hex_to_rgb_int("#F0FDF4")
        c2.Line.Visible = -1
        c2.Line.ForeColor.RGB = hex_to_rgb_int("#BBF7D0")
        c2.Line.Weight = 1.0

        b2_tag = slide.Shapes.AddShape(5, right_x + 18, top + 16, 65, 20)
        b2_tag.Fill.Solid()
        b2_tag.Fill.ForeColor.RGB = hex_to_rgb_int("#10B981")
        b2_tag.Line.Visible = 0
        b2t_tr = b2_tag.TextFrame.TextRange
        b2t_tr.Text = "AFTER"
        b2t_tr.Font.Name = "Segoe UI"
        b2t_tr.Font.Size = 9
        b2t_tr.Font.Bold = 1
        b2t_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
        b2t_tr.ParagraphFormat.Alignment = 2

        t2 = slide.Shapes.AddTextbox(1, right_x + 18, top + 42, col_w - 36, 30)
        t2.Fill.Background()
        t2_tr = t2.TextFrame.TextRange
        t2_tr.Text = after_title
        t2_tr.Font.Name = "Segoe UI"
        t2_tr.Font.Size = 16
        t2_tr.Font.Bold = 1
        t2_tr.Font.Color.RGB = hex_to_rgb_int("#065F46")

        a_text = "\n\n".join([f"+ {it}" for it in after_items])
        tx2 = slide.Shapes.AddTextbox(1, right_x + 18, top + 80, col_w - 36, height - 90)
        tx2.Fill.Background()
        tx2_tf = tx2.TextFrame
        tx2_tf.WordWrap = -1
        tx2_tr = tx2_tf.TextRange
        tx2_tr.Text = a_text
        tx2_tr.Font.Name = "Segoe UI"
        tx2_tr.Font.Size = 12
        tx2_tr.Font.Color.RGB = hex_to_rgb_int("#064E3B")

        return {
            "status": "success",
            "slide_index": slide_index,
            "before_title": before_title,
            "after_title": after_title,
        }

    def add_stat_hero(
        self,
        slide_index: int,
        number: str,
        label: str,
        trend: Optional[str] = None,
        trend_positive: bool = True,
        details: Optional[str] = None,
        left: float = 80,
        top: float = 140,
        width: float = 380,
        height: float = 240,
        accent_color_hex: str = "#2563EB",
    ) -> Dict[str, Any]:
        """Create a prominent KPI Hero stat card with large number and trend pill."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Main Card
        card = slide.Shapes.AddShape(5, left, top, width, height)
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
        card.Line.Visible = -1
        card.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
        card.Line.Weight = 1.0

        # Top accent strip
        accent = slide.Shapes.AddShape(5, left + 4, top + 2, width - 8, 4)
        accent.Fill.Solid()
        accent.Fill.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
        accent.Line.Visible = 0

        # Big Number
        num_box = slide.Shapes.AddTextbox(1, left + 20, top + 22, width - 40, 60)
        num_box.Fill.Background()
        n_tr = num_box.TextFrame.TextRange
        n_tr.Text = number
        n_tr.Font.Name = "Segoe UI"
        n_tr.Font.Size = 44
        n_tr.Font.Bold = 1
        n_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

        # Trend Badge (if present)
        curr_y = top + 90
        if trend:
            t_color = "#10B981" if trend_positive else "#EF4444"
            t_bg = "#ECFDF5" if trend_positive else "#FEF2F2"
            pill = slide.Shapes.AddShape(5, left + 20, curr_y, min(len(trend) * 8 + 20, 140), 24)
            pill.Fill.Solid()
            pill.Fill.ForeColor.RGB = hex_to_rgb_int(t_bg)
            pill.Line.Visible = -1
            pill.Line.ForeColor.RGB = hex_to_rgb_int(t_color)
            pill.Line.Weight = 1.0
            p_tr = pill.TextFrame.TextRange
            p_tr.Text = trend
            p_tr.Font.Name = "Segoe UI"
            p_tr.Font.Size = 10
            p_tr.Font.Bold = 1
            p_tr.Font.Color.RGB = hex_to_rgb_int(t_color)
            p_tr.ParagraphFormat.Alignment = 2
            curr_y += 34

        # Label / Metric name
        lbl_box = slide.Shapes.AddTextbox(1, left + 20, curr_y, width - 40, 28)
        lbl_box.Fill.Background()
        l_tr = lbl_box.TextFrame.TextRange
        l_tr.Text = label
        l_tr.Font.Name = "Segoe UI"
        l_tr.Font.Size = 14
        l_tr.Font.Bold = 1
        l_tr.Font.Color.RGB = hex_to_rgb_int("#334155")

        # Details
        if details:
            d_box = slide.Shapes.AddTextbox(1, left + 20, curr_y + 30, width - 40, 40)
            d_box.Fill.Background()
            d_tf = d_box.TextFrame
            d_tf.WordWrap = -1
            d_tr = d_tf.TextRange
            d_tr.Text = details
            d_tr.Font.Name = "Segoe UI"
            d_tr.Font.Size = 11
            d_tr.Font.Color.RGB = hex_to_rgb_int("#64748B")

        return {
            "status": "success",
            "slide_index": slide_index,
            "number": number,
            "label": label,
        }

    def add_faq_accordion(
        self,
        slide_index: int,
        qa_pairs: List[Dict[str, str]],
        left: float = 80,
        top: float = 130,
        width: float = 800,
        height: float = 340,
    ) -> Dict[str, Any]:
        """Create a professional FAQ or Q&A card list."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_items = len(qa_pairs)
        if num_items == 0:
            return {"status": "error", "message": "qa_pairs list is empty"}

        gap = 12
        card_h = (height - (gap * (num_items - 1))) / num_items

        for i, qa in enumerate(qa_pairs):
            q_text = str(qa.get("question", f"Question {i + 1}"))
            a_text = str(qa.get("answer", ""))
            cy = top + (i * (card_h + gap))

            # Container card
            card = slide.Shapes.AddShape(5, left, cy, width, card_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
            card.Line.Visible = -1
            card.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
            card.Line.Weight = 1.0

            # "Q" badge
            q_badge = slide.Shapes.AddShape(9, left + 14, cy + 12, 24, 24)
            q_badge.Fill.Solid()
            q_badge.Fill.ForeColor.RGB = hex_to_rgb_int("#2563EB")
            q_badge.Line.Visible = 0
            qb_tr = q_badge.TextFrame.TextRange
            qb_tr.Text = "Q"
            qb_tr.Font.Name = "Segoe UI"
            qb_tr.Font.Size = 11
            qb_tr.Font.Bold = 1
            qb_tr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF")
            qb_tr.ParagraphFormat.Alignment = 2

            # Question Text
            q_box = slide.Shapes.AddTextbox(1, left + 46, cy + 10, width - 60, 24)
            q_box.Fill.Background()
            q_tr = q_box.TextFrame.TextRange
            q_tr.Text = q_text
            q_tr.Font.Name = "Segoe UI"
            q_tr.Font.Size = 13
            q_tr.Font.Bold = 1
            q_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

            # Answer Text
            a_box = slide.Shapes.AddTextbox(1, left + 46, cy + 34, width - 60, card_h - 40)
            a_box.Fill.Background()
            a_tf = a_box.TextFrame
            a_tf.WordWrap = -1
            a_tr = a_tf.TextRange
            a_tr.Text = a_text
            a_tr.Font.Name = "Segoe UI"
            a_tr.Font.Size = 11
            a_tr.Font.Color.RGB = hex_to_rgb_int("#475569")

        return {
            "status": "success",
            "slide_index": slide_index,
            "qa_count": num_items,
        }

    def add_checklist(
        self,
        slide_index: int,
        items: List[Dict[str, Any]],
        left: float = 80,
        top: float = 130,
        width: float = 800,
        height: float = 330,
    ) -> Dict[str, Any]:
        """Create a styled checklist with status badges (completed, in_progress, pending, blocked)."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        num_items = len(items)
        if num_items == 0:
            return {"status": "error", "message": "items list is empty"}

        gap = 10
        row_h = (height - (gap * (num_items - 1))) / num_items

        status_map = {
            "completed": {"icon": "[v]", "color": "#10B981", "bg": "#ECFDF5", "tag": "DONE"},
            "in_progress": {"icon": "[>]", "color": "#2563EB", "bg": "#EFF6FF", "tag": "IN PROGRESS"},
            "pending": {"icon": "[~]", "color": "#F59E0B", "bg": "#FFFBEB", "tag": "PENDING"},
            "blocked": {"icon": "[x]", "color": "#EF4444", "bg": "#FEF2F2", "tag": "BLOCKED"},
        }

        for i, item in enumerate(items):
            text = str(item.get("text", f"Checklist item {i + 1}"))
            status_key = str(item.get("status", "completed")).lower()
            cfg = status_map.get(status_key, status_map["completed"])
            tag = str(item.get("badge", cfg["tag"]))
            iy = top + (i * (row_h + gap))

            # Row Card
            row = slide.Shapes.AddShape(5, left, iy, width, row_h)
            row.Fill.Solid()
            row.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
            row.Line.Visible = -1
            row.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
            row.Line.Weight = 1.0

            # Status Icon badge
            ic_box = slide.Shapes.AddShape(5, left + 14, iy + (row_h - 26) / 2, 32, 26)
            ic_box.Fill.Solid()
            ic_box.Fill.ForeColor.RGB = hex_to_rgb_int(cfg["bg"])
            ic_box.Line.Visible = -1
            ic_box.Line.ForeColor.RGB = hex_to_rgb_int(cfg["color"])
            ic_box.Line.Weight = 1.0
            ic_tr = ic_box.TextFrame.TextRange
            ic_tr.Text = cfg["icon"]
            ic_tr.Font.Name = "Segoe UI"
            ic_tr.Font.Size = 10
            ic_tr.Font.Bold = 1
            ic_tr.Font.Color.RGB = hex_to_rgb_int(cfg["color"])
            ic_tr.ParagraphFormat.Alignment = 2

            # Text
            t_box = slide.Shapes.AddTextbox(1, left + 56, iy + (row_h - 24) / 2, width - 180, 24)
            t_box.Fill.Background()
            t_tr = t_box.TextFrame.TextRange
            t_tr.Text = text
            t_tr.Font.Name = "Segoe UI"
            t_tr.Font.Size = 12
            t_tr.Font.Bold = 1
            t_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

            # Tag pill on right
            if tag:
                pill = slide.Shapes.AddShape(5, left + width - 110, iy + (row_h - 22) / 2, 95, 22)
                pill.Fill.Solid()
                pill.Fill.ForeColor.RGB = hex_to_rgb_int(cfg["bg"])
                pill.Line.Visible = 0
                p_tr = pill.TextFrame.TextRange
                p_tr.Text = tag.upper()
                p_tr.Font.Name = "Segoe UI"
                p_tr.Font.Size = 9
                p_tr.Font.Bold = 1
                p_tr.Font.Color.RGB = hex_to_rgb_int(cfg["color"])
                p_tr.ParagraphFormat.Alignment = 2

        return {
            "status": "success",
            "slide_index": slide_index,
            "checklist_items": num_items,
        }

    def add_review_card(
        self,
        slide_index: int,
        rating: float = 5.0,
        quote: str = "",
        author: str = "Client Reviewer",
        company: str = "",
        platform: str = "G2 Leader",
        left: float = 80,
        top: float = 140,
        width: float = 380,
        height: float = 260,
    ) -> Dict[str, Any]:
        """Create a 5-star customer review or testimonial trust card."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Card
        card = slide.Shapes.AddShape(5, left, top, width, height)
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
        card.Line.Visible = -1
        card.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
        card.Line.Weight = 1.0

        # Star Rating text (e.g. ★★★★★)
        num_stars = min(max(int(round(rating)), 1), 5)
        stars_str = ("★ " * num_stars).strip()
        star_box = slide.Shapes.AddTextbox(1, left + 20, top + 16, 160, 26)
        star_box.Fill.Background()
        s_tr = star_box.TextFrame.TextRange
        s_tr.Text = f"{stars_str}  ({rating:.1f}/5.0)"
        s_tr.Font.Name = "Segoe UI"
        s_tr.Font.Size = 12
        s_tr.Font.Bold = 1
        s_tr.Font.Color.RGB = hex_to_rgb_int("#F59E0B")

        # Platform Pill on right
        if platform:
            plat_pill = slide.Shapes.AddShape(5, left + width - 110, top + 16, 92, 22)
            plat_pill.Fill.Solid()
            plat_pill.Fill.ForeColor.RGB = hex_to_rgb_int("#EFF6FF")
            plat_pill.Line.Visible = -1
            plat_pill.Line.ForeColor.RGB = hex_to_rgb_int("#BFDBFE")
            pl_tr = plat_pill.TextFrame.TextRange
            pl_tr.Text = platform
            pl_tr.Font.Name = "Segoe UI"
            pl_tr.Font.Size = 9
            pl_tr.Font.Bold = 1
            pl_tr.Font.Color.RGB = hex_to_rgb_int("#2563EB")
            pl_tr.ParagraphFormat.Alignment = 2

        # Quote
        q_box = slide.Shapes.AddTextbox(1, left + 20, top + 48, width - 40, 110)
        q_box.Fill.Background()
        q_tf = q_box.TextFrame
        q_tf.WordWrap = -1
        q_tr = q_tf.TextRange
        q_tr.Text = f'"{quote}"'
        q_tr.Font.Name = "Segoe UI"
        q_tr.Font.Size = 12
        q_tr.Font.Italic = 1
        q_tr.Font.Color.RGB = hex_to_rgb_int("#334155")

        # Divider
        div = slide.Shapes.AddShape(1, left + 20, top + height - 60, width - 40, 1)
        div.Fill.Solid()
        div.Fill.ForeColor.RGB = hex_to_rgb_int("#F1F5F9")
        div.Line.Visible = 0

        # Author and company
        a_box = slide.Shapes.AddTextbox(1, left + 20, top + height - 52, width - 40, 40)
        a_box.Fill.Background()
        a_tr = a_box.TextFrame.TextRange
        sub = f" - {company}" if company else ""
        a_tr.Text = f"{author}{sub}"
        a_tr.Font.Name = "Segoe UI"
        a_tr.Font.Size = 11
        a_tr.Font.Bold = 1
        a_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

        return {
            "status": "success",
            "slide_index": slide_index,
            "author": author,
            "rating": rating,
        }

    def style_shape(
        self,
        slide_index: int,
        shape_id_or_name: Any,
        fill_color_hex: Optional[str] = None,
        border_color_hex: Optional[str] = None,
        border_weight: Optional[float] = None,
        opacity: Optional[float] = None,
        has_shadow: Optional[bool] = None,
    ) -> Dict[str, Any]:
        """Format and style an existing shape (color, border, opacity, shadow)."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)

        target_shape = None
        for s in slide.Shapes:
            if str(s.Id) == str(shape_id_or_name) or s.Name.lower() == str(shape_id_or_name).lower():
                target_shape = s
                break

        if not target_shape:
            raise ValueError(f"Shape '{shape_id_or_name}' not found on slide {slide_index}")

        if fill_color_hex:
            target_shape.Fill.Solid()
            target_shape.Fill.ForeColor.RGB = hex_to_rgb_int(fill_color_hex)

        if opacity is not None:
            # PowerPoint COM Fill.Transparency: 0.0 is opaque, 1.0 is transparent
            target_shape.Fill.Transparency = max(0.0, min(1.0, 1.0 - opacity))

        if border_color_hex:
            target_shape.Line.Visible = -1
            target_shape.Line.ForeColor.RGB = hex_to_rgb_int(border_color_hex)

        if border_weight is not None:
            target_shape.Line.Visible = -1 if border_weight > 0 else 0
            if border_weight > 0:
                target_shape.Line.Weight = border_weight

        if has_shadow is not None:
            target_shape.Shadow.Visible = -1 if has_shadow else 0

        return {
            "status": "success",
            "shape_name": target_shape.Name,
            "shape_id": target_shape.Id,
            "slide_index": slide_index,
        }

    def split_slide_to_progressive_deck(
        self,
        slide_index: int,
        shape_ids_or_names: Optional[List[Any]] = None,
        mode: str = "dim",
        highlight_border_color_hex: str = "#2563EB",
        keep_master_slide: bool = False,
    ) -> Dict[str, Any]:
        """Split a multi-element slide into N progressive sequential slides,
        where on each slide one element is highlighted while previous are dimmed and future are hidden/dimmed.
        """
        pres = self.get_presentation()
        master_slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        # Identify targets
        targets = []
        if shape_ids_or_names:
            targets = list(shape_ids_or_names)
        else:
            for s in master_slide.Shapes:
                if s.Top > 100:
                    targets.append(s.Id)

        if not targets:
            raise ValueError("No target shapes found to create progressive deck")

        num_steps = len(targets)
        created_slide_indices = []

        # Duplicate slide for each step
        for k in range(num_steps):
            dup_range = master_slide.Duplicate()
            dup_slide = dup_range.Item(1)
            dup_slide.SlideShowTransition.EntryEffect = 1793  # fade
            dup_slide.SlideShowTransition.Speed = 2  # medium

            for idx, target_ref in enumerate(targets):
                target_shape = None
                for s in dup_slide.Shapes:
                    if str(s.Id) == str(target_ref) or s.Name.lower() == str(target_ref).lower():
                        target_shape = s
                        break

                if not target_shape:
                    continue

                if idx == k:
                    target_shape.Fill.Transparency = 0.0
                    target_shape.Line.Visible = -1
                    target_shape.Line.ForeColor.RGB = hex_to_rgb_int(highlight_border_color_hex)
                    target_shape.Line.Weight = 2.5
                elif idx < k:
                    target_shape.Fill.Transparency = 0.72
                    target_shape.Line.Visible = -1
                    target_shape.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
                    target_shape.Line.Weight = 1.0
                    if target_shape.HasTextFrame:
                        try:
                            target_shape.TextFrame.TextRange.Font.Color.RGB = hex_to_rgb_int("#94A3B8")
                        except Exception:
                            pass
                else:
                    if mode.lower() == "hide":
                        target_shape.Visible = 0
                    else:
                        target_shape.Fill.Transparency = 0.88
                        target_shape.Line.Visible = -1
                        target_shape.Line.ForeColor.RGB = hex_to_rgb_int("#F1F5F9")
                        target_shape.Line.Weight = 1.0
                        if target_shape.HasTextFrame:
                            try:
                                target_shape.TextFrame.TextRange.Font.Color.RGB = hex_to_rgb_int("#CBD5E1")
                            except Exception:
                                pass

            created_slide_indices.append(dup_slide.SlideIndex)

        if not keep_master_slide:
            master_slide.Delete()

        return {
            "status": "success",
            "progressive_slides_count": num_steps,
            "mode": mode,
            "message": f"Created {num_steps} progressive sequential slides",
        }

    def create_progressive_deck(
        self,
        title: str,
        steps: List[Dict[str, str]],
        accent_color_hex: str = "#2563EB",
        bg_color_hex: str = "#FFFFFF",
    ) -> Dict[str, Any]:
        """Generate a complete multi-slide progressive reveal presentation from a list of steps.
        Each slide highlights the current step, advances the top progress tracker, and dims prior steps.
        """
        pres = self.get_presentation()
        num_steps = len(steps)
        if num_steps == 0:
            return {"status": "error", "message": "steps list is empty"}

        w = pres.PageSetup.SlideWidth
        h = pres.PageSetup.SlideHeight

        created_slides = []

        for k in range(num_steps):
            slide = pres.Slides.Add(pres.Slides.Count + 1, 12)  # blank layout
            self._app.ActiveWindow.View.GotoSlide(slide.SlideIndex)
            slide.SlideShowTransition.EntryEffect = 1793  # fade
            slide.SlideShowTransition.Speed = 2

            if bg_color_hex != "#FFFFFF":
                slide.Background.Fill.Solid()
                slide.Background.Fill.ForeColor.RGB = hex_to_rgb_int(bg_color_hex)

            # Slide Main Header
            t_box = slide.Shapes.AddTextbox(1, 80, 40, w - 160, 40)
            t_tr = t_box.TextFrame.TextRange
            t_tr.Text = f"{title}"
            t_tr.Font.Name = "Segoe UI"
            t_tr.Font.Size = 24
            t_tr.Font.Bold = 1
            t_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A")

            # Progress Tracker Header (Step k+1 of N)
            pill_w = (w - 160 - (12 * (num_steps - 1))) / num_steps
            pill_y = 90
            for idx, st in enumerate(steps):
                px = 80 + (idx * (pill_w + 12))
                pill = slide.Shapes.AddShape(5, px, pill_y, pill_w, 28)
                pill.Fill.Solid()

                if idx < k:
                    pill.Fill.ForeColor.RGB = hex_to_rgb_int("#ECFDF5")
                    pill.Line.Visible = -1
                    pill.Line.ForeColor.RGB = hex_to_rgb_int("#A7F3D0")
                    p_text = f"[v] {st.get('title', f'Step {idx+1}')}"
                    p_color = "#10B981"
                elif idx == k:
                    pill.Fill.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
                    pill.Line.Visible = -1
                    pill.Line.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
                    p_text = f"[*] {st.get('title', f'Step {idx+1}')}"
                    p_color = "#FFFFFF"
                else:
                    pill.Fill.ForeColor.RGB = hex_to_rgb_int("#F8FAFC")
                    pill.Line.Visible = -1
                    pill.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
                    p_text = f" {idx+1}. {st.get('title', f'Step {idx+1}')}"
                    p_color = "#94A3B8"

                ptr = pill.TextFrame.TextRange
                ptr.Text = p_text
                ptr.Font.Name = "Segoe UI"
                ptr.Font.Size = 10
                ptr.Font.Bold = 1 if idx == k else 0
                ptr.Font.Color.RGB = hex_to_rgb_int(p_color)
                ptr.ParagraphFormat.Alignment = 2

            # Main Content Area: Cards for each step
            card_gap = 18
            card_w = (w - 160 - (card_gap * (num_steps - 1))) / num_steps
            card_y = 140
            card_h = 320

            for idx, st in enumerate(steps):
                cx = 80 + (idx * (card_w + card_gap))
                card = slide.Shapes.AddShape(5, cx, card_y, card_w, card_h)
                card.Fill.Solid()

                is_active = (idx == k)
                is_past = (idx < k)

                if is_active:
                    card.Fill.ForeColor.RGB = hex_to_rgb_int("#FFFFFF")
                    card.Line.Visible = -1
                    card.Line.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
                    card.Line.Weight = 2.5
                    top_bar = slide.Shapes.AddShape(5, cx + 4, card_y + 2, card_w - 8, 4)
                    top_bar.Fill.Solid()
                    top_bar.Fill.ForeColor.RGB = hex_to_rgb_int(accent_color_hex)
                    top_bar.Line.Visible = 0
                elif is_past:
                    card.Fill.ForeColor.RGB = hex_to_rgb_int("#F8FAFC")
                    card.Line.Visible = -1
                    card.Line.ForeColor.RGB = hex_to_rgb_int("#E2E8F0")
                    card.Line.Weight = 1.0
                else:
                    card.Fill.ForeColor.RGB = hex_to_rgb_int("#FAFAFA")
                    card.Fill.Transparency = 0.5
                    card.Line.Visible = -1
                    card.Line.ForeColor.RGB = hex_to_rgb_int("#F1F5F9")
                    card.Line.Weight = 1.0

                # Step Number Pill
                np = slide.Shapes.AddShape(9, cx + 16, card_y + 16, 32, 32)
                np.Fill.Solid()
                np_bg = accent_color_hex if is_active else ("#10B981" if is_past else "#E2E8F0")
                np.Fill.ForeColor.RGB = hex_to_rgb_int(np_bg)
                np.Line.Visible = 0
                nptr = np.TextFrame.TextRange
                nptr.Text = "[v]" if is_past else f"{idx+1}"
                nptr.Font.Name = "Segoe UI"
                nptr.Font.Size = 12
                nptr.Font.Bold = 1
                nptr.Font.Color.RGB = hex_to_rgb_int("#FFFFFF" if (is_active or is_past) else "#94A3B8")
                nptr.ParagraphFormat.Alignment = 2

                # Step Title
                st_box = slide.Shapes.AddTextbox(1, cx + 16, card_y + 56, card_w - 32, 40)
                st_tr = st_box.TextFrame.TextRange
                st_tr.Text = st.get("title", f"Step {idx+1}")
                st_tr.Font.Name = "Segoe UI"
                st_tr.Font.Size = 15 if is_active else 13
                st_tr.Font.Bold = 1
                st_tr.Font.Color.RGB = hex_to_rgb_int("#0F172A" if is_active else ("#475569" if is_past else "#94A3B8"))

                # Step Description
                sd_box = slide.Shapes.AddTextbox(1, cx + 16, card_y + 100, card_w - 32, card_h - 110)
                sd_tf = sd_box.TextFrame
                sd_tf.WordWrap = -1
                sd_tr = sd_tf.TextRange
                sd_tr.Text = st.get("description", "")
                sd_tr.Font.Name = "Segoe UI"
                sd_tr.Font.Size = 12 if is_active else 11
                sd_tr.Font.Color.RGB = hex_to_rgb_int("#334155" if is_active else ("#64748B" if is_past else "#CBD5E1"))

            created_slides.append(slide.SlideIndex)

        return {
            "status": "success",
            "slides_count": num_steps,
            "created_slides": created_slides,
            "title": title,
        }

    def animate_sequence(
        self,
        slide_index: int,
        shape_ids_or_names: List[Any],
        effect: str = "fade",
        trigger: str = "on_click",
        duration: float = 0.5,
    ) -> Dict[str, Any]:
        """Apply sequential entrance animations to a list of shapes on a single slide."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        seq = slide.TimeLine.MainSequence
        effect_code = ANIMATION_EFFECTS.get(effect.lower(), 10)

        trigger_type = 1
        if trigger.lower() == "after_previous":
            trigger_type = 3
        elif trigger.lower() == "with_previous":
            trigger_type = 2

        animated = []
        for i, shape_ref in enumerate(shape_ids_or_names):
            target_shape = None
            for s in slide.Shapes:
                if str(s.Id) == str(shape_ref) or s.Name.lower() == str(shape_ref).lower():
                    target_shape = s
                    break

            if not target_shape:
                continue

            eff = seq.AddEffect(target_shape, effect_code, 0, trigger_type)
            eff.Timing.Duration = duration
            animated.append(target_shape.Name)

        return {
            "status": "success",
            "slide_index": slide_index,
            "animated_count": len(animated),
            "shapes": animated,
            "effect": effect,
            "trigger": trigger,
        }

    def animate_bullet_points(
        self,
        slide_index: int,
        shape_id_or_name: Any,
        effect: str = "fade",
    ) -> Dict[str, Any]:
        """Animate bullet points or paragraphs in a textbox one by one on click."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        target_shape = None
        for s in slide.Shapes:
            if str(s.Id) == str(shape_id_or_name) or s.Name.lower() == str(shape_id_or_name).lower():
                target_shape = s
                break

        if not target_shape:
            raise ValueError(f"Shape '{shape_id_or_name}' not found on slide {slide_index}")

        effect_code = ANIMATION_EFFECTS.get(effect.lower(), 10)
        eff = slide.TimeLine.MainSequence.AddEffect(target_shape, effect_code, 1, 1)

        return {
            "status": "success",
            "slide_index": slide_index,
            "shape_name": target_shape.Name,
            "effect": effect,
        }

    def animate_emphasis(
        self,
        slide_index: int,
        shape_id_or_name: Any,
        effect: str = "pulse",
        trigger: str = "on_click",
    ) -> Dict[str, Any]:
        """Add emphasis animation (pulse, grow_shrink, spin, teeter) to a shape."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        target_shape = None
        for s in slide.Shapes:
            if str(s.Id) == str(shape_id_or_name) or s.Name.lower() == str(shape_id_or_name).lower():
                target_shape = s
                break

        if not target_shape:
            raise ValueError(f"Shape '{shape_id_or_name}' not found on slide {slide_index}")

        emphasis_map = {
            "pulse": 29,
            "grow_shrink": 31,
            "spin": 28,
            "teeter": 33,
            "color_wave": 34,
        }
        eff_code = emphasis_map.get(effect.lower(), 29)
        trigger_type = 3 if trigger.lower() == "after_previous" else (2 if trigger.lower() == "with_previous" else 1)

        eff = slide.TimeLine.MainSequence.AddEffect(target_shape, eff_code, 0, trigger_type)
        return {
            "status": "success",
            "slide_index": slide_index,
            "shape_name": target_shape.Name,
            "effect": effect,
        }

    def animate_exit(
        self,
        slide_index: int,
        shape_id_or_name: Any,
        effect: str = "fade",
        trigger: str = "on_click",
    ) -> Dict[str, Any]:
        """Add exit animation (fade_out, disappear, fly_out) to dismiss a shape."""
        pres = self.get_presentation()
        slide = pres.Slides(slide_index)
        self._app.ActiveWindow.View.GotoSlide(slide_index)

        target_shape = None
        for s in slide.Shapes:
            if str(s.Id) == str(shape_id_or_name) or s.Name.lower() == str(shape_id_or_name).lower():
                target_shape = s
                break

        if not target_shape:
            raise ValueError(f"Shape '{shape_id_or_name}' not found on slide {slide_index}")

        eff_code = ANIMATION_EFFECTS.get(effect.lower(), 10)
        trigger_type = 3 if trigger.lower() == "after_previous" else 1
        eff = slide.TimeLine.MainSequence.AddEffect(target_shape, eff_code, 0, trigger_type)
        eff.Exit = -1

        return {
            "status": "success",
            "slide_index": slide_index,
            "shape_name": target_shape.Name,
            "exit_effect": effect,
        }

    def clear_animations(self, slide_index: Optional[int] = None) -> Dict[str, Any]:
        """Clear all animation effects from a slide or all slides."""
        pres = self.get_presentation()
        target_slides = [pres.Slides(slide_index)] if slide_index else [pres.Slides(i) for i in range(1, pres.Slides.Count + 1)]

        cleared_total = 0
        for s in target_slides:
            seq = s.TimeLine.MainSequence
            cnt = seq.Count
            for i in range(cnt, 0, -1):
                seq(i).Delete()
            cleared_total += cnt

        return {
            "status": "success",
            "cleared_effects_count": cleared_total,
            "scope": f"slide {slide_index}" if slide_index else "all slides",
        }

    def add_animated_stepper(
        self,
        slide_index: int,
        steps: List[Dict[str, str]],
        left: float = 80,
        top: float = 150,
        width: float = 800,
        height: float = 260,
        animate: bool = True,
        trigger: str = "on_click",
    ) -> Dict[str, Any]:
        """Create a process stepper pipeline with automatic sequential entrance animations."""
        res = self.add_process_steps(
            slide_index=slide_index,
            steps=steps,
            left=left,
            top=top,
            width=width,
            height=height,
        )

        if animate:
            pres = self.get_presentation()
            slide = pres.Slides(slide_index)
            recent_shapes = [slide.Shapes(i) for i in range(max(1, slide.Shapes.Count - (len(steps) * 4)), slide.Shapes.Count + 1)]
            seq = slide.TimeLine.MainSequence
            trig_code = 1 if trigger.lower() == "on_click" else 3

            for s in recent_shapes:
                if s.Type == 1 or s.Type == 5:
                    eff = seq.AddEffect(s, 10, 0, trig_code)
                    eff.Timing.Duration = 0.4

        return {
            "status": "success",
            "slide_index": slide_index,
            "steps_count": len(steps),
            "animated": animate,
        }


# Global singleton controller instance
controller = PowerPointController()


