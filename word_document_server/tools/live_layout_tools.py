"""COM-based layout tools for Microsoft Word.

These tools operate on documents currently open in Word via COM automation.
They provide layout, header/footer, spacing, bookmark, watermark, and section
management for files that are open (and locked) in Word.
"""

import json
import sys

# macOS JXA dispatch
_MAC_AVAILABLE = sys.platform == 'darwin'

# 1 inch = 72 points (avoid app.InchesToPoints which can fail on some COM setups)
_PTS_PER_INCH = 72.0


async def word_live_set_page_layout(
    filename: str = None,
    section_index: int = 1,
    orientation: str = None,
    page_width_inches: float = None,
    page_height_inches: float = None,
    margin_top_inches: float = None,
    margin_bottom_inches: float = None,
    margin_left_inches: float = None,
    margin_right_inches: float = None,
) -> str:
    """Set page layout for a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed, COM style). Default 1.
        orientation: "portrait" or "landscape".
        page_width_inches: Page width in inches.
        page_height_inches: Page height in inches.
        margin_top_inches: Top margin in inches.
        margin_bottom_inches: Bottom margin in inches.
        margin_left_inches: Left margin in inches.
        margin_right_inches: Right margin in inches.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_set_page_layout
        # Mac JXA page setup values are in points; the tool contract is inches.
        def _pts(inches):
            return None if inches is None else inches * _PTS_PER_INCH

        return mac_set_page_layout(
            filename=filename, section_index=section_index, orientation=orientation,
            page_width=_pts(page_width_inches), page_height=_pts(page_height_inches),
            top_margin=_pts(margin_top_inches), bottom_margin=_pts(margin_bottom_inches),
            left_margin=_pts(margin_left_inches), right_margin=_pts(margin_right_inches),
        )

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        app = doc.Application

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Set Page Layout"):
            ps = doc.Sections(section_index).PageSetup
            changes = []

            if orientation is not None:
                # wdOrientPortrait=0, wdOrientLandscape=1
                if orientation.lower() == "landscape":
                    ps.Orientation = 1
                    changes.append("orientation=landscape")
                elif orientation.lower() == "portrait":
                    ps.Orientation = 0
                    changes.append("orientation=portrait")

            if page_width_inches is not None:
                ps.PageWidth = page_width_inches * _PTS_PER_INCH
                changes.append(f"width={page_width_inches}in")
            if page_height_inches is not None:
                ps.PageHeight = page_height_inches * _PTS_PER_INCH
                changes.append(f"height={page_height_inches}in")
            if margin_top_inches is not None:
                ps.TopMargin = margin_top_inches * _PTS_PER_INCH
                changes.append(f"margin_top={margin_top_inches}in")
            if margin_bottom_inches is not None:
                ps.BottomMargin = margin_bottom_inches * _PTS_PER_INCH
                changes.append(f"margin_bottom={margin_bottom_inches}in")
            if margin_left_inches is not None:
                ps.LeftMargin = margin_left_inches * _PTS_PER_INCH
                changes.append(f"margin_left={margin_left_inches}in")
            if margin_right_inches is not None:
                ps.RightMargin = margin_right_inches * _PTS_PER_INCH
                changes.append(f"margin_right={margin_right_inches}in")

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "changes": changes,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_header_footer(
    filename: str = None,
    section_index: int = 1,
    header_text: str = None,
    footer_text: str = None,
    header_alignment: str = "center",
    footer_alignment: str = "center",
) -> str:
    """Add header and/or footer text to a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed).
        header_text: Text for the header. None = don't change.
        footer_text: Text for the footer. None = don't change.
        header_alignment: "left", "center", "right".
        footer_alignment: "left", "center", "right".

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_header_footer
        return mac_add_header_footer(filename=filename, section_index=section_index, header_text=header_text, footer_text=footer_text)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        app = doc.Application

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Add Header/Footer"):
            # Alignment map: 0=left, 1=center, 2=right
            align_map = {"left": 0, "center": 1, "right": 2}
            added = []

            # wdHeaderFooterPrimary = 1
            section = doc.Sections(section_index)

            if header_text is not None:
                hdr = section.Headers(1)  # Primary header
                hdr.Range.Text = header_text
                hdr.Range.ParagraphFormat.Alignment = align_map.get(
                    header_alignment.lower(), 1
                )
                added.append("header")

            if footer_text is not None:
                ftr = section.Footers(1)  # Primary footer
                ftr.Range.Text = footer_text
                ftr.Range.ParagraphFormat.Alignment = align_map.get(
                    footer_alignment.lower(), 1
                )
                added.append("footer")

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "added": added,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_page_numbers(
    filename: str = None,
    section_index: int = 1,
    position: str = "footer",
    alignment: str = "center",
    prefix: str = "",
    suffix: str = "",
    include_total: bool = False,
) -> str:
    """Add page numbers to header or footer in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed).
        position: "header" or "footer".
        alignment: "left", "center", "right".
        prefix: Text before page number.
        suffix: Text after page number.
        include_total: If True, adds " / N" after page number.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_page_numbers
        return mac_add_page_numbers(filename=filename, section_index=section_index, position=position, alignment=alignment, prefix=prefix, suffix=suffix, include_total=include_total)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        app = doc.Application

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        if position not in ('header', 'footer') or alignment not in ('left', 'center', 'right'):
            raise ValueError('Invalid page-number position or alignment')
        from word_document_server.utils.text_safety import reject_control_chars
        reject_control_chars('prefix', prefix)
        reject_control_chars('suffix', suffix)
        section = doc.Sections(section_index)
        target = section.Headers(1) if position == 'header' else section.Footers(1)
        if any(field.Type == 33 for field in target.Range.Fields):
            raise ValueError('This header/footer already has a PAGE field; edit it instead of adding a duplicate')
        with undo_record(app, "MCP: Add Page Numbers"):
            # Isolate the requested section while retaining its existing contents.
            if section_index > 1:
                target.LinkToPrevious = False
            rng = target.Range.Duplicate
            rng.End -= 1
            rng.Collapse(0)
            if target.Range.Text.rstrip('\r\x07'):
                rng.InsertAfter('\r')
                rng.Collapse(0)
            rng.ParagraphFormat.Alignment = {'left': 0, 'center': 1, 'right': 2}[alignment]
            if prefix:
                rng.InsertAfter(prefix)
                rng.Collapse(0)
            field = target.Range.Fields.Add(Range=rng, Type=33)
            rng = field.Result.Duplicate
            rng.Collapse(0)
            # Move beyond the field terminator before appending text/fields.
            rng.SetRange(rng.End + 1, rng.End + 1)
            if include_total:
                rng.InsertAfter(' / ')
                rng.Collapse(0)
                field = target.Range.Fields.Add(Range=rng, Type=26)
                rng = field.Result.Duplicate
                rng.Collapse(0)
                rng.SetRange(rng.End + 1, rng.End + 1)
            if suffix:
                rng.InsertAfter(suffix)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "position": position,
            "alignment": alignment,
            "include_total": include_total,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_section_break(
    filename: str = None,
    break_type: str = "new_page",
) -> str:
    """Add a section break to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        break_type: "new_page", "continuous", "even_page", "odd_page".

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_section_break
        return mac_add_section_break(filename=filename, break_type=break_type)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        app = doc.Application

        # wdSectionBreakNextPage=2, Continuous=3, EvenPage=4, OddPage=5
        type_map = {
            "new_page": 2,
            "continuous": 3,
            "even_page": 4,
            "odd_page": 5,
        }

        if break_type not in type_map:
            return json.dumps({
                "error": f"Invalid break_type: {break_type}. Use: {list(type_map.keys())}"
            })

        with undo_record(app, "MCP: Add Section Break"):
            # Insert at end of document
            end_pos = doc.Content.End - 1
            rng = doc.Range(end_pos, end_pos)
            rng.InsertBreak(Type=type_map[break_type])

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "break_type": break_type,
            "total_sections": doc.Sections.Count,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_paragraph_spacing(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    space_before_pt: float = None,
    space_after_pt: float = None,
    line_spacing: float = None,
    line_spacing_rule: str = None,
    keep_with_next: bool = None,
    keep_together: bool = None,
    alignment: str = None,
) -> str:
    """Set paragraph spacing and layout properties in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed). Ignored if start/end given.
        start_paragraph: Start of range (1-indexed, inclusive).
        end_paragraph: End of range (1-indexed, inclusive).
        space_before_pt: Space before paragraph in points.
        space_after_pt: Space after paragraph in points.
        line_spacing: Line spacing value IN POINTS (depends on rule).
            IMPORTANT: For "multiple" rule, value is in points, NOT a multiplier.
            Single spacing (1.0) = 12pt. So: 1.15 lines = 13.8pt, 1.5 lines = 18pt,
            2.0 lines = 24pt. Formula: desired_lines * 12 = points_value.
        line_spacing_rule: "single"(0), "1.5_lines"(1), "double"(2),
                           "at_least"(3), "exactly"(4), "multiple"(5).
        keep_with_next: Keep paragraph with next paragraph on same page (True/False).
        keep_together: Keep all lines of paragraph on same page (True/False).
        alignment: Paragraph alignment - "left"(0), "center"(1), "right"(2), "justify"(3).

    Returns:
        JSON with count of affected paragraphs.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_set_paragraph_spacing
        return mac_set_paragraph_spacing(filename=filename, paragraph_index=paragraph_index, start_paragraph=start_paragraph, end_paragraph=end_paragraph, space_before=space_before_pt, space_after=space_after_pt, line_spacing=line_spacing, keep_with_next=keep_with_next, keep_together=keep_together, alignment=alignment)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        app = doc.Application

        total = doc.Paragraphs.Count

        # wdLineSpacing rules
        rule_map = {
            "single": 0,
            "1.5_lines": 1,
            "double": 2,
            "at_least": 3,
            "exactly": 4,
            "multiple": 5,
        }

        import math
        if line_spacing_rule is not None and line_spacing_rule not in rule_map:
            raise ValueError('Invalid line_spacing_rule')
        if alignment is not None and alignment not in ('left', 'center', 'right', 'justify'):
            raise ValueError('Invalid alignment')
        for value in (space_before_pt, space_after_pt, line_spacing):
            if value is not None and (not math.isfinite(value) or value < 0):
                raise ValueError('Spacing must be finite and non-negative')
        if (start_paragraph is None) != (end_paragraph is None):
            raise ValueError('Supply both start_paragraph and end_paragraph')
        # Determine range of paragraphs (1-indexed)
        if start_paragraph is not None and end_paragraph is not None:
            if not 1 <= start_paragraph <= end_paragraph <= total:
                raise ValueError("Invalid paragraph range")
            indices = range(start_paragraph, end_paragraph + 1)
        elif paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > total:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{total})"
                })
            indices = [paragraph_index]
        else:
            indices = range(1, total + 1)

        with undo_record(app, "MCP: Set Paragraph Spacing"):
            count = 0
            for i in indices:
                pf = doc.Paragraphs(i).Format
                if space_before_pt is not None:
                    pf.SpaceBefore = space_before_pt
                if space_after_pt is not None:
                    pf.SpaceAfter = space_after_pt
                if line_spacing_rule is not None and line_spacing_rule in rule_map:
                    pf.LineSpacingRule = rule_map[line_spacing_rule]
                if line_spacing is not None:
                    pf.LineSpacing = line_spacing
                if keep_with_next is not None:
                    pf.KeepWithNext = keep_with_next
                if keep_together is not None:
                    pf.KeepTogether = keep_together
                if alignment is not None:
                    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                    if alignment in align_map:
                        pf.Alignment = align_map[alignment]
                count += 1

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_affected": count,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_bookmark(
    filename: str = None,
    paragraph_index: int = 1,
    bookmark_name: str = "",
) -> str:
    """Add a named bookmark at a paragraph in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Paragraph to bookmark (1-indexed).
        bookmark_name: Bookmark name (alphanumeric + underscore, no spaces).

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_bookmark
        return mac_add_bookmark(filename=filename, paragraph_index=paragraph_index, bookmark_name=bookmark_name)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    if not bookmark_name:
        return json.dumps({"error": "bookmark_name is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        app = doc.Application

        if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
            return json.dumps({
                "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
            })

        with undo_record(app, "MCP: Add Bookmark"):
            rng = doc.Paragraphs(paragraph_index).Range
            doc.Bookmarks.Add(bookmark_name, rng)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "bookmark_name": bookmark_name,
            "paragraph_index": paragraph_index,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_watermark(
    filename: str = None,
    text: str = "TASLAK",
    font_size: int = 72,
    font_color: str = "C0C0C0",
    rotation: int = -45,
    section_index: int = 1,
) -> str:
    """Add a diagonal text watermark to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        text: Watermark text (e.g. "TASLAK", "DRAFT", "GİZLİ").
        font_size: Font size in points.
        font_color: Hex color without # (e.g. "C0C0C0").
        rotation: Rotation angle in degrees (e.g. -45).
        section_index: Section number (1-indexed).

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_add_watermark is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        app = doc.Application

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Add Watermark"):
            section = doc.Sections(section_index)
            header = section.Headers(1)  # wdHeaderFooterPrimary

            # Parse color
            c = font_color.lstrip("#")
            r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
            rgb_color = r + (g << 8) + (b << 16)

            # AddTextEffect(PresetTextEffect, Text, FontName, FontSize,
            #               FontBold, FontItalic, Left, Top)
            # COM requires positional args
            shape = header.Shapes.AddTextEffect(
                0, text, "Calibri", font_size, False, False, 0, 0
            )

            # Configure shape
            shape.Fill.ForeColor.RGB = rgb_color
            shape.Fill.Transparency = 0.5
            shape.Line.Visible = False  # msoFalse
            shape.Rotation = rotation
            shape.LockAspectRatio = False

            # Position relative to page center
            # msoRelativeHorizontalPositionMargin = 0
            # msoRelativeVerticalPositionMargin = 0
            shape.RelativeHorizontalPosition = 0
            shape.RelativeVerticalPosition = 0
            shape.Left = -999995  # wdShapeCenter (magic value for centering)
            shape.Top = -999995  # wdShapeCenter

            # Send behind text
            shape.WrapFormat.Type = 3  # wdWrapBehind
            shape.WrapFormat.AllowOverlap = True

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "text": text,
            "font_size": font_size,
            "color": font_color,
            "rotation": rotation,
            "section": section_index,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})
