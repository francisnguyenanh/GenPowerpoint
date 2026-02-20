"""
deep_scanner.py
~~~~~~~~~~~~~~~
Deep-scan a PowerPoint master file and produce two JSON-ready dicts that
mirror the hand-crafted ``master_profile.json`` and ``master_structure.json``
formats — but derived entirely from the PPTX itself.

Public API
----------
deep_scan(pptx_path) -> dict
    Returns::

        {
          "master_profile": { ... },   # same schema as master_profile.json
          "master_structure": { ... },  # same schema as master_structure.json
        }
"""

from __future__ import annotations

import os
from collections import Counter
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn

# ── Placeholder-type categories ─────────────────────────────────────────────

_FIXED_PH_TYPES = {
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.SLIDE_NUMBER,
}

_TITLE_PH_TYPES = {
    PP_PLACEHOLDER.TITLE,
    PP_PLACEHOLDER.CENTER_TITLE,
    PP_PLACEHOLDER.VERTICAL_TITLE,
}

_BODY_PH_TYPES = {
    PP_PLACEHOLDER.BODY,
    PP_PLACEHOLDER.OBJECT,
    PP_PLACEHOLDER.SUBTITLE,
    PP_PLACEHOLDER.TABLE,
    PP_PLACEHOLDER.CHART,
    PP_PLACEHOLDER.ORG_CHART,
    PP_PLACEHOLDER.MEDIA_CLIP,
}

_PICTURE_PH_TYPES = {
    PP_PLACEHOLDER.PICTURE,
}


# ── Helpers ──────────────────────────────────────────────────────────────────

def _emu_to_in(emu) -> float | None:
    """EMU → inches, rounded to 4 dp."""
    if emu is None:
        return None
    return round(int(emu) / 914400, 4)


def _emu_to_pt(emu) -> float | None:
    """EMU → points, rounded to 2 dp."""
    if emu is None:
        return None
    return round(int(emu) / 12700, 2)


def _rgb_hex(color) -> str | None:
    """Try to extract '#RRGGBB' from a ColorFormat. Returns None on failure."""
    if color is None:
        return None
    try:
        rgb = color.rgb
        if rgb is not None:
            return f"#{rgb}"
    except Exception:
        pass
    return None


def _theme_color_name(color) -> str | None:
    """Try to extract the theme-color enum name."""
    try:
        tc = color.theme_color
        if tc is not None:
            return tc.name
    except Exception:
        pass
    return None


def _font_info(font) -> dict[str, Any]:
    """Extract font attributes into a compact dict."""
    d: dict[str, Any] = {}
    try:
        d["name"] = font.name
    except Exception:
        d["name"] = None
    try:
        d["size"] = _emu_to_pt(font.size) if font.size else None
    except Exception:
        d["size"] = None
    try:
        d["bold"] = font.bold
    except Exception:
        d["bold"] = None
    try:
        d["italic"] = font.italic
    except Exception:
        d["italic"] = None
    try:
        hex_val = _rgb_hex(font.color)
        d["color"] = hex_val
        d["theme_color"] = _theme_color_name(font.color)
    except Exception:
        d["color"] = None
        d["theme_color"] = None
    return d


def _para_info(para) -> dict[str, Any]:
    """Extract paragraph formatting."""
    d: dict[str, Any] = {}
    try:
        fmt = para.paragraph_format
        try:
            d["alignment"] = fmt.alignment.name if fmt.alignment else None
        except Exception:
            d["alignment"] = None
        try:
            ls = fmt.line_spacing
            d["line_spacing"] = float(ls) if ls is not None else None
        except Exception:
            d["line_spacing"] = None
        try:
            d["space_before"] = _emu_to_pt(fmt.space_before) if fmt.space_before else None
        except Exception:
            d["space_before"] = None
        try:
            d["space_after"] = _emu_to_pt(fmt.space_after) if fmt.space_after else None
        except Exception:
            d["space_after"] = None
    except AttributeError:
        # Some paragraph objects (layout-level) don't have paragraph_format
        d["alignment"] = None
        d["line_spacing"] = None
        d["space_before"] = None
        d["space_after"] = None
    try:
        d["level"] = para.level
    except Exception:
        d["level"] = 0
    # Bullet info from XML
    d["bullet"] = _bullet_info(para)
    return d


def _bullet_info(para) -> dict | None:
    """Extract bullet/numbering info from the paragraph XML."""
    pPr = para._p.find(qn("a:pPr"))
    if pPr is None:
        return None
    buNone = pPr.find(qn("a:buNone"))
    if buNone is not None:
        return {"type": "none"}
    buChar = pPr.find(qn("a:buChar"))
    if buChar is not None:
        return {"type": "char", "char": buChar.get("char", "•")}
    buAutoNum = pPr.find(qn("a:buAutoNum"))
    if buAutoNum is not None:
        return {"type": "auto_num", "scheme": buAutoNum.get("type", "")}
    buFont = pPr.find(qn("a:buFont"))
    buClr = pPr.find(qn("a:buClr"))
    buSzPct = pPr.find(qn("a:buSzPct"))
    if buFont is not None or buClr is not None or buSzPct is not None:
        info: dict[str, Any] = {"type": "formatted"}
        if buFont is not None:
            info["font"] = buFont.get("typeface", "")
        return info
    return None


def _shape_color(shape) -> str | None:
    """Try to get fill color of a non-placeholder shape."""
    try:
        fill = shape.fill
        if fill.type is not None:
            fc = fill.fore_color
            if fc and fc.rgb:
                return f"#{fc.rgb}"
    except Exception:
        pass
    return None


def _extract_placeholder_detail(ph, master_styles: dict | None = None,
                                 theme_fonts: dict | None = None,
                                 theme_colors: dict | None = None) -> dict[str, Any]:
    """Build a detailed placeholder descriptor for the profile JSON.

    Cascades font info: placeholder XML → master text styles → theme fonts.
    """
    ph_fmt = ph.placeholder_format
    ph_type = ph_fmt.type
    idx = ph_fmt.idx

    d: dict[str, Any] = {
        "idx": idx,
        "name": ph.name,
        "ph_type": ph_type.name if ph_type else None,
        "pos": {
            "left":   _emu_to_in(ph.left),
            "top":    _emu_to_in(ph.top),
            "width":  _emu_to_in(ph.width),
            "height": _emu_to_in(ph.height),
        },
    }

    # ── 1. Extract font directly from placeholder XML ─────────────────────
    xml_font = _extract_placeholder_xml_font(ph)

    # ── 2. Determine master style to cascade from ─────────────────────────
    is_title = (ph_type in _TITLE_PH_TYPES) or idx == 0
    cascade_style = None
    if master_styles:
        if is_title and "titleStyle" in master_styles:
            cascade_style = master_styles["titleStyle"]
        elif "bodyStyle" in master_styles:
            cascade_style = master_styles["bodyStyle"]

    # ── 3. Build merged font: XML > master style > theme ──────────────────
    merged: dict[str, Any] = {}

    # Start with theme font as base
    if theme_fonts:
        base_font = theme_fonts.get("major_latin") if is_title else theme_fonts.get("minor_latin")
        if base_font:
            merged["name"] = base_font
        ea_font = theme_fonts.get("major_ea") if is_title else theme_fonts.get("minor_ea")
        if ea_font:
            merged["ea_name"] = ea_font

    # Overlay master text style
    if cascade_style:
        for k in ("name", "ea_name", "size", "bold", "color"):
            if cascade_style.get(k) is not None:
                merged[k] = cascade_style[k]

    # Overlay placeholder-level XML (highest priority)
    for k in ("name", "ea_name", "size", "bold", "color", "align"):
        if xml_font.get(k) is not None:
            merged[k] = xml_font[k]

    # Track whether the color came directly from the placeholder XML
    # (not cascaded from master text style or theme defaults)
    xml_color_found = bool(xml_font.get("color"))

    # Resolve scheme colors to hex — including common OOXML aliases
    _SCHEME_ALIASES = {"tx1": "dk1", "tx2": "dk2", "bg1": "lt1", "bg2": "lt2"}
    if isinstance(merged.get("color"), str) and merged["color"].startswith("scheme:"):
        scheme_key = merged["color"].split(":")[1]
        resolved = None
        if theme_colors:
            resolved = (theme_colors.get(scheme_key)
                        or theme_colors.get(_SCHEME_ALIASES.get(scheme_key, "")))
        # If we can't resolve it, don't keep the unresolvable string — it would crash
        merged["color"] = resolved  # None → gets filtered out below

    # Build font dict (omit None values)
    d["font"] = {k: v for k, v in merged.items() if v is not None}

    # Tag the color with its source so the generator can decide whether to
    # apply it as an override or let the master theme handle it naturally.
    # color_from_xml=True  → explicitly set in the placeholder XML → apply as override
    # color_from_xml=False → only from master-style/theme cascade → let theme inherit
    if "color" in d["font"]:
        d["font"]["color_from_xml"] = xml_color_found

    # ── 4. Paragraph info from XML ────────────────────────────────────────
    if ph.has_text_frame:
        d["word_wrap"] = ph.text_frame.word_wrap
        paras_collected = []
        for para in ph.text_frame.paragraphs:
            paras_collected.append(_para_info(para))
        merged_para: dict[str, Any] = {}
        for key in ("line_spacing", "space_before", "space_after"):
            for p in paras_collected:
                if p.get(key) is not None:
                    merged_para[key] = p[key]
                    break
        for p in paras_collected:
            if p.get("bullet") and p["bullet"].get("type") != "none":
                merged_para["bullet"] = p["bullet"]
                break
        if merged_para:
            d["paragraph"] = merged_para

    return d


def _classify_placeholder(ph) -> str:
    """Classify a placeholder for layout purpose."""
    pt = ph.placeholder_format.type
    idx = ph.placeholder_format.idx
    if pt in _TITLE_PH_TYPES or idx == 0:
        return "title"
    if pt in _FIXED_PH_TYPES:
        return "fixed"
    if pt in _PICTURE_PH_TYPES:
        return "picture"
    return "content"


def _detect_non_placeholder_shapes(layout) -> list[dict]:
    """Find decorative shapes (lines, rectangles, images) on a layout."""
    shapes_info = []
    for shape in layout.shapes:
        if shape.is_placeholder:
            continue
        info: dict[str, Any] = {
            "shape_type": shape.shape_type.name if shape.shape_type else "UNKNOWN",
            "name": shape.name,
            "pos": {
                "left":   _emu_to_in(shape.left),
                "top":    _emu_to_in(shape.top),
                "width":  _emu_to_in(shape.width),
                "height": _emu_to_in(shape.height),
            },
        }
        color = _shape_color(shape)
        if color:
            info["fill_color"] = color
        shapes_info.append(info)
    return shapes_info


def _detect_slide_master_shapes(prs) -> list[dict]:
    """Find decorative shapes on the slide master itself."""
    shapes_info = []
    try:
        master = prs.slide_masters[0]
        for shape in master.shapes:
            if shape.is_placeholder:
                continue
            info: dict[str, Any] = {
                "shape_type": shape.shape_type.name if shape.shape_type else "UNKNOWN",
                "name": shape.name,
                "pos": {
                    "left":   _emu_to_in(shape.left),
                    "top":    _emu_to_in(shape.top),
                    "width":  _emu_to_in(shape.width),
                    "height": _emu_to_in(shape.height),
                },
            }
            color = _shape_color(shape)
            if color:
                info["fill_color"] = color
            shapes_info.append(info)
    except Exception:
        pass
    return shapes_info


def _extract_theme_colors(prs) -> dict[str, str]:
    """Extract theme/scheme colors from the slide master's theme part."""
    colors: dict[str, str] = {}
    theme_el = _get_theme_element(prs)
    if theme_el is None:
        return colors

    cs = theme_el.find(".//" + qn("a:clrScheme"))
    if cs is None:
        return colors

    for child in cs:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        srgb = child.find(qn("a:srgbClr"))
        if srgb is not None:
            val = srgb.get("val", "")
            if val:
                colors[tag] = f"#{val}"
            continue
        sys_clr = child.find(qn("a:sysClr"))
        if sys_clr is not None:
            val = sys_clr.get("lastClr", sys_clr.get("val", ""))
            if val and len(val) == 6:
                colors[tag] = f"#{val}"

    return colors


def _get_theme_element(prs):
    """Parse the theme XML from the slide master's related theme part."""
    try:
        from lxml import etree
        master = prs.slide_masters[0]
        for rel in master.part.rels.values():
            if "theme" in rel.reltype:
                return etree.fromstring(rel.target_part.blob)
    except Exception:
        pass
    return None


def _extract_theme_fonts(prs) -> dict[str, str | None]:
    """Extract major/minor theme font names."""
    result = {"major_latin": None, "major_ea": None, "minor_latin": None, "minor_ea": None}
    theme_el = _get_theme_element(prs)
    if theme_el is None:
        return result
    try:
        mj = theme_el.find(".//" + qn("a:majorFont"))
        if mj is not None:
            lat = mj.find(qn("a:latin"))
            ea = mj.find(qn("a:ea"))
            if lat is not None:
                result["major_latin"] = lat.get("typeface") or None
            if ea is not None:
                result["major_ea"] = ea.get("typeface") or None
        mn = theme_el.find(".//" + qn("a:minorFont"))
        if mn is not None:
            lat = mn.find(qn("a:latin"))
            ea = mn.find(qn("a:ea"))
            if lat is not None:
                result["minor_latin"] = lat.get("typeface") or None
            if ea is not None:
                result["minor_ea"] = ea.get("typeface") or None
    except Exception:
        pass
    return result


def _extract_master_text_styles(prs) -> dict[str, dict]:
    """Extract default text styles (title/body/other) from the slide master."""
    styles: dict[str, dict] = {}
    try:
        master = prs.slide_masters[0]
        txStyles = master._element.find(qn("p:txStyles"))
        if txStyles is None:
            return styles
        for style_name in ("titleStyle", "bodyStyle", "otherStyle"):
            style_el = txStyles.find(qn("p:" + style_name))
            if style_el is None:
                continue
            # Get lvl1pPr (the default level)
            lvl1 = style_el.find(qn("a:lvl1pPr"))
            if lvl1 is None:
                continue
            info = _parse_rpr_xml(lvl1)
            # Also check defRPr inside lvl1pPr
            defRPr = lvl1.find(qn("a:defRPr"))
            if defRPr is not None:
                sub = _parse_rpr_xml(defRPr)
                # Merge: defRPr values override
                for k, v in sub.items():
                    if v is not None:
                        info[k] = v
            styles[style_name] = info
    except Exception:
        pass
    return styles


def _parse_rpr_xml(el) -> dict[str, Any]:
    """Parse a run-property-like XML element for font info."""
    d: dict[str, Any] = {"name": None, "ea_name": None, "size": None, "bold": None, "color": None}
    latin = el.find(qn("a:latin"))
    if latin is not None:
        tf = latin.get("typeface")
        if tf and not tf.startswith("+"):
            d["name"] = tf
    ea = el.find(qn("a:ea"))
    if ea is not None:
        tf = ea.get("typeface")
        if tf and not tf.startswith("+"):
            d["ea_name"] = tf
    sz = el.get("sz")
    if sz:
        d["size"] = round(int(sz) / 100, 1)
    b = el.get("b")
    if b is not None:
        d["bold"] = b == "1" or b.lower() == "true"
    # Color
    solidFill = el.find(qn("a:solidFill"))
    if solidFill is not None:
        srgb = solidFill.find(qn("a:srgbClr"))
        schm = solidFill.find(qn("a:schemeClr"))
        if srgb is not None:
            d["color"] = f"#{srgb.get('val', '')}"
        elif schm is not None:
            d["color"] = f"scheme:{schm.get('val', '?')}"
    return d


def _extract_placeholder_xml_font(ph) -> dict[str, Any]:
    """Extract font info directly from placeholder XML (defRPr in txBody)."""
    result: dict[str, Any] = {"name": None, "ea_name": None, "size": None, "bold": None, "color": None, "align": None}
    try:
        sp = ph._element
        txBody = sp.find(qn("p:txBody"))
        if txBody is None:
            return result
        for p_el in txBody.findall(qn("a:p")):
            pPr = p_el.find(qn("a:pPr"))
            if pPr is not None:
                algn = pPr.get("algn")
                if algn and result["align"] is None:
                    align_map = {"l": "LEFT", "ctr": "CENTER", "r": "RIGHT", "just": "JUSTIFY"}
                    result["align"] = align_map.get(algn, algn)
                defRPr = pPr.find(qn("a:defRPr"))
                if defRPr is not None:
                    parsed = _parse_rpr_xml(defRPr)
                    for k, v in parsed.items():
                        if v is not None and result.get(k) is None:
                            result[k] = v
            # Also check endParaRPr
            endRPr = p_el.find(qn("a:endParaRPr"))
            if endRPr is not None:
                parsed = _parse_rpr_xml(endRPr)
                for k, v in parsed.items():
                    if v is not None and result.get(k) is None:
                        result[k] = v
            # Check runs
            for r_el in p_el.findall(qn("a:r")):
                rPr = r_el.find(qn("a:rPr"))
                if rPr is not None:
                    parsed = _parse_rpr_xml(rPr)
                    for k, v in parsed.items():
                        if v is not None and result.get(k) is None:
                            result[k] = v
    except Exception:
        pass
    return result


def _build_color_palette(theme_colors: dict, all_fonts: list[dict]) -> dict[str, str]:
    """Build a simplified color_palette from theme colors and font colors found."""
    palette: dict[str, str] = {}

    # Map theme color slots to our palette keys
    if "dk1" in theme_colors:
        palette["text_main"] = theme_colors["dk1"]
    if "dk2" in theme_colors:
        palette["text_sub"] = theme_colors["dk2"]
    if "lt1" in theme_colors:
        palette["background"] = theme_colors["lt1"]
    if "lt2" in theme_colors:
        palette["background_alt"] = theme_colors["lt2"]

    # Accent colors
    for i in range(1, 7):
        key = f"accent{i}"
        if key in theme_colors:
            if "primary" not in palette:
                palette["primary"] = theme_colors[key]
            else:
                palette[key] = theme_colors[key]

    # Hyperlink colors
    if "hlink" in theme_colors:
        palette["hyperlink"] = theme_colors["hlink"]
    if "folHlink" in theme_colors:
        palette["followed_hyperlink"] = theme_colors["folHlink"]

    # If we still don't have primary, pick the most common non-black font color
    if "primary" not in palette:
        color_counter: Counter[str] = Counter()
        for f in all_fonts:
            c = f.get("color")
            if c and c not in ("#000000", "#FFFFFF", None):
                color_counter[c] += 1
        if color_counter:
            palette["primary"] = color_counter.most_common(1)[0][0]

    return palette


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC: deep_scan
# ══════════════════════════════════════════════════════════════════════════════

def deep_scan(pptx_path: str) -> dict:
    """
    Deep-scan a PPTX master file and return two objects matching the schemas
    of ``master_profile.json`` and ``master_structure.json``.

    Returns::

        {
          "master_profile": {
            "template_identity": str,
            "canvas_size": { "width": float, "height": float },
            "color_palette": { ... },
            "theme_colors": { ... },
            "layouts": [ ... ],          # "main" layouts (Title, Content, Two Content)
            "fixed_elements": { ... }
          },
          "master_structure": {
            "additional_layouts": [ ... ],
            "common_elements": { ... }
          }
        }
    """
    prs = Presentation(pptx_path)
    stem = os.path.splitext(os.path.basename(pptx_path))[0]

    # ── Canvas ────────────────────────────────────────────────────────────────
    canvas = {
        "width":  round(prs.slide_width / 914400, 3),
        "height": round(prs.slide_height / 914400, 3),
    }

    # ── Theme colors ──────────────────────────────────────────────────────────
    theme_colors = _extract_theme_colors(prs)

    # ── Theme fonts ───────────────────────────────────────────────────────────
    theme_fonts = _extract_theme_fonts(prs)

    # ── Master text styles ────────────────────────────────────────────────────
    master_styles = _extract_master_text_styles(prs)

    # ── Scan all layouts ──────────────────────────────────────────────────────
    all_fonts: list[dict] = []
    all_layouts: list[dict] = []
    fixed_elements: dict[str, Any] = {}

    # Track which layout types we've seen (for main vs additional split)
    _MAIN_LAYOUT_TYPES = {"Title Slide", "Title and Content", "Two Content",
                          "TITLE", "OBJECT", "TWO_OBJECTS"}

    for li, layout in enumerate(prs.slide_layouts):
        layout_info: dict[str, Any] = {
            "layout_index": li,
            "name": layout.name,
            "pptx_layout": None,
            "placeholders": [],
            "decorative_shapes": [],
        }

        # Try to get the layout type from XML
        try:
            layout_type = layout.element.get("type")
            if layout_type:
                layout_info["pptx_layout"] = layout_type
        except Exception:
            pass

        content_placeholders = []
        for ph in layout.placeholders:
            detail = _extract_placeholder_detail(
                ph,
                master_styles=master_styles,
                theme_fonts=theme_fonts,
                theme_colors=theme_colors,
            )
            cls = _classify_placeholder(ph)

            if detail.get("font"):
                all_fonts.append(detail["font"])

            if cls == "fixed":
                # Capture fixed elements once
                ph_type = ph.placeholder_format.type
                if ph_type == PP_PLACEHOLDER.FOOTER and "footer" not in fixed_elements:
                    fixed_elements["footer"] = {
                        "idx": detail["idx"],
                        "font_size": detail["font"].get("size", 10),
                        "color": detail["font"].get("color", "#A6A6A6"),
                    }
                elif ph_type == PP_PLACEHOLDER.SLIDE_NUMBER and "page_number" not in fixed_elements:
                    fixed_elements["page_number"] = {
                        "idx": detail["idx"],
                        "pos": "bottom_right",
                        "font_size": detail["font"].get("size", 10),
                    }
                elif ph_type == PP_PLACEHOLDER.DATE and "date" not in fixed_elements:
                    fixed_elements["date"] = {
                        "idx": detail["idx"],
                        "font_size": detail["font"].get("size", 10),
                    }
                continue  # don't include fixed ph in layout placeholders

            content_placeholders.append(detail)

        layout_info["placeholders"] = content_placeholders

        # Decorative shapes on this layout
        deco = _detect_non_placeholder_shapes(layout)
        if deco:
            layout_info["decorative_shapes"] = deco

        all_layouts.append(layout_info)

    # ── Master-level decorative shapes ────────────────────────────────────────
    master_shapes = _detect_slide_master_shapes(prs)

    # ── Color palette ─────────────────────────────────────────────────────────
    color_palette = _build_color_palette(theme_colors, all_fonts)

    # ── Split layouts into "main" (first 3 usable) vs "additional" ────────────
    # Heuristic: Title Slide (idx 0), first content layout, two-column layout
    #            are "main". The rest are "additional".
    main_layouts = []
    additional_layouts = []

    for lo in all_layouts:
        name_lower = lo["name"].lower()
        pptx_type = (lo.get("pptx_layout") or "").lower()

        is_main = False
        # Title slide is always main
        if lo["layout_index"] == 0:
            is_main = True
        # Two Content / Two Objects (check BEFORE single-content to avoid overlap)
        elif pptx_type in ("twoobj", "twotxtwoobj") or any(
            k in name_lower for k in ("two content", "two objects", "two_objects", "two ")
        ):
            if not any(
                "two" in m["name"].lower() or m.get("pptx_layout", "").lower().startswith("two")
                for m in main_layouts
            ):
                is_main = True
        # Single-content layout (obj, title and content, etc.)
        elif pptx_type in ("obj", "object") or any(
            k in name_lower for k in ("title and content", "content", "object")
        ):
            if "caption" not in name_lower:
                if not any(
                    m.get("pptx_layout", "").lower() in ("obj", "object")
                    or "content" in m["name"].lower()
                    for m in main_layouts if m["layout_index"] != 0
                ):
                    is_main = True

        if is_main and len(main_layouts) < 3:
            main_layouts.append(lo)
        else:
            extra = dict(lo)
            extra["description"] = _auto_description(lo)
            additional_layouts.append(extra)

    # ── Assemble master_profile ───────────────────────────────────────────────
    master_profile: dict[str, Any] = {
        "template_identity": stem,
        "canvas_size": canvas,
        "color_palette": color_palette,
        "theme_colors": theme_colors,
        "theme_fonts": {k: v for k, v in theme_fonts.items() if v},
        "master_text_styles": master_styles,
        "pptx_layouts_reference": _build_layouts_reference(prs),
        "layouts": main_layouts,
        "fixed_elements": fixed_elements,
    }

    # ── Assemble master_structure ─────────────────────────────────────────────
    common_elements: dict[str, Any] = {}
    if master_shapes:
        common_elements["master_shapes"] = master_shapes
    # Detect if any layout has a prominent colored bar/line
    for lo in all_layouts:
        for s in lo.get("decorative_shapes", []):
            pos = s.get("pos", {})
            w = pos.get("width", 0) or 0
            h = pos.get("height", 0) or 0
            if s.get("fill_color") and (h < 0.15 and w > 0.5):
                common_elements["color_bar"] = {
                    "type": "shape",
                    "pos": pos,
                    "color": s["fill_color"],
                }
                break
        if "color_bar" in common_elements:
            break

    master_structure: dict[str, Any] = {
        "additional_layouts": additional_layouts,
        "common_elements": common_elements,
    }

    return {
        "master_profile": master_profile,
        "master_structure": master_structure,
    }


def _build_layouts_reference(prs) -> str:
    """Build a one-line reference string like 'file: [0]=TITLE, [1]=OBJECT, ...'."""
    parts = []
    for i, layout in enumerate(prs.slide_layouts):
        lt = layout.element.get("type") or layout.name
        parts.append(f"[{i}]={lt}")
    return ", ".join(parts)


def _auto_description(layout_info: dict) -> str:
    """Generate a human-readable description for additional layouts."""
    name = layout_info["name"]
    phs = layout_info["placeholders"]
    n_ph = len(phs)
    ph_names = [p["name"] for p in phs]

    name_lower = name.lower()
    if "blank" in name_lower:
        return "Slide trống, dùng cho hình ảnh hoặc nội dung tùy chỉnh"
    if "section" in name_lower:
        return "Trang đệm ngăn cách giữa các chương lớn"
    if "title only" in name_lower:
        return "Chỉ có tiêu đề, phần còn lại để chèn sơ đồ/hình ảnh"
    if "picture" in name_lower:
        return "Layout có chỗ chèn hình ảnh"
    if "caption" in name_lower:
        return "Nội dung chính kèm phần mô tả bổ sung"
    if "vertical" in name_lower:
        return "Layout văn bản dọc (thường dùng cho tiếng Nhật/Trung)"
    if n_ph == 0:
        return "Layout không có placeholder văn bản"
    return f"Layout với {n_ph} placeholder: {', '.join(ph_names)}"


# ── CLI ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) < 2:
        print("Usage: python deep_scanner.py <path.pptx> [output_prefix]")
        sys.exit(1)

    path = sys.argv[1]
    prefix = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(path)[0]

    result = deep_scan(path)

    profile_path = f"{prefix}_profile.json"
    structure_path = f"{prefix}_structure.json"

    with open(profile_path, "w", encoding="utf-8") as f:
        json.dump(result["master_profile"], f, ensure_ascii=False, indent=2)
    with open(structure_path, "w", encoding="utf-8") as f:
        json.dump(result["master_structure"], f, ensure_ascii=False, indent=2)

    print(f"Profile  → {profile_path}")
    print(f"Structure → {structure_path}")
    n_main = len(result["master_profile"]["layouts"])
    n_extra = len(result["master_structure"]["additional_layouts"])
    print(f"  {n_main} main layouts + {n_extra} additional layouts")
