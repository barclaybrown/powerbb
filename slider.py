#!/usr/bin/env python3
"""
pptx_inspect.py

Open a PowerPoint file, list slides with layout indicator, layout name, and title.
Prompt the user to choose a slide, then display a structured report of all components
on that slide (type, placeholder info, position/size, text, images, tables, charts, etc.).
Loops until the user quits.

Install dependency:
    pip install python-pptx
"""

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Emu
import sys
import textwrap

# --- EDIT THIS PATH OR SET VIA COMMAND LINE ARGUMENT ---
PPTX_PATH = r"Q:\_pro\Johns Hopkins Univ JHU\AI and SE Course\powerbb generation\deck.pptx"

# ---- Helpers ----

EMU_PER_INCH = 914400
EMU_PER_CM = 360000

def emu_to_in(val) -> float:
    try:
        return float(val) / EMU_PER_INCH
    except Exception:
        return None

def emu_to_cm(val) -> float:
    try:
        return float(val) / EMU_PER_CM
    except Exception:
        return None

def fmt_len(val) -> str:
    """Format an EMU length as 'X.XX in / Y.Y cm'."""
    if val is None:
        return "N/A"
    x_in = emu_to_in(val)
    x_cm = emu_to_cm(val)
    if x_in is None or x_cm is None:
        return "N/A"
    return f"{x_in:.2f} in / {x_cm:.2f} cm"

def safe_enum_name(e) -> str:
    """Return Enum member name cleanly, even if None or library changes."""
    try:
        return e.name  # python-pptx enums are IntEnums with .name
    except Exception:
        return str(e)

def get_layout_indicator(prs: Presentation, slide) -> str:
    """
    Return 'master_idx:layout_idx' where slide.layout matches a layout
    within prs.slide_masters[master_idx].slide_layouts[layout_idx].
    """
    target = slide.slide_layout
    for mi, master in enumerate(prs.slide_masters):
        for li, layout in enumerate(master.slide_layouts):
            if layout is target:
                return f"{mi}:{li}"
    return "?:?"

def get_slide_title(slide) -> str:
    """
    Best-effort retrieval of a slide title:
    1) slide.shapes.title
    2) a TITLE or CENTER_TITLE placeholder
    3) first text-containing shape
    """
    try:
        t = slide.shapes.title
        if t is not None and hasattr(t, "text"):
            txt = (t.text or "").strip()
            if txt:
                return txt
    except Exception:
        pass

    # Try title placeholders explicitly
    for shp in slide.shapes:
        if getattr(shp, "is_placeholder", False):
            try:
                pht = shp.placeholder_format.type
                if pht in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.SUBTITLE):
                    txt = (getattr(shp, "text", "") or "").strip()
                    if txt:
                        return txt
            except Exception:
                pass

    # Fallback: first text-bearing shape
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False):
            txt = (getattr(shp, "text", "") or "").strip()
            if txt:
                return txt

    return "(no title)"

def wrap_if_needed(text: str, width: int = 100) -> str:
    text = text.strip()
    if not text:
        return ""
    return "\n".join(textwrap.wrap(text, width=width, replace_whitespace=False))

def describe_text_frame(shp, indent: str) -> list[str]:
    lines = []
    try:
        tf = shp.text_frame
    except Exception:
        return lines

    if tf is None:
        return lines

    # Text (flattened)
    try:
        txt = (shp.text or "").strip()
    except Exception:
        txt = ""
    if txt:
        lines.append(f"{indent}- Text (flattened):")
        for ln in wrap_if_needed(txt, 100).splitlines():
            lines.append(f"{indent}  {ln}")

    # Margins & autofit hints (best-effort: some decks have None)
    def get_attr(obj, name, default=None):
        try:
            return getattr(obj, name)
        except Exception:
            return default

    margins = {
        "left": get_attr(tf, "margin_left"),
        "right": get_attr(tf, "margin_right"),
        "top": get_attr(tf, "margin_top"),
        "bottom": get_attr(tf, "margin_bottom"),
    }
    if any(v is not None for v in margins.values()):
        lines.append(f"{indent}- TextFrame margins:")
        for k, v in margins.items():
            if isinstance(v, Emu):
                lines.append(f"{indent}  {k}: {fmt_len(v)}")
            elif v is not None:
                lines.append(f"{indent}  {k}: {v}")
    # Autofit / word wrap
    autosize = get_attr(tf, "auto_size", None)
    word_wrap = get_attr(tf, "word_wrap", None)
    if autosize is not None or word_wrap is not None:
        lines.append(f"{indent}- TextFrame properties:")
        if autosize is not None:
            lines.append(f"{indent}  auto_size: {autosize}")
        if word_wrap is not None:
            lines.append(f"{indent}  word_wrap: {word_wrap}")

    # Paragraph overview
    try:
        if tf.paragraphs:
            lines.append(f"{indent}- Paragraphs ({len(tf.paragraphs)}):")
            for i, p in enumerate(tf.paragraphs, 1):
                lvl = getattr(p, "level", None)
                al = getattr(p, "alignment", None)
                al_name = safe_enum_name(al) if al is not None else None
                ptxt = (p.text or "").strip()
                lines.append(f"{indent}  [{i}] level={lvl} alignment={al_name} text={ptxt!r}")
    except Exception:
        pass

    return lines

def describe_table(shp, indent: str) -> list[str]:
    lines = []
    try:
        tbl = shp.table
        rows = len(tbl.rows)
        cols = len(tbl.columns)
        lines.append(f"{indent}- Table: {rows} rows x {cols} columns")
        # Show first few cells’ text
        preview_rows = min(rows, 5)
        preview_cols = min(cols, 5)
        for r in range(preview_rows):
            row_txts = []
            for c in range(preview_cols):
                try:
                    cell_txt = tbl.cell(r, c).text
                except Exception:
                    cell_txt = ""
                row_txts.append(cell_txt.replace("\n", " ").strip())
            lines.append(f"{indent}  Row {r}: {row_txts}")
    except Exception as e:
        lines.append(f"{indent}- Table: <access error: {e}>")
    return lines

def describe_chart(shp, indent: str) -> list[str]:
    lines = []
    try:
        ch = shp.chart  # only exists on chart shapes
        ctype = getattr(ch, "chart_type", None)
        lines.append(f"{indent}- Chart type: {ctype}")
        # Series names & point counts
        try:
            series = list(ch.series)
            lines.append(f"{indent}- Series ({len(series)}):")
            for i, s in enumerate(series, 1):
                name = getattr(s, "name", "")
                # category/values counts (best-effort)
                cat_ct = len(list(s.categories)) if getattr(s, "categories", None) else None
                val_ct = len(list(s.values)) if getattr(s, "values", None) else None
                lines.append(f"{indent}  [{i}] name={name!r} categories={cat_ct} values={val_ct}")
        except Exception:
            pass
    except Exception as e:
        lines.append(f"{indent}- Chart: <access error: {e}>")
    return lines

def describe_picture(shp, indent: str) -> list[str]:
    lines = []
    try:
        img = shp.image
        if img is not None:
            fn = getattr(img, "filename", None)
            ct = getattr(img, "content_type", None)
            ext = getattr(img, "ext", None)
            size_px = getattr(img, "size", None)  # sometimes returns (px_w, px_h)
            lines.append(f"{indent}- Image:")
            if fn: lines.append(f"{indent}  filename: {fn}")
            if ct: lines.append(f"{indent}  content_type: {ct}")
            if ext: lines.append(f"{indent}  extension: {ext}")
            if size_px:
                try:
                    lines.append(f"{indent}  intrinsic_px: {size_px[0]}x{size_px[1]}")
                except Exception:
                    pass
    except Exception as e:
        lines.append(f"{indent}- Image: <access error: {e}>")
    return lines

def describe_placeholder(shp, indent: str) -> list[str]:
    lines = []
    try:
        pf = shp.placeholder_format
        ptype = safe_enum_name(getattr(pf, "type", None))
        pidx = getattr(pf, "idx", None)
        lines.append(f"{indent}- Placeholder: type={ptype} idx={pidx}")
    except Exception:
        pass
    return lines

def describe_common_geometry(shp, indent: str) -> list[str]:
    lines = []
    # Position/size
    try:
        lines.append(f"{indent}- Position: left={fmt_len(getattr(shp, 'left', None))}, "
                     f"top={fmt_len(getattr(shp, 'top', None))}")
        lines.append(f"{indent}- Size: width={fmt_len(getattr(shp, 'width', None))}, "
                     f"height={fmt_len(getattr(shp, 'height', None))}")
    except Exception:
        pass

    # Rotation & z-order
    try:
        rot = getattr(shp, "rotation", None)
        if rot not in (None, 0):
            lines.append(f"{indent}- Rotation: {rot:.2f}°")
    except Exception:
        pass
    try:
        z = getattr(shp, "z_order_position", None)
        if z is not None:
            lines.append(f"{indent}- Z-order: {z}")
    except Exception:
        pass

    # Alt text and name/id
    nm = getattr(shp, "name", None)
    sid = getattr(shp, "shape_id", None)
    if nm is not None or sid is not None:
        lines.append(f"{indent}- Identity: name={nm!r} shape_id={sid}")
    alt = getattr(shp, "alternative_text", None) or getattr(shp, "alt_text", None)
    if alt:
        lines.append(f"{indent}- Alt text: {alt!r}")

    return lines

def describe_shape(shp, indent_level: int = 0) -> list[str]:
    indent = "  " * indent_level
    lines = []

    # Header line: index handled by caller; we show type and (placeholder?) flags
    stype = safe_enum_name(getattr(shp, "shape_type", None))
    header = f"{indent}• Shape type: {stype}"
    if getattr(shp, "is_placeholder", False):
        header += " (placeholder)"
    lines.append(header)

    # Common geometry & identity
    lines.extend(describe_common_geometry(shp, indent))

    # Placeholder info
    if getattr(shp, "is_placeholder", False):
        lines.extend(describe_placeholder(shp, indent))

    # Type-specific sections
    try:
        if getattr(shp, "has_text_frame", False):
            lines.extend(describe_text_frame(shp, indent))
    except Exception:
        pass

    try:
        if getattr(shp, "has_table", False):
            lines.extend(describe_table(shp, indent))
    except Exception:
        pass

    try:
        if getattr(shp, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
            lines.extend(describe_picture(shp, indent))
    except Exception:
        pass

    try:
        if getattr(shp, "shape_type", None) == MSO_SHAPE_TYPE.CHART:
            lines.extend(describe_chart(shp, indent))
    except Exception:
        pass

    # Grouped shapes (recurse)
    try:
        if getattr(shp, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            grp = getattr(shp, "shapes", None) or getattr(shp, "group_items", None)
            if grp is not None:
                lines.append(f"{indent}- Group contains {len(grp)} shapes:")
                for gidx, gshp in enumerate(grp, 1):
                    lines.append(f"{indent}  [{gidx}]")
                    lines.extend(describe_shape(gshp, indent_level + 2))
    except Exception:
        pass

    return lines

def list_slides(prs: Presentation):
    print("\nSlides:")
    print("-------")
    for idx, slide in enumerate(prs.slides, start=1):
        layout_indicator = get_layout_indicator(prs, slide)
        layout_name = getattr(slide.slide_layout, "name", "(unknown layout)")
        title = get_slide_title(slide)
        print(f"[{idx:>2}] {layout_indicator}  |  {layout_name}  |  {title}")

def show_slide_components(prs: Presentation, slide_index: int):
    if slide_index < 1 or slide_index > len(prs.slides):
        print("Invalid slide number.")
        return
    slide = prs.slides[slide_index - 1]
    layout_indicator = get_layout_indicator(prs, slide)
    layout_name = getattr(slide.slide_layout, "name", "(unknown layout)")
    title = get_slide_title(slide)
    slide_id = getattr(slide, "slide_id", None)

    print("\n" + "=" * 80)
    print(f"Slide {slide_index}  |  ID: {slide_id}  |  Layout: {layout_indicator} '{layout_name}'")
    print(f"Title: {title}")
    print("-" * 80)

    shapes = slide.shapes
    print(f"\nShapes on slide: {len(shapes)}")
    for i, shp in enumerate(shapes, 1):
        print(f"\n[{i}] -----------------------------")
        for line in describe_shape(shp, indent_level=0):
            print(line)
    print("=" * 80 + "\n")

def main():
    path = PPTX_PATH
    if len(sys.argv) > 1:
        path = sys.argv[1]

    try:
        prs = Presentation(path)
    except Exception as e:
        print(f"Error opening {path!r}: {e}")
        sys.exit(1)

    print(f"\nOpened: {path}\n")
    while True:
        list_slides(prs)
        print("\nEnter a slide number to inspect, or 'r' to reload, or 'q' to quit.")
        choice = input("> ").strip().lower()

        if choice in ("q", "quit", "exit"):
            break
        if choice in ("r", "reload"):
            try:
                prs = Presentation(path)
                print("Reloaded presentation.")
            except Exception as e:
                print(f"Reload failed: {e}")
            continue

        try:
            idx = int(choice)
        except ValueError:
            print("Please enter a valid number, 'r' to reload, or 'q' to quit.")
            continue

        show_slide_components(prs, idx)

        _ = input('Press ENTER to continue...')

if __name__ == "__main__":
    main()
