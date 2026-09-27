"""Sistema de diseño y helpers para el rediseño de 'Actuadores' (python-pptx)."""
import re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.enum.text import MSO_AUTO_SIZE
import copy

# ------------------------------------------------------------------
# Paleta y tipografías
# ------------------------------------------------------------------
NAVY = RGBColor(0x14, 0x23, 0x3B)
PAPER = RGBColor(0xF5, 0xF3, 0xEE)
LIGHTBLUE = RGBColor(0xE9, 0xEE, 0xF3)
ORANGE = RGBColor(0xE0, 0x69, 0x2A)
ORANGE_TXT = RGBColor(0xB8, 0x50, 0x1A)
TEAL = RGBColor(0x1B, 0x8A, 0x8C)
TEAL_TXT = RGBColor(0x14, 0x6E, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_TXT = RGBColor(0x5B, 0x66, 0x78)
CODE_BG = NAVY

FONT_TITLE = "Space Grotesk"
FONT_BODY = "IBM Plex Sans"
FONT_CODE = "JetBrains Mono"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

MARGIN_X = Inches(0.6)
CONTENT_TOP = Inches(1.85)
CONTENT_BOTTOM = Inches(7.0)
LABEL_Y = Inches(0.45)
TITLE_Y = Inches(0.78)
TITLE_H = Inches(0.85)


def new_presentation():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_slide(prs, bg=PAPER):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = bg
    bg_shape.line.fill.background()
    bg_shape.shadow.inherit = False
    # send to back
    sp = bg_shape._element
    sp.getparent().remove(sp)
    slide.shapes._spTree.insert(2, sp)
    return slide


def _set_run(run, text, font=FONT_BODY, size=16, color=NAVY, bold=False, italic=False, spacing=None):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    if spacing is not None:
        rPr = run._r.get_or_add_rPr()
        rPr.set('spc', str(int(spacing * 100)))
    ea = run.font._rPr.get_or_change_to_latin() if False else None
    return run


def add_textbox(slide, x, y, w, h, word_wrap=True, anchor=MSO_ANCHOR.TOP, autosize=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = word_wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if autosize:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    return tb, tf


def set_paragraph(p, align=PP_ALIGN.LEFT, space_after=0, space_before=0, line_spacing=None, level=0):
    p.alignment = align
    p.space_after = Pt(space_after)
    p.space_before = Pt(space_before)
    if line_spacing:
        p.line_spacing = line_spacing
    p.level = level
    return p


def add_label_title(slide, section_label, title, title_size=30, label_color=TEAL_TXT, two_line_ok=True):
    """Rotulo pequeño en mayusculas + titulo a la misma altura fija."""
    _, tf = add_textbox(slide, MARGIN_X, LABEL_Y, Inches(11.5), Inches(0.3))
    p = tf.paragraphs[0]
    set_paragraph(p)
    r = p.add_run()
    _set_run(r, section_label.upper(), font=FONT_BODY, size=12.5, color=label_color, bold=True, spacing=1.5)

    _, tf2 = add_textbox(slide, MARGIN_X, TITLE_Y, Inches(12.1), TITLE_H, anchor=MSO_ANCHOR.TOP)
    p2 = tf2.paragraphs[0]
    set_paragraph(p2, line_spacing=1.0)
    r2 = p2.add_run()
    _set_run(r2, title, font=FONT_TITLE, size=title_size, color=NAVY, bold=True)
    return tf2


def add_page_footer(slide, page_num, total=57):
    _, tf = add_textbox(slide, Inches(12.3), Inches(7.15), Inches(0.8), Inches(0.3))
    p = tf.paragraphs[0]
    set_paragraph(p, align=PP_ALIGN.RIGHT)
    r = p.add_run()
    _set_run(r, f"{page_num:02d}", font=FONT_BODY, size=10, color=GRAY_TXT, bold=False)


def add_rect(slide, x, y, w, h, fill=LIGHTBLUE, line=None, shadow=False, radius=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    if radius:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


def add_card(slide, x, y, w, h, title=None, body=None, accent=TEAL, title_color=None,
             body_color=NAVY, fill=LIGHTBLUE, title_size=15, body_size=13, valign=MSO_ANCHOR.TOP):
    card = add_rect(slide, x, y, w, h, fill=fill, radius=0.06)
    pad = Inches(0.22)
    tx, ty, tw, th = x + pad, y + Inches(0.16), w - 2 * pad, h - Inches(0.3)
    _, tf = add_textbox(slide, tx, ty, tw, th, anchor=valign)
    first = True
    if title:
        p = tf.paragraphs[0]
        set_paragraph(p, space_after=4)
        r = p.add_run()
        _set_run(r, title, font=FONT_TITLE, size=title_size, color=title_color or accent, bold=True)
        first = False
    if body:
        lines = body if isinstance(body, list) else [body]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            set_paragraph(p, space_after=3, line_spacing=1.08)
            r = p.add_run()
            _set_run(r, line, font=FONT_BODY, size=body_size, color=body_color)
    return card


def add_bullets(slide, x, y, w, h, items, size=14.5, color=NAVY, bullet_color=None,
                space_after=8, bold_lead=None, line_spacing=1.15, anchor=MSO_ANCHOR.TOP):
    """items: list of str, or (lead_bold, rest) tuples."""
    _, tf = add_textbox(slide, x, y, w, h, anchor=anchor)
    bullet_color = bullet_color or ORANGE
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        set_paragraph(p, space_after=space_after, line_spacing=line_spacing)
        # bullet glyph as its own colored run
        rb = p.add_run()
        _set_run(rb, "■  ", font=FONT_BODY, size=size * 0.62, color=bullet_color, bold=True)
        if isinstance(item, tuple):
            lead, rest = item
            r1 = p.add_run()
            _set_run(r1, lead, font=FONT_BODY, size=size, color=color, bold=True)
            if rest:
                r2 = p.add_run()
                _set_run(r2, rest, font=FONT_BODY, size=size, color=color, bold=False)
        else:
            r1 = p.add_run()
            _set_run(r1, item, font=FONT_BODY, size=size, color=color, bold=False)
    return tf


def add_picture_contain(slide, path, x, y, w, h, align="center"):
    """Inserta imagen respetando proporción, centrada en el area (x,y,w,h)."""
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    ratio_box = w / h
    ratio_img = iw / ih
    if ratio_img > ratio_box:
        new_w = w
        new_h = int(w / ratio_img)
    else:
        new_h = h
        new_w = int(h * ratio_img)
    off_x = x + (w - new_w) // 2
    off_y = y + (h - new_h) // 2
    pic = slide.shapes.add_picture(path, off_x, off_y, width=new_w, height=new_h)
    return pic


def add_picture_cover(slide, path, x, y, w, h):
    """Inserta imagen recortando para llenar completamente el area (cover)."""
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    ratio_box = w / h
    ratio_img = iw / ih
    pic = slide.shapes.add_picture(path, x, y, width=w, height=h)
    if ratio_img > ratio_box:
        # image wider than box -> crop left/right
        visible_ratio = ratio_box / ratio_img
        crop = (1 - visible_ratio) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    else:
        visible_ratio = ratio_img / ratio_box
        crop = (1 - visible_ratio) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop
    return pic


def rounded_image_frame(slide, path, x, y, w, h, cover=True, border_color=None, shadow=True):
    if cover:
        pic = add_picture_cover(slide, path, x, y, w, h)
    else:
        pic = add_picture_contain(slide, path, x, y, w, h)
    return pic


# ------------------------------------------------------------------
# Diapositivas divisoras
# ------------------------------------------------------------------
def add_divider_slide(prs, number, section_title, subtitle, image_path, credit=None):
    slide = blank_slide(prs, bg=NAVY)
    # Big number top-left
    _, tf = add_textbox(slide, Inches(0.7), Inches(0.55), Inches(4.0), Inches(2.0))
    p = tf.paragraphs[0]
    set_paragraph(p)
    r = p.add_run()
    _set_run(r, number, font=FONT_TITLE, size=110, color=ORANGE, bold=True)

    # Section title
    _, tf2 = add_textbox(slide, Inches(0.75), Inches(2.55), Inches(6.6), Inches(2.6))
    p2 = tf2.paragraphs[0]
    set_paragraph(p2, line_spacing=1.05)
    r2 = p2.add_run()
    _set_run(r2, section_title, font=FONT_TITLE, size=40, color=WHITE, bold=True)

    if subtitle:
        p3 = tf2.add_paragraph()
        set_paragraph(p3, space_before=16, line_spacing=1.2)
        r3 = p3.add_run()
        _set_run(r3, subtitle, font=FONT_BODY, size=16, color=RGBColor(0xC9, 0xD3, 0xE0))

    # Photo on right side, full-bleed
    if image_path:
        img_x, img_y = Inches(7.9), Inches(0)
        img_w, img_h = SLIDE_W - img_x, SLIDE_H
        rounded_image_frame(slide, image_path, img_x, img_y, img_w, img_h, cover=True)
        # subtle overlay strip for separation
        strip = add_rect(slide, img_x - Inches(0.02), 0, Inches(0.04), SLIDE_H, fill=NAVY)

    if credit:
        _, tfc = add_textbox(slide, Inches(0.75), Inches(7.05), Inches(6), Inches(0.3))
        pc = tfc.paragraphs[0]
        rc = pc.add_run()
        _set_run(rc, credit, font=FONT_BODY, size=9, color=RGBColor(0x8A, 0x94, 0xA6), italic=True)
    return slide


# ------------------------------------------------------------------
# Tablas nativas
# ------------------------------------------------------------------
def add_table(slide, x, y, w, h, headers, rows, col_widths=None, header_fill=NAVY,
              header_color=WHITE, body_size=13, header_size=13, row_fill_alt=LIGHTBLUE,
              highlight_rows=None, align_cols=None):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    gframe = slide.shapes.add_table(n_rows, n_cols, x, y, w, h)
    table = gframe.table
    # remove default banding style visuals by setting explicit fills
    if col_widths:
        for i, cw in enumerate(col_widths):
            table.columns[i].width = cw
    # header
    for c, htext in enumerate(headers):
        cell = table.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        cell.margin_left = Inches(0.1)
        cell.margin_right = Inches(0.1)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        set_paragraph(p, align=(align_cols[c] if align_cols else PP_ALIGN.LEFT))
        r = p.add_run()
        _set_run(r, htext, font=FONT_BODY, size=header_size, color=header_color, bold=True)
    # body
    for ri, row in enumerate(rows, start=1):
        is_hl = highlight_rows and (ri - 1) in highlight_rows
        for c, val in enumerate(row):
            cell = table.cell(ri, c)
            cell.fill.solid()
            if is_hl:
                cell.fill.fore_color.rgb = RGBColor(0xF5, 0xDE, 0xCC)
            else:
                cell.fill.fore_color.rgb = row_fill_alt if ri % 2 == 0 else WHITE
            cell.margin_left = Inches(0.1)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            set_paragraph(p, align=(align_cols[c] if align_cols else PP_ALIGN.LEFT))
            r = p.add_run()
            bold = val.startswith("**") and val.endswith("**")
            text = val[2:-2] if bold else val
            _set_run(r, text, font=FONT_BODY, size=body_size, color=NAVY, bold=bold)
    # strip default table style banding (first row/col special formatting)
    tbl = table._tbl
    tblPr = tbl.find(qn('a:tblPr'))
    if tblPr is not None:
        tblPr.set('firstRow', '0')
        tblPr.set('bandRow', '0')
    return gframe


# ------------------------------------------------------------------
# Bloque de codigo con resaltado de sintaxis simple (Arduino/C++)
# ------------------------------------------------------------------
KEYWORDS = set("void int bool if else for while return const digitalWrite analogWrite "
               "pinMode delay delayMicroseconds Serial begin OUTPUT INPUT HIGH LOW map "
               "setup loop include define".split())

TYPES = set("int bool void float".split())


def _tokenize_code(line):
    tokens = []
    pattern = re.compile(r'(//.*$)|("([^"\\]|\\.)*")|(\b\d+\b)|(\b\w+\b)|([^\w\s])|(\s+)')
    for m in pattern.finditer(line):
        if m.group(1):
            tokens.append(("comment", m.group(1)))
        elif m.group(2):
            tokens.append(("string", m.group(2)))
        elif m.group(4):
            tokens.append(("number", m.group(4)))
        elif m.group(5):
            word = m.group(5)
            if word in KEYWORDS:
                tokens.append(("keyword", word))
            else:
                tokens.append(("default", word))
        elif m.group(6):
            tokens.append(("punct", m.group(6)))
        elif m.group(7):
            tokens.append(("default", m.group(7)))
    return tokens


CODE_COLORS = {
    "keyword": RGBColor(0x6E, 0xD6, 0xD1),
    "string": RGBColor(0xF2, 0xA6, 0x71),
    "number": RGBColor(0xF2, 0xA6, 0x71),
    "comment": RGBColor(0x7C, 0x8A, 0xA5),
    "punct": RGBColor(0xE9, 0xEE, 0xF3),
    "default": RGBColor(0xF5, 0xF3, 0xEE),
}


def add_code_block(slide, x, y, w, h, code_lines, font_size=13.5, title=None):
    panel = add_rect(slide, x, y, w, h, fill=CODE_BG, radius=0.035)
    pad = Inches(0.22)
    ty = y + pad
    if title:
        _, tft = add_textbox(slide, x + pad, ty, w - 2 * pad, Inches(0.3))
        p = tft.paragraphs[0]
        r = p.add_run()
        _set_run(r, title, font=FONT_CODE, size=11.5, color=RGBColor(0x8A, 0x94, 0xA6), bold=True)
        ty += Inches(0.34)
    _, tf = add_textbox(slide, x + pad, ty, w - 2 * pad, y + h - ty - pad)
    for i, line in enumerate(code_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        set_paragraph(p, space_after=2, line_spacing=1.25)
        if line.strip() == "":
            r = p.add_run()
            _set_run(r, " ", font=FONT_CODE, size=font_size, color=WHITE)
            continue
        for kind, text in _tokenize_code(line):
            r = p.add_run()
            _set_run(r, text, font=FONT_CODE, size=font_size, color=CODE_COLORS.get(kind, WHITE))
    return panel


# ------------------------------------------------------------------
# Numeros grandes / stat callouts
# ------------------------------------------------------------------
def add_big_stat(slide, x, y, w, h, number, label, color=ORANGE_TXT, number_size=44, label_size=12.5):
    _, tf = add_textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP)
    p = tf.paragraphs[0]
    set_paragraph(p, space_after=2)
    r = p.add_run()
    _set_run(r, number, font=FONT_TITLE, size=number_size, color=color, bold=True)
    p2 = tf.add_paragraph()
    set_paragraph(p2, line_spacing=1.1)
    r2 = p2.add_run()
    _set_run(r2, label, font=FONT_BODY, size=label_size, color=NAVY)
    return tf


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text
