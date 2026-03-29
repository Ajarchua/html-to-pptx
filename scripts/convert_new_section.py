# 这个文件只是用来生成替换内容，不是独立脚本

NEW_PROCESS_SLIDE = '''
# -- 容器包裹写法处理 --
def add_shape_with_children(slide, shape_el, shape_style):
    left, top, w, h = parse_geometry(shape_style)
    shape_type = shape_el.get('data-type', 'rect')
    type_map = {'rect':1, 'ellipse':9, 'arrow':13}
    mso_type = type_map.get(shape_type, 1)

    shape = slide.shapes.add_shape(mso_type, left, top, w, h)
    fill_color = parse_background_color(shape_style)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    shape.line.fill.background()

    cursor_top = top

    for child in shape_el.find_all(['div','p','span'], recursive=True):
        child_cls   = " ".join(child.get("class", []))
        child_style = parse_style(child.get("style", ""))
        child_dtype = child.get("data-type", "")

        if child_dtype != "text" and "text-box" not in child_cls:
            continue
        text = child.get_text(strip=True)
        if not text:
            continue

        # padding 作为偏移
        pt_px = parse_px(child_style.get("padding-top", child_style.get("padding", "0")))
        pl_px = parse_px(child_style.get("padding-left", child_style.get("padding", "16")))
        font_size_px = parse_px(child_style.get("font-size", "16px"))
        try:
            lh = float(str(child_style.get("line-height", "1.3")).rstrip("px"))
        except Exception:
            lh = 1.3

        line_h_emu = px_to_emu(font_size_px * lh * 1.4, "y")
        txt_left = left + px_to_emu(pl_px, "x")
        txt_top  = cursor_top + px_to_emu(pt_px, "y")
        txt_w    = max(px_to_emu(10, "x"), w - px_to_emu(pl_px * 2, "x"))
        txt_h    = max(line_h_emu, px_to_emu(font_size_px * 2, "y"))

        if txt_top + txt_h > top + h + px_to_emu(20, "y"):
            break

        txBox = slide.shapes.add_textbox(txt_left, txt_top, txt_w, txt_h)
        tf = txBox.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = parse_text_align(child_style)
        run = p.add_run()
        run.text = text
        font = run.font
        font.size   = Pt(px_to_pt(font_size_px))
        font.bold   = parse_font_weight(child_style)
        font.italic = parse_font_italic(child_style)
        color = parse_color(child_style.get("color", "#ffffff"))
        if color: font.color.rgb = color
        ff = child_style.get("font-family", "Calibri").split(",")[0].strip().strip("\\'\\\"")
        font.name = ff
        cursor_top = txt_top + txt_h


def _has_child_textboxes(el):
    for child in el.find_all(["div","p","span"], recursive=True):
        child_cls = " ".join(child.get("class", []))
        if "text-box" in child_cls or child.get("data-type") == "text":
            child_style = parse_style(child.get("style", ""))
            if "absolute" not in child_style.get("position", ""):
                return True
    return False


# -- 主转换逻辑 --
def process_slide(prs, section):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    slide_style = parse_style(section.get("style",""))
    set_slide_background(slide, slide_style)

    skip_els = set()

    for el in section.find_all(True, recursive=True):
        if id(el) in skip_els:
            continue

        el_style = parse_style(el.get("style",""))
        dtype    = el.get("data-type","")
        cls      = " ".join(el.get("class",[]))

        if "absolute" not in el_style.get("position",""):
            if el.name != "table" or "pptx-table" not in cls:
                continue

        if "placeholder" in cls and el.get("data-chart"):
            add_chart_from_placeholder(slide, el, el_style)

        elif dtype == "text" or "text-box" in cls:
            add_text_box(slide, el, el_style)

        elif dtype in ("rect","ellipse","arrow","right-arrow") or "shape" in cls:
            if _has_child_textboxes(el):
                add_shape_with_children(slide, el, el_style)
                for child in el.find_all(True, recursive=True):
                    skip_els.add(id(child))
            else:
                add_shape(slide, el, el_style, dtype or "rect")

        elif el.name == "img":
            add_image(slide, el, el_style)

        elif el.name == "table" and "pptx-table" in cls:
            add_table(slide, el, el_style)

        elif "chart-box" in cls or dtype in ("bar","line","pie","area","column"):
            add_chart_inline(slide, el, el_style)

    return slide
'''
