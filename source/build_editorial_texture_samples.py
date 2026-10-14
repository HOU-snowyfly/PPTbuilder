"""Warm editorial review: three revised pages and editable side-by-side comparison.

Inputs are the unchanged original deck and approved existing story pictures.
Temporary renders remain under work/. No original delivery file is written.
"""
from __future__ import annotations

import hashlib
import io
import json
from collections import Counter
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches

from build_global_style_samples import (
    THEME, box, canonical, emergency_icon, original_texts, picture, rgb, text,
)


ROOT = Path(__file__).resolve().parents[1]
ORIGINAL = ROOT / "交付文件/合理用药_双人35分钟.pptx"
PAPER = ROOT / "source/assets/editorial_surface/warm_paper.png"
OUT = ROOT / "work/editorial-texture-samples"
PAGES = (5, 11, 17)
THEME.update({"ink": "173B50", "accent": "D95361", "paper": "F7F2E8",
              "tint": "EFE5D7", "muted": "DDD7CA", "rule": "CFC6B8"})


def gradient(slide, x, y, w, h, start, end, angle=90):
    obj = box(slide, x, y, w, h, start)
    props = obj._element.spPr
    solid = props.find(qn("a:solidFill"))
    index = list(props).index(solid)
    props.remove(solid)
    grad = OxmlElement("a:gradFill")
    grad.set("rotWithShape", "1")
    stops = OxmlElement("a:gsLst")
    for position, color in ((0, start), (100000, end)):
        stop = OxmlElement("a:gs")
        stop.set("pos", str(position))
        value = OxmlElement("a:srgbClr")
        value.set("val", color)
        stop.append(value)
        stops.append(stop)
    grad.append(stops)
    linear = OxmlElement("a:lin")
    linear.set("ang", str(angle * 60000))
    linear.set("scaled", "1")
    grad.append(linear)
    props.insert(index, grad)
    return obj


def base(presentation, original):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(THEME["paper"])
    backdrop = slide.shapes.add_picture(str(PAPER), 0, 0,
                                       width=presentation.slide_width,
                                       height=presentation.slide_height)
    backdrop.name = "STYLE_WARM_PAPER"
    slide.notes_slide.notes_text_frame.text = original.notes_slide.notes_text_frame.text
    return slide, original_texts(original)


def nav(slide, label, color=None):
    box(slide, .64, .27, .07, .18, THEME["accent"])
    text(slide, label, .89, .23, 5.0, .31, 13, color=color, bold=True)


def footer(slide, label):
    text(slide, label, 11.68, 7.18, 1.0, .2, 11, align=PP_ALIGN.RIGHT)


def scene(presentation, original):
    slide, values = base(presentation, original)
    nav(slide, values[0])
    title = values[1].replace("，", "，\n", 1)
    text(slide, title, .64, .61, 12.02, 1.38, 40, bold=True)
    panels = [s for s in original.shapes if s.shape_type == 13]
    # One dominant image reaches the left edge; all original panel content is retained.
    picture(slide, panels[0], 0, 2.28, 6.64, 4.28)
    gradient(slide, .34, 6.63, 6.02, .46, "EFE5D7", "F7F2E8", angle=0)
    text(slide, values[4], .51, 6.66, 5.72, .37, 22, bold=True)
    for panel, caption, y in ((panels[1], values[5], 2.21),
                              (panels[2], values[6], 4.24)):
        picture(slide, panel, 7.13, y, 2.74, 1.83)
        gradient(slide, 10.12, y + .28, 2.54, 1.2, "EFE5D7", "F7F2E8", angle=0)
        text(slide, caption.replace("，", "，\n", 1), 10.29, y + .43,
             2.2, .88, 20, bold=True)
    gradient(slide, 7.13, 6.32, 5.53, .8, "143446", "23485C", angle=0)
    text(slide, values[2].replace("，", "，\n", 1), 7.37, 6.39,
         5.05, .65, 21, color=THEME["paper"], bold=True)
    footer(slide, values[3])
    return slide


def data(presentation, original):
    slide, values = base(presentation, original)
    gradient(slide, 0, 0, 5.24, 7.5, "122C3C", "254B60")
    nav(slide, values[0], color=THEME["paper"])
    title = values[1].replace("，", "，\n", 1).replace("真实伤害", "\n真实伤害")
    text(slide, title, .65, 1.04, 4.08, 2.58, 32,
         color=THEME["paper"], bold=True)
    box(slide, .65, 4.94, .5, .045, THEME["accent"])
    text(slide, values[2].replace("，", "，\n", 1), .65, 5.48, 4.03, 1.22, 23,
         color=THEME["paper"], bold=True)
    metrics = values[4:]
    text(slide, metrics[1], 5.86, .53, 3.1, .35, 21,
         color=THEME["accent"], bold=True)
    text(slide, metrics[2], 5.8, 1.03, 3.37, 1.34, 78,
         color=THEME["accent"], bold=True)
    text(slide, metrics[0].replace("门诊", "门诊\n", 1), 5.86, 2.49,
         3.27, .78, 23, bold=True)
    for index in range(100):
        row, col = divmod(index, 10)
        obj = box(slide, 9.64 + col * .273, .68 + row * .273, .222, .222,
                  THEME["accent"] if index < 28 else THEME["muted"], rounded=True)
        obj.name = f"DATA_GRID_{index+1:03d}_{'HIGHLIGHT' if index < 28 else 'REST'}"
    text(slide, metrics[3], 5.86, 3.39, 6.8, .4, 23, bold=True)
    box(slide, 5.86, 4.05, 6.8, .016, THEME["rule"])
    text(slide, metrics[4], 5.86, 4.29, 6.8, .39, 23, bold=True)
    text(slide, metrics[5], 5.86, 4.94, 4.56, .61, 34,
         color=THEME["accent"], bold=True)
    text(slide, metrics[6], 10.38, 4.96, 2.28, .6, 18, bold=True)
    for i in range(5):
        emergency_icon(slide, 5.89 + i * 1.365, 5.78, 1.04,
                       THEME["accent"] if i == 0 else THEME["ink"],
                       f"DATA_ED_{i+1}_{'HIGHLIGHT' if i == 0 else 'REST'}")
    text(slide, metrics[7], 5.86, 6.97, 6.8, .31, 18, bold=True)
    footer(slide, values[3])
    return slide


def framework(presentation, original):
    slide, values = base(presentation, original)
    gradient(slide, 4.07, 0, presentation.slide_width / 914400 - 4.07,
             7.5, "18374A", "295165")
    nav(slide, values[0])
    title = values[1].replace("，", "，\n", 1).replace("用药动作", "\n用药动作")
    text(slide, title, .65, 1.06, 2.94, 3.08, 30, bold=True)
    box(slide, .65, 4.72, .5, .045, THEME["accent"])
    text(slide, values[2].replace("，", "，\n", 1), .65, 5.21, 2.94, 1.18,
         22, bold=True)
    actions = values[4:]
    for index, (x, y) in enumerate(((4.67, 1.04), (8.96, 1.04),
                                  (4.67, 4.08), (8.96, 4.08))):
        text(slide, actions[index * 2], x, y, 1.58, 1.29, 70,
             color="F2B1AA", bold=True)
        text(slide, actions[index * 2 + 1], x, y + 1.58, 3.58, 1.0,
             24, color=THEME["paper"], bold=True)
    box(slide, 4.67, 3.68, 7.73, .012, "507082")
    box(slide, 8.39, 1.12, .012, 5.78, "507082")
    text(slide, values[3], 11.68, 7.18, 1.0, .2, 11,
         color=THEME["paper"], align=PP_ALIGN.RIGHT)
    return slide


def clone_shapes(source, target, dx=0, dy=0):
    """Copy original editable shapes, remapping image relationships to the target part."""
    for obj in source.shapes:
        element = deepcopy(obj._element)
        element.find(".//" + qn("p:cNvPr")).set("id", str(target.shapes._next_shape_id))
        if obj.shape_type == 13:
            _, relationship = target.part.get_or_add_image_part(io.BytesIO(obj.image.blob))
            element.find(".//" + qn("a:blip")).set(qn("r:embed"), relationship)
        target.shapes._spTree.insert_element_before(element, "p:extLst")
        fresh = target.shapes[-1]
        fresh.left += Inches(dx)
        fresh.top += Inches(dy)


def compare(original, revised, baseline_pages=PAGES, baseline_label="原版",
            revision_label="暖白纸感与深色编辑版"):
    result = Presentation(ORIGINAL)  # Retain the original theme so baseline looks unchanged.
    old_ids = list(result.slides._sldIdLst)
    result.slide_width = original.slide_width * 2
    result.slide_height = original.slide_height + Inches(.6)
    width = original.slide_width / 914400
    for index, page in enumerate(PAGES):
        slide = result.slides.add_slide(result.slide_layouts[6])
        box(slide, 0, 0, width * 2, .6, "F7F2E8")
        box(slide, 0, .6, width, 7.5, "FFFFFF")
        source = original.slides[baseline_pages[index] - 1]
        text(slide, f"{baseline_label} · 第{page}页", .65, .12, 8.0, .36, 20, bold=True)
        text(slide, f"新版 · {revision_label}", width + .65, .12,
             10.0, .36, 20, bold=True)
        clone_shapes(source, slide, dy=.6)
        clone_shapes(revised.slides[index], slide, dx=width, dy=.6)
        box(slide, width, .6, .016, 7.5, "B6B8B4")
        slide.notes_slide.notes_text_frame.text = source.notes_slide.notes_text_frame.text
    for entry in old_ids:
        result.part.drop_rel(entry.rId)
        result.slides._sldIdLst.remove(entry)
    return result


def verify(original, revised):
    for page, slide in zip(PAGES, revised.slides):
        expected = Counter(canonical(t) for t in original_texts(original.slides[page - 1]))
        actual = Counter(canonical(t) for t in original_texts(slide))
        assert actual == expected, (page, expected - actual, actual - expected)
        assert slide.notes_slide.notes_text_frame.text == original.slides[page - 1].notes_slide.notes_text_frame.text
        expected_images = Counter((s.name, hashlib.sha256(s.image.blob).hexdigest())
                                  for s in original.slides[page - 1].shapes if s.shape_type == 13)
        actual_images = Counter((s.name, hashlib.sha256(s.image.blob).hexdigest())
                                for s in slide.shapes if s.shape_type == 13 and s.name != "STYLE_WARM_PAPER")
        assert actual_images == expected_images
        for obj in slide.shapes:
            assert obj.left >= 0 and obj.top >= 0, (page, obj.name)
            # Native background is precisely aligned to the slide bounds.
            assert obj.left + obj.width <= revised.slide_width + 9144, (page, obj.name)
            assert obj.top + obj.height <= revised.slide_height, (page, obj.name)
    chart = revised.slides[1]
    for prefix, total, highlight in (("DATA_GRID_", 100, 28), ("DATA_ED_", 5, 1)):
        objects = [s for s in chart.shapes if s.name.startswith(prefix)]
        assert len(objects) == total
        assert sum(s.name.endswith("HIGHLIGHT") for s in objects) == highlight


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    before = hashlib.sha256(ORIGINAL.read_bytes()).hexdigest()
    original = Presentation(ORIGINAL)
    revised = Presentation()
    revised.slide_width, revised.slide_height = original.slide_width, original.slide_height
    for page, layout in zip(PAGES, (scene, data, framework)):
        layout(revised, original.slides[page - 1])
    verify(original, revised)
    sample = OUT / "合理用药_质感改版_三页样页.pptx"
    comparison = OUT / "合理用药_原版与质感改版_并排对比.pptx"
    revised.save(sample)
    combined = compare(original, revised)
    assert len(combined.slides) == len(PAGES)
    combined.save(comparison)
    assert hashlib.sha256(ORIGINAL.read_bytes()).hexdigest() == before
    report = {"pages": PAGES, "original_unchanged": True,
              "content_notes_and_story_pixels_preserved": True,
              "data_grid": [100, 28], "data_ed": [5, 1],
              "surface": "warm paper and editable dark gradients",
              "comparison": "editable side-by-side, original left / revision right",
              "sample": str(sample.relative_to(ROOT)),
              "comparison_file": str(comparison.relative_to(ROOT))}
    (OUT / "checks.json").write_text(json.dumps(report, ensure_ascii=False, indent=2))
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
