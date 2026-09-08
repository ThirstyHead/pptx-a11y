"""Generate accessible No Knead Bread.pptx and barrier-filled No Knead Bread-test.pptx."""
from pathlib import Path
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx_a11y.audit import audit_file
from pptx_a11y.rules import RULES

EXAMPLES_DIR = Path(__file__).parent
IMAGE_PATH = EXAMPLES_DIR / "bread.jpg"


def set_run_lang(run, lang="en-US"):
    """Ensure run has an a:rPr element with lang attribute."""
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", lang)


def format_run(run, font_size_pt: float = 15.0, bold: bool = False, color_rgb=(0x0F, 0x17, 0x2A), lang: str = "en-US"):
    """Format run typography: language, size, weight, and WCAG-compliant color."""
    set_run_lang(run, lang=lang)
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.name = "Calibri"
    run.font.color.rgb = RGBColor(*color_rgb)


def remove_bullets(paragraph):
    """Ensure paragraph renders as clean body prose without bullet symbols."""
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith(("buClr", "buSzPct", "buSzPts", "buFont", "buChar", "buAutoNum", "buBlip")):
            pPr.remove(child)
    if pPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}buNone") is None:
        etree.SubElement(pPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}buNone")


def set_numbered_list(paragraph, start_at: int | None = None):
    """Configure paragraph as an accessible OpenXML auto-numbered list item."""
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith(("buClr", "buSzPct", "buSzPts", "buFont", "buChar", "buAutoNum", "buBlip", "buNone")):
            pPr.remove(child)
    attrib = {"type": "arabicPeriod"}
    if start_at is not None:
        attrib["startAt"] = str(start_at)
    etree.SubElement(pPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum", attrib=attrib)
    pPr.set("marL", "342900")
    pPr.set("indent", "-342900")


def set_bullet_list(paragraph, char: str = "•"):
    """Configure paragraph as an accessible OpenXML bulleted list item."""
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith(("buClr", "buSzPct", "buSzPts", "buFont", "buChar", "buAutoNum", "buBlip", "buNone")):
            pPr.remove(child)
    etree.SubElement(pPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}buChar", attrib={"char": char})
    pPr.set("marL", "342900")
    pPr.set("indent", "-342900")


def style_card_placeholder(shape, fill_hex="F8FAFC", border_hex="CBD5E1"):
    """Style placeholder shape with a modern card container appearance."""
    shape.fill.solid()
    r = int(fill_hex[0:2], 16)
    g = int(fill_hex[2:4], 16)
    b = int(fill_hex[4:6], 16)
    shape.fill.fore_color.rgb = RGBColor(r, g, b)

    shape.line.color.rgb = RGBColor(
        int(border_hex[0:2], 16), int(border_hex[2:4], 16), int(border_hex[4:6], 16)
    )
    shape.line.width = Pt(1)

    shape.text_frame.margin_left = Inches(0.28)
    shape.text_frame.margin_right = Inches(0.28)
    shape.text_frame.margin_top = Inches(0.28)
    shape.text_frame.margin_bottom = Inches(0.28)
    shape.text_frame.word_wrap = True


def add_accessible_sections(prs: Presentation):
    """Add accessible, uniquely named sections using presentation extension list."""
    extLst = prs._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}extLst")
    if extLst is None:
        extLst = etree.SubElement(
            prs._element,
            "{http://schemas.openxmlformats.org/presentationml/2006/main}extLst",
        )
    ext = etree.SubElement(
        extLst,
        "{http://schemas.openxmlformats.org/presentationml/2006/main}ext",
        attrib={"uri": "{521415D9-36F7-43E2-AB2F-EE9909263876}"},
    )
    sectionLst = etree.SubElement(
        ext,
        "{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst",
        nsmap={"p14": "http://schemas.microsoft.com/office/powerpoint/2010/main"},
    )
    for name in ["Introduction", "Preparation", "Baking & Serving"]:
        etree.SubElement(
            sectionLst,
            "{http://schemas.microsoft.com/office/powerpoint/2010/main}section",
            attrib={"name": name, "id": f"{{{name.lower()}}}"},
        )


def build_accessible_deck(output_path: Path):
    prs = Presentation()
    prs.core_properties.title = "No Knead Bread Recipe"
    add_accessible_sections(prs)

    # -------------------------------------------------------------
    # Slide 1: Title Slide
    # -------------------------------------------------------------
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    title1 = slide1.shapes.title
    title1.text = "No Knead Bread"
    format_run(title1.text_frame.paragraphs[0].runs[0], font_size_pt=40, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    sub1 = slide1.placeholders[1]
    sub1.text = "No kneading required, 4 simple ingredients, baked in a Dutch Oven."
    format_run(sub1.text_frame.paragraphs[0].runs[0], font_size_pt=18, bold=False, color_rgb=(0x47, 0x55, 0x69))

    # -------------------------------------------------------------
    # Slide 2: Introduction & Overview (Two-column layout)
    # -------------------------------------------------------------
    slide2 = prs.slides.add_slide(prs.slide_layouts[3])
    title2 = slide2.shapes.title
    title2.text = "Introduction & Overview"
    format_run(title2.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    # Left content placeholder for text
    body2 = slide2.placeholders[1]
    body2.left = Inches(0.8)
    body2.top = Inches(1.6)
    body2.width = Inches(4.5)
    body2.height = Inches(5.2)
    style_card_placeholder(body2)

    p_lead = body2.text_frame.paragraphs[0]
    p_lead.text = "OVERVIEW & PHILOSOPHY"
    p_lead.space_before = Pt(0)
    p_lead.space_after = Pt(12)
    remove_bullets(p_lead)
    format_run(p_lead.runs[0], font_size_pt=11, bold=True, color_rgb=(0x03, 0x69, 0xA1))

    p1 = body2.text_frame.add_paragraph()
    p1.text = (
        "The simplicity of this no-knead bread is what makes it a household favorite. "
        "Your entire home will fill with the aroma of fresh artisan bakery bread as it bakes."
    )
    p1.space_before = Pt(0)
    p1.space_after = Pt(12)
    remove_bullets(p1)
    format_run(p1.runs[0], font_size_pt=13.5, color_rgb=(0x33, 0x41, 0x55))

    p2 = body2.text_frame.add_paragraph()
    p2.text = (
        "Requiring zero special equipment and just four basic pantry staples, "
        "this method delivers a golden, blistered crust and a moist, airy crumb."
    )
    p2.space_before = Pt(0)
    p2.space_after = Pt(12)
    remove_bullets(p2)
    format_run(p2.runs[0], font_size_pt=13.5, color_rgb=(0x33, 0x41, 0x55))

    p3 = body2.text_frame.add_paragraph()
    p3.text = (
        "A slow, 12 to 18-hour room-temperature fermentation develops rich flavor "
        "and a chewy artisan texture without any manual kneading."
    )
    p3.space_before = Pt(0)
    p3.space_after = Pt(0)
    remove_bullets(p3)
    format_run(p3.runs[0], font_size_pt=13.5, color_rgb=(0x33, 0x41, 0x55))

    # Remove unused right placeholder from DOM
    slide2.shapes._spTree.remove(slide2.placeholders[2]._element)

    # Add bread photo cleanly on the right side
    left = Inches(5.6)
    top = Inches(1.6)
    pic = slide2.shapes.add_picture(str(IMAGE_PATH), left, top, width=Inches(3.7))
    pic.name = "Artisan Bread Loaf"
    cNvPr = pic._element.xpath(".//p:cNvPr")[0]
    cNvPr.set(
        "descr",
        "A freshly baked round loaf of golden-brown artisanal no-knead bread with a crispy, flour-dusted crust in a Dutch oven.",
    )

    # -------------------------------------------------------------
    # Slide 3: Ingredients
    # -------------------------------------------------------------
    slide3 = prs.slides.add_slide(prs.slide_layouts[1])
    title3 = slide3.shapes.title
    title3.text = "Ingredients"
    format_run(title3.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    # Hide default content placeholder and add accessible table
    slide3.shapes._spTree.remove(slide3.placeholders[1]._element)

    rows, cols = 5, 3
    left = Inches(1.0)
    top = Inches(1.8)
    width = Inches(8.0)
    height = Inches(3.2)
    tbl_shape = slide3.shapes.add_table(rows, cols, left, top, width, height)
    tbl_shape.name = "Ingredients Table"
    table = tbl_shape.table

    # Set Header Row explicitly
    tblPr = tbl_shape._element.xpath(".//a:tblPr")[0]
    tblPr.set("firstRow", "1")

    headers = ["Ingredient", "Quantity", "Notes"]
    data = [
        ["All-Purpose Flour", "3 cups", "Unbleached white flour preferred"],
        ["Salt", "1 3/4 tsp", "Fine sea salt or kosher salt"],
        ["Active Dry Yeast", "1/2 tsp", "Instant or dry active yeast"],
        ["Water", "1 1/2 cups", "Room temperature (approx. 70°F)"],
    ]

    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x0F, 0x17, 0x2A)
        cell.text = h_text
        format_run(cell.text_frame.paragraphs[0].runs[0], font_size_pt=14, bold=True, color_rgb=(0xFF, 0xFF, 0xFF))

    for row_idx, row_data in enumerate(data, start=1):
        bg_rgb = RGBColor(0xF8, 0xFA, 0xFC) if row_idx % 2 == 1 else RGBColor(0xFF, 0xFF, 0xFF)
        for col_idx, val in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_rgb
            cell.text = val
            format_run(cell.text_frame.paragraphs[0].runs[0], font_size_pt=13.5, bold=False, color_rgb=(0x33, 0x41, 0x55))

    # -------------------------------------------------------------
    # Slide 4: Instructions (Exemplary Accessible Numbered List)
    # -------------------------------------------------------------
    slide4 = prs.slides.add_slide(prs.slide_layouts[1])
    title4 = slide4.shapes.title
    title4.text = "Step-by-Step Instructions"
    format_run(title4.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    body4 = slide4.placeholders[1]
    body4.left = Inches(0.8)
    body4.top = Inches(1.6)
    body4.width = Inches(8.4)
    body4.height = Inches(5.2)
    style_card_placeholder(body4)

    steps = [
        (
            "Whisk Dry Ingredients — ",
            "In a large bowl, whisk together 3 cups all-purpose flour, 1¾ tsp fine salt, and ½ tsp active dry yeast.",
        ),
        (
            "Form Shaggy Dough — ",
            "Pour in 1½ cups room-temperature water. Mix with a wooden spoon or spatula until thoroughly incorporated.",
        ),
        (
            "Overnight Counter Ferment — ",
            "Cover bowl tightly with plastic wrap. Let rest at room temperature for 12 to 18 hours until bubbly and doubled.",
        ),
        (
            "Preheat Oven & Dutch Oven — ",
            "Place covered cast iron Dutch oven inside oven. Preheat both together to 450°F (230°C) for at least 30 minutes.",
        ),
        (
            "Shape & Transfer Dough — ",
            "With generously floured hands, turn dough onto a floured counter, shape into a ball, and drop into hot pot.",
        ),
        (
            "Covered & Uncovered Bake — ",
            "Bake covered 30 minutes to trap steam. Remove lid and bake 15 to 20 minutes more until crust is deeply golden.",
        ),
    ]

    p0 = body4.text_frame.paragraphs[0]
    p0.space_before = Pt(0)
    p0.space_after = Pt(10)
    set_numbered_list(p0)
    r_lead = p0.runs[0] if p0.runs else p0.add_run()
    r_lead.text = steps[0][0]
    format_run(r_lead, font_size_pt=13.5, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
    r_body = p0.add_run()
    r_body.text = steps[0][1]
    format_run(r_body, font_size_pt=13.0, bold=False, color_rgb=(0x33, 0x41, 0x55))

    for step_title, step_desc in steps[1:]:
        p = body4.text_frame.add_paragraph()
        p.space_before = Pt(0)
        p.space_after = Pt(10)
        set_numbered_list(p)
        r_l = p.add_run()
        r_l.text = step_title
        format_run(r_l, font_size_pt=13.5, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
        r_d = p.add_run()
        r_d.text = step_desc
        format_run(r_d, font_size_pt=13.0, bold=False, color_rgb=(0x33, 0x41, 0x55))

    # -------------------------------------------------------------
    # Slide 5: Serving & Storage Tips (Exemplary Accessible Bullet List)
    # -------------------------------------------------------------
    slide5 = prs.slides.add_slide(prs.slide_layouts[1])
    title5 = slide5.shapes.title
    title5.text = "Serving & Storage Tips"
    format_run(title5.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    body5 = slide5.placeholders[1]
    body5.left = Inches(0.8)
    body5.top = Inches(1.6)
    body5.width = Inches(8.4)
    body5.height = Inches(5.2)
    style_card_placeholder(body5)

    tips = [
        (
            "Cool Completely Before Slicing: ",
            "Allow loaf to rest on a wire cooling rack for at least 1 full hour. Slicing warm bread compresses steam and creates a gummy, dense crumb.",
        ),
        (
            "Preserve Crust Crispness: ",
            "Store cut-side down in a breathable paper bag or wooden bread box at room temperature for up to 3 days. Never store in plastic bags, which soften the crust.",
        ),
        (
            "Artisanal Serving Pairings: ",
            "Slice thick and serve warm with salted European cultured butter, extra virgin olive oil with flaky sea salt, artisan cheeses, or alongside hearty soups.",
        ),
        (
            "Baker's Companion Guide: ",
            "Looking for crumb analysis, hydration scaling, and Dutch oven troubleshooting? ",
        ),
    ]

    p0 = body5.text_frame.paragraphs[0]
    p0.space_before = Pt(0)
    p0.space_after = Pt(14)
    set_bullet_list(p0)
    r_lead = p0.runs[0] if p0.runs else p0.add_run()
    r_lead.text = tips[0][0]
    format_run(r_lead, font_size_pt=13.5, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
    r_body = p0.add_run()
    r_body.text = tips[0][1]
    format_run(r_body, font_size_pt=13.0, bold=False, color_rgb=(0x33, 0x41, 0x55))

    for idx, (tip_title, tip_desc) in enumerate(tips[1:], start=1):
        p = body5.text_frame.add_paragraph()
        p.space_before = Pt(0)
        p.space_after = Pt(14)
        set_bullet_list(p)
        r_l = p.add_run()
        r_l.text = tip_title
        format_run(r_l, font_size_pt=13.5, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
        r_d = p.add_run()
        r_d.text = tip_desc
        format_run(r_d, font_size_pt=13.0, bold=False, color_rgb=(0x33, 0x41, 0x55))

        if idx == 3:
            r_link = p.add_run()
            r_link.text = "Explore the Complete Artisan Dutch Oven Guide"
            r_link.hyperlink.address = "https://example.com/artisan-bread-guide"
            format_run(r_link, font_size_pt=13.0, bold=True, color_rgb=(0x03, 0x69, 0xA1))

    prs.save(str(output_path))
    print(f"Saved accessible presentation to: {output_path}")


def build_test_deck(output_path: Path):
    """Build presentation that visually mirrors No Knead Bread.pptx but contains accessibility barriers."""
    prs = Presentation()

    # Barrier 1: title-missing (strip core properties title)
    prs.core_properties.title = ""

    # Barrier 2: document-restricted-access (MS Accessibility: Restricted access)
    p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    mv = etree.SubElement(prs._element, f"{{{p_ns}}}modifyVerifier")
    mv.set("cryptProviderType", "rsaAES")
    mv.set("cryptAlgorithmClass", "hash")

    # Barrier 3: section-name-default & section-name-duplicate (MS Accessibility: Default/Duplicate section name)
    extLst = etree.SubElement(
        prs._element,
        "{http://schemas.openxmlformats.org/presentationml/2006/main}extLst",
    )
    ext = etree.SubElement(
        extLst,
        "{http://schemas.openxmlformats.org/presentationml/2006/main}ext",
        attrib={"uri": "{521415D9-36F7-43E2-AB2F-EE9909263876}"},
    )
    sectionLst = etree.SubElement(
        ext,
        "{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst",
        nsmap={"p14": "http://schemas.microsoft.com/office/powerpoint/2010/main"},
    )
    for idx, name in enumerate(["Default Section", "Default Section", "Default Section"]):
        etree.SubElement(
            sectionLst,
            "{http://schemas.microsoft.com/office/powerpoint/2010/main}section",
            attrib={"name": name, "id": f"{{default-sec-{idx}}}"},
        )

    # -------------------------------------------------------------
    # Slide 1: Title Slide (with reading-order-inverted & color-contrast-insufficient)
    # -------------------------------------------------------------
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    title1 = slide1.shapes.title
    title1.text = "No Knead Bread"
    format_run(title1.text_frame.paragraphs[0].runs[0], font_size_pt=40, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    sub1 = slide1.placeholders[1]
    sub1.text = "No kneading required, 4 simple ingredients, baked in a Dutch Oven."
    # Barrier 4: color-contrast-insufficient (MS Accessibility: Hard-to-read text contrast)
    # Light slate #94A3B8 on white background gives contrast ratio 2.44:1 (< 3.0:1 / 4.5:1 required)
    format_run(sub1.text_frame.paragraphs[0].runs[0], font_size_pt=18, bold=False, color_rgb=(0x94, 0xA3, 0xB8))

    # Strip language tag on Slide 1
    rPr = sub1.text_frame.paragraphs[0].runs[0]._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
    if rPr is not None and "lang" in rPr.attrib:
        del rPr.attrib["lang"]

    # Barrier 5: reading-order-inverted (MS Accessibility: Check reading order)
    spTree1 = slide1.shapes._spTree
    spTree1.remove(title1._element)
    spTree1.append(title1._element)

    # -------------------------------------------------------------
    # Slide 2: Introduction & Overview (with slide-title-missing & image-alt-missing)
    # -------------------------------------------------------------
    slide2 = prs.slides.add_slide(prs.slide_layouts[3])
    title2 = slide2.shapes.title
    # Barrier 6: slide-title-missing (MS Accessibility: Missing slide title)
    if title2:
        title2.text = ""

    body2 = slide2.placeholders[1]
    body2.left = Inches(0.8)
    body2.top = Inches(1.6)
    body2.width = Inches(4.5)
    body2.height = Inches(5.2)
    style_card_placeholder(body2)

    p_lead = body2.text_frame.paragraphs[0]
    p_lead.text = "OVERVIEW & PHILOSOPHY"
    p_lead.space_before = Pt(0)
    p_lead.space_after = Pt(12)
    remove_bullets(p_lead)
    format_run(p_lead.runs[0], font_size_pt=11, bold=True, color_rgb=(0x03, 0x69, 0xA1))

    p1 = body2.text_frame.add_paragraph()
    p1.text = (
        "The simplicity of this no-knead bread is what makes it a household favorite. "
        "Your entire home will fill with the aroma of fresh artisan bakery bread as it bakes."
    )
    p1.space_before = Pt(0)
    p1.space_after = Pt(12)
    remove_bullets(p1)
    format_run(p1.runs[0], font_size_pt=13.5, color_rgb=(0x33, 0x41, 0x55))

    p2 = body2.text_frame.add_paragraph()
    p2.text = (
        "Requiring zero special equipment and just four basic pantry staples, "
        "this method delivers a golden, blistered crust and a moist, airy crumb."
    )
    p2.space_before = Pt(0)
    p2.space_after = Pt(12)
    remove_bullets(p2)
    format_run(p2.runs[0], font_size_pt=13.5, color_rgb=(0x33, 0x41, 0x55))

    p3 = body2.text_frame.add_paragraph()
    p3.text = (
        "A slow, 12 to 18-hour room-temperature fermentation develops rich flavor "
        "and a chewy artisan texture without any manual kneading."
    )
    p3.space_before = Pt(0)
    p3.space_after = Pt(0)
    remove_bullets(p3)
    format_run(p3.runs[0], font_size_pt=13.5, color_rgb=(0x33, 0x41, 0x55))

    slide2.shapes._spTree.remove(slide2.placeholders[2]._element)

    # Bread photo WITHOUT alt text (Barrier 6)
    left = Inches(5.6)
    top = Inches(1.6)
    pic = slide2.shapes.add_picture(str(IMAGE_PATH), left, top, width=Inches(3.7))
    pic.name = "Artisan Bread Loaf"
    cNvPr = pic._element.xpath(".//p:cNvPr")[0]
    if "descr" in cNvPr.attrib:
        del cNvPr.attrib["descr"]

    # -------------------------------------------------------------
    # Slide 3: Ingredients (with Barrier 7: table-header-missing)
    # -------------------------------------------------------------
    slide3 = prs.slides.add_slide(prs.slide_layouts[1])
    title3 = slide3.shapes.title
    title3.text = "Ingredients"
    format_run(title3.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    slide3.shapes._spTree.remove(slide3.placeholders[1]._element)

    rows, cols = 5, 3
    left = Inches(1.0)
    top = Inches(1.8)
    width = Inches(8.0)
    height = Inches(3.2)
    tbl_shape = slide3.shapes.add_table(rows, cols, left, top, width, height)
    tbl_shape.name = "Ingredients Table"
    table = tbl_shape.table

    # Barrier 7: table-header-missing (firstRow="0")
    tblPr = tbl_shape._element.xpath(".//a:tblPr")[0]
    tblPr.set("firstRow", "0")

    headers = ["Ingredient", "Quantity", "Notes"]
    data = [
        ["All-Purpose Flour", "3 cups", "Unbleached white flour preferred"],
        ["Salt", "1 3/4 tsp", "Fine sea salt or kosher salt"],
        ["Active Dry Yeast", "1/2 tsp", "Instant or dry active yeast"],
        ["Water", "1 1/2 cups", "Room temperature (approx. 70°F)"],
    ]

    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x0F, 0x17, 0x2A)
        cell.text = h_text
        format_run(cell.text_frame.paragraphs[0].runs[0], font_size_pt=14, bold=True, color_rgb=(0xFF, 0xFF, 0xFF))

    for row_idx, row_data in enumerate(data, start=1):
        bg_rgb = RGBColor(0xF8, 0xFA, 0xFC) if row_idx % 2 == 1 else RGBColor(0xFF, 0xFF, 0xFF)
        for col_idx, val in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_rgb
            cell.text = val
            format_run(cell.text_frame.paragraphs[0].runs[0], font_size_pt=13.5, bold=False, color_rgb=(0x33, 0x41, 0x55))

    # Barrier 9: table-merged-cells (MS Accessibility: Use of merged or split cells)
    cell_water = table.cell(4, 1)
    cell_water.text = "1 1/2 cups (Room temperature, approx. 70°F)"
    cell_water._tc.set("gridSpan", "2")
    table.cell(4, 2).text = ""

    # -------------------------------------------------------------
    # Slide 4: Instructions (Step-by-Step Instructions)
    # -------------------------------------------------------------
    slide4 = prs.slides.add_slide(prs.slide_layouts[1])
    title4 = slide4.shapes.title
    if title4:
        title4.text = "Step-by-Step Instructions"
        format_run(title4.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    body4 = slide4.placeholders[1]
    body4.left = Inches(0.8)
    body4.top = Inches(1.6)
    body4.width = Inches(8.4)
    body4.height = Inches(5.2)
    style_card_placeholder(body4)

    steps = [
        (
            "Whisk Dry Ingredients — ",
            "In a large bowl, whisk together 3 cups all-purpose flour, 1¾ tsp fine salt, and ½ tsp active dry yeast.",
        ),
        (
            "Form Shaggy Dough — ",
            "Pour in 1½ cups room-temperature water. Mix with a wooden spoon or spatula until thoroughly incorporated.",
        ),
        (
            "Overnight Counter Ferment — ",
            "Cover bowl tightly with plastic wrap. Let rest at room temperature for 12 to 18 hours until bubbly and doubled.",
        ),
        (
            "Preheat Oven & Dutch Oven — ",
            "Place covered cast iron Dutch oven inside oven. Preheat both together to 450°F (230°C) for at least 30 minutes.",
        ),
        (
            "Shape & Transfer Dough — ",
            "With generously floured hands, turn dough onto a floured counter, shape into a ball, and drop into hot pot.",
        ),
        (
            "Covered & Uncovered Bake — ",
            "Bake covered 30 minutes to trap steam. Remove lid and bake 15 to 20 minutes more until crust is deeply golden.",
        ),
    ]

    p0 = body4.text_frame.paragraphs[0]
    p0.space_before = Pt(0)
    p0.space_after = Pt(10)
    set_numbered_list(p0)
    r_lead = p0.runs[0] if p0.runs else p0.add_run()
    r_lead.text = steps[0][0]
    format_run(r_lead, font_size_pt=13.5, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
    r_body = p0.add_run()
    r_body.text = steps[0][1]
    format_run(r_body, font_size_pt=13.0, bold=False, color_rgb=(0x33, 0x41, 0x55))

    for step_title, step_desc in steps[1:]:
        p = body4.text_frame.add_paragraph()
        p.space_before = Pt(0)
        p.space_after = Pt(10)
        set_numbered_list(p)
        r_l = p.add_run()
        r_l.text = step_title
        format_run(r_l, font_size_pt=13.5, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
        r_d = p.add_run()
        r_d.text = step_desc
        format_run(r_d, font_size_pt=13.0, bold=False, color_rgb=(0x33, 0x41, 0x55))

    # -------------------------------------------------------------
    # Slide 5: Serving & Storage Tips (with slide-title-duplicate & link-text-vague)
    # -------------------------------------------------------------
    slide5 = prs.slides.add_slide(prs.slide_layouts[1])
    title5 = slide5.shapes.title
    if title5:
        # Barrier 10: slide-title-duplicate (MS Accessibility: Duplicate slide title)
        title5.text = "Step-by-Step Instructions"
        format_run(title5.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    body5 = slide5.placeholders[1]
    body5.left = Inches(0.8)
    body5.top = Inches(1.6)
    body5.width = Inches(8.4)
    body5.height = Inches(5.2)
    style_card_placeholder(body5)

    tips = [
        (
            "Cool Completely Before Slicing: ",
            "Allow loaf to rest on a wire cooling rack for at least 1 full hour. Slicing warm bread compresses steam and creates a gummy, dense crumb.",
        ),
        (
            "Preserve Crust Crispness: ",
            "Store cut-side down in a breathable paper bag or wooden bread box at room temperature for up to 3 days. Never store in plastic bags, which soften the crust.",
        ),
        (
            "Artisanal Serving Pairings: ",
            "Slice thick and serve warm with salted European cultured butter, extra virgin olive oil with flaky sea salt, artisan cheeses, or alongside hearty soups.",
        ),
        (
            "Baker's Companion Guide: ",
            "Looking for crumb analysis, hydration scaling, and Dutch oven troubleshooting? ",
        ),
    ]

    p0 = body5.text_frame.paragraphs[0]
    p0.space_before = Pt(0)
    p0.space_after = Pt(14)
    set_bullet_list(p0)
    r_lead = p0.runs[0] if p0.runs else p0.add_run()
    r_lead.text = tips[0][0]
    format_run(r_lead, font_size_pt=13.5, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
    r_body = p0.add_run()
    r_body.text = tips[0][1]
    format_run(r_body, font_size_pt=13.0, bold=False, color_rgb=(0x33, 0x41, 0x55))

    for idx, (tip_title, tip_desc) in enumerate(tips[1:], start=1):
        p = body5.text_frame.add_paragraph()
        p.space_before = Pt(0)
        p.space_after = Pt(14)
        set_bullet_list(p)
        r_l = p.add_run()
        r_l.text = tip_title
        format_run(r_l, font_size_pt=13.5, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
        r_d = p.add_run()
        r_d.text = tip_desc
        format_run(r_d, font_size_pt=13.0, bold=False, color_rgb=(0x33, 0x41, 0x55))

        if idx == 3:
            # Barrier 8: link-text-vague ("click here")
            r_link = p.add_run()
            r_link.text = "click here"
            r_link.hyperlink.address = "https://example.com/artisan-bread-guide"
            format_run(r_link, font_size_pt=13.0, bold=True, color_rgb=(0x03, 0x69, 0xA1))

    # Strip language tags from several text runs across slides to trigger language-missing
    for s in prs.slides:
        for sh in s.shapes:
            if getattr(sh, "has_text_frame", False):
                for par in sh.text_frame.paragraphs:
                    for run in par.runs:
                        rPr_elem = run._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
                        if rPr_elem is not None and "lang" in rPr_elem.attrib:
                            del rPr_elem.attrib["lang"]

    prs.save(str(output_path))
    print(f"Saved test presentation with barriers to: {output_path}")


if __name__ == "__main__":
    good_deck = EXAMPLES_DIR / "No Knead Bread.pptx"
    bad_deck = EXAMPLES_DIR / "No Knead Bread-test.pptx"

    build_accessible_deck(good_deck)
    build_test_deck(bad_deck)

    print("\n--- Auditing Accessible Deck ---")
    good_res = audit_file(good_deck)
    print(f"Findings: {good_res['summary']['total']}")
    print(f"Pass: {good_res['summary']['pass']}")

    print("\n--- Auditing Test Barrier Deck ---")
    bad_res = audit_file(bad_deck)
    print(f"Findings: {bad_res['summary']['total']}")
    rule_ids = {f['rule_id'] for f in bad_res['findings']}
    print(f"Rules triggered ({len(rule_ids)}):")
    for r in sorted(rule_ids):
        print(f"  - {r}")
