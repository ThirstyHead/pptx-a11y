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


def format_run(run, font_size_pt=15, bold=False, color_rgb=(0x0F, 0x17, 0x2A), lang="en-US"):
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
    body2.top = Inches(1.8)
    body2.width = Inches(4.5)
    body2.height = Inches(4.8)

    p1 = body2.text_frame.paragraphs[0]
    p1.text = (
        "The simplicity of this no-knead bread is what makes it a household favorite. "
        "Your entire home will fill with the aroma of fresh artisan bakery bread as it bakes."
    )
    p1.space_before = Pt(0)
    p1.space_after = Pt(12)
    remove_bullets(p1)
    format_run(p1.runs[0], font_size_pt=15, color_rgb=(0x33, 0x41, 0x55))

    p2 = body2.text_frame.add_paragraph()
    p2.text = (
        "Requiring zero special equipment and just four basic pantry staples, "
        "this method delivers a golden, blistered crust and a moist, airy crumb."
    )
    p2.space_before = Pt(0)
    p2.space_after = Pt(12)
    remove_bullets(p2)
    format_run(p2.runs[0], font_size_pt=15, color_rgb=(0x33, 0x41, 0x55))

    p3 = body2.text_frame.add_paragraph()
    p3.text = (
        "A slow, 12 to 18-hour room-temperature fermentation develops rich flavor "
        "and a chewy artisan texture without any manual kneading."
    )
    p3.space_before = Pt(0)
    p3.space_after = Pt(0)
    remove_bullets(p3)
    format_run(p3.runs[0], font_size_pt=15, color_rgb=(0x33, 0x41, 0x55))

    # Remove unused right placeholder from DOM
    slide2.shapes._spTree.remove(slide2.placeholders[2]._element)

    # Add bread photo cleanly on the right side
    left = Inches(5.5)
    top = Inches(1.8)
    pic = slide2.shapes.add_picture(str(IMAGE_PATH), left, top, width=Inches(3.8))
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
        cell.text = h_text
        format_run(cell.text_frame.paragraphs[0].runs[0], font_size_pt=15, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, val in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.text = val
            format_run(cell.text_frame.paragraphs[0].runs[0], font_size_pt=14, bold=False, color_rgb=(0x33, 0x41, 0x55))

    # -------------------------------------------------------------
    # Slide 4: Instructions
    # -------------------------------------------------------------
    slide4 = prs.slides.add_slide(prs.slide_layouts[1])
    title4 = slide4.shapes.title
    title4.text = "Step-by-Step Instructions"
    format_run(title4.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    body4 = slide4.placeholders[1]
    body4.top = Inches(1.8)
    body4.height = Inches(5.0)
    steps = [
        "1. In a large bowl, whisk together flour, salt, and yeast.",
        "2. Add room-temperature water; stir with a wooden spoon until a shaggy dough forms.",
        "3. Cover tightly with plastic wrap; let rest at room temperature for 12 to 18 hours.",
        "4. Preheat oven and covered cast iron Dutch oven to 450°F (230°C).",
        "5. Gently turn dough onto a floured surface, shape into a ball, and transfer to hot pot.",
        "6. Bake covered for 30 minutes, then uncovered for 15 to 20 minutes until golden brown.",
    ]
    p0 = body4.text_frame.paragraphs[0]
    p0.text = steps[0]
    p0.space_before = Pt(0)
    p0.space_after = Pt(6)
    remove_bullets(p0)
    format_run(p0.runs[0], font_size_pt=14, color_rgb=(0x33, 0x41, 0x55))

    for s in steps[1:]:
        p = body4.text_frame.add_paragraph()
        p.text = s
        p.space_before = Pt(0)
        p.space_after = Pt(6)
        remove_bullets(p)
        format_run(p.runs[0], font_size_pt=14, color_rgb=(0x33, 0x41, 0x55))

    # -------------------------------------------------------------
    # Slide 5: Serving & Storage Tips
    # -------------------------------------------------------------
    slide5 = prs.slides.add_slide(prs.slide_layouts[1])
    title5 = slide5.shapes.title
    title5.text = "Serving & Storage Tips"
    format_run(title5.text_frame.paragraphs[0].runs[0], font_size_pt=30, bold=True, color_rgb=(0x0F, 0x17, 0x2A))

    body5 = slide5.placeholders[1]
    body5.top = Inches(1.8)
    body5.height = Inches(4.8)
    tips = [
        ("Cool Completely: ", "Allow loaf to rest on a wire cooling rack for at least 1 hour before slicing to prevent gumminess."),
        ("Storage: ", "Store at room temperature in a paper bag or bread box for up to 3 days. Avoid plastic bags."),
        ("Serving: ", "Perfect warm with salted European butter, extra virgin olive oil, or hearty soups."),
    ]
    p0 = body5.text_frame.paragraphs[0]
    p0.space_before = Pt(0)
    p0.space_after = Pt(10)
    remove_bullets(p0)
    r_lead = p0.runs[0] if p0.runs else p0.add_run()
    r_lead.text = tips[0][0]
    format_run(r_lead, font_size_pt=15, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
    r_body = p0.add_run()
    r_body.text = tips[0][1]
    format_run(r_body, font_size_pt=15, bold=False, color_rgb=(0x33, 0x41, 0x55))

    for lead, text in tips[1:]:
        p = body5.text_frame.add_paragraph()
        p.space_before = Pt(0)
        p.space_after = Pt(10)
        remove_bullets(p)
        r_l = p.add_run()
        r_l.text = lead
        format_run(r_l, font_size_pt=15, bold=True, color_rgb=(0x0F, 0x17, 0x2A))
        r_b = p.add_run()
        r_b.text = text
        format_run(r_b, font_size_pt=15, bold=False, color_rgb=(0x33, 0x41, 0x55))

    # Add accessible link with descriptive text
    p_link = body5.text_frame.add_paragraph()
    p_link.space_before = Pt(8)
    remove_bullets(p_link)
    r_link = p_link.add_run()
    r_link.text = "Explore the complete Artisan Dutch Oven Baking Guide"
    r_link.hyperlink.address = "https://example.com/artisan-bread-guide"
    format_run(r_link, font_size_pt=15, bold=True, color_rgb=(0x02, 0x84, 0xC7))

    prs.save(str(output_path))
    print(f"Saved accessible presentation to: {output_path}")


def build_test_deck(output_path: Path):
    """Build presentation containing at least one instance of EVERY rule tested by pptx-a11y."""
    prs = Presentation()

    # Barrier 1: title-missing (remove core title)
    prs.core_properties.title = ""

    # Barrier 2: section-name-default & Barrier 3: section-name-duplicate
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
    for name in ["Default Section", "Baking Section", "Baking Section"]:
        etree.SubElement(
            sectionLst,
            "{http://schemas.microsoft.com/office/powerpoint/2010/main}section",
            attrib={"name": name, "id": f"{{{name.lower()}}}"},
        )

    # Barrier 4: document-restricted-access (<p:modifyVerifier>)
    etree.SubElement(
        prs._element,
        "{http://schemas.openxmlformats.org/presentationml/2006/main}modifyVerifier",
        attrib={"cryptProviderType": "rsaAES", "hashData": "dummy"},
    )

    # -------------------------------------------------------------
    # Slide 1: Title Slide (with reading-order-inverted & language-missing)
    # -------------------------------------------------------------
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    title1 = slide1.shapes.title
    title1.text = "No Knead Bread"
    title1.top = Inches(1.5)
    title1.left = Inches(1.0)

    sub1 = slide1.placeholders[1]
    sub1.text = "No kneading required, 4 simple ingredients."
    sub1.top = Inches(4.0)
    sub1.left = Inches(1.0)

    # Barrier 5: language-missing (no lang on subtitle text run)
    rPr = sub1.text_frame.paragraphs[0].runs[0]._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
    if rPr is not None and "lang" in rPr.attrib:
        del rPr.attrib["lang"]

    # Barrier 6: reading-order-inverted
    # Invert order in spTree: place sub1 before title1 in spTree, but sub1 is spatially lower
    spTree = slide1.shapes._spTree
    spTree.remove(title1._element)
    spTree.append(title1._element)  # title is now last in DOM but topmost spatially!

    # -------------------------------------------------------------
    # Slide 2: Missing Slide Title & Missing Alt Text & Vague Link
    # -------------------------------------------------------------
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    title2 = slide2.shapes.title
    # Barrier 7: slide-title-missing (empty title)
    title2.text = ""

    body2 = slide2.placeholders[1]
    body2.text = "Overview of the bread baking process and ingredients."
    set_run_lang(body2.text_frame.paragraphs[0].runs[0])

    # Barrier 8: link-text-vague ("click here")
    p_link = body2.text_frame.add_paragraph()
    r_before = p_link.add_run()
    r_before.text = "For baking tips, "
    set_run_lang(r_before)
    r_link = p_link.add_run()
    r_link.text = "click here"
    r_link.hyperlink.address = "https://example.com/recipe"
    set_run_lang(r_link)

    # Barrier 9: image-alt-missing (no alt text)
    pic = slide2.shapes.add_picture(str(IMAGE_PATH), Inches(5.5), Inches(2.2), width=Inches(3.8))
    pic.name = "Bread Photo"
    # Ensure no descr attribute on cNvPr
    cNvPr = pic._element.xpath(".//p:cNvPr")[0]
    if "descr" in cNvPr.attrib:
        del cNvPr.attrib["descr"]

    # -------------------------------------------------------------
    # Slide 3: Ingredients (Table Header Missing & Table Merged Cells)
    # -------------------------------------------------------------
    # Barrier 10: slide-title-duplicate (name this slide "Recipe Details")
    slide3 = prs.slides.add_slide(prs.slide_layouts[1])
    title3 = slide3.shapes.title
    title3.text = "Recipe Details"
    set_run_lang(title3.text_frame.paragraphs[0].runs[0])

    slide3.shapes._spTree.remove(slide3.placeholders[1]._element)
    tbl_shape = slide3.shapes.add_table(3, 3, Inches(1.0), Inches(2.0), Inches(8.0), Inches(2.5))
    tbl_shape.name = "Ingredients Table"

    # Barrier 11: table-header-missing (firstRow="0")
    tblPr = tbl_shape._element.xpath(".//a:tblPr")[0]
    tblPr.set("firstRow", "0")

    # Barrier 12: table-merged-cells (gridSpan="2")
    cell0 = tbl_shape.table.cell(0, 0)
    cell0._tc.set("gridSpan", "2")
    cell0.text = "Merged Ingredients Header"
    set_run_lang(cell0.text_frame.paragraphs[0].runs[0])

    # -------------------------------------------------------------
    # Slide 4: Instructions (Chart Missing Alt & Media Subtitles Missing & Duplicate Slide Title)
    # -------------------------------------------------------------
    slide4 = prs.slides.add_slide(prs.slide_layouts[1])
    title4 = slide4.shapes.title
    # Barrier 13: slide-title-duplicate ("Recipe Details" repeated)
    title4.text = "Recipe Details"
    set_run_lang(title4.text_frame.paragraphs[0].runs[0])

    # Barrier 14: chart-missing-alt (chart with no alt text)
    chart_data = CategoryChartData()
    chart_data.categories = ["Flour", "Water", "Salt", "Yeast"]
    chart_data.add_series("Grams", (400, 300, 10, 3))
    chart_shape = slide4.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1.0),
        Inches(2.0),
        Inches(4.0),
        Inches(3.0),
        chart_data,
    )
    chart_shape.name = "Ingredient Proportions Chart"
    chart_cNvPr = chart_shape._element.xpath(".//p:cNvPr")[0]
    if "descr" in chart_cNvPr.attrib:
        del chart_cNvPr.attrib["descr"]

    # Barrier 15: media-subtitles-missing (media element with no vtt tracks)
    media_shape = slide4.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(2.0), Inches(3.5), Inches(2.5)
    )
    media_shape.name = "Bread Baking Video"
    etree.SubElement(
        media_shape._element,
        "{http://schemas.openxmlformats.org/presentationml/2006/main}videoFile",
        nsmap={"p": "http://schemas.openxmlformats.org/presentationml/2006/main"},
    )

    # -------------------------------------------------------------
    # Slide 5: Semantic Placeholders Missing (Blank layout with floating txBox)
    # -------------------------------------------------------------
    # Barrier 16: semantic-placeholders-missing (layout 6 is blank, has 0 placeholders)
    slide5 = prs.slides.add_slide(prs.slide_layouts[6])
    txBox = slide5.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.0))
    txBox.name = "Floating Text Box"
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Serving Tips: Let bread cool before slicing. This slide uses floating text without semantic layout."
    set_run_lang(p.runs[0])

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
