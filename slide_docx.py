"""Rebuild PowerPoint PDF pages as graphics with editable Word text boxes."""

from io import BytesIO
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
import pymupdf


_FONT_NAMES = {
    "TimesNewRomanPSMT": "Times New Roman",
    "TimesNewRomanPS-BoldMT": "Times New Roman",
    "TimesNewRomanPS-ItalicMT": "Times New Roman",
    "TimesNewRomanPS-BoldItalicMT": "Times New Roman",
    "ArialMT": "Arial",
    "Arial-BoldMT": "Arial",
    "Arial-ItalicMT": "Arial",
    "Arial-BoldItalicMT": "Arial",
    "CambriaMath": "Cambria Math",
}


def page_anchor(x, y, width, height, *, behind_text=False):
    """Create a floating object measured from the top left of the Word page."""
    return parse_xml(f"""
        <wp:anchor {nsdecls('wp')} distT="0" distB="0" distL="0" distR="0"
                   simplePos="0" relativeHeight="{0 if behind_text else 1}"
                   behindDoc="{int(behind_text)}" locked="0" layoutInCell="0" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="page"><wp:posOffset>{Pt(x)}</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="page"><wp:posOffset>{Pt(y)}</wp:posOffset></wp:positionV>
          <wp:extent cx="{Pt(width)}" cy="{Pt(height)}"/>
          <wp:wrapNone/>
        </wp:anchor>
    """)


def add_backdrop(paragraph, page, width, height):
    image = page.get_pixmap(dpi=144).tobytes("png")
    inline = paragraph.add_run().add_picture(BytesIO(image), width=Pt(width), height=Pt(height))._inline
    shape_id = int(inline.docPr.id)
    anchor = page_anchor(0, 0, width, height, behind_text=True)
    anchor.append(inline.docPr)
    anchor.append(inline.graphic)
    inline.getparent().replace(inline, anchor)
    return shape_id


def add_editable_span(paragraph, span, scale, shape_id):
    x0, y0, x1, y1 = span["bbox"]
    size = span["size"] * scale
    width = (x1 - x0) * scale + max(size, 5)
    height = max((y1 - y0) * scale, size * 1.4)
    anchor = page_anchor(x0 * scale, y0 * scale, width, height)
    properties = OxmlElement("wp:docPr")
    properties.set("id", str(shape_id))
    properties.set("name", f"Text {shape_id}")
    anchor.append(properties)
    graphic = parse_xml(f"""
        <a:graphic {nsdecls('a', 'w')} xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="{Pt(width)}" cy="{Pt(height)}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:noFill/><a:ln><a:noFill/></a:ln>
              </wps:spPr>
              <wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t">
                <a:noAutofit/>
              </wps:bodyPr>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
    """)
    box_paragraph = next(graphic.iter(qn("w:p")))
    text = Paragraph(box_paragraph, paragraph)
    text.paragraph_format.space_before = Pt(0)
    text.paragraph_format.space_after = Pt(0)
    text.paragraph_format.line_spacing = 1.0
    run = text.add_run(span["text"])
    pdf_font = span["font"]
    run.font.name = _FONT_NAMES.get(pdf_font, re.sub(r"[-,](?:BoldItalic|Bold|Italic|Regular)$", "", pdf_font))
    run.font.size = Pt(size)
    run.font.bold = bool(span["flags"] & pymupdf.TEXT_FONT_BOLD)
    run.font.italic = bool(span["flags"] & pymupdf.TEXT_FONT_ITALIC)
    run.font.color.rgb = RGBColor.from_string(f"{span['color']:06X}")
    anchor.append(graphic)
    drawing = OxmlElement("w:drawing")
    drawing.append(anchor)
    paragraph.add_run()._r.append(drawing)


def slides_to_word(pdf, output):
    document = Document()
    last_size = None
    for index, page in enumerate(pdf):
        scale = min(1, 1584 / max(page.rect.width, page.rect.height))
        size = (page.rect.width * scale, page.rect.height * scale)
        if size != last_size:
            section = document.add_section(WD_SECTION.NEW_PAGE) if index else document.sections[0]
            section.page_width, section.page_height = map(Pt, size)
            section.left_margin = section.right_margin = Pt(0)
            section.top_margin = section.bottom_margin = Pt(0)

        paragraph = document.add_paragraph()
        paragraph.paragraph_format.page_break_before = index > 0 and size == last_size
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = Pt(1)
        last_size = size

        spans = []
        if page.rotation == 0:
            blocks = page.get_text("dict", flags=pymupdf.TEXTFLAGS_DICT & ~pymupdf.TEXT_PRESERVE_IMAGES)["blocks"]
            spans = [span for block in blocks for line in block["lines"] if line["dir"] == (1.0, 0.0)
                     for span in line["spans"] if span["text"].strip() and span.get("alpha", 255)]
        for span in spans:
            page.add_redact_annot(span["bbox"], fill=False, cross_out=False)
        if spans:
            page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                                  graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
        shape_id = add_backdrop(paragraph, page, *size)
        for offset, span in enumerate(spans, 1):
            add_editable_span(paragraph, span, scale, shape_id + offset)

    document.save(output)
