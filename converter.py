"""Local file conversions for the Windows desktop application."""

import os
from contextlib import contextmanager
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Lock

import img2pdf
import pypdfium2 as pdfium


_pdf_word_lock = Lock()


def pdf_to_word(source: Path, output: Path) -> None:
    from fontTools.ttLib import TTFont
    from pdf2docx import Converter
    from pdf2docx.font.Fonts import Fonts
    from pdf2docx.page.RawPageFactory import RawPageFactory
    from pdf2docx.page.RawPageFitz import RawPageFitz
    from pdf2docx.text.TextSpan import TextSpan

    def is_folio(text, bbox, width, height, page_number):
        x0, y0, x1, _ = bbox
        return (text.strip() == str(page_number) and x1 - x0 <= width * 0.05
                and abs((x0 + x1) / 2 - width / 2) <= width * 0.03 and y0 >= height * 0.92)

    moved_folios = set()
    embedded_fonts = {}
    preserve_folios = False

    def read_font(pdf, xref):
        if xref not in embedded_fonts:
            embedded_fonts[xref] = None
            try:
                _, extension, _, data = pdf.extract_font(xref)
                if extension in ("ttf", "otf") and data:
                    with TTFont(BytesIO(data)) as font:
                        family = font["name"].getBestFamilyName()
                        if family:
                            embedded_fonts[xref] = (family, None)
                            embedded_fonts[xref] = (family, Fonts.get_line_height_factor(font))
            except Exception:
                # Unavailable or invalid embedded metadata should leave the original fallback intact.
                pass
        return embedded_fonts[xref]

    class DocumentPage(RawPageFitz):
        def process_font(self, fonts):
            super().process_font(fonts)
            aliases = {}
            for xref, _, _, name, *_ in self.page_engine.get_fonts():
                name = name.casefold()
                if name.startswith("cidfont+"):
                    aliases.setdefault(name, set()).add(xref)
            if not aliases:
                return
            for line in self.blocks:
                for span in line.spans:
                    if not isinstance(span, TextSpan):
                        continue
                    refs = aliases.get(span.font.casefold())
                    if refs:
                        candidates = {read_font(self.page_engine.parent, xref) for xref in refs}
                        if len(candidates) == 1 and None not in candidates:
                            span.font, line_height = candidates.pop()
                            if line_height:
                                span.line_height = line_height * span.size

        def parse_section(self, **settings):
            if not preserve_folios:
                return super().parse_section(**settings)
            # A centered folio in the gutter can be mistaken for a third column.
            # pdf2docx then selects it as the right column and loses the body text.
            blocks = list(self.blocks)
            page_number = self.page_engine.number + 1
            body = [block for block in blocks if not is_folio(block.text, block.bbox,
                                                              self.width, self.height, page_number)]
            if len(body) != len(blocks) - 1:
                return super().parse_section(**settings)

            moved_folios.add(page_number)
            self.blocks.reset(body)
            try:
                return super().parse_section(**settings)
            finally:
                self.blocks.reset(blocks)

    with _pdf_word_lock:
        backend = RawPageFactory.MAP["PYMUPDF"]
        converter = Converter(str(source))
        try:
            pages = converter.fitz_doc
            page_count = len(pages)
            if page_count and not pages.needs_pass and any(
                "powerpoint" in (pages.metadata.get(field) or "").casefold()
                for field in ("creator", "producer")
            ):
                from slide_docx import slides_to_word

                slides_to_word(pages, output)
                return
            preserve_folios = bool(page_count and not pages.needs_pass and all(
                page.rotation == 0 and any(
                    is_folio(word[4], word[:4], page.rect.width, page.rect.height, page.number + 1)
                    for word in page.get_text("words")
                ) for page in pages
            ))
            RawPageFactory.MAP["PYMUPDF"] = DocumentPage
            converter.convert(str(output), ignore_page_error=False)
            if moved_folios and len(moved_folios) != page_count:
                preserve_folios = False
                fallback = Converter(str(source))
                try:
                    fallback.convert(str(output), ignore_page_error=False)
                finally:
                    fallback.close()
        finally:
            RawPageFactory.MAP["PYMUPDF"] = backend
            converter.close()

    if moved_folios and len(moved_folios) == page_count:
        add_page_numbers(output)


def add_page_numbers(output: Path) -> None:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    document = Document(output)
    for section in document.sections:
        section.footer_distance = Pt(18)
    paragraph = document.sections[0].footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    paragraph._p.append(field)
    document.save(output)


@contextmanager
def word_application():
    import pythoncom
    import pywintypes
    from win32com.client import DispatchEx

    pythoncom.CoInitialize()
    try:
        try:
            word = DispatchEx("Word.Application")
        except pywintypes.com_error as exc:
            raise RuntimeError("Cần cài và mở Microsoft Word ít nhất một lần trước khi chuyển đổi.") from exc
        try:
            word.Visible = False
            word.DisplayAlerts = 0
            word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            yield word
        finally:
            word.Quit(False)
    finally:
        pythoncom.CoUninitialize()


def word_to_pdf(source: Path, output: Path, word=None) -> None:
    if word is None:
        with word_application() as application:
            word_to_pdf(source, output, application)
        return
    update_links = word.Options.UpdateLinksAtOpen
    try:
        word.Options.UpdateLinksAtOpen = False
        document = word.Documents.Open(
            str(source), ConfirmConversions=False, ReadOnly=True,
            AddToRecentFiles=False, Visible=False, NoEncodingDialog=True,
            PasswordDocument="_unsupported_password_",
        )
        try:
            document.ExportAsFixedFormat(str(output), 17)  # wdExportFormatPDF
        finally:
            document.Close(False)
    finally:
        word.Options.UpdateLinksAtOpen = update_links


def convert_to_file(mode: str, sources: list[Path], output: Path, *, overwrite: bool = True, word=None) -> Path:
    output = output.absolute()
    sources = [source.resolve() for source in sources]
    if output.resolve() in sources:
        raise ValueError("Hãy chọn đường dẫn đầu ra khác tệp gốc.")

    with TemporaryDirectory(prefix=".chuyen-doi-", dir=output.parent) as temporary:
        staged = Path(temporary) / output.name
        if mode == "images_pdf":
            with staged.open("wb") as stream:
                img2pdf.convert(
                    *(str(source) for source in sources), outputstream=stream,
                    rotation=img2pdf.Rotation.ifvalid,
                )
        elif mode == "word_pdf":
            word_to_pdf(sources[0], staged, word)
        elif mode == "pdf_word":
            pdf_to_word(sources[0], staged)
        else:
            raise ValueError(f"Chế độ không hợp lệ: {mode}")
        if not staged.is_file() or staged.stat().st_size == 0:
            raise RuntimeError("Không tạo được tệp đầu ra.")
        if overwrite:
            os.replace(staged, output)
        else:
            staged.rename(output)
    return output


def pdf_to_images(source: Path, destination: Path, image_format: str, dpi: int, progress=None) -> Path:
    with TemporaryDirectory(prefix=".chuyen-doi-", dir=destination) as temporary:
        pages = Path(temporary) / "pages"
        pages.mkdir()
        with pdfium.PdfDocument(source) as pdf:
            total = len(pdf)
            if not total:
                raise ValueError("PDF không có trang nào.")
            digits = max(3, len(str(total)))
            extension = "jpg" if image_format == "JPEG" else image_format.lower()
            options = {"quality": 95} if image_format == "JPEG" else {}
            for index in range(total):
                if progress:
                    progress(index + 1, total)
                page = pdf[index]
                output = pages / f"trang-{index + 1:0{digits}}.{extension}"
                try:
                    bitmap = page.render(scale=dpi / 72)
                    try:
                        with bitmap.to_pil() as image:
                            image.save(output, dpi=(dpi, dpi), **options)
                    finally:
                        bitmap.close()
                finally:
                    page.close()

        output = destination / source.stem
        number = 1
        while True:
            try:
                # On Windows rename fails if the destination already exists.
                pages.rename(output)
                return output
            except FileExistsError:
                number += 1
                output = destination / f"{source.stem} ({number})"
