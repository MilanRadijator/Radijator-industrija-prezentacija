from __future__ import annotations

import html
import re
import shutil
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
DOCX = Path(r"D:\Prezentacija nikola\KATALOG INDUSTRIJSKIH KOTLOVA.docx")
DOCS = ROOT / "docs"
ASSET_DIR = DOCS / "assets" / "full-catalog"
OUT_HTML = DOCS / "index.html"
ALIAS_HTML = DOCS / "full-catalog.html"


def iter_blocks(document: Document):
    body = document.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def uppercase_ratio(text: str) -> float:
    letters = [character for character in text if character.isalpha()]
    if not letters:
        return 0
    return sum(1 for character in letters if character.upper() == character) / len(letters)


def is_heading(paragraph: Paragraph, text: str) -> bool:
    if not text or len(text) > 110:
        return False
    if paragraph.style and paragraph.style.name.startswith("List"):
        return False
    return uppercase_ratio(text) > 0.78 or text.lower() == "o nama"


def is_subheading(paragraph: Paragraph, text: str) -> bool:
    if not text or len(text) > 90:
        return False
    if paragraph.style and paragraph.style.name.startswith("List"):
        return uppercase_ratio(text) > 0.78
    prefixes = (
        "serija ",
        "kaskadni ",
        "dodatna ",
        "primena ",
        "specifičnosti ",
        "hidrauličko ",
        "pozicioniranje ",
    )
    return text.lower().startswith(prefixes)


def table_to_html(table: Table) -> str:
    rows_html: list[str] = []
    for row_index, row in enumerate(table.rows):
        cells = [clean_text(cell.text) for cell in row.cells]
        tag = "th" if row_index == 0 else "td"
        cell_html = "".join(f"<{tag}>{html.escape(cell)}</{tag}>" for cell in cells)
        rows_html.append(f"<tr>{cell_html}</tr>")
    header = rows_html[0] if rows_html else ""
    body = "".join(rows_html[1:])
    return (
        "<div class=\"table-scroll\"><table>"
        f"<thead>{header}</thead><tbody>{body}</tbody>"
        "</table></div>"
    )


def extract_images() -> list[dict[str, str]]:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    images: list[dict[str, str]] = []

    with zipfile.ZipFile(DOCX) as archive:
        media_entries = [
            entry for entry in archive.infolist() if entry.filename.startswith("word/media/")
        ]
        for index, entry in enumerate(media_entries, start=1):
            original_name = Path(entry.filename).name
            suffix = Path(original_name).suffix.lower()
            raw_path = ASSET_DIR / original_name
            raw_path.write_bytes(archive.read(entry))

            target_name = f"catalog-image-{index:02d}"
            display_path: Path | None = None
            status = "included"

            if suffix in {".png", ".jpg", ".jpeg"}:
                display_path = ASSET_DIR / f"{target_name}{suffix}"
                shutil.copyfile(raw_path, display_path)
            elif suffix in {".tif", ".tiff"}:
                display_path = ASSET_DIR / f"{target_name}.png"
                with Image.open(raw_path) as image:
                    image.seek(0)
                    image.convert("RGBA").save(display_path)
            elif suffix == ".wmf":
                # Browser support for WMF is not reliable; keep the original asset visible as a download item.
                status = "original-wmf"
                display_path = None
            else:
                status = "unsupported"

            images.append(
                {
                    "label": f"Slika {index}",
                    "original": original_name,
                    "key": Path(original_name).name.lower(),
                    "display": display_path.name if display_path else "",
                    "status": status,
                }
            )

    return images


def paragraph_image_keys(paragraph: Paragraph) -> list[str]:
    keys: list[str] = []
    for blip in paragraph._element.xpath('.//*[local-name()="blip"]'):
        rel_id = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        if not rel_id:
            continue
        rel = paragraph.part.rels.get(rel_id)
        if rel is not None:
            keys.append(Path(rel.target_ref).name.lower())
    return keys


def image_figure(item: dict[str, str], index: int) -> str:
    if not item.get("display"):
        return (
            "<div class=\"catalog-original-note\">"
            f"<strong>{html.escape(item['label'])}</strong>: originalni fajl "
            f"{html.escape(item['original'])} je sacuvan u assets/full-catalog."
            "</div>"
        )
    return f"""
    <figure class="catalog-figure">
      <img src="assets/full-catalog/{html.escape(item['display'])}" alt="{html.escape(item['label'])} - {html.escape(item['original'])}" />
      <figcaption data-index="{index:02d}">{index:02d}. {html.escape(item['original'])}</figcaption>
    </figure>
    """


def render_copy(blocks: list[dict[str, object]]) -> str:
    rendered: list[str] = []
    for block in blocks:
        block_type = block["type"]
        if block_type == "paragraph":
            rendered.append(f"<p>{html.escape(str(block['text']))}</p>")
        elif block_type == "subheading":
            rendered.append(f"<h3>{html.escape(str(block['text']))}</h3>")
        elif block_type == "list":
            items = "".join(f"<li>{html.escape(str(item))}</li>" for item in block["items"])
            rendered.append(f"<ul class=\"feature-list\">{items}</ul>")
    return "".join(rendered)


def render_section(section: dict[str, object], section_number: int) -> str:
    blocks = list(section["blocks"])
    rendered: list[str] = []
    media_index = 0
    index = 0

    while index < len(blocks):
        block = blocks[index]
        block_type = block["type"]

        if block_type in {"paragraph", "subheading", "list"}:
            copy_end = index
            while copy_end < len(blocks) and blocks[copy_end]["type"] in {"paragraph", "subheading", "list"}:
                copy_end += 1
            if copy_end < len(blocks) and blocks[copy_end]["type"] == "figure":
                media_index += 1
                side = "media-block--image-right" if media_index % 2 else "media-block--image-left"
                paired_copy = blocks[index:copy_end]
                after_figure = copy_end + 1
                copy_length = sum(len(str(item.get("text", ""))) for item in paired_copy)
                if copy_length < 260:
                    while (
                        after_figure < len(blocks)
                        and blocks[after_figure]["type"] in {"paragraph", "subheading", "list"}
                        and len(paired_copy) < 3
                    ):
                        paired_copy.append(blocks[after_figure])
                        after_figure += 1
                rendered.append(
                    f"<div class=\"media-block {side}\">"
                    f"<div class=\"media-copy\">{render_copy(paired_copy)}</div>"
                    f"{blocks[copy_end]['html']}"
                    "</div>"
                )
                index = after_figure
                continue
            rendered.append(render_copy(blocks[index:copy_end]))
            index = copy_end
            continue

        if block_type == "figure":
            copy_end = index + 1
            while copy_end < len(blocks) and blocks[copy_end]["type"] in {"paragraph", "subheading", "list"}:
                copy_end += 1
            if copy_end > index + 1:
                media_index += 1
                side = "media-block--image-left" if media_index % 2 else "media-block--image-right"
                rendered.append(
                    f"<div class=\"media-block {side}\">"
                    f"{block['html']}"
                    f"<div class=\"media-copy\">{render_copy(blocks[index + 1:copy_end])}</div>"
                    "</div>"
                )
                index = copy_end
                continue
            rendered.append(f"<div class=\"technical-visual\">{block['html']}</div>")
        elif block_type == "table":
            rendered.append(str(block["html"]))
        index += 1

    media_class = " catalog-section--media" if any(block["type"] == "figure" for block in blocks) else ""
    return (
        f"<section class=\"catalog-section{media_class}\" id=\"{section['id']}\" data-section=\"{section_number:02d}\">"
        f"<div class=\"section-heading\"><span>{section_number:02d}</span>"
        f"<h2>{html.escape(str(section['title']))}</h2></div>"
        f"{''.join(rendered)}"
        "</section>"
    )


def build_content(images_by_key: dict[str, dict[str, str]]) -> tuple[str, list[str], int, int, set[str]]:
    document = Document(DOCX)
    sections: list[dict[str, object]] = []
    toc: list[str] = []
    used_images: set[str] = set()
    current_section: dict[str, object] | None = None
    paragraph_count = 0
    table_count = 0
    figure_count = 0
    paragraph_buffer: list[str] = []
    list_buffer: list[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph_buffer
        if paragraph_buffer and current_section is not None:
            current_section["blocks"].append(
                {"type": "paragraph", "text": " ".join(paragraph_buffer)}
            )
            paragraph_buffer = []

    def flush_list() -> None:
        nonlocal list_buffer
        if list_buffer and current_section is not None:
            current_section["blocks"].append({"type": "list", "items": list_buffer})
            list_buffer = []

    def ensure_section(title: str = "Uvod") -> dict[str, object]:
        nonlocal current_section
        if current_section is None:
            section_id = f"section-{len(sections) + 1:02d}"
            current_section = {"id": section_id, "title": title, "blocks": []}
            sections.append(current_section)
            toc.append(f"<a href=\"#{section_id}\">{html.escape(title)}</a>")
        return current_section

    def start_section(title: str) -> None:
        nonlocal current_section
        flush_paragraph()
        flush_list()
        section_id = f"section-{len(sections) + 1:02d}"
        current_section = {"id": section_id, "title": title, "blocks": []}
        sections.append(current_section)
        toc.append(f"<a href=\"#{section_id}\">{html.escape(title)}</a>")

    for block in iter_blocks(document):
        if isinstance(block, Paragraph):
            text = clean_text(block.text)
            image_keys = paragraph_image_keys(block)
            if not text and not image_keys:
                continue
            if is_heading(block, text):
                start_section(text)
            else:
                section = ensure_section()
                if text:
                    paragraph_count += 1
                    if is_subheading(block, text):
                        flush_paragraph()
                        flush_list()
                        section["blocks"].append({"type": "subheading", "text": text})
                    elif block.style and block.style.name.startswith("List"):
                        flush_paragraph()
                        list_buffer.append(text)
                    else:
                        flush_list()
                        paragraph_buffer.append(text)
                        ends_with_year = bool(re.search(r"\b(?:19|20)\d{2}\.\s*$", text))
                        if (re.search(r"[.!?]\s*$", text) and not ends_with_year) or len(text) > 240:
                            flush_paragraph()
                if image_keys:
                    flush_paragraph()
                    flush_list()
                    figure_group: list[str] = []
                    for key in image_keys:
                        item = images_by_key.get(key)
                        if not item:
                            continue
                        used_images.add(key)
                        figure_count += 1
                        figure_group.append(image_figure(item, figure_count))
                    if figure_group:
                        section["blocks"].append(
                            {"type": "figure", "html": f"<div class=\"figure-row\">{''.join(figure_group)}</div>"}
                        )
        elif isinstance(block, Table):
            section = ensure_section()
            flush_paragraph()
            flush_list()
            table_count += 1
            section["blocks"].append({"type": "table", "html": table_to_html(block)})

    flush_paragraph()
    flush_list()
    rendered_sections = [render_section(section, index) for index, section in enumerate(sections, start=1)]
    return "\n".join(rendered_sections), toc, paragraph_count, table_count, used_images


def render_page() -> None:
    images = extract_images()
    images_by_key = {item["key"]: item for item in images}
    body_html, toc, _, _, _ = build_content(images_by_key)

    page = f"""<!doctype html>
<html lang="sr">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Kompletan katalog | Radijator Inzenjering</title>
    <link rel="icon" href="assets/favicon.svg" />
    <link rel="stylesheet" href="styles.css" />
    <link rel="stylesheet" href="catalog-premium.css" />
    <link rel="stylesheet" href="catalog-web.css" />
  </head>
  <body class="catalog-page" id="top">
    <div class="catalog-progress" aria-hidden="true"><span></span></div>
    <nav class="catalog-web-nav" aria-label="Glavna navigacija">
      <a class="web-nav-brand" href="#top" aria-label="Radijator Inženjering - vrh kataloga">
        <img src="assets/logo.png" alt="" />
        <span>Industrijski katalog</span>
      </a>
      <div class="web-nav-links">
        <a href="#section-03">Kotlovi</a>
        <a href="#section-11">Sistemi</a>
        <a href="#section-12">Oprema</a>
        <a href="#kontakt">Kontakt</a>
      </div>
      <a class="web-nav-pdf" href="radijator-industrijski-kotlovi.pdf">PDF katalog</a>
    </nav>
    <header class="catalog-hero">
      <div class="catalog-hero-topline">
        <div class="catalog-logo-card">
          <img src="assets/logo.png" alt="Radijator Inzenjering" />
        </div>
        <span>Industrijska termoenergetska rešenja</span>
      </div>
      <div class="catalog-hero-copy">
        <p class="eyebrow">Kompletan proizvodni katalog / 2026</p>
        <h1>Industrijski kotlovi <em>na biomasu</em></h1>
        <p class="catalog-lead">Pouzdani sistemi visokih snaga, projektovani za efikasnost, dug radni vek i potpunu kontrolu procesa sagorevanja.</p>
        <div class="catalog-actions">
          <a class="action-primary" href="#section-01">Pregledaj katalog</a>
          <a class="action-secondary" href="radijator-industrijski-kotlovi.pdf">Preuzmi PDF</a>
        </div>
      </div>
      <div class="hero-machine">
        <span class="machine-orbit machine-orbit--outer"></span>
        <span class="machine-orbit machine-orbit--inner"></span>
        <div class="machine-card">
          <span class="machine-series">Industrijska serija <strong>TKAN</strong></span>
          <img src="assets/full-catalog/catalog-image-13.png" alt="Industrijski kotao TKAN na biomasu" />
          <div class="machine-specs">
            <span><strong>80–500</strong> kW</span>
            <span><strong>EU</strong> standard</span>
          </div>
        </div>
      </div>
      <a class="scroll-cue" href="#section-01" aria-label="Nastavi na sadržaj"><span></span>Skrolujte</a>
    </header>
    <main class="catalog-layout">
      <details class="catalog-toc" open>
        <summary>Sadržaj kataloga</summary>
        <a class="back-link" href="#top">Vrh kataloga</a>
        <nav>{"".join(toc)}</nav>
      </details>
      <article class="catalog-content">
        {body_html}
      </article>
    </main>
    <footer class="catalog-footer" id="kontakt">
      <div class="catalog-footer-main">
        <p class="footer-kicker">Projektovanje / proizvodnja / podrška</p>
        <h2>Partner za kompletna termoenergetska rešenja.</h2>
        <a class="footer-mail" href="mailto:radijator@radijator.rs">radijator@radijator.rs</a>
      </div>
      <div class="catalog-footer-contact">
        <p><strong>Radijator Inženjering d.o.o.</strong><br />Živojina Lazića Solunca 6<br />36000 Kraljevo, Srbija</p>
        <p><a href="tel:+38136399140">+381 36 399 140</a><br /><a href="https://www.radijator.rs/">www.radijator.rs</a></p>
      </div>
      <div class="catalog-footer-bottom">
        <span>Industrijski kotlovi na biomasu</span>
        <a href="#top">Nazad na vrh</a>
      </div>
    </footer>
    <dialog class="catalog-lightbox" aria-label="Uvećani tehnički prikaz">
      <button class="lightbox-close" type="button" aria-label="Zatvori uvećani prikaz">Zatvori</button>
      <div class="lightbox-stage">
        <img alt="" />
        <p></p>
      </div>
    </dialog>
    <script src="catalog.js" defer></script>
  </body>
</html>
"""
    OUT_HTML.write_text(page, encoding="utf-8")
    ALIAS_HTML.write_text(page, encoding="utf-8")


if __name__ == "__main__":
    render_page()
