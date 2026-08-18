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
OUT_HTML = DOCS / "full-catalog.html"


def iter_blocks(document: Document):
    body = document.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def is_heading(text: str) -> bool:
    if not text or len(text) > 95:
        return False
    letters = [c for c in text if c.isalpha()]
    if len(letters) < 4:
        return False
    upper_ratio = sum(1 for c in letters if c.upper() == c) / len(letters)
    heading_words = (
        "model",
        "modeli",
        "sistemi",
        "oprema",
        "kotao",
        "kaskad",
        "tabela",
        "presek",
        "polozaj",
        "transport",
        "osiguranje",
        "ciklon",
        "silos",
        "automatsko",
    )
    return upper_ratio > 0.78 or any(word in text.lower() for word in heading_words)


def table_to_html(table: Table) -> str:
    rows_html: list[str] = []
    for row_index, row in enumerate(table.rows):
        cells = [clean_text(cell.text) for cell in row.cells]
        tag = "th" if row_index == 0 else "td"
        cell_html = "".join(f"<{tag}>{html.escape(cell)}</{tag}>" for cell in cells)
        rows_html.append(f"<tr>{cell_html}</tr>")
    return f"<div class=\"table-scroll\"><table>{''.join(rows_html)}</table></div>"


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
                    "display": display_path.name if display_path else "",
                    "status": status,
                }
            )

    return images


def build_content() -> tuple[str, list[str], int, int]:
    document = Document(DOCX)
    sections: list[str] = []
    toc: list[str] = []
    section_open = False
    paragraph_count = 0
    table_count = 0
    paragraph_buffer: list[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph_buffer
        if paragraph_buffer:
            sections.append(f"<p>{html.escape(' '.join(paragraph_buffer))}</p>")
            paragraph_buffer = []

    def close_section() -> None:
        nonlocal section_open
        if section_open:
            flush_paragraph()
            sections.append("</section>")
            section_open = False

    for block in iter_blocks(document):
        if isinstance(block, Paragraph):
            text = clean_text(block.text)
            if not text:
                continue
            if is_heading(text):
                close_section()
                section_id = f"section-{len(toc) + 1:02d}"
                toc.append(f"<a href=\"#{section_id}\">{html.escape(text)}</a>")
                sections.append(f"<section class=\"catalog-section\" id=\"{section_id}\">")
                sections.append(f"<h2>{html.escape(text)}</h2>")
                section_open = True
            else:
                if not section_open:
                    toc.append("<a href=\"#section-01\">Uvod</a>")
                    sections.append("<section class=\"catalog-section\" id=\"section-01\">")
                    sections.append("<h2>Uvod</h2>")
                    section_open = True
                paragraph_count += 1
                paragraph_buffer.append(text)
        elif isinstance(block, Table):
            if not section_open:
                toc.append("<a href=\"#section-01\">Uvod</a>")
                sections.append("<section class=\"catalog-section\" id=\"section-01\">")
                sections.append("<h2>Uvod</h2>")
                section_open = True
            flush_paragraph()
            table_count += 1
            sections.append(table_to_html(block))

    close_section()
    return "\n".join(sections), toc, paragraph_count, table_count


def render_page() -> None:
    images = extract_images()
    body_html, toc, paragraph_count, table_count = build_content()
    visible_images = [item for item in images if item["display"]]
    original_only = [item for item in images if not item["display"]]

    gallery = []
    for item in visible_images:
        gallery.append(
            f"""
            <figure class="gallery-card">
              <img src="assets/full-catalog/{html.escape(item['display'])}" alt="{html.escape(item['label'])} iz Word kataloga" />
              <figcaption>{html.escape(item['label'])} - {html.escape(item['original'])}</figcaption>
            </figure>
            """
        )

    if original_only:
        gallery.append(
            "<div class=\"original-only\"><h3>Originalni format u dokumentu</h3>"
            + "".join(
                f"<p>{html.escape(item['label'])}: {html.escape(item['original'])} je sacuvan u assets/full-catalog kao originalni fajl.</p>"
                for item in original_only
            )
            + "</div>"
        )

    page = f"""<!doctype html>
<html lang="sr">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Kompletan katalog | Radijator Inzenjering</title>
    <link rel="icon" href="assets/favicon.svg" />
    <link rel="stylesheet" href="styles.css" />
  </head>
  <body class="catalog-page">
    <header class="catalog-hero">
      <img src="assets/logo.png" alt="Radijator Inzenjering" />
      <div>
        <p class="eyebrow">Kompletan katalog iz Word dokumenta</p>
        <h1>Industrijski kotlovi na biomasu</h1>
        <p>U ovu verziju je prenet kompletan tekstualni materijal iz dokumenta i galerija svih slika koje su izdvojene iz DOCX fajla.</p>
        <div class="catalog-meta">
          <span>{paragraph_count} tekstualnih pasusa</span>
          <span>{table_count} tabela</span>
          <span>{len(images)} slika iz dokumenta</span>
        </div>
      </div>
    </header>
    <main class="catalog-layout">
      <aside class="catalog-toc">
        <a class="back-link" href="index.html">Nazad na prezentaciju</a>
        <h2>Sadrzaj</h2>
        <nav>{"".join(toc)}</nav>
      </aside>
      <article class="catalog-content">
        {body_html}
        <section class="catalog-section image-gallery" id="sve-slike">
          <h2>Sve slike iz dokumenta</h2>
          <div class="gallery-grid">{"".join(gallery)}</div>
        </section>
      </article>
    </main>
  </body>
</html>
"""
    OUT_HTML.write_text(page, encoding="utf-8")


if __name__ == "__main__":
    render_page()
