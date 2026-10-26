"""Build the complete English edition without changing figures or technical values."""
from html import escape, unescape
from html.parser import HTMLParser
import json
import base64
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
TRANSLATIONS = json.loads((ROOT / "scripts/catalog-en.json").read_text(encoding="utf-8"))


class ContentInventory(HTMLParser):
    def __init__(self):
        super().__init__()
        self.texts = []

    def handle_data(self, text):
        if text.strip():
            self.texts.append(text.strip())

    def handle_starttag(self, tag, attrs):
        self.texts.extend(value for key, value in attrs if key in ("alt", "aria-label", "title") and value)


def build_english_catalog():
    source = (DOCS / "index.html").read_text(encoding="utf-8")
    inventory = ContentInventory()
    inventory.feed(source)
    # Every unchanged string must be an English description, proper name,
    # technical symbol/model, contact detail, or numerical value.
    allowed = re.compile(r"^(?:[\d\W]+|TKAN [\d/]+(?: Integra)?|DN\d+ NP\d+|[ABCEGHV]|[ABCD] \(mm\)|A1|B[12]|D[1-6]|ØD|As|Hs|EU|SR|RO|BG|EN|CS|kW|kPa|kg|mm|Language|Radijator Inzenjering|Radijator Inženjering d\.o\.o\.|Živojina Lazića Solunca 6|radijator@radijator\.rs|www\.radijator\.rs|Radijator Inženjering factory complex in Kraljevo|Automated sheet metal processing in production|Laser cutting of boiler sheet metal|Operator supervising a modern production process|Welding industrial boiler components)$")
    missing = sorted({text for text in inventory.texts if text not in TRANSLATIONS and not allowed.fullmatch(text)})
    if missing:
        raise ValueError(f"Untranslated catalogue content: {missing}")

    def translate(text):
        decoded = unescape(text)
        core = decoded.strip()
        if core not in TRANSLATIONS:
            return text
        before = decoded[:len(decoded) - len(decoded.lstrip())]
        after = decoded[len(decoded.rstrip()):]
        return before + escape(TRANSLATIONS[core], quote=True) + after

    # Preserve the original HTML structure and all numerical data exactly.
    parts = re.split(r"(<[^>]+>)", source)
    for index, part in enumerate(parts):
        if part.startswith("<"):
            parts[index] = re.sub(r'\b(alt|aria-label|title)="([^"]*)"', lambda match: f'{match[1]}="{translate(match[2])}"', part)
        else:
            parts[index] = translate(part)
    page = "".join(parts)
    page = page.replace('<html lang="sr">', '<html lang="en">')
    page = page.replace("radijator-industrijski-kotlovi.pdf", "radijator-industrijski-kotlovi-en.pdf")
    page = page.replace("assets/editorial/anniversary-badge.png", "assets/editorial/anniversary-badge-en.png")
    # Keep the original engineering drawing intact. Localise its single prose
    # annotation as native SVG text without regenerating any geometry or symbols.
    drawing = base64.b64encode((DOCS / "assets/full-catalog/catalog-image-07.png").read_bytes()).decode("ascii")
    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="1430" height="889" viewBox="0 0 1430 889">
<title>TKAN Integra dimensions. Silo for 2000 kg of pellets.</title>
<defs><mask id="annotation-mask"><rect width="1430" height="889" fill="white"/><rect x="793" y="16" width="195" height="24" fill="black"/></mask></defs>
<image width="1430" height="889" href="data:image/png;base64,{drawing}" mask="url(#annotation-mask)"/>
<text x="795" y="33" font-family="Arial, sans-serif" font-size="14" fill="black">Silo for 2000 kg of pellets</text>
</svg>'''
    (DOCS / "assets/full-catalog/catalog-image-07-en.svg").write_text(svg, encoding="utf-8")
    page = page.replace("assets/full-catalog/catalog-image-07.png", "assets/full-catalog/catalog-image-07-en.svg")
    page = page.replace('    <link rel="stylesheet" href="catalog-web.css?v=20260825-company-fit" />', '    <link rel="stylesheet" href="catalog-web.css?v=20260825-company-fit" />\n    <link rel="stylesheet" href="catalog-en.css?v=20260918" />')
    page = re.sub(r'<div class="language-switch"[^>]*>.*?</div>', language_switch("en"), page)
    page = "\n".join(line.rstrip() for line in page.splitlines()) + "\n"
    for filename in ("index-en.html", "full-catalog-en.html"):
        (DOCS / filename).write_text(page, encoding="utf-8")
    for language in ("sr", "ro", "bg"):
        suffix = "" if language == "sr" else f"-{language}"
        for filename in (f"index{suffix}.html", f"full-catalog{suffix}.html"):
            path = DOCS / filename
            original = path.read_text(encoding="utf-8")
            updated = re.sub(r'<div class="language-switch"[^>]*>.*?</div>', language_switch(language), original)
            path.write_text(updated, encoding="utf-8")
    print(f"English catalogue complete: {len(TRANSLATIONS)} translated strings, all content accounted for.")


def language_switch(current):
    links = []
    for language, label in (("sr", "SR"), ("en", "EN"), ("ro", "RO"), ("bg", "BG"), ("cs", "CS")):
        suffix = "" if language == "sr" else f"-{language}"
        active = ' class="is-active" aria-current="page"' if language == current else ' class=""'
        links.append(f'<a{active} href="index{suffix}.html">{label}</a>')
    return '<div class="language-switch" aria-label="Language">' + "".join(links) + ' </div>'


if __name__ == "__main__":
    build_english_catalog()
