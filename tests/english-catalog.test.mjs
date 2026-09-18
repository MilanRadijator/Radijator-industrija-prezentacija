import assert from "node:assert/strict";
import { readFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const read = (path) => readFileSync(resolve(root, path), "utf8");
const source = read("docs/index.html");
const english = read("docs/index-en.html");
const translations = JSON.parse(read("scripts/catalog-en.json"));
const tables = (html) => [...html.matchAll(/<table\b[^>]*>[\s\S]*?<\/table>/g)].map(([table]) => table);
const numericCells = (table) => [...table.matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/g)]
  .map(([, cell]) => cell.replace(/<[^>]+>/g, "").replace(/&quot;/g, '"'))
  .filter((cell) => /\d/.test(cell));

test("English catalogue retains all sections, tables, figures and technical values", () => {
  for (const pattern of [/<section\b/g, /<table\b/g, /<figure\b/g, /<p\b/g, /<li\b/g]) {
    assert.equal([...english.matchAll(pattern)].length, [...source.matchAll(pattern)].length);
  }
  for (let section = 1; section <= 13; section++) {
    assert.match(english, new RegExp(`id="section-${String(section).padStart(2, "0")}"`));
  }
  assert.deepEqual(tables(english).map(numericCells), tables(source).map(numericCells));
  assert.equal(english, read("docs/full-catalog-en.html"));
  for (const phrase of Object.keys(translations)) {
    // Check actual text and accessibility strings, excluding source code/URLs.
    assert.ok(!english.includes(`>${phrase}<`), `Untranslated text: ${phrase}`);
    assert.ok(!english.includes(`alt="${phrase}"`), `Untranslated alt: ${phrase}`);
  }
});

test("English navigation, downloadable PDF, images and local links are complete", () => {
  assert.match(english, /<html lang="en">/);
  assert.match(english, /aria-current="page" href="index-en.html">EN<\/a>/);
  assert.match(english, /href="radijator-industrijski-kotlovi-en.pdf"/);
  for (const [, target] of english.matchAll(/(?:href|src)="([^"]+)"/g)) {
    if (target.startsWith("#")) {
      assert.ok(english.includes(`id="${target.slice(1)}"`), `Missing anchor: ${target}`);
    } else if (!/^(?:https?:|mailto:|tel:)/.test(target)) {
      assert.ok(existsSync(resolve(root, "docs", target.split("?")[0])), `Missing local file: ${target}`);
    }
  }
  for (const language of ["", "-ro", "-bg"]) {
    assert.match(read(`docs/index${language}.html`), /href="index-en.html">EN<\/a>/);
  }
});

test("Localised engineering annotation retains the exact original drawing", () => {
  const svg = read("docs/assets/full-catalog/catalog-image-07-en.svg");
  const [, embedded] = svg.match(/href="data:image\/png;base64,([^"]+)"/);
  assert.deepEqual(Buffer.from(embedded, "base64"), readFileSync(resolve(root, "docs/assets/full-catalog/catalog-image-07.png")));
  assert.match(svg, /Silo for 2000 kg of pellets/);
  assert.match(english, /src="assets\/full-catalog\/catalog-image-07-en.svg"/);
});
