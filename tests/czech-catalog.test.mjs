import assert from "node:assert/strict";
import { readFileSync, existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const read = (path) => readFileSync(resolve(root, path), "utf8");
const source = read("docs/index.html");
const czech = read("docs/index-cs.html");
const translations = JSON.parse(read("scripts/catalog-cs.json"));
const count = (html, pattern) => [...html.matchAll(pattern)].length;

test("Czech catalogue retains source structure and technical values", () => {
  for (const pattern of [/<section\b/g, /<table\b/g, /<figure\b/g, /<p\b/g, /<li\b/g]) {
    assert.equal(count(czech, pattern), count(source, pattern));
  }
  assert.equal(czech, read("docs/full-catalog-cs.html"));
  const numericalTableCells = (html) => [...html.matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/g)]
    .map(([, cell]) => cell.replace(/<[^>]+>/g, "").replace(/&quot;/g, '"').trim())
    .filter((cell) => /\d/.test(cell));
  assert.deepEqual(numericalTableCells(czech), numericalTableCells(source));
  assert.match(czech, /<html lang="cs">/);
  assert.match(czech, /aria-current="page" href="index-cs.html">CS<\/a>/);
  assert.match(czech, /href="radijator-industrijski-kotlovi-cs\.pdf"/);
  assert.match(czech, /src="assets\/full-catalog\/catalog-image-07-cs\.svg"/);
  for (const phrase of Object.keys(translations)) {
    if (translations[phrase] === phrase) continue; // Valid shared terminology, e.g. "Kontakt".
    assert.ok(!czech.includes(`>${phrase}<`), `Untranslated text: ${phrase}`);
    assert.ok(!czech.includes(`alt="${phrase}"`), `Untranslated alt: ${phrase}`);
  }
});

test("Czech local assets, anchors and CSS-generated labels are complete", () => {
  for (const [, target] of czech.matchAll(/(?:href|src)="([^"]+)"/g)) {
    if (target.startsWith("#")) {
      assert.ok(czech.includes(`id="${target.slice(1)}"`), `Missing anchor: ${target}`);
    } else if (!/^(?:https?:|mailto:|tel:)/.test(target)) {
      assert.ok(existsSync(resolve(root, "docs", target.split("?")[0])), `Missing local file: ${target}`);
    }
  }
  assert.match(czech, /href="catalog-cs.css/);
  assert.match(read("docs/catalog-cs.css"), /PRŮMYSLOVÉ KOTLE NA BIOMASU/);
  assert.match(czech, /anniversary-badge-cs.svg/);
  assert.match(read("scripts/catalog-cs.json"), /"Snaga": "Výkon"/);
  const svg = read("docs/assets/full-catalog/catalog-image-07-cs.svg");
  const [, embedded] = svg.match(/href="data:image\/png;base64,([^"]+)"/);
  assert.deepEqual(Buffer.from(embedded, "base64"), readFileSync(resolve(root, "docs/assets/full-catalog/catalog-image-07.png")));
});

test("Czech is available from every existing catalogue language", () => {
  for (const language of ["", "-en", "-ro", "-bg"]) {
    assert.match(read(`docs/index${language}.html`), /href="index-cs.html">CS<\/a>/);
  }
});
