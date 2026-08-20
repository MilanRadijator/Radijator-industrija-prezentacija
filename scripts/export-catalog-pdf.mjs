import { copyFileSync, createReadStream, existsSync, mkdtempSync, rmSync, statSync, unlinkSync } from "node:fs";
import { createServer } from "node:http";
import { tmpdir } from "node:os";
import { extname, join, normalize, resolve } from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const docs = join(root, "docs");
const output = join(docs, "radijator-industrijski-kotlovi.pdf");
const printOutput = join(tmpdir(), `radijator-industrijski-kotlovi-${process.pid}.pdf`);
const userDataDir = mkdtempSync(join(tmpdir(), "radijator-edge-"));

const edgeCandidates = [
  process.env.CHROME_PATH,
  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
  process.env.MSEDGE_PATH,
  "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
  "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
].filter(Boolean);
const edge = edgeCandidates.find(existsSync);

if (!edge) {
  throw new Error("Microsoft Edge nije pronađen. Postavite MSEDGE_PATH i pokušajte ponovo.");
}

const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".jpeg": "image/jpeg",
  ".jpg": "image/jpeg",
  ".js": "text/javascript; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

const server = createServer((request, response) => {
  const requested = decodeURIComponent(new URL(request.url, "http://127.0.0.1").pathname);
  const relative = requested === "/" ? "index.html" : requested.slice(1);
  const target = normalize(join(docs, relative));

  if (!target.startsWith(docs) || !existsSync(target) || !statSync(target).isFile()) {
    response.writeHead(404).end("Not found");
    return;
  }

  response.writeHead(200, {
    "Cache-Control": "no-store",
    "Content-Type": contentTypes[extname(target).toLowerCase()] || "application/octet-stream",
  });
  createReadStream(target).pipe(response);
});

await new Promise((resolveListen) => server.listen(0, "127.0.0.1", resolveListen));
const address = server.address();
const url = `http://127.0.0.1:${address.port}/index.html`;

const args = [
  "--headless=new",
  "--disable-gpu",
  "--disable-crash-reporter",
  "--hide-scrollbars",
  "--no-pdf-header-footer",
  "--run-all-compositor-stages-before-draw",
  `--user-data-dir=${userDataDir}`,
  `--print-to-pdf=${printOutput}`,
  url,
];

if (existsSync(printOutput)) {
  unlinkSync(printOutput);
}

const exitCode = await new Promise((resolveExit, reject) => {
  const child = spawn(edge, args, { stdio: ["ignore", "inherit", "inherit"] });
  child.once("error", reject);
  child.once("exit", resolveExit);
});

await new Promise((resolveClose) => server.close(resolveClose));

if (exitCode !== 0 || !existsSync(printOutput) || statSync(printOutput).size === 0) {
  rmSync(userDataDir, { recursive: true, force: true });
  throw new Error(`PDF izvoz nije uspeo. Edge exit code: ${exitCode}`);
}

copyFileSync(printOutput, output);
unlinkSync(printOutput);
rmSync(userDataDir, { recursive: true, force: true });

console.log(`PDF exported: ${output}`);
