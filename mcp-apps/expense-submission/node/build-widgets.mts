import { build } from "vite";
import { viteSingleFile } from "vite-plugin-singlefile";
import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const WIDGETS_DIR = path.join(__dirname, "widgets");
const ASSETS_DIR = path.join(__dirname, "assets");

// Find all widget directories that have an index.html
const widgetDirs = fs
  .readdirSync(WIDGETS_DIR, { withFileTypes: true })
  .filter(
    (d) =>
      d.isDirectory() &&
      fs.existsSync(path.join(WIDGETS_DIR, d.name, "index.html"))
  )
  .map((d) => d.name);

console.log(`Building ${widgetDirs.length} widget(s): ${widgetDirs.join(", ")}\n`);

if (!fs.existsSync(ASSETS_DIR)) {
  fs.mkdirSync(ASSETS_DIR, { recursive: true });
}

for (const widget of widgetDirs) {
  console.log(`  Building ${widget}...`);
  await build({
    root: path.join(WIDGETS_DIR, widget),
    plugins: [viteSingleFile()],
    build: {
      outDir: ASSETS_DIR,
      emptyOutDir: false,
      target: "esnext",
    },
    logLevel: "warn",
  });

  // vite-plugin-singlefile always outputs index.html — rename to <widget>.html
  const srcHtml = path.join(ASSETS_DIR, "index.html");
  const destHtml = path.join(ASSETS_DIR, `${widget}.html`);
  if (fs.existsSync(srcHtml)) {
    fs.renameSync(srcHtml, destHtml);
  }

  console.log(`  ✅ assets/${widget}.html\n`);
}

console.log("Done. All widgets built to assets/");
