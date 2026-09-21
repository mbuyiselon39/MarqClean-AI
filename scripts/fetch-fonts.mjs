import { mkdir, writeFile, readFile } from "node:fs/promises";
import { join, basename } from "node:path";

const cssUrl = "https://api.fontshare.com/v2/css?f[]=clash-grotesk@700&f[]=satoshi@400,500,700&display=swap";
const fontDir = join(process.cwd(), "public", "fonts");
const indexPath = join(process.cwd(), "index.html");

const response = await fetch(cssUrl);
if (!response.ok) throw new Error(`Fontshare stylesheet request failed: ${response.status}`);
const sourceCss = await response.text();
const urls = [...new Set([...sourceCss.matchAll(/url\((['"]?)(https?:\/\/[^'")]+)\1\)/g)].map((match) => match[2]))];
if (!urls.length) throw new Error("Fontshare returned no WOFF2 assets.");

await mkdir(fontDir, { recursive: true });
const localNames = [];

for (const url of urls) {
  const name = basename(new URL(url).pathname);
  if (!name.endsWith(".woff2")) continue;
  const asset = await fetch(url);
  if (!asset.ok) throw new Error(`Font asset request failed for ${name}: ${asset.status}`);
  await writeFile(join(fontDir, name), Buffer.from(await asset.arrayBuffer()));
  localNames.push(name);
}

const localCss = sourceCss.replace(/url\((['"]?)https?:\/\/[^'")]+\1\)/g, (_, quote, url) => {
  const name = basename(new URL(url).pathname);
  return `url("/fonts/${name}")`;
});
await writeFile(join(fontDir, "fonts.css"), localCss, "utf8");

let index = await Bun.file(indexPath).text().catch(async () => {
  const { readFile } = await import("node:fs/promises");
  return readFile(indexPath, "utf8");
});
const preloads = localNames.map((name) => `    <link rel="preload" href="/fonts/${name}" as="font" type="font/woff2" crossorigin />`).join("\n");
index = index.replace(/<!-- FONT_PRELOADS_START -->[\s\S]*?<!-- FONT_PRELOADS_END -->/, `<!-- FONT_PRELOADS_START -->\n${preloads}\n    <!-- FONT_PRELOADS_END -->`);
await writeFile(indexPath, index, "utf8");

console.log(`Self-hosted ${localNames.length} Fontshare WOFF2 files: ${localNames.join(", ")}`);
