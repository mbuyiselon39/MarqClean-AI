import { copyFile, mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";

// Ship predictable, self-hosted latin subsets from npm instead of depending on
// a third-party font API being reachable during every build and local preview.
const root = process.cwd();
const fontDir = join(root, "public", "fonts");
const fonts = [
  { file: "manrope-latin-wght-normal.woff2", family: "Manrope", package: "manrope" },
  { file: "dm-sans-latin-wght-normal.woff2", family: "DM Sans", package: "dm-sans" },
];

await mkdir(fontDir, { recursive: true });
for (const font of fonts) {
  await copyFile(join(root, "node_modules", "@fontsource-variable", font.package, "files", font.file), join(fontDir, font.file));
}
await writeFile(join(fontDir, "fonts.css"), fonts.map((font) =>
  `@font-face { font-family: "${font.family}"; font-style: normal; font-weight: 400 800; font-display: swap; src: url("/fonts/${font.file}") format("woff2"); }`
).join("\n") + "\n", "utf8");

const indexPath = join(root, "index.html");
let index = await readFile(indexPath, "utf8");
const preloads = fonts.map((font) => `    <link rel="preload" href="/fonts/${font.file}" as="font" type="font/woff2" crossorigin />`).join("\n");
index = index.replace(/<!-- FONT_PRELOADS_START -->[\s\S]*?<!-- FONT_PRELOADS_END -->/,
  `<!-- FONT_PRELOADS_START -->\n${preloads}\n    <!-- FONT_PRELOADS_END -->`);
await writeFile(indexPath, index, "utf8");
console.log("Prepared self-hosted variable fonts: Manrope and DM Sans");
