import * as pdfjsLib from "pdfjs-dist";

type RequestMessage = { buffer: ArrayBuffer };

type TextItemLike = {
  str?: string;
  transform?: number[];
};

self.onmessage = async (event: MessageEvent<RequestMessage>) => {
  try {
    const pdf = await pdfjsLib.getDocument({ data: event.data.buffer }).promise;
    const pages: string[] = [];

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      const items = content.items as TextItemLike[];

      const lines: Array<{ y: number; parts: Array<{ x: number; str: string }> }> = [];
      for (const item of items) {
        const str = item.str?.trim() ?? "";
        const transform = item.transform ?? [];
        if (!str || transform.length < 6) continue;

        const x = Number(transform[4]);
        const y = Number(transform[5]);
        if (!Number.isFinite(x) || !Number.isFinite(y)) continue;

        // PDF text fragments rarely share exactly the same baseline. Match the
        // nearest existing line instead of the first one, with a small tolerance.
        // Keeping the original baseline avoids cumulative drift as fragments arrive.
        let nearest: { line: { y: number; parts: Array<{ x: number; str: string }> }; distance: number } | null = null;
        for (const candidate of lines) {
          const distance = Math.abs(candidate.y - y);
          if (distance <= 4.5 && (!nearest || distance < nearest.distance)) {
            nearest = { line: candidate, distance };
          }
        }
        if (nearest) {
          nearest.line.parts.push({ x, str });
        } else {
          lines.push({ y, parts: [{ x, str }] });
        }
      }

      lines.sort((a, b) => b.y - a.y);
      pages.push(
        lines
          .map((line) => line.parts.sort((a, b) => a.x - b.x).map((part) => part.str).join(" ").replace(/\s+/g, " ").trim())
          .filter(Boolean)
          .join("\n"),
      );
      page.cleanup();
    }

    await pdf.destroy();
    self.postMessage({ text: pages.join("\n") });
  } catch (error) {
    self.postMessage({
      error: error instanceof Error ? error.message : "PDF parsing failed.",
    });
  }
};
