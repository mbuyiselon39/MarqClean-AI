export function extractPdfTextInWorker(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./pdfWorker.ts", import.meta.url), { type: "module" });

    worker.onmessage = (event: MessageEvent<{ text?: string; error?: string }>) => {
      worker.terminate();
      if (event.data.error) reject(new Error(event.data.error));
      else resolve(event.data.text ?? "");
    };

    worker.onerror = (event) => {
      worker.terminate();
      reject(new Error(event.message || "PDF worker failed."));
    };

    file.arrayBuffer()
      .then((buffer) => worker.postMessage({ buffer }, [buffer]))
      .catch((error) => {
        worker.terminate();
        reject(error);
      });
  });
}
