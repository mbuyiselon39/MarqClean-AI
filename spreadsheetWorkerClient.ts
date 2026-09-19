export type SpreadsheetWorkerResult = {
  sheetNames: string[];
  headers: string[];
  preview: string[][];
};

export function inspectSpreadsheetInWorker(file: File): Promise<SpreadsheetWorkerResult> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./spreadsheetWorker.ts", import.meta.url), { type: "module" });
    worker.onmessage = (event: MessageEvent<SpreadsheetWorkerResult & { error?: string }>) => {
      worker.terminate();
      if (event.data.error) reject(new Error(event.data.error));
      else resolve(event.data);
    };
    worker.onerror = (event) => {
      worker.terminate();
      reject(new Error(event.message || "Spreadsheet worker failed."));
    };
    file.arrayBuffer().then((buffer) => worker.postMessage({ buffer }, [buffer])).catch((error) => {
      worker.terminate();
      reject(error);
    });
  });
}
