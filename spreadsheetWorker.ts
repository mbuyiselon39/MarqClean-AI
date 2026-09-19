import * as XLSX from "xlsx";

type WorkerRequest = { buffer: ArrayBuffer };

self.onmessage = (event: MessageEvent<WorkerRequest>) => {
  try {
    const workbook = XLSX.read(event.data.buffer, {
      type: "array",
      dense: true,
      cellDates: true,
      cellNF: false,
      cellStyles: false,
    });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = sheet ? XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false }) : [];
    const headers = (matrix[0] ?? []).map((value) => String(value ?? "").trim());
    const preview = matrix.slice(1, 11).map((row) => row.map((value) => String(value ?? "")));
    self.postMessage({ sheetNames: workbook.SheetNames, headers, preview });
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : "Workbook parsing failed." });
  }
};
