import * as XLSX from "xlsx";

type WorkerRequest = { buffer: ArrayBuffer };

self.onmessage = (event: MessageEvent<WorkerRequest>) => {
  try {
    const workbook = XLSX.read(event.data.buffer, {
      type: "array",
      dense: true,
      cellDates: true,
      cellNF: true,
      cellStyles: false,
      raw: true,
    });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = sheet ? XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true }) : [];
    const isDateFormat = (format: unknown) => {
      if (typeof format !== "string" || !format) return false;
      const cleaned = format.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "").replace(/\\./g, "");
      return /(^|[^a-z])(?:d{1,4}|m{1,4}|y{2,4})(?:[^a-z]|$)/i.test(cleaned);
    };
    const toDate = (serial: number) => {
      const date = new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
      return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
    };
    const normalizedMatrix = matrix.map((row, rowIndex) => row.map((value, columnIndex) => {
      if (value instanceof Date) return toDate(value.getTime() / 86400000 + 25569);
      if (typeof value === "number") {
        const cell = sheet?.[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
        if (cell && isDateFormat(cell.z)) return toDate(value);
      }
      return String(value ?? "");
    }));
    const headers = (normalizedMatrix[0] ?? []).map((value) => String(value ?? "").trim());
    const preview = normalizedMatrix.slice(1, 11).map((row) => row.map((value) => String(value ?? "")));
    self.postMessage({ sheetNames: workbook.SheetNames, headers, preview });
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : "Workbook parsing failed." });
  }
};
