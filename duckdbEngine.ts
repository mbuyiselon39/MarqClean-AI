import * as duckdb from "@duckdb/duckdb-wasm";
import duckdbMvp from "@duckdb/duckdb-wasm/dist/duckdb-mvp.wasm?url";
import mvpWorker from "@duckdb/duckdb-wasm/dist/duckdb-browser-mvp.worker.js?url";
import duckdbEh from "@duckdb/duckdb-wasm/dist/duckdb-eh.wasm?url";
import ehWorker from "@duckdb/duckdb-wasm/dist/duckdb-browser-eh.worker.js?url";

let databasePromise: Promise<duckdb.AsyncDuckDB> | null = null;

async function getDatabase() {
  if (!databasePromise) {
    databasePromise = (async () => {
      const bundles: duckdb.DuckDBBundles = {
        mvp: { mainModule: duckdbMvp, mainWorker: mvpWorker },
        eh: { mainModule: duckdbEh, mainWorker: ehWorker },
      };
      const bundle = await duckdb.selectBundle(bundles);
      const worker = new Worker(bundle.mainWorker!);
      const db = new duckdb.AsyncDuckDB(new duckdb.ConsoleLogger(), worker);
      await db.instantiate(bundle.mainModule, bundle.pthreadWorker);
      return db;
    })();
  }
  return databasePromise;
}

function quoteIdentifier(value: string) {
  return '"' + value.replace(/"/g, '""') + '"';
}

function quotePath(value: string) {
  return "'" + value.replace(/'/g, "''") + "'";
}

export async function queryCsv(file: File, sql: string) {
  const db = await getDatabase();
  const conn = await db.connect();
  const fileName = `marq_${crypto.randomUUID().replace(/-/g, "")}.csv`;
  try {
    await db.registerFileBuffer(fileName, new Uint8Array(await file.arrayBuffer()));
    const source = quotePath(fileName);
    const normalizedSql = sql.replace(/__CSV__/g, source);
    const result = await conn.query(normalizedSql);
    return result.toArray().map((row) => Object.fromEntries(Object.entries(row)));
  } finally {
    await conn.close();
    try { await db.dropFile(fileName); } catch { /* best effort */ }
  }
}

export function safeIdentifier(value: string) {
  return quoteIdentifier(value);
}
