import * as duckdb from "@duckdb/duckdb-wasm";

let databasePromise: Promise<duckdb.AsyncDuckDB> | null = null;

async function getDatabase() {
  if (!databasePromise) {
    databasePromise = (async () => {
      // Keep the static app under Cloudflare Pages' 25 MiB asset limit.
      // DuckDB's official CDN bundles are selected at runtime, while the application
      // remains fully serverless and user data still stays in the browser.
      const bundle = await duckdb.selectBundle(duckdb.getJsDelivrBundles());
      const workerUrl = URL.createObjectURL(
        new Blob([`importScripts("${bundle.mainWorker!}");`], { type: "text/javascript" }),
      );
      const worker = new Worker(workerUrl);
      const db = new duckdb.AsyncDuckDB(new duckdb.ConsoleLogger(), worker);
      await db.instantiate(bundle.mainModule, bundle.pthreadWorker);
      URL.revokeObjectURL(workerUrl);
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
