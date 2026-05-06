import { Database } from 'bun:sqlite';

import migration001 from './migrations/001_initial.sql' with { type: 'text' };
import migration002 from './migrations/002_provider_mcp.sql' with { type: 'text' };

const MIGRATIONS: Array<[number, string]> = [
  [1, migration001],
  [2, migration002],
];

export type DbConfig = { url: string };

export function openDb(cfg: DbConfig): Database {
  const db = new Database(cfg.url, { create: true, strict: true });
  db.exec('PRAGMA journal_mode = WAL');
  db.exec('PRAGMA foreign_keys = ON');

  db.exec(`
    CREATE TABLE IF NOT EXISTS schema_migrations (
      version INTEGER PRIMARY KEY,
      applied_at INTEGER NOT NULL
    )
  `);

  const applied = new Set<number>(
    (db.query('SELECT version FROM schema_migrations').all() as Array<{ version: number }>).map(
      (r) => r.version,
    ),
  );

  const insertStmt = db.prepare(
    'INSERT INTO schema_migrations (version, applied_at) VALUES (?, ?)',
  );

  for (const [version, sql] of MIGRATIONS) {
    if (applied.has(version)) continue;
    db.exec('BEGIN');
    try {
      db.exec(sql);
      insertStmt.run(version, Date.now());
      db.exec('COMMIT');
    } catch (err) {
      db.exec('ROLLBACK');
      throw new Error(`Migration ${version} failed: ${(err as Error).message}`);
    }
  }

  return db;
}
