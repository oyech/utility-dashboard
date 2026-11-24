const path = require('path');
const fs = require('fs');

function findDbPath() {
  const candidates = [
    path.join(__dirname, 'utility.db'),
    path.join(process.cwd(), 'utility.db'),
    path.join(__dirname, '..', 'utility.db')
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return path.join(__dirname, 'utility.db');
}

let dbPromise = null;

async function initDb() {
  const dbPath = findDbPath();

  // 1. Try built-in node:sqlite
  try {
    const { DatabaseSync } = require('node:sqlite');
    const db = new DatabaseSync(dbPath);
    return {
      all(sql, params = []) {
        const stmt = db.prepare(sql);
        return stmt.all(...params);
      },
      run(sql, params = []) {
        return db.prepare(sql).run(...params);
      },
      exec(sql) {
        return db.exec(sql);
      }
    };
  } catch (err) {
    console.warn('[DB] node:sqlite not available, using sql.js WebAssembly:', err.message);
  }

  // 2. Fallback to sql.js (Pure WebAssembly SQLite - 100% universal across all Node/Vercel versions)
  const initSqlJs = require('sql.js');
  const SQL = await initSqlJs();
  const filebuffer = fs.readFileSync(dbPath);
  const db = new SQL.Database(filebuffer);

  return {
    all(sql, params = []) {
      const stmt = db.prepare(sql);
      stmt.bind(params);
      const rows = [];
      while (stmt.step()) {
        rows.push(stmt.getAsObject());
      }
      stmt.free();
      return rows;
    },
    run(sql, params = []) {
      db.run(sql, params);
      return { changes: 1 };
    },
    exec(sql) {
      db.run(sql);
    }
  };
}

function getDb() {
  if (!dbPromise) {
    dbPromise = initDb();
  }
  return dbPromise;
}

module.exports = {
  getDb,
  findDbPath
};
