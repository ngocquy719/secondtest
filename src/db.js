const path = require('path');
const fs = require('fs');
const Database = require('better-sqlite3');
const bcrypt = require('bcryptjs');

const dbPath = path.join(__dirname, '..', 'database.sqlite');
const dbDir = path.dirname(dbPath);

if (!fs.existsSync(dbDir)) {
  fs.mkdirSync(dbDir, { recursive: true });
}

const db = new Database(dbPath);

function runAsync(sql, params = []) {
  return Promise.resolve(db.prepare(sql).run(...params));
}

function getAsync(sql, params = []) {
  return Promise.resolve(db.prepare(sql).get(...params));
}

async function initDb() {
  // Users: admin / leader / user
  await runAsync(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL CHECK (role IN ('admin', 'leader', 'user')),
      created_by INTEGER,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (created_by) REFERENCES users(id)
    )
  `);

  // Sheets
  await runAsync(`
    CREATE TABLE IF NOT EXISTS sheets (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      owner_id INTEGER NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (owner_id) REFERENCES users(id)
    )
  `);

  // Sheet permissions: owner / editor / viewer
  await runAsync(`
    CREATE TABLE IF NOT EXISTS sheet_permissions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      sheet_id INTEGER NOT NULL,
      user_id INTEGER NOT NULL,
      role TEXT NOT NULL CHECK (role IN ('owner', 'editor', 'viewer')),
      UNIQUE (sheet_id, user_id),
      FOREIGN KEY (sheet_id) REFERENCES sheets(id),
      FOREIGN KEY (user_id) REFERENCES users(id)
    )
  `);

  // Cells
  await runAsync(`
    CREATE TABLE IF NOT EXISTS cells (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      sheet_id INTEGER NOT NULL,
      row INTEGER NOT NULL,
      column INTEGER NOT NULL,
      value TEXT,
      updated_at TEXT NOT NULL,
      updated_by INTEGER NOT NULL,
      UNIQUE (sheet_id, row, column),
      FOREIGN KEY (sheet_id) REFERENCES sheets(id),
      FOREIGN KEY (updated_by) REFERENCES users(id)
    )
  `);

  // Seed default admin user
  const existingAdmin = await getAsync(
    `SELECT id FROM users WHERE role = 'admin' LIMIT 1`
  );

  if (!existingAdmin) {
    const username = 'admin';
    const password = 'admin123';
    const passwordHash = await bcrypt.hash(password, 10);

    await runAsync(
      `INSERT INTO users (username, password_hash, role, created_by)
       VALUES (?, ?, 'admin', NULL)`,
      [username, passwordHash]
    );

    console.log('Seeded default admin user:');
    console.log('  username: admin');
    console.log('  password: admin123');
  }
}

module.exports = {
  db,
  initDb,
  runAsync,
  getAsync
};

