const path = require('path');
const sqlite3 = require('sqlite3').verbose();
const bcrypt = require('bcryptjs');

const dbPath = path.join(__dirname, '..', 'database.sqlite');

const db = new sqlite3.Database(dbPath, (err) => {
  if (err) {
    console.error('DB error:', err);
  } else {
    console.log('SQLite connected');
  }
});

function runAsync(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function (err) {
      if (err) return reject(err);
      resolve(this);
    });
  });
}

function getAsync(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => {
      if (err) return reject(err);
      resolve(row);
    });
  });
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

  // Sheet tabs (worksheets inside one document, like Google Sheets)
  await runAsync(`
    CREATE TABLE IF NOT EXISTS sheet_tabs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      sheet_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      order_index INTEGER NOT NULL DEFAULT 0,
      FOREIGN KEY (sheet_id) REFERENCES sheets(id)
    )
  `);

  const info = await new Promise((resolve, reject) => {
    db.all(`PRAGMA table_info(cells)`, (err, rows) => { if (err) reject(err); else resolve(rows || []); });
  });
  const hasCol = info.some((r) => r.name === 'sheet_tab_id');
  if (info.length === 0) {
    await runAsync(`
      CREATE TABLE cells (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sheet_id INTEGER NOT NULL,
        sheet_tab_id INTEGER NOT NULL,
        row INTEGER NOT NULL,
        column INTEGER NOT NULL,
        value TEXT,
        updated_at TEXT NOT NULL,
        updated_by INTEGER NOT NULL,
        UNIQUE (sheet_id, sheet_tab_id, row, column),
        FOREIGN KEY (sheet_id) REFERENCES sheets(id),
        FOREIGN KEY (sheet_tab_id) REFERENCES sheet_tabs(id),
        FOREIGN KEY (updated_by) REFERENCES users(id)
      )
    `);
  } else if (!hasCol) {
    const sheets = await new Promise((resolve, reject) => {
      db.all(`SELECT id FROM sheets`, (err, rows) => { if (err) reject(err); else resolve(rows); });
    });
    for (const s of sheets) {
      const tabRow = await getAsync(`SELECT id FROM sheet_tabs WHERE sheet_id = ? LIMIT 1`, [s.id]);
      const tabId = tabRow ? tabRow.id : (await runAsync(`INSERT INTO sheet_tabs (sheet_id, name, order_index) VALUES (?, 'Sheet1', 0)`, [s.id])).lastID;
    }
    await runAsync(`
      CREATE TABLE cells_new (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sheet_id INTEGER NOT NULL,
        sheet_tab_id INTEGER NOT NULL,
        row INTEGER NOT NULL,
        column INTEGER NOT NULL,
        value TEXT,
        updated_at TEXT NOT NULL,
        updated_by INTEGER NOT NULL,
        UNIQUE (sheet_id, sheet_tab_id, row, column),
        FOREIGN KEY (sheet_id) REFERENCES sheets(id),
        FOREIGN KEY (sheet_tab_id) REFERENCES sheet_tabs(id),
        FOREIGN KEY (updated_by) REFERENCES users(id)
      )
    `);
    for (const s of sheets) {
      const tabRow = await getAsync(`SELECT id FROM sheet_tabs WHERE sheet_id = ? ORDER BY order_index LIMIT 1`, [s.id]);
      const tabId = tabRow.id;
      await runAsync(
        `INSERT INTO cells_new (id, sheet_id, sheet_tab_id, row, column, value, updated_at, updated_by)
         SELECT id, sheet_id, ?, row, column, value, updated_at, updated_by FROM cells WHERE sheet_id = ?`,
        [tabId, s.id]
      );
    }
    await runAsync(`DROP TABLE cells`);
    await runAsync(`ALTER TABLE cells_new RENAME TO cells`);
  }

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

module.exports = db;
module.exports.initDb = initDb;
