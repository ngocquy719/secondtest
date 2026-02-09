const express = require('express');
const { db } = require('../db');

const router = express.Router();

// Helper: build luckysheet sheet object from DB rows
function buildLuckysheetData(sheetRow, cellRows) {
  const celldata = cellRows.map((c) => ({
    r: c.row,
    c: c.column,
    v: {
      v: c.value,
      m: c.value
    }
  }));

  return {
    id: sheetRow.id,
    name: sheetRow.name || 'Sheet1',
    index: 0,
    row: 100,
    column: 26,
    celldata
  };
}

// List sheets current user can access
router.get('/', (req, res) => {
  const current = req.user;

  db.all(
    `SELECT s.id, s.name, s.owner_id, s.created_at, s.updated_at, sp.role AS permission
     FROM sheets s
     JOIN sheet_permissions sp ON sp.sheet_id = s.id
     WHERE sp.user_id = ?
     ORDER BY s.id`,
    [current.id],
    (err, rows) => {
      if (err) {
        console.error('sheet list error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      res.json({ sheets: rows });
    }
  );
});

// Create a new sheet; creator becomes owner
router.post('/', (req, res) => {
  const current = req.user;
  const { name } = req.body;
  const sheetName = name || 'Sheet1';

  const now = new Date().toISOString();

  db.run(
    `INSERT INTO sheets (name, owner_id, created_at, updated_at)
     VALUES (?, ?, ?, ?)`,
    [sheetName, current.id, now, now],
    function (err) {
      if (err) {
        console.error('sheet create error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }

      const sheetId = this.lastID;

      db.run(
        `INSERT INTO sheet_permissions (sheet_id, user_id, role)
         VALUES (?, ?, 'owner')`,
        [sheetId, current.id],
        (err2) => {
          if (err2) {
            console.error('sheet perm owner error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }

          // New sheet must be initialized with an empty celldata array
          const sheetObj = {
            id: sheetId,
            name: 'Sheet1',
            index: 0,
            row: 100,
            column: 26,
            celldata: []
          };

          res.status(201).json({ sheet: sheetObj });
        }
      );
    }
  );
});

// Get sheet data for Luckysheet (single fetch per sheet on frontend)
router.get('/:id', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);

  if (Number.isNaN(sheetId)) {
    return res.status(400).json({ error: 'Invalid sheet id' });
  }

  // Check permission
  db.get(
    `SELECT s.id, s.name, s.owner_id, s.created_at, s.updated_at, sp.role AS permission
     FROM sheets s
     JOIN sheet_permissions sp ON sp.sheet_id = s.id
     WHERE s.id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, sheetRow) => {
      if (err) {
        console.error('sheet get error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!sheetRow) {
        return res.status(404).json({ error: 'Sheet not found' });
      }

      db.all(
        `SELECT id, row, column, value, updated_at, updated_by
         FROM cells
         WHERE sheet_id = ?
         ORDER BY row, column`,
        [sheetId],
        (err2, cellRows) => {
          if (err2) {
            console.error('sheet cells error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }

          const sheetObj = buildLuckysheetData(sheetRow, cellRows);
          res.json({ sheet: sheetObj, permission: sheetRow.permission });
        }
      );
    }
  );
});

// Get metadata for a single cell (who last updated, when)
router.get('/:id/cell-meta', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const row = Number(req.query.row);
  const column = Number(req.query.column);

  if (Number.isNaN(sheetId) || Number.isNaN(row) || Number.isNaN(column)) {
    return res.status(400).json({ error: 'Invalid sheet/cell coordinates' });
  }

  // Check permission (viewer/editor/owner all can see metadata)
  db.get(
    `SELECT sp.role
     FROM sheet_permissions sp
     WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('cell-meta perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm) {
        return res.status(403).json({ error: 'No access' });
      }

      db.get(
        `SELECT c.value, c.updated_at, c.updated_by, u.username AS updated_by_username
         FROM cells c
         JOIN users u ON u.id = c.updated_by
         WHERE c.sheet_id = ? AND c.row = ? AND c.column = ?`,
        [sheetId, row, column],
        (err2, cell) => {
          if (err2) {
            console.error('cell-meta lookup error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }

          if (!cell) {
            return res.json({
              row,
              column,
              value: null,
              updatedAt: null,
              updatedBy: null
            });
          }

          res.json({
            row,
            column,
            value: cell.value,
            updatedAt: cell.updated_at,
            updatedBy: {
              id: cell.updated_by,
              username: cell.updated_by_username
            }
          });
        }
      );
    }
  );
});

// Share sheet with another user (editor / viewer)
router.post('/:id/share', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const { userId, role } = req.body;

  if (Number.isNaN(sheetId) || !userId || !role) {
    return res
      .status(400)
      .json({ error: 'sheetId, userId and role (editor/viewer) are required' });
  }

  if (!['editor', 'viewer'].includes(role)) {
    return res.status(400).json({ error: 'Invalid share role' });
  }

  // Only owner can share the sheet
  db.get(
    `SELECT sp.role
     FROM sheet_permissions sp
     WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('share perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm || perm.role !== 'owner') {
        return res.status(403).json({ error: 'Only owner can share sheet' });
      }

      db.run(
        `INSERT INTO sheet_permissions (sheet_id, user_id, role)
         VALUES (?, ?, ?)
         ON CONFLICT(sheet_id, user_id) DO UPDATE SET role = excluded.role`,
        [sheetId, userId, role],
        (err2) => {
          if (err2) {
            console.error('share insert error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }
          res.json({ ok: true });
        }
      );
    }
  );
});

module.exports = router;

