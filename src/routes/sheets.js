const express = require('express');
const db = require('../db');

if (!db || typeof db.get !== 'function' || typeof db.all !== 'function' || typeof db.run !== 'function') {
  throw new Error('DB instance is not initialized');
}

const router = express.Router();

// Build one sheet for Luckysheet multi-sheet workbook. index = 0-based position (Luckysheet uses this for setSheetActive), order = same, status = 1 active / 0 inactive.
function buildLuckysheetData(tabId, tabName, orderIndex, isActive, cellRows) {
  const celldata = (cellRows || []).map((c) => ({
    r: c.row,
    c: c.column,
    v: { v: c.value, m: c.value }
  }));
  return {
    id: tabId,
    name: tabName || 'Sheet1',
    index: orderIndex,
    order: orderIndex,
    status: isActive ? 1 : 0,
    row: 100,
    column: 26,
    celldata,
    config: {}
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

// Create a new sheet (document); creator becomes owner; one default tab "Sheet1"
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
        `INSERT INTO sheet_permissions (sheet_id, user_id, role) VALUES (?, ?, 'owner')`,
        [sheetId, current.id],
        (err2) => {
          if (err2) {
            console.error('sheet perm owner error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }
          db.run(
            `INSERT INTO sheet_tabs (sheet_id, name, order_index) VALUES (?, 'Sheet1', 0)`,
            [sheetId],
            function (err3) {
              if (err3) {
                console.error('sheet tab create error', err3);
                return res.status(500).json({ error: 'Internal server error' });
              }
              const tabId = this.lastID;
              res.status(201).json({
                sheet: { id: sheetId, name: sheetName },
                tab: { id: tabId, name: 'Sheet1', order_index: 0 }
              });
            }
          );
        }
      );
    }
  );
});

// Get sheet (document) + all tabs with cell data for Luckysheet
router.get('/:id', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  if (Number.isNaN(sheetId)) {
    return res.status(400).json({ error: 'Invalid sheet id' });
  }
  db.get(
    `SELECT s.id, s.name, sp.role AS permission
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
        `SELECT id, name, order_index FROM sheet_tabs WHERE sheet_id = ? ORDER BY order_index, id`,
        [sheetId],
        (err2, tabRows) => {
          if (err2) {
            console.error('sheet tabs error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }
          const tabs = tabRows || [];
          if (tabs.length === 0) {
            return res.json({
              sheet: { id: sheetRow.id, name: sheetRow.name },
              permission: sheetRow.permission,
              tabs: []
            });
          }
          let pending = tabs.length;
          let responded = false;
          const tabData = [];
          tabs.forEach((tab, idx) => {
            db.all(
              `SELECT row, column, value FROM cells WHERE sheet_id = ? AND sheet_tab_id = ? ORDER BY row, column`,
              [sheetId, tab.id],
              (err3, cellRows) => {
                if (err3) {
                  if (!responded) {
                    responded = true;
                    console.error('sheet get cells error', err3);
                    res.status(500).json({ error: 'Internal server error' });
                  }
                  return;
                }
                tabData[idx] = buildLuckysheetData(tab.id, tab.name, tab.order_index ?? idx, idx === 0, cellRows || []);
                pending--;
                if (pending === 0 && !responded) {
                  responded = true;
                  res.json({
                    sheet: { id: sheetRow.id, name: sheetRow.name },
                    permission: sheetRow.permission,
                    tabs: tabData
                  });
                }
              }
            );
          });
        }
      );
    }
  );
});

// Add a new tab to a sheet (document)
router.post('/:id/tabs', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const { name } = req.body;
  if (Number.isNaN(sheetId)) return res.status(400).json({ error: 'Invalid sheet id' });
  db.get(
    `SELECT sp.role FROM sheet_permissions sp WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('add tab perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm || (perm.role !== 'owner' && perm.role !== 'editor')) {
        return res.status(403).json({ error: 'No permission to add tab' });
      }
      db.get(
        `SELECT COALESCE(MAX(order_index), -1) + 1 AS next FROM sheet_tabs WHERE sheet_id = ?`,
        [sheetId],
        (e, r) => {
          if (e) {
            console.error('add tab next order error', e);
            return res.status(500).json({ error: 'Internal server error' });
          }
          const orderIndex = r ? r.next : 0;
          const tabName = (name && String(name).trim()) || 'Sheet' + (orderIndex + 1);
          db.run(
            `INSERT INTO sheet_tabs (sheet_id, name, order_index) VALUES (?, ?, ?)`,
            [sheetId, tabName, orderIndex],
            function (err2) {
              if (err2) {
                console.error('add tab insert error', err2);
                return res.status(500).json({ error: 'Internal server error' });
              }
              res.status(201).json({ id: this.lastID, name: tabName, order_index: orderIndex });
            }
          );
        }
      );
    }
  );
});

// Rename a tab
router.patch('/:id/tabs/:tabId', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const tabId = Number(req.params.tabId);
  const { name } = req.body;
  if (Number.isNaN(sheetId) || Number.isNaN(tabId) || typeof name !== 'string' || !name.trim()) {
    return res.status(400).json({ error: 'Invalid request' });
  }
  db.get(
    `SELECT sp.role FROM sheet_permissions sp WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('rename tab perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm || (perm.role !== 'owner' && perm.role !== 'editor')) {
        return res.status(403).json({ error: 'No permission to rename tab' });
      }
      db.run(
        `UPDATE sheet_tabs SET name = ? WHERE id = ? AND sheet_id = ?`,
        [name.trim(), tabId, sheetId],
        function (err2) {
          if (err2) {
            console.error('rename tab update error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }
          if (this.changes === 0) return res.status(404).json({ error: 'Tab not found' });
          res.json({ ok: true });
        }
      );
    }
  );
});

// Delete a tab (sheet_tab in this document). Must keep at least one tab.
router.delete('/:id/tabs/:tabId', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const tabId = Number(req.params.tabId);
  if (Number.isNaN(sheetId) || Number.isNaN(tabId)) {
    return res.status(400).json({ error: 'Invalid request' });
  }
  db.get(
    `SELECT sp.role FROM sheet_permissions sp WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('delete tab perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm || (perm.role !== 'owner' && perm.role !== 'editor')) {
        return res.status(403).json({ error: 'No permission to delete tab' });
      }
      db.get(
        `SELECT COUNT(*) AS cnt FROM sheet_tabs WHERE sheet_id = ?`,
        [sheetId],
        (err2, row) => {
          if (err2) {
            console.error('delete tab count error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }
          if (row && row.cnt <= 1) {
            return res.status(400).json({ error: 'Cannot delete the only sheet tab' });
          }
          db.run('DELETE FROM cells WHERE sheet_id = ? AND sheet_tab_id = ?', [sheetId, tabId], (err3) => {
            if (err3) {
              console.error('delete tab cells error', err3);
              return res.status(500).json({ error: 'Internal server error' });
            }
            db.run(
              'DELETE FROM sheet_tabs WHERE id = ? AND sheet_id = ?',
              [tabId, sheetId],
              function (err4) {
                if (err4) {
                  console.error('delete tab error', err4);
                  return res.status(500).json({ error: 'Internal server error' });
                }
                if (this.changes === 0) return res.status(404).json({ error: 'Tab not found' });
                res.json({ ok: true });
              }
            );
          });
        }
      );
    }
  );
});

// Duplicate a tab (new sheet_tab in this document with copied cells)
router.post('/:id/tabs/:tabId/duplicate', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const tabId = Number(req.params.tabId);
  if (Number.isNaN(sheetId) || Number.isNaN(tabId)) {
    return res.status(400).json({ error: 'Invalid request' });
  }
  db.get(
    `SELECT sp.role FROM sheet_permissions sp WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('duplicate tab perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm || (perm.role !== 'owner' && perm.role !== 'editor')) {
        return res.status(403).json({ error: 'No permission' });
      }
      db.get(
        `SELECT id, name FROM sheet_tabs WHERE id = ? AND sheet_id = ?`,
        [tabId, sheetId],
        (err2, tab) => {
          if (err2) {
            console.error('duplicate tab get error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }
          if (!tab) return res.status(404).json({ error: 'Tab not found' });
          db.get(
            `SELECT COALESCE(MAX(order_index), -1) + 1 AS next FROM sheet_tabs WHERE sheet_id = ?`,
            [sheetId],
            (e, r) => {
              if (e) {
                console.error('duplicate tab order error', e);
                return res.status(500).json({ error: 'Internal server error' });
              }
              const orderIndex = r ? r.next : 0;
              const newName = (tab.name || 'Sheet') + ' Copy';
              db.run(
                `INSERT INTO sheet_tabs (sheet_id, name, order_index) VALUES (?, ?, ?)`,
                [sheetId, newName, orderIndex],
                function (err3) {
                  if (err3) {
                    console.error('duplicate tab insert error', err3);
                    return res.status(500).json({ error: 'Internal server error' });
                  }
                  const newTabId = this.lastID;
                  db.all(
                    `SELECT row, column, value FROM cells WHERE sheet_id = ? AND sheet_tab_id = ?`,
                    [sheetId, tabId],
                    (err4, cells) => {
                      if (err4) {
                        console.error('duplicate tab cells error', err4);
                        return res.status(500).json({ error: 'Internal server error' });
                      }
                      const now = new Date().toISOString();
                      let pending = (cells || []).length;
                      if (pending === 0) {
                        return res.json({ id: newTabId, name: newName, order_index: orderIndex });
                      }
                      (cells || []).forEach((c) => {
                        db.run(
                          `INSERT INTO cells (sheet_id, sheet_tab_id, row, column, value, updated_at, updated_by)
                           VALUES (?, ?, ?, ?, ?, ?, ?)`,
                          [sheetId, newTabId, c.row, c.column, c.value, now, current.id],
                          (err5) => {
                            if (err5) console.error('cell copy error', err5);
                            pending--;
                            if (pending === 0) {
                              res.json({ id: newTabId, name: newName, order_index: orderIndex });
                            }
                          }
                        );
                      });
                    }
                  );
                }
              );
            }
          );
        }
      );
    }
  );
});

// Move tab order (body: { order_index: number } = new 0-based position)
router.post('/:id/tabs/:tabId/move', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const tabId = Number(req.params.tabId);
  const newIdx = Number(req.body.order_index);
  if (Number.isNaN(sheetId) || Number.isNaN(tabId) || Number.isNaN(newIdx) || newIdx < 0) {
    return res.status(400).json({ error: 'Invalid request' });
  }
  db.get(
    `SELECT sp.role FROM sheet_permissions sp WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('move tab perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm || (perm.role !== 'owner' && perm.role !== 'editor')) {
        return res.status(403).json({ error: 'No permission' });
      }
      db.all(
        `SELECT id, order_index FROM sheet_tabs WHERE sheet_id = ? ORDER BY order_index, id`,
        [sheetId],
        (err2, rows) => {
          if (err2) {
            console.error('move tab list error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }
          const tabs = rows || [];
          const fromIdx = tabs.findIndex((t) => t.id === tabId);
          if (fromIdx < 0) return res.status(404).json({ error: 'Tab not found' });
          if (fromIdx === newIdx) return res.json({ ok: true });
          const [rem] = tabs.splice(fromIdx, 1);
          tabs.splice(newIdx, 0, rem);
          let pending = tabs.length;
          const done = (e) => {
            if (e) console.error('move tab update error', e);
            pending--;
            if (pending === 0) res.json({ ok: true });
          };
          tabs.forEach((t, i) => {
            db.run(`UPDATE sheet_tabs SET order_index = ? WHERE id = ? AND sheet_id = ?`, [i, t.id, sheetId], done);
          });
        }
      );
    }
  );
});

// Get metadata for a single cell (who last updated, when)
router.get('/:id/cell-meta', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const tabId = Number(req.query.tabId);
  const row = Number(req.query.row);
  const column = Number(req.query.column);

  if (Number.isNaN(sheetId) || Number.isNaN(row) || Number.isNaN(column)) {
    return res.status(400).json({ error: 'Invalid sheet/cell coordinates' });
  }

  db.get(
    `SELECT sp.role FROM sheet_permissions sp WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('cell-meta perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm) return res.status(403).json({ error: 'No access' });

      const needTab = tabId && !Number.isNaN(tabId);
      const cellSql = needTab
        ? `SELECT c.value, c.updated_at, c.updated_by, u.username AS updated_by_username
           FROM cells c JOIN users u ON u.id = c.updated_by
           WHERE c.sheet_id = ? AND c.sheet_tab_id = ? AND c.row = ? AND c.column = ?`
        : `SELECT c.value, c.updated_at, c.updated_by, u.username AS updated_by_username
           FROM cells c JOIN users u ON u.id = c.updated_by
           WHERE c.sheet_id = ? AND c.row = ? AND c.column = ?`;
      const cellParams = needTab ? [sheetId, tabId, row, column] : [sheetId, row, column];
      db.get(cellSql, cellParams,
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

// Rename sheet (owner or editor)
router.patch('/:id', (req, res) => {
  const current = req.user;
  const sheetId = Number(req.params.id);
  const { name } = req.body;

  if (Number.isNaN(sheetId)) {
    return res.status(400).json({ error: 'Invalid sheet id' });
  }
  if (typeof name !== 'string' || !name.trim()) {
    return res.status(400).json({ error: 'Name is required' });
  }

  db.get(
    `SELECT sp.role FROM sheet_permissions sp
     WHERE sp.sheet_id = ? AND sp.user_id = ?`,
    [sheetId, current.id],
    (err, perm) => {
      if (err) {
        console.error('sheet rename perm error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!perm || (perm.role !== 'owner' && perm.role !== 'editor')) {
        return res.status(403).json({ error: 'No permission to rename' });
      }
      db.run(
        `UPDATE sheets SET name = ?, updated_at = ? WHERE id = ?`,
        [name.trim(), new Date().toISOString(), sheetId],
        (err2) => {
          if (err2) {
            console.error('sheet rename error', err2);
            return res.status(500).json({ error: 'Internal server error' });
          }
          res.json({ ok: true });
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

