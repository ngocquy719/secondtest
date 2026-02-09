const express = require('express');
const bcrypt = require('bcryptjs');
const db = require('../db');
const { requireRole } = require('../auth');

if (!db || typeof db.get !== 'function' || typeof db.all !== 'function' || typeof db.run !== 'function') {
  throw new Error('DB instance is not initialized');
}

const router = express.Router();

// List users visible to current user
router.get('/', (req, res) => {
  const current = req.user;

  if (current.role === 'admin') {
    db.all(
      `SELECT id, username, role, created_by, created_at FROM users ORDER BY id`,
      [],
      (err, rows) => {
        if (err) {
          console.error('users list error', err);
          return res.status(500).json({ error: 'Internal server error' });
        }
        res.json({ users: rows });
      }
    );
  } else if (current.role === 'leader') {
    db.all(
      `SELECT id, username, role, created_by, created_at
       FROM users
       WHERE created_by = ?
       ORDER BY id`,
      [current.id],
      (err, rows) => {
        if (err) {
          console.error('users list error', err);
          return res.status(500).json({ error: 'Internal server error' });
        }
        res.json({ users: rows });
      }
    );
  } else {
    // Regular users don't see others per rules (no created users)
    res.json({ users: [] });
  }
});

// Create user
// - Admin: can create leaders and users
// - Leader: can create users (created_by = leader id)
router.post('/', requireRole(['admin', 'leader']), async (req, res) => {
  try {
    const current = req.user;
    const { username, password, role } = req.body;

    if (!username || !password || !role) {
      return res.status(400).json({ error: 'username, password, and role are required' });
    }

    if (!['admin', 'leader', 'user'].includes(role)) {
      return res.status(400).json({ error: 'Invalid role' });
    }

    if (current.role === 'leader' && role !== 'user') {
      return res.status(403).json({ error: 'Leaders can only create users' });
    }

    const passwordHash = await bcrypt.hash(password, 10);

    db.run(
      `INSERT INTO users (username, password_hash, role, created_by)
       VALUES (?, ?, ?, ?)`,
      [username, passwordHash, role, current.id],
      function (err) {
        if (err) {
          if (err.code === 'SQLITE_CONSTRAINT') {
            return res.status(400).json({ error: 'Username already taken' });
          }
          console.error('user create error', err);
          return res.status(500).json({ error: 'Internal server error' });
        }

        res.status(201).json({
          user: {
            id: this.lastID,
            username,
            role,
            created_by: current.id
          }
        });
      }
    );
  } catch (err) {
    console.error('user create exception', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Update user (username / password / role)
// - Admin: can edit leaders and users, but not other admins' role
// - Leader: can edit only users they created
router.put('/:id', requireRole(['admin', 'leader']), async (req, res) => {
  try {
    const current = req.user;
    const targetId = Number(req.params.id);
    const { username, password, role } = req.body;

    if (Number.isNaN(targetId)) {
      return res.status(400).json({ error: 'Invalid user id' });
    }

    // Cannot edit self here to avoid locking yourself out by mistake
    if (targetId === current.id) {
      return res.status(400).json({ error: 'Cannot edit your own account here' });
    }

    db.get(
      `SELECT id, username, role, created_by FROM users WHERE id = ?`,
      [targetId],
      async (err, user) => {
        if (err) {
          console.error('user update lookup error', err);
          return res.status(500).json({ error: 'Internal server error' });
        }
        if (!user) {
          return res.status(404).json({ error: 'User not found' });
        }

        // Permission checks
        if (current.role === 'leader') {
          // Leaders can only modify users they created and only role "user"
          if (user.created_by !== current.id || user.role !== 'user') {
            return res.status(403).json({ error: 'Forbidden' });
          }
        } else if (current.role === 'admin') {
          // Admins cannot downgrade other admins through this endpoint
          if (user.role === 'admin') {
            return res.status(403).json({ error: 'Cannot modify another admin here' });
          }
        }

        const updates = [];
        const params = [];

        if (typeof username === 'string' && username.trim()) {
          updates.push('username = ?');
          params.push(username.trim());
        }

        if (typeof role === 'string') {
          if (!['admin', 'leader', 'user'].includes(role)) {
            return res.status(400).json({ error: 'Invalid role' });
          }
          if (current.role === 'leader' && role !== 'user') {
            return res.status(403).json({ error: 'Leaders can only assign user role' });
          }
          if (current.role === 'admin' && user.role === 'admin' && role !== 'admin') {
            return res
              .status(403)
              .json({ error: 'Cannot change role of another admin here' });
          }
          updates.push('role = ?');
          params.push(role);
        }

        if (typeof password === 'string' && password.length > 0) {
          const passwordHash = await bcrypt.hash(password, 10);
          updates.push('password_hash = ?');
          params.push(passwordHash);
        }

        if (updates.length === 0) {
          return res.status(400).json({ error: 'No fields to update' });
        }

        params.push(targetId);

        db.run(
          `UPDATE users SET ${updates.join(', ')} WHERE id = ?`,
          params,
          function (err2) {
            if (err2) {
              if (err2.code === 'SQLITE_CONSTRAINT') {
                return res.status(400).json({ error: 'Username already taken' });
              }
              console.error('user update error', err2);
              return res.status(500).json({ error: 'Internal server error' });
            }
            res.json({ ok: true });
          }
        );
      }
    );
  } catch (err) {
    console.error('user update exception', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Delete user
// - Admin: can delete leaders and users (not other admins)
// - Leader: can delete only users they created
router.delete('/:id', requireRole(['admin', 'leader']), (req, res) => {
  const current = req.user;
  const targetId = Number(req.params.id);

  if (Number.isNaN(targetId)) {
    return res.status(400).json({ error: 'Invalid user id' });
  }

  if (targetId === current.id) {
    return res.status(400).json({ error: 'Cannot delete your own account' });
  }

  db.get(
    `SELECT id, role, created_by FROM users WHERE id = ?`,
    [targetId],
    (err, user) => {
      if (err) {
        console.error('user delete lookup error', err);
        return res.status(500).json({ error: 'Internal server error' });
      }
      if (!user) {
        return res.status(404).json({ error: 'User not found' });
      }

      if (current.role === 'leader') {
        if (user.created_by !== current.id || user.role !== 'user') {
          return res.status(403).json({ error: 'Forbidden' });
        }
      } else if (current.role === 'admin') {
        if (user.role === 'admin') {
          return res.status(403).json({ error: 'Cannot delete another admin' });
        }
      }

      db.run(`DELETE FROM users WHERE id = ?`, [targetId], function (err2) {
        if (err2) {
          console.error('user delete error', err2);
          return res.status(500).json({ error: 'Internal server error' });
        }
        res.json({ ok: true });
      });
    }
  );
});

module.exports = router;

