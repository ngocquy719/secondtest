require('dotenv').config();

const path = require('path');
const http = require('http');
const express = require('express');
const cors = require('cors');
const { Server } = require('socket.io');

const { db, initDb } = require('./src/db');
const { authMiddleware, socketAuthMiddleware, generateToken } = require('./src/auth');
const userRoutes = require('./src/routes/users');
const sheetRoutes = require('./src/routes/sheets');

const app = express();
const server = http.createServer(app);

const io = new Server(server, {
  cors: {
    origin: '*',
    methods: ['GET', 'POST']
  }
});

// Basic middleware
app.use(cors());
app.use(express.json());

// Static frontend
app.use(express.static(path.join(__dirname, 'public')));

// Health check
app.get('/api/health', (req, res) => {
  res.json({ ok: true });
});

// Auth routes
app.post('/api/auth/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    if (!username || !password) {
      return res.status(400).json({ error: 'Username and password required' });
    }

    db.get(
      'SELECT id, username, password_hash AS passwordHash, role FROM users WHERE username = ?',
      [username],
      async (err, user) => {
        if (err) {
          console.error('Login error:', err);
          return res.status(500).json({ error: 'Internal server error' });
        }
        if (!user) {
          return res.status(401).json({ error: 'Invalid credentials' });
        }

        const bcrypt = require('bcryptjs');
        const match = await bcrypt.compare(password, user.passwordHash);
        if (!match) {
          return res.status(401).json({ error: 'Invalid credentials' });
        }

        const token = generateToken({ id: user.id, username: user.username, role: user.role });
        res.json({
          token,
          user: { id: user.id, username: user.username, role: user.role }
        });
      }
    );
  } catch (e) {
    console.error('Login exception:', e);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Current user
app.get('/api/auth/me', authMiddleware, (req, res) => {
  res.json({ user: req.user });
});

// User & sheet routes
app.use('/api/users', authMiddleware, userRoutes);
app.use('/api/sheets', authMiddleware, sheetRoutes);

// Socket.IO for realtime cell updates
io.use(socketAuthMiddleware);

io.on('connection', (socket) => {
  const user = socket.user;
  console.log('Socket connected:', user.username, socket.id);

  // Join a sheet room
  socket.on('join_sheet', ({ sheetId }) => {
    if (!sheetId) return;

    db.get(
      `SELECT sp.role
       FROM sheet_permissions sp
       WHERE sp.sheet_id = ? AND sp.user_id = ?`,
      [sheetId, user.id],
      (err, row) => {
        if (err) {
          console.error('join_sheet error', err);
          return;
        }
        if (!row) {
          // No access
          return;
        }
        const room = `sheet_${sheetId}`;
        socket.join(room);
      }
    );
  });

  // Presence / cursor location within a sheet
  socket.on('presence', (payload) => {
    const { sheetId, row, column } = payload || {};
    if (typeof sheetId !== 'number') return;
    const room = `sheet_${sheetId}`;
    // Broadcast to others in the same sheet
    socket.to(room).emit('presence', {
      sheetId,
      row,
      column,
      userId: user.id,
      username: user.username
    });
  });

  // Cell update handling
  socket.on('cell_update', (payload) => {
    const { sheetId, row, column, value, timestamp } = payload || {};
    if (
      typeof sheetId !== 'number' ||
      typeof row !== 'number' ||
      typeof column !== 'number'
    ) {
      return;
    }

    // Check permissions (owner or editor)
    db.get(
      `SELECT sp.role
       FROM sheet_permissions sp
       WHERE sp.sheet_id = ? AND sp.user_id = ?`,
      [sheetId, user.id],
      (err, perm) => {
        if (err) {
          console.error('cell_update perm error', err);
          return;
        }
        if (!perm || (perm.role !== 'owner' && perm.role !== 'editor')) {
          // Read-only or no access
          return;
        }

        const now = new Date().toISOString();

        // Upsert cell
        db.get(
          `SELECT id FROM cells WHERE sheet_id = ? AND row = ? AND column = ?`,
          [sheetId, row, column],
          (err2, cell) => {
            if (err2) {
              console.error('cell select error', err2);
              return;
            }

            const userId = user.id;

            if (!value || value === '') {
              // If empty, delete cell if exists
              if (cell) {
                db.run('DELETE FROM cells WHERE id = ?', [cell.id], (err3) => {
                  if (err3) console.error('cell delete error', err3);
                });
              }
            } else if (cell) {
              db.run(
                `UPDATE cells
                 SET value = ?, updated_at = ?, updated_by = ?
                 WHERE id = ?`,
                [String(value), now, userId, cell.id],
                (err3) => {
                  if (err3) console.error('cell update error', err3);
                }
              );
            } else {
              db.run(
                `INSERT INTO cells (sheet_id, row, column, value, updated_at, updated_by)
                 VALUES (?, ?, ?, ?, ?, ?)`,
                [sheetId, row, column, String(value), now, userId],
                (err3) => {
                  if (err3) console.error('cell insert error', err3);
                }
              );
            }

            const room = `sheet_${sheetId}`;
            const out = {
              sheetId,
              row,
              column,
              value,
              userId,
              timestamp: timestamp || now
            };
            io.to(room).emit('cell_update', out);
          }
        );
      }
    );
  });

  socket.on('disconnect', () => {
    console.log('Socket disconnected:', user.username, socket.id);
    // Let others know this user went offline / left sheet
    io.emit('presence_leave', {
      userId: user.id,
      username: user.username
    });
  });
});

const PORT = process.env.PORT || 3000;

initDb()
  .then(() => {
    server.listen(PORT, () => {
      console.log(`Server listening on http://localhost:${PORT}`);
    });
  })
  .catch((err) => {
    console.error('Failed to initialize database', err);
    process.exit(1);
  });

