(() => {
  if (!window.auth.requireAuth()) return;

  const apiBase = '/api';
  let authToken = window.auth.getToken();
  let currentUser = null;
  let currentPermission = null;
  let socket = null;
  let luckysheetInitialized = false;
  let isApplyingRemoteUpdate = false;
  const lastCellValues = {};
  let knownUsers = [];

  const userInfoEl = document.getElementById('user-info');
  const sheetTitleEl = document.getElementById('sheet-title');
  const sheetPermissionEl = document.getElementById('sheet-permission');
  const shareForm = document.getElementById('share-form');
  const shareUserIdInput = document.getElementById('share-user-id');
  const shareRoleSelect = document.getElementById('share-role');
  const shareMessage = document.getElementById('share-message');
  const cellInfoForm = document.getElementById('cell-info-form');
  const cellRowInput = document.getElementById('cell-row');
  const cellColumnInput = document.getElementById('cell-column');
  const cellInfoMessage = document.getElementById('cell-info-message');
  const cellPresenceMessage = document.getElementById('cell-presence-message');
  const onlineUsersEl = document.getElementById('online-users');
  const remotePresenceByUser = {};
  const toggleCellInfoBtn = document.getElementById('toggle-cell-info');
  const toggleShareBtn = document.getElementById('toggle-share');
  const cellInfoSection = document.getElementById('cell-info-section');
  const shareSection = document.getElementById('share-section');

  const sheetId = (() => {
    const m = /[?&]id=(\d+)/.exec(window.location.search);
    return m ? Number(m[1]) : null;
  })();

  if (!sheetId) {
    sheetTitleEl.textContent = 'Invalid sheet (missing id)';
    shareForm.classList.add('hidden');
    return;
  }

  async function apiRequest(path, options = {}) {
    const headers = options.headers || {};
    if (authToken) {
      headers['Authorization'] = `Bearer ${authToken}`;
    }
    if (!(options.body instanceof FormData)) {
      headers['Content-Type'] = 'application/json';
    }

    const res = await fetch(`${apiBase}${path}`, {
      ...options,
      headers,
      body:
        options.body && !(options.body instanceof FormData)
          ? JSON.stringify(options.body)
          : options.body
    });

    if (!res.ok) {
      let msg = 'Request failed';
      try {
        const data = await res.json();
        if (data && data.error) msg = data.error;
      } catch (_) {}
      throw new Error(msg);
    }
    return res.json();
  }

  async function loadCurrentUser() {
    const data = await apiRequest('/auth/me');
    currentUser = data.user;
    userInfoEl.textContent = `${currentUser.username} (${currentUser.role})`;
  }

  async function loadUsersForShare() {
    try {
      const data = await apiRequest('/users');
      knownUsers = data.users || [];
      shareUserIdInput.innerHTML = '';
      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = 'Select user';
      shareUserIdInput.appendChild(placeholder);

      knownUsers.forEach((u) => {
        if (currentUser && u.id === currentUser.id) return;
        const opt = document.createElement('option');
        opt.value = String(u.id);
        opt.textContent = `${u.username} (id:${u.id})`;
        shareUserIdInput.appendChild(opt);
      });
    } catch (err) {
      console.error('Failed to load users for share', err);
    }
  }

  async function loadSheet() {
    const data = await apiRequest(`/sheets/${sheetId}`);
    const sheetObj = data.sheet;
    currentPermission = data.permission;

    sheetTitleEl.textContent = sheetObj.name || 'Sheet';
    sheetPermissionEl.textContent = currentPermission;

    if (!luckysheetInitialized) {
      luckysheet.create({
        container: 'luckysheet',
        data: [sheetObj],
        showinfobar: false,
        enableAddRow: true,
        enableAddCol: true,
        sheetFormulaBar: true,
        hook: {
          // Commit event (Enter / blur / move), giống Google Sheets hơn
          cellUpdated: (r, c, oldValue, newValue) => {
            handleLocalCellChange(r, c, newValue);
          },
          // Khi chọn cell: gửi presence + tự load meta cho ô đó
          cellSelected: (r, c) => {
            emitPresence(r, c);
            loadCellMeta(r, c);
          }
        }
      });
      luckysheetInitialized = true;
    }
  }

  function connectSocket() {
    if (socket) {
      socket.disconnect();
      socket = null;
    }
    socket = io('', {
      auth: {
        token: authToken
      }
    });

    socket.on('connect', () => {
      socket.emit('join_sheet', { sheetId });
      // Đánh dấu mình online trong sheet này
      socket.emit('presence', { sheetId, row: null, column: null });
    });

    socket.on('cell_update', (payload) => {
      const { sheetId: eventSheetId, row, column, value } = payload || {};
      if (!luckysheetInitialized) return;
      if (eventSheetId !== sheetId) return;

      isApplyingRemoteUpdate = true;
      try {
        luckysheet.setCellValue(row, column, value);
        const key = `${row}:${column}`;
        lastCellValues[key] = value;
      } finally {
        isApplyingRemoteUpdate = false;
      }
    });

    socket.on('presence', (payload) => {
      const { userId, username, row, column } = payload || {};
      if (!userId || !username) return;
      if (currentUser && userId === currentUser.id) return;
      remotePresenceByUser[userId] = { username, row, column };
      renderOnlineUsers();
      updateCellPresenceMessage();
    });

    socket.on('presence_leave', (payload) => {
      const { userId } = payload || {};
      if (!userId) return;
      delete remotePresenceByUser[userId];
      renderOnlineUsers();
      updateCellPresenceMessage();
    });
  }

  function handleLocalCellChange(r, c, rawValue) {
    if (!socket || !socket.connected) return;
    if (isApplyingRemoteUpdate) return;
    if (currentPermission === 'viewer') return;

    // Chuẩn hoá value về dạng primitive string để lưu DB và so sánh
    let value = rawValue;
    if (value && typeof value === 'object') {
      if (value.v != null) value = value.v;
      else if (value.m != null) value = value.m;
    }

    const key = `${r}:${c}`;
    if (lastCellValues[key] === value) {
      return;
    }
    lastCellValues[key] = value;

    const payload = {
      sheetId,
      row: r,
      column: c,
      value,
      userId: currentUser?.id,
      timestamp: new Date().toISOString()
    };
    socket.emit('cell_update', payload);
  }

  function emitPresence(r, c) {
    if (!socket || !socket.connected) return;
    socket.emit('presence', {
      sheetId,
      row: r,
      column: c
    });
  }

  async function loadCellMeta(r, c) {
    if (!Number.isInteger(r) || !Number.isInteger(c)) return;
    try {
      const meta = await apiRequest(
        `/sheets/${sheetId}/cell-meta?row=${r}&column=${c}`
      );
      if (!meta || !meta.updatedAt || !meta.updatedBy) {
        cellInfoMessage.textContent = 'No data for this cell yet.';
      } else {
        const when = new Date(meta.updatedAt).toLocaleString();
        cellInfoMessage.textContent = `Last edited by ${meta.updatedBy.username} at ${when}`;
      }
    } catch (err) {
      cellInfoMessage.textContent = err.message;
    }
  }

  function renderOnlineUsers() {
    onlineUsersEl.innerHTML = '';
    const entries = Object.values(remotePresenceByUser);
    if (!entries.length) return;
    entries.forEach((u) => {
      const li = document.createElement('li');
      const avatar = document.createElement('div');
      avatar.className = 'avatar-circle';
      avatar.textContent = (u.username || '?').charAt(0).toUpperCase();
      const label = document.createElement('span');
      label.textContent = u.username;
      li.appendChild(avatar);
      li.appendChild(label);
      onlineUsersEl.appendChild(li);
    });
  }

  function updateCellPresenceMessage() {
    if (!cellPresenceMessage) return;
    const currentRange = luckysheet.getRange?.();
    if (!currentRange || !Array.isArray(currentRange) || !currentRange[0]) {
      cellPresenceMessage.textContent = '';
      return;
    }
    const r = currentRange[0].row?.[0];
    const c = currentRange[0].column?.[0];
    if (typeof r !== 'number' || typeof c !== 'number') {
      cellPresenceMessage.textContent = '';
      return;
    }
    const atSameCell = Object.values(remotePresenceByUser).filter(
      (u) => u.row === r && u.column === c
    );
    if (!atSameCell.length) {
      cellPresenceMessage.textContent = '';
      return;
    }
    const names = atSameCell.map((u) => u.username).join(', ');
    cellPresenceMessage.textContent = `${names} đang chỉnh ô này`;
  }

  shareForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    shareMessage.textContent = '';

    try {
      const userId = Number(shareUserIdInput.value || '');
      const role = shareRoleSelect.value;
      if (!userId) {
        shareMessage.textContent = 'Please select a user.';
        return;
      }

      await apiRequest(`/sheets/${sheetId}/share`, {
        method: 'POST',
        body: { userId, role }
      });
      shareMessage.textContent = 'Shared successfully.';
    } catch (err) {
      shareMessage.textContent = err.message;
    }
  });

  cellInfoForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    cellInfoMessage.textContent = '';

    const row1 = Number(cellRowInput.value);
    const col1 = Number(cellColumnInput.value);
    if (!row1 || !col1) {
      cellInfoMessage.textContent = 'Enter row and column (1-based).';
      return;
    }

    const row = row1 - 1;
    const column = col1 - 1;

    try {
      const meta = await apiRequest(`/sheets/${sheetId}/cell-meta?row=${row}&column=${column}`);
      if (!meta || !meta.updatedAt || !meta.updatedBy) {
        cellInfoMessage.textContent = 'No data for this cell yet.';
      } else {
        const when = new Date(meta.updatedAt).toLocaleString();
        cellInfoMessage.textContent = `Last edited by ${meta.updatedBy.username} at ${when}`;
      }
    } catch (err) {
      cellInfoMessage.textContent = err.message;
    }
  });

  // Toggle panels from small icon buttons
  toggleCellInfoBtn.addEventListener('click', () => {
    cellInfoSection.classList.toggle('hidden');
  });

  toggleShareBtn.addEventListener('click', () => {
    shareSection.classList.toggle('hidden');
  });

  // Init
  (async () => {
    try {
      await loadCurrentUser();
      await loadUsersForShare();
      await loadSheet();
      connectSocket();
    } catch (err) {
      console.error(err);
      window.auth.clearToken();
      window.location.href = '/';
    }
  })();
})();

