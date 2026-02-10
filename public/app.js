(() => {
  const apiBase = '/api';
  let authToken = null;
  let currentUser = null;
  let currentSheetId = null;
  let currentTabId = null;
  let documentTabs = [];
  let currentPermission = null;
  let socket = null;
  let luckysheetInitialized = false;
  let isApplyingRemoteUpdate = false;
  const lastCellValues = {};
  let lastSelectedCell = { r: 0, c: 0 };

  const topUserNameEl = document.getElementById('top-user-name');
  const topLogoutBtn = document.getElementById('top-logout-btn');
  const topSheetNameEl = document.getElementById('top-sheet-name');
  const topUserAvatar = document.getElementById('top-user-avatar');
  const headerShareBtn = document.getElementById('header-share-btn');
  const shareModal = document.getElementById('share-modal');
  const shareForm = document.getElementById('share-form');
  const shareUserId = document.getElementById('share-user-id');
  const shareRoleSelect = document.getElementById('share-role');
  const shareMessage = document.getElementById('share-message');
  const shareCancelBtn = document.getElementById('share-cancel');
  const menuDropdowns = document.getElementById('gs-menu-dropdowns');
  const ribbonTabs = document.getElementById('ribbon-tabs');
  const sheetTabsList = document.getElementById('sheet-tabs-list');
  const sheetTabAddBtn = document.getElementById('sheet-tab-add');

  const sidebarOverlay = document.getElementById('sidebar-overlay');
  const sidebarBackdrop = document.getElementById('sidebar-backdrop');
  const sidebarEl = document.getElementById('sidebar');
  const sidebarToggleBtn = document.getElementById('sidebar-toggle');
  const navHome = document.getElementById('nav-home');
  const navMySheets = document.getElementById('nav-my-sheets');
  const navSharedSheets = document.getElementById('nav-shared-sheets');
  const navUsers = document.getElementById('nav-users');
  const navSettings = document.getElementById('nav-settings');

  // Views
  const viewLogin = document.getElementById('view-login');
  const viewHome = document.getElementById('view-home');
  const viewMySheets = document.getElementById('view-my-sheets');
  const viewSharedSheets = document.getElementById('view-shared-sheets');
  const viewUsers = document.getElementById('view-users');
  const viewSettings = document.getElementById('view-settings');
  const viewSheet = document.getElementById('view-sheet');

  // Login view
  const loginForm = document.getElementById('login-form');
  const loginUsernameInput = document.getElementById('login-username');
  const loginPasswordInput = document.getElementById('login-password');
  const loginError = document.getElementById('login-error');

  // Home view
  const homeRecentList = document.getElementById('home-recent-sheets');
  const homeSharedList = document.getElementById('home-shared-sheets');
  const homeCreateSheetBtn = document.getElementById('home-create-sheet-btn');

  // My sheets view
  const myCreateSheetBtn = document.getElementById('my-create-sheet-btn');
  const mySheetsTableBody = document.getElementById('my-sheets-table-body');

  // Shared sheets view
  const sharedSheetsTableBody = document.getElementById('shared-sheets-table-body');

  // User management view
  const umUsernameInput = document.getElementById('um-username');
  const umPasswordInput = document.getElementById('um-password');
  const umRoleSelect = document.getElementById('um-role');
  const umCreateBtn = document.getElementById('um-create-btn');
  const umMessage = document.getElementById('um-message');
  const umTableBody = document.getElementById('um-table-body');

  // Sidebar sheet list (quick access)
  const sheetsListEl = document.getElementById('sheets-list');
  const createSheetBtn = document.getElementById('create-sheet-btn');

  function setAuth(token, user) {
    authToken = token;
    currentUser = user;
    if (token) {
      localStorage.setItem('jwt', token);
    } else {
      localStorage.removeItem('jwt');
    }
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

  function showView(viewName) {
    const views = [
      viewLogin,
      viewHome,
      viewMySheets,
      viewSharedSheets,
      viewUsers,
      viewSettings,
      viewSheet
    ];
    views.forEach((v) => v && v.classList.add('hidden'));

    const navItems = [navHome, navMySheets, navSharedSheets, navUsers, navSettings];
    navItems.forEach((n) => n && n.classList.remove('active'));

    document.body.classList.toggle('gs-view-sheet', viewName === 'sheet');

    switch (viewName) {
      case 'login':
        document.body.classList.remove('gs-logged-in');
        if (viewLogin) viewLogin.classList.remove('hidden');
        if (topUserNameEl) topUserNameEl.textContent = '';
        if (topSheetNameEl) topSheetNameEl.textContent = '';
        if (topUserAvatar) topUserAvatar.textContent = '';
        break;
      case 'home':
        if (viewHome) viewHome.classList.remove('hidden');
        if (navHome) navHome.classList.add('active');
        break;
      case 'my-sheets':
        if (viewMySheets) viewMySheets.classList.remove('hidden');
        if (navMySheets) navMySheets.classList.add('active');
        break;
      case 'shared-sheets':
        if (viewSharedSheets) viewSharedSheets.classList.remove('hidden');
        if (navSharedSheets) navSharedSheets.classList.add('active');
        break;
      case 'users':
        if (viewUsers) viewUsers.classList.remove('hidden');
        if (navUsers) navUsers.classList.add('active');
        break;
      case 'settings':
        if (viewSettings) viewSettings.classList.remove('hidden');
        if (navSettings) navSettings.classList.add('active');
        break;
      case 'sheet':
        if (viewSheet) viewSheet.classList.remove('hidden');
        break;
      default:
        if (viewHome) viewHome.classList.remove('hidden');
        if (navHome) navHome.classList.add('active');
        break;
    }
  }

  async function tryRestoreSession() {
    const token = localStorage.getItem('jwt');
    if (!token) {
      showView('login');
      return;
    }
    authToken = token;
    try {
      const data = await apiRequest('/auth/me');
      currentUser = data.user;
      document.body.classList.add('gs-logged-in');
      if (topUserNameEl) topUserNameEl.textContent = `${currentUser.username} (${currentUser.role})`;
      if (topUserAvatar) topUserAvatar.textContent = (currentUser.username || '').charAt(0).toUpperCase();
      if (navUsers) {
        navUsers.style.display = (currentUser.role === 'admin' || currentUser.role === 'leader') ? '' : 'none';
      }
      connectSocket();
      await loadHome();
      showView('home');
    } catch (_) {
      setAuth(null, null);
      showView('login');
    }
  }

  if (loginForm) loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (loginError) loginError.textContent = '';
    try {
      const data = await apiRequest('/auth/login', {
        method: 'POST',
        body: {
          username: (loginUsernameInput && loginUsernameInput.value || '').trim(),
          password: (loginPasswordInput && loginPasswordInput.value) || ''
        }
      });
      setAuth(data.token, data.user);
      document.body.classList.add('gs-logged-in');
      if (topUserNameEl) topUserNameEl.textContent = `${data.user.username} (${data.user.role})`;
      if (topUserAvatar) topUserAvatar.textContent = (data.user.username || '').charAt(0).toUpperCase();
      if (navUsers) navUsers.style.display = (data.user.role === 'admin' || data.user.role === 'leader') ? '' : 'none';
      connectSocket();
      await loadHome();
      showView('home');
    } catch (err) {
      if (loginError) loginError.textContent = err.message;
    }
  });

  async function loadHome() {
    await loadSheetsList();
    if (!homeRecentList || !homeSharedList) return;
    // recent sheets
    try {
      const recentRes = await apiRequest('/sheets');
      homeRecentList.innerHTML = '';
      if (recentRes.sheets && recentRes.sheets.length) {
        recentRes.sheets.forEach((s) => {
          const li = document.createElement('li');
          const left = document.createElement('div');
          const right = document.createElement('div');
          left.textContent = s.name || 'Sheet';
          const meta = document.createElement('span');
          meta.className = 'sheet-permission';
          meta.textContent = s.permission;
          left.appendChild(meta);
          const openBtn = document.createElement('button');
          openBtn.textContent = 'Open';
          openBtn.addEventListener('click', () => openSheetView(s.id, s.permission));
          right.appendChild(openBtn);
          li.appendChild(left);
          li.appendChild(right);
          homeRecentList.appendChild(li);
        });
      } else {
        const li = document.createElement('li');
        li.textContent = 'No sheets yet. Create one!';
        homeRecentList.appendChild(li);
      }
    } catch (err) {
      homeRecentList.innerHTML = '';
      const li = document.createElement('li');
      li.textContent = 'Failed to load sheets.';
      homeRecentList.appendChild(li);
    }

    // Shared with me – for now use same source; in a real app this would be a separate API
    try {
      const sharedRes = await apiRequest('/sheets');
      homeSharedList.innerHTML = '';
      if (sharedRes.sheets && sharedRes.sheets.length) {
        sharedRes.sheets.forEach((s) => {
          const li = document.createElement('li');
          const left = document.createElement('div');
          const right = document.createElement('div');
          left.textContent = s.name || 'Sheet';
          const meta = document.createElement('span');
          meta.className = 'sheet-permission';
          meta.textContent = s.permission;
          left.appendChild(meta);
          const openBtn = document.createElement('button');
          openBtn.textContent = 'Open';
          openBtn.addEventListener('click', () => openSheetView(s.id, s.permission));
          right.appendChild(openBtn);
          li.appendChild(left);
          li.appendChild(right);
          homeSharedList.appendChild(li);
        });
      } else {
        const li = document.createElement('li');
        li.textContent = 'No shared sheets.';
        homeSharedList.appendChild(li);
      }
    } catch {
      homeSharedList.innerHTML = '';
      const li = document.createElement('li');
      li.textContent = 'Failed to load shared sheets.';
      homeSharedList.appendChild(li);
    }

    await loadSheetsList();
  }

  async function loadSheetsList() {
    try {
      const data = await apiRequest('/sheets');
      const sheets = data.sheets || [];
      if (sheetsListEl) {
        sheetsListEl.innerHTML = '';
        sheets.forEach((s) => {
          const li = document.createElement('li');
          li.dataset.id = String(s.id);
          li.textContent = s.name || 'Sheet';
          const perm = document.createElement('span');
          perm.className = 'sheet-permission';
          perm.textContent = s.permission;
          li.appendChild(perm);
          li.addEventListener('click', () => openSheetView(s.id, s.permission));
          sheetsListEl.appendChild(li);
        });
      }
    } catch (err) {
      console.error('Failed to load sheets', err);
    }
  }

  function renderSheetTabs() {
    if (!sheetTabsList) return;
    sheetTabsList.innerHTML = '';
    documentTabs.forEach((tab) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'sheet-tab' + (tab.id === currentTabId ? ' active' : '');
      btn.dataset.tabId = String(tab.id);
      btn.textContent = tab.name || 'Sheet';
      btn.title = tab.name || 'Sheet';
      btn.addEventListener('click', () => switchToTab(tab.id));
      btn.addEventListener('dblclick', (e) => {
        e.preventDefault();
        const newName = prompt('Rename tab', tab.name || 'Sheet');
        if (newName == null || newName.trim() === '') return;
        apiRequest(`/sheets/${currentSheetId}/tabs/${tab.id}`, { method: 'PATCH', body: { name: newName.trim() } })
          .then(() => {
            tab.name = newName.trim();
            btn.textContent = tab.name;
            btn.title = tab.name;
          })
          .catch((err) => alert(err.message));
      });
      sheetTabsList.appendChild(btn);
    });
  }

  function switchToTab(tabId) {
    if (tabId === currentTabId) return;
    const tab = documentTabs.find((t) => t.id === tabId);
    if (!tab) return;
    currentTabId = tabId;
    renderSheetTabs();
    if (typeof luckysheet.destroy === 'function') {
      try { luckysheet.destroy(); } catch (_) {}
    }
    luckysheetInitialized = false;
    luckysheet.create({
      container: 'luckysheet',
      data: [tab],
      showinfobar: false,
      enableAddRow: true,
      enableAddCol: true,
      sheetFormulaBar: true,
      hook: {
        cellUpdated: (r, c, oldValue, newValue) => {
          handleLocalCellChange(r, c, newValue);
          lastSelectedCell = { r, c };
        }
      }
    });
    luckysheetInitialized = true;
  }

  if (createSheetBtn) {
    createSheetBtn.addEventListener('click', async () => {
    try {
      const name = prompt('Sheet name?', 'Sheet1');
      const data = await apiRequest('/sheets', {
        method: 'POST',
        body: { name }
      });
      await loadHome();
      if (data.sheet && data.sheet.id) {
        openSheetView(data.sheet.id, 'owner');
      }
    } catch (err) {
      alert(err.message);
    }
  });
  }

  if (homeCreateSheetBtn) {
    homeCreateSheetBtn.addEventListener('click', async () => {
    try {
      const name = prompt('Sheet name?', 'Sheet1');
      const data = await apiRequest('/sheets', {
        method: 'POST',
        body: { name }
      });
      await loadHome();
      if (data.sheet && data.sheet.id) {
        openSheetView(data.sheet.id, 'owner');
      }
    } catch (err) {
      alert(err.message);
    }
  });
  }

  if (myCreateSheetBtn) {
    myCreateSheetBtn.addEventListener('click', async () => {
    try {
      const name = prompt('Sheet name?', 'Sheet1');
      const data = await apiRequest('/sheets', {
        method: 'POST',
        body: { name }
      });
      await loadMySheets();
      if (data.sheet && data.sheet.id) {
        openSheetView(data.sheet.id, 'owner');
      }
    } catch (err) {
      alert(err.message);
    }
  });
  }

  if (topLogoutBtn) {
    topLogoutBtn.addEventListener('click', () => {
    if (socket) {
      socket.disconnect();
      socket = null;
    }
    setAuth(null, null);
    luckysheetInitialized = false;
    currentSheetId = null;
    currentTabId = null;
    documentTabs = [];
    currentPermission = null;
    topUserNameEl.textContent = '';
    topSheetNameEl.textContent = '';
    showView('login');
  });
  }

  function closeSidebar() {
    if (sidebarOverlay) sidebarOverlay.classList.remove('open');
  }
  function openSidebar() {
    if (sidebarOverlay) sidebarOverlay.classList.add('open');
  }
  if (sidebarToggleBtn) {
    sidebarToggleBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      if (sidebarOverlay && sidebarOverlay.classList.contains('open')) closeSidebar();
      else openSidebar();
    });
  }
  if (sidebarBackdrop) {
    sidebarBackdrop.addEventListener('click', closeSidebar);
  }

  if (navHome) navHome.addEventListener('click', async () => {
    await loadHome();
    showView('home');
    closeSidebar();
  });
  if (navMySheets) navMySheets.addEventListener('click', async () => {
    await loadMySheets();
    showView('my-sheets');
    closeSidebar();
  });
  if (navSharedSheets) navSharedSheets.addEventListener('click', async () => {
    await loadSharedSheets();
    showView('shared-sheets');
    closeSidebar();
  });
  if (navUsers) navUsers.addEventListener('click', async () => {
    await loadUsersView();
    showView('users');
    closeSidebar();
  });
  if (navSettings) navSettings.addEventListener('click', () => {
    showView('settings');
    closeSidebar();
  });

  if (ribbonTabs) {
    ribbonTabs.querySelectorAll('.ribbon-tab').forEach((tabEl) => {
      tabEl.addEventListener('click', (e) => {
        e.stopPropagation();
        const tab = tabEl.dataset.tab;
        if (tab === 'file') {
          const dropdown = document.getElementById('dropdown-file');
          const open = dropdown && dropdown.classList.toggle('open');
          menuDropdowns.querySelectorAll('.menu-dropdown').forEach((d) => d.classList.remove('open'));
          if (dropdown && open) {
            const rect = tabEl.getBoundingClientRect();
            dropdown.style.left = rect.left + 'px';
            dropdown.style.top = (rect.bottom + 2) + 'px';
            dropdown.classList.add('open');
          }
          return;
        }
        ribbonTabs.querySelectorAll('.ribbon-tab').forEach((t) => t.classList.remove('active'));
        tabEl.classList.add('active');
        document.querySelectorAll('.ribbon-panel').forEach((p) => p.classList.remove('active'));
        const panel = document.getElementById('ribbon-' + tab);
        if (panel) panel.classList.add('active');
      });
    });
  }
  document.addEventListener('click', () => {
    if (menuDropdowns) menuDropdowns.querySelectorAll('.menu-dropdown').forEach((d) => d.classList.remove('open'));
  });
  if (menuDropdowns) {
    menuDropdowns.addEventListener('click', (e) => e.stopPropagation());
    menuDropdowns.addEventListener('click', (e) => {
      const item = e.target.closest('.menu-dropdown-item');
      if (!item) return;
      const action = item.dataset.action;
      menuDropdowns.querySelectorAll('.menu-dropdown').forEach((d) => d.classList.remove('open'));
      if (action === 'new-sheet') {
        (async () => {
          try {
            const name = prompt('Sheet name?', 'Sheet1');
            const data = await apiRequest('/sheets', { method: 'POST', body: { name: name || 'Sheet1' } });
            await loadSheetsList();
            openSheetView(data.sheet.id, 'owner');
          } catch (err) { alert(err.message); }
        })();
      }
      if (action === 'permissions' && headerShareBtn) headerShareBtn.click();
    });
  }

  document.querySelectorAll('.ribbon-panel').forEach((panel) => {
    panel.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-action]');
      if (!btn) return;
      const action = btn.dataset.action;
      if (action === 'insert-row-above' || action === 'insert-row-below' || action === 'insert-column-left' || action === 'insert-column-right') return;
      if (action === 'permissions') { if (headerShareBtn) headerShareBtn.click(); return; }
      if (action === 'history' || action === 'export') return;
    });
  });

  if (sheetTabAddBtn) {
    sheetTabAddBtn.addEventListener('click', async () => {
      if (!currentSheetId) return;
      try {
        const name = prompt('Tab name?', 'Sheet' + (documentTabs.length + 1));
        const tab = await apiRequest(`/sheets/${currentSheetId}/tabs`, {
          method: 'POST',
          body: { name: name || ('Sheet' + (documentTabs.length + 1)) }
        });
        documentTabs.push({
          id: tab.id,
          name: tab.name,
          index: documentTabs.length,
          row: 100,
          column: 26,
          celldata: []
        });
        renderSheetTabs();
        switchToTab(tab.id);
      } catch (err) {
        alert(err.message);
      }
    });
  }

  if (headerShareBtn) {
    headerShareBtn.addEventListener('click', async () => {
      if (!currentSheetId || !shareModal) return;
      shareMessage.textContent = '';
      shareUserId.innerHTML = '<option value="">Select user</option>';
      try {
        const data = await apiRequest('/users');
        (data.users || []).forEach((u) => {
          if (u.id === currentUser.id) return;
          const opt = document.createElement('option');
          opt.value = u.id;
          opt.textContent = u.username;
          shareUserId.appendChild(opt);
        });
        shareModal.classList.remove('hidden');
      } catch (err) {
        shareMessage.textContent = err.message;
      }
    });
  }
  if (shareCancelBtn && shareModal) {
    shareCancelBtn.addEventListener('click', () => shareModal.classList.add('hidden'));
  }
  if (shareModal) {
    shareModal.querySelector('.modal-backdrop').addEventListener('click', () => shareModal.classList.add('hidden'));
  }
  if (shareForm) {
    shareForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (!currentSheetId || !shareUserId || !shareRoleSelect) return;
      const userId = Number(shareUserId.value);
      if (!userId) {
        shareMessage.textContent = 'Select a user';
        return;
      }
      shareMessage.textContent = '';
      try {
        await apiRequest(`/sheets/${currentSheetId}/share`, {
          method: 'POST',
          body: { userId, role: shareRoleSelect.value }
        });
        shareMessage.textContent = 'Shared successfully.';
        shareModal.classList.add('hidden');
      } catch (err) {
        shareMessage.textContent = err.message;
      }
    });
  }

  async function loadMySheets() {
    if (!mySheetsTableBody) return;
    const data = await apiRequest('/sheets');
    mySheetsTableBody.innerHTML = '';
    data.sheets.forEach((s) => {
      if (s.permission !== 'owner') return;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${s.name || 'Sheet'}</td>
        <td>${s.updated_at || ''}</td>
        <td>${s.permission}</td>
        <td><button data-open-id="${s.id}">Open</button></td>
      `;
      mySheetsTableBody.appendChild(tr);
    });
    mySheetsTableBody.querySelectorAll('button[data-open-id]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-open-id'));
        const row = Array.from(btn.closest('tr').children);
        const perm = row[2].textContent || 'owner';
        openSheetView(id, perm);
      });
    });
  }

  async function loadSharedSheets() {
    if (!sharedSheetsTableBody) return;
    const data = await apiRequest('/sheets');
    sharedSheetsTableBody.innerHTML = '';
    data.sheets.forEach((s) => {
      if (s.permission === 'owner') return;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${s.name || 'Sheet'}</td>
        <td>${s.owner_id || ''}</td>
        <td>${s.permission}</td>
        <td>${s.updated_at || ''}</td>
        <td><button data-open-id="${s.id}">Open</button></td>
      `;
      sharedSheetsTableBody.appendChild(tr);
    });
    sharedSheetsTableBody.querySelectorAll('button[data-open-id]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-open-id'));
        const row = Array.from(btn.closest('tr').children);
        const perm = row[2].textContent || 'viewer';
        openSheetView(id, perm);
      });
    });
  }

  async function loadUsersView() {
    if (!umTableBody) return;
    if (!currentUser || (currentUser.role !== 'admin' && currentUser.role !== 'leader')) {
      umTableBody.innerHTML = '';
      return;
    }
    try {
      const data = await apiRequest('/users');
      umTableBody.innerHTML = '';
      data.users.forEach((u) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${u.username}</td>
          <td>${u.role}</td>
          <td>${u.created_by || ''}</td>
          <td><button data-del-id="${u.id}">Remove</button></td>
        `;
        umTableBody.appendChild(tr);
      });
      umTableBody.querySelectorAll('button[data-del-id]').forEach((btn) => {
        btn.addEventListener('click', async () => {
          const id = Number(btn.getAttribute('data-del-id'));
          if (!confirm('Remove this user?')) return;
          try {
            await apiRequest(`/users/${id}`, { method: 'DELETE' });
            await loadUsersView();
          } catch (err) {
            umMessage.textContent = err.message;
          }
        });
      });
    } catch (err) {
      umMessage.textContent = err.message;
    }
  }

  if (umCreateBtn) umCreateBtn.addEventListener('click', async () => {
    umMessage.textContent = '';
    const username = umUsernameInput.value.trim();
    const password = umPasswordInput.value;
    const role = umRoleSelect.value;
    if (!username || !password) {
      umMessage.textContent = 'Username and password required';
      return;
    }
    try {
      await apiRequest('/users', {
        method: 'POST',
        body: { username, password, role }
      });
      umUsernameInput.value = '';
      umPasswordInput.value = '';
      umMessage.textContent = 'User created';
      await loadUsersView();
    } catch (err) {
      umMessage.textContent = err.message;
    }
  });

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
      if (currentSheetId) {
        socket.emit('join_sheet', { sheetId: currentSheetId });
      }
    });

    socket.on('cell_update', (payload) => {
      const { sheetId, sheetTabId, row, column, value } = payload || {};
      if (!luckysheetInitialized) return;
      if (sheetId !== currentSheetId || sheetTabId !== currentTabId) return;

      isApplyingRemoteUpdate = true;
      try {
        luckysheet.setCellValue(row, column, value);
        const key = `${row}:${column}`;
        lastCellValues[key] = value;
      } finally {
        isApplyingRemoteUpdate = false;
      }
    });
  }

  async function openSheetView(sheetId, permission) {
    currentSheetId = sheetId;
    currentPermission = permission;

    if (sheetsListEl) {
      Array.from(sheetsListEl.children).forEach((li) => {
        li.classList.toggle('active', Number(li.dataset.id) === sheetId);
      });
    }

    const data = await apiRequest(`/sheets/${sheetId}`);
    const tabs = data.tabs || [];
    documentTabs = tabs;
    currentPermission = data.permission || permission;
    if (topSheetNameEl) topSheetNameEl.textContent = data.sheet?.name || 'Sheet';

    if (typeof luckysheet.destroy === 'function') {
      try { luckysheet.destroy(); } catch (_) {}
    }
    luckysheetInitialized = false;

    if (tabs.length === 0) {
      const newTab = await apiRequest(`/sheets/${sheetId}/tabs`, { method: 'POST', body: { name: 'Sheet1' } });
      documentTabs = [{ id: newTab.id, name: newTab.name, index: 0, row: 100, column: 26, celldata: [] }];
      currentTabId = newTab.id;
      luckysheet.create({
        container: 'luckysheet',
        data: [documentTabs[0]],
        showinfobar: false,
        enableAddRow: true,
        enableAddCol: true,
        sheetFormulaBar: true,
        hook: {
          cellUpdated: (r, c, oldValue, newValue) => {
            handleLocalCellChange(r, c, newValue);
            lastSelectedCell = { r, c };
          }
        }
      });
      luckysheetInitialized = true;
    } else {
      currentTabId = tabs[0].id;
      luckysheet.create({
        container: 'luckysheet',
        data: [tabs[0]],
        showinfobar: false,
        enableAddRow: true,
        enableAddCol: true,
        sheetFormulaBar: true,
        hook: {
          cellUpdated: (r, c, oldValue, newValue) => {
            handleLocalCellChange(r, c, newValue);
            lastSelectedCell = { r, c };
          }
        }
      });
      luckysheetInitialized = true;
    }

    if (socket && socket.connected) {
      socket.emit('join_sheet', { sheetId });
    }

    if (sidebarOverlay) sidebarOverlay.classList.remove('open');
    renderSheetTabs();
    showView('sheet');
  }

  function handleLocalCellChange(r, c, rawValue) {
    if (!socket || !socket.connected) return;
    if (!currentSheetId) return;
    if (isApplyingRemoteUpdate) return;
    if (currentPermission === 'viewer') {
      // Read-only: do not send updates
      return;
    }

    // Normalize Luckysheet value object to primitive
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
      sheetId: currentSheetId,
      sheetTabId: currentTabId,
      row: r,
      column: c,
      value,
      userId: currentUser?.id,
      timestamp: new Date().toISOString()
    };
    socket.emit('cell_update', payload);
  }

  // Initialize
  tryRestoreSession();
})();

