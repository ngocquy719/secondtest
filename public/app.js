(() => {
  const apiBase = '/api';
  let authToken = null;
  let currentUser = null;
  let currentSheetId = null;
  let currentPermission = null;
  let socket = null;
  let luckysheetInitialized = false;
  let isApplyingRemoteUpdate = false;
  const lastCellValues = {};
  let lastSelectedCell = { r: 0, c: 0 };

  function getSheets() { return (window.SheetManager && SheetManager.getState().sheets) || []; }
  function getActiveSheetId() { return (window.SheetManager && SheetManager.getState().activeSheetId) || null; }
  function setActiveSheetId(id) { if (window.SheetManager) SheetManager.switchSheet(id); }

  const topUserNameEl = document.getElementById('top-user-name');
  const topLogoutBtn = document.getElementById('top-logout-btn');
  const gsDocTitleEl = document.getElementById('gs-doc-title');
  const gsSaveStatusEl = document.getElementById('gs-save-status');
  const topUserAvatar = document.getElementById('top-user-avatar');
  const headerShareBtn = document.getElementById('header-share-btn');
  const shareModal = document.getElementById('share-modal');
  const shareForm = document.getElementById('share-form');
  const shareUserId = document.getElementById('share-user-id');
  const shareRoleSelect = document.getElementById('share-role');
  const shareMessage = document.getElementById('share-message');
  const shareCancelBtn = document.getElementById('share-cancel');
  const menuDropdowns = document.getElementById('gs-menu-dropdowns');
  const sheetTabsList = document.getElementById('sheet-tabs-list');
  const sheetTabAddBtn = document.getElementById('sheet-tab-add');
  const sheetTabsListBtn = document.getElementById('sheet-tabs-list-btn');
  const sheetTabContextMenu = document.getElementById('sheet-tab-context-menu');
  const sheetTabsListModal = document.getElementById('sheet-tabs-list-modal');
  const sheetTabsListModalUl = document.getElementById('sheet-tabs-list-modal-ul');
  const currentCellRefEl = document.getElementById('current-cell-ref');

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

  const minimalBar = document.getElementById('minimal-bar');
  const minimalUserName = document.getElementById('minimal-user-name');
  const minimalLogoutBtn = document.getElementById('minimal-logout-btn');
  const logoLink = document.getElementById('logo-link');
  const logoLinkSheet = document.getElementById('logo-link-sheet');

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
        if (gsDocTitleEl) gsDocTitleEl.textContent = '';
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
      if (minimalUserName) minimalUserName.textContent = `${currentUser.username} (${currentUser.role})`;
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
      if (minimalUserName) minimalUserName.textContent = `${data.user.username} (${data.user.role})`;
      if (navUsers) navUsers.style.display = (data.user.role === 'admin' || data.user.role === 'leader') ? '' : 'none';
      connectSocket();
      await loadHome();
      showView('home');
    } catch (err) {
      if (loginError) loginError.textContent = err.message;
    }
  });

  async function loadHome() {
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
  }

  function renderSheetTabs() {
    if (!sheetTabsList) return;
    const sheets = getSheets();
    const activeId = getActiveSheetId();
    sheetTabsList.innerHTML = '';
    sheets.forEach((tab) => {
      const wrap = document.createElement('div');
      wrap.className = 'gs-sheet-tab' + (tab.id === activeId ? ' active' : '');
      wrap.dataset.tabId = String(tab.id);
      wrap.title = tab.name || 'Sheet';
      const label = document.createElement('span');
      label.textContent = tab.name || 'Sheet';
      label.style.flex = '1';
      label.style.overflow = 'hidden';
      label.style.textOverflow = 'ellipsis';
      const arrow = document.createElement('span');
      arrow.className = 'gs-sheet-tab-arrow';
      arrow.textContent = '▼';
      arrow.setAttribute('aria-label', 'Tab menu');
      wrap.appendChild(label);
      wrap.appendChild(arrow);

      const openContextMenu = (e) => {
        e.preventDefault();
        e.stopPropagation();
        sheetTabContextMenu.dataset.tabId = String(tab.id);
        sheetTabContextMenu.classList.remove('hidden');
        sheetTabContextMenu.style.left = e.clientX + 'px';
        sheetTabContextMenu.style.top = e.clientY + 'px';
      };

      label.addEventListener('click', (e) => {
        e.stopPropagation();
        switchToTab(tab.id);
      });
      arrow.addEventListener('click', openContextMenu);
      wrap.addEventListener('contextmenu', openContextMenu);

      sheetTabsList.appendChild(wrap);
    });
  }

  function closeTabContextMenu() {
    if (sheetTabContextMenu) {
      sheetTabContextMenu.classList.add('hidden');
      sheetTabContextMenu.dataset.tabId = '';
    }
  }

  function handleTabContextAction(action) {
    const tabId = Number(sheetTabContextMenu.dataset.tabId);
    if (!tabId) return;
    closeTabContextMenu();
    const sheets = getSheets();
    const tab = sheets.find((t) => t.id === tabId);
    if (!tab) return;
    if (action === 'rename') {
      const newName = prompt('Rename sheet', tab.name || 'Sheet');
      if (newName == null || newName.trim() === '') return;
      apiRequest(`/sheets/${currentSheetId}/tabs/${tabId}`, { method: 'PATCH', body: { name: newName.trim() } })
        .then(() => {
          if (window.SheetManager) SheetManager.renameSheet(tabId, newName.trim());
          renderSheetTabs();
        })
        .catch((err) => {
          console.error('Rename tab failed', err);
          alert(err && err.message ? err.message : 'Failed to rename');
        });
      return;
    }
    if (action === 'duplicate') {
      apiRequest(`/sheets/${currentSheetId}/tabs/${tabId}/duplicate`, { method: 'POST' })
        .then((newTab) => {
          return apiRequest(`/sheets/${currentSheetId}`).then((data) => {
          if (!window.SheetManager) return;
          SheetManager.setDocument(currentSheetId, data.tabs || [], data.permission);
          if (window.SpreadsheetEngine) {
            SpreadsheetEngine.loadFromTabs(SheetManager.getState().sheets);
            createLuckysheetWithTabs(SpreadsheetEngine.toLuckysheetTabs());
          } else {
            createLuckysheetWithTabs(SheetManager.getState().sheets);
          }
          renderSheetTabs();
          setActiveSheetId(newTab.id);
            const order = SheetManager.getOrderById(newTab.id);
            if (order >= 0) triggerLuckysheetSheetTabClick(order);
          });
        })
        .catch((err) => {
          console.error('Duplicate sheet failed', err);
          alert(err && err.message ? err.message : 'Failed to duplicate sheet');
        });
      return;
    }
    if (action === 'delete') {
      if (sheets.length <= 1) {
        alert('Cannot delete the only sheet. Add another sheet first.');
        return;
      }
      if (!confirm('Delete sheet "' + (tab.name || 'Sheet') + '"?')) return;
      apiRequest(`/sheets/${currentSheetId}/tabs/${tabId}`, { method: 'DELETE' })
        .then(() => {
          if (window.SheetManager) SheetManager.removeSheet(tabId);
          const nextSheets = getSheets();
          if (nextSheets.length === 0) return;
          if (window.SpreadsheetEngine) SpreadsheetEngine.removeTab(tabId);
          createLuckysheetWithTabs(window.SpreadsheetEngine ? SpreadsheetEngine.toLuckysheetTabs() : nextSheets);
          renderSheetTabs();
          triggerLuckysheetSheetTabClick(0);
        })
        .catch((err) => {
          console.error('Delete tab failed', err);
          alert(err && err.message ? err.message : 'Failed to delete sheet');
        });
      return;
    }
    if (action === 'move-left' || action === 'move-right') {
      const idx = sheets.findIndex((t) => t.id === tabId);
      if (idx < 0) return;
      const newIdx = action === 'move-left' ? Math.max(0, idx - 1) : Math.min(sheets.length - 1, idx + 1);
      if (newIdx === idx) return;
      apiRequest(`/sheets/${currentSheetId}/tabs/${tabId}/move`, { method: 'POST', body: { order_index: newIdx } })
        .then(() => {
          const arr = sheets.slice();
          const [rem] = arr.splice(idx, 1);
          arr.splice(newIdx, 0, rem);
          arr.forEach((s, i) => { s.order = i; s.index = i; });
          if (window.SheetManager) SheetManager.setSheets(arr);
          if (window.SpreadsheetEngine) SpreadsheetEngine.setTabs(getSheets());
          createLuckysheetWithTabs(window.SpreadsheetEngine ? SpreadsheetEngine.toLuckysheetTabs() : getSheets());
          renderSheetTabs();
          setActiveSheetId(tabId);
          triggerLuckysheetSheetTabClick(newIdx);
        })
        .catch((err) => {
          console.error('Move tab failed', err);
          alert(err && err.message ? err.message : 'Failed to move sheet');
        });
    }
  }

  if (sheetTabContextMenu) {
    sheetTabContextMenu.querySelectorAll('.gs-context-item').forEach((el) => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        handleTabContextAction(el.dataset.tabAction);
      });
    });
  }
  document.addEventListener('click', closeTabContextMenu);

  if (sheetTabsListBtn && sheetTabsListModal) {
    sheetTabsListBtn.addEventListener('click', () => {
      if (!sheetTabsListModalUl) return;
      const sheets = getSheets();
      const activeId = getActiveSheetId();
      sheetTabsListModalUl.innerHTML = '';
      sheets.forEach((tab) => {
        const li = document.createElement('li');
        li.textContent = tab.name || 'Sheet';
        li.dataset.tabId = String(tab.id);
        if (tab.id === activeId) li.classList.add('active');
        li.addEventListener('click', () => {
          switchToTab(tab.id);
          sheetTabsListModal.classList.add('hidden');
        });
        sheetTabsListModalUl.appendChild(li);
      });
      sheetTabsListModal.classList.remove('hidden');
    });
  }
  if (sheetTabsListModal) {
    const backdrop = document.getElementById('sheet-tabs-list-backdrop');
    const closeBtn = document.getElementById('sheet-tabs-list-close');
    if (backdrop) backdrop.addEventListener('click', () => sheetTabsListModal.classList.add('hidden'));
    if (closeBtn) closeBtn.addEventListener('click', () => sheetTabsListModal.classList.add('hidden'));
  }

  function switchToTab(tabId) {
    if (tabId === getActiveSheetId()) return;
    const order = window.SheetManager ? SheetManager.getOrderById(tabId) : -1;
    if (order < 0) return;
    setActiveSheetId(tabId);
    renderSheetTabs();
    if (luckysheetInitialized) triggerLuckysheetSheetTabClick(order);
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

  function doLogout() {
    if (socket) {
      socket.disconnect();
      socket = null;
    }
    setAuth(null, null);
    luckysheetInitialized = false;
    currentSheetId = null;
    currentPermission = null;
    if (window.SheetManager) SheetManager.reset();
    if (window.SpreadsheetEngine) SpreadsheetEngine.reset();
    if (topUserNameEl) topUserNameEl.textContent = '';
    if (gsDocTitleEl) gsDocTitleEl.textContent = '';
    if (minimalUserName) minimalUserName.textContent = '';
    showView('login');
  }
  if (topLogoutBtn) topLogoutBtn.addEventListener('click', doLogout);
  if (minimalLogoutBtn) minimalLogoutBtn.addEventListener('click', doLogout);

  if (logoLink) {
    logoLink.addEventListener('click', (e) => {
      e.preventDefault();
      closeSidebar();
      loadHome();
      showView('home');
    });
  }
  if (logoLinkSheet) {
    logoLinkSheet.addEventListener('click', (e) => {
      e.preventDefault();
      closeSidebar();
      loadHome();
      showView('home');
    });
  }

  if (gsDocTitleEl) {
    gsDocTitleEl.addEventListener('blur', () => {
      const sheetId = gsDocTitleEl.dataset.sheetId;
      const name = (gsDocTitleEl.textContent || '').trim() || 'Untitled spreadsheet';
      if (!sheetId || !currentSheetId || Number(sheetId) !== currentSheetId) return;
      apiRequest(`/sheets/${currentSheetId}`, { method: 'PATCH', body: { name } })
        .then(() => { if (gsSaveStatusEl) gsSaveStatusEl.textContent = 'Saved'; })
        .catch(() => { gsDocTitleEl.textContent = name; });
    });
    gsDocTitleEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') e.target.blur();
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

  const fileMenuBtn = document.querySelector('.gs-menu-item[data-menu="file"]');
  if (fileMenuBtn && menuDropdowns) {
    fileMenuBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const dropdown = document.getElementById('dropdown-file');
      const open = dropdown && dropdown.classList.toggle('open');
      menuDropdowns.querySelectorAll('.gs-menu-dropdown').forEach((d) => d.classList.remove('open'));
      if (dropdown && open) {
        const rect = fileMenuBtn.getBoundingClientRect();
        dropdown.style.left = rect.left + 'px';
        dropdown.style.top = (rect.bottom + 2) + 'px';
        dropdown.classList.add('open');
      }
    });
  }
  document.addEventListener('click', () => {
    if (menuDropdowns) menuDropdowns.querySelectorAll('.gs-menu-dropdown').forEach((d) => d.classList.remove('open'));
  });
  if (menuDropdowns) {
    menuDropdowns.addEventListener('click', (e) => e.stopPropagation());
    menuDropdowns.addEventListener('click', (e) => {
      const item = e.target.closest('.gs-menu-dropdown-item');
      if (!item) return;
      const action = item.dataset.action;
      menuDropdowns.querySelectorAll('.gs-menu-dropdown').forEach((d) => d.classList.remove('open'));
      if (action === 'new-sheet') {
        (async () => {
          try {
            const data = await apiRequest('/sheets', { method: 'POST', body: { name: 'Untitled spreadsheet' } });
            openSheetView(data.sheet.id, 'owner');
          } catch (err) { alert(err.message); }
        })();
      }
      if (action === 'permissions' && headerShareBtn) headerShareBtn.click();
      if (action === 'history') { /* version history placeholder */ }
      if (action === 'export') { /* download placeholder */ }
    });
  }

  if (sheetTabAddBtn) {
    sheetTabAddBtn.addEventListener('click', async () => {
      if (!currentSheetId || !window.SheetManager) return;
      try {
        const sheets = getSheets();
        const defaultName = 'Sheet' + (sheets.length + 1);
        const newTab = await apiRequest(`/sheets/${currentSheetId}/tabs`, {
          method: 'POST',
          body: { name: defaultName }
        });
        const order = sheets.length;
        const minimalSheet = {
          id: newTab.id,
          name: newTab.name,
          index: order,
          order,
          status: 0,
          row: 100,
          column: 26,
          celldata: [],
          config: {}
        };
        SheetManager.addSheet(minimalSheet);
        if (window.SpreadsheetEngine) SpreadsheetEngine.setTabs(getSheets());
        createLuckysheetWithTabs(window.SpreadsheetEngine ? SpreadsheetEngine.toLuckysheetTabs() : getSheets());
        renderSheetTabs();
        triggerLuckysheetSheetTabClick(order);
      } catch (err) {
        console.error('Add sheet tab failed', err);
        alert(err && err.message ? err.message : 'Failed to add sheet');
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
      if (sheetId !== currentSheetId) return;

      isApplyingRemoteUpdate = true;
      try {
        if (window.SpreadsheetEngine) {
          const updates = SpreadsheetEngine.setCell(sheetTabId, row, column, value);
          applyEngineUpdatesToLuckysheet(updates);
          updates.forEach((u) => {
            lastCellValues[`${u.tabId}:${u.r}:${u.c}`] = u.value;
          });
        } else {
          if (sheetTabId !== getActiveSheetId()) {
            isApplyingRemoteUpdate = false;
            return;
          }
          luckysheet.setCellValue(row, column, value);
          lastCellValues[`${row}:${column}`] = value;
        }
      } finally {
        isApplyingRemoteUpdate = false;
      }
    });
  }

  function applyEngineUpdatesToLuckysheet(updates) {
    if (!luckysheetInitialized || !updates || !updates.length) return;
    if (typeof luckysheet.setSheetActive !== 'function' || typeof luckysheet.setCellValue !== 'function') return;
    const currentOrder = window.SheetManager ? SheetManager.getOrderById(getActiveSheetId()) : 0;
    const byTab = new Map();
    updates.forEach((u) => {
      const order = window.SheetManager ? SheetManager.getOrderById(u.tabId) : 0;
      if (order < 0) return;
      if (!byTab.has(order)) byTab.set(order, []);
      byTab.get(order).push(u);
    });
    byTab.forEach((list, order) => {
      if (Number(order) !== currentOrder) try { luckysheet.setSheetActive(Number(order)); } catch (_) {}
      list.forEach((u) => { try { luckysheet.setCellValue(u.r, u.c, u.value); } catch (_) {} });
      if (Number(order) !== currentOrder) try { luckysheet.setSheetActive(currentOrder); } catch (_) {}
    });
  }

  function createLuckysheetWithTabs(tabsData) {
    if (typeof luckysheet.destroy === 'function') {
      try { luckysheet.destroy(); } catch (_) {}
    }
    luckysheetInitialized = false;
    if (!tabsData || tabsData.length === 0) return;

    luckysheet.create({
      container: 'luckysheet',
      data: tabsData,
      showinfobar: false,
      showsheetbar: true,
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

  function triggerLuckysheetSheetTabClick(order) {
    if (typeof order !== 'number' || order < 0) return;
    if (typeof luckysheet.setSheetActive === 'function') {
      try { luckysheet.setSheetActive(order); return; } catch (_) {}
    }
    var container = document.querySelector('#luckysheet-sheet-area .luckysheet-sheet-container-c') || document.querySelector('.luckysheet-sheet-container-c') || document.querySelector('[class*="sheet-container"]');
    if (container) {
      var items = container.querySelectorAll('[class*="sheet-item"], [class*="sheet-container-item"], .luckysheet-sheet-list-item, div[class*="sheet"]');
      if (items.length > order && items[order]) items[order].click();
    }
  }

  async function openSheetView(sheetId, permission) {
    currentSheetId = sheetId;
    currentPermission = permission;
    try {
    const data = await apiRequest(`/sheets/${sheetId}`);
    let tabs = data.tabs || [];
    currentPermission = data.permission || permission;
    if (gsDocTitleEl) {
      gsDocTitleEl.textContent = data.sheet?.name || 'Untitled spreadsheet';
      gsDocTitleEl.dataset.sheetId = String(currentSheetId);
    }
    if (gsSaveStatusEl) gsSaveStatusEl.textContent = 'Saved';

    if (tabs.length === 0) {
      const newTab = await apiRequest(`/sheets/${sheetId}/tabs`, { method: 'POST', body: { name: 'Sheet1' } });
      tabs = [{
        id: newTab.id,
        name: newTab.name,
        index: 0,
        order: 0,
        status: 1,
        row: 100,
        column: 26,
        celldata: [],
        config: {}
      }];
    }

    if (typeof luckysheet.destroy === 'function') {
      try { luckysheet.destroy(); } catch (_) {}
    }
    luckysheetInitialized = false;

    if (window.SheetManager) {
      SheetManager.setDocument(sheetId, tabs, currentPermission);
      setActiveSheetId(tabs[0].id);
    }
    if (window.SpreadsheetEngine) {
      SpreadsheetEngine.setDocumentId(sheetId);
      SpreadsheetEngine.loadFromTabs(getSheets());
    }
    createLuckysheetWithTabs(window.SpreadsheetEngine ? SpreadsheetEngine.toLuckysheetTabs() : getSheets());

    if (socket && socket.connected) {
      socket.emit('join_sheet', { sheetId });
    }

    if (sidebarOverlay) sidebarOverlay.classList.remove('open');
    renderSheetTabs();
    showView('sheet');
    } catch (err) {
      console.error('Open sheet failed', err);
      alert(err && err.message ? err.message : 'Failed to open spreadsheet');
    }
  }

  function handleLocalCellChange(r, c, rawValue) {
    if (!currentSheetId) return;
    if (isApplyingRemoteUpdate) return;
    if (currentPermission === 'viewer') return;

    let input = rawValue;
    if (input != null && typeof input === 'object') {
      input = input.v != null ? input.v : input.m;
    }

    const activeTabId = getActiveSheetId();
    if (!activeTabId) return;

    if (window.SpreadsheetEngine) {
      const updates = SpreadsheetEngine.setCell(activeTabId, r, c, input);
      applyEngineUpdatesToLuckysheet(updates);
      if (updates[0]) {
        lastCellValues[`${activeTabId}:${r}:${c}`] = updates[0].value;
        if (socket && socket.connected) {
          socket.emit('cell_update', {
            sheetId: currentSheetId,
            sheetTabId: activeTabId,
            row: r,
            column: c,
            value: input,
            userId: currentUser?.id,
            timestamp: new Date().toISOString()
          });
        }
      }
      return;
    }

    const key = `${r}:${c}`;
    if (lastCellValues[key] === input) return;
    lastCellValues[key] = input;
    if (socket && socket.connected) {
      socket.emit('cell_update', {
        sheetId: currentSheetId,
        sheetTabId: activeTabId,
        row: r,
        column: c,
        value: input,
        userId: currentUser?.id,
        timestamp: new Date().toISOString()
      });
    }
  }

  function colIndexToLetters(c) {
    let s = '';
    let n = c + 1;
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function bindToolbarCommands() {
    if (typeof ToolbarCommands === 'undefined' || !ToolbarCommands.executeCommand) return;
    const ex = (cmd, payload) => () => { if (luckysheetInitialized) ToolbarCommands.executeCommand(cmd, payload); };
    const byId = (id) => document.getElementById(id);
    byId('toolbar-bold')?.addEventListener('click', ex('bold'));
    byId('toolbar-italic')?.addEventListener('click', ex('italic'));
    byId('toolbar-underline')?.addEventListener('click', ex('underline'));
    byId('toolbar-strike')?.addEventListener('click', ex('strikethrough'));
    byId('toolbar-align-left')?.addEventListener('click', ex('alignLeft'));
    byId('toolbar-align-center')?.addEventListener('click', ex('alignCenter'));
    byId('toolbar-align-right')?.addEventListener('click', ex('alignRight'));
    byId('toolbar-valign')?.addEventListener('click', () => { if (luckysheetInitialized) ToolbarCommands.executeCommand('alignMiddle'); });
    byId('toolbar-merge')?.addEventListener('click', ex('merge'));
    const sizeEl = byId('toolbar-size');
    if (sizeEl) sizeEl.addEventListener('change', function () { if (luckysheetInitialized) ToolbarCommands.executeCommand('fontSize', parseInt(this.value, 10) || 10); });
    byId('toolbar-sum')?.addEventListener('click', () => {
      if (!luckysheetInitialized || !window.SpreadsheetEngine) return;
      const range = ToolbarCommands.getSelectionRange && ToolbarCommands.getSelectionRange();
      const activeTabId = getActiveSheetId();
      if (!range || !range.length || !activeTabId) return;
      const r0 = range[0].row ? range[0].row[0] : 0;
      const r1 = range[0].row ? range[0].row[1] : 0;
      const c0 = range[0].column ? range[0].column[0] : 0;
      const c1 = range[0].column ? range[0].column[1] : 0;
      const sheetName = window.SpreadsheetEngine.getTabNameById && SpreadsheetEngine.getTabNameById(activeTabId);
      const rangeRef = (sheetName ? sheetName + '!' : '') + colIndexToLetters(c0) + (r0 + 1) + ':' + colIndexToLetters(c1) + (r1 + 1);
      const formula = '=SUM(' + rangeRef + ')';
      const updates = SpreadsheetEngine.setCell(activeTabId, r0, c0, formula);
      applyEngineUpdatesToLuckysheet(updates);
    });
  }

  // Initialize
  tryRestoreSession();
  bindToolbarCommands();
})();

