(() => {
  if (!window.auth.requireAuth()) return;

  const apiBase = '/api';
  let currentUser = null;
  let knownUsers = [];

  const userInfoEl = document.getElementById('user-info');
  const sheetsListEl = document.getElementById('sheets-list');
  const sheetsTableBody = document.querySelector('#sheets-table tbody');
  const createSheetBtn = document.getElementById('create-sheet-btn');
  const usersSection = document.getElementById('users-section');
  const userForm = document.getElementById('user-form');
  const userIdInput = document.getElementById('user-id');
  const userUsernameInput = document.getElementById('user-username');
  const userPasswordInput = document.getElementById('user-password');
  const userRoleSelect = document.getElementById('user-role');
  const userSaveBtn = document.getElementById('user-save-btn');
  const userCancelBtn = document.getElementById('user-cancel-btn');
  const userMessage = document.getElementById('user-message');
  const usersListEl = document.getElementById('users-list');

  async function apiRequest(path, options = {}) {
    const headers = options.headers || {};
    Object.assign(headers, window.auth.authHeaders());
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
    if (currentUser.role === 'user') {
      usersSection.classList.add('hidden');
    } else {
      usersSection.classList.remove('hidden');
      if (currentUser.role === 'leader') {
        userRoleSelect.value = 'user';
        userRoleSelect.disabled = true;
      } else {
        userRoleSelect.disabled = false;
      }
      loadUsersList();
    }
  }

  async function loadSheets() {
    const data = await apiRequest('/sheets');
    sheetsListEl.innerHTML = '';
    sheetsTableBody.innerHTML = '';

    data.sheets.forEach((s) => {
      const li = document.createElement('li');
      li.textContent = s.name || 'Sheet';
      const perm = document.createElement('span');
      perm.className = 'sheet-permission';
      perm.textContent = s.permission;
      li.appendChild(perm);
      li.addEventListener('click', () => {
        window.location.href = `/sheet.html?id=${s.id}`;
      });
      sheetsListEl.appendChild(li);

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${s.id}</td>
        <td>${s.name || 'Sheet'}</td>
        <td>${s.permission}</td>
        <td><button data-open-id="${s.id}">Open</button></td>
      `;
      sheetsTableBody.appendChild(tr);
    });

    sheetsTableBody
      .querySelectorAll('button[data-open-id]')
      .forEach((btn) => {
        btn.addEventListener('click', () => {
          const id = btn.getAttribute('data-open-id');
          window.location.href = `/sheet.html?id=${id}`;
        });
      });
  }

  async function loadUsersList() {
    if (!currentUser || (currentUser.role !== 'admin' && currentUser.role !== 'leader')) {
      usersListEl.innerHTML = '';
      return;
    }
    try {
      const data = await apiRequest('/users');
      knownUsers = data.users || [];
      usersListEl.innerHTML = '';
      knownUsers.forEach((u) => {
        const li = document.createElement('li');
        const left = document.createElement('div');
        const right = document.createElement('div');
        right.className = 'user-actions';

        left.textContent = u.username;
        const meta = document.createElement('span');
        meta.className = 'user-meta';
        meta.textContent = `(${u.role}) id:${u.id}`;
        left.appendChild(meta);

        const editBtn = document.createElement('button');
        editBtn.textContent = 'Edit';
        editBtn.addEventListener('click', () => {
          startEditUser(u);
        });

        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = 'Remove';
        deleteBtn.addEventListener('click', async () => {
          if (!confirm(`Remove user "${u.username}"?`)) return;
          try {
            await apiRequest(`/users/${u.id}`, { method: 'DELETE' });
            await loadUsersList();
          } catch (err) {
            userMessage.textContent = err.message;
          }
        });

        right.appendChild(editBtn);
        right.appendChild(deleteBtn);

        li.appendChild(left);
        li.appendChild(right);
        usersListEl.appendChild(li);
      });
    } catch (err) {
      console.error('Failed to load users', err);
    }
  }

  function resetUserForm() {
    userIdInput.value = '';
    userUsernameInput.value = '';
    userPasswordInput.value = '';
    userMessage.textContent = '';
    if (currentUser?.role === 'admin') {
      userRoleSelect.value = 'leader';
    } else {
      userRoleSelect.value = 'user';
    }
    userSaveBtn.textContent = 'Create';
  }

  function startEditUser(user) {
    userIdInput.value = String(user.id);
    userUsernameInput.value = user.username;
    userPasswordInput.value = '';
    userRoleSelect.value = user.role;
    if (currentUser?.role === 'leader') {
      userRoleSelect.value = 'user';
      userRoleSelect.disabled = true;
    } else {
      userRoleSelect.disabled = false;
    }
    userSaveBtn.textContent = 'Update';
  }

  createSheetBtn.addEventListener('click', async () => {
    try {
      const name = prompt('Sheet name?', 'Sheet1');
      const data = await apiRequest('/sheets', {
        method: 'POST',
        body: { name }
      });
      await loadSheets();
      if (data.sheet && data.sheet.id) {
        window.location.href = `/sheet.html?id=${data.sheet.id}`;
      }
    } catch (err) {
      alert(err.message);
    }
  });

  userForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    userMessage.textContent = '';

    const id = userIdInput.value ? Number(userIdInput.value) : null;
    const username = userUsernameInput.value.trim();
    const password = userPasswordInput.value;
    const role = userRoleSelect.value;

    if (!username && !id) {
      userMessage.textContent = 'Username is required for new users.';
      return;
    }

    try {
      if (!id) {
        await apiRequest('/users', {
          method: 'POST',
          body: { username, password, role }
        });
        userMessage.textContent = 'User created.';
      } else {
        const body = {};
        if (username) body.username = username;
        if (password) body.password = password;
        if (role) body.role = role;
        await apiRequest(`/users/${id}`, {
          method: 'PUT',
          body
        });
        userMessage.textContent = 'User updated.';
      }
      resetUserForm();
      loadUsersList();
    } catch (err) {
      userMessage.textContent = err.message;
    }
  });

  userCancelBtn.addEventListener('click', () => {
    resetUserForm();
  });

  // Init
  (async () => {
    try {
      await loadCurrentUser();
      await loadSheets();
    } catch (err) {
      console.error(err);
      window.auth.clearToken();
      window.location.href = '/';
    }
  })();
})();

