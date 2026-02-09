(() => {
  const form = document.getElementById('login-form');
  const usernameInput = document.getElementById('username');
  const passwordInput = document.getElementById('password');
  const errorEl = document.getElementById('login-error');

  async function apiLogin(username, password) {
    const res = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    if (!res.ok) {
      let msg = 'Login failed';
      try {
        const data = await res.json();
        if (data && data.error) msg = data.error;
      } catch (_) {}
      throw new Error(msg);
    }
    return res.json();
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    errorEl.textContent = '';
    try {
      const data = await apiLogin(
        usernameInput.value.trim(),
        passwordInput.value
      );
      window.auth.setToken(data.token);
      window.location.href = '/home.html';
    } catch (err) {
      errorEl.textContent = err.message;
    }
  });
})();

