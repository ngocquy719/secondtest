// Simple auth helper shared across pages
(() => {
  const STORAGE_KEY = 'jwt';

  function getToken() {
    return localStorage.getItem(STORAGE_KEY) || null;
  }

  function setToken(token) {
    if (token) {
      localStorage.setItem(STORAGE_KEY, token);
    } else {
      localStorage.removeItem(STORAGE_KEY);
    }
  }

  function clearToken() {
    localStorage.removeItem(STORAGE_KEY);
  }

  function authHeaders() {
    const token = getToken();
    return token ? { Authorization: `Bearer ${token}` } : {};
  }

  function requireAuth() {
    const token = getToken();
    if (!token) {
      window.location.href = '/';
      return false;
    }
    return true;
  }

  window.auth = {
    getToken,
    setToken,
    clearToken,
    authHeaders,
    requireAuth
  };
})();

