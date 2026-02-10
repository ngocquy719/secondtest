/**
 * SheetManager - Single source of truth for sheet tabs (in-memory).
 * 1 file = N sheet tabs. Switch = only change activeSheetId. NO reload, NO refetch, NO socket reconnect.
 */
(function (global) {
  const state = {
    documentId: null,
    sheets: [],
    activeSheetId: null,
    permission: null
  };

  const listeners = [];

  function emit() {
    listeners.forEach((fn) => fn(state));
  }

  function getState() {
    return {
      documentId: state.documentId,
      sheets: state.sheets.slice(),
      activeSheetId: state.activeSheetId,
      permission: state.permission
    };
  }

  function getOrderById(tabId) {
    return state.sheets.findIndex((s) => s.id === tabId);
  }

  function setDocument(documentId, sheets, permission) {
    state.documentId = documentId;
    state.sheets = Array.isArray(sheets) ? sheets : [];
    state.activeSheetId = state.sheets.length ? state.sheets[0].id : null;
    state.permission = permission || null;
    emit();
  }

  function switchSheet(tabId) {
    if (state.activeSheetId === tabId) return;
    const idx = getOrderById(tabId);
    if (idx < 0) return;
    state.activeSheetId = tabId;
    emit();
  }

  function addSheet(sheetObject) {
    if (!sheetObject || !sheetObject.id) return;
    state.sheets.push(sheetObject);
    state.activeSheetId = sheetObject.id;
    emit();
  }

  function removeSheet(tabId) {
    const idx = getOrderById(tabId);
    if (idx < 0) return;
    state.sheets.splice(idx, 1);
    if (state.activeSheetId === tabId) {
      state.activeSheetId = state.sheets.length ? state.sheets[0].id : null;
    }
    emit();
  }

  function renameSheet(tabId, name) {
    const s = state.sheets.find((x) => x.id === tabId);
    if (s) {
      s.name = name;
      emit();
    }
  }

  function setSheets(sheets) {
    state.sheets = Array.isArray(sheets) ? sheets : [];
    if (state.activeSheetId && getOrderById(state.activeSheetId) < 0) {
      state.activeSheetId = state.sheets.length ? state.sheets[0].id : null;
    }
    emit();
  }

  function subscribe(fn) {
    listeners.push(fn);
    return () => {
      const i = listeners.indexOf(fn);
      if (i >= 0) listeners.splice(i, 1);
    };
  }

  function reset() {
    state.documentId = null;
    state.sheets = [];
    state.activeSheetId = null;
    state.permission = null;
    emit();
  }

  global.SheetManager = {
    getState,
    getOrderById,
    setDocument,
    switchSheet,
    addSheet,
    removeSheet,
    renameSheet,
    setSheets,
    subscribe,
    reset
  };
})(typeof window !== 'undefined' ? window : this);
