/**
 * Single global spreadsheet engine.
 * - One cell store: key = sheetTabId + cellAddress (e.g. "123!A1" -> tabId 123, row 0, col 0)
 * - Cross-sheet formulas: =Sheet2!A1 + Sheet1!B3
 * - One dependency graph: when a cell changes, recompute all dependents across sheets
 * - Tab switch is UI-only: no reload, no re-fetch, no re-parse
 */
(function (global) {
  const cells = new Map(); // key -> { formula?, value, type? }
  const dependents = new Map(); // key -> Set of keys that depend on this key
  let documentId = null;
  let tabs = []; // { id, name, index, order, ... }

  function key(tabId, r, c) {
    return tabId + '!' + r + ':' + c;
  }

  function parseKey(k) {
    const ex = k.split('!');
    if (ex.length !== 2) return null;
    const [r, c] = ex[1].split(':').map(Number);
    return { tabId: ex[0], r, c };
  }

  function colLettersToIndex(letters) {
    let n = 0;
    const s = (letters || '').toUpperCase();
    for (let i = 0; i < s.length; i++) {
      n = n * 26 + (s.charCodeAt(i) - 64);
    }
    return n - 1; // 0-based
  }

  function rowNumberToIndex(numStr) {
    const n = parseInt(numStr, 10);
    return Number.isNaN(n) ? 0 : Math.max(0, n - 1); // 1-based in A1 -> 0-based
  }

  function getTabIdByName(sheetName) {
    if (!sheetName) return null;
    const name = String(sheetName).trim();
    const t = tabs.find((x) => (x.name || '').trim() === name);
    return t ? t.id : null;
  }

  function getTabNameById(tabId) {
    const t = tabs.find((x) => x.id === tabId);
    return t ? (t.name || 'Sheet') : null;
  }

  function resolveRef(sheetNameOrNull, colLetters, rowNum, currentTabId) {
    const tabId = sheetNameOrNull ? getTabIdByName(sheetNameOrNull) : currentTabId;
    if (tabId == null) return null;
    const r = rowNumberToIndex(rowNum);
    const c = colLettersToIndex(colLetters);
    return { tabId, r, c };
  }

  function keyFromRef(ref, currentTabId) {
    if (!ref) return null;
    const tid = ref.tabId != null ? ref.tabId : currentTabId;
    return key(tid, ref.r, ref.c);
  }

  // Parse formula: extract refs and ranges. Supports Sheet1!A1, A1, Sheet1!A1:B2, A1:B2, SUM(...), + - * /
  const REF_RE = /(?:([A-Za-z0-9\u4e00-\u9fa5]+)!)?([A-Z]+)([0-9]+)(?=\s*([:+\-*\/\)\,]|$))/gi;
  const SUM_RE = /SUM\s*\(\s*([^)]+)\s*\)/gi;

  function extractRefs(formula, currentTabId) {
    const refs = [];
    const seen = new Set();
    let m;
    REF_RE.lastIndex = 0;
    while ((m = REF_RE.exec(formula)) !== null) {
      const sheetName = m[1] || null;
      const col = m[2];
      const row = m[3];
      const ref = resolveRef(sheetName, col, row, currentTabId);
      if (ref) {
        const k = key(ref.tabId, ref.r, ref.c);
        if (!seen.has(k)) {
          seen.add(k);
          refs.push(ref);
        }
      }
    }
    return refs;
  }

  function extractRanges(formula, currentTabId) {
    const ranges = [];
    SUM_RE.lastIndex = 0;
    let m;
    while ((m = SUM_RE.exec(formula)) !== null) {
      const arg = m[1].trim();
      const colon = arg.indexOf(':');
      if (colon !== -1) {
        const left = arg.slice(0, colon).trim();
        const right = arg.slice(colon + 1).trim();
        const leftM = left.match(/(?:([A-Za-z0-9\u4e00-\u9fa5]+)!)?([A-Z]+)([0-9]+)$/i);
        const rightM = right.match(/(?:([A-Za-z0-9\u4e00-\u9fa5]+)!)?([A-Z]+)([0-9]+)$/i);
        if (leftM && rightM) {
          const r1 = resolveRef(leftM[1] || null, leftM[2], leftM[3], currentTabId);
          const r2 = resolveRef(rightM[1] || null, rightM[2], rightM[3], currentTabId);
          if (r1 && r2) ranges.push({ r1, r2 });
        }
      } else {
        const singleM = arg.match(/(?:([A-Za-z0-9\u4e00-\u9fa5]+)!)?([A-Z]+)([0-9]+)$/i);
        if (singleM) {
          const r = resolveRef(singleM[1] || null, singleM[2], singleM[3], currentTabId);
          if (r) ranges.push({ r1: r, r2: r });
        }
      }
    }
    return ranges;
  }

  function getAllRefKeys(formula, currentTabId) {
    const keys = new Set();
    extractRefs(formula, currentTabId).forEach((ref) => keys.add(key(ref.tabId, ref.r, ref.c)));
    extractRanges(formula, currentTabId).forEach(({ r1, r2 }) => {
      const minR = Math.min(r1.r, r2.r);
      const maxR = Math.max(r1.r, r2.r);
      const minC = Math.min(r1.c, r2.c);
      const maxC = Math.max(r1.c, r2.c);
      const tid = r1.tabId != null ? r1.tabId : currentTabId;
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          keys.add(key(tid, r, c));
        }
      }
    });
    return keys;
  }

  function getCellValue(tabId, r, c) {
    const k = key(tabId, r, c);
    const cell = cells.get(k);
    if (!cell) return null;
    const v = cell.value;
    if (typeof v === 'number' && !Number.isNaN(v)) return v;
    if (typeof v === 'string') return v;
    if (v == null) return '';
    return v;
  }

  function getCellValueByKey(k) {
    const p = parseKey(k);
    if (!p) return null;
    const tid = typeof p.tabId === 'string' && /^\d+$/.test(p.tabId) ? parseInt(p.tabId, 10) : getTabIdByName(p.tabId);
    if (tid == null) return null;
    return getCellValue(tid, p.r, p.c);
  }

  function toNumber(x) {
    if (typeof x === 'number') return Number.isNaN(x) ? 0 : x;
    if (x === null || x === undefined || x === '') return 0;
    const n = Number(x);
    return Number.isNaN(n) ? 0 : n;
  }

  function evaluateFormula(formula, currentTabId, visitedKeys) {
    if (!formula || typeof formula !== 'string') return null;
    const f = formula.trim();
    if (!f.startsWith('=')) return f;

    const expr = f.slice(1).trim();
    if (!expr) return null;

    const kCurrent = key(currentTabId, -1, -1);
    if (visitedKeys.has(kCurrent)) return '#CIRCULAR!';
    visitedKeys.add(kCurrent);

    try {
      // SUM(range) - range like A1:B2 or Sheet2!A1:B2
      const sumMatch = expr.match(/^\s*SUM\s*\(\s*([^)]+)\s*\)\s*$/i);
      if (sumMatch) {
        const arg = sumMatch[1].trim();
        const colon = arg.indexOf(':');
        let total = 0;
        if (colon !== -1) {
          const left = arg.slice(0, colon).trim().match(/(?:([A-Za-z0-9\u4e00-\u9fa5]+)!)?([A-Z]+)([0-9]+)/i);
          const right = arg.slice(colon + 1).trim().match(/(?:([A-Za-z0-9\u4e00-\u9fa5]+)!)?([A-Z]+)([0-9]+)/i);
          if (left && right) {
            const r1 = resolveRef(left[1] || null, left[2], left[3], currentTabId);
            const r2 = resolveRef(right[1] || null, right[2], right[3], currentTabId);
            if (r1 && r2) {
              const minR = Math.min(r1.r, r2.r);
              const maxR = Math.max(r1.r, r2.r);
              const minC = Math.min(r1.c, r2.c);
              const maxC = Math.max(r1.c, r2.c);
              const tid = r1.tabId != null ? r1.tabId : currentTabId;
              for (let row = minR; row <= maxR; row++) {
                for (let col = minC; col <= maxC; col++) {
                  total += toNumber(getCellValue(tid, row, col));
                }
              }
            }
          }
        } else {
          const one = arg.match(/(?:([A-Za-z0-9\u4e00-\u9fa5]+)!)?([A-Z]+)([0-9]+)/i);
          if (one) {
            const r = resolveRef(one[1] || null, one[2], one[3], currentTabId);
            if (r) total = toNumber(getCellValue(r.tabId, r.r, r.c));
          }
        }
        return total;
      }

      // Simple expression: ref, ref+ref, ref-ref, ref*ref, ref/ref, number
      const tokens = [];
      let rest = expr;
      const refOrNum = /^\s*((?:[A-Za-z0-9\u4e00-\u9fa5]+!)?[A-Z]+\d+|\d+(?:\.\d+)?)\s*([+\-*\/])?/i;
      while (rest.length) {
        const opMatch = rest.match(/^\s*([+\-*\/])\s*/);
        if (opMatch) {
          tokens.push({ type: 'op', value: opMatch[1] });
          rest = rest.slice(opMatch[0].length);
          continue;
        }
        const refMatch = rest.match(refOrNum);
        if (refMatch) {
          const val = refMatch[1];
          if (/^\d/.test(val)) {
            tokens.push({ type: 'num', value: parseFloat(val) });
          } else {
            const sheetCell = val.match(/(?:([A-Za-z0-9\u4e00-\u9fa5]+)!)?([A-Z]+)([0-9]+)/i);
            if (sheetCell) {
              const ref = resolveRef(sheetCell[1] || null, sheetCell[2], sheetCell[3], currentTabId);
              const v = ref ? getCellValue(ref.tabId, ref.r, ref.c) : null;
              tokens.push({ type: 'num', value: toNumber(v) });
            } else {
              tokens.push({ type: 'num', value: 0 });
            }
          }
          rest = rest.slice(refMatch[0].length);
          if (refMatch[2]) tokens.push({ type: 'op', value: refMatch[2] });
          continue;
        }
        break;
      }

      let acc = tokens.length && tokens[0].type === 'num' ? tokens[0].value : 0;
      for (let i = 1; i < tokens.length; i += 2) {
        if (tokens[i].type === 'op' && tokens[i + 1] && tokens[i + 1].type === 'num') {
          const n = tokens[i + 1].value;
          switch (tokens[i].value) {
            case '+': acc += n; break;
            case '-': acc -= n; break;
            case '*': acc *= n; break;
            case '/': acc = n === 0 ? 0 : acc / n; break;
            default: break;
          }
        }
      }
      return acc;
    } catch (e) {
      console.warn('Formula eval error', expr, e);
      return '#ERROR!';
    }
  }

  function removeDependentsOf(cellKey) {
    const set = dependents.get(cellKey);
    if (!set) return;
    set.forEach((depKey) => {
      const cell = cells.get(depKey);
      if (cell && cell.formula) {
        const refKeys = getAllRefKeys(cell.formula, null);
        refKeys.forEach((refKey) => {
          const s = dependents.get(refKey);
          if (s) s.delete(depKey);
        });
      }
    });
    dependents.set(cellKey, new Set());
  }

  function addDependencies(cellKey, refKeys) {
    refKeys.forEach((refKey) => {
      let set = dependents.get(refKey);
      if (!set) {
        set = new Set();
        dependents.set(refKey, set);
      }
      set.add(cellKey);
    });
  }

  function recomputeDependents(cellKey, currentTabId, updated, visited) {
    if (visited.has(cellKey)) return;
    visited.add(cellKey);
    const set = dependents.get(cellKey);
    if (!set) return;
    const toRecompute = Array.from(set);
    toRecompute.forEach((depKey) => {
      const cell = cells.get(depKey);
      if (!cell || !cell.formula) return;
      const parsed = parseKey(depKey);
      if (!parsed) return;
      const tabId = typeof parsed.tabId === 'string' && /^\d+$/.test(parsed.tabId) ? parseInt(parsed.tabId, 10) : getTabIdByName(parsed.tabId);
      if (tabId == null) return;
      const newVal = evaluateFormula(cell.formula, tabId, new Set());
      const prev = cell.value;
      cell.value = newVal;
      updated.push({ tabId, r: parsed.r, c: parsed.c, value: newVal });
      recomputeDependents(depKey, tabId, updated, visited);
    });
  }

  function setDocumentId(id) {
    documentId = id;
  }

  function setTabs(tabsArray) {
    tabs = Array.isArray(tabsArray) ? tabsArray : [];
  }

  function loadFromTabs(tabsData) {
    setTabs(tabsData || []);
    cells.clear();
    dependents.clear();
    (tabsData || []).forEach((tab) => {
      const tabId = tab.id;
      const celldata = tab.celldata || [];
      celldata.forEach((item) => {
        const r = item.r;
        const c = item.c;
        let raw = item.v;
        if (raw && typeof raw === 'object') {
          raw = raw.v != null ? raw.v : raw.m;
        }
        if (raw == null) return;
        const k = key(tabId, r, c);
        const isFormula = typeof raw === 'string' && raw.trim().startsWith('=');
        const value = isFormula ? evaluateFormula(raw, tabId, new Set()) : raw;
        cells.set(k, { formula: isFormula ? String(raw) : undefined, value });
      });
    });
    // Build dependency graph for formula cells
    cells.forEach((cell, k) => {
      if (!cell.formula) return;
      const parsed = parseKey(k);
      if (!parsed) return;
      const tabId = typeof parsed.tabId === 'string' && /^\d+$/.test(parsed.tabId) ? parseInt(parsed.tabId, 10) : null;
      const refKeys = getAllRefKeys(cell.formula, tabId);
      addDependencies(k, refKeys);
    });
  }

  function toLuckysheetTabs() {
    return tabs.map((tab) => {
      const tabId = tab.id;
      const tabIdStr = String(tabId);
      const celldata = [];
      cells.forEach((cell, k) => {
        const p = parseKey(k);
        if (!p) return;
        if (String(p.tabId) !== tabIdStr) return;
        const v = cell.value;
        const display = v == null ? '' : (typeof v === 'object' ? JSON.stringify(v) : String(v));
        celldata.push({
          r: p.r,
          c: p.c,
          v: { v: v, m: display }
        });
      });
      return {
        id: tab.id,
        name: tab.name || 'Sheet',
        index: tab.index != null ? tab.index : 0,
        order: tab.order != null ? tab.order : 0,
        status: tab.status != null ? tab.status : 0,
        row: tab.row != null ? tab.row : 100,
        column: tab.column != null ? tab.column : 26,
        celldata,
        config: tab.config || {}
      };
    });
  }

  function getCell(tabId, r, c) {
    const k = key(tabId, r, c);
    return cells.get(k);
  }

  function setCell(tabId, r, c, input) {
    const k = key(tabId, r, c);
    const updates = [{ tabId, r, c, value: null }];

    const isFormula = typeof input === 'string' && input.trim().startsWith('=');
    removeDependentsOf(k);

    if (isFormula) {
      const formula = input.trim();
      const refKeys = getAllRefKeys(formula, tabId);
      addDependencies(k, refKeys);
      const value = evaluateFormula(formula, tabId, new Set());
      cells.set(k, { formula, value });
      updates[0].value = value;
      const visited = new Set();
      recomputeDependents(k, tabId, updates, visited);
    } else {
      const value = input != null && typeof input === 'object' && (input.v !== undefined || input.m !== undefined)
        ? (input.v != null ? input.v : input.m)
        : input;
      cells.set(k, { value });
      updates[0].value = value;
      const visited = new Set();
      recomputeDependents(k, tabId, updates, visited);
    }

    return updates;
  }

  function getDocumentId() {
    return documentId;
  }

  function getTabs() {
    return tabs.slice();
  }

  function removeTab(tabId) {
    const idStr = String(tabId);
    tabs = tabs.filter((t) => t.id !== tabId && String(t.id) !== idStr);
    const keysToDelete = [];
    cells.forEach((_, k) => {
      if (k.startsWith(idStr + '!')) keysToDelete.push(k);
    });
    keysToDelete.forEach((k) => {
      removeDependentsOf(k);
      cells.delete(k);
      dependents.delete(k);
    });
  }

  function reset() {
    documentId = null;
    tabs = [];
    cells.clear();
    dependents.clear();
  }

  global.SpreadsheetEngine = {
    key,
    getCell,
    setCell,
    getCellValue,
    getCellValueByKey,
    setDocumentId,
    setTabs,
    getTabIdByName,
    getTabNameById,
    loadFromTabs,
    toLuckysheetTabs,
    getDocumentId,
    getTabs,
    removeTab,
    reset,
    extractRefs,
    evaluateFormula
  };
})(typeof window !== 'undefined' ? window : this);
