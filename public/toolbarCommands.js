/**
 * Toolbar commands - executeCommand(name, payload). All actions call Luckysheet API on current selection.
 */
(function (global) {
  function getSelectionRange() {
    if (typeof luckysheet === 'undefined' || !luckysheet.getRange) return null;
    const range = luckysheet.getRange();
    if (!range || !range.length) return null;
    return range;
  }

  function forEachCellInRange(range, fn) {
    if (!range || !range.length) return;
    range.forEach((r) => {
      const [r0, r1] = r.row || [0, 0];
      const [c0, c1] = r.column || [0, 0];
      for (let row = r0; row <= r1; row++) {
        for (let col = c0; col <= c1; col++) {
          fn(row, col);
        }
      }
    });
  }

  function executeCommand(commandName, payload) {
    if (typeof luckysheet === 'undefined') {
      console.error('ToolbarCommands: luckysheet not loaded');
      return;
    }
    const range = getSelectionRange();
    if (!range) return;

    const applyToSelection = (fn) => {
      forEachCellInRange(range, (r, c) => {
        try {
          fn(r, c);
        } catch (e) {
          console.warn('ToolbarCommands: apply failed for cell', r, c, commandName, e);
        }
      });
    };

    switch (commandName) {
      case 'bold':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'bl', 1));
        break;
      case 'italic':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'it', 1));
        break;
      case 'underline':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'ul', 1));
        break;
      case 'strikethrough':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'st', 1));
        break;
      case 'alignLeft':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'ht', 0));
        break;
      case 'alignCenter':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'ht', 1));
        break;
      case 'alignRight':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'ht', 2));
        break;
      case 'alignTop':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'vt', 0));
        break;
      case 'alignMiddle':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'vt', 1));
        break;
      case 'alignBottom':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'vt', 2));
        break;
      case 'fontSize':
        applyToSelection((r, c) => luckysheet.setCellFormat(r, c, 'fs', payload != null ? payload : 10));
        break;
      case 'merge':
        if (range.length && range[0].row && range[0].column) {
          const [r0, r1] = range[0].row;
          const [c0, c1] = range[0].column;
          if (r1 > r0 || c1 > c0) {
            try {
              luckysheet.setCellFormat(r0, c0, 'merge', { row: [r0, r1], column: [c0, c1] });
            } catch (e) {
              console.warn('ToolbarCommands: merge failed', e);
            }
          }
        }
        break;
      case 'clearFormat':
        applyToSelection((r, c) => {
          try {
            luckysheet.setCellFormat(r, c, 'bl', 0);
            luckysheet.setCellFormat(r, c, 'it', 0);
            luckysheet.setCellFormat(r, c, 'ul', 0);
            luckysheet.setCellFormat(r, c, 'ht', 1);
            luckysheet.setCellFormat(r, c, 'vt', 1);
          } catch (e) {}
        });
        break;
      default:
        break;
    }
  }

  global.ToolbarCommands = { executeCommand, getSelectionRange };
})(typeof window !== 'undefined' ? window : this);
