// ============================================================================
// APP CONTROLLER - SISTEMA EXCEL ORDENES WHOLESALE (Pacifica Produce Farms)
// ============================================================================

(function () {
  const { colToIndex, indexToCol, parseCoord } = window.ExcelUtils;

  let workbook = null;
  let engine = null;
  let activeSheet = 'PO Nuevo';
  let selectedCoord = 'F11';
  let selectionEndCoord = 'F11';
  let isDraggingSelection = false;
  let isDraggingFillHandle = false;
  let fillPreviewTargetCoord = null;
  let clipboardState = null; // { sheetName, minRow, maxRow, minCol, maxCol, isCut, grid, tsv }
  let inlineEditor = null;
  let zoomLevel = 80;
  let searchQuery = '';
  let emailSendMode = null;

  // Filters for DayOrders
  let dayFilterDate = 'ALL';
  let dayFilterClient = 'ALL';

  // Undo / Redo history
  const undoStack = [];
  const redoStack = [];
  const MAX_HISTORY = 35;

  // Protected formula columns in PO Nuevo rows 22..41 so pressing Delete clears input without breaking formulas
  const PO_FORMULA_COLS = new Set(['A', 'C', 'D', 'E', 'F', 'G', 'J']);
  const PO_HEADER_FORMULA_COORDS = new Set(['E12', 'F15', 'F16', 'B18', 'B19', 'H42', 'I42']);

  // DOM References
  const gridHead = document.getElementById('excelGridHead');
  const gridBody = document.getElementById('excelGridBody');
  const sheetCanvasWrapper = document.getElementById('sheetCanvasWrapper');
  const cellNameBox = document.getElementById('cellNameBox');
  const formulaInput = document.getElementById('formulaInput');
  const cellMetaBadge = document.getElementById('cellMetaBadge');
  const saveStatusBadge = document.getElementById('saveStatusBadge');
  const validationDropdown = document.getElementById('validationDropdown');
  const validationSearchInput = document.getElementById('validationSearchInput');
  const validationList = document.getElementById('validationList');
  const validationCountBadge = document.getElementById('validationCountBadge');
  const dayOrdersFilterGroup = document.getElementById('dayOrdersFilterGroup');
  const filterDayOrdersDate = document.getElementById('filterDayOrdersDate');
  const filterDayOrdersClient = document.getElementById('filterDayOrdersClient');

  // Initialize App
  async function init() {
    try {
      let res = await fetch('/api/workbook');
      if (!res.ok) {
        res = await fetch('data/initial_workbook.json');
      }
      workbook = await res.json();
    } catch (err) {
      const localBackup = localStorage.getItem('wholesale_workbook_backup');
      if (localBackup) {
        workbook = JSON.parse(localBackup);
      } else {
        const res = await fetch('data/initial_workbook.json');
        workbook = await res.json();
      }
    }

    const formulasMigrated = migratePONuevoDefaultFormulas();
    engine = new window.FormulaEngine(workbook);
    engine.recalculateAll();
    if (formulasMigrated) saveWorkbookState();

    sheetCanvasWrapper.style.zoom = `${zoomLevel}%`;
    const zLbl = document.getElementById('zoomLabel');
    if (zLbl) zLbl.textContent = `${zoomLevel}%`;

    populateDayOrdersFilters();
    renderSheet();
    selectCell('F11', false);
    bindEvents();
    updateStatusMetrics();
  }

  function pushUndoState() {
    const snap = JSON.stringify(workbook.sheets);
    if (undoStack.length > 0 && undoStack[undoStack.length - 1] === snap) return;
    undoStack.push(snap);
    if (undoStack.length > MAX_HISTORY) undoStack.shift();
    redoStack.length = 0;
    updateUndoRedoButtons();
  }

  function updateUndoRedoButtons() {
    document.getElementById('btnUndo').disabled = undoStack.length === 0;
    document.getElementById('btnRedo').disabled = redoStack.length === 0;
  }

  function undo() {
    if (undoStack.length === 0) return;
    redoStack.push(JSON.stringify(workbook.sheets));
    workbook.sheets = JSON.parse(undoStack.pop());
    engine = new window.FormulaEngine(workbook);
    engine.recalculateAll();
    updateUndoRedoButtons();
    renderSheet();
    selectCell(selectedCoord, false);
    saveWorkbookState();
    showToast('Acción deshecha (Ctrl+Z)');
  }

  function redo() {
    if (redoStack.length === 0) return;
    undoStack.push(JSON.stringify(workbook.sheets));
    workbook.sheets = JSON.parse(redoStack.pop());
    engine = new window.FormulaEngine(workbook);
    engine.recalculateAll();
    updateUndoRedoButtons();
    renderSheet();
    selectCell(selectedCoord, false);
    saveWorkbookState();
    showToast('Acción rehecha (Ctrl+Y)');
  }

  let saveTimeout = null;
  function saveWorkbookState() {
    if (saveStatusBadge) saveStatusBadge.textContent = '● Guardando...';
    clearTimeout(saveTimeout);
    saveTimeout = setTimeout(async () => {
      try {
        localStorage.setItem('wholesale_workbook_backup', JSON.stringify(workbook));
        const res = await fetch('/api/workbook', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(workbook)
        });
        if (res.ok && saveStatusBadge) {
          saveStatusBadge.textContent = '● Guardado automáticamente';
        }
      } catch (e) {
        if (saveStatusBadge) saveStatusBadge.textContent = '● Guardado en navegador';
      }
    }, 250);
  }

  // Formatters
  function formatDateDisplay(val) {
    if (!val) return '';
    const s = String(val).trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) {
      return `${m[2]}/${m[3]}/${m[1].slice(2)}`;
    }
    return s;
  }

  function formatCurrency(val) {
    if (val === '' || val === null || val === undefined) return '';
    const n = Number(val);
    if (isNaN(n)) return String(val);
    return '$ ' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function isBomPendingStatus(value) {
    const status = String(value || '').trim().toUpperCase();
    return status === 'PENDIENTE DE BOM' || status === 'PENDIENTE DE BOOM';
  }

  function getPendingBomRows() {
    const ws = workbook?.sheets?.['PO Nuevo'];
    if (!ws) return [];
    const lastProdRow = getPONuevoTotalRow() - 1;
    const pending = [];
    for (let r = 22; r <= lastProdRow; r += 1) {
      const flower = ws.cells[`B${r}`]?.v;
      if (!flower || !isBomPendingStatus(ws.cells[`J${r}`]?.v)) continue;
      pending.push({
        row: r,
        product: flower,
        recommendation: [
          flower,
          ws.cells[`C${r}`]?.v,
          ws.cells[`D${r}`]?.v ? `STEM/${ws.cells[`D${r}`].v}` : '',
          ws.cells[`E${r}`]?.v ? `BU/${ws.cells[`E${r}`].v}` : '',
          ws.cells[`F${r}`]?.v
        ].filter(Boolean).join(' ') || 'Producto Wholesale'
      });
    }
    return pending;
  }

  function getPONuevoTotalRow() {
    const ws = workbook?.sheets?.['PO Nuevo'];
    if (!ws) return 42;
    if (ws.poTotalRow && ws.poTotalRow >= 23) return ws.poTotalRow;
    for (let r = 23; r <= 300; r++) {
      if (String(ws.cells[`G${r}`]?.v || '').trim().toUpperCase() === 'TOTAL') {
        ws.poTotalRow = r;
        return r;
      }
    }
    return 42;
  }

  function formatCellDisplay(sheetName, coord, rawVal, colStr, row) {
    if (rawVal === null || rawVal === undefined || rawVal === '') return '';

    if (sheetName === 'PO Nuevo') {
      const totalRow = getPONuevoTotalRow();
      if (coord === 'B19' || coord === 'F19') return formatDateDisplay(rawVal);
      if (colStr === 'I' && row >= 22 && row <= totalRow) return formatCurrency(rawVal);
    } else if (sheetName === 'DayOrders' && row >= 2) {
      if (colStr === 'F') return formatDateDisplay(rawVal);
      if (colStr === 'O' || colStr === 'P') return formatCurrency(rawVal);
    } else if (sheetName === 'PASSING SHEET') {
      if (coord === 'B6' || coord === 'B7') return formatDateDisplay(rawVal);
    }
    return String(rawVal);
  }

  // Determine merged cell map for active sheet
  function buildMergeMap(sheet) {
    const skipCells = new Set();
    const spanMap = {};

    const merges = (sheet.merges || []).filter(rng => rng !== 'J10:L11');
    if (activeSheet === 'PO Nuevo') {
      merges.push('A1:B4');
    }

    for (const rng of merges) {
      const parts = rng.split(':');
      if (parts.length !== 2) continue;
      const start = parseCoord(parts[0]);
      const end = parseCoord(parts[1]);
      if (!start || !end) continue;
      const colspan = end.col - start.col + 1;
      const rowspan = end.row - start.row + 1;
      spanMap[start.coord] = { colspan, rowspan };
      for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
          const coord = `${indexToCol(c)}${r}`;
          if (coord !== start.coord) skipCells.add(coord);
        }
      }
    }
    return { skipCells, spanMap };
  }

  function isValidationCell(sheetName, coord) {
    if (sheetName !== 'PO Nuevo') return null;
    if (coord === 'F11') return 'CUSTOMER';
    const parsed = parseCoord(coord);
    const lastProdRow = getPONuevoTotalRow() - 1;
    if (parsed && parsed.colStr === 'B' && parsed.row >= 22 && parsed.row <= lastProdRow) {
      return 'PRODUCT';
    }
    return null;
  }

  function isDatePickerCell(sheetName, coord) {
    if (sheetName === 'PO Nuevo' && coord === 'F19') return true;
    const parsed = parseCoord(coord);
    if (sheetName === 'DayOrders' && parsed && parsed.colStr === 'F' && parsed.row >= 2) return true;
    return false;
  }

  function getCellClasses(sheetName, coord, colStr, row, val) {
    const classes = [];

    if (sheetName === 'PO Nuevo') {
      const totalRow = getPONuevoTotalRow();
      const lastProdRow = totalRow - 1;
      if (['A8', 'A9', 'E11', 'B19', 'F19'].includes(coord)) classes.push('po-yellow-cell');
      if (coord === 'A5') classes.push('po-title-brand');
      if (['A11', 'E11', 'A15', 'E15', 'E16', 'A18', 'D18', 'A19', 'D19'].includes(coord)) {
        classes.push('po-label-bold');
      }
      if (row === 21 && colStr >= 'A' && colStr <= 'I') {
        if (['B', 'H', 'I'].includes(colStr)) classes.push('po-header-yellow');
        else classes.push('po-header-blue');
      }
      if (row >= 22 && row <= lastProdRow && colStr >= 'A' && colStr <= 'I') {
        classes.push('po-table-line');
        if (['D', 'E', 'F', 'G', 'H'].includes(colStr)) classes.push('center-cell');
        if (colStr === 'I') classes.push('num-cell');
      }
      if (row === totalRow && colStr >= 'A' && colStr <= 'I') {
        classes.push('po-total-row');
        if (colStr === 'H') classes.push('center-cell');
        if (colStr === 'I') classes.push('num-cell');
      }
      if (colStr === 'J' && row >= 22 && row <= lastProdRow && val === 'PENDIENTE DE BOM') {
        classes.push('po-bom-warning');
      }
    } else if (sheetName === 'DayOrders') {
      if (row === 1) {
        if (colStr === 'P') classes.push('dayorders-header-green');
        else classes.push('dayorders-header-blue');
      } else {
        if (colStr === 'P') classes.push('dayorders-final-price');
        else if (['J', 'K', 'N', 'O'].includes(colStr)) classes.push('num-cell');
        else if (['F', 'L', 'M'].includes(colStr)) classes.push('center-cell');
      }
    } else if (sheetName === 'Productos' || sheetName === 'Transporte') {
      if (row === 1) classes.push('catalog-header');
      if (sheetName === 'Transporte' && row >= 2 && colStr >= 'B' && colStr <= 'H') {
        classes.push('center-cell');
      }
    } else if (sheetName === 'PASSING SHEET') {
      if (coord === 'A1') classes.push('po-title-brand');
      if (colStr === 'A' && row >= 4 && row <= 12) classes.push('po-label-bold');
    }

    if (searchQuery && val !== '' && String(val).toLowerCase().includes(searchQuery)) {
      classes.push('search-match');
    }

    return classes.join(' ');
  }

  // Render Active Sheet Grid
  function renderSheet() {
    closeInlineEditor(false);
    hideValidationDropdown();

    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;

    const lastReal = getLastRealRow(activeSheet);
    sheet.maxCol = Math.max(sheet.maxCol || 16, 16);
    sheet.maxRow = Math.max(sheet.maxRow || 100, lastReal + 40, 100);

    const maxCol = sheet.maxCol;
    const maxRow = sheet.maxRow;
    const { skipCells, spanMap } = buildMergeMap(sheet);

    // 1. Render Column Headers
    let headHtml = '<tr><th class="corner-header" title="Seleccionar toda la hoja">◢</th>';
    for (let c = 1; c <= maxCol; c++) {
      const colStr = indexToCol(c);
      const w = Math.max(65, Math.round((sheet.colWidths[colStr] || 14) * 7.5));
      headHtml += `<th data-col="${colStr}" style="width:${w}px;min-width:${w}px;">${colStr}</th>`;
    }
    headHtml += '</tr>';
    gridHead.innerHTML = headHtml;

    // 2. Render Rows
    const rowsHtml = [];
    for (let r = 1; r <= maxRow; r++) {
      // Apply DayOrders filtering on rows >= 2 when a filter is active
      if (activeSheet === 'DayOrders' && r >= 2) {
        const hasAny = ['A', 'B', 'C', 'F', 'H', 'N', 'O'].some(col => {
          const cell = sheet.cells[`${col}${r}`];
          return cell && cell.v !== '' && cell.v !== null && cell.v !== undefined;
        });

        if (dayFilterDate !== 'ALL' && hasAny) {
          const rDate = sheet.cells[`F${r}`]?.v || '';
          if (rDate !== dayFilterDate) continue;
        }
        if (dayFilterClient !== 'ALL' && hasAny) {
          const rClient = sheet.cells[`B${r}`]?.v || '';
          if (rClient !== dayFilterClient) continue;
        }
      }

      let tr = `<tr data-row="${r}"><th class="row-header" data-row="${r}">${r}</th>`;

      for (let c = 1; c <= maxCol; c++) {
        const colStr = indexToCol(c);
        const coord = `${colStr}${r}`;
        if (skipCells.has(coord)) continue;

        const span = spanMap[coord];
        const spanAttrs = span
          ? `${span.colspan > 1 ? ` colspan="${span.colspan}"` : ''}${span.rowspan > 1 ? ` rowspan="${span.rowspan}"` : ''}`
          : '';

        const cellObj = sheet.cells[coord];
        let rawVal = cellObj ? cellObj.v : '';

        // Special logo in PO Nuevo
        if (activeSheet === 'PO Nuevo') {
          if (coord === 'J10' && rawVal === 'Ejecutar Generar Orden') {
            rawVal = '';
          }
          if (coord === 'A1') {
            tr += `<td data-coord="A1"${spanAttrs} style="background:#ffffff;padding:4px 8px;">
              <div class="po-logo-widget">
                <img src="logo.png" alt="Pacifica Produce Farms" onerror="this.style.display='none'" />
              </div>
            </td>`;
            continue;
          }
        }

        const displayVal = formatCellDisplay(activeSheet, coord, rawVal, colStr, r);
        const cls = getCellClasses(activeSheet, coord, colStr, r, rawVal);
        const valType = isValidationCell(activeSheet, coord);
        const isDateCell = isDatePickerCell(activeSheet, coord);
        const triggerHtml = valType
          ? `<span class="cell-dropdown-trigger" data-dropdown="${coord}" title="Abrir lista desplegable">▾</span>`
          : (isDateCell
              ? `<span class="cell-dropdown-trigger" data-datepicker="${coord}" title="Elegir fecha en calendario">📅</span>`
              : '');

        tr += `<td data-coord="${coord}"${spanAttrs}${cls ? ` class="${cls}"` : ''}>${escapeHtml(displayVal)}${triggerHtml}</td>`;
      }

      tr += '</tr>';
      rowsHtml.push(tr);
    }

    gridBody.innerHTML = rowsHtml.join('');
    highlightSelectedCellInDOM();
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function getSelectionBounds(startC = selectedCoord, endC = selectionEndCoord) {
    const p1 = parseCoord(startC) || parseCoord('A1');
    const p2 = parseCoord(endC || startC) || p1;
    const minRow = Math.min(p1.row, p2.row);
    const maxRow = Math.max(p1.row, p2.row);
    const minCol = Math.min(p1.col, p2.col);
    const maxCol = Math.max(p1.col, p2.col);
    return {
      minRow,
      maxRow,
      minCol,
      maxCol,
      isMulti: minRow !== maxRow || minCol !== maxCol,
      startParsed: p1,
      endParsed: p2
    };
  }

  // Cell Selection & Formula Bar Sync
  function selectCell(coord, openDropdownIfList = false, extendSelection = false) {
    const parsed = parseCoord(coord);
    if (!parsed) return;

    if (extendSelection) {
      selectionEndCoord = parsed.coord;
    } else {
      selectedCoord = parsed.coord;
      selectionEndCoord = parsed.coord;
    }

    const bounds = getSelectionBounds();
    if (bounds.isMulti) {
      cellNameBox.value = `${indexToCol(bounds.minCol)}${bounds.minRow}:${indexToCol(bounds.maxCol)}${bounds.maxRow}`;
    } else {
      cellNameBox.value = selectedCoord;
    }

    const anchorParsed = parseCoord(selectedCoord);
    const sheet = workbook.sheets[activeSheet];
    const cellObj = sheet?.cells[selectedCoord];

    if (cellObj && cellObj.f) {
      formulaInput.value = cellObj.f;
      cellMetaBadge.style.display = 'inline-block';
      cellMetaBadge.className = 'cell-meta-badge formula';
      cellMetaBadge.textContent = `= ${formatCellDisplay(activeSheet, selectedCoord, cellObj.v, anchorParsed.colStr, anchorParsed.row) || 'vacío'}`;
    } else {
      formulaInput.value = cellObj && cellObj.v !== undefined && cellObj.v !== null ? cellObj.v : '';
      const vType = isValidationCell(activeSheet, selectedCoord);
      if (vType) {
        cellMetaBadge.style.display = 'inline-block';
        cellMetaBadge.className = 'cell-meta-badge dropdown';
        cellMetaBadge.textContent = vType === 'CUSTOMER' ? 'Lista: Transporte!A:A' : 'Lista: Productos!F:F';
      } else {
        cellMetaBadge.style.display = 'none';
      }
    }

    highlightSelectedCellInDOM();
    updateStatusMetrics();

    if (openDropdownIfList && !bounds.isMulti) {
      const vType = isValidationCell(activeSheet, selectedCoord);
      if (vType) openValidationDropdown(selectedCoord, vType);
      else hideValidationDropdown();
    } else {
      hideValidationDropdown();
    }
  }

  function selectRange(startCoordStr, endCoordStr) {
    const p1 = parseCoord(startCoordStr);
    const p2 = parseCoord(endCoordStr);
    if (!p1 || !p2) return;
    selectedCoord = p1.coord;
    selectionEndCoord = p2.coord;
    selectCell(p2.coord, false, true);
  }

  function highlightSelectedCellInDOM() {
    gridBody.querySelectorAll('td.selected-cell, td.in-selection-range, td.sel-edge-top, td.sel-edge-bottom, td.sel-edge-left, td.sel-edge-right, td.copied-range-cell, td.autofill-preview-cell').forEach(el => {
      el.classList.remove('selected-cell', 'in-selection-range', 'sel-edge-top', 'sel-edge-bottom', 'sel-edge-left', 'sel-edge-right', 'copied-range-cell', 'autofill-preview-cell');
    });
    gridBody.querySelectorAll('.cell-fill-handle').forEach(el => el.remove());
    gridHead.querySelectorAll('th.active-col-header').forEach(el => el.classList.remove('active-col-header'));
    gridBody.querySelectorAll('th.active-row-header').forEach(el => el.classList.remove('active-row-header'));

    const bounds = getSelectionBounds();
    const { minRow, maxRow, minCol, maxCol } = bounds;

    const anchorTd = gridBody.querySelector(`td[data-coord="${selectedCoord}"]`);
    if (anchorTd) {
      anchorTd.classList.add('selected-cell');
    }

    const focusTd = gridBody.querySelector(`td[data-coord="${selectionEndCoord || selectedCoord}"]`);
    if (focusTd) {
      focusTd.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }

    for (let c = minCol; c <= maxCol; c++) {
      const colStr = indexToCol(c);
      const colTh = gridHead.querySelector(`th[data-col="${colStr}"]`);
      if (colTh) colTh.classList.add('active-col-header');
    }

    for (let r = minRow; r <= maxRow; r++) {
      const rowTh = gridBody.querySelector(`th.row-header[data-row="${r}"]`);
      if (rowTh) rowTh.classList.add('active-row-header');

      for (let c = minCol; c <= maxCol; c++) {
        const coord = `${indexToCol(c)}${r}`;
        const td = gridBody.querySelector(`td[data-coord="${coord}"]`);
        if (!td) continue;
        td.classList.add('in-selection-range');
        if (r === minRow) td.classList.add('sel-edge-top');
        if (r === maxRow) td.classList.add('sel-edge-bottom');
        if (c === minCol) td.classList.add('sel-edge-left');
        if (c === maxCol) td.classList.add('sel-edge-right');
      }
    }

    // Place draggable Excel Fill Handle on the bottom-right corner cell of the selection
    const brCoord = `${indexToCol(maxCol)}${maxRow}`;
    const brTd = gridBody.querySelector(`td[data-coord="${brCoord}"]`) || anchorTd;
    if (brTd && !inlineEditor) {
      const handle = document.createElement('div');
      handle.className = 'cell-fill-handle';
      handle.title = 'Arrastrar para rellenar / copiar celdas';
      brTd.appendChild(handle);
    }

    // Highlight copied/cut marching ants range if on activeSheet
    if (clipboardState && clipboardState.sheetName === activeSheet) {
      for (let r = clipboardState.minRow; r <= clipboardState.maxRow; r++) {
        for (let c = clipboardState.minCol; c <= clipboardState.maxCol; c++) {
          const cTd = gridBody.querySelector(`td[data-coord="${indexToCol(c)}${r}"]`);
          if (cTd) cTd.classList.add('copied-range-cell');
        }
      }
    }
  }

  function getDefaultFormula(sheetName, coord) {
    const parsed = parseCoord(coord);
    if (!parsed) return null;
    const { colStr, row } = parsed;

    if (sheetName === 'PASSING SHEET' && coord === 'A6') {
      return `="Load Date (Arrival to "&'PO Nuevo'!B18&"):"`;
    }

    if (sheetName === 'PO Nuevo') {
      const totalRow = getPONuevoTotalRow();
      const lastProdRow = totalRow - 1;
      const prodMax = Math.max(234, getLastRealRow('Productos'));

      if (coord === 'E12') return '=+VLOOKUP(F11,Transporte!A:L,12,FALSE)';
      if (coord === 'B18') return '=IFERROR(VLOOKUP(F11,Transporte!A:M,13,FALSE),"")';
      if (coord === 'F15') return '=+VLOOKUP(F11,Transporte!A:K,9,FALSE)';
      if (coord === 'F16') return '=+VLOOKUP(F11,Transporte!A:K,10,FALSE)';
      if (coord === 'B19') return '=F19-1';
      if (coord === `H${totalRow}`) return `=SUM(H22:H${lastProdRow})`;
      if (coord === `I${totalRow}`) return `=SUM(I22:I${lastProdRow})`;

      if (row >= 22 && row <= lastProdRow) {
        if (colStr === 'A') return row === 22 ? '=+A9' : `=IF(B${row}="","",$A$9)`;
        const selectedProductRow = `IF(PRODUCTROW(B${row})>0,PRODUCTROW(B${row}),`;
        const descriptionMatch = `MATCH(B${row},Productos!$F$1:$F$${prodMax},0)`;
        const boxDimensionMatch = `MATCH(1,INDEX((Productos!$K$1:$K$${prodMax}=F${row})*(Productos!$L$1:$L$${prodMax}<>""),0),0)`;
        const lookupByRow = (range, fallbackMatch) => {
          const rowExpression = `${selectedProductRow}${fallbackMatch})`;
          return `=IF(ISERROR(INDEX(${range},${rowExpression})),"",INDEX(${range},${rowExpression}))`;
        };
        if (colStr === 'C') return lookupByRow(`Productos!$E$1:$E$${prodMax}`, descriptionMatch);
        if (colStr === 'D') return lookupByRow(`Productos!$H$1:$H$${prodMax}`, descriptionMatch);
        if (colStr === 'E') return lookupByRow(`Productos!$G$1:$G$${prodMax}`, descriptionMatch);
        if (colStr === 'F') return lookupByRow(`Productos!$I$1:$I$${prodMax}`, descriptionMatch);
        // BOX DIMENSION comes from the box-code mapping table (K:L), not from
        // the selected product row. Several products share the same box type.
        if (colStr === 'G') {
          return `=IF(ISERROR(INDEX(Productos!$L$1:$L$${prodMax},${boxDimensionMatch})),"",INDEX(Productos!$L$1:$L$${prodMax},${boxDimensionMatch}))`;
        }
        if (colStr === 'J') return `=IF(B${row}="","",IF(COUNTIF(Productos!$F$1:$F$${prodMax},B${row})=0,"PENDIENTE DE BOM",""))`;
      }
    } else if (sheetName === 'DayOrders' && row >= 2 && colStr === 'P') {
      return `=+O${row}*N${row}*K${row}`;
    }
    return null;
  }

  function migratePONuevoDefaultFormulas() {
    const sheet = workbook?.sheets?.['PO Nuevo'];
    if (!sheet) return false;
    const totalRow = getPONuevoTotalRow();
    const changedRows = totalRow - 1;
    let changed = false;
    for (let row = 22; row <= changedRows; row++) {
      const hasProduct = String(sheet.cells[`B${row}`]?.v ?? '').trim() !== '';
      for (const col of ['A', 'C', 'D', 'E', 'F', 'G', 'J']) {
        const coord = `${col}${row}`;
        const expected = getDefaultFormula('PO Nuevo', coord);
        if (!expected) continue;
        const cell = sheet.cells[coord];
        if ((cell?.f || hasProduct) && cell?.f !== expected) {
          sheet.cells[coord] = { ...(cell || {}), f: expected, v: '' };
          changed = true;
        }
      }
    }

    const farmShipToFormula = getDefaultFormula('PO Nuevo', 'B18');
    const farmShipToCell = sheet.cells.B18;
    if (farmShipToFormula && farmShipToCell?.f !== farmShipToFormula) {
      sheet.cells.B18 = { ...(farmShipToCell || {}), f: farmShipToFormula, v: farmShipToCell?.v ?? '' };
      changed = true;
    }

    const passingSheet = workbook?.sheets?.['PASSING SHEET'];
    const passingLabelFormula = getDefaultFormula('PASSING SHEET', 'A6');
    const passingLabelCell = passingSheet?.cells?.A6;
    if (passingSheet && passingLabelFormula && passingLabelCell?.f !== passingLabelFormula) {
      passingSheet.cells.A6 = { ...(passingLabelCell || {}), f: passingLabelFormula, v: passingLabelCell?.v ?? '' };
      changed = true;
    }

    return changed;
  }

  function restoreRowFormulasPONuevo(sheet, row) {
    for (const c of ['A', 'C', 'D', 'E', 'F', 'G', 'J']) {
      const f = getDefaultFormula('PO Nuevo', `${c}${row}`);
      if (f) sheet.cells[`${c}${row}`] = { f, v: '' };
    }
  }

  // Raw single-cell mutation without triggering full render (used for both single and bulk operations)
  function applyCellValueRaw(sheetName, coord, rawInput) {
    const sheet = workbook.sheets[sheetName];
    if (!sheet) return;

    const parsed = parseCoord(coord);
    const trimmed = String(rawInput ?? '').trim();
    const lastProdRow = getPONuevoTotalRow() - 1;

    if (trimmed === '') {
      const defFormula = getDefaultFormula(sheetName, coord);
      if (defFormula) {
        sheet.cells[coord] = { f: defFormula, v: '' };
      } else {
        if (sheet.cells[coord]) {
          delete sheet.cells[coord];
        }
        if (sheetName === 'PO Nuevo' && parsed && parsed.colStr === 'B' && parsed.row >= 22 && parsed.row <= lastProdRow) {
          restoreRowFormulasPONuevo(sheet, parsed.row);
        }
      }
    } else if (trimmed.startsWith('=')) {
      sheet.cells[coord] = { f: trimmed, v: '' };
    } else {
      let val = trimmed;
      if ((sheetName === 'PO Nuevo' && (coord === 'F19' || coord === 'B19')) ||
          (sheetName === 'DayOrders' && parsed?.colStr === 'F')) {
        const mDate = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
        if (mDate) {
          const mm = mDate[1].padStart(2, '0');
          const dd = mDate[2].padStart(2, '0');
          const yyyy = mDate[3].length === 2 ? `20${mDate[3]}` : mDate[3];
          val = `${yyyy}-${mm}-${dd}`;
        }
      } else if (/^-?\d+(\.\d+)?$/.test(trimmed) && !(sheetName === 'PO Nuevo' && coord === 'A9')) {
        val = Number(trimmed);
      }
      sheet.cells[coord] = { v: val };

      if (sheetName === 'PO Nuevo' && parsed && parsed.colStr === 'B' && parsed.row >= 22 && parsed.row <= lastProdRow) {
        restoreRowFormulasPONuevo(sheet, parsed.row);
      }
      // If editing a row in DayOrders, ensure column P has its Final PRICE formula
      if (sheetName === 'DayOrders' && parsed && parsed.row >= 2 && parsed.colStr !== 'P' && !sheet.cells[`P${parsed.row}`]?.f) {
        sheet.cells[`P${parsed.row}`] = { f: `=+O${parsed.row}*N${parsed.row}*K${parsed.row}`, v: 0 };
      }
    }
  }

  // Update Cell Value or Formula
  function setCellContent(sheetName, coord, rawInput, metadata = null) {
    const sheet = workbook.sheets[sheetName];
    if (!sheet) return;
    pushUndoState();
    applyCellValueRaw(sheetName, coord, rawInput);
    if (
      metadata?.productRow &&
      sheetName === 'PO Nuevo' &&
      parseCoord(coord)?.colStr === 'B' &&
      sheet.cells[coord]
    ) {
      sheet.cells[coord].productRow = Number(metadata.productRow);
    }
    engine.recalculateAll();
    if (sheetName === 'DayOrders') populateDayOrdersFilters();
    renderSheet();
    selectCell(coord, false);
    saveWorkbookState();
  }

  // Clear all cells in current selection (Delete / Backspace) and restore any underlying formulas
  function clearSelectedRange() {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    pushUndoState();
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const coord = `${indexToCol(c)}${r}`;
        if (activeSheet === 'PO Nuevo' && coord === 'A1') continue;
        applyCellValueRaw(activeSheet, coord, '');
      }
    }
    engine.recalculateAll();
    if (activeSheet === 'DayOrders') populateDayOrdersFilters();
    renderSheet();
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
  }

  // Copy / Cut Selection (Ctrl+C / Ctrl+X)
  function copySelection(isCut = false) {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    const grid = [];
    const tsvRows = [];

    for (let r = minRow; r <= maxRow; r++) {
      const rowCells = [];
      const rowText = [];
      for (let c = minCol; c <= maxCol; c++) {
        const coord = `${indexToCol(c)}${r}`;
        const cellObj = sheet.cells[coord];
        const rawVal = cellObj?.v !== undefined && cellObj?.v !== null ? cellObj.v : '';
        const hasDefaultFormula = Boolean(getDefaultFormula(activeSheet, coord));
        rowCells.push({
          v: rawVal,
          f: (!hasDefaultFormula && cellObj?.f) ? cellObj.f : null,
          productRow: cellObj?.productRow || null,
          srcRow: r,
          srcCol: c
        });
        rowText.push(String(rawVal));
      }
      grid.push(rowCells);
      tsvRows.push(rowText.join('\t'));
    }

    const tsv = tsvRows.join('\r\n');
    clipboardState = {
      sheetName: activeSheet,
      minRow,
      maxRow,
      minCol,
      maxCol,
      isCut,
      grid,
      tsv
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsv).catch(() => {});
    }

    highlightSelectedCellInDOM();
    const count = (maxRow - minRow + 1) * (maxCol - minCol + 1);
    showToast(`${isCut ? 'Cortado' : 'Copiado'}: ${count} celda(s) (${indexToCol(minCol)}${minRow}:${indexToCol(maxCol)}${maxRow})`);
  }

  // Paste into Selection (Ctrl+V)
  function pasteDataMatrix(matrix, fromInternal = false) {
    if (!matrix || matrix.length === 0) return;
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;

    pushUndoState();
    const bounds = getSelectionBounds();
    const srcRows = matrix.length;
    const srcCols = matrix[0].length;

    const selRows = bounds.maxRow - bounds.minRow + 1;
    const selCols = bounds.maxCol - bounds.minCol + 1;

    // If copying 1x1 cell and selecting multiple cells, fill entire selection; otherwise paste block starting at top-left
    const targetRows = (srcRows === 1 && srcCols === 1 && (selRows > 1 || selCols > 1)) ? selRows : srcRows;
    const targetCols = (srcRows === 1 && srcCols === 1 && (selRows > 1 || selCols > 1)) ? selCols : srcCols;

    for (let dr = 0; dr < targetRows; dr++) {
      for (let dc = 0; dc < targetCols; dc++) {
        const destR = bounds.minRow + dr;
        const destC = bounds.minCol + dc;
        const destCoord = `${indexToCol(destC)}${destR}`;
        if (activeSheet === 'PO Nuevo' && destCoord === 'A1') continue;

        const item = matrix[dr % srcRows][dc % srcCols];
        const rawToWrite = (fromInternal && typeof item === 'object' && item !== null)
          ? (item.f ? item.f : item.v)
          : item;
        applyCellValueRaw(activeSheet, destCoord, rawToWrite);
        if (
          fromInternal &&
          item &&
          typeof item === 'object' &&
          item.productRow &&
          activeSheet === 'PO Nuevo' &&
          parseCoord(destCoord)?.colStr === 'B' &&
          sheet.cells[destCoord]
        ) {
          sheet.cells[destCoord].productRow = Number(item.productRow);
        }
      }
    }

    // If it was a Cut operation (Ctrl+X), clear the original cells that aren't overlapped by destination
    if (fromInternal && clipboardState && clipboardState.isCut) {
      const cutSheet = workbook.sheets[clipboardState.sheetName];
      if (cutSheet) {
        for (let r = clipboardState.minRow; r <= clipboardState.maxRow; r++) {
          for (let c = clipboardState.minCol; c <= clipboardState.maxCol; c++) {
            const overlapsDest =
              clipboardState.sheetName === activeSheet &&
              r >= bounds.minRow && r < bounds.minRow + targetRows &&
              c >= bounds.minCol && c < bounds.minCol + targetCols;
            if (!overlapsDest) {
              applyCellValueRaw(clipboardState.sheetName, `${indexToCol(c)}${r}`, '');
            }
          }
        }
      }
      clipboardState = null;
    }

    const endR = bounds.minRow + targetRows - 1;
    const endC = bounds.minCol + targetCols - 1;
    if (endR + 3 > sheet.maxRow) sheet.maxRow = endR + 5;

    engine.recalculateAll();
    if (activeSheet === 'DayOrders') populateDayOrdersFilters();
    renderSheet();
    selectedCoord = `${indexToCol(bounds.minCol)}${bounds.minRow}`;
    selectionEndCoord = `${indexToCol(endC)}${endR}`;
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
  }

  // Ctrl+D (Fill Down) and Ctrl+R (Fill Right)
  function fillSelectionDown() {
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    if (maxRow <= minRow) return;
    const sheet = workbook.sheets[activeSheet];
    pushUndoState();
    for (let c = minCol; c <= maxCol; c++) {
      const colStr = indexToCol(c);
      const topVal = sheet.cells[`${colStr}${minRow}`]?.v ?? '';
      for (let r = minRow + 1; r <= maxRow; r++) {
        applyCellValueRaw(activeSheet, `${colStr}${r}`, topVal);
      }
    }
    engine.recalculateAll();
    renderSheet();
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
    showToast('Rellenado hacia abajo (Ctrl+D)');
  }

  function fillSelectionRight() {
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    if (maxCol <= minCol) return;
    const sheet = workbook.sheets[activeSheet];
    pushUndoState();
    for (let r = minRow; r <= maxRow; r++) {
      const leftVal = sheet.cells[`${indexToCol(minCol)}${r}`]?.v ?? '';
      for (let c = minCol + 1; c <= maxCol; c++) {
        applyCellValueRaw(activeSheet, `${indexToCol(c)}${r}`, leftVal);
      }
    }
    engine.recalculateAll();
    renderSheet();
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
    showToast('Rellenado hacia la derecha (Ctrl+R)');
  }


  // Inline Editor inside Grid Cell (with live autocomplete for validation cells)
  function startInlineEdit(coord, initialText = null) {
    const td = gridBody.querySelector(`td[data-coord="${coord}"]`);
    if (!td) return;
    closeInlineEditor(false);

    const sheet = workbook.sheets[activeSheet];
    const cellObj = sheet?.cells[coord];
    const vType = isValidationCell(activeSheet, coord);
    const startVal = initialText !== null
      ? initialText
      : (cellObj?.f ? cellObj.f : (cellObj?.v ?? ''));

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'cell-inline-editor';
    input.value = startVal;

    td.appendChild(input);
    input.focus();
    if (initialText === null) input.select();
    else input.setSelectionRange(input.value.length, input.value.length);

    inlineEditor = { coord, input };

    // If this cell has a validation list (F11 Customer or B22:B41 Flower), open live dropdown below cell
    if (vType) {
      openValidationDropdown(coord, vType, false, initialText !== null ? startVal : '');
    }

    input.addEventListener('input', () => {
      formulaInput.value = input.value;
      if (vType) {
        if (validationDropdown.style.display === 'none') {
          openValidationDropdown(coord, vType, false, input.value);
        } else {
          validationSearchInput.value = input.value;
          renderValidationList(input.value);
        }
      }
    });

    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      const isDropOpen = vType && validationDropdown.style.display !== 'none';

      if (isDropOpen && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        e.preventDefault();
        moveValidationHighlight(e.key === 'ArrowDown' ? 1 : -1);
        return;
      }

      if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
        e.preventDefault();
        const val = input.value;
        hideValidationDropdown();
        closeInlineEditor(false);
        setCellContent(activeSheet, coord, val);
        if (e.key === 'ArrowDown') moveSelection(1, 0);
        else if (e.key === 'ArrowUp') moveSelection(-1, 0);
        else if (e.key === 'ArrowRight') moveSelection(0, 1);
        else if (e.key === 'ArrowLeft') moveSelection(0, -1);
        return;
      }

      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        let val = input.value;
        let selectedItem = null;
        if (isDropOpen && val.trim() !== '') {
          const activeItem = validationList.querySelector('.validation-item.active') ||
                             validationList.querySelector('.validation-item');
          if (activeItem) {
            selectedItem = getValidationItem(activeItem, val);
            val = selectedItem.value;
          }
        }
        hideValidationDropdown();
        closeInlineEditor(false);
        setCellContent(activeSheet, coord, val, getValidationMetadata(coord, selectedItem));
        if (e.key === 'Enter') moveSelection(1, 0);
        else moveSelection(0, e.shiftKey ? -1 : 1);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        hideValidationDropdown();
        closeInlineEditor(false);
        selectCell(coord, false);
      }
    });

    input.addEventListener('blur', () => {
      setTimeout(() => {
        if (inlineEditor && inlineEditor.coord === coord) {
          const val = input.value;
          hideValidationDropdown();
          closeInlineEditor(false);
          setCellContent(activeSheet, coord, val);
        }
      }, 120);
    });
  }

  function closeInlineEditor(commit = false) {
    if (!inlineEditor) return;
    const { coord, input } = inlineEditor;
    const val = input.value;
    inlineEditor = null;
    if (input.parentNode) input.parentNode.removeChild(input);
    if (commit) setCellContent(activeSheet, coord, val);
  }

  function commitPendingEditor() {
    if (!inlineEditor) return;
    hideValidationDropdown();
    closeInlineEditor(true);
  }

  function moveSelection(dRow, dCol, extendSelection = false) {
    const baseCoord = extendSelection ? (selectionEndCoord || selectedCoord) : selectedCoord;
    const parsed = parseCoord(baseCoord);
    if (!parsed) return;
    const sheet = workbook.sheets[activeSheet];
    let needsRender = false;

    const targetRow = Math.max(1, parsed.row + dRow);
    const targetCol = Math.max(1, parsed.col + dCol);

    if (targetRow > (sheet.maxRow || 100)) {
      sheet.maxRow = targetRow + 25;
      needsRender = true;
    }
    if (targetCol > (sheet.maxCol || 16)) {
      sheet.maxCol = targetCol + 4;
      needsRender = true;
    }

    if (needsRender) {
      renderSheet();
    }
    selectCell(`${indexToCol(targetCol)}${targetRow}`, false, extendSelection);
  }

  // Insert Rows (smartly expands PO Nuevo product table if inside rows 22..TOTAL, or shifts rows on any sheet)
  function insertRowsAt(insertRow, count = 1) {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet || count < 1) return;
    pushUndoState();

    if (activeSheet === 'PO Nuevo') {
      const oldTotalRow = getPONuevoTotalRow();
      // If user inserts inside or right at the PO Nuevo product table (rows 22..oldTotalRow)
      if (insertRow >= 22 && insertRow <= oldTotalRow) {
        const newTotalRow = oldTotalRow + count;
        sheet.poTotalRow = newTotalRow;
        const maxR = Math.max(sheet.maxRow || 100, newTotalRow + 20);

        // Shift rows from bottom up (maxR down to insertRow)
        for (let r = maxR; r >= insertRow; r--) {
          for (let c = 1; c <= (sheet.maxCol || 16); c++) {
            const colStr = indexToCol(c);
            const srcCell = sheet.cells[`${colStr}${r}`];
            if (srcCell) {
              sheet.cells[`${colStr}${r + count}`] = JSON.parse(JSON.stringify(srcCell));
              delete sheet.cells[`${colStr}${r}`];
            } else {
              delete sheet.cells[`${colStr}${r + count}`];
            }
          }
        }

        // Refresh all product row formulas for rows 22 .. newTotalRow - 1 so row references match their new row numbers
        for (let r = 22; r < newTotalRow; r++) {
          if (r >= insertRow && r < insertRow + count) {
            // Newly inserted blank product row: set default formulas
            for (let c = 1; c <= 12; c++) {
              delete sheet.cells[`${indexToCol(c)}${r}`];
            }
            restoreRowFormulasPONuevo(sheet, r);
          } else {
            // Existing product row that shifted: update formula cells (A, C..G, J) if they still have formulas
            for (const c of ['A', 'C', 'D', 'E', 'F', 'G', 'J']) {
              if (!sheet.cells[`${c}${r}`] || sheet.cells[`${c}${r}`].f) {
                const f = getDefaultFormula('PO Nuevo', `${c}${r}`);
                if (f) sheet.cells[`${c}${r}`] = { f, v: '' };
              }
            }
          }
        }

        // Update TOTAL row formulas at newTotalRow
        sheet.cells[`G${newTotalRow}`] = { v: 'TOTAL' };
        sheet.cells[`H${newTotalRow}`] = { f: `=SUM(H22:H${newTotalRow - 1})`, v: 0 };
        sheet.cells[`I${newTotalRow}`] = { f: `=SUM(I22:I${newTotalRow - 1})`, v: 0 };

        // Update PASSING SHEET B10 link to PO Nuevo total boxes
        const wsPassing = workbook.sheets['PASSING SHEET'];
        if (wsPassing && wsPassing.cells['B10']) {
          wsPassing.cells['B10'].f = `='PO Nuevo'!H${newTotalRow}`;
        }

        sheet.maxRow = Math.max(sheet.maxRow || 100, newTotalRow + 30);
        engine.recalculateAll();
        renderSheet();
        selectCell(`B${insertRow}`, false);
        saveWorkbookState();
        showToast(`${count} fila(s) de producto insertadas en PO Nuevo (Total ahora en fila ${newTotalRow})`);
        return;
      }
    }

    // Generic row insertion for DayOrders, Productos, Transporte, or below PO Nuevo table
    const maxR = Math.max(sheet.maxRow || 100, getLastRealRow(activeSheet) + 5);
    for (let r = maxR; r >= insertRow; r--) {
      for (let c = 1; c <= (sheet.maxCol || 16); c++) {
        const colStr = indexToCol(c);
        const srcCell = sheet.cells[`${colStr}${r}`];
        if (srcCell) {
          const cloned = JSON.parse(JSON.stringify(srcCell));
          if (activeSheet === 'DayOrders' && colStr === 'P' && cloned.f) {
            cloned.f = `=+O${r + count}*N${r + count}*K${r + count}`;
          }
          sheet.cells[`${colStr}${r + count}`] = cloned;
          delete sheet.cells[`${colStr}${r}`];
        } else {
          delete sheet.cells[`${colStr}${r + count}`];
        }
      }
    }

    if (activeSheet === 'DayOrders') {
      for (let r = insertRow; r < insertRow + count; r++) {
        if (r >= 2) {
          sheet.cells[`P${r}`] = { f: `=+O${r}*N${r}*K${r}`, v: 0 };
        }
      }
    }

    sheet.maxRow = Math.max(sheet.maxRow || 100, maxR + count + 20);
    engine.recalculateAll();
    if (activeSheet === 'DayOrders') populateDayOrdersFilters();
    renderSheet();
    selectCell(`A${insertRow}`, false);
    saveWorkbookState();
    showToast(`${count} fila(s) insertadas en fila ${insertRow}`);
  }

  // Delete Selected Row(s)
  function deleteSelectedRows() {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    const { minRow, maxRow } = getSelectionBounds();
    const count = maxRow - minRow + 1;

    if (activeSheet === 'PO Nuevo') {
      const oldTotalRow = getPONuevoTotalRow();
      if (minRow < 22) {
        showToast('No se pueden eliminar las filas del encabezado principal de PO Nuevo', true);
        return;
      }
      if (minRow >= 22 && maxRow < oldTotalRow) {
        const remainingProdRows = (oldTotalRow - 22) - count;
        if (remainingProdRows < 1) {
          showToast('Debe quedar al menos 1 fila de producto en la orden', true);
          return;
        }
        pushUndoState();
        const newTotalRow = oldTotalRow - count;
        sheet.poTotalRow = newTotalRow;
        const maxR = Math.max(sheet.maxRow || 100, oldTotalRow + 20);

        for (let r = minRow; r <= maxR; r++) {
          for (let c = 1; c <= (sheet.maxCol || 16); c++) {
            const colStr = indexToCol(c);
            const srcCell = sheet.cells[`${colStr}${r + count}`];
            if (srcCell) {
              sheet.cells[`${colStr}${r}`] = JSON.parse(JSON.stringify(srcCell));
            } else {
              delete sheet.cells[`${colStr}${r}`];
            }
          }
        }

        for (let r = 22; r < newTotalRow; r++) {
          for (const c of ['A', 'C', 'D', 'E', 'F', 'G', 'J']) {
            if (!sheet.cells[`${c}${r}`] || sheet.cells[`${c}${r}`].f) {
              const f = getDefaultFormula('PO Nuevo', `${c}${r}`);
              if (f) sheet.cells[`${c}${r}`] = { f, v: '' };
            }
          }
        }

        sheet.cells[`G${newTotalRow}`] = { v: 'TOTAL' };
        sheet.cells[`H${newTotalRow}`] = { f: `=SUM(H22:H${newTotalRow - 1})`, v: 0 };
        sheet.cells[`I${newTotalRow}`] = { f: `=SUM(I22:I${newTotalRow - 1})`, v: 0 };

        const wsPassing = workbook.sheets['PASSING SHEET'];
        if (wsPassing && wsPassing.cells['B10']) {
          wsPassing.cells['B10'].f = `='PO Nuevo'!H${newTotalRow}`;
        }

        engine.recalculateAll();
        renderSheet();
        selectCell(`B${Math.min(minRow, newTotalRow - 1)}`, false);
        saveWorkbookState();
        showToast(`${count} fila(s) eliminadas de PO Nuevo`);
        return;
      }
    }

    if (minRow === 1) {
      showToast('No se puede eliminar la fila 1 de encabezados', true);
      return;
    }

    pushUndoState();
    const maxR = Math.max(sheet.maxRow || 100, getLastRealRow(activeSheet) + 5);
    for (let r = minRow; r <= maxR; r++) {
      for (let c = 1; c <= (sheet.maxCol || 16); c++) {
        const colStr = indexToCol(c);
        const srcCell = sheet.cells[`${colStr}${r + count}`];
        if (srcCell) {
          const cloned = JSON.parse(JSON.stringify(srcCell));
          if (activeSheet === 'DayOrders' && colStr === 'P' && cloned.f) {
            cloned.f = `=+O${r}*N${r}*K${r}`;
          }
          sheet.cells[`${colStr}${r}`] = cloned;
        } else {
          delete sheet.cells[`${colStr}${r}`];
        }
      }
    }

    engine.recalculateAll();
    if (activeSheet === 'DayOrders') populateDayOrdersFilters();
    renderSheet();
    selectCell(`A${minRow}`, false);
    saveWorkbookState();
    showToast(`${count} fila(s) eliminadas`);
  }

  function expandGridDimensions(extraRows = 0, extraCols = 0) {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    if (extraRows > 0) sheet.maxRow = (sheet.maxRow || 100) + extraRows;
    if (extraCols > 0) sheet.maxCol = (sheet.maxCol || 16) + extraCols;
    renderSheet();
    saveWorkbookState();
    showToast(
      extraRows > 0
        ? `Se añadieron ${extraRows} filas más hacia abajo (Total: ${sheet.maxRow} filas)`
        : `Se añadieron ${extraCols} columnas más a la derecha (Hasta columna ${indexToCol(sheet.maxCol)})`
    );
  }

  // ============================================================================
  // SEARCHABLE DATA VALIDATION DROPDOWN (Transporte!A:A and Productos!F:F)
  // ============================================================================
  let currentDropdownCoord = null;
  let currentDropdownItems = [];
  let activeValidationIdx = 0;

  function getValidationItem(element, fallbackValue = '') {
    const sourceRow = Number(element?.getAttribute('data-source-row'));
    if (Number.isInteger(sourceRow) && sourceRow > 0) {
      const sourceItem = currentDropdownItems.find(item => item.sourceRow === sourceRow);
      if (sourceItem) return sourceItem;
    }
    const value = element?.getAttribute('data-val') ?? fallbackValue;
    return currentDropdownItems.find(item => item.value === value) || { value };
  }

  function getValidationMetadata(targetCoord, item) {
    const type = isValidationCell(activeSheet, targetCoord);
    return type === 'PRODUCT' && item?.sourceRow
      ? { productRow: item.sourceRow }
      : null;
  }

  function commitValidationItem(targetCoord, element, fallbackValue = '') {
    const item = getValidationItem(element, fallbackValue);
    setCellContent(activeSheet, targetCoord, item.value, getValidationMetadata(targetCoord, item));
  }

  function openValidationDropdown(coord, type, focusSearchBox = true, initialFilter = '') {
    const td = gridBody.querySelector(`td[data-coord="${coord}"]`);
    if (!td) return;

    currentDropdownCoord = coord;
    const items = [];

    if (type === 'CUSTOMER') {
      const trSheet = workbook.sheets['Transporte'];
      const maxR = Math.max(trSheet.maxRow || 77, getLastRealRow('Transporte'));
      for (let r = 2; r <= maxR; r++) {
        const cust = trSheet.cells[`A${r}`]?.v;
        if (cust !== undefined && cust !== null && String(cust).trim() !== '') {
          const carrier = String(trSheet.cells[`I${r}`]?.v || '').replace(/\s+/g, ' ').trim();
          const city = String(trSheet.cells[`L${r}`]?.v || '').replace(/\s+/g, ' ').trim();
          items.push({
            value: String(cust),
            sub: [carrier, city].filter(Boolean).join(' • '),
            sourceRow: r
          });
        }
      }
    } else if (type === 'PRODUCT') {
      const prSheet = workbook.sheets['Productos'];
      const maxR = Math.max(prSheet.maxRow || 234, getLastRealRow('Productos'));
      for (let r = 2; r <= maxR; r++) {
        const desc = prSheet.cells[`F${r}`]?.v;
        if (desc !== undefined && desc !== null && String(desc).trim() !== '') {
          const color = String(prSheet.cells[`E${r}`]?.v || '').trim();
          const box = String(prSheet.cells[`I${r}`]?.v || '').trim();
          const stems = prSheet.cells[`H${r}`]?.v ?? '';
          const pack = prSheet.cells[`G${r}`]?.v ?? '';
          items.push({
            value: String(desc),
            sub: `${color ? color + ' | ' : ''}St:${stems} Bu:${pack} ${box}`,
            sourceRow: r
          });
        }
      }
    }

    currentDropdownItems = items;

    const searchHeader = validationDropdown.querySelector('.validation-search-header');
    if (searchHeader) {
      searchHeader.style.display = focusSearchBox ? 'flex' : 'none';
    }

    validationDropdown.style.display = 'flex';
    validationDropdown.style.left = `${Math.max(42, td.offsetLeft)}px`;
    validationDropdown.style.top = `${td.offsetTop + td.offsetHeight + 3}px`;
    validationDropdown.style.width = `${Math.max(440, td.offsetWidth)}px`;

    validationSearchInput.value = initialFilter;
    renderValidationList(initialFilter);
    if (focusSearchBox) {
      setTimeout(() => validationSearchInput.focus(), 10);
    }
  }

  function renderValidationList(filterText) {
    const q = String(filterText || '').trim().toLowerCase();
    let filtered = currentDropdownItems;

    if (q) {
      // Prioritize items that start with the typed query (e.g. "mex" -> "Mex y Can USA" first), then contains
      const startsWithMatches = [];
      const containsMatches = [];
      for (const it of currentDropdownItems) {
        const vLow = it.value.toLowerCase();
        const sLow = it.sub.toLowerCase();
        if (vLow.startsWith(q)) {
          startsWithMatches.push(it);
        } else if (vLow.includes(q) || sLow.includes(q)) {
          containsMatches.push(it);
        }
      }
      filtered = [...startsWithMatches, ...containsMatches];
    }

    activeValidationIdx = 0;
    validationCountBadge.textContent = `${filtered.length} opciones`;
    validationList.innerHTML = filtered
      .map(
        (it, idx) => `<div class="validation-item${idx === 0 ? ' active' : ''}" data-idx="${idx}" data-source-row="${it.sourceRow ?? ''}" data-val="${escapeHtml(it.value)}">
          <span class="validation-item-main">${escapeHtml(it.value)}</span>
          <span class="validation-item-sub">${escapeHtml(it.sub)}</span>
        </div>`
      )
      .join('');
  }

  function moveValidationHighlight(delta) {
    const els = Array.from(validationList.querySelectorAll('.validation-item'));
    if (els.length === 0) return;
    els.forEach(el => el.classList.remove('active'));
    activeValidationIdx = Math.max(0, Math.min(els.length - 1, activeValidationIdx + delta));
    const target = els[activeValidationIdx];
    if (target) {
      target.classList.add('active');
      target.scrollIntoView({ block: 'nearest' });
    }
  }

  function hideValidationDropdown() {
    validationDropdown.style.display = 'none';
    currentDropdownCoord = null;
  }

  // ============================================================================
  // EXCEL VBA MACROS (100% FAITHFUL TO ORIGINAL .XLSM)
  // ============================================================================

  function getLastRealRow(sheetName) {
    const sheet = workbook.sheets[sheetName];
    if (!sheet) return 1;
    let lastRow = 1;
    for (const [coord, cell] of Object.entries(sheet.cells)) {
      if (cell && cell.v !== '' && cell.v !== null && cell.v !== undefined) {
        const p = parseCoord(coord);
        if (p && p.row > lastRow) lastRow = p.row;
      }
    }
    return lastRow;
  }

  // MACRO 1: Module1.bas -> Sub Guardar_Acumulado()
  function ejecutarGuardarDatos() {
    const wsOrigen = workbook.sheets['PO Nuevo'];
    const wsDestino = workbook.sheets['DayOrders'];
    if (!wsOrigen || !wsDestino) return;

    const lastProdRow = getPONuevoTotalRow() - 1;
    const linesToCopy = [];
    for (let fila = 22; fila <= lastProdRow; fila++) {
      const flowerVal = wsOrigen.cells[`B${fila}`]?.v;
      if (flowerVal !== undefined && flowerVal !== null && String(flowerVal).trim() !== '') {
        linesToCopy.push(fila);
      }
    }

    if (linesToCopy.length === 0) {
      showToast(`No hay líneas para guardar en PO Nuevo (filas 22 a ${lastProdRow} vacías)`, true);
      return;
    }

    pushUndoState();
    let filaDestino = Math.max(2, getLastRealRow('DayOrders') + 1);
    const firstAddedRow = filaDestino;

    const valA9 = wsOrigen.cells['A9']?.v ?? '';
    const valF11 = wsOrigen.cells['F11']?.v ?? '';
    const valF15 = wsOrigen.cells['F15']?.v ?? '';
    const valF16 = wsOrigen.cells['F16']?.v ?? 0;
    const valB18 = wsOrigen.cells['B18']?.v ?? '';
    const valF19 = wsOrigen.cells['F19']?.v ?? '';

    for (const fila of linesToCopy) {
      wsDestino.cells[`A${filaDestino}`] = { v: valA9 };
      wsDestino.cells[`B${filaDestino}`] = { v: valF11 };
      wsDestino.cells[`C${filaDestino}`] = { v: valF15 };
      wsDestino.cells[`D${filaDestino}`] = { v: valF16 };
      wsDestino.cells[`E${filaDestino}`] = { v: valB18 };
      wsDestino.cells[`F${filaDestino}`] = { v: valF19 };

      wsDestino.cells[`G${filaDestino}`] = { v: wsOrigen.cells[`A${fila}`]?.v ?? '' };
      wsDestino.cells[`H${filaDestino}`] = { v: wsOrigen.cells[`B${fila}`]?.v ?? '' };
      wsDestino.cells[`I${filaDestino}`] = { v: wsOrigen.cells[`C${fila}`]?.v ?? '' };
      wsDestino.cells[`J${filaDestino}`] = { v: wsOrigen.cells[`D${fila}`]?.v ?? '' };
      wsDestino.cells[`K${filaDestino}`] = { v: wsOrigen.cells[`E${fila}`]?.v ?? '' };
      wsDestino.cells[`L${filaDestino}`] = { v: wsOrigen.cells[`F${fila}`]?.v ?? '' };
      wsDestino.cells[`M${filaDestino}`] = { v: wsOrigen.cells[`G${fila}`]?.v ?? '' };
      wsDestino.cells[`N${filaDestino}`] = { v: wsOrigen.cells[`H${fila}`]?.v ?? 0 };
      wsDestino.cells[`O${filaDestino}`] = { v: wsOrigen.cells[`I${fila}`]?.v ?? 0 };
      wsDestino.cells[`P${filaDestino}`] = {
        f: `=+O${filaDestino}*N${filaDestino}*K${filaDestino}`,
        v: 0
      };

      filaDestino++;
    }

    if (filaDestino + 5 > wsDestino.maxRow) {
      wsDestino.maxRow = filaDestino + 10;
    }

    engine.recalculateAll();
    populateDayOrdersFilters();
    renderSheet();
    saveWorkbookState();

    showToast(
      `Registro agregado correctamente (${linesToCopy.length} línea(s) guardadas en DayOrders fila ${firstAddedRow})`,
      false,
      () => {
        switchSheet('DayOrders');
        selectCell(`B${firstAddedRow}`, false);
      },
      'Ver en DayOrders →'
    );
  }

  // Helper LimpiarTexto from Module3.bas
  function limpiarTexto(texto) {
    return String(texto || '')
      .replace(/[\/\\:*?"<>|]/g, '-')
      .trim();
  }

  // MACRO 3 & 4: Module3.bas / Module4.bas -> Generar_PDF_PO_Orden (PO Nuevo A1:I{totalRow})
  function ejecutarGenerarOrdenPDF() {
    commitPendingEditor();
    const ws = workbook.sheets['PO Nuevo'];
    const totalRow = getPONuevoTotalRow();
    const lastProdRow = totalRow - 1;
    const lineCount = Array.from({ length: Math.max(0, lastProdRow - 21) }, (_, index) => index + 22)
      .filter(row => String(ws.cells[`B${row}`]?.v || '').trim() !== '').length;
    if (lineCount === 0) {
      showToast(`No hay líneas para generar la orden (filas 22 a ${lastProdRow} vacías)`, true);
      return false;
    }
    const cliente = limpiarTexto(ws.cells['F11']?.v || 'Cliente');
    const fecha = limpiarTexto(formatDateDisplay(ws.cells['F19']?.v || 'Fecha'));
    const po = limpiarTexto(ws.cells['A9']?.v || 'PO');
    const suggestedName = `${cliente}_${fecha}_${po}.pdf`;

    let rowsHtml = '';
    for (let r = 22; r <= lastProdRow; r++) {
      const bVal = ws.cells[`B${r}`]?.v;
      if (!bVal && r > 28) continue; // Keep compact like 1-page print area
      rowsHtml += `<tr>
        <td style="text-align:center;font-weight:600;">${escapeHtml(ws.cells[`A${r}`]?.v ?? '')}</td>
        <td>${escapeHtml(bVal ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`C${r}`]?.v ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`D${r}`]?.v ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`E${r}`]?.v ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`F${r}`]?.v ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`G${r}`]?.v ?? '')}</td>
        <td style="text-align:center;font-weight:600;">${escapeHtml(ws.cells[`H${r}`]?.v ?? '')}</td>
        <td style="text-align:right;">${escapeHtml(formatCurrency(ws.cells[`I${r}`]?.v))}</td>
      </tr>`;
    }

    const html = `
      <div style="font-family:'Inter',sans-serif;color:#111827;">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #14532d;padding-bottom:12px;margin-bottom:14px;">
          <div style="display:flex;align-items:center;gap:14px;">
            <img src="logo.png" alt="Logo" style="height:64px;width:auto;" onerror="this.style.display='none'" />
            <div>
              <h2 style="font-size:18px;font-weight:800;color:#14532d;margin:0;">PACIFICA PRODUCE FARMS LLC</h2>
              <p style="font-size:12px;margin:2px 0;">PO Box 530990 • San Diego, CA 92154</p>
            </div>
          </div>
          <div style="text-align:right;">
            <div style="background:#ffff00;padding:4px 10px;border:1px solid #000;font-weight:700;font-size:14px;">
              Customer PO: ${escapeHtml(ws.cells['A9']?.v ?? '')}
            </div>
            <div style="font-size:11px;color:#475569;margin-top:4px;">Komet Order: ${escapeHtml(ws.cells['A8']?.v ?? '-')}</div>
          </div>
        </div>

        <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:14px;font-size:12px;">
          <div style="border:1px solid #cbd5e1;padding:10px;border-radius:4px;">
            <div><strong>SHIPPER:</strong> Pacifica Produce Farms LLC</div>
            <div style="margin-top:6px;"><strong>Farm Ship to:</strong> ${escapeHtml(ws.cells['B18']?.v ?? '')}</div>
            <div style="margin-top:4px;"><strong>Load Date:</strong> <span style="background:#ffff00;padding:1px 6px;font-weight:700;">${escapeHtml(formatDateDisplay(ws.cells['B19']?.v))}</span></div>
          </div>
          <div style="border:1px solid #cbd5e1;padding:10px;border-radius:4px;">
            <div><strong style="background:#ffff00;padding:1px 4px;">SHIP TO:</strong> <strong>${escapeHtml(ws.cells['F11']?.v ?? '')}</strong> (${escapeHtml(ws.cells['E12']?.v ?? '')})</div>
            <div style="margin-top:6px;"><strong>TRUCK LINE:</strong> ${escapeHtml(ws.cells['F15']?.v ?? '')} &nbsp;|&nbsp; <strong>ACCOUNT #:</strong> ${escapeHtml(ws.cells['F16']?.v ?? '')}</div>
            <div style="margin-top:4px;"><strong>Tender Date:</strong> <span style="background:#ffff00;padding:1px 6px;font-weight:700;">${escapeHtml(formatDateDisplay(ws.cells['F19']?.v))}</span></div>
          </div>
        </div>

        <table class="pdf-table">
          <thead>
            <tr>
              <th style="background:#d9e1f2;">PO</th>
              <th style="background:#ffff00;">FLOWER</th>
              <th style="background:#d9e1f2;">COLOR</th>
              <th style="background:#d9e1f2;">STEM/BCH</th>
              <th style="background:#d9e1f2;">BU/BOX</th>
              <th style="background:#d9e1f2;">BOX TYPE</th>
              <th style="background:#d9e1f2;">BOX DIMENSION</th>
              <th style="background:#ffff00;">CASES</th>
              <th style="background:#ffff00;">Price $</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
          <tfoot>
            <tr style="font-weight:800;background:#f8fafc;">
              <td colspan="7" style="text-align:right;">TOTAL</td>
              <td style="text-align:center;">${escapeHtml(ws.cells[`H${totalRow}`]?.v ?? 0)}</td>
              <td style="text-align:right;">${escapeHtml(formatCurrency(ws.cells[`I${totalRow}`]?.v ?? 0))}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    `;

    openPdfPreviewModal(`Orden de Compra — PO Nuevo (A1:I${totalRow})`, suggestedName, html);
    return true;
  }

  async function generarArchivoPosco() {
    const button = document.getElementById('btnVisualGenerarOrden');
    if (button) {
      button.disabled = true;
      button.dataset.originalLabel = button.textContent.trim();
      button.querySelector('.btn-dot')?.classList.add('is-busy');
    }

    try {
      const response = await fetch('/api/posco/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(workbook)
      });

      if (!response.ok) {
        let detail = 'No se pudo crear el archivo POSCO.';
        try {
          const payload = await response.json();
          detail = payload.error || detail;
        } catch {
          // Conserva el mensaje genérico si el servidor no devuelve JSON.
        }
        throw new Error(detail);
      }

      await response.arrayBuffer();
      emailSendMode = 'order';
      const sendButton = document.getElementById('btnSendOrder');
      if (sendButton) {
        sendButton.style.display = 'inline-flex';
        sendButton.disabled = false;
      }

      let pending = getPendingBomRows();
      const pendingHeader = response.headers.get('X-BOM-Pending');
      if (pendingHeader) {
        try {
          pending = JSON.parse(decodeURIComponent(pendingHeader));
        } catch {
          // Usa la detección local si el encabezado no se puede interpretar.
        }
      }

      const rowCount = response.headers.get('X-POSCO-Rows') || '0';
      if (pending.length > 0) {
        showToast(
          `Orden generada: PDF y formato POSCO listos (${rowCount} línea(s)). Hay ${pending.length} producto(s) PENDIENTE(S) DE BOM. Pulsa Enviar para mandar todo.`,
          true,
          abrirModalHomologeados,
          'Revisar BOM'
        );
      } else {
        showToast(`Orden generada: PDF y formato POSCO listos (${rowCount} línea(s)). Pulsa Enviar para mandar todo por correo.`);
      }
    } catch (error) {
      showToast(`No se pudo generar el Excel POSCO: ${error.message}`, true);
    } finally {
      if (button) {
        button.disabled = false;
        button.querySelector('.btn-dot')?.classList.remove('is-busy');
      }
    }
  }

  async function enviarOrdenPorCorreo() {
    commitPendingEditor();
    const button = document.getElementById('btnSendOrder');
    if (!button) return;
    const isPassingSheet = emailSendMode === 'passing';
    const originalLabel = button.textContent;
    button.disabled = true;
    button.textContent = '✉️ Enviando...';
    try {
      const response = await fetch(isPassingSheet ? '/api/passing/send' : '/api/order/send', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(workbook)
      });
      let payload = {};
      try {
        payload = await response.json();
      } catch {
        payload = {};
      }
      if (!response.ok) throw new Error(payload.error || 'No se pudo enviar la orden por correo.');
      const recipients = Array.isArray(payload.sentTo) ? payload.sentTo.join(', ') : 'los destinatarios configurados';
      const pendingMessage = payload.pendingBom > 0 ? ` BOM pendiente incluido: ${payload.pendingBom}.` : '';
      showToast(`${isPassingSheet ? 'Passing Sheet enviado' : 'Orden enviada'} a ${recipients}.${pendingMessage}`);
    } catch (error) {
      showToast(`No se pudo enviar la orden: ${error.message}`, true);
    } finally {
      button.disabled = false;
      button.textContent = originalLabel;
    }
  }

  // MACRO 5 & 6: Module5.bas / Module6.bas -> DayOrders PDF (con o sin precios)
  function ejecutarPdfDayOrders(includePrices = true) {
    const ws = workbook.sheets['DayOrders'];
    const lastRow = getLastRealRow('DayOrders');
    if (lastRow <= 1) {
      showToast('Solo hay encabezados, no hay registros para generar el PDF', true);
      return;
    }

    const rand = Math.floor(1000 + Math.random() * 9000);
    const nowStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const prefix = includePrices ? 'DayOrders' : 'PO';
    const suggestedName = `${prefix}_${nowStr}_${rand}.pdf`;

    let rowsHtml = '';
    let rowCount = 0;
    let totalBoxes = 0;
    let totalAmount = 0;

    for (let r = 2; r <= lastRow; r++) {
      const client = ws.cells[`B${r}`]?.v;
      const flower = ws.cells[`H${r}`]?.v;
      if (!client && !flower) continue;

      if (dayFilterDate !== 'ALL' && (ws.cells[`F${r}`]?.v || '') !== dayFilterDate) continue;
      if (dayFilterClient !== 'ALL' && (ws.cells[`B${r}`]?.v || '') !== dayFilterClient) continue;

      rowCount++;
      const boxes = Number(ws.cells[`N${r}`]?.v || 0);
      const finalPrice = Number(ws.cells[`P${r}`]?.v || 0);
      totalBoxes += boxes;
      totalAmount += finalPrice;

      rowsHtml += `<tr>
        <td>${escapeHtml(ws.cells[`A${r}`]?.v ?? '')}</td>
        <td><strong>${escapeHtml(client ?? '')}</strong></td>
        <td>${escapeHtml(ws.cells[`C${r}`]?.v ?? '')}</td>
        <td>${escapeHtml(ws.cells[`D${r}`]?.v ?? '')}</td>
        <td>${escapeHtml(ws.cells[`E${r}`]?.v ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(formatDateDisplay(ws.cells[`F${r}`]?.v))}</td>
        <td>${escapeHtml(ws.cells[`G${r}`]?.v ?? '')}</td>
        <td>${escapeHtml(flower ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`J${r}`]?.v ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`K${r}`]?.v ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`L${r}`]?.v ?? '')}</td>
        <td style="text-align:center;">${escapeHtml(ws.cells[`M${r}`]?.v ?? '')}</td>
        <td style="text-align:center;font-weight:700;">${escapeHtml(ws.cells[`N${r}`]?.v ?? '')}</td>
        ${includePrices ? `
          <td style="text-align:right;">${escapeHtml(formatCurrency(ws.cells[`O${r}`]?.v))}</td>
          <td style="text-align:right;background:#d9f99d;font-weight:700;">${escapeHtml(formatCurrency(ws.cells[`P${r}`]?.v))}</td>
        ` : ''}
      </tr>`;
    }

    const filterNote = (dayFilterDate !== 'ALL' || dayFilterClient !== 'ALL')
      ? `Filtro activo: Fecha=${dayFilterDate} | Cliente=${dayFilterClient} (${rowCount} registros)`
      : `Todos los registros (${rowCount} filas)`;

    const html = `
      <div style="font-family:'Inter',sans-serif;color:#111827;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
          <div>
            <h2 style="font-size:16px;font-weight:800;color:#0070c0;margin:0;">
              DAY ORDERS — ${includePrices ? 'REPORTE CON PRECIOS (A:P)' : 'PASSING SHEET / SIN PRECIOS (A:N)'}
            </h2>
            <p style="font-size:11px;color:#475569;margin-top:2px;">${escapeHtml(filterNote)}</p>
          </div>
          <div style="font-size:12px;font-weight:700;">
            Total Cajas: ${totalBoxes} ${includePrices ? ` | Total Final: ${formatCurrency(totalAmount)}` : ''}
          </div>
        </div>
        <table class="pdf-table" style="font-size:10px;">
          <thead>
            <tr style="background:#0070c0;color:#ffffff;">
              <th>PO</th>
              <th>Client</th>
              <th>TRUCK LINE</th>
              <th>ACCOUNT #</th>
              <th>FLORAL TRADE</th>
              <th>Tender Date</th>
              <th>PO</th>
              <th>FLOWER</th>
              <th>Stems</th>
              <th>Pack</th>
              <th>Caja</th>
              <th>Medidas</th>
              <th>Num Cajas</th>
              ${includePrices ? `<th style="background:#0070c0;">Bu Price</th><th style="background:#92d050;color:#000;">Final PRICE</th>` : ''}
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    `;

    openPdfPreviewModal(
      includePrices ? 'PDF DayOrders (Con Precios)' : 'PDF DayOrders (Sin Precios)',
      suggestedName,
      html
    );
  }

  // MACRO 7: Módulo1.bas -> Public Sub Generar_PDF_Passing_Sheet()
  function ejecutarPdfPassingSheet() {
    commitPendingEditor();
    const ws = workbook.sheets['PASSING SHEET'];
    const arrivalLabel = ws.cells['A6']?.v || 'Load Date (Arrival to Floral Trade):';
    const html = `
      <div style="font-family:'Inter',sans-serif;max-width:680px;margin:0 auto;padding:16px;border:1px solid #cbd5e1;">
        <h2 style="font-size:18px;font-weight:800;margin-bottom:8px;">PASSING SHEET – FLORAL TRADE</h2>
        <p style="font-size:12.5px;color:#334155;margin-bottom:18px;">
          ${escapeHtml(ws.cells['A2']?.v || '')}
        </p>
        <table class="pdf-table" style="font-size:13px;margin-bottom:20px;">
          <tbody>
            <tr><td style="font-weight:700;width:55%;">Farm / Shipper:</td><td>${escapeHtml(ws.cells['B4']?.v ?? '')}</td></tr>
            <tr><td style="font-weight:700;">Customer Name:</td><td><strong>${escapeHtml(ws.cells['B5']?.v ?? '')}</strong></td></tr>
            <tr><td style="font-weight:700;">${escapeHtml(arrivalLabel)}</td><td>${escapeHtml(formatDateDisplay(ws.cells['B6']?.v))}</td></tr>
            <tr><td style="font-weight:700;">Release Date (Deliver to Trucking Line):</td><td>${escapeHtml(formatDateDisplay(ws.cells['B7']?.v))}</td></tr>
            <tr><td style="font-weight:700;">Trucking Line:</td><td>${escapeHtml(ws.cells['B8']?.v ?? '')}</td></tr>
            <tr><td style="font-weight:700;">Account # (if applicable):</td><td>${escapeHtml(ws.cells['B9']?.v ?? '')}</td></tr>
            <tr><td style="font-weight:700;">Number of Boxes:</td><td style="font-weight:800;font-size:15px;">${escapeHtml(ws.cells['B10']?.v ?? 0)}</td></tr>
          </tbody>
        </table>
        <div style="margin-top:18px;font-size:13px;">
          <strong>Special Instructions:</strong>
          <div style="border-bottom:1px solid #334155;height:26px;margin-top:8px;">${escapeHtml(ws.cells['B13']?.v ?? '')}</div>
          <div style="border-bottom:1px solid #334155;height:26px;">${escapeHtml(ws.cells['B14']?.v ?? '')}</div>
          <div style="border-bottom:1px solid #334155;height:26px;">${escapeHtml(ws.cells['B15']?.v ?? '')}</div>
        </div>
        <div style="margin-top:32px;font-size:13px;">
          <p style="margin-bottom:14px;"><strong>Authorized by Pacifica Farms:</strong> ________________________________</p>
          <p><strong>Contact Phone:</strong> ______________________ &nbsp;&nbsp; <strong>Email:</strong> ______________________</p>
        </div>
      </div>
    `;
    openPdfPreviewModal('PASSING SHEET – FLORAL TRADE ($A$1:$B$18)', 'Passing_Sheet.pdf', html);
    emailSendMode = 'passing';
    const sendButton = document.getElementById('btnSendOrder');
    if (sendButton) {
      sendButton.style.display = 'inline-flex';
      sendButton.disabled = false;
    }
  }

  function openPdfPreviewModal(title, filename, contentHtml) {
    document.getElementById('pdfModalTitle').textContent = title;
    document.getElementById('pdfModalSubtitle').textContent = `Nombre sugerido por macro: ${filename}`;
    document.getElementById('pdfFilenameInput').value = filename;
    document.getElementById('pdfPreviewContainer').innerHTML = contentHtml;
    document.getElementById('printOnlyRoot').innerHTML = contentHtml;
    emailSendMode = null;
    const sendButton = document.getElementById('btnSendOrder');
    if (sendButton) {
      sendButton.style.display = 'none';
      sendButton.disabled = true;
    }
    document.getElementById('pdfModalOverlay').style.display = 'flex';
  }

  // BOM / Homologados Inspector Modal
  function abrirModalHomologeados() {
    const ws = workbook.sheets['PO Nuevo'];
    const lastProdRow = getPONuevoTotalRow() - 1;
    const rows = [];
    for (let r = 22; r <= lastProdRow; r++) {
      const flower = ws.cells[`B${r}`]?.v;
      if (flower) {
        const status = ws.cells[`J${r}`]?.v || 'OK EN CATÁLOGO';
        rows.push({
          row: r,
          flower,
          color: ws.cells[`C${r}`]?.v || '-',
          box: ws.cells[`F${r}`]?.v || '-',
          dim: ws.cells[`G${r}`]?.v || '-',
          pending: isBomPendingStatus(status),
          recommendation: [
            flower,
            ws.cells[`C${r}`]?.v,
            ws.cells[`D${r}`]?.v ? `STEM/${ws.cells[`D${r}`].v}` : '',
            ws.cells[`E${r}`]?.v ? `BU/${ws.cells[`E${r}`].v}` : '',
            ws.cells[`F${r}`]?.v
          ].filter(Boolean).join(' ') || 'Producto Wholesale'
        });
      }
    }

    const body = document.getElementById('bomModalBody');
    if (rows.length === 0) {
      body.innerHTML = `<p style="color:#475569;">No hay flores capturadas actualmente en PO Nuevo (filas 22 a ${lastProdRow}).</p>`;
    } else {
      body.innerHTML = `
        <table class="pdf-table" style="font-size:12.5px;">
          <thead>
            <tr style="background:#f1f5f9;">
              <th>Fila PO</th>
              <th>Descripción Flor (B22:B${lastProdRow})</th>
              <th>Color</th>
              <th>Caja</th>
              <th>Medidas</th>
              <th>Propuesta BOM</th>
              <th>Estado BOM (Columna J)</th>
              <th>Acción</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(it => `
              <tr>
                <td style="text-align:center;">Fila ${it.row}</td>
                <td><strong>${escapeHtml(it.flower)}</strong></td>
                <td style="text-align:center;">${escapeHtml(it.color)}</td>
                <td style="text-align:center;">${escapeHtml(it.box)}</td>
                <td style="text-align:center;">${escapeHtml(it.dim)}</td>
                <td>${escapeHtml(it.recommendation)}</td>
                <td style="text-align:center;font-weight:700;color:${it.pending ? '#dc2626' : '#15803d'};">
                  ${it.pending ? '⚠️ PENDIENTE DE BOM' : '✓ HOMOLOGADO'}
                </td>
                <td style="text-align:center;">
                  ${it.pending ? `<button class="mini-reset-btn" style="margin:0;background:#15803d;color:#fff;" data-add-bom="${escapeHtml(it.flower)}">+ Crear BOM / Agregar</button>` : 'Listo'}
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;

      body.querySelectorAll('button[data-add-bom]').forEach(btn => {
        btn.addEventListener('click', () => {
          const flowerName = btn.getAttribute('data-add-bom');
          agregarProductoRapido(flowerName);
          abrirModalHomologeados();
        });
      });
    }
    document.getElementById('bomModalOverlay').style.display = 'flex';
  }

  function agregarProductoRapido(flowerName) {
    const wsProd = workbook.sheets['Productos'];
    const newRow = getLastRealRow('Productos') + 1;
    pushUndoState();
    wsProd.cells[`A${newRow}`] = { v: 'PACIFICA PRODUCE FARMS' };
    wsProd.cells[`B${newRow}`] = { v: 'WHOLESALE' };
    wsProd.cells[`C${newRow}`] = { v: 'BOUQUET' };
    wsProd.cells[`D${newRow}`] = { v: flowerName.split(' ')[0] || 'FLOR' };
    wsProd.cells[`E${newRow}`] = { v: 'ASSORTED' };
    wsProd.cells[`F${newRow}`] = { v: flowerName };
    wsProd.cells[`G${newRow}`] = { v: 10 };
    wsProd.cells[`H${newRow}`] = { v: 10 };
    wsProd.cells[`I${newRow}`] = { v: 'F4P' };
    wsProd.maxRow = Math.max(wsProd.maxRow, newRow + 2);

    // Expand PO Nuevo lookup ranges if needed
    engine.recalculateAll();
    renderSheet();
    saveWorkbookState();
    showToast(`Producto "${flowerName}" agregado a Productos (fila ${newRow})`);
  }

  function limpiarLineasPONuevo() {
    const ws = workbook.sheets['PO Nuevo'];
    const lastProdRow = getPONuevoTotalRow() - 1;
    pushUndoState();
    for (let r = 22; r <= lastProdRow; r++) {
      delete ws.cells[`B${r}`];
      delete ws.cells[`H${r}`];
      delete ws.cells[`I${r}`];
      restoreRowFormulasPONuevo(ws, r);
    }
    engine.recalculateAll();
    renderSheet();
    selectCell('B22', false);
    saveWorkbookState();
    showToast(`Líneas de PO Nuevo (filas 22 a ${lastProdRow}) limpiadas; fórmulas listas para nueva orden`);
  }

  // Populate DayOrders date & client filters
  function populateDayOrdersFilters() {
    if (!filterDayOrdersDate || !filterDayOrdersClient) return;
    const ws = workbook.sheets['DayOrders'];
    if (!ws) return;
    const dates = new Set();
    const clients = new Set();
    const lastR = getLastRealRow('DayOrders');

    for (let r = 2; r <= lastR; r++) {
      const d = ws.cells[`F${r}`]?.v;
      const c = ws.cells[`B${r}`]?.v;
      if (d) dates.add(String(d));
      if (c) clients.add(String(c));
    }

    const sortedDates = Array.from(dates).sort().reverse();
    const sortedClients = Array.from(clients).sort();

    filterDayOrdersDate.innerHTML = `<option value="ALL">Todas las fechas (${lastR - 1} filas)</option>` +
      sortedDates.map(d => `<option value="${escapeHtml(d)}"${d === dayFilterDate ? ' selected' : ''}>${escapeHtml(formatDateDisplay(d))} (${escapeHtml(d)})</option>`).join('');

    filterDayOrdersClient.innerHTML = `<option value="ALL">Todos los clientes (${sortedClients.length})</option>` +
      sortedClients.map(c => `<option value="${escapeHtml(c)}"${c === dayFilterClient ? ' selected' : ''}>${escapeHtml(c)}</option>`).join('');
  }

  function switchSheet(sheetName) {
    if (!workbook.sheets[sheetName]) return;
    activeSheet = sheetName;
    document.querySelectorAll('.sheet-tab').forEach(btn => {
      btn.classList.toggle('active', btn.getAttribute('data-sheet') === sheetName);
    });
    const defaultCell = sheetName === 'PO Nuevo' ? 'F11' : 'A1';
    renderSheet();
    selectCell(defaultCell, false);
  }

  function updateStatusMetrics() {
    const wsPO = workbook.sheets['PO Nuevo'];
    const totalRow = getPONuevoTotalRow();
    const totalBoxes = wsPO?.cells[`H${totalRow}`]?.v ?? 0;
    const totalPrice = wsPO?.cells[`I${totalRow}`]?.v ?? 0;

    const bounds = getSelectionBounds();
    if (bounds.isMulti) {
      const sheet = workbook.sheets[activeSheet];
      let nonEmptyCount = 0;
      let numericCount = 0;
      let numericSum = 0;
      for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
        for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
          const val = sheet?.cells[`${indexToCol(c)}${r}`]?.v;
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            nonEmptyCount++;
            const num = Number(val);
            if (!isNaN(num)) {
              numericCount++;
              numericSum += num;
            }
          }
        }
      }
      const rangeStr = `${indexToCol(bounds.minCol)}${bounds.minRow}:${indexToCol(bounds.maxCol)}${bounds.maxRow}`;
      document.getElementById('metricCount').textContent = `Rango: ${activeSheet}!${rangeStr} (Recuento: ${nonEmptyCount})`;
      if (numericCount > 0) {
        const avg = numericSum / numericCount;
        document.getElementById('metricBoxes').textContent = `Promedio sel.: ${avg.toLocaleString('en-US', { maximumFractionDigits: 2 })}`;
        document.getElementById('metricSum').textContent = `Suma sel.: ${numericSum.toLocaleString('en-US', { maximumFractionDigits: 2 })}`;
      } else {
        document.getElementById('metricBoxes').textContent = `Total Cajas PO: ${totalBoxes}`;
        document.getElementById('metricSum').textContent = `Total PO: ${formatCurrency(totalPrice) || '$0.00'}`;
      }
    } else {
      document.getElementById('metricBoxes').textContent = `Total Cajas PO: ${totalBoxes}`;
      document.getElementById('metricSum').textContent = `Total PO: ${formatCurrency(totalPrice) || '$0.00'}`;
      document.getElementById('metricCount').textContent = `Celda: ${activeSheet}!${selectedCoord}`;
    }
  }

  function showToast(msg, isWarn = false, actionFn = null, actionLabel = '') {
    const container = document.getElementById('toastContainer');
    const el = document.createElement('div');
    el.className = `excel-toast${isWarn ? ' warn' : ''}`;
    el.innerHTML = `<span>${escapeHtml(msg)}</span>`;
    if (actionFn && actionLabel) {
      const btn = document.createElement('button');
      btn.className = 'mini-reset-btn';
      btn.style.margin = '0 0 0 8px';
      btn.textContent = actionLabel;
      btn.addEventListener('click', () => {
        actionFn();
        el.remove();
      });
      el.appendChild(btn);
    }
    container.appendChild(el);
    setTimeout(() => {
      if (el.parentNode) el.remove();
    }, 4500);
  }

  // Export active sheet to CSV
  function exportActiveSheetCsv() {
    const sheet = workbook.sheets[activeSheet];
    const lastR = getLastRealRow(activeSheet);
    const maxC = sheet.maxCol || 12;
    const lines = [];
    for (let r = 1; r <= lastR; r++) {
      const rowVals = [];
      for (let c = 1; c <= maxC; c++) {
        const coord = `${indexToCol(c)}${r}`;
        const v = sheet.cells[coord]?.v ?? '';
        rowVals.push(`"${String(v).replace(/"/g, '""')}"`);
      }
      lines.push(rowVals.join(','));
    }
    const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${activeSheet}_PacificaFarms.csv`;
    a.click();
    showToast(`Hoja "${activeSheet}" exportada a CSV`);
  }

  function openCellDatePicker(coord) {
    const td = gridBody.querySelector(`td[data-coord="${coord}"]`);
    if (!td) return;
    closeInlineEditor(false);

    const sheet = workbook.sheets[activeSheet];
    const curVal = String(sheet?.cells[coord]?.v || '').trim();

    const picker = document.createElement('input');
    picker.type = 'date';
    picker.style.position = 'fixed';
    const rect = td.getBoundingClientRect();
    picker.style.left = `${rect.left}px`;
    picker.style.top = `${rect.bottom}px`;
    picker.style.width = '1px';
    picker.style.height = '1px';
    picker.style.opacity = '0';
    picker.style.pointerEvents = 'none';

    if (/^\d{4}-\d{2}-\d{2}$/.test(curVal)) {
      picker.value = curVal;
    }

    document.body.appendChild(picker);

    const cleanup = () => {
      if (picker.parentNode) picker.parentNode.removeChild(picker);
    };

    picker.addEventListener('change', () => {
      if (picker.value) {
        setCellContent(activeSheet, coord, picker.value);
      }
      cleanup();
    });

    picker.addEventListener('blur', () => {
      setTimeout(cleanup, 200);
    });

    setTimeout(() => {
      try {
        picker.showPicker();
      } catch (err) {
        cleanup();
        startInlineEdit(coord);
      }
    }, 10);
  }

  function updateFillHandlePreview(hoverCoord) {
    gridBody.querySelectorAll('td.autofill-preview-cell').forEach(el => el.classList.remove('autofill-preview-cell'));
    const pHover = parseCoord(hoverCoord);
    if (!pHover) return;
    fillPreviewTargetCoord = hoverCoord;

    const bounds = getSelectionBounds();
    const minR = Math.min(bounds.minRow, pHover.row);
    const maxR = Math.max(bounds.maxRow, pHover.row);
    const minC = Math.min(bounds.minCol, pHover.col);
    const maxC = Math.max(bounds.maxCol, pHover.col);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (r >= bounds.minRow && r <= bounds.maxRow && c >= bounds.minCol && c <= bounds.maxCol) continue;
        const td = gridBody.querySelector(`td[data-coord="${indexToCol(c)}${r}"]`);
        if (td) td.classList.add('autofill-preview-cell');
      }
    }
  }

  function commitFillHandleDrag() {
    if (!fillPreviewTargetCoord) return;
    const pHover = parseCoord(fillPreviewTargetCoord);
    fillPreviewTargetCoord = null;
    if (!pHover) return;

    const bounds = getSelectionBounds();
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;

    const srcRows = bounds.maxRow - bounds.minRow + 1;
    const srcCols = bounds.maxCol - bounds.minCol + 1;

    const minR = Math.min(bounds.minRow, pHover.row);
    const maxR = Math.max(bounds.maxRow, pHover.row);
    const minC = Math.min(bounds.minCol, pHover.col);
    const maxC = Math.max(bounds.maxCol, pHover.col);

    if (minR === bounds.minRow && maxR === bounds.maxRow && minC === bounds.minCol && maxC === bounds.maxCol) {
      return;
    }

    // Capture source values
    const srcMatrix = [];
    for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
      const rowArr = [];
      for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
        const coord = `${indexToCol(c)}${r}`;
        rowArr.push(sheet.cells[coord]?.v ?? '');
      }
      srcMatrix.push(rowArr);
    }

    pushUndoState();
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (r >= bounds.minRow && r <= bounds.maxRow && c >= bounds.minCol && c <= bounds.maxCol) continue;
        const destCoord = `${indexToCol(c)}${r}`;
        if (activeSheet === 'PO Nuevo' && destCoord === 'A1') continue;
        const sr = ((r - bounds.minRow) % srcRows + srcRows) % srcRows;
        const sc = ((c - bounds.minCol) % srcCols + srcCols) % srcCols;
        applyCellValueRaw(activeSheet, destCoord, srcMatrix[sr][sc]);
      }
    }

    engine.recalculateAll();
    if (activeSheet === 'DayOrders') populateDayOrdersFilters();
    renderSheet();
    selectedCoord = `${indexToCol(minC)}${minR}`;
    selectionEndCoord = `${indexToCol(maxC)}${maxR}`;
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
    showToast(`Autorrelleno aplicado (${indexToCol(minC)}${minR}:${indexToCol(maxC)}${maxR})`);
  }

  // Event Bindings
  function bindEvents() {
    // Column Header Click -> Select Entire Column
    gridHead.addEventListener('mousedown', (e) => {
      const cornerTh = e.target.closest('th.corner-header');
      if (cornerTh) {
        e.preventDefault();
        const sheet = workbook.sheets[activeSheet];
        const maxColStr = indexToCol(sheet.maxCol || 12);
        const maxR = getLastRealRow(activeSheet) || sheet.maxRow || 42;
        selectRange('A1', `${maxColStr}${maxR}`);
        return;
      }
      const colTh = e.target.closest('th[data-col]');
      if (!colTh) return;
      e.preventDefault();
      closeInlineEditor(true);
      const colStr = colTh.getAttribute('data-col');
      const sheet = workbook.sheets[activeSheet];
      const maxR = Math.max(getLastRealRow(activeSheet), sheet.maxRow || 42);
      if (e.shiftKey) {
        const startP = parseCoord(selectedCoord) || { colStr: 'A' };
        selectRange(`${startP.colStr}1`, `${colStr}${maxR}`);
      } else {
        selectRange(`${colStr}1`, `${colStr}${maxR}`);
      }
    });

    // Grid click, drag-to-select, fill handle drag & row header click
    gridBody.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;

      const fillHandle = e.target.closest('.cell-fill-handle');
      if (fillHandle) {
        e.preventDefault();
        e.stopPropagation();
        closeInlineEditor(true);
        isDraggingFillHandle = true;
        fillPreviewTargetCoord = null;
        return;
      }

      const rowTh = e.target.closest('th.row-header[data-row]');
      if (rowTh) {
        e.preventDefault();
        closeInlineEditor(true);
        const r = Number(rowTh.getAttribute('data-row'));
        const sheet = workbook.sheets[activeSheet];
        const maxColStr = indexToCol(sheet.maxCol || 12);
        if (e.shiftKey) {
          const startP = parseCoord(selectedCoord) || { row: r };
          selectRange(`A${startP.row}`, `${maxColStr}${r}`);
        } else {
          selectRange(`A${r}`, `${maxColStr}${r}`);
        }
        return;
      }

      const macroTrigger = e.target.closest('[data-macro]');
      if (macroTrigger) {
        const m = macroTrigger.getAttribute('data-macro');
        if (m === 'GUARDAR') ejecutarGuardarDatos();
        else if (m === 'PDF_PO') ejecutarGenerarOrdenPDF();
        else if (m === 'HOMOLOGADOS') abrirModalHomologeados();
        return;
      }

      const dateTrigger = e.target.closest('[data-datepicker]');
      if (dateTrigger) {
        e.preventDefault();
        e.stopPropagation();
        const coord = dateTrigger.getAttribute('data-datepicker');
        selectCell(coord, false);
        openCellDatePicker(coord);
        return;
      }

      const dropTrigger = e.target.closest('[data-dropdown]');
      if (dropTrigger) {
        e.preventDefault();
        e.stopPropagation();
        const coord = dropTrigger.getAttribute('data-dropdown');
        selectCell(coord, false);
        startInlineEdit(coord);
        return;
      }

      const td = e.target.closest('td[data-coord]');
      if (!td) return;
      const coord = td.getAttribute('data-coord');
      if (inlineEditor && inlineEditor.coord !== coord) {
        closeInlineEditor(true);
      }
      if (!inlineEditor) {
        e.preventDefault(); // Prevent native text selection while dragging cells
        isDraggingSelection = true;
        selectCell(coord, false, e.shiftKey);
      }
    });

    gridBody.addEventListener('mouseover', (e) => {
      if (!isDraggingSelection && !isDraggingFillHandle) return;
      const td = e.target.closest('td[data-coord]');
      if (!td) return;
      const coord = td.getAttribute('data-coord');
      if (isDraggingFillHandle) {
        updateFillHandlePreview(coord);
      } else if (isDraggingSelection) {
        if (selectionEndCoord !== coord) {
          selectCell(coord, false, true);
        }
      }
    });

    window.addEventListener('mouseup', () => {
      if (isDraggingFillHandle) {
        isDraggingFillHandle = false;
        commitFillHandleDrag();
      }
      if (isDraggingSelection) {
        isDraggingSelection = false;
      }
    });

    gridBody.addEventListener('dblclick', (e) => {
      const td = e.target.closest('td[data-coord]');
      if (!td) return;
      const coord = td.getAttribute('data-coord');
      if (td.hasAttribute('data-macro') || coord === 'A1') return;
      if (isDatePickerCell(activeSheet, coord)) {
        openCellDatePicker(coord);
        return;
      }
      startInlineEdit(coord);
    });

    // Validation list item click & filter
    validationSearchInput.addEventListener('input', () => {
      renderValidationList(validationSearchInput.value);
    });

    validationSearchInput.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        hideValidationDropdown();
      } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        moveValidationHighlight(e.key === 'ArrowDown' ? 1 : -1);
      } else if (e.key === 'Enter') {
        const active = validationList.querySelector('.validation-item.active') ||
                       validationList.querySelector('.validation-item');
        if (active && currentDropdownCoord) {
          const targetCoord = currentDropdownCoord;
          hideValidationDropdown();
          closeInlineEditor(false);
          commitValidationItem(targetCoord, active);
        }
      }
    });

    validationList.addEventListener('mousedown', (e) => {
      const item = e.target.closest('.validation-item');
      if (!item || !currentDropdownCoord) return;
      e.preventDefault();
      const targetCoord = currentDropdownCoord;
      hideValidationDropdown();
      closeInlineEditor(false);
      commitValidationItem(targetCoord, item);
    });

    // Formula Bar & Name Box (supports single coord like F11 or range like B22:I30)
    cellNameBox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        const raw = cellNameBox.value.trim().toUpperCase();
        if (raw.includes(':')) {
          const [startStr, endStr] = raw.split(':');
          selectRange(startStr, endStr);
        } else {
          selectCell(raw, false);
        }
        cellNameBox.blur();
      }
    });

    formulaInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        setCellContent(activeSheet, selectedCoord, formulaInput.value);
        moveSelection(1, 0);
      } else if (e.key === 'Escape') {
        selectCell(selectedCoord, false);
        formulaInput.blur();
      }
    });

    // System Clipboard Paste Listener (handles pasting from Excel, Google Sheets, or internal Ctrl+C)
    window.addEventListener('paste', (e) => {
      const activeTag = document.activeElement?.tagName;
      if (activeTag === 'INPUT' || activeTag === 'SELECT' || activeTag === 'TEXTAREA') return;
      if (document.getElementById('pdfModalOverlay').style.display !== 'none') return;
      if (document.getElementById('bomModalOverlay').style.display !== 'none') return;

      e.preventDefault();
      const pastedText = (e.clipboardData || window.clipboardData)?.getData('text/plain') || '';
      if (clipboardState && (!pastedText || pastedText === clipboardState.tsv)) {
        pasteDataMatrix(clipboardState.grid, true);
      } else if (pastedText) {
        const rows = pastedText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
        // Remove trailing empty newline that Excel often adds on copy
        if (rows.length > 1 && rows[rows.length - 1] === '') rows.pop();
        const matrix = rows.map(line => line.split('\t'));
        pasteDataMatrix(matrix, false);
      }
    });

    // Global Keyboard Navigation & Excel Shortcuts
    window.addEventListener('keydown', (e) => {
      const activeTag = document.activeElement?.tagName;
      if (activeTag === 'INPUT' || activeTag === 'SELECT' || activeTag === 'TEXTAREA') return;
      if (document.getElementById('pdfModalOverlay').style.display !== 'none') return;
      if (document.getElementById('bomModalOverlay').style.display !== 'none') return;

      const keyLow = e.key.toLowerCase();

      if ((e.ctrlKey || e.metaKey) && keyLow === 'z') {
        e.preventDefault();
        undo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'y') {
        e.preventDefault();
        redo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'c') {
        e.preventDefault();
        copySelection(false);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'x') {
        e.preventDefault();
        copySelection(true);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'v') {
        // If browser doesn't fire 'paste' event automatically on Ctrl+V without focused input, fallback to internal or async clipboard read
        if (navigator.clipboard && navigator.clipboard.readText) {
          e.preventDefault();
          navigator.clipboard.readText().then(txt => {
            if (clipboardState && (!txt || txt === clipboardState.tsv)) {
              pasteDataMatrix(clipboardState.grid, true);
            } else if (txt) {
              const rows = txt.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
              if (rows.length > 1 && rows[rows.length - 1] === '') rows.pop();
              const matrix = rows.map(line => line.split('\t'));
              pasteDataMatrix(matrix, false);
            } else if (clipboardState) {
              pasteDataMatrix(clipboardState.grid, true);
            }
          }).catch(() => {
            if (clipboardState) pasteDataMatrix(clipboardState.grid, true);
          });
          return;
        }
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'd') {
        e.preventDefault();
        fillSelectionDown();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'r') {
        e.preventDefault();
        fillSelectionRight();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'a') {
        e.preventDefault();
        const sheet = workbook.sheets[activeSheet];
        if (activeSheet === 'PO Nuevo') {
          selectRange('A22', 'I41');
        } else {
          const maxColStr = indexToCol(sheet.maxCol || 12);
          const maxR = getLastRealRow(activeSheet) || sheet.maxRow || 42;
          selectRange('A1', `${maxColStr}${maxR}`);
        }
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 's') {
        e.preventDefault();
        saveWorkbookState();
        showToast('Libro guardado correctamente');
        return;
      }

      if (e.key === 'Escape') {
        if (clipboardState) {
          clipboardState = null;
          highlightSelectedCellInDOM();
          showToast('Selección de copiado cancelada');
        } else {
          selectCell(selectedCoord, false, false);
        }
        return;
      }

      if (e.key === 'ArrowDown') { e.preventDefault(); moveSelection(1, 0, e.shiftKey); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); moveSelection(-1, 0, e.shiftKey); }
      else if (e.key === 'ArrowRight') { e.preventDefault(); moveSelection(0, 1, e.shiftKey); }
      else if (e.key === 'ArrowLeft') { e.preventDefault(); moveSelection(0, -1, e.shiftKey); }
      else if (e.key === 'Tab') { e.preventDefault(); moveSelection(0, e.shiftKey ? -1 : 1, false); }
      else if ((e.ctrlKey || e.metaKey) && (e.key === '+' || e.key === '=')) {
        e.preventDefault();
        const { minRow } = getSelectionBounds();
        insertRowsAt(minRow, 1);
      } else if ((e.ctrlKey || e.metaKey) && e.key === '-') {
        e.preventDefault();
        deleteSelectedRows();
      } else if (e.key === 'Enter') {
        e.preventDefault();
        startInlineEdit(selectedCoord);
      } else if (e.key === 'F2') {
        e.preventDefault();
        startInlineEdit(selectedCoord);
      } else if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        clearSelectedRange();
      } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        startInlineEdit(selectedCoord, e.key);
      }
    });

    // Auto-expand rows when scrolling near bottom of viewport (like Excel infinite scroll)
    const viewportEl = document.getElementById('spreadsheetViewport');
    if (viewportEl) {
      viewportEl.addEventListener('scroll', () => {
        if (viewportEl.scrollTop + viewportEl.clientHeight >= viewportEl.scrollHeight - 25) {
          const sheet = workbook?.sheets?.[activeSheet];
          if (sheet && (sheet.maxRow || 100) < 2000) {
            sheet.maxRow = (sheet.maxRow || 100) + 30;
            renderSheet();
          }
        }
      });
    }

    // Excel Right-Click Context Menu
    const ctxMenu = document.createElement('div');
    ctxMenu.className = 'excel-context-menu no-print';
    document.body.appendChild(ctxMenu);

    const hideCtxMenu = () => {
      ctxMenu.style.display = 'none';
    };

    window.addEventListener('mousedown', (e) => {
      if (!ctxMenu.contains(e.target)) hideCtxMenu();
    });

    gridBody.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      closeInlineEditor(true);

      const rowTh = e.target.closest('th.row-header[data-row]');
      const td = e.target.closest('td[data-coord]');
      if (rowTh) {
        const r = Number(rowTh.getAttribute('data-row'));
        const sheet = workbook.sheets[activeSheet];
        const maxColStr = indexToCol(sheet.maxCol || 16);
        selectRange(`A${r}`, `${maxColStr}${r}`);
      } else if (td) {
        const coord = td.getAttribute('data-coord');
        const p = parseCoord(coord);
        const b = getSelectionBounds();
        if (p && (p.row < b.minRow || p.row > b.maxRow || p.col < b.minCol || p.col > b.maxCol)) {
          selectCell(coord, false);
        }
      }

      const bounds = getSelectionBounds();
      const isPOTable = activeSheet === 'PO Nuevo' && bounds.minRow >= 22 && bounds.minRow <= getPONuevoTotalRow();
      const rowLabel = isPOTable ? 'de producto en PO' : `en fila ${bounds.minRow}`;

      ctxMenu.innerHTML = `
        <button class="excel-context-item" data-ctx="copy">
          <span>📋 Copiar</span><span class="excel-context-shortcut">Ctrl+C</span>
        </button>
        <button class="excel-context-item" data-ctx="cut">
          <span>✂️ Cortar</span><span class="excel-context-shortcut">Ctrl+X</span>
        </button>
        <button class="excel-context-item" data-ctx="paste">
          <span>📥 Pegar</span><span class="excel-context-shortcut">Ctrl+V</span>
        </button>
        <button class="excel-context-item" data-ctx="clear">
          <span>🧹 Borrar contenido</span><span class="excel-context-shortcut">Supr</span>
        </button>
        <div class="excel-context-divider"></div>
        <button class="excel-context-item" data-ctx="insert-above">
          <span>➕ Insertar 1 fila arriba (${rowLabel})</span><span class="excel-context-shortcut">Ctrl++</span>
        </button>
        <button class="excel-context-item" data-ctx="insert-below">
          <span>➕ Insertar 1 fila abajo (${rowLabel})</span>
        </button>
        <button class="excel-context-item" data-ctx="insert-5">
          <span>➕ Insertar 5 filas aquí</span>
        </button>
        <button class="excel-context-item danger" data-ctx="delete-row">
          <span>🗑️ Eliminar fila(s) seleccionada(s)</span><span class="excel-context-shortcut">Ctrl+-</span>
        </button>
        <div class="excel-context-divider"></div>
        <button class="excel-context-item" data-ctx="add-50-rows">
          <span>⬇️ Añadir 50 filas más al final</span>
        </button>
        <button class="excel-context-item" data-ctx="add-5-cols">
          <span>➡️ Añadir 5 columnas más a la derecha</span>
        </button>
      `;

      ctxMenu.style.display = 'flex';
      const menuW = 270;
      const menuH = 340;
      ctxMenu.style.left = `${Math.min(window.innerWidth - menuW - 10, e.clientX)}px`;
      ctxMenu.style.top = `${Math.min(window.innerHeight - menuH - 10, e.clientY)}px`;
    });

    ctxMenu.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-ctx]');
      if (!btn) return;
      const action = btn.getAttribute('data-ctx');
      hideCtxMenu();
      const bounds = getSelectionBounds();

      if (action === 'copy') copySelection(false);
      else if (action === 'cut') copySelection(true);
      else if (action === 'paste') {
        if (navigator.clipboard && navigator.clipboard.readText) {
          navigator.clipboard.readText().then(txt => {
            if (clipboardState && (!txt || txt === clipboardState.tsv)) {
              pasteDataMatrix(clipboardState.grid, true);
            } else if (txt) {
              const rows = txt.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
              if (rows.length > 1 && rows[rows.length - 1] === '') rows.pop();
              pasteDataMatrix(rows.map(line => line.split('\t')), false);
            } else if (clipboardState) {
              pasteDataMatrix(clipboardState.grid, true);
            }
          }).catch(() => {
            if (clipboardState) pasteDataMatrix(clipboardState.grid, true);
          });
        } else if (clipboardState) {
          pasteDataMatrix(clipboardState.grid, true);
        }
      } else if (action === 'clear') clearSelectedRange();
      else if (action === 'insert-above') insertRowsAt(bounds.minRow, 1);
      else if (action === 'insert-below') insertRowsAt(bounds.maxRow + 1, 1);
      else if (action === 'insert-5') insertRowsAt(bounds.minRow, 5);
      else if (action === 'delete-row') deleteSelectedRows();
      else if (action === 'add-50-rows') expandGridDimensions(50, 0);
      else if (action === 'add-5-cols') expandGridDimensions(0, 5);
    });

    // Sheet Tabs
    document.querySelectorAll('.sheet-tab').forEach(btn => {
      btn.addEventListener('click', () => {
        switchSheet(btn.getAttribute('data-sheet'));
      });
    });

    const sheetOrder = ['PO Nuevo', 'DayOrders', 'Productos', 'Transporte', 'PASSING SHEET'];
    document.getElementById('btnPrevSheet').addEventListener('click', () => {
      const idx = sheetOrder.indexOf(activeSheet);
      if (idx > 0) switchSheet(sheetOrder[idx - 1]);
    });
    document.getElementById('btnNextSheet').addEventListener('click', () => {
      const idx = sheetOrder.indexOf(activeSheet);
      if (idx < sheetOrder.length - 1) switchSheet(sheetOrder[idx + 1]);
    });

    // Utility buttons
    document.getElementById('btnUndo').addEventListener('click', undo);
    document.getElementById('btnRedo').addEventListener('click', redo);
    document.getElementById('btnVisualGenerarOrden').addEventListener('click', () => {
      if (ejecutarGenerarOrdenPDF()) void generarArchivoPosco();
    });
    document.getElementById('btnVisualGenerarPassing').addEventListener('click', () => {
      ejecutarPdfPassingSheet();
    });
    document.getElementById('btnSendOrder').addEventListener('click', () => {
      void enviarOrdenPorCorreo();
    });
    document.getElementById('btnSaveManual')?.addEventListener('click', () => {
      saveWorkbookState();
      showToast('Cambios guardados');
    });
    document.getElementById('btnExportCsv').addEventListener('click', exportActiveSheetCsv);

    document.getElementById('btnResetWorkbook').addEventListener('click', async () => {
      if (!confirm('¿Restaurar el libro exactamente al estado original del archivo .xlsm?')) return;
      localStorage.removeItem('wholesale_workbook_backup');
      const res = await fetch('/api/reset', { method: 'POST' });
      workbook = await res.json();
      engine = new window.FormulaEngine(workbook);
      engine.recalculateAll();
      populateDayOrdersFilters();
      renderSheet();
      selectCell('F11', false);
      showToast('Libro restaurado al archivo .xlsm original');
    });

    // Zoom controls
    document.getElementById('btnZoomIn').addEventListener('click', () => {
      zoomLevel = Math.min(140, zoomLevel + 10);
      sheetCanvasWrapper.style.zoom = `${zoomLevel}%`;
      document.getElementById('zoomLabel').textContent = `${zoomLevel}%`;
    });
    document.getElementById('btnZoomOut').addEventListener('click', () => {
      zoomLevel = Math.max(70, zoomLevel - 10);
      sheetCanvasWrapper.style.zoom = `${zoomLevel}%`;
      document.getElementById('zoomLabel').textContent = `${zoomLevel}%`;
    });

    // Modals close & print
    document.getElementById('btnClosePdfModal').addEventListener('click', () => {
      document.getElementById('pdfModalOverlay').style.display = 'none';
    });
    document.getElementById('btnCloseBomModal').addEventListener('click', () => {
      document.getElementById('bomModalOverlay').style.display = 'none';
    });
    document.getElementById('btnPrintPdfNow').addEventListener('click', () => {
      const prevTitle = document.title;
      const customName = document.getElementById('pdfFilenameInput').value.replace(/\.pdf$/i, '');
      document.title = customName || prevTitle;
      window.print();
      setTimeout(() => { document.title = prevTitle; }, 500);
    });
  }

  init();
})();
