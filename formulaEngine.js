// ============================================================================
// EXCEL FORMULA ENGINE - Pacifica Produce Farms Wholesale Workbook
// Supports cross-sheet references, VLOOKUP, INDEX, MATCH, array criteria MATCH,
// IF, IFERROR, ISERROR, COUNTIF, SUM, AVERAGE, MIN, MAX, date arithmetic,
// and standard arithmetic/comparison expressions.
// ============================================================================

(function (global) {
  function colToIndex(colStr) {
    let idx = 0;
    const clean = colStr.replace(/\$/g, '').toUpperCase();
    for (let i = 0; i < clean.length; i++) {
      idx = idx * 26 + (clean.charCodeAt(i) - 64);
    }
    return idx; // 1-based
  }

  function indexToCol(idx) {
    let col = '';
    let n = idx;
    while (n > 0) {
      const rem = (n - 1) % 26;
      col = String.fromCharCode(65 + rem) + col;
      n = Math.floor((n - 1) / 26);
    }
    return col;
  }

  function parseCoord(coordStr) {
    const m = String(coordStr).trim().replace(/\$/g, '').match(/^([A-Z]+)(\d+)$/i);
    if (!m) return null;
    return {
      colStr: m[1].toUpperCase(),
      col: colToIndex(m[1]),
      row: parseInt(m[2], 10),
      coord: `${m[1].toUpperCase()}${parseInt(m[2], 10)}`
    };
  }

  // Split top-level arguments of a function call: "a, b(1,2), 'c,d'" -> ["a", "b(1,2)", "'c,d'"]
  function splitArgs(argsStr) {
    const args = [];
    let cur = '';
    let depth = 0;
    let inQuotes = false;
    for (let i = 0; i < argsStr.length; i++) {
      const ch = argsStr[i];
      if (ch === '"') {
        inQuotes = !inQuotes;
        cur += ch;
      } else if (!inQuotes) {
        if (ch === '(') {
          depth++;
          cur += ch;
        } else if (ch === ')') {
          depth--;
          cur += ch;
        } else if (ch === ',' && depth === 0) {
          args.push(cur.trim());
          cur = '';
        } else {
          cur += ch;
        }
      } else {
        cur += ch;
      }
    }
    if (cur.trim() !== '' || argsStr.endsWith(',')) {
      args.push(cur.trim());
    }
    return args;
  }

  function normalizeString(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim().toUpperCase();
  }

  function isIsoDate(str) {
    return typeof str === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(str.trim());
  }

  function addDaysToIsoDate(isoStr, days) {
    const parts = isoStr.trim().split('-').map(Number);
    const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    dt.setUTCDate(dt.getUTCDate() + days);
    const y = dt.getUTCFullYear();
    const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const d = String(dt.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  class FormulaEngine {
    constructor(workbook) {
      this.workbook = workbook;
      this.memo = new Map();
      this.visiting = new Set();
    }

    recalculateAll() {
      this.memo.clear();
      this.visiting.clear();
      const sheetNames = Object.keys(this.workbook.sheets);
      for (const sName of sheetNames) {
        const sheet = this.workbook.sheets[sName];
        for (const [coord, cell] of Object.entries(sheet.cells)) {
          if (cell && cell.f) {
            cell.v = this.getCellValue(sName, coord);
          }
        }
      }
    }

    getCellValue(sheetName, coord) {
      const cleanCoord = coord.replace(/\$/g, '').toUpperCase();
      const key = `${sheetName}!${cleanCoord}`;
      if (this.memo.has(key)) return this.memo.get(key);

      const sheet = this.workbook.sheets[sheetName];
      if (!sheet || !sheet.cells[cleanCoord]) return '';

      const cell = sheet.cells[cleanCoord];
      if (!cell.f) {
        return cell.v !== undefined && cell.v !== null ? cell.v : '';
      }

      if (this.visiting.has(key)) {
        return '#REF!';
      }

      this.visiting.add(key);
      let result = '';
      try {
        result = this.evaluateFormula(cell.f, sheetName);
        if (result && typeof result === 'object' && result.__error) {
          result = result.code || '#VALUE!';
        }
      } catch (e) {
        result = '#ERROR!';
      }
      this.visiting.delete(key);
      this.memo.set(key, result);
      cell.v = result;
      return result;
    }

    // Parse range reference like "Productos!$E$1:$E$234", "Transporte!A:L", "H22:H41", "Productos!F:F"
    parseRange(rangeExpr, currentSheet) {
      let sheetName = currentSheet;
      let ref = rangeExpr.trim();
      if (ref.includes('!')) {
        const bangIdx = ref.indexOf('!');
        sheetName = ref.slice(0, bangIdx).replace(/^'|'$/g, '');
        ref = ref.slice(bangIdx + 1);
      }
      ref = ref.replace(/\$/g, '');
      const sheet = this.workbook.sheets[sheetName];
      const maxRow = sheet ? (sheet.maxRow || 300) : 300;

      // Full column range e.g. A:L or F:F
      const colRangeMatch = ref.match(/^([A-Z]+):([A-Z]+)$/i);
      if (colRangeMatch) {
        return {
          sheetName,
          startCol: colToIndex(colRangeMatch[1]),
          endCol: colToIndex(colRangeMatch[2]),
          startRow: 1,
          endRow: maxRow
        };
      }

      // Cell range e.g. E1:E234 or H22:H41
      const cellRangeMatch = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
      if (cellRangeMatch) {
        return {
          sheetName,
          startCol: colToIndex(cellRangeMatch[1]),
          startRow: parseInt(cellRangeMatch[2], 10),
          endCol: colToIndex(cellRangeMatch[3]),
          endRow: parseInt(cellRangeMatch[4], 10)
        };
      }

      // Single cell
      const single = parseCoord(ref);
      if (single) {
        return {
          sheetName,
          startCol: single.col,
          startRow: single.row,
          endCol: single.col,
          endRow: single.row
        };
      }
      return null;
    }

    evaluateFormula(rawFormula, currentSheet) {
      if (typeof rawFormula !== 'string') return rawFormula;
      let expr = rawFormula.trim();
      if (expr.startsWith('=+')) {
        expr = expr.slice(2).trim();
      } else if (expr.startsWith('=')) {
        expr = expr.slice(1).trim();
      }
      return this.evalExpr(expr, currentSheet);
    }

    evalExpr(expr, currentSheet) {
      expr = expr.trim();
      if (expr === '') return '';

      // Strip outer balanced parentheses if they wrap the whole expression
      while (expr.startsWith('(') && expr.endsWith(')')) {
        let depth = 0;
        let balanced = true;
        let inQuotes = false;
        for (let i = 0; i < expr.length; i++) {
          if (expr[i] === '"') inQuotes = !inQuotes;
          if (!inQuotes) {
            if (expr[i] === '(') depth++;
            else if (expr[i] === ')') depth--;
            if (depth === 0 && i < expr.length - 1) {
              balanced = false;
              break;
            }
          }
        }
        if (balanced) expr = expr.slice(1, -1).trim();
        else break;
      }

      // String literal
      if (expr.startsWith('"') && expr.endsWith('"') && expr.slice(1, -1).indexOf('"') === -1) {
        return expr.slice(1, -1);
      }

      // Booleans
      if (/^TRUE$/i.test(expr)) return true;
      if (/^FALSE$/i.test(expr)) return false;

      // Numeric literal
      if (/^-?\d+(\.\d+)?$/.test(expr)) {
        return Number(expr);
      }

      // Check top-level comparison operators (=, <>, <=, >=, <, >) outside parens/quotes
      const compOp = this.findTopLevelOperator(expr, ['<>', '<=', '>=', '=', '<', '>']);
      if (compOp) {
        const left = this.evalExpr(compOp.left, currentSheet);
        const right = this.evalExpr(compOp.right, currentSheet);
        if (left && left.__error) return left;
        if (right && right.__error) return right;
        return this.compareValues(left, right, compOp.op);
      }

      // Check top-level string concatenation (&)
      const concatOp = this.findTopLevelOperator(expr, ['&']);
      if (concatOp) {
        const left = this.evalExpr(concatOp.left, currentSheet);
        const right = this.evalExpr(concatOp.right, currentSheet);
        return `${left ?? ''}${right ?? ''}`;
      }

      // Check top-level addition / subtraction (+, -) from right to left
      const addSubOp = this.findTopLevelAddSub(expr);
      if (addSubOp) {
        const left = this.evalExpr(addSubOp.left, currentSheet);
        const right = this.evalExpr(addSubOp.right, currentSheet);
        if (left && left.__error) return left;
        if (right && right.__error) return right;

        // Date arithmetic support (e.g. =F19-1)
        if (isIsoDate(left) && typeof right === 'number') {
          return addDaysToIsoDate(left, addSubOp.op === '+' ? right : -right);
        }
        const lNum = Number(left || 0);
        const rNum = Number(right || 0);
        if (isNaN(lNum) || isNaN(rNum)) return { __error: true, code: '#VALUE!' };
        return addSubOp.op === '+' ? lNum + rNum : lNum - rNum;
      }

      // Check top-level multiplication / division (*, /)
      const mulDivOp = this.findTopLevelOperator(expr, ['*', '/'], true);
      if (mulDivOp) {
        const left = this.evalExpr(mulDivOp.left, currentSheet);
        const right = this.evalExpr(mulDivOp.right, currentSheet);
        if (left && left.__error) return left;
        if (right && right.__error) return right;
        const lNum = Number(left || 0);
        const rNum = Number(right || 0);
        if (isNaN(lNum) || isNaN(rNum)) return { __error: true, code: '#VALUE!' };
        if (mulDivOp.op === '/') {
          if (rNum === 0) return { __error: true, code: '#DIV/0!' };
          return lNum / rNum;
        }
        return Math.round(lNum * rNum * 1e6) / 1e6;
      }

      // Function call e.g. IF(...), VLOOKUP(...), INDEX(...), _xlfn.XLOOKUP(...)
      const fnMatch = expr.match(/^([A-Za-z0-9_.]+)\(([\s\S]*)\)$/);
      if (fnMatch) {
        const fnName = fnMatch[1].replace(/^_xlfn\./i, '').toUpperCase();
        const argsStr = fnMatch[2];
        return this.evalFunction(fnName, argsStr, currentSheet);
      }

      // Cross-sheet cell reference e.g. 'PO Nuevo'!A12 or Productos!$F$1
      if (expr.includes('!')) {
        const bangIdx = expr.indexOf('!');
        const targetSheet = expr.slice(0, bangIdx).replace(/^'|'$/g, '');
        const targetCoord = expr.slice(bangIdx + 1).replace(/\$/g, '');
        if (parseCoord(targetCoord)) {
          return this.getCellValue(targetSheet, targetCoord);
        }
      }

      // Local cell reference e.g. A9, $A$9, F11
      const parsedCell = parseCoord(expr);
      if (parsedCell) {
        return this.getCellValue(currentSheet, parsedCell.coord);
      }

      return expr;
    }

    findTopLevelOperator(expr, ops, rightToLeft = false) {
      let depth = 0;
      let inQuotes = false;
      const indices = [];
      for (let i = 0; i < expr.length; i++) {
        const ch = expr[i];
        if (ch === '"') inQuotes = !inQuotes;
        else if (!inQuotes) {
          if (ch === '(') depth++;
          else if (ch === ')') depth--;
          else if (depth === 0) {
            for (const op of ops) {
              if (expr.slice(i, i + op.length) === op) {
                indices.push({ idx: i, op });
                i += op.length - 1;
                break;
              }
            }
          }
        }
      }
      if (indices.length === 0) return null;
      const chosen = rightToLeft ? indices[indices.length - 1] : indices[0];
      return {
        op: chosen.op,
        left: expr.slice(0, chosen.idx),
        right: expr.slice(chosen.idx + chosen.op.length)
      };
    }

    findTopLevelAddSub(expr) {
      let depth = 0;
      let inQuotes = false;
      let lastMatch = null;
      for (let i = 0; i < expr.length; i++) {
        const ch = expr[i];
        if (ch === '"') inQuotes = !inQuotes;
        else if (!inQuotes) {
          if (ch === '(') depth++;
          else if (ch === ')') depth--;
          else if (depth === 0 && (ch === '+' || ch === '-') && i > 0) {
            const prev = expr.slice(0, i).trim();
            if (prev && !/[+\-*/=(<>^,]$/.test(prev)) {
              lastMatch = { idx: i, op: ch };
            }
          }
        }
      }
      if (!lastMatch) return null;
      return {
        op: lastMatch.op,
        left: expr.slice(0, lastMatch.idx),
        right: expr.slice(lastMatch.idx + 1)
      };
    }

    compareValues(a, b, op) {
      const aEmpty = a === '' || a === null || a === undefined;
      const bEmpty = b === '' || b === null || b === undefined;
      if (aEmpty && bEmpty) {
        return op === '=' || op === '<=' || op === '>=';
      }
      if (typeof a === 'number' && typeof b === 'number') {
        if (op === '=') return a === b;
        if (op === '<>') return a !== b;
        if (op === '<') return a < b;
        if (op === '>') return a > b;
        if (op === '<=') return a <= b;
        if (op === '>=') return a >= b;
      }
      const sa = normalizeString(a);
      const sb = normalizeString(b);
      if (op === '=') return sa === sb;
      if (op === '<>') return sa !== sb;
      if (op === '<') return sa < sb;
      if (op === '>') return sa > sb;
      if (op === '<=') return sa <= sb;
      if (op === '>=') return sa >= sb;
      return false;
    }

    evalFunction(fnName, argsStr, currentSheet) {
      const args = splitArgs(argsStr);

      if (fnName === 'IF') {
        const cond = this.evalExpr(args[0] || '', currentSheet);
        if (cond && cond.__error) return cond;
        if (Boolean(cond)) {
          return args[1] !== undefined ? this.evalExpr(args[1], currentSheet) : true;
        } else {
          return args[2] !== undefined ? this.evalExpr(args[2], currentSheet) : false;
        }
      }

      if (fnName === 'ISERROR') {
        const val = this.evalExpr(args[0] || '', currentSheet);
        return Boolean(val && typeof val === 'object' && val.__error) ||
          (typeof val === 'string' && val.startsWith('#'));
      }

      if (fnName === 'IFERROR') {
        const val = this.evalExpr(args[0] || '', currentSheet);
        const isErr = Boolean(val && typeof val === 'object' && val.__error) ||
          (typeof val === 'string' && val.startsWith('#'));
        return isErr ? this.evalExpr(args[1] || '""', currentSheet) : val;
      }

      if (fnName === 'SUM' || fnName === 'AVERAGE' || fnName === 'COUNT' || fnName === 'COUNTA') {
        const nums = [];
        let nonEmpty = 0;
        for (const arg of args) {
          const rng = this.parseRange(arg, currentSheet);
          if (rng) {
            for (let r = rng.startRow; r <= rng.endRow; r++) {
              for (let c = rng.startCol; c <= rng.endCol; c++) {
                const v = this.getCellValue(rng.sheetName, `${indexToCol(c)}${r}`);
                if (v !== '' && v !== null && v !== undefined) {
                  nonEmpty++;
                  const n = Number(v);
                  if (!isNaN(n) && typeof v !== 'boolean') nums.push(n);
                }
              }
            }
          } else {
            const v = this.evalExpr(arg, currentSheet);
            if (v !== '' && v !== null && v !== undefined) {
              nonEmpty++;
              const n = Number(v);
              if (!isNaN(n)) nums.push(n);
            }
          }
        }
        if (fnName === 'COUNTA') return nonEmpty;
        if (fnName === 'COUNT') return nums.length;
        const sum = nums.reduce((acc, x) => acc + x, 0);
        if (fnName === 'SUM') return Math.round(sum * 1e6) / 1e6;
        if (fnName === 'AVERAGE') return nums.length ? sum / nums.length : 0;
      }

      if (fnName === 'COUNTIF') {
        const rng = this.parseRange(args[0] || '', currentSheet);
        const critVal = this.evalExpr(args[1] || '', currentSheet);
        if (!rng) return 0;
        const targetNorm = normalizeString(critVal);
        let count = 0;
        for (let r = rng.startRow; r <= rng.endRow; r++) {
          for (let c = rng.startCol; c <= rng.endCol; c++) {
            const v = this.getCellValue(rng.sheetName, `${indexToCol(c)}${r}`);
            if (normalizeString(v) === targetNorm && targetNorm !== '') {
              count++;
            }
          }
        }
        return count;
      }

      if (fnName === 'VLOOKUP') {
        const lookupVal = this.evalExpr(args[0] || '', currentSheet);
        const rng = this.parseRange(args[1] || '', currentSheet);
        const colIdx = Number(this.evalExpr(args[2] || '1', currentSheet));
        if (!rng || lookupVal === '' || lookupVal === null || lookupVal === undefined) {
          return { __error: true, code: '#N/A' };
        }
        const targetNorm = normalizeString(lookupVal);
        for (let r = rng.startRow; r <= rng.endRow; r++) {
          const cellVal = this.getCellValue(rng.sheetName, `${indexToCol(rng.startCol)}${r}`);
          if (normalizeString(cellVal) === targetNorm) {
            const targetCol = rng.startCol + colIdx - 1;
            const res = this.getCellValue(rng.sheetName, `${indexToCol(targetCol)}${r}`);
            return res === '' ? 0 : res;
          }
        }
        return { __error: true, code: '#N/A' };
      }

      if (fnName === 'MATCH') {
        const lookupVal = this.evalExpr(args[0] || '', currentSheet);
        const lookupArrayExpr = (args[1] || '').trim();

        // Special array formula pattern in PO Nuevo col G:
        // MATCH(1, INDEX((Productos!$K$1:$K$234=F22)*(Productos!$L$1:$L$234<>""), 0), 0)
        const boxDimMatch = lookupArrayExpr.match(
          /^INDEX\(\(([^=]+)=([A-Z]+\d+)\)\*\(([^<]+ font?)<>""\),\s*0\)$/i
        ) || lookupArrayExpr.match(
          /^INDEX\(\(([^=]+)=([A-Za-z0-9$]+)\)\*\(([^<]+)\s*<>\s*""\),\s*0\)$/i
        );
        if (boxDimMatch) {
          const kRng = this.parseRange(boxDimMatch[1], currentSheet);
          const fVal = this.evalExpr(boxDimMatch[2], currentSheet);
          const lRng = this.parseRange(boxDimMatch[3], currentSheet);
          if (!kRng || !lRng || !fVal) return { __error: true, code: '#N/A' };
          const targetBox = normalizeString(fVal);
          for (let r = kRng.startRow; r <= kRng.endRow; r++) {
            const kVal = normalizeString(this.getCellValue(kRng.sheetName, `${indexToCol(kRng.startCol)}${r}`));
            const lVal = this.getCellValue(lRng.sheetName, `${indexToCol(lRng.startCol)}${r}`);
            if (kVal === targetBox && lVal !== '' && lVal !== null && lVal !== undefined) {
              return r - kRng.startRow + 1;
            }
          }
          // Fallback: if F22 is e.g. F4P or F4, also check prefix match if exact not found
          return { __error: true, code: '#N/A' };
        }

        const rng = this.parseRange(lookupArrayExpr, currentSheet);
        if (!rng || lookupVal === '' || lookupVal === null || lookupVal === undefined) {
          return { __error: true, code: '#N/A' };
        }
        const targetNorm = normalizeString(lookupVal);
        for (let r = rng.startRow; r <= rng.endRow; r++) {
          const v = this.getCellValue(rng.sheetName, `${indexToCol(rng.startCol)}${r}`);
          if (normalizeString(v) === targetNorm) {
            return r - rng.startRow + 1;
          }
        }
        return { __error: true, code: '#N/A' };
      }

      if (fnName === 'INDEX') {
        const rng = this.parseRange(args[0] || '', currentSheet);
        const rowOffset = this.evalExpr(args[1] || '1', currentSheet);
        if (!rng || (rowOffset && rowOffset.__error) || typeof rowOffset !== 'number' || rowOffset < 1) {
          return { __error: true, code: '#N/A' };
        }
        const colOffset = args[2] !== undefined ? Number(this.evalExpr(args[2], currentSheet)) : 1;
        const targetRow = rng.startRow + rowOffset - 1;
        const targetCol = rng.startCol + (colOffset > 0 ? colOffset - 1 : 0);
        const val = this.getCellValue(rng.sheetName, `${indexToCol(targetCol)}${targetRow}`);
        return val;
      }

      return '';
    }
  }

  global.FormulaEngine = FormulaEngine;
  global.ExcelUtils = { colToIndex, indexToCol, parseCoord };
})(typeof window !== 'undefined' ? window : globalThis);
