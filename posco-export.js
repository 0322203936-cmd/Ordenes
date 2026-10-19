const ExcelJS = require('exceljs');

const POSCO_HEADERS = [
  'Distribuidor',
  'Cust Name',
  'Descripcion',
  'No.',
  'CATEGORIA',
  'FLOR',
  'Flor Color',
  'Load Date',
  '#cajas',
  'Pack',
  'Stems',
  'Fecha Produccion',
  'Caja'
];

const PRODUCT_HEADERS = [
  'Distribuidor',
  'Cust Name',
  'CATEGORIA',
  'FLOR',
  'Flor Color',
  'Descripcion',
  'Ramos/Cajas',
  'txr_orden',
  'Caja'
];

function clean(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function cellValue(sheet, coord) {
  const cell = sheet?.cells?.[coord];
  if (cell && typeof cell === 'object' && Object.prototype.hasOwnProperty.call(cell, 'v')) {
    return cell.v;
  }
  return cell ?? '';
}

function normalize(value) {
  return clean(value).toLocaleLowerCase();
}

function similarity(left, right) {
  const a = normalize(left);
  const b = normalize(right);
  if (!a && !b) return 1;
  if (!a || !b) return 0;
  const previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let row = 1; row <= a.length; row += 1) {
    let diagonal = previous[0];
    previous[0] = row;
    for (let column = 1; column <= b.length; column += 1) {
      const above = previous[column];
      const cost = a[row - 1] === b[column - 1] ? 0 : 1;
      previous[column] = Math.min(previous[column] + 1, previous[column - 1] + 1, diagonal + cost);
      diagonal = above;
    }
  }
  return 1 - previous[b.length] / Math.max(a.length, b.length);
}

function parseDate(value, label) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    if (!Number.isNaN(date.getTime())) return date;
  }

  const text = clean(value);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
    if (!Number.isNaN(date.getTime())) return date;
  }

  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) {
    return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  }
  throw new Error(`${label} está vacío o no es una fecha válida en PO Nuevo.`);
}

function parsePositiveInteger(value, label) {
  const text = clean(value);
  const match = text.match(/^(\d+(?:\.\d+)?)\s*(?:st(?:em(?:s)?)?|bu(?:nch(?:es)?)?|pk|box)?$/i);
  const number = Number(match ? match[1] : value);
  if (!Number.isInteger(number) || number <= 0) {
    throw new Error(`${label} debe ser un entero positivo; llegó ${JSON.stringify(value)}.`);
  }
  return number;
}

function carrierAbbreviation(value) {
  const carrier = clean(value);
  if (!carrier) throw new Error('Truck Line está vacío en PO Nuevo.');
  if (normalize(carrier) === 'floral trade distributors') return 'FTD';

  const name = carrier.split(/\s+[-–—]\s+/, 1)[0];
  const words = name.match(/[A-Za-z]+/g) || [];
  if (words.length >= 2 && words[0].length >= 2) {
    return `${words[0].slice(0, 2)}${words[1][0]}`.toUpperCase();
  }
  if (!words[0] || words[0].length < 3) {
    throw new Error(`Truck Line ${JSON.stringify(carrier)} no permite obtener una abreviación.`);
  }
  return words[0].slice(0, 3).toUpperCase();
}

function isBomPending(value) {
  return ['pendiente de bom', 'pendiente de boom'].includes(normalize(value));
}

function boomRecommendation(flower, color, stems, bunches, box) {
  return [
    clean(flower),
    clean(color),
    clean(stems) ? `STEM/${clean(stems)}` : '',
    clean(bunches) ? `BU/${clean(bunches)}` : '',
    clean(box)
  ].filter(Boolean).join(' ') || 'Producto Wholesale';
}

function productLookup(productsSheet) {
  const headings = 'ABCDEFGHI'.split('').map(column => clean(cellValue(productsSheet, `${column}1`)));
  if (headings.some((heading, index) => heading !== PRODUCT_HEADERS[index])) {
    throw new Error('La hoja Productos no tiene las columnas A:I esperadas para POSCO.');
  }

  const lookup = new Map();
  const maxRow = Number(productsSheet.maxRow || 300);
  for (let row = 2; row <= maxRow; row += 1) {
    const description = clean(cellValue(productsSheet, `F${row}`));
    if (!description) continue;
    const item = {
      categoria: clean(cellValue(productsSheet, `C${row}`)),
      flor: clean(cellValue(productsSheet, `D${row}`)),
      color: clean(cellValue(productsSheet, `E${row}`)),
      descripcion: description,
      caja: clean(cellValue(productsSheet, `I${row}`))
    };
    const key = normalize(description);
    if (!lookup.has(key)) lookup.set(key, []);
    lookup.get(key).push(item);
  }
  if (lookup.size === 0) throw new Error('La hoja Productos no contiene productos homologados.');
  return lookup;
}

function selectProduct(lookup, description, color, box) {
  const normalized = normalize(description);
  let candidates = lookup.get(normalized) || [];
  if (candidates.length === 0) {
    candidates = [...lookup.values()]
      .flat()
      .filter(item => similarity(item.descripcion, description) >= 0.92);
  }
  if (candidates.length === 0) return null;

  const exactBox = candidates.filter(item => normalize(item.caja) === normalize(box));
  const byBox = exactBox.length ? exactBox : candidates;
  const exactColor = color
    ? byBox.filter(item => normalize(item.color) === normalize(color))
    : [];
  const selected = exactColor.length ? exactColor : byBox;
  return selected.length === 1 ? selected[0] : null;
}

function findTotalRow(orderSheet) {
  if (Number(orderSheet.poTotalRow) >= 23) return Number(orderSheet.poTotalRow);
  const maxRow = Number(orderSheet.maxRow || 100);
  for (let row = 23; row <= maxRow; row += 1) {
    if (normalize(cellValue(orderSheet, `G${row}`)) === 'total') return row;
  }
  return 42;
}

async function createPoscoWorkbook(workbook) {
  const orderSheet = workbook?.sheets?.['PO Nuevo'];
  const productsSheet = workbook?.sheets?.Productos;
  if (!orderSheet || !productsSheet) {
    throw new Error('El libro debe incluir las hojas PO Nuevo y Productos.');
  }

  const poNumber = clean(cellValue(orderSheet, 'A9'));
  const customer = clean(cellValue(orderSheet, 'F11'));
  const truckLine = clean(cellValue(orderSheet, 'F15'));
  const loadDate = parseDate(cellValue(orderSheet, 'B19'), 'Load Date');
  if (!poNumber || !customer || !truckLine) {
    throw new Error('PO Nuevo requiere PO, Customer Ship y Truck Line para crear el formato POSCO.');
  }

  const lookup = productLookup(productsSheet);
  const totalRow = findTotalRow(orderSheet);
  const rows = [];
  const boomPending = [];

  for (let row = 22; row < totalRow; row += 1) {
    const descriptionInput = clean(cellValue(orderSheet, `B${row}`));
    if (!descriptionInput) continue;

    const colorInput = clean(cellValue(orderSheet, `C${row}`));
    const stemsInput = cellValue(orderSheet, `D${row}`);
    const bunchesInput = cellValue(orderSheet, `E${row}`);
    const boxInput = clean(cellValue(orderSheet, `F${row}`));
    const casesInput = cellValue(orderSheet, `H${row}`);
    const product = selectProduct(lookup, descriptionInput, colorInput, boxInput);
    const stems = parsePositiveInteger(stemsInput, `STEM/BCH de la fila ${row}`);
    const bunches = parsePositiveInteger(bunchesInput, `BU/BOX de la fila ${row}`);
    const cases = parsePositiveInteger(casesInput, `CASES de la fila ${row}`);

    if (isBomPending(cellValue(orderSheet, `J${row}`))) {
      boomPending.push({
        row,
        product: descriptionInput,
        recommendation: boomRecommendation(descriptionInput, colorInput, stemsInput, bunchesInput, boxInput)
      });
    }

    const description = product?.descripcion || descriptionInput;
    const category = product?.categoria || 'CB / BULK';
    const flower = product?.flor || descriptionInput;
    const color = product?.color || colorInput;
    const box = product?.caja || boxInput;
    const productionDate = new Date(loadDate.getTime() - 3 * 86400000);

    rows.push([
      'PACIFICA PRODUCE FARMS',
      'WHOLESALE',
      description,
      `${customer} - ${carrierAbbreviation(truckLine)} ${poNumber}`,
      category,
      flower,
      color,
      loadDate,
      cases,
      bunches,
      stems,
      productionDate,
      box
    ]);
  }

  if (rows.length === 0) {
    throw new Error('PO Nuevo no contiene productos para crear el formato POSCO.');
  }

  const output = new ExcelJS.Workbook();
  output.creator = 'ORDENES WHOLESALE';
  output.created = new Date();
  const sheet = output.addWorksheet('New Order', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });
  sheet.addRow(POSCO_HEADERS);
  rows.forEach(row => sheet.addRow(row));
  sheet.autoFilter = { from: 'A1', to: `M${rows.length + 1}` };
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.paperSize = 1;

  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF14532D' } };
  sheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle' };
  sheet.getRow(1).height = 22;
  sheet.columns = [
    { width: 26 }, { width: 16 }, { width: 48 }, { width: 30 }, { width: 16 },
    { width: 24 }, { width: 18 }, { width: 14 }, { width: 10 }, { width: 10 },
    { width: 10 }, { width: 18 }, { width: 12 }
  ];
  for (let row = 2; row <= rows.length + 1; row += 1) {
    sheet.getCell(`H${row}`).numFmt = 'mm-dd-yy';
    sheet.getCell(`L${row}`).numFmt = 'mm-dd-yy';
    sheet.getRow(row).alignment = { vertical: 'middle' };
  }

  const buffer = Buffer.from(await output.xlsx.writeBuffer());
  const safeCustomer = customer.replace(/[^A-Za-z0-9._ -]+/g, ' ').trim() || 'Cliente';
  const safeDate = loadDate.toISOString().slice(0, 10);
  return {
    buffer,
    fileName: `POSCO Wholesale - ${safeCustomer} - ${safeDate}.xlsx`,
    rows: rows.length,
    boomPending
  };
}

module.exports = { createPoscoWorkbook };
