const PDFDocument = require('pdfkit');

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

function parseDate(value, label) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    if (!Number.isNaN(date.getTime())) return date;
  }
  const text = clean(value);
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) {
    return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  }
  throw new Error(`${label} está vacío o no es una fecha válida en PO Nuevo.`);
}

function formatDate(value) {
  const date = value instanceof Date ? value : parseDate(value, 'Fecha');
  return date.toLocaleDateString('en-US', { timeZone: 'UTC' });
}

function findTotalRow(orderSheet) {
  if (Number(orderSheet.poTotalRow) >= 23) return Number(orderSheet.poTotalRow);
  const maxRow = Number(orderSheet.maxRow || 100);
  for (let row = 23; row <= maxRow; row += 1) {
    if (normalize(cellValue(orderSheet, `G${row}`)) === 'total') return row;
  }
  return 42;
}

function numberValue(value) {
  const number = Number(String(value ?? '').replace(/[$,]/g, '').trim());
  return Number.isFinite(number) ? number : 0;
}

function money(value) {
  return `$ ${numberValue(value).toFixed(2)}`;
}

function drawText(doc, value, x, y, width, options = {}) {
  doc.text(clean(value) || '-', x + 3, y + 4, {
    width: Math.max(1, width - 6),
    height: Math.max(1, options.height - 6),
    lineBreak: false,
    ellipsis: true,
    align: options.align || 'left'
  });
}

function drawTableHeader(doc, columns, y) {
  let x = doc.page.margins.left;
  doc.font('Helvetica-Bold').fontSize(8).fillColor('#1f2937');
  columns.forEach(column => {
    doc.rect(x, y, column.width, 22).fillAndStroke('#f2e600', '#475569');
    doc.fillColor('#1f2937');
    drawText(doc, column.label, x, y, column.width, { height: 22, align: 'center' });
    x += column.width;
  });
  return y + 22;
}

function drawTableRow(doc, columns, values, y, height) {
  let x = doc.page.margins.left;
  doc.font('Helvetica').fontSize(8).fillColor('#111827');
  columns.forEach(column => {
    doc.rect(x, y, column.width, height).fillAndStroke('#ffffff', '#64748b');
    drawText(doc, values[column.key], x, y, column.width, { height, align: column.align });
    x += column.width;
  });
  return y + height;
}

async function createOrderPdf(workbook) {
  const orderSheet = workbook?.sheets?.['PO Nuevo'];
  if (!orderSheet) throw new Error('El libro debe incluir la hoja PO Nuevo.');

  const poNumber = clean(cellValue(orderSheet, 'A9'));
  const customer = clean(cellValue(orderSheet, 'F11'));
  const truckLine = clean(cellValue(orderSheet, 'F15'));
  const loadDate = parseDate(cellValue(orderSheet, 'B19'), 'Load Date');
  if (!poNumber || !customer || !truckLine) {
    throw new Error('PO Nuevo requiere PO, Customer Ship y Truck Line para crear el PDF.');
  }

  const totalRow = findTotalRow(orderSheet);
  const rows = [];
  let totalCases = 0;
  let totalPrice = 0;
  for (let row = 22; row < totalRow; row += 1) {
    const flower = clean(cellValue(orderSheet, `B${row}`));
    if (!flower) continue;
    const cases = numberValue(cellValue(orderSheet, `H${row}`));
    const price = numberValue(cellValue(orderSheet, `I${row}`));
    totalCases += cases;
    totalPrice += cases * price;
    rows.push({
      po: poNumber,
      flower,
      color: cellValue(orderSheet, `C${row}`),
      stems: cellValue(orderSheet, `D${row}`),
      bunches: cellValue(orderSheet, `E${row}`),
      box: cellValue(orderSheet, `F${row}`),
      dimension: cellValue(orderSheet, `G${row}`),
      cases,
      price
    });
  }
  if (!rows.length) throw new Error('PO Nuevo no contiene productos para crear el PDF.');

  const safeCustomer = customer.replace(/[^A-Za-z0-9._ -]+/g, ' ').trim() || 'Cliente';
  const filenameDate = loadDate.toISOString().slice(0, 10);
  const fileName = `Orden Wholesale - ${safeCustomer} - ${filenameDate}.pdf`;
  const doc = new PDFDocument({
    size: 'LETTER',
    layout: 'landscape',
    margin: 28,
    info: { Title: `Orden Wholesale - ${customer}`, Author: 'ORDENES WHOLESALE' }
  });
  const chunks = [];
  doc.on('data', chunk => chunks.push(chunk));
  const completed = new Promise((resolve, reject) => {
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);
  });

  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  doc.font('Helvetica-Bold').fontSize(17).fillColor('#14532d').text('PACIFICA PRODUCE FARMS LLC', left, 28);
  doc.font('Helvetica').fontSize(8).fillColor('#334155').text('PO Box 530990 • San Diego, CA 92154', left, 50);
  doc.font('Helvetica-Bold').fontSize(12).fillColor('#111827').text(`Customer PO: ${poNumber}`, right - 190, 30, { width: 190, align: 'right' });
  doc.moveTo(left, 70).lineTo(right, 70).strokeColor('#14532d').lineWidth(1.5).stroke();

  doc.font('Helvetica-Bold').fontSize(9).fillColor('#111827');
  doc.text(`SHIPPER: `, left, 86, { continued: true }).font('Helvetica').text('Pacifica Produce Farms LLC');
  doc.font('Helvetica-Bold').text(`SHIP TO: `, left + 260, 86, { continued: true }).font('Helvetica').text(customer);
  doc.font('Helvetica-Bold').text(`Farm Ship to: `, left, 101, { continued: true }).font('Helvetica').text('FLORAL TRADE');
  doc.font('Helvetica-Bold').text(`TRUCK LINE: `, left + 260, 101, { continued: true }).font('Helvetica').text(truckLine);
  doc.font('Helvetica-Bold').text(`Load Date: `, left, 116, { continued: true }).font('Helvetica').text(formatDate(loadDate));
  doc.font('Helvetica-Bold').text(`Tender Date: `, left + 260, 116, { continued: true }).font('Helvetica').text(formatDate(cellValue(orderSheet, 'F19') || loadDate));

  const columns = [
    { key: 'po', label: 'PO', width: 48, align: 'center' },
    { key: 'flower', label: 'FLOWER', width: 170 },
    { key: 'color', label: 'COLOR', width: 72, align: 'center' },
    { key: 'stems', label: 'STEM/BCH', width: 66, align: 'center' },
    { key: 'bunches', label: 'BU/BOX', width: 58, align: 'center' },
    { key: 'box', label: 'BOX TYPE', width: 70, align: 'center' },
    { key: 'dimension', label: 'BOX DIMENSION', width: 92, align: 'center' },
    { key: 'cases', label: 'CASES', width: 52, align: 'center' },
    { key: 'price', label: 'PRICE $', width: 62, align: 'right' }
  ];
  let y = drawTableHeader(doc, columns, 140);
  rows.forEach(row => {
    if (y + 23 + 24 > doc.page.height - 48) {
      doc.addPage();
      y = drawTableHeader(doc, columns, 40);
    }
    const values = { ...row, price: money(row.price) };
    y = drawTableRow(doc, columns, values, y, 23);
  });
  const totalValues = { po: '', flower: '', color: '', stems: '', bunches: '', box: '', dimension: 'TOTAL', cases: totalCases, price: money(totalPrice) };
  if (y + 24 > doc.page.height - 48) {
    doc.addPage();
    y = drawTableHeader(doc, columns, 40);
  }
  drawTableRow(doc, columns, totalValues, y, 24);

  doc.font('Helvetica').fontSize(8).fillColor('#64748b').text(
    `Generado por ORDENES WHOLESALE • ${new Date().toLocaleString('en-US')}`,
    left,
    doc.page.height - 35,
    { width: right - left, align: 'center' }
  );
  doc.end();
  return { buffer: await completed, fileName, rows: rows.length };
}

module.exports = { createOrderPdf };
