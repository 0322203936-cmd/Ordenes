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

function parseDate(value) {
  const text = clean(value);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
}

function formatDate(value) {
  const date = parseDate(value);
  return date ? date.toLocaleDateString('en-US', { timeZone: 'UTC' }) : clean(value) || '-';
}

function safePart(value, fallback) {
  return clean(value).replace(/[^A-Za-z0-9._ -]+/g, ' ').trim() || fallback;
}

function drawTableRow(doc, label, value, x, y, labelWidth, valueWidth, height = 25, boldValue = false) {
  const totalWidth = labelWidth + valueWidth;
  doc.lineWidth(0.7).strokeColor('#334155').rect(x, y, totalWidth, height).stroke();
  doc.moveTo(x + labelWidth, y).lineTo(x + labelWidth, y + height).stroke();
  doc.font('Helvetica-Bold').fontSize(9.5).fillColor('#111827').text(`${clean(label)}:`, x + 7, y + 7, {
    width: labelWidth - 14,
    height: height - 10,
    ellipsis: true
  });
  doc.font(boldValue ? 'Helvetica-Bold' : 'Helvetica').fontSize(9.5).fillColor('#111827').text(clean(value) || '-', x + labelWidth + 7, y + 7, {
    width: valueWidth - 14,
    height: height - 10,
    ellipsis: true
  });
}

async function createPassingSheetPdf(workbook) {
  const sheet = workbook?.sheets?.['PASSING SHEET'];
  if (!sheet) throw new Error('El libro debe incluir la hoja PASSING SHEET.');

  const customer = clean(cellValue(sheet, 'B5')) || 'Cliente';
  const loadDate = cellValue(sheet, 'B6');
  const arrivalLabel = clean(cellValue(sheet, 'A6')) || 'Load Date (Arrival to Floral Trade):';
  const farmShipTo = clean(cellValue(workbook?.sheets?.['PO Nuevo'], 'B18')) || 'FLORAL TRADE';
  const fileName = `Passing Sheet Wholesale - ${safePart(customer, 'Cliente')} - ${formatDate(loadDate).replace(/\//g, '-')}.pdf`;
  const doc = new PDFDocument({
    size: 'LETTER',
    layout: 'portrait',
    margin: 48,
    info: { Title: `Passing Sheet Wholesale - ${customer}`, Author: 'ORDENES WHOLESALE' }
  });
  const chunks = [];
  doc.on('data', chunk => chunks.push(chunk));
  const completed = new Promise((resolve, reject) => {
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);
  });

  const pageLeft = doc.page.margins.left;
  const pageRight = doc.page.width - doc.page.margins.right;
  const sheetLeft = pageLeft + 16;
  const sheetRight = pageRight - 16;
  const sheetWidth = sheetRight - sheetLeft;
  doc.lineWidth(0.7).strokeColor('#cbd5e1').rect(pageLeft, 36, pageRight - pageLeft, 650).stroke();
  doc.font('Helvetica-Bold').fontSize(18).fillColor('#14532d').text('PACIFICA PRODUCE FARMS LLC', sheetLeft, 54);
  doc.font('Helvetica').fontSize(8.5).fillColor('#334155').text('PO Box 530990 • San Diego, CA 92154', sheetLeft, 77);
  doc.moveTo(sheetLeft, 94).lineTo(sheetRight, 94).strokeColor('#14532d').lineWidth(1.2).stroke();
  doc.font('Helvetica-Bold').fontSize(16).fillColor('#111827').text(`PASSING SHEET – ${farmShipTo}`, sheetLeft, 112);
  doc.font('Helvetica').fontSize(10).fillColor('#334155').text(
    clean(cellValue(sheet, 'A2')) || 'Please receive the following shipment and release it to the designated trucking line as instructed below.',
    sheetLeft,
    136,
    { width: sheetWidth }
  );

  const tableTop = 164;
  const labelWidth = 270;
  const valueWidth = sheetWidth - labelWidth;
  const rows = [
    ['Farm / Shipper', cellValue(sheet, 'B4'), false],
    ['Customer Name', customer, true],
    [arrivalLabel.replace(/:$/, ''), formatDate(loadDate), false],
    ['Release Date (Deliver to Trucking Line)', formatDate(cellValue(sheet, 'B7')), false],
    ['Trucking Line', cellValue(sheet, 'B8'), false],
    ['Account # (if applicable)', cellValue(sheet, 'B9'), false],
    ['Number of Boxes', clean(cellValue(sheet, 'B10')) || '0', true]
  ];
  rows.forEach(([label, value, boldValue], index) => {
    drawTableRow(doc, label, value, sheetLeft, tableTop + index * 25, labelWidth, valueWidth, 25, boldValue);
  });

  doc.font('Helvetica-Bold').fontSize(11).fillColor('#111827').text('Special Instructions:', sheetLeft, 354);
  [13, 14, 15].forEach((row, index) => {
    const line = clean(cellValue(sheet, `B${row}`)) || ' ';
    const y = 380 + index * 27;
    doc.moveTo(sheetLeft, y + 15).lineTo(sheetRight, y + 15).strokeColor('#334155').lineWidth(0.6).stroke();
    if (line !== ' ') doc.font('Helvetica').fontSize(10).fillColor('#334155').text(line, sheetLeft + 4, y, { width: sheetWidth - 8 });
  });

  doc.font('Helvetica').fontSize(10).fillColor('#111827').text('Authorized by Pacifica Farms: ________________________________', sheetLeft, 486);
  doc.text('Contact Phone: ______________________   Email: ______________________', sheetLeft, 512);
  doc.end();
  return { buffer: await completed, fileName };
}

module.exports = { createPassingSheetPdf };
