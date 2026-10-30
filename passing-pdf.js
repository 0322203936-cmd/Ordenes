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

function addField(doc, label, value, x, y, width) {
  doc.font('Helvetica-Bold').fontSize(10).fillColor('#111827').text(`${label}:`, x, y, { continued: true });
  doc.font('Helvetica').text(` ${clean(value) || '-'}`, { width });
}

async function createPassingSheetPdf(workbook) {
  const sheet = workbook?.sheets?.['PASSING SHEET'];
  if (!sheet) throw new Error('El libro debe incluir la hoja PASSING SHEET.');

  const customer = clean(cellValue(sheet, 'B5')) || 'Cliente';
  const loadDate = cellValue(sheet, 'B6');
  const arrivalLabel = clean(cellValue(sheet, 'A6')) || 'Load Date (Arrival to Floral Trade):';
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

  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  doc.font('Helvetica-Bold').fontSize(18).fillColor('#14532d').text('PACIFICA PRODUCE FARMS LLC', left, 52);
  doc.font('Helvetica').fontSize(9).fillColor('#334155').text('PO Box 530990 • San Diego, CA 92154', left, 78);
  doc.moveTo(left, 96).lineTo(right, 96).strokeColor('#14532d').lineWidth(1.5).stroke();
  doc.font('Helvetica-Bold').fontSize(17).fillColor('#111827').text('PASSING SHEET – FLORAL TRADE', left, 120);
  doc.font('Helvetica').fontSize(10).fillColor('#334155').text(
    clean(cellValue(sheet, 'A2')) || 'Please receive the following shipment and release it to the designated trucking line as instructed below.',
    left,
    150,
    { width: right - left }
  );

  const boxTop = 195;
  doc.roundedRect(left, boxTop, right - left, 210, 4).fillAndStroke('#f8fafc', '#cbd5e1');
  addField(doc, 'Farm / Shipper', cellValue(sheet, 'B4'), left + 18, boxTop + 22, 300);
  addField(doc, 'Customer Name', customer, left + 18, boxTop + 55, 300);
  addField(doc, arrivalLabel.replace(/:$/, ''), formatDate(loadDate), left + 18, boxTop + 88, 300);
  addField(doc, 'Release Date (Deliver to Trucking Line)', formatDate(cellValue(sheet, 'B7')), left + 18, boxTop + 121, 300);
  addField(doc, 'Trucking Line', cellValue(sheet, 'B8'), left + 18, boxTop + 154, 300);
  addField(doc, 'Account # (if applicable)', cellValue(sheet, 'B9'), left + 18, boxTop + 187, 300);
  doc.font('Helvetica-Bold').fontSize(12).fillColor('#111827').text('Number of Boxes:', left + 340, boxTop + 88);
  doc.font('Helvetica-Bold').fontSize(22).fillColor('#14532d').text(clean(cellValue(sheet, 'B10')) || '0', left + 340, boxTop + 116);

  doc.font('Helvetica-Bold').fontSize(11).fillColor('#111827').text('Special Instructions', left, 445);
  [13, 14, 15].forEach((row, index) => {
    const line = clean(cellValue(sheet, `B${row}`)) || ' ';
    const y = 470 + index * 27;
    doc.moveTo(left, y + 15).lineTo(right, y + 15).strokeColor('#64748b').lineWidth(0.6).stroke();
    if (line !== ' ') doc.font('Helvetica').fontSize(10).fillColor('#334155').text(line, left + 4, y, { width: right - left - 8 });
  });

  doc.font('Helvetica').fontSize(10).fillColor('#111827').text('Authorized by Pacifica Farms: ________________________________', left, 585);
  doc.text('Contact Phone: ______________________   Email: ______________________', left, 612);
  doc.font('Helvetica').fontSize(8).fillColor('#64748b').text(
    `Generado por ORDENES WHOLESALE • ${new Date().toLocaleString('en-US')}`,
    left,
    doc.page.height - 35,
    { width: right - left, align: 'center' }
  );
  doc.end();
  return { buffer: await completed, fileName };
}

module.exports = { createPassingSheetPdf };
