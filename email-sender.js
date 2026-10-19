const DEFAULT_RECIPIENTS = 'Jesús|jesus.sandoval@cfbc.co';
const GRAPH_URL = 'https://graph.microsoft.com/v1.0';

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

function escapeHtml(value) {
  return clean(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function parseRecipients(value = process.env.WHOLESALE_PDF_RECIPIENTS || DEFAULT_RECIPIENTS) {
  const recipients = [];
  const seen = new Set();
  for (const raw of String(value).split(/[;\r\n]+/)) {
    const entry = raw.trim();
    if (!entry) continue;
    const separator = entry.indexOf('|');
    const name = separator >= 0 ? entry.slice(0, separator).trim() : '';
    const address = (separator >= 0 ? entry.slice(separator + 1) : entry).trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address)) {
      throw new Error(`Destinatario inválido: ${address || '(vacío)'}. Usa Nombre|correo@empresa.com.`);
    }
    const key = address.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    recipients.push({ name: name || address.split('@')[0], address });
  }
  if (!recipients.length) throw new Error('WHOLESALE_PDF_RECIPIENTS debe tener al menos un destinatario.');
  return recipients;
}

function requiredEnv(...names) {
  for (const name of names) {
    const value = clean(process.env[name]);
    if (value) return value;
  }
  throw new Error(`Falta configurar ${names.join(' o ')} en Render.`);
}

async function graphToken() {
  const tenant = requiredEnv('SHAREPOINT_TENANT_ID', 'GRAPH_TENANT_ID');
  const clientId = requiredEnv('SHAREPOINT_CLIENT_ID', 'GRAPH_CLIENT_ID');
  const clientSecret = requiredEnv('SHAREPOINT_CLIENT_SECRET', 'GRAPH_CLIENT_SECRET');
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const response = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  if (!response.ok) {
    throw new Error(`Microsoft Graph no autorizó el envío (${response.status}).`);
  }
  const payload = await response.json();
  if (!payload.access_token) throw new Error('Microsoft Graph no devolvió un token de acceso.');
  return payload.access_token;
}

function attachment(name, contentType, buffer) {
  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name,
    contentType,
    contentBytes: buffer.toString('base64')
  };
}

function metadata(workbook) {
  const sheet = workbook?.sheets?.['PO Nuevo'];
  return {
    customer: clean(cellValue(sheet, 'F11')) || 'Cliente',
    poNumber: clean(cellValue(sheet, 'A9')) || 'Sin PO',
    loadDate: clean(cellValue(sheet, 'B19')) || 'Sin fecha'
  };
}

function passingMetadata(workbook) {
  const sheet = workbook?.sheets?.['PASSING SHEET'];
  return {
    customer: clean(cellValue(sheet, 'B5')) || 'Cliente',
    loadDate: clean(cellValue(sheet, 'B6')) || 'Sin fecha'
  };
}

function pendingHtml(boomPending) {
  if (!boomPending.length) return '';
  const items = boomPending.map(item => (
    `<li><strong>${escapeHtml(item.product)}</strong><br>Nombre recomendado: ${escapeHtml(item.recommendation)}</li>`
  )).join('');
  return '<p><strong style="color:#b91c1c">PENDIENTE DE BOM</strong></p>'
    + '<p>Se requiere revisar o crear el BOM para:</p><ul>' + items + '</ul>';
}

function buildWholesaleMailPayload({ workbook, recipient, pdfBuffer, pdfFileName, poscoBuffer, poscoFileName, boomPending }) {
  const order = metadata(workbook);
  const pending = Array.isArray(boomPending) ? boomPending : [];
  const body = `<p>Hola ${escapeHtml(recipient.name)},</p>`
    + `<p>Adjunto encontrarás la Orden Wholesale y el formato POSCO. PO: <strong>${escapeHtml(order.poNumber)}</strong>. Load Date: <strong>${escapeHtml(order.loadDate)}</strong>.</p>`
    + pendingHtml(pending)
    + '<p>Saludos.</p>';
  return {
    message: {
      subject: `Orden Wholesale - ${order.customer}`,
      body: { contentType: 'HTML', content: body },
      toRecipients: [{ emailAddress: { address: recipient.address } }],
      ccRecipients: [],
      bccRecipients: [],
      attachments: [
        attachment(pdfFileName, 'application/pdf', pdfBuffer),
        attachment(poscoFileName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', poscoBuffer)
      ]
    },
    saveToSentItems: true
  };
}

function buildPassingSheetMailPayload({ workbook, recipient, pdfBuffer, pdfFileName }) {
  const passing = passingMetadata(workbook);
  const body = `<p>Hola ${escapeHtml(recipient.name)},</p>`
    + `<p>Adjunto encontrarás el Passing Sheet generado. Load Date: <strong>${escapeHtml(passing.loadDate)}</strong>.</p>`
    + '<p>Saludos.</p>';
  return {
    message: {
      subject: `Passing Sheet Wholesale - ${passing.customer}`,
      body: { contentType: 'HTML', content: body },
      toRecipients: [{ emailAddress: { address: recipient.address } }],
      ccRecipients: [],
      bccRecipients: [],
      attachments: [attachment(pdfFileName, 'application/pdf', pdfBuffer)]
    },
    saveToSentItems: true
  };
}

async function sendWholesaleEmail({ workbook, pdfBuffer, pdfFileName, poscoBuffer, poscoFileName, boomPending }) {
  const sender = requiredEnv('MAIL_SENDER');
  const recipients = parseRecipients();
  const token = await graphToken();
  const results = [];
  for (const recipient of recipients) {
    const payload = buildWholesaleMailPayload({
      workbook,
      recipient,
      pdfBuffer,
      pdfFileName,
      poscoBuffer,
      poscoFileName,
      boomPending
    });
    const response = await fetch(`${GRAPH_URL}/users/${encodeURIComponent(sender)}/sendMail`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!response.ok) {
      const detail = (await response.text()).slice(0, 300);
      throw new Error(`Microsoft Graph rechazó el correo a ${recipient.address} (${response.status}). ${detail}`);
    }
    results.push(recipient.address);
  }
  return { sentTo: results, pendingBom: boomPending.length };
}

async function sendPassingSheetEmail({ workbook, pdfBuffer, pdfFileName }) {
  const sender = requiredEnv('MAIL_SENDER');
  const recipients = parseRecipients();
  const token = await graphToken();
  const results = [];
  for (const recipient of recipients) {
    const payload = buildPassingSheetMailPayload({ workbook, recipient, pdfBuffer, pdfFileName });
    const response = await fetch(`${GRAPH_URL}/users/${encodeURIComponent(sender)}/sendMail`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!response.ok) {
      const detail = (await response.text()).slice(0, 300);
      throw new Error(`Microsoft Graph rechazó el Passing Sheet a ${recipient.address} (${response.status}). ${detail}`);
    }
    results.push(recipient.address);
  }
  return { sentTo: results };
}

module.exports = {
  buildPassingSheetMailPayload,
  buildWholesaleMailPayload,
  parseRecipients,
  sendPassingSheetEmail,
  sendWholesaleEmail
};
