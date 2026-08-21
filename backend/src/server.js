import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import multer from 'multer';
import ExcelJS from 'exceljs';
import PDFDocument from 'pdfkit';
import { existsSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
import {
  dbPath,
  saveImport,
  getDashboard,
  getTaxpayers,
  getTaxpayer,
  setPaidStatus,
  getImportHistory,
  getSourceRows,
  addMessage,
  getMessages,
  updateMessageStatusByWamid,
  getWhatsappStatus,
  getWhatsappSummary,
  getWhatsappStatusScoped,
  getWhatsappSummaryScoped,
  getPermissionsList,
  getRoles,
  saveRole,
  deleteRole,
  getUsers,
  getUser,
  loginUser,
  saveUser,
  deleteUser,
  resetUserPassword,
  getSettings,
  saveSettings,
  getSentCountForPackage,
  getEffectiveMessageLimit,
  getVillages,
  saveVillage,
  getCoupons,
  createCoupon,
  redeemCoupon
} from './db.js';
import { clean, extractWorkbook } from './excelExtractors.js';

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const PORT = process.env.PORT || 4200;
const bundledGujaratiFontPath = join(__dirname, 'fonts', 'NotoSansGujarati-Regular.ttf');
const gujaratiFontPath = process.env.GUJARATI_FONT_PATH || bundledGujaratiFontPath;

app.use(cors());
app.use(express.json({ limit: '2mb' }));

function makeBillPdf(taxpayer, res) {
  const doc = new PDFDocument({ size: [226, 680], margin: 14 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="bill-${taxpayer.propertyNo}.pdf"`);
  doc.pipe(res);

  if (existsSync(gujaratiFontPath)) {
    doc.registerFont('Gujarati', gujaratiFontPath);
    doc.font('Gujarati');
  }

  const village = taxpayer.village || process.env.VILLAGE_NAME || 'ગ્રામપંચાયત';
  const amount = taxpayer.pendingTax || taxpayer.currentTax || taxpayer.demandTotal || 0;
  const receiptNo = `${taxpayer.propertyNo}-${new Date().getFullYear()}`;

  doc.fontSize(14).text(village, { align: 'center' });
  doc.fontSize(9).text('મિલકત વેરા બિલ / પહોંચ', { align: 'center' });
  doc.moveDown(0.35);
  doc.fontSize(8).text('----------------------------------------', { align: 'center' });
  doc.moveDown(0.3);

  doc.fontSize(9);
  doc.text(`પહોંચ નં.: ${receiptNo}`);
  doc.text(`તારીખ: ${new Date().toLocaleDateString('en-IN')}`);
  doc.text(`મિલકત નં.: ${taxpayer.propertyNo}`);
  if (taxpayer.houseNo) doc.text(`મકાન નં.: ${taxpayer.houseNo}`);
  doc.text(`માલિકનું નામ: ${taxpayer.holderName || '-'}`);
  if (taxpayer.occupantName) doc.text(`કબજેદારનું નામ: ${taxpayer.occupantName}`);
  doc.text(`વિસ્તાર: ${taxpayer.area || '-'}`);
  doc.text(`મોબાઇલ નં.: ${taxpayer.mobile || '-'}`);
  doc.text(`મિલકતનો પ્રકાર: ${taxpayer.category || '-'}`);
  if (taxpayer.description) doc.text(`વર્ણન: ${taxpayer.description}`);
  doc.moveDown(0.35);
  doc.fontSize(8).text('----------------------------------------', { align: 'center' });
  doc.moveDown(0.2);

  doc.fontSize(9);
  doc.text('વેરાની વિગત', { underline: true });
  doc.text(`માંગણા રકમ: રૂ. ${taxpayer.demandTotal || 0}`);
  doc.text(`પાછલી / બાકી રકમ: રૂ. ${taxpayer.pendingTax || 0}`);
  doc.text(`ચાલુ રકમ: રૂ. ${taxpayer.currentTax || amount}`);
  doc.text(`કુલ ચુકવવાની રકમ: રૂ. ${amount}`);
  doc.text(`ચુકવણી સ્થિતિ: ${taxpayer.paid ? 'ચૂકવેલ' : 'બાકી'}`);
  doc.moveDown(0.35);
  doc.fontSize(8).text('----------------------------------------', { align: 'center' });
  doc.moveDown(0.3);

  doc.fontSize(8);
  doc.text('નોંધ: કૃપા કરીને આ બિલ ઓફિસ રેકોર્ડ માટે સાચવી રાખશો.');
  doc.text('આ બિલ WhatsApp Billing Panel દ્વારા જનરેટ કરવામાં આવ્યું છે.');
  doc.moveDown(1.4);
  doc.text('સહી / સિક્કો: __________________', { align: 'right' });
  doc.end();
}

async function makeTaxpayerReportWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'WhatsApp Billing Panel';
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet('Excel Data Report');

  worksheet.columns = [
    { header: 'Property No', key: 'propertyNo', width: 14 },
    { header: 'Old Property No', key: 'oldPropertyNo', width: 16 },
    { header: 'House No', key: 'houseNo', width: 12 },
    { header: 'Owner Name', key: 'holderName', width: 34 },
    { header: 'Occupant Name', key: 'occupantName', width: 28 },
    { header: 'Area', key: 'area', width: 28 },
    { header: 'Mobile No', key: 'mobile', width: 16 },
    { header: 'Category', key: 'category', width: 18 },
    { header: 'Description', key: 'description', width: 42 },
    { header: 'Demand Total', key: 'demandTotal', width: 16 },
    { header: 'Pending Amount', key: 'pendingTax', width: 16 },
    { header: 'Current Amount', key: 'currentTax', width: 16 },
    { header: 'Status', key: 'status', width: 14 }
  ];

  rows.forEach((item) => {
    worksheet.addRow({
      ...item,
      status: item.paid ? 'Paid' : 'Unpaid'
    });
  });

  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF1EB' } };
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
  worksheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: worksheet.columns.length } };

  return workbook;
}

async function sendWhatsAppMessage({ taxpayer, message }) {
  const settings = getSettings();
  const token = settings.whatsapp_access_token || process.env.WHATSAPP_ACCESS_TOKEN;
  const phoneNumberId = settings.whatsapp_phone_number_id || process.env.WHATSAPP_PHONE_NUMBER_ID;
  if (!token || !phoneNumberId) {
    return { status: 'pending', detail: 'WhatsApp credentials are not configured.' };
  }

  const response = await fetch(`https://graph.facebook.com/v20.0/${phoneNumberId}/messages`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      messaging_product: 'whatsapp',
      to: taxpayer.mobile,
      type: 'text',
      text: { body: message }
    })
  });

  if (!response.ok) {
    return { status: 'failed', detail: await response.text() };
  }
  return { status: 'delivered', detail: await response.json() };
}

function normalizeIndianMobile(mobile) {
  const digits = String(mobile || '').replace(/\D/g, '');
  if (digits.length === 10) return `91${digits}`;
  return digits;
}

async function sendBillWhatsApp(taxpayer) {
  const settings = getSettings();
  const token = settings.whatsapp_access_token || process.env.WHATSAPP_ACCESS_TOKEN;
  const phoneNumberId = settings.whatsapp_phone_number_id || process.env.WHATSAPP_PHONE_NUMBER_ID;
  const templateName = settings.whatsapp_template_name || process.env.WHATSAPP_TEMPLATE_NAME || 'bill_pdf_notification';
  const languageCode = settings.whatsapp_language_code || process.env.WHATSAPP_LANGUAGE_CODE || 'gu';
  const backendUrl = process.env.BACKEND_PUBLIC_URL || '';

  if (!token || !phoneNumberId) {
    return { status: 'pending', detail: 'WhatsApp credentials are not configured.' };
  }
  if (!backendUrl) {
    return { status: 'failed', detail: 'BACKEND_PUBLIC_URL is not configured, cannot build a public bill link.' };
  }

  const to = normalizeIndianMobile(taxpayer.mobile);
  if (!to) {
    return { status: 'failed', detail: 'No valid mobile number for this taxpayer.' };
  }

  const village = taxpayer.village || process.env.VILLAGE_NAME || 'ગ્રામપંચાયત';
  const amount = taxpayer.pendingTax || taxpayer.currentTax || taxpayer.demandTotal || 0;
  const link = `${backendUrl}/api/bills/${encodeURIComponent(taxpayer.propertyNo)}.pdf`;

  const response = await fetch(`https://graph.facebook.com/v21.0/${phoneNumberId}/messages`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      messaging_product: 'whatsapp',
      to,
      type: 'template',
      template: {
        name: templateName,
        language: { code: languageCode },
        components: [
          {
            type: 'header',
            parameters: [{ type: 'document', document: { link, filename: `bill-${taxpayer.propertyNo}.pdf` } }]
          },
          {
            type: 'body',
            parameters: [
              { type: 'text', text: taxpayer.holderName || 'ગ્રાહક' },
              { type: 'text', text: String(amount) },
              { type: 'text', text: village }
            ]
          }
        ]
      }
    })
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    return { status: 'failed', detail: data };
  }
  return { status: 'sent', detail: data, wamid: data?.messages?.[0]?.id || '' };
}

function currentUser(req) {
  const id = req.header('x-user-id') || req.query.userId;
  return id ? getUser(id) : null;
}

function requirePermission(permission) {
  return (req, res, next) => {
    const user = currentUser(req);
    if (!user) return res.status(401).json({ error: 'Login required.' });
    if (!user.permissions.includes(permission)) return res.status(403).json({ error: 'Permission denied.' });
    req.user = user;
    next();
  };
}

app.get('/api/health', (req, res) => {
  res.json({ ok: true, village: process.env.VILLAGE_NAME || 'KHIJADIYA ગ્રામપંચાયત', dbPath });
});

app.post('/api/auth/login', (req, res) => {
  const user = loginUser(clean(req.body.username), clean(req.body.password));
  if (!user) return res.status(401).json({ error: 'Invalid login or expired access.' });
  res.json(user);
});

app.get('/api/auth/me', (req, res) => {
  const user = currentUser(req);
  if (!user) return res.status(401).json({ error: 'Login required.' });
  res.json(user);
});

app.get('/api/dashboard', requirePermission('dashboard.view'), (req, res) => {
  res.json(getDashboard(req.query, req.user));
});

app.get('/api/taxpayers', requirePermission('billing.view'), (req, res) => {
  res.json(getTaxpayers(req.query, req.user));
});

app.get('/api/imports', requirePermission('excel.view'), (req, res) => {
  res.json(getImportHistory(req.query, req.user));
});

app.get('/api/source-rows', requirePermission('excel.view'), (req, res) => {
  res.json(getSourceRows(req.query.kind, req.query, req.user));
});

app.get('/api/reports/taxpayers.xlsx', requirePermission('reports.view'), async (req, res) => {
  const rows = getTaxpayers(req.query, req.user);
  const workbook = await makeTaxpayerReportWorkbook(rows);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="taxpayer-excel-report.xlsx"');
  await workbook.xlsx.write(res);
  res.end();
});

app.post('/api/import/:kind', requirePermission('excel.import'), upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Excel file is required.' });
  try {
    const extracted = await extractWorkbook(req.file.buffer, req.params.kind);
    const result = saveImport({
      kind: req.params.kind,
      fileName: req.file.originalname,
      extracted,
      village: req.body.villageName || process.env.VILLAGE_NAME || 'KHIJADIYA ગ્રામપંચાયત',
      userId: req.user.id,
      villageId: Number(req.body.villageId || 1)
    });
    res.json({ ...result, sourceRows: extracted.rows.length, dbPath });
  } catch (error) {
    res.status(422).json({ error: error.message });
  }
});

app.patch('/api/taxpayers/:propertyNo/status', requirePermission('billing.update'), (req, res) => {
  const taxpayer = setPaidStatus(req.params.propertyNo, Boolean(req.body.paid), req.body, req.user);
  if (!taxpayer) return res.status(404).json({ error: 'Taxpayer not found.' });
  res.json(taxpayer);
});

app.get('/api/bills/:propertyNo.pdf', (req, res) => {
  const taxpayer = getTaxpayer(req.params.propertyNo, req.query, currentUser(req));
  if (!taxpayer) return res.status(404).json({ error: 'Taxpayer not found.' });
  makeBillPdf(taxpayer, res);
});

app.post('/api/bills/:propertyNo/send', requirePermission('broadcast.send'), async (req, res) => {
  const taxpayer = getTaxpayer(req.params.propertyNo, req.query, req.user);
  if (!taxpayer) return res.status(404).json({ error: 'Taxpayer not found.' });
  if (!taxpayer.mobile) return res.status(400).json({ error: 'This taxpayer has no mobile number on file.' });

  const settings = getSettings();
  const isSuperAdmin = req.user.permissions.includes('superadmin.view_all');
  const packageLimit = isSuperAdmin ? 0 : (getEffectiveMessageLimit(req.user) || Number(req.user.messageLimit || settings.daily_message_limit || 0));
  const sentInPackage = getSentCountForPackage(req.user.id);
  if (packageLimit > 0 && sentInPackage + 1 > packageLimit) {
    return res.status(429).json({ error: `WhatsApp package limit exceeded. Limit ${packageLimit}, already used ${sentInPackage}. Use recharge coupon for more messages.` });
  }

  const result = await sendBillWhatsApp(taxpayer);
  const row = {
    userId: req.user.id,
    villageId: taxpayer.villageId,
    propertyNo: taxpayer.propertyNo,
    mobile: taxpayer.mobile,
    message: `Bill PDF (${settings.whatsapp_template_name || 'bill_pdf_notification'})`,
    ...result,
    sentAt: new Date().toISOString()
  };
  addMessage(row);
  res.json(row);
});

app.post('/api/messages/send', requirePermission('broadcast.send'), async (req, res) => {
  const recipients = getTaxpayers(req.body.filters || {}, req.user).filter((item) => item.mobile);
  const message = clean(req.body.message) || 'તમારું પંચાયત બિલ તૈયાર છે. કૃપા કરીને ચુકવણી કરો.';
  const settings = getSettings();
  const isSuperAdmin = req.user.permissions.includes('superadmin.view_all');
  const packageLimit = isSuperAdmin ? 0 : (getEffectiveMessageLimit(req.user) || Number(req.user.messageLimit || settings.daily_message_limit || 0));
  const sentInPackage = getSentCountForPackage(req.user.id);
  if (packageLimit > 0 && sentInPackage + recipients.length > packageLimit) {
    return res.status(429).json({ error: `WhatsApp package limit exceeded. Limit ${packageLimit}, already used ${sentInPackage}. Use recharge coupon for more messages.` });
  }
  const results = [];

  for (const taxpayer of recipients) {
    const result = await sendWhatsAppMessage({ taxpayer, message });
    const row = { userId: req.user.id, villageId: taxpayer.villageId, propertyNo: taxpayer.propertyNo, mobile: taxpayer.mobile, message, ...result, sentAt: new Date().toISOString() };
    addMessage(row);
    results.push(row);
  }

  res.json({ sent: results.length, results });
});

app.get('/api/messages/report', requirePermission('reports.view'), (req, res) => {
  res.json(getMessages(req.query, req.user));
});

app.get('/api/whatsapp/status', requirePermission('reports.view'), (req, res) => {
  res.json({
    summary: getWhatsappSummaryScoped(req.query, req.user),
    rows: getWhatsappStatusScoped(req.query, req.user)
  });
});

app.get('/api/villages', requirePermission('excel.view'), (req, res) => {
  res.json(getVillages(req.query, req.user));
});

app.post('/api/villages', requirePermission('excel.import'), (req, res) => {
  res.json(saveVillage(req.body, req.user));
});

app.get('/api/coupons', requirePermission('reports.view'), (req, res) => {
  res.json(getCoupons(req.query, req.user));
});

app.post('/api/coupons/redeem', requirePermission('broadcast.send'), (req, res) => {
  try {
    res.json(redeemCoupon(req.body.code, req.user));
  } catch (error) {
    res.status(422).json({ error: error.message });
  }
});

app.get('/api/admin/bootstrap', requirePermission('admin.manage'), (req, res) => {
  res.json({
    permissions: getPermissionsList(),
    roles: getRoles(),
    users: getUsers(),
    settings: getSettings(),
    villages: getVillages({}, req.user),
    coupons: getCoupons({}, req.user)
  });
});

app.post('/api/admin/roles', requirePermission('admin.manage'), (req, res) => {
  res.json(saveRole(req.body));
});

app.delete('/api/admin/roles/:id', requirePermission('admin.manage'), (req, res) => {
  try {
    res.json(deleteRole(Number(req.params.id)));
  } catch (error) {
    res.status(422).json({ error: error.message });
  }
});

app.post('/api/admin/users', requirePermission('admin.manage'), (req, res) => {
  res.json(saveUser(req.body));
});

app.delete('/api/admin/users/:id', requirePermission('admin.manage'), (req, res) => {
  try {
    res.json(deleteUser(Number(req.params.id)));
  } catch (error) {
    res.status(422).json({ error: error.message });
  }
});

app.post('/api/admin/users/:id/reset-password', requirePermission('admin.manage'), (req, res) => {
  try {
    res.json(resetUserPassword(Number(req.params.id), req.body.password));
  } catch (error) {
    res.status(422).json({ error: error.message });
  }
});

app.post('/api/admin/settings', requirePermission('admin.manage'), (req, res) => {
  res.json(saveSettings(req.body));
});

app.post('/api/admin/coupons', requirePermission('admin.manage'), (req, res) => {
  res.json(createCoupon(req.body));
});

// Meta calls this to verify the webhook URL when you configure it in the app dashboard.
app.get('/api/webhook/whatsapp', (req, res) => {
  const mode = req.query['hub.mode'];
  const token = req.query['hub.verify_token'];
  const challenge = req.query['hub.challenge'];
  const expected = process.env.WHATSAPP_WEBHOOK_VERIFY_TOKEN;

  if (mode === 'subscribe' && expected && token === expected) {
    res.status(200).send(challenge);
  } else {
    res.status(403).end();
  }
});

// Meta posts delivery/read/failed status updates here for messages we sent.
app.post('/api/webhook/whatsapp', (req, res) => {
  res.status(200).end();

  const statuses = req.body?.entry?.[0]?.changes?.[0]?.value?.statuses || [];
  for (const status of statuses) {
    if (status.id) {
      updateMessageStatusByWamid(status.id, status.status, status.errors || status);
    }
  }
});

app.listen(PORT, () => {
  console.log(`WhatsApp billing backend running on http://localhost:${PORT}`);
  console.log(`SQLite database: ${dbPath}`);
});
