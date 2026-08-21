import { mkdirSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import { DatabaseSync } from 'node:sqlite';
import { money } from './excelExtractors.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
export const dbPath = process.env.SQLITE_DB_PATH || join(__dirname, '..', 'data', 'billing.sqlite');

mkdirSync(dirname(dbPath), { recursive: true });

export const db = new DatabaseSync(dbPath);
db.exec('PRAGMA journal_mode = WAL');
db.exec('PRAGMA foreign_keys = ON');

db.exec(`
CREATE TABLE IF NOT EXISTS taxpayers (
  property_no TEXT PRIMARY KEY,
  old_property_no TEXT DEFAULT '',
  house_no TEXT DEFAULT '',
  holder_name TEXT DEFAULT '',
  occupant_name TEXT DEFAULT '',
  area TEXT DEFAULT '',
  category TEXT DEFAULT '',
  mobile TEXT DEFAULT '',
  description TEXT DEFAULT '',
  demand_total REAL DEFAULT 0,
  pending_tax REAL DEFAULT 0,
  current_tax REAL DEFAULT 0,
  paid INTEGER DEFAULT 0,
  village TEXT DEFAULT '',
  updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS excel_imports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kind TEXT NOT NULL,
  file_name TEXT NOT NULL,
  row_count INTEGER NOT NULL,
  imported_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS excel_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  import_id INTEGER NOT NULL REFERENCES excel_imports(id) ON DELETE CASCADE,
  kind TEXT NOT NULL,
  property_no TEXT DEFAULT '',
  receipt_no TEXT DEFAULT '',
  receipt_date TEXT DEFAULT '',
  old_property_no TEXT DEFAULT '',
  house_no TEXT DEFAULT '',
  holder_name TEXT DEFAULT '',
  occupant_name TEXT DEFAULT '',
  area TEXT DEFAULT '',
  mobile TEXT DEFAULT '',
  category TEXT DEFAULT '',
  description TEXT DEFAULT '',
  previous_total REAL DEFAULT 0,
  current_total REAL DEFAULT 0,
  grand_total REAL DEFAULT 0,
  taxes_json TEXT DEFAULT '{}'
);

CREATE TABLE IF NOT EXISTS messages (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  property_no TEXT DEFAULT '',
  mobile TEXT DEFAULT '',
  message TEXT DEFAULT '',
  status TEXT DEFAULT '',
  detail TEXT DEFAULT '',
  sent_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS roles (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  permissions_json TEXT NOT NULL,
  message_limit INTEGER DEFAULT 0,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  username TEXT NOT NULL UNIQUE,
  password TEXT NOT NULL,
  role_id INTEGER NOT NULL REFERENCES roles(id),
  active INTEGER DEFAULT 1,
  access_expires_at TEXT DEFAULT '',
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS app_settings (
  key TEXT PRIMARY KEY,
  value TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS villages (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id INTEGER NOT NULL DEFAULT 1,
  name TEXT NOT NULL,
  taluka TEXT DEFAULT '',
  district TEXT DEFAULT '',
  created_at TEXT NOT NULL,
  UNIQUE(user_id, name)
);

CREATE TABLE IF NOT EXISTS coupons (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  code TEXT NOT NULL UNIQUE,
  message_credit INTEGER NOT NULL DEFAULT 0,
  assigned_user_id INTEGER DEFAULT 0,
  redeemed_user_id INTEGER DEFAULT 0,
  status TEXT NOT NULL DEFAULT 'active',
  expires_at TEXT DEFAULT '',
  created_at TEXT NOT NULL,
  redeemed_at TEXT DEFAULT ''
);
`);

function migrateTaxpayersForSaas() {
  const columns = db.prepare('PRAGMA table_info(taxpayers)').all();
  const propertyColumn = columns.find((item) => item.name === 'property_no');
  if (!propertyColumn?.pk) return;

  db.exec(`
    ALTER TABLE taxpayers RENAME TO taxpayers_legacy;
    CREATE TABLE taxpayers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL DEFAULT 1,
      village_id INTEGER NOT NULL DEFAULT 1,
      property_no TEXT NOT NULL,
      old_property_no TEXT DEFAULT '',
      house_no TEXT DEFAULT '',
      holder_name TEXT DEFAULT '',
      occupant_name TEXT DEFAULT '',
      area TEXT DEFAULT '',
      category TEXT DEFAULT '',
      mobile TEXT DEFAULT '',
      description TEXT DEFAULT '',
      demand_total REAL DEFAULT 0,
      pending_tax REAL DEFAULT 0,
      current_tax REAL DEFAULT 0,
      paid INTEGER DEFAULT 0,
      village TEXT DEFAULT '',
      updated_at TEXT NOT NULL,
      UNIQUE(user_id, village_id, property_no)
    );
    INSERT INTO taxpayers (
      user_id, village_id, property_no, old_property_no, house_no, holder_name, occupant_name,
      area, category, mobile, description, demand_total, pending_tax, current_tax, paid, village, updated_at
    )
    SELECT
      1, 1, property_no, old_property_no, house_no, holder_name, occupant_name,
      area, category, mobile, description, demand_total, pending_tax, current_tax, paid, village, updated_at
    FROM taxpayers_legacy;
    DROP TABLE taxpayers_legacy;
  `);
}

migrateTaxpayersForSaas();

const allPermissions = [
  'dashboard.view',
  'excel.import',
  'excel.view',
  'billing.view',
  'billing.update',
  'broadcast.send',
  'reports.view',
  'admin.manage',
  'superadmin.view_all'
];

function seedAdminData() {
  const now = new Date().toISOString();
  const roleCount = db.prepare('SELECT COUNT(*) AS total FROM roles').get().total;
  if (!roleCount) {
    const insertRole = db.prepare('INSERT INTO roles (name, permissions_json, message_limit, created_at) VALUES (?, ?, ?, ?)');
    insertRole.run('Admin', JSON.stringify(allPermissions), 0, now);
    insertRole.run('Manager', JSON.stringify(['dashboard.view', 'excel.view', 'billing.view', 'broadcast.send', 'reports.view']), 500, now);
    insertRole.run('Operator', JSON.stringify(['dashboard.view', 'billing.view', 'broadcast.send']), 100, now);
  }

  const userCount = db.prepare('SELECT COUNT(*) AS total FROM users').get().total;
  if (!userCount) {
    const adminRole = db.prepare('SELECT id FROM roles WHERE name = ?').get('Admin');
    db.prepare('INSERT INTO users (name, username, password, role_id, active, access_expires_at, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
      .run('System Admin', 'admin', 'admin123', adminRole.id, 1, '', now);
  }

  const defaults = {
    whatsapp_phone_number_id: process.env.WHATSAPP_PHONE_NUMBER_ID || '',
    whatsapp_access_token: process.env.WHATSAPP_ACCESS_TOKEN || '',
    whatsapp_template_name: process.env.WHATSAPP_TEMPLATE_NAME || 'bill_pdf_notification',
    whatsapp_language_code: process.env.WHATSAPP_LANGUAGE_CODE || 'gu',
    daily_message_limit: '1000'
  };
  const insertSetting = db.prepare('INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)');
  Object.entries(defaults).forEach(([key, value]) => insertSetting.run(key, value));
}

seedAdminData();

function ensureColumn(table, column, definition) {
  const exists = db.prepare(`PRAGMA table_info(${table})`).all().some((item) => item.name === column);
  if (!exists) db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`);
}

ensureColumn('messages', 'user_id', 'INTEGER DEFAULT 0');
ensureColumn('messages', 'village_id', 'INTEGER DEFAULT 0');
ensureColumn('messages', 'wamid', "TEXT DEFAULT ''");
ensureColumn('excel_imports', 'user_id', 'INTEGER DEFAULT 1');
ensureColumn('excel_imports', 'village_id', 'INTEGER DEFAULT 1');
ensureColumn('excel_rows', 'user_id', 'INTEGER DEFAULT 1');
ensureColumn('excel_rows', 'village_id', 'INTEGER DEFAULT 1');
ensureColumn('users', 'package_type', "TEXT DEFAULT 'monthly'");
ensureColumn('users', 'package_limit', 'INTEGER DEFAULT 1000');
ensureColumn('users', 'bonus_limit', 'INTEGER DEFAULT 0');
ensureColumn('users', 'package_started_at', 'TEXT DEFAULT ""');
ensureColumn('users', 'package_expires_at', 'TEXT DEFAULT ""');

if (!db.prepare('SELECT COUNT(*) AS total FROM villages').get().total) {
  db.prepare('INSERT INTO villages (id, user_id, name, taluka, district, created_at) VALUES (?, ?, ?, ?, ?, ?)')
    .run(1, 1, 'KHIJADIYA ગ્રામપંચાયત', 'Wankaner', 'MORBI', new Date().toISOString());
}

function ensureAdminSuperPermission() {
  const admin = db.prepare("SELECT * FROM roles WHERE name = 'Admin'").get();
  if (!admin) return;
  const permissions = parsePermissions(admin.permissions_json);
  if (!permissions.includes('superadmin.view_all')) {
    permissions.push('superadmin.view_all');
    db.prepare('UPDATE roles SET permissions_json = ? WHERE id = ?').run(JSON.stringify(permissions), admin.id);
  }
}

const taxpayerUpsert = db.prepare(`
INSERT INTO taxpayers (
  user_id, village_id, property_no, old_property_no, house_no, holder_name, occupant_name, area, category,
  mobile, description, demand_total, pending_tax, current_tax, paid, village, updated_at
) VALUES (
  ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
)
ON CONFLICT(user_id, village_id, property_no) DO UPDATE SET
  old_property_no = CASE WHEN excluded.old_property_no != '' THEN excluded.old_property_no ELSE taxpayers.old_property_no END,
  house_no = CASE WHEN excluded.house_no != '' THEN excluded.house_no ELSE taxpayers.house_no END,
  holder_name = CASE WHEN excluded.holder_name != '' THEN excluded.holder_name ELSE taxpayers.holder_name END,
  occupant_name = CASE WHEN excluded.occupant_name != '' THEN excluded.occupant_name ELSE taxpayers.occupant_name END,
  area = CASE WHEN excluded.area != '' THEN excluded.area ELSE taxpayers.area END,
  category = CASE WHEN excluded.category != '' THEN excluded.category ELSE taxpayers.category END,
  mobile = CASE WHEN excluded.mobile != '' THEN excluded.mobile ELSE taxpayers.mobile END,
  description = CASE WHEN excluded.description != '' THEN excluded.description ELSE taxpayers.description END,
  demand_total = CASE WHEN excluded.demand_total != 0 THEN excluded.demand_total ELSE taxpayers.demand_total END,
  pending_tax = CASE WHEN excluded.pending_tax != 0 THEN excluded.pending_tax ELSE taxpayers.pending_tax END,
  current_tax = CASE WHEN excluded.current_tax != 0 THEN excluded.current_tax ELSE taxpayers.current_tax END,
  paid = CASE WHEN excluded.paid != taxpayers.paid THEN excluded.paid ELSE taxpayers.paid END,
  village = CASE WHEN excluded.village != '' THEN excluded.village ELSE taxpayers.village END,
  updated_at = excluded.updated_at
`);

const sourceRowInsert = db.prepare(`
INSERT INTO excel_rows (
  import_id, user_id, village_id, kind, property_no, receipt_no, receipt_date, old_property_no, house_no,
  holder_name, occupant_name, area, mobile, category, description,
  previous_total, current_total, grand_total, taxes_json
) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
`);

function canViewAll(user) {
  return user?.permissions?.includes('superadmin.view_all');
}

function scopedQuery(query = {}, user = null) {
  return {
    ...query,
    userId: canViewAll(user) ? query.userId : user?.id,
    villageId: query.villageId || query.village_id
  };
}

export function saveImport({ kind, fileName, extracted, village, userId = 1, villageId = 1 }) {
  const now = new Date().toISOString();
  const insertImport = db.prepare('INSERT INTO excel_imports (kind, file_name, row_count, imported_at, user_id, village_id) VALUES (?, ?, ?, ?, ?, ?)');

  db.exec('BEGIN');
  try {
    const importResult = insertImport.run(kind, fileName, extracted.rows.length, now, userId, villageId);
    const importId = Number(importResult.lastInsertRowid);

    for (const row of extracted.rows) {
      sourceRowInsert.run(
        importId,
        userId,
        villageId,
        kind,
        row.propertyNo || '',
        row.receiptNo || '',
        row.receiptDate || '',
        row.oldPropertyNo || '',
        row.houseNo || '',
        row.holderName || '',
        row.occupantName || '',
        row.area || '',
        row.mobile || '',
        row.category || '',
        row.description || '',
        money(row.previousTotal),
        money(row.currentTotal),
        money(row.grandTotal || row.pendingTax || row.currentTax),
        JSON.stringify(row.taxes || {})
      );
    }

    for (const item of extracted.taxpayers) {
      taxpayerUpsert.run(
        userId,
        villageId,
        item.propertyNo,
        item.oldPropertyNo || '',
        item.houseNo || '',
        item.holderName || '',
        item.occupantName || '',
        item.area || '',
        item.category || '',
        item.mobile || '',
        item.description || '',
        money(item.demandTotal),
        money(item.pendingTax),
        money(item.currentTax || item.pendingTax || item.demandTotal),
        item.paid ? 1 : 0,
        item.village || village || '',
        now
      );
    }

    db.exec('COMMIT');
    return { importId, imported: extracted.rows.length, total: getDashboard().total };
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

function filterSql(query = {}) {
  const clauses = [];
  const params = [];
  const col = (name) => `taxpayers.${name}`;

  if (query.q) {
    clauses.push(`(${col('property_no')} LIKE ? OR ${col('holder_name')} LIKE ? OR ${col('occupant_name')} LIKE ? OR ${col('mobile')} LIKE ? OR ${col('area')} LIKE ? OR ${col('category')} LIKE ? OR ${col('description')} LIKE ?)`);
    const value = `%${query.q}%`;
    params.push(value, value, value, value, value, value, value);
  }
  if (query.userId) {
    clauses.push(`${col('user_id')} = ?`);
    params.push(Number(query.userId));
  }
  if (query.villageId) {
    clauses.push(`${col('village_id')} = ?`);
    params.push(Number(query.villageId));
  }
  if (query.propertyNo) {
    clauses.push(`${col('property_no')} LIKE ?`);
    params.push(`%${query.propertyNo}%`);
  }
  if (query.houseNo) {
    clauses.push(`${col('house_no')} LIKE ?`);
    params.push(`%${query.houseNo}%`);
  }
  if (query.ownerName) {
    clauses.push(`${col('holder_name')} LIKE ?`);
    params.push(`%${query.ownerName}%`);
  }
  if (query.occupantName) {
    clauses.push(`${col('occupant_name')} LIKE ?`);
    params.push(`%${query.occupantName}%`);
  }
  if (query.mobile) {
    clauses.push(`${col('mobile')} LIKE ?`);
    params.push(`%${query.mobile}%`);
  }
  if (query.area) {
    clauses.push(`${col('area')} = ?`);
    params.push(query.area);
  }
  if (query.category) {
    clauses.push(`${col('category')} = ?`);
    params.push(query.category);
  }
  if (query.status === 'paid' || query.status === 'unpaid') {
    clauses.push(`${col('paid')} = ?`);
    params.push(query.status === 'paid' ? 1 : 0);
  }
  if (query.minAmount) {
    clauses.push(`COALESCE(NULLIF(${col('pending_tax')}, 0), ${col('current_tax')}, ${col('demand_total')}, 0) >= ?`);
    params.push(Number(query.minAmount));
  }
  if (query.maxAmount) {
    clauses.push(`COALESCE(NULLIF(${col('pending_tax')}, 0), ${col('current_tax')}, ${col('demand_total')}, 0) <= ?`);
    params.push(Number(query.maxAmount));
  }
  if (query.messageStatus === 'not_sent') {
    clauses.push('NOT EXISTS (SELECT 1 FROM messages WHERE messages.property_no = taxpayers.property_no)');
  } else if (query.messageStatus === 'sent') {
    clauses.push('EXISTS (SELECT 1 FROM messages WHERE messages.property_no = taxpayers.property_no)');
  } else if (['pending', 'delivered', 'failed'].includes(query.messageStatus)) {
    clauses.push(`EXISTS (
      SELECT 1 FROM messages
      WHERE messages.property_no = taxpayers.property_no
      AND messages.status = ?
    )`);
    params.push(query.messageStatus);
  }

  return {
    where: clauses.length ? `WHERE ${clauses.join(' AND ')}` : '',
    params
  };
}

function mapTaxpayer(row) {
  return {
    propertyNo: row.property_no,
    userId: row.user_id,
    villageId: row.village_id,
    oldPropertyNo: row.old_property_no,
    houseNo: row.house_no,
    holderName: row.holder_name,
    occupantName: row.occupant_name,
    area: row.area,
    category: row.category,
    mobile: row.mobile,
    description: row.description,
    demandTotal: row.demand_total,
    pendingTax: row.pending_tax,
    currentTax: row.current_tax,
    paid: Boolean(row.paid),
    village: row.village
  };
}

export function getTaxpayers(query = {}, user = null) {
  query = scopedQuery(query, user);
  const { where, params } = filterSql(query);
  return db.prepare(`SELECT * FROM taxpayers ${where} ORDER BY CAST(property_no AS INTEGER), property_no`).all(...params).map(mapTaxpayer);
}

export function getTaxpayer(propertyNo, query = {}, user = null) {
  query = scopedQuery(query, user);
  const clauses = ['property_no = ?'];
  const params = [propertyNo];
  if (query.userId) { clauses.push('user_id = ?'); params.push(Number(query.userId)); }
  if (query.villageId) { clauses.push('village_id = ?'); params.push(Number(query.villageId)); }
  const row = db.prepare(`SELECT * FROM taxpayers WHERE ${clauses.join(' AND ')}`).get(...params);
  return row ? mapTaxpayer(row) : null;
}

export function setPaidStatus(propertyNo, paid, query = {}, user = null) {
  query = scopedQuery(query, user);
  db.prepare('UPDATE taxpayers SET paid = ?, updated_at = ? WHERE property_no = ? AND user_id = COALESCE(?, user_id) AND village_id = COALESCE(?, village_id)')
    .run(paid ? 1 : 0, new Date().toISOString(), propertyNo, query.userId || null, query.villageId || null);
  return getTaxpayer(propertyNo, query, user);
}

export function getDashboard(query = {}, user = null) {
  query = scopedQuery(query, user);
  const { where, params } = filterSql(query);
  const row = db.prepare(`
    SELECT
      COUNT(*) AS total,
      SUM(CASE WHEN paid = 1 THEN 1 ELSE 0 END) AS paid,
      SUM(CASE WHEN mobile != '' THEN 1 ELSE 0 END) AS mobileReady,
      SUM(CASE WHEN paid = 0 THEN COALESCE(NULLIF(pending_tax, 0), current_tax, demand_total, 0) ELSE 0 END) AS pendingAmount
    FROM taxpayers
    ${where}
  `).get(...params);

  return {
    total: row.total || 0,
    paid: row.paid || 0,
    unpaid: (row.total || 0) - (row.paid || 0),
    mobileReady: row.mobileReady || 0,
    pendingAmount: row.pendingAmount || 0
  };
}

export function getImportHistory(query = {}, user = null) {
  query = scopedQuery(query, user);
  const clauses = [];
  const params = [];
  if (query.userId) { clauses.push('excel_imports.user_id = ?'); params.push(Number(query.userId)); }
  if (query.villageId) { clauses.push('excel_imports.village_id = ?'); params.push(Number(query.villageId)); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  return db.prepare(`
    SELECT excel_imports.*, villages.name AS village_name, users.name AS user_name
    FROM excel_imports
    LEFT JOIN villages ON villages.id = excel_imports.village_id
    LEFT JOIN users ON users.id = excel_imports.user_id
    ${where}
    ORDER BY excel_imports.id DESC LIMIT 50
  `).all(...params);
}

export function getSourceRows(kind, query = {}, user = null) {
  query = scopedQuery(query, user);
  const params = [];
  const clauses = [];
  if (kind) {
    clauses.push('kind = ?');
    params.push(kind);
  }
  if (query.userId) { clauses.push('user_id = ?'); params.push(Number(query.userId)); }
  if (query.villageId) { clauses.push('village_id = ?'); params.push(Number(query.villageId)); }
  if (query.importId) { clauses.push('import_id = ?'); params.push(Number(query.importId)); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  return db.prepare(`
    SELECT * FROM excel_rows
    ${where}
    ORDER BY CAST(COALESCE(NULLIF(property_no, ''), receipt_no) AS INTEGER), id
    LIMIT 1000
  `).all(...params);
}

export function addMessage(row) {
  db.prepare('INSERT INTO messages (user_id, village_id, property_no, mobile, message, status, detail, sent_at, wamid) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)')
    .run(row.userId || 0, row.villageId || 0, row.propertyNo || '', row.mobile || '', row.message || '', row.status || '', typeof row.detail === 'string' ? row.detail : JSON.stringify(row.detail || ''), row.sentAt, row.wamid || '');
}

export function updateMessageStatusByWamid(wamid, status, detail) {
  db.prepare('UPDATE messages SET status = ?, detail = ? WHERE wamid = ?')
    .run(status, typeof detail === 'string' ? detail : JSON.stringify(detail || ''), wamid);
}

export function getMessages(query = {}, user = null) {
  query = scopedQuery(query, user);
  const clauses = [];
  const params = [];
  if (query.userId) { clauses.push('user_id = ?'); params.push(Number(query.userId)); }
  if (query.villageId) { clauses.push('village_id = ?'); params.push(Number(query.villageId)); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  return db.prepare(`SELECT property_no AS propertyNo, mobile, message, status, detail, sent_at AS sentAt FROM messages ${where} ORDER BY id DESC LIMIT 200`).all(...params);
}

export function getWhatsappStatus(query = {}) {
  const { where, params } = filterSql(query);
  return db.prepare(`
    SELECT
      taxpayers.property_no AS propertyNo,
      taxpayers.house_no AS houseNo,
      taxpayers.holder_name AS holderName,
      taxpayers.occupant_name AS occupantName,
      taxpayers.area,
      taxpayers.category,
      taxpayers.mobile,
      taxpayers.pending_tax AS pendingTax,
      taxpayers.current_tax AS currentTax,
      taxpayers.paid,
      latest.status AS messageStatus,
      latest.message AS lastMessage,
      latest.sent_at AS lastSentAt,
      latest.detail AS detail,
      COALESCE(counts.total_sent, 0) AS totalSent
    FROM taxpayers
    LEFT JOIN (
      SELECT messages.*
      FROM messages
      JOIN (
        SELECT property_no, MAX(id) AS id
        FROM messages
        GROUP BY property_no, user_id, village_id
      ) last_message ON last_message.id = messages.id
    ) latest ON latest.property_no = taxpayers.property_no AND latest.user_id = taxpayers.user_id AND latest.village_id = taxpayers.village_id
    LEFT JOIN (
      SELECT property_no, COUNT(*) AS total_sent
      FROM messages
      GROUP BY property_no, user_id, village_id
    ) counts ON counts.property_no = taxpayers.property_no AND counts.user_id = taxpayers.user_id AND counts.village_id = taxpayers.village_id
    ${where}
    ORDER BY CAST(taxpayers.property_no AS INTEGER), taxpayers.property_no
    LIMIT 1500
  `).all(...params).map((row) => ({
    ...row,
    paid: Boolean(row.paid),
    messageStatus: row.messageStatus || 'not_sent'
  }));
}

export function getWhatsappSummary(query = {}) {
  const rows = getWhatsappStatus(query);
  return {
    total: rows.length,
    sent: rows.filter((row) => row.messageStatus !== 'not_sent').length,
    notSent: rows.filter((row) => row.messageStatus === 'not_sent').length,
    pending: rows.filter((row) => row.messageStatus === 'pending').length,
    delivered: rows.filter((row) => row.messageStatus === 'delivered').length,
    failed: rows.filter((row) => row.messageStatus === 'failed').length
  };
}

export function getWhatsappStatusScoped(query = {}, user = null) {
  return getWhatsappStatus(scopedQuery(query, user));
}

export function getWhatsappSummaryScoped(query = {}, user = null) {
  return getWhatsappSummary(scopedQuery(query, user));
}

function parsePermissions(value) {
  try {
    return JSON.parse(value || '[]');
  } catch {
    return [];
  }
}

ensureAdminSuperPermission();

function mapRole(row) {
  return {
    id: row.id,
    name: row.name,
    permissions: parsePermissions(row.permissions_json),
    messageLimit: row.message_limit,
    createdAt: row.created_at
  };
}

function mapUser(row) {
  return {
    id: row.id,
    name: row.name,
    username: row.username,
    roleId: row.role_id,
    roleName: row.role_name,
    permissions: parsePermissions(row.permissions_json),
    messageLimit: row.message_limit,
    active: Boolean(row.active),
    accessExpiresAt: row.access_expires_at,
    packageType: row.package_type,
    packageLimit: row.package_limit,
    bonusLimit: row.bonus_limit,
    packageStartedAt: row.package_started_at,
    packageExpiresAt: row.package_expires_at,
    createdAt: row.created_at
  };
}

export function getPermissionsList() {
  return allPermissions;
}

export function getRoles() {
  return db.prepare('SELECT * FROM roles ORDER BY id').all().map(mapRole);
}

export function saveRole(payload) {
  const permissions = JSON.stringify(payload.permissions || []);
  const limit = Number(payload.messageLimit || 0);
  if (payload.id) {
    db.prepare('UPDATE roles SET name = ?, permissions_json = ?, message_limit = ? WHERE id = ?')
      .run(payload.name, permissions, limit, payload.id);
    return getRoles().find((role) => role.id === Number(payload.id));
  }
  const result = db.prepare('INSERT INTO roles (name, permissions_json, message_limit, created_at) VALUES (?, ?, ?, ?)')
    .run(payload.name, permissions, limit, new Date().toISOString());
  return getRoles().find((role) => role.id === Number(result.lastInsertRowid));
}

export function deleteRole(id) {
  const role = db.prepare('SELECT * FROM roles WHERE id = ?').get(id);
  if (!role) throw new Error('Role not found.');
  if (role.name === 'Admin') throw new Error('Admin role cannot be deleted.');
  const inUse = db.prepare('SELECT COUNT(*) AS total FROM users WHERE role_id = ?').get(id).total;
  if (inUse) throw new Error('Role is assigned to users. Move users before deleting.');
  db.prepare('DELETE FROM roles WHERE id = ?').run(id);
  return { deleted: true };
}

export function getUsers() {
  return db.prepare(`
    SELECT users.*, roles.name AS role_name, roles.permissions_json, roles.message_limit
    FROM users
    JOIN roles ON roles.id = users.role_id
    ORDER BY users.id
  `).all().map(mapUser);
}

export function getUser(id) {
  const row = db.prepare(`
    SELECT users.*, roles.name AS role_name, roles.permissions_json, roles.message_limit
    FROM users
    JOIN roles ON roles.id = users.role_id
    WHERE users.id = ?
  `).get(id);
  return row ? mapUser(row) : null;
}

export function loginUser(username, password) {
  const row = db.prepare(`
    SELECT users.*, roles.name AS role_name, roles.permissions_json, roles.message_limit
    FROM users
    JOIN roles ON roles.id = users.role_id
    WHERE users.username = ? AND users.password = ? AND users.active = 1
  `).get(username, password);
  if (!row) return null;
  if (row.access_expires_at && new Date(row.access_expires_at) < new Date()) return null;
  return mapUser(row);
}

export function saveUser(payload) {
  const now = new Date().toISOString();
  const active = payload.active === false ? 0 : 1;
  if (payload.id) {
    const existing = db.prepare('SELECT password FROM users WHERE id = ?').get(payload.id);
    db.prepare(`UPDATE users SET
      name = ?, username = ?, password = ?, role_id = ?, active = ?, access_expires_at = ?,
      package_type = ?, package_limit = ?, bonus_limit = ?, package_started_at = ?, package_expires_at = ?
      WHERE id = ?`)
      .run(payload.name, payload.username, payload.password || existing.password, payload.roleId, active, payload.accessExpiresAt || '',
        payload.packageType || 'monthly', Number(payload.packageLimit || 1000), Number(payload.bonusLimit || 0), payload.packageStartedAt || '', payload.packageExpiresAt || '', payload.id);
    return getUser(payload.id);
  }
  const result = db.prepare(`INSERT INTO users
    (name, username, password, role_id, active, access_expires_at, package_type, package_limit, bonus_limit, package_started_at, package_expires_at, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`)
    .run(payload.name, payload.username, payload.password || '123456', payload.roleId, active, payload.accessExpiresAt || '',
      payload.packageType || 'monthly', Number(payload.packageLimit || 1000), Number(payload.bonusLimit || 0), payload.packageStartedAt || now.slice(0, 10), payload.packageExpiresAt || '', now);
  return getUser(Number(result.lastInsertRowid));
}

export function deleteUser(id) {
  const user = getUser(id);
  if (!user) throw new Error('User not found.');
  if (user.permissions.includes('superadmin.view_all')) throw new Error('Super admin user cannot be deleted.');
  db.prepare('UPDATE users SET active = 0 WHERE id = ?').run(id);
  return { deleted: true };
}

export function resetUserPassword(id, password) {
  const nextPassword = String(password || '').trim();
  if (!nextPassword) throw new Error('Password is required.');
  const user = getUser(id);
  if (!user) throw new Error('User not found.');
  db.prepare('UPDATE users SET password = ? WHERE id = ?').run(nextPassword, id);
  return { reset: true };
}

export function getSettings() {
  const rows = db.prepare('SELECT key, value FROM app_settings ORDER BY key').all();
  return Object.fromEntries(rows.map((row) => [row.key, row.value]));
}

export function saveSettings(settings) {
  const stmt = db.prepare('INSERT INTO app_settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value');
  Object.entries(settings).forEach(([key, value]) => stmt.run(key, String(value ?? '')));
  return getSettings();
}

export function getSentCountToday(userId) {
  const start = new Date();
  start.setHours(0, 0, 0, 0);
  return db.prepare('SELECT COUNT(*) AS total FROM messages WHERE user_id = ? AND sent_at >= ?').get(userId || 0, start.toISOString()).total || 0;
}

export function getSentCountForPackage(userId) {
  const user = getUser(userId);
  const start = user?.packageStartedAt || new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString();
  const end = user?.packageExpiresAt || '';
  const clauses = ['user_id = ?', 'sent_at >= ?'];
  const params = [userId || 0, start];
  if (end) {
    clauses.push('sent_at <= ?');
    params.push(end);
  }
  return db.prepare(`SELECT COUNT(*) AS total FROM messages WHERE ${clauses.join(' AND ')}`).get(...params).total || 0;
}

export function getEffectiveMessageLimit(user) {
  return Number(user?.packageLimit || 0) + Number(user?.bonusLimit || 0);
}

export function getVillages(query = {}, user = null) {
  query = scopedQuery(query, user);
  const clauses = [];
  const params = [];
  if (query.userId) { clauses.push('villages.user_id = ?'); params.push(Number(query.userId)); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  return db.prepare(`
    SELECT villages.*, users.name AS user_name
    FROM villages
    LEFT JOIN users ON users.id = villages.user_id
    ${where}
    ORDER BY villages.name
  `).all(...params);
}

export function saveVillage(payload, user = null) {
  const userId = canViewAll(user) ? Number(payload.userId || user.id) : user.id;
  if (payload.id) {
    db.prepare('UPDATE villages SET name = ?, taluka = ?, district = ? WHERE id = ? AND user_id = COALESCE(?, user_id)')
      .run(payload.name, payload.taluka || '', payload.district || '', payload.id, canViewAll(user) ? null : user.id);
    return getVillages({ userId }, user).find((item) => item.id === Number(payload.id));
  }
  const result = db.prepare('INSERT INTO villages (user_id, name, taluka, district, created_at) VALUES (?, ?, ?, ?, ?)')
    .run(userId, payload.name, payload.taluka || '', payload.district || '', new Date().toISOString());
  return getVillages({ userId }, user).find((item) => item.id === Number(result.lastInsertRowid));
}

export function getCoupons(query = {}, user = null) {
  const clauses = [];
  const params = [];
  if (!canViewAll(user)) {
    clauses.push('(assigned_user_id = ? OR redeemed_user_id = ?)');
    params.push(user.id, user.id);
  } else if (query.userId) {
    clauses.push('(assigned_user_id = ? OR redeemed_user_id = ?)');
    params.push(Number(query.userId), Number(query.userId));
  }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  return db.prepare(`SELECT * FROM coupons ${where} ORDER BY id DESC LIMIT 100`).all(...params);
}

export function createCoupon(payload) {
  const code = String(payload.code || `CPN${Date.now()}`).trim().toUpperCase();
  const result = db.prepare(`INSERT INTO coupons
    (code, message_credit, assigned_user_id, status, expires_at, created_at)
    VALUES (?, ?, ?, 'active', ?, ?)`)
    .run(code, Number(payload.messageCredit || 0), Number(payload.assignedUserId || 0), payload.expiresAt || '', new Date().toISOString());
  return db.prepare('SELECT * FROM coupons WHERE id = ?').get(Number(result.lastInsertRowid));
}

export function redeemCoupon(code, user) {
  const coupon = db.prepare('SELECT * FROM coupons WHERE code = ?').get(String(code || '').trim().toUpperCase());
  if (!coupon) throw new Error('Invalid coupon code.');
  if (coupon.status !== 'active') throw new Error('Coupon is not active.');
  if (coupon.expires_at && new Date(coupon.expires_at) < new Date()) throw new Error('Coupon expired.');
  if (coupon.assigned_user_id && coupon.assigned_user_id !== user.id) throw new Error('Coupon assigned to another user.');
  db.exec('BEGIN');
  try {
    db.prepare('UPDATE users SET bonus_limit = bonus_limit + ? WHERE id = ?').run(coupon.message_credit, user.id);
    db.prepare("UPDATE coupons SET status = 'redeemed', redeemed_user_id = ?, redeemed_at = ? WHERE id = ?")
      .run(user.id, new Date().toISOString(), coupon.id);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
  return { coupon: db.prepare('SELECT * FROM coupons WHERE id = ?').get(coupon.id), user: getUser(user.id) };
}
