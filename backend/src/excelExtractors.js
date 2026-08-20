import ExcelJS from 'exceljs';

const TAX_COLUMNS = [
  'houseTax',
  'sapaniTax',
  'khapaniTax',
  'cleaningTax',
  'gutterTax',
  'lightTax',
  'notice',
  'advance',
  'other1',
  'other2',
  'other3',
  'other4',
  'other5'
];

export const clean = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();
export const money = (value) => Number.parseFloat(value || 0) || 0;

export function normalizeProperty(value) {
  const text = clean(value);
  if (!text) return '';
  return text.endsWith('.0') ? text.slice(0, -2) : text;
}

function isNumericProperty(value) {
  return /^\d+$/.test(normalizeProperty(value));
}

function cellText(cell) {
  if (!cell) return '';
  const value = cell.value;
  if (value == null) return '';
  if (typeof value === 'object') {
    if (value.richText) return clean(value.richText.map((part) => part.text).join(''));
    if (value.text) return clean(value.text);
    if (value.result != null) return clean(value.result);
    return '';
  }
  return clean(value);
}

function rowValues(row, columnCount) {
  const values = [];
  for (let index = 1; index <= columnCount; index += 1) {
    values[index - 1] = cellText(row.getCell(index));
  }
  return values;
}

function toRows(sheet) {
  const rows = [];
  for (let index = 1; index <= sheet.rowCount; index += 1) {
    rows[index - 1] = rowValues(sheet.getRow(index), sheet.columnCount);
  }
  return rows;
}

function parseBracketFields(value) {
  return [...clean(value).matchAll(/\[([^\]]*)\]/g)].map((match) => clean(match[1]));
}

function splitOwnerOccupant(value) {
  const [holderName = '', occupantName = ''] = clean(value).split('/').map(clean);
  return { holderName, occupantName };
}

function taxBreakup(row, startIndex) {
  const taxes = {};
  let previousTotal = 0;
  let currentTotal = 0;

  TAX_COLUMNS.forEach((key, index) => {
    const previous = money(row[startIndex + (index * 2)]);
    const current = money(row[startIndex + (index * 2) + 1]);
    taxes[key] = { previous, current, total: previous + current };
    previousTotal += previous;
    currentTotal += current;
  });

  return { taxes, previousTotal, currentTotal, grandTotal: previousTotal + currentTotal };
}

async function workbookRows(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.worksheets[0];
  return toRows(sheet);
}

export async function extractMobileWorkbook(buffer) {
  const rows = await workbookRows(buffer);
  const headerIndex = rows.findIndex((row) => row.some((value) => clean(value).includes('મિલ્કત નંબર')) && row.some((value) => clean(value).includes('મોબાઇલ')));
  const dataRows = rows.slice(headerIndex + 1);
  const records = dataRows
    .map((row) => ({
      propertyNo: normalizeProperty(row[0]),
      holderName: clean(row[1]),
      occupantName: clean(row[2]),
      pendingTax: money(row[3]),
      mobile: clean(row[4]),
      category: clean(row[5]),
      description: clean(row[6])
    }))
    .filter((item) => item.propertyNo);

  return {
    kind: 'mobile',
    taxpayers: records,
    rows: records
  };
}

export async function extractDemandWorkbook(buffer) {
  const rows = await workbookRows(buffer);
  const records = [];

  for (let index = 7; index < rows.length; index += 3) {
    const infoRow = rows[index] || [];
    const amountRow = rows[index + 1] || [];
    const propertyNo = normalizeProperty(infoRow[0]);
    if (!isNumericProperty(propertyNo)) continue;

    const names = splitOwnerOccupant(infoRow[1]);
    const totals = taxBreakup(amountRow, 1);
    records.push({
      propertyNo,
      oldPropertyNo: '',
      holderName: names.holderName,
      occupantName: names.occupantName,
      ...totals
    });
  }

  return {
    kind: 'demand',
    taxpayers: records.map((item) => ({
      propertyNo: item.propertyNo,
      oldPropertyNo: item.oldPropertyNo,
      holderName: item.holderName,
      occupantName: item.occupantName,
      currentTax: item.grandTotal,
      demandTotal: item.grandTotal
    })),
    rows: records
  };
}

export async function extractPendingWorkbook(buffer) {
  const rows = await workbookRows(buffer);
  const records = [];

  for (let index = 6; index < rows.length; index += 3) {
    const infoRow = rows[index] || [];
    const amountRow = rows[index + 1] || [];
    const propertyNo = normalizeProperty(infoRow[0]);
    if (!isNumericProperty(propertyNo)) continue;

    const [houseNo = '', area = '', holderName = '', occupantName = ''] = parseBracketFields(infoRow[2]);
    const totals = taxBreakup(amountRow, 2);
    records.push({
      propertyNo,
      houseNo: normalizeProperty(infoRow[1] || houseNo),
      area,
      holderName,
      occupantName,
      ...totals
    });
  }

  return {
    kind: 'pending',
    taxpayers: records.map((item) => ({
      propertyNo: item.propertyNo,
      houseNo: item.houseNo,
      area: item.area,
      holderName: item.holderName,
      occupantName: item.occupantName,
      pendingTax: item.grandTotal
    })),
    rows: records
  };
}

export async function extractRecoveryWorkbook(buffer) {
  const rows = await workbookRows(buffer);
  const records = [];

  for (let index = 6; index < rows.length; index += 3) {
    const infoRow = rows[index] || [];
    const amountRow = rows[index + 1] || [];
    const receiptNo = normalizeProperty(infoRow[0]);
    const receiptDate = clean(amountRow[0]);
    const [propertyNo = '', oldPropertyNo = '', area = '', holderName = '', occupantName = '', houseNo = ''] = parseBracketFields(infoRow[2]);
    if (!receiptNo || !propertyNo) continue;

    const totals = taxBreakup(amountRow, 2);
    records.push({
      receiptNo,
      receiptDate,
      propertyNo: normalizeProperty(propertyNo),
      oldPropertyNo: normalizeProperty(oldPropertyNo),
      houseNo: normalizeProperty(houseNo),
      area,
      holderName,
      occupantName,
      ...totals
    });
  }

  return {
    kind: 'recovery',
    taxpayers: records.map((item) => ({
      propertyNo: item.propertyNo,
      oldPropertyNo: item.oldPropertyNo,
      houseNo: item.houseNo,
      area: item.area,
      holderName: item.holderName,
      occupantName: item.occupantName,
      paid: true
    })),
    rows: records
  };
}

export async function extractWorkbook(buffer, kind) {
  if (kind === 'mobile') return extractMobileWorkbook(buffer);
  if (kind === 'demand') return extractDemandWorkbook(buffer);
  if (kind === 'pending') return extractPendingWorkbook(buffer);
  if (kind === 'recovery') return extractRecoveryWorkbook(buffer);
  throw new Error(`Unsupported Excel import type: ${kind}`);
}
