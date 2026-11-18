// يقرا العربيات من Firebase، يبني ملف إكسل، ويبعته بالإيميل.
// المتغيرات (بتتحط كـ Secrets في GitHub، مش في الكود):
//   FIREBASE_SERVICE_ACCOUNT  محتوى ملف المفتاح (JSON)
//   SMTP_USER / SMTP_PASS     إيميل جيميل وكلمة مرور التطبيقات
//   MAIL_TO                   الإيميلات المستلمة، مفصولة بفاصلة
import ExcelJS from 'exceljs';
import nodemailer from 'nodemailer';
import { writeFileSync } from 'node:fs';

const MOCK = process.env.REPORT_MOCK === '1';
const USERS = { zizo: 'Zizo', ussef: 'Ussef', user: 'User' };
const nameOf = k => USERS[k] || k || '';
const has = v => v != null && v !== '';
const fmtNum = n => Number(n).toLocaleString('en-US', { maximumFractionDigits: 2 });
const fmtDate = ms => ms ? new Date(ms).toLocaleDateString('en-CA', { timeZone: 'Africa/Cairo' }) : '';

async function load() {
  if (MOCK) {
    return { rate: 50, rows: [
      { chassis: 'AAA111', make: 'مان', description: 'تراكتور 2019', color: 'أبيض', buyer: 'أحمد', buyPrice: 100000, buyCur: 'EGP', extraPrice: 5000, extraCur: 'EGP', shipPrice: 10000, shipCur: 'EGP', salePrice: 130000, saleCur: 'EGP', transportDone: true, acidDone: false, shipNo: 'S-1', shipDone: false, notes: '', createdBy: 'zizo', createdAt: Date.now() - 86400000 },
      { chassis: 'BBB222', make: 'فولفو', description: '', color: 'أزرق', buyer: '', buyPrice: 2000, buyCur: 'EUR', extraPrice: null, extraCur: 'EGP', shipPrice: 300, shipCur: 'EUR', salePrice: 200000, saleCur: 'EGP', transportDone: false, acidDone: true, shipNo: '', shipDone: true, notes: 'ملاحظة', createdBy: 'ussef', createdAt: Date.now() }
    ] };
  }
  const { default: admin } = await import('firebase-admin');
  admin.initializeApp({ credential: admin.credential.cert(JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT)) });
  const db = admin.firestore();
  const snap = await db.collection('trucks').get();
  const r = await db.doc('settings/rate').get();
  return { rate: r.exists ? Number(r.data().eurToEgp) || 0 : 0, rows: snap.docs.map(d => ({ id: d.id, ...d.data() })) };
}

const { rate, rows } = await load();
const egp = (n, c) => c === 'EUR' ? (rate > 0 ? n * rate : null) : n;
function calc(r) {
  let cost = 0, ok = true, any = false;
  [['buyPrice', 'buyCur'], ['extraPrice', 'extraCur'], ['shipPrice', 'shipCur']].forEach(([p, c]) => {
    if (!has(r[p])) return; any = true;
    const e = egp(Number(r[p]), r[c]); if (e == null) ok = false; else cost += e;
  });
  const sale = has(r.salePrice) ? egp(Number(r.salePrice), r.saleCur) : null;
  return { any, cost: ok ? cost : null, profit: (ok && sale != null) ? sale - cost : null };
}

const list = rows.slice().sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));
const NC = 19, HEAD = 4, first = HEAD + 1;
const wb = new ExcelJS.Workbook(); wb.creator = 'سجل العربيات'; wb.created = new Date();
const ws = wb.addWorksheet('العربيات', { views: [{ rightToLeft: true, state: 'frozen', xSplit: 2, ySplit: HEAD, showGridLines: false }], properties: { tabColor: { argb: 'FF1F4F60' } }, pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
[6, 24, 13, 34, 12, 22, 17, 17, 17, 19, 17, 19, 11, 11, 18, 13, 34, 14, 14].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
const FONT = 'Arial', thin = { style: 'thin', color: { argb: 'FFD8DCD8' } }, border = { top: thin, left: thin, bottom: thin, right: thin };
const fill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + c } });
const al = h => ({ horizontal: h, vertical: 'middle', wrapText: true, readingOrder: 'rtl' });

ws.mergeCells(1, 1, 1, NC); const t = ws.getCell(1, 1); t.value = 'سجل العربيات';
t.font = { name: FONT, size: 20, bold: true, color: { argb: 'FFFFFFFF' } }; t.fill = fill('1F4F60'); t.alignment = al('center'); ws.getRow(1).height = 38;
const nT = list.filter(r => r.transportDone).length, nA = list.filter(r => r.acidDone).length, nS = list.filter(r => r.shipDone).length;
const today = fmtDate(Date.now());
ws.mergeCells(2, 1, 2, NC); const s = ws.getCell(2, 1);
s.value = `تاريخ التصدير: ${today}   •   عدد العربيات: ${list.length}   •   تم النقل: ${nT}   •   تم الأسيد: ${nA}   •   تم الشحن: ${nS}` + (rate > 0 ? `   •   سعر اليورو: ${fmtNum(rate)} جنيه` : '');
s.font = { name: FONT, size: 11, color: { argb: 'FF3E4A52' } }; s.fill = fill('E6EEF0'); s.alignment = al('center'); ws.getRow(2).height = 26; ws.getRow(3).height = 8;

['م', 'رقم الشاسيه', 'النوع', 'الوصف', 'اللون', 'المشتري', 'سعر الشراء', 'مصاريف إضافية', 'سعر الشحن', 'إجمالي التكلفة (ج.م)', 'سعر البيع', 'الربح (ج.م)', 'النقل', 'الأسيد', 'رقم الشحن', 'تم الشحن', 'ملاحظات', 'سجّلها', 'تاريخ التسجيل']
  .forEach((h, i) => { const c = ws.getRow(HEAD).getCell(i + 1); c.value = h; c.font = { name: FONT, size: 12, bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = fill('2B6A80'); c.border = border; c.alignment = al('center'); });
ws.getRow(HEAD).height = 30;

const cf = c => c === 'EUR' ? '#,##0.##" €"' : '#,##0.##" ج.م"', EG = '#,##0.##" ج.م"';
let sumCost = 0, sumProfit = 0;
list.forEach((r, i) => {
  const row = ws.getRow(first + i), zebra = (i % 2 === 1) ? 'F5F6F4' : 'FFFFFF', c = calc(r);
  const n = p => has(r[p]) ? Number(r[p]) : null;
  const vals = [i + 1, r.chassis || '', r.make || '', r.description || '', r.color || '', r.buyer || '', n('buyPrice'), n('extraPrice'), n('shipPrice'), c.any ? c.cost : null, n('salePrice'), c.profit,
    r.transportDone ? 'تم' : 'لا', r.acidDone ? 'تم' : 'لا', r.shipNo || '', r.shipDone ? 'تم' : 'لا', r.notes || '', nameOf(r.createdBy), fmtDate(r.createdAt)];
  if (c.any && c.cost != null) sumCost += c.cost; if (c.profit != null) sumProfit += c.profit;
  vals.forEach((v, k) => { const cell = row.getCell(k + 1); cell.value = v; cell.font = { name: FONT, size: 11, color: { argb: 'FF1B2328' } }; cell.fill = fill(zebra); cell.border = border; cell.alignment = al([3, 5, 16].includes(k + 1) ? 'right' : 'center'); });
  row.getCell(2).font = { name: 'Courier New', size: 11, bold: true, color: { argb: 'FF1B2328' } };
  row.getCell(7).numFmt = cf(r.buyCur); row.getCell(8).numFmt = cf(r.extraCur); row.getCell(9).numFmt = cf(r.shipCur); row.getCell(11).numFmt = cf(r.saleCur);
  row.getCell(10).numFmt = EG; row.getCell(12).numFmt = EG;
  if (c.profit != null) row.getCell(12).font = { name: FONT, size: 11, bold: true, color: { argb: c.profit >= 0 ? 'FF1D7449' : 'FFA63D2F' } };
  [13, 14, 16].forEach(k => { const cell = row.getCell(k), ok = cell.value === 'تم'; cell.fill = fill(ok ? 'DDF0E5' : 'F1F2F0'); cell.font = { name: FONT, size: 11, bold: ok, color: { argb: ok ? 'FF1D7449' : 'FF8A949B' } }; });
  row.height = Math.max(22, Math.min(60, Math.ceil(Math.max((r.description || '').length, (r.notes || '').length) / 32) * 16 + 8));
});
const last = first + list.length - 1, tr = last + 1;
if (list.length) {
  const lab = ws.getCell(tr, 6); lab.value = 'الإجمالي (ج.م)';
  const tc = ws.getCell(tr, 10); tc.value = { formula: `SUM(J${first}:J${last})`, result: sumCost }; tc.numFmt = EG;
  const tp = ws.getCell(tr, 12); tp.value = { formula: `SUM(L${first}:L${last})`, result: sumProfit }; tp.numFmt = EG;
  [lab, tc, tp].forEach(c => { c.font = { name: FONT, size: 12, bold: true, color: { argb: 'FF3D2F05' } }; c.fill = fill('FBEBC0'); c.border = border; c.alignment = al('center'); });
  ws.getRow(tr).height = 26;
  ws.autoFilter = { from: { row: HEAD, column: 1 }, to: { row: last, column: NC } };
}
const buf = Buffer.from(await wb.xlsx.writeBuffer());
const filename = `سجل-العربيات-${today}.xlsx`;

if (MOCK) { writeFileSync('mock-report.xlsx', buf); console.log('mock ok', list.length, 'rows,', buf.length, 'bytes, cost', sumCost, 'profit', sumProfit); process.exit(0); }

const to = (process.env.MAIL_TO || '').split(',').map(x => x.trim()).filter(Boolean);
if (!to.length) throw new Error('MAIL_TO فاضي');
const mailer = nodemailer.createTransport({ service: 'gmail', auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS } });
await mailer.sendMail({
  from: `"سجل العربيات" <${process.env.SMTP_USER}>`, to,
  subject: `سجل العربيات - ${today}`,
  text: `مرفق آخر نسخة من سجل العربيات.\n\nعدد العربيات: ${list.length}\nتم النقل: ${nT}\nتم الأسيد: ${nA}\nتم الشحن: ${nS}\nإجمالي الربح: ${fmtNum(sumProfit)} ج.م`,
  attachments: [{ filename, content: buf, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }]
});
console.log('sent to', to.length, 'recipients,', list.length, 'rows');
