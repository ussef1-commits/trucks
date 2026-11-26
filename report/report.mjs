// يقرا المركبات (عربيات وبرادات) من Firebase، يبني ملف إكسل، ويبعته بالإيميل. اليورو فقط.
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

/* ===== shared by the page and the daily report: money maths + Excel layout (EUR only) ===== */
const has = v => v != null && v !== '';
const nf = n => Number(n).toLocaleString('en-US', { maximumFractionDigits: 2 });
const dstr = ms => ms ? new Date(ms).toLocaleDateString('en-CA', { timeZone: 'Africa/Cairo' }) : '';
const VLABEL = { car: 'عربية', reefer: 'براد' };
const vtypeOf = r => r && r.vehicleType === 'reefer' ? 'reefer' : 'car';
function calcRow(r) {
  let cost = 0, any = false;
  ['buyPrice', 'extraPrice', 'shipPrice', 'transportPrice'].forEach(p => { if (has(r[p])) { any = true; cost += Number(r[p]); } });
  const sale = has(r.salePrice) ? Number(r.salePrice) : null;
  return { any, cost: any ? cost : null, profit: sale != null ? sale - cost : null };
}
function isoToDate(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s || '')) return null;
  const p = s.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2]));
}
async function buildWorkbookBuffer(ExcelJS, list, nameOf) {
  const EUR = '#,##0.##" €"', DATEF = 'yyyy-mm-dd', FONT = 'Arial';
  const thin = { style: 'thin', color: { argb: 'FFD8DCD8' } }, border = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + c } });
  const al = h => ({ horizontal: h, vertical: 'middle', wrapText: true, readingOrder: 'rtl' });
  const num = p => r => has(r[p]) ? Number(r[p]) : null;
  const dt = p => r => isoToDate(r[p]);
  const cols = [
    { h: 'م', w: 6, v: (r, i) => i + 1, a: 'center' },
    { h: 'رقم الشاسيه', w: 24, v: r => r.chassis || '', a: 'center', mono: true },
    { h: 'المركبة', w: 11, v: r => VLABEL[vtypeOf(r)], a: 'center' },
    { h: 'النوع', w: 20, v: r => r.make || '', a: 'center' },
    { h: 'الوصف', w: 34, v: r => r.description || '', a: 'right' },
    { h: 'اللون', w: 12, v: r => r.color || '', a: 'center' },
    { h: 'المشتري', w: 22, v: r => r.buyer || '', a: 'right' },
    { h: 'رقم المشتري', w: 18, v: r => r.buyerPhone || '', a: 'center', text: true },
    { h: 'الشركة البائعة', w: 24, v: r => r.sellerName || '', a: 'right' },
    { h: 'هاتف الشركة', w: 18, v: r => r.sellerPhone || '', a: 'center', text: true },
    { h: 'سعر الشراء (€)', w: 16, v: num('buyPrice'), a: 'center', fmt: EUR },
    { h: 'مصاريف النقل (€)', w: 17, v: num('transportPrice'), a: 'center', fmt: EUR },
    { h: 'سعر الشحن (€)', w: 16, v: num('shipPrice'), a: 'center', fmt: EUR },
    { h: 'مصاريف إضافية (€)', w: 18, v: num('extraPrice'), a: 'center', fmt: EUR },
    { h: 'إجمالي التكلفة (€)', w: 19, v: (r, i, c) => c.any ? c.cost : null, a: 'center', fmt: EUR, key: 'cost' },
    { h: 'سعر البيع (€)', w: 16, v: num('salePrice'), a: 'center', fmt: EUR },
    { h: 'الربح (€)', w: 16, v: (r, i, c) => c.profit, a: 'center', fmt: EUR, key: 'profit' },
    { h: 'تاريخ الشراء', w: 14, v: dt('purchaseDate'), a: 'center', fmt: DATEF },
    { h: 'تاريخ البيع', w: 14, v: dt('saleDate'), a: 'center', fmt: DATEF },
    { h: 'تاريخ النقل', w: 14, v: dt('transportDate'), a: 'center', fmt: DATEF },
    { h: 'تاريخ الشحن', w: 14, v: dt('shipDate'), a: 'center', fmt: DATEF },
    { h: 'تاريخ الوصول', w: 14, v: dt('arrivalDate'), a: 'center', fmt: DATEF },
    { h: 'النقل', w: 11, v: r => r.transportDone ? 'تم' : 'لا', a: 'center', status: true },
    { h: 'الأسيد', w: 11, v: r => r.acidDone ? 'تم' : 'لا', a: 'center', status: true },
    { h: 'رقم الشحن', w: 18, v: r => r.shipNo || '', a: 'center' },
    { h: 'اسم الشاحن', w: 22, v: r => r.shipperName || '', a: 'right' },
    { h: 'تم الشحن', w: 13, v: r => r.shipDone ? 'تم' : 'لا', a: 'center', status: true },
    { h: 'ملاحظات', w: 34, v: r => r.notes || '', a: 'right' },
    { h: 'سجّلها', w: 14, v: r => nameOf(r.createdBy), a: 'center' },
    { h: 'تاريخ التسجيل', w: 14, v: r => dstr(r.createdAt), a: 'center' }
  ];
  const NC = cols.length, HEAD = 4, first = HEAD + 1;
  const wb = new ExcelJS.Workbook(); wb.creator = 'سجل العربيات'; wb.created = new Date();
  const ws = wb.addWorksheet('المركبات', { views: [{ rightToLeft: true, state: 'frozen', xSplit: 2, ySplit: HEAD, showGridLines: false }], properties: { tabColor: { argb: 'FF1F4F60' } }, pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  cols.forEach((c, i) => { ws.getColumn(i + 1).width = c.w; });

  ws.mergeCells(1, 1, 1, NC); const t = ws.getCell(1, 1); t.value = 'سجل العربيات';
  t.font = { name: FONT, size: 20, bold: true, color: { argb: 'FFFFFFFF' } }; t.fill = fill('1F4F60'); t.alignment = al('center'); ws.getRow(1).height = 38;
  const nCar = list.filter(r => vtypeOf(r) === 'car').length, nRf = list.length - nCar;
  const nT = list.filter(r => r.transportDone).length, nA = list.filter(r => r.acidDone).length, nS = list.filter(r => r.shipDone).length;
  ws.mergeCells(2, 1, 2, NC); const s = ws.getCell(2, 1);
  s.value = 'تاريخ التصدير: ' + dstr(Date.now()) + '   •   عدد المركبات: ' + list.length + ' (عربيات ' + nCar + ' / برادات ' + nRf + ')   •   تم النقل: ' + nT + '   •   تم الأسيد: ' + nA + '   •   تم الشحن: ' + nS;
  s.font = { name: FONT, size: 11, color: { argb: 'FF3E4A52' } }; s.fill = fill('E6EEF0'); s.alignment = al('center'); ws.getRow(2).height = 26; ws.getRow(3).height = 8;

  const hr = ws.getRow(HEAD); hr.height = 30;
  cols.forEach((c, i) => { const cell = hr.getCell(i + 1); cell.value = c.h; cell.font = { name: FONT, size: 12, bold: true, color: { argb: 'FFFFFFFF' } }; cell.fill = fill('2B6A80'); cell.border = border; cell.alignment = al('center'); });

  let sumCost = 0, sumProfit = 0;
  list.forEach((r, i) => {
    const row = ws.getRow(first + i), zebra = (i % 2 === 1) ? 'F5F6F4' : 'FFFFFF', c = calcRow(r);
    if (c.any) sumCost += c.cost; if (c.profit != null) sumProfit += c.profit;
    cols.forEach((col, k) => {
      const cell = row.getCell(k + 1); cell.value = col.v(r, i, c);
      cell.font = { name: col.mono ? 'Courier New' : FONT, size: 11, bold: !!col.mono, color: { argb: 'FF1B2328' } };
      cell.fill = fill(zebra); cell.border = border; cell.alignment = al(col.a);
      if (col.fmt) cell.numFmt = col.fmt; if (col.text) cell.numFmt = '@';
      if (col.status) { const ok = cell.value === 'تم'; cell.fill = fill(ok ? 'DDF0E5' : 'F1F2F0'); cell.font = { name: FONT, size: 11, bold: ok, color: { argb: ok ? 'FF1D7449' : 'FF8A949B' } }; }
      if (col.key === 'profit' && c.profit != null) cell.font = { name: FONT, size: 11, bold: true, color: { argb: c.profit >= 0 ? 'FF1D7449' : 'FFA63D2F' } };
    });
    row.height = Math.max(22, Math.min(60, Math.ceil(Math.max((r.description || '').length, (r.notes || '').length) / 32) * 16 + 8));
  });
  const last = first + list.length - 1, tr = last + 1;
  if (list.length) {
    const ci = cols.findIndex(c => c.key === 'cost') + 1, pi = cols.findIndex(c => c.key === 'profit') + 1;
    const col = n => ws.getColumn(n).letter;
    const lab = ws.getCell(tr, ci - 1); lab.value = 'الإجمالي (€)';
    const tc = ws.getCell(tr, ci); tc.value = { formula: 'SUM(' + col(ci) + first + ':' + col(ci) + last + ')', result: sumCost };
    const tp = ws.getCell(tr, pi); tp.value = { formula: 'SUM(' + col(pi) + first + ':' + col(pi) + last + ')', result: sumProfit };
    tc.numFmt = EUR; tp.numFmt = EUR;
    [lab, tc, tp].forEach(x => { x.font = { name: FONT, size: 12, bold: true, color: { argb: 'FF3D2F05' } }; x.fill = fill('FBEBC0'); x.border = border; x.alignment = al('center'); });
    ws.getRow(tr).height = 26;
    ws.autoFilter = { from: { row: HEAD, column: 1 }, to: { row: last, column: NC } };
  }
  const buf = await wb.xlsx.writeBuffer();
  return { buf, sumCost, sumProfit, nT, nA, nS, nCar, nRf };
}


async function load() {
  if (MOCK) {
    return [
      { chassis: 'AAA111', vehicleType: 'car', make: 'مان', description: 'تراكتور 2019', color: 'أبيض', buyer: 'أحمد', sellerName: 'Muller GmbH', sellerPhone: '+49 170 1234567', buyPrice: 10000, extraPrice: 500, shipPrice: 1000, salePrice: 13000, purchaseDate: '2026-10-01', saleDate: '2026-11-02', transportDate: '2026-10-10', shipDate: '2026-10-20', arrivalDate: '2026-11-01', transportDone: true, acidDone: false, shipNo: 'S-1', shipDone: false, notes: '', createdBy: 'zizo', createdAt: Date.now() - 86400000 },
      { chassis: 'BBB222', vehicleType: 'reefer', make: 'كرونا (Krone)', description: '', color: 'أزرق', buyer: '', sellerName: '', sellerPhone: '', buyPrice: 2000, extraPrice: null, shipPrice: 300, salePrice: null, transportDone: false, acidDone: true, shipNo: '', shipDone: true, notes: 'ملاحظة', createdBy: 'ussef', createdAt: Date.now() }
    ];
  }
  const { default: admin } = await import('firebase-admin');
  admin.initializeApp({ credential: admin.credential.cert(JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT)) });
  const snap = await admin.firestore().collection('trucks').get();
  return snap.docs.map(d => ({ id: d.id, ...d.data() }));
}

const rows = await load();
const list = rows.slice().sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));
const out = await buildWorkbookBuffer(ExcelJS, list, nameOf);
const buf = Buffer.from(out.buf), today = dstr(Date.now());
const filename = `سجل-العربيات-${today}.xlsx`;

if (MOCK) { writeFileSync('mock-report.xlsx', buf); console.log('mock ok', list.length, 'rows,', buf.length, 'bytes, cost', out.sumCost, 'profit', out.sumProfit); process.exit(0); }

const to = (process.env.MAIL_TO || '').split(',').map(x => x.trim()).filter(Boolean);
if (!to.length) throw new Error('MAIL_TO فاضي');
const mailer = nodemailer.createTransport({ service: 'gmail', auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS } });
await mailer.sendMail({
  from: `"سجل العربيات" <${process.env.SMTP_USER}>`, to,
  subject: `سجل العربيات - ${today}`,
  text: `مرفق آخر نسخة من سجل العربيات والبرادات.\n\nإجمالي المركبات: ${list.length} (عربيات ${out.nCar} • برادات ${out.nRf})\nتم النقل: ${out.nT}\nتم الأسيد: ${out.nA}\nتم الشحن: ${out.nS}\nإجمالي الربح: ${nf(out.sumProfit)} €`,
  attachments: [{ filename, content: buf, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }]
});
console.log('sent to', to.length, 'recipients,', list.length, 'rows');
