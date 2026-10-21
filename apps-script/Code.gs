// מערכת גבייה – שרת Apps Script. מגיש את Index.html וקורא/כותב ל-Google Sheet.
// פריסה: Execute as = User accessing the web app; Who has access = Anyone with Google account.
// הגישה לנתונים נשלטת בהרשאות השיתוף של הגיליון עצמו (רק מי שיש לו עריכה בגיליון יכול לקרוא/לכתוב).
const SHEET_ID = '1Yvw4z8rsw5L8egJC_WaJuNNmm6ZE67rDvxeXCntvpz8';
const TAB = 'גבייה';
const NCOLS = 17; // A..Q ; Q = JSON נסתר (תשלומים, יומן פניות)
const STATUS = {hold: 'בהמתנה', quote: 'הצעת מחיר', legal: 'משפטי'};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('מערכת גבייה')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const sh_ = () => SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
const iso_ = v => v instanceof Date ? Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd') : (v || '');
const toDate_ = s => s ? new Date(s + 'T12:00:00') : '';

function getItems() {
  const sh = sh_(), last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, NCOLS).getValues();
  const rev = {}; Object.keys(STATUS).forEach(k => rev[STATUS[k]] = k);
  const out = [];
  rows.forEach((r, k) => {
    if (!r[0]) return;
    let x = {}; try { x = r[16] ? JSON.parse(r[16]) : {}; } catch (e) {}
    const paid = (x.payments || []).reduce((s, p) => s + p.amount, 0);
    const amount = Number(r[6]) || 0;
    // העמודות הגלויות גוברות על ה-JSON, כדי שעריכה ידנית בגיליון תיקלט
    const it = Object.assign({id: 'r' + (k + 2), payments: [], log: []}, x, {
      customer: String(r[0]), contact: String(r[1] || ''), phone: String(r[2] || ''), email: String(r[3] || ''),
      invoice: String(r[4] || ''), due: iso_(r[5]), amount: amount,
      status: r[10] === 'שולם' ? '' : (rev[r[10]] || ''), done: r[10] === 'שולם' || (amount > 0 && paid >= amount),
      promise: iso_(r[13]), next: iso_(r[14]), notes: String(r[15] || '')
    });
    if (r[10] === 'שולם' && amount > 0 && paid < amount) it.payments.push({amount: amount - paid, date: iso_(new Date()), method: 'גיליון'});
    out.push(it);
  });
  return out;
}

function saveItems(json) {
  const items = JSON.parse(json);
  const sh = sh_(), lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const n = items.length, last = sh.getLastRow();
    const vals = [];
    items.forEach(it => {
      const paid = (it.payments || []).reduce((s, p) => s + p.amount, 0);
      const lastContact = (it.log || []).find(l => l.t === 'contact');
      const state = it.done ? 'שולם' : (STATUS[it.status] || '');
      vals.push([it.customer, it.contact || '', it.phone || '', it.email || '', it.invoice || '', toDate_(it.due),
        it.amount > 0 ? it.amount : '', paid || '', null, null, state,
        lastContact ? toDate_(lastContact.d) : '', lastContact ? lastContact.txt : '',
        toDate_(it.promise), toDate_(it.next), it.notes || '',
        JSON.stringify({id: it.id, payments: it.payments || [], log: it.log || [], date: it.date || ''})]);
    });
    // כותבים עמודות A:H ו-K:Q; עמודות I,J (נוסחאות יתרה/פיגור) נשארות
    if (n) {
      sh.getRange(2, 1, n, 8).setValues(vals.map(v => v.slice(0, 8)));
      sh.getRange(2, 11, n, 7).setValues(vals.map(v => v.slice(10)));
      const f = sh.getRange(2, 9, 1, 2);
      if (n > 1) f.copyTo(sh.getRange(3, 9, n - 1, 2));
    }
    if (last > n + 1) sh.getRange(n + 2, 1, last - n - 1, NCOLS).clearContent();
    return {ok: true, count: n};
  } finally { lock.releaseLock(); }
}
