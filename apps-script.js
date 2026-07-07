// ★ VERSION 2026-07-07f（日均餐數改除以「有回報天數」：新增 monthDates 不重複回報日 + baseDays 開帳結轉天數；分母＝baseDays＋Set(monthDates∪今日)；f修正D欄Date型別解析）
/**
 * 業績回報-鴻海 累積 backend (hh-v2 — 建全 2026-07-04 16:09 規格：
 * 餐券改 黃券/藍券 各自張數+金額；當月/對帳單雙軌加總含 營業額/餐券/進貨)
 *
 * Tab 命名：民國年月 (例 `115_07`)
 *
 * Endpoints:
 *   GET ?ym=115_07                    → { ok, ym, monthTotal, monthHours, monthPurchase,
 *                                         monthVouchers:{yCnt,yAmt,bCnt,bAmt} }
 *   GET ?action=list&ym=&userId=      → { ok, items: [...] }
 *   GET ?action=statement&date=ISO    → { ok, label, amt, purchase, yCnt, yAmt, bCnt, bAmt }
 *                                        對帳窗 = 21號→下月20號
 *   POST { action:'submit', date, iso, ym, site, reporter, userId,
 *          dayAmt, dayHours, dayPurchase,
 *          vouchers:{ yellow:{count,amt}, blue:{count,amt} },
 *          purchases:{veg,egg,central}, meals, persons }
 *   POST { action:'delete', ym, id }
 *
 * 欄位 (A-Q):
 *   A 時間戳 | B record_id | C ROC日期 | D ISO日期 | E 場域 | F 填寫人 | G userId
 *   H 全日營業額 | I 今日總工時 | J 全日進貨
 *   K 黃券張 | L 黃券金額 | M 藍券張 | N 藍券金額
 *   O meals_json | P persons_json | Q purchases_json
 */

const HEADERS = ['時間戳', 'record_id', 'ROC日期', 'ISO日期', '場域', '填寫人', 'userId',
                 '全日營業額', '今日總工時', '全日進貨',
                 '黃券張', '黃券金額', '藍券張', '藍券金額',
                 'meals_json', 'persons_json', 'purchases_json'];

function ensureTab(ym) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ym);
  if (!sheet) {
    sheet = ss.insertSheet(ym);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    const lastCol = sheet.getLastColumn();
    if (lastCol < HEADERS.length) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
  }
  // ISO日期(D欄)強制純文字：Sheets 會把 "2026-07-03" 自動轉成 Date 物件，
  // 造成 statementTotals 抓不到、對帳單恆為 0。設 @ 讓 appendRow 存字串。
  sheet.getRange(2, 4, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
  return sheet;
}

// 把 D 欄的值轉成 yyyy-MM-dd。⚠️ Apps Script 的 getValues() 有時回 Date 物件、有時回字串，
// 且 `instanceof Date` 在 Apps Script 不可靠（已實測踩過），改用 duck-typing 判斷。
function isoFromCell_(v) {
  if (v == null || v === '') return '';
  if (typeof v.getFullYear === 'function') {              // Date-like
    return Utilities.formatDate(v, 'GMT+8', 'yyyy-MM-dd');
  }
  var s = String(v);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);            // 已是 ISO 字串
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var d = new Date(s);                                    // 保險：其他可解析格式
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'GMT+8', 'yyyy-MM-dd');
  return s.slice(0, 10);
}

function emptyTotals_() {
  return { monthTotal: 0, monthHours: 0, monthPurchase: 0, monthGuests: 0,
           monthDates: [],
           monthVouchers: { yCnt: 0, yAmt: 0, bCnt: 0, bAmt: 0 } };
}

function monthTotalsFor(ym) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ym);
  if (!sheet) return emptyTotals_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return emptyTotals_();
  const vals = sheet.getRange(2, 8, lastRow - 1, 8).getValues();   // H..O (O=meals_json 用來算餐數)
  // 日均餐數用「有回報天數」當分母：收集不重複的回報日期（排除 SYSTEM 期初結轉列）
  const meta = sheet.getRange(2, 4, lastRow - 1, 4).getValues();   // D..G: D ISO日期, E 場域, F 填寫人, G userId
  const dateSet = {};
  meta.forEach(function(m) {
    if (String(m[3]) === 'SYSTEM' || String(m[2]) === '期初結轉') return;   // 結轉列不算一個回報日
    var iso = isoFromCell_(m[0]);
    if (iso) dateSet[iso] = 1;
  });
  const t = emptyTotals_();
  t.monthDates = Object.keys(dateSet);
  vals.forEach(function(r) {
    t.monthTotal    += Number(r[0]) || 0;   // H
    t.monthHours    += Number(r[1]) || 0;   // I
    t.monthPurchase += Number(r[2]) || 0;   // J
    t.monthVouchers.yCnt += Number(r[3]) || 0;   // K
    t.monthVouchers.yAmt += Number(r[4]) || 0;   // L
    t.monthVouchers.bCnt += Number(r[5]) || 0;   // M
    t.monthVouchers.bAmt += Number(r[6]) || 0;   // N
    // 餐數 = 該日各餐別 count 加總（存在 meals_json / O 欄）
    try {
      var meals = r[7] ? JSON.parse(r[7]) : [];
      if (meals && meals.length) meals.forEach(function(m){ t.monthGuests += Number(m.count) || 0; });
    } catch (_) {}
  });
  return t;
}

function listForUser(ym, userId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ym);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const items = [];
  rows.forEach(function(r) {
    if (userId && String(r[6]) !== String(userId)) return;
    var meals = [], persons = [], purchases = {};
    try { if (r[14]) meals = JSON.parse(r[14]); } catch (_) {}
    try { if (r[15]) persons = JSON.parse(r[15]); } catch (_) {}
    try { if (r[16]) purchases = JSON.parse(r[16]); } catch (_) {}
    items.push({
      id: r[1], date: r[2], iso: r[3], site: r[4], reporter: r[5], userId: r[6],
      dayAmt: Number(r[7]) || 0, dayHours: Number(r[8]) || 0, dayPurchase: Number(r[9]) || 0,
      vouchers: {
        yellow: { count: Number(r[10]) || 0, amt: Number(r[11]) || 0 },
        blue:   { count: Number(r[12]) || 0, amt: Number(r[13]) || 0 }
      },
      meals: meals, persons: persons, purchases: purchases
    });
  });
  return items;
}

function deleteRow(ym, id) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ym);
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function pad2_(n) { return ('0' + n).slice(-2); }

// ===== 修改數值（人工調整）=====
// 建全 2026-07-04：可直接修正「當月」或「對帳單」的 營業額/黃券/藍券/進貨/工時。
// 存進獨立「調整」分頁（不動每日回報列），以 delta 累加方式套到累計值上。
// 調整分頁欄位：A 時間戳 | B scope(month|statement) | C key(month=ym；statement=窗label) | D field | E value(delta；券為張數 delta)
const ADJ_TAB = '調整';
const VPRICE = { yellow: 70, blue: 110 };

function ensureAdjTab_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ADJ_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(ADJ_TAB);
    sheet.getRange(1, 1, 1, 5).setValues([['時間戳', 'scope', 'key', 'field', 'value']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 回傳某 scope+key 下所有調整聚合值
function adjMap_(scope, key) {
  const out = { amt: 0, hours: 0, purchase: 0, guests: 0, days: 0, yCnt: 0, yAmt: 0, bCnt: 0, bAmt: 0 };
  const sheet = SpreadsheetApp.getActive().getSheetByName(ADJ_TAB);
  if (!sheet) return out;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return out;
  const rows = sheet.getRange(2, 2, lastRow - 1, 4).getValues();   // B..E
  rows.forEach(function(r) {
    if (String(r[0]) !== String(scope) || String(r[1]) !== String(key)) return;
    const field = String(r[2]);
    const v = Number(r[3]) || 0;
    if (field === 'amt') out.amt += v;
    else if (field === 'hours') out.hours += v;
    else if (field === 'purchase') out.purchase += v;
    else if (field === 'guests') out.guests += v;
    else if (field === 'days') out.days += v;   // 開帳結轉涵蓋的回報天數（日均餐數分母基底）
    else if (field === 'yellow') { out.yCnt += v; out.yAmt += v * VPRICE.yellow; }
    else if (field === 'blue') { out.bCnt += v; out.bAmt += v * VPRICE.blue; }
  });
  return out;
}

// upsert：同 scope+key+field 只留一筆（覆蓋）
function setAdjust_(scope, key, field, delta) {
  const sheet = ensureAdjTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 2, lastRow - 1, 3).getValues();   // B(scope) C(key) D(field)
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(scope) && String(rows[i][1]) === String(key) && String(rows[i][2]) === String(field)) {
        sheet.getRange(i + 2, 5).setValue(delta);   // E
        return;
      }
    }
  }
  sheet.appendRow([new Date().toISOString(), scope, key, field, delta]);
}

// 依 scope+field 取「純每日回報列」現值（不含調整），用來算 delta
function rowsValueFor_(scope, key, field, refDate) {
  if (scope === 'month') {
    const t = monthTotalsFor(key);
    if (field === 'amt') return t.monthTotal;
    if (field === 'hours') return t.monthHours;
    if (field === 'purchase') return t.monthPurchase;
    if (field === 'guests') return t.monthGuests;
    if (field === 'yellow') return t.monthVouchers.yCnt;
    if (field === 'blue') return t.monthVouchers.bCnt;
  } else {
    const s = statementTotals(refDate);
    if (field === 'amt') return s.amt;
    if (field === 'purchase') return s.purchase;
    if (field === 'yellow') return s.yCnt;
    if (field === 'blue') return s.bCnt;
  }
  return 0;
}

// 對帳窗：date 落在 [X月21號, X+1月20號] → 回窗起訖與涉及的 ym tabs
function statementWindow(isoDate) {
  const d = new Date(isoDate + 'T00:00:00');
  var startY = d.getFullYear(), startM = d.getMonth();   // 窗起點（西元）
  if (d.getDate() < 21) {
    startM -= 1;
    if (startM < 0) { startM = 11; startY -= 1; }
  }
  const start = new Date(startY, startM, 21);
  const end = new Date(startY, startM + 1, 20);   // Date 自動跨年
  function ymOf(dt) { return (dt.getFullYear() - 1911) + '_' + pad2_(dt.getMonth() + 1); }
  const label = (start.getFullYear() - 1911) + '/' + (start.getMonth() + 1) + '/21〜' +
                (end.getFullYear() - 1911) + '/' + (end.getMonth() + 1) + '/20';
  return { start: start, end: end, yms: [ymOf(start), ymOf(end)], label: label };
}

function statementTotals(isoDate) {
  const w = statementWindow(isoDate);
  var amt = 0, purchase = 0, yCnt = 0, yAmt = 0, bCnt = 0, bAmt = 0;
  const ss = SpreadsheetApp.getActive();
  w.yms.forEach(function(ym) {
    const sheet = ss.getSheetByName(ym);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const rows = sheet.getRange(2, 4, lastRow - 1, 11).getValues();   // D(iso)..N(藍券金額)
    rows.forEach(function(r) {
      // r[0] = ISO日期(D 欄)。Sheets 常把 "2026-07-03" 這種字串自動轉成 Date 物件，
      // 也可能存成純字串，兩種都要能解析（否則對帳單一律算不到、回 0）。
      var rd;
      if (r[0] instanceof Date) {
        rd = new Date(r[0].getFullYear(), r[0].getMonth(), r[0].getDate());
      } else {
        const iso = String(r[0] || '').slice(0, 10);
        if (!iso) return;
        rd = new Date(iso + 'T00:00:00');
      }
      if (!rd || isNaN(rd.getTime()) || rd < w.start || rd > w.end) return;
      amt      += Number(r[4])  || 0;   // H 全日營業額（自 D 起 offset 4）
      purchase += Number(r[6])  || 0;   // J 全日進貨
      yCnt     += Number(r[7])  || 0;   // K 黃券張
      yAmt     += Number(r[8])  || 0;   // L 黃券金額
      bCnt     += Number(r[9])  || 0;   // M 藍券張
      bAmt     += Number(r[10]) || 0;   // N 藍券金額
    });
  });
  return { label: w.label, amt: amt, purchase: purchase,
           yCnt: yCnt, yAmt: yAmt, bCnt: bCnt, bAmt: bAmt };
}

function newId_() { return Utilities.getUuid().slice(0, 8); }

function totalsPayload_(ym, extra) {
  const t = monthTotalsFor(ym);
  const a = adjMap_('month', ym);
  const out = extra || {};
  out.monthTotal = t.monthTotal + a.amt;
  out.monthHours = t.monthHours + a.hours;
  out.monthPurchase = t.monthPurchase + a.purchase;
  out.monthGuests = t.monthGuests + a.guests;
  out.monthDates = t.monthDates;   // 真實回報的不重複日期（前端 Set 併今日算「有回報天數」）
  out.baseDays = a.days;           // 開帳結轉涵蓋的天數（如7月結轉=7/1~7/3=3天）
  out.monthVouchers = {
    yCnt: t.monthVouchers.yCnt + a.yCnt, yAmt: t.monthVouchers.yAmt + a.yAmt,
    bCnt: t.monthVouchers.bCnt + a.bCnt, bAmt: t.monthVouchers.bAmt + a.bAmt
  };
  return out;
}

function doGet(e) {
  try {
    const action = (e.parameter.action || 'monthTotal');
    if (action === 'statement') {
      const dateIso = e.parameter.date;
      if (!dateIso) return _json({ ok: false, error: 'missing date' });
      const s = statementTotals(dateIso);
      const a = adjMap_('statement', s.label);
      return _json({ ok: true, label: s.label,
                     amt: s.amt + a.amt, purchase: s.purchase + a.purchase,
                     yCnt: s.yCnt + a.yCnt, yAmt: s.yAmt + a.yAmt,
                     bCnt: s.bCnt + a.bCnt, bAmt: s.bAmt + a.bAmt });
    }
    const ym = e.parameter.ym;
    if (!ym) return _json({ ok: false, error: 'missing ym' });
    if (action === 'list') {
      const userId = e.parameter.userId || '';
      return _json(totalsPayload_(ym, { ok: true, ym: ym, items: listForUser(ym, userId) }));
    }
    return _json(totalsPayload_(ym, { ok: true, ym: ym }));
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'submit';
    if (action === 'delete') {
      const ym = data.ym, id = data.id;
      if (!ym || !id) return _json({ ok: false, error: 'missing ym/id' });
      const ok = deleteRow(ym, id);
      return _json(totalsPayload_(ym, { ok: ok }));
    }
    if (action === 'adjust') {
      // { scope:'month'|'statement', field:'amt'|'hours'|'purchase'|'yellow'|'blue', target, ym, date }
      const scope = String(data.scope || '');
      const field = String(data.field || '');
      const target = Number(data.target);
      if ((scope !== 'month' && scope !== 'statement') || !field || isNaN(target)) {
        return _json({ ok: false, error: 'bad adjust params' });
      }
      if (scope === 'statement' && (field === 'hours' || field === 'guests')) {
        return _json({ ok: false, error: '對帳單不記工時/餐數' });
      }
      const ym = data.ym;
      const refDate = data.date;   // ISO，用於算對帳窗與當月ym
      const key = (scope === 'month') ? ym : statementWindow(refDate).label;
      if (!key) return _json({ ok: false, error: 'missing key' });
      const rowsVal = rowsValueFor_(scope, key, field, refDate);
      const delta = target - rowsVal;   // 券為張數 delta
      setAdjust_(scope, key, field, delta);
      // 回傳更新後累計（含當月與對帳單）
      const out = totalsPayload_(ym, { ok: true });
      if (refDate) {
        const s2 = statementTotals(refDate);
        const a2 = adjMap_('statement', s2.label);
        out.statement = { label: s2.label,
          amt: s2.amt + a2.amt, purchase: s2.purchase + a2.purchase,
          yCnt: s2.yCnt + a2.yCnt, yAmt: s2.yAmt + a2.yAmt,
          bCnt: s2.bCnt + a2.bCnt, bAmt: s2.bAmt + a2.bAmt };
      }
      return _json(out);
    }
    const ym = data.ym;
    if (!ym) return _json({ ok: false, error: 'missing ym' });
    const sheet = ensureTab(ym);
    const id = newId_();
    const v = data.vouchers || {};
    const vy = v.yellow || {}, vb = v.blue || {};
    sheet.appendRow([
      new Date().toISOString(),
      id,
      data.date || '',
      data.iso || '',
      data.site || '',
      data.reporter || '',
      data.userId || '',
      Number(data.dayAmt || 0),
      Number(data.dayHours || 0),
      Number(data.dayPurchase || 0),
      Number(vy.count || 0),
      Number(vy.amt || 0),
      Number(vb.count || 0),
      Number(vb.amt || 0),
      JSON.stringify(data.meals || []),
      JSON.stringify(data.persons || []),
      JSON.stringify(data.purchases || {})
    ]);
    return _json(totalsPayload_(ym, { ok: true, id: id }));
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
