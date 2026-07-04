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
  return sheet;
}

function emptyTotals_() {
  return { monthTotal: 0, monthHours: 0, monthPurchase: 0,
           monthVouchers: { yCnt: 0, yAmt: 0, bCnt: 0, bAmt: 0 } };
}

function monthTotalsFor(ym) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ym);
  if (!sheet) return emptyTotals_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return emptyTotals_();
  const vals = sheet.getRange(2, 8, lastRow - 1, 7).getValues();   // H..N
  const t = emptyTotals_();
  vals.forEach(function(r) {
    t.monthTotal    += Number(r[0]) || 0;   // H
    t.monthHours    += Number(r[1]) || 0;   // I
    t.monthPurchase += Number(r[2]) || 0;   // J
    t.monthVouchers.yCnt += Number(r[3]) || 0;   // K
    t.monthVouchers.yAmt += Number(r[4]) || 0;   // L
    t.monthVouchers.bCnt += Number(r[5]) || 0;   // M
    t.monthVouchers.bAmt += Number(r[6]) || 0;   // N
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
  const out = extra || {};
  out.monthTotal = t.monthTotal;
  out.monthHours = t.monthHours;
  out.monthPurchase = t.monthPurchase;
  out.monthVouchers = t.monthVouchers;
  return out;
}

function doGet(e) {
  try {
    const action = (e.parameter.action || 'monthTotal');
    if (action === 'statement') {
      const dateIso = e.parameter.date;
      if (!dateIso) return _json({ ok: false, error: 'missing date' });
      const s = statementTotals(dateIso);
      return _json({ ok: true, label: s.label, amt: s.amt, purchase: s.purchase,
                     yCnt: s.yCnt, yAmt: s.yAmt, bCnt: s.bCnt, bAmt: s.bAmt });
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
