/**
 * 羽球點名板的共用資料庫（Google Apps Script）
 *
 * 資料存在這份試算表的「資料」分頁 A1，每一次修改都會記在「紀錄」分頁。
 * 設定步驟請看 badminton-setup.md。
 */

var EMPTY = { members: [], pool: [], leaves: {}, invites: {}, offDays: [] };

function sheet_(name) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function read_() {
  var v = sheet_('資料').getRange('A1').getValue();
  if (!v) return JSON.parse(JSON.stringify(EMPTY));
  var s = JSON.parse(v);
  Object.keys(EMPTY).forEach(function (k) { if (!s[k]) s[k] = JSON.parse(JSON.stringify(EMPTY[k])); });
  return s;
}

function write_(s) {
  sheet_('資料').getRange('A1').setValue(JSON.stringify(s));
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function addTo_(list, v) { if (list.indexOf(v) === -1) list.push(v); }
function removeFrom_(list, v) { var i = list.indexOf(v); if (i !== -1) list.splice(i, 1); }

// 一次修改一件事，這樣兩個人同時改不同的東西也不會互相蓋掉
function apply_(s, op) {
  var name = String(op.name || '').trim().slice(0, 40);
  var date = /^\d{4}-\d{2}-\d{2}$/.test(op.date) ? op.date : null;
  switch (op.type) {
    case 'member': if (name) (op.on ? addTo_ : removeFrom_)(s.members, name); return name + (op.on ? ' 加入固定成員' : ' 移出固定成員');
    case 'pool': if (name) (op.on ? addTo_ : removeFrom_)(s.pool, name); return name + (op.on ? ' 加入可以約的人' : ' 移出可以約的人');
    case 'leave':
      if (!name || !date) return null;
      s.leaves[date] = s.leaves[date] || [];
      (op.on ? addTo_ : removeFrom_)(s.leaves[date], name);
      if (!s.leaves[date].length) delete s.leaves[date];
      return date + ' ' + name + (op.on ? ' 請假' : ' 取消請假');
    case 'sub':
      if (!name || !date) return null;
      s.invites[date] = s.invites[date] || {};
      if (op.on) s.invites[date][name] = 'yes'; else delete s.invites[date][name];
      if (!Object.keys(s.invites[date]).length) delete s.invites[date];
      return date + ' ' + name + (op.on ? ' 加入候補' : ' 取消候補');
    case 'off':
      if (!date) return null;
      (op.on ? addTo_ : removeFrom_)(s.offDays, date);
      return date + (op.on ? ' 設為全員請假' : ' 取消全員請假');
  }
  return null;
}

function doGet() {
  return out_({ ok: true, state: read_() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    var s = read_();
    var log = sheet_('紀錄');
    (body.ops || []).slice(0, 50).forEach(function (op) {
      var what = apply_(s, op);
      if (what) log.appendRow([new Date(), what]);
    });
    write_(s);
    return out_({ ok: true, state: s });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
