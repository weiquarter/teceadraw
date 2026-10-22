/** 前台 API：stageLoad、stageAsset、recordDraw、heartbeat，以及作用中前台代號的判斷。前台的所有動作以操作紀錄（_操作紀錄）為準，得獎名單等由它推導寫入。 */

// _前台狀態 固定列位置（A 欄項目、B 欄值）
const ST = {
  ACTIVE: 1, CODE: 2, PRIZE: 3, PAGE: 4, PENDING: 5, REPORTED: 6, FORCE_HOME: 7, REJECTED: 8, REJECTED_AT: 9,
};
const ST_KEYS = ['作用中前台代號', '前台代號', '目前獎項', '目前頁面', '待同步筆數', '最後回報時間', '指令_回主頁', '被拒絕的補寫', '最後被拒絕時間'];
const OPLOG_HEADERS = ['操作 ID', '內容', '時間', '前台代號', '結果'];
const BACKGROUND_KEYS = { home: '主頁背景', prize: '獎項頁背景', draw: '抽獎頁背景', logo: 'logo' };

function stageAuth_(p) {
  const key = PropertiesService.getScriptProperties().getProperty('STAGE_KEY');
  return !!key && String(p.k || '') === key;
}

function stageSheet_() {
  const sh = sheet_(SHEET.STAGE);
  if (sh.getRange(1, 1).getValue() !== ST_KEYS[0]) {
    sh.clearContents();
    sh.getRange(1, 1, ST_KEYS.length, 1).setValues(ST_KEYS.map(k => [k]));
  }
  return sh;
}

function stageValues_() {
  return stageSheet_().getRange(1, 2, ST_KEYS.length, 1).getValues().map(r => r[0]);
}

// 沒有作用中前台時，第一個呼叫的自動成為作用中
function claimActive_(code) {
  const sh = stageSheet_();
  const cur = String(sh.getRange(ST.ACTIVE, 2).getValue());
  if (cur) return cur === code;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const again = String(sh.getRange(ST.ACTIVE, 2).getValue());
    if (again) return again === code;
    sh.getRange(ST.ACTIVE, 2).setValue(code);
    log_('前台啟用', true, null, '', code);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function stageCode_(p) {
  const code = String(p.code || '');
  return /^\d{4}$/.test(code) ? code : '';
}

function apiHeartbeat_(p) {
  if (!stageAuth_(p)) return { ok: false, code: 'badKey', message: '前台連結已失效，請向後台索取新連結' };
  const code = stageCode_(p);
  if (!code) return { ok: false, code: 'badCode' };
  const state = currentState_();
  const active = claimActive_(code);
  const out = { ok: true, state: state, active: active, apiVersion: API_VERSION };
  if (!active) return out;

  const sh = stageSheet_();
  sh.getRange(ST.CODE, 2, 5, 1).setValues([[code], [String(p.prize || '')], [String(p.page || '')], [Number(p.pending) || 0], [new Date()]]);
  const vals = stageValues_();
  out.forceHome = vals[ST.FORCE_HOME - 1] ? new Date(vals[ST.FORCE_HOME - 1]).getTime() : 0;
  out.nextPrize = String(sheet_(SHEET.CONSOLE).getRange('B17').getValue() || '');
  if (state === STATE.OPEN) out.checkins = Number(sheet_(SHEET.CONSOLE).getRange('B7').getValue()) || 0;
  out.version = Number(sysGet_('名單版本') || 0);
  out.prizeChanged = sysGet_('獎項鎖定後變更') === '是';
  return out;
}

function apiStageLoad_(p) {
  if (!stageAuth_(p)) return { ok: false, code: 'badKey', message: '前台連結已失效，請向後台索取新連結' };
  const code = stageCode_(p);
  if (!code) return { ok: false, code: 'badCode' };
  const state = currentState_();
  // 鎖定前只給背景與 Logo，報到期間的主頁也要有主視覺
  if ([STATE.LOCKED, STATE.DRAWING].indexOf(state) < 0) return { ok: true, state: state, ready: false, assets: stageAssets_([]) };
  const active = claimActive_(code);

  const sys = readKeyValue_(SHEET.SYSTEM);
  const prizes = sortedPrizes_();
  const snap = sheet_(SHEET.SNAPSHOT);
  const last = snap.getLastRow();
  const pool = last >= 2 ? snap.getRange(2, 1, last - 1, SNAPSHOT_HEADERS.length).getValues()
    .map(r => ({ id: String(r[0]), unit: String(r[1]), name: String(r[2]), display: String(r[3]) })) : [];

  return {
    ok: true, state: state, ready: true, active: active,
    version: Number(sys['名單版本'] || 0),
    poolHash: String(sys['名單檢查碼'] || ''),
    prizeHash: String(sys['獎項檢查碼'] || ''),
    prizeChanged: sys['獎項鎖定後變更'] === '是' || prizeHash_(prizes) !== String(sys['獎項檢查碼'] || ''),
    prizes: prizes.map(x => ({ sort: x.sort, name: x.name, qty: x.qty, mode: x.mode, photo: photoKey_(x.photo) })),
    pool: pool,
    ops: appliedOps_().map(o => o.op),
    assets: stageAssets_(prizes),
  };
}

// 素材資料夾裡用得到的檔案：三張背景、Logo、各獎項照片
function stageAssets_(prizes) {
  const folderId = sysGet_('素材資料夾ID');
  if (!folderId) return [];
  const wanted = {};
  Object.keys(BACKGROUND_KEYS).forEach(k => { wanted[photoKey_(BACKGROUND_KEYS[k])] = k; });
  prizes.forEach(x => { if (x.photo) wanted[photoKey_(x.photo)] = 'photo:' + photoKey_(x.photo); });
  const out = [];
  const it = DriveApp.getFolderById(folderId).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const role = wanted[photoKey_(f.getName())];
    if (role && f.getMimeType().indexOf('image/') === 0) out.push({ role: role, id: f.getId(), size: f.getSize() });
  }
  CacheService.getScriptCache().put('assets', JSON.stringify(out.map(a => a.id)), 21600);
  return out;
}

function apiStageAsset_(p) {
  if (!stageAuth_(p)) return { ok: false, code: 'badKey' };
  const allowed = JSON.parse(CacheService.getScriptCache().get('assets') || '[]');
  if (allowed.indexOf(String(p.id)) < 0) return { ok: false, code: 'badAsset', message: '不是素材資料夾裡的檔案' };
  const blob = DriveApp.getFileById(String(p.id)).getBlob();
  return { ok: true, mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}

// 一次寫入多筆操作；同一個操作 ID 只套用一次
function apiRecordDraw_(p) {
  if (!stageAuth_(p)) return { ok: false, code: 'badKey', message: '前台連結已失效' };
  const code = stageCode_(p);
  const ops = Array.isArray(p.ops) ? p.ops : [];
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const state = currentState_();
    if ([STATE.LOCKED, STATE.DRAWING].indexOf(state) < 0) return { ok: false, code: 'badState', message: '目前狀態是「' + state + '」，不能寫入抽獎結果' };
    const sh = stageSheet_();
    if (String(sh.getRange(ST.ACTIVE, 2).getValue()) !== code) {
      rejectOps_(sh, code, ops, '不是作用中的前台');
      return { ok: false, code: 'notActive', message: '這台前台已被取代，結果沒有寫入' };
    }
    const sys = readKeyValue_(SHEET.SYSTEM);
    const prizes = sortedPrizes_();
    if (Number(p.v) !== Number(sys['名單版本'])) return { ok: false, code: 'versionChanged', message: '名單版本已變更，請重新載入前台' };
    if (sys['獎項鎖定後變更'] === '是' || prizeHash_(prizes) !== String(sys['獎項檢查碼'])) {
      return { ok: false, code: 'prizeChanged', message: '獎項在鎖定後被修改，請後台解除鎖定後重新鎖定' };
    }

    const known = {};
    appliedOps_().forEach(o => { known[o.op.id] = true; });
    const fresh = ops.filter(o => o && o.id && !known[o.id]);
    if (fresh.length) applyOps_(fresh, prizes, code, state);
    if (fresh.some(o => o.t === 'draw')) log_('recordDraw', true, Date.now() - t0, fresh.length + ' 筆', code);
    return { ok: true, acked: ops.map(o => o && o.id).filter(Boolean) };
  } catch (err) {
    log_('recordDraw', false, Date.now() - t0, '', code, err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function rejectOps_(sh, code, ops, reason) {
  const now = new Date();
  const oplog = oplogSheet_();
  ops.forEach(o => oplog.appendRow([String(o.id || ''), JSON.stringify(o), now, code, '拒絕：' + reason]));
  const n = Number(sh.getRange(ST.REJECTED, 2).getValue()) || 0;
  sh.getRange(ST.REJECTED, 2, 2, 1).setValues([[n + ops.length], [now]]);
  log_('recordDraw', false, null, ops.length + ' 筆', code, '舊前台補寫被拒絕：' + reason);
}

function oplogSheet_() {
  const sh = sheet_(SHEET.OPLOG);
  if (sh.getRange(1, 1).getValue() !== OPLOG_HEADERS[0]) {
    sh.clearContents();
    sh.getRange(1, 1, 1, OPLOG_HEADERS.length).setValues([OPLOG_HEADERS]);
  }
  return sh;
}

function appliedOps_() {
  const sh = oplogSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, OPLOG_HEADERS.length).getValues()
    .filter(r => r[4] === '套用')
    .map(r => ({ op: JSON.parse(r[1]) }));
}

function sortedPrizes_() {
  return readPrizes_().slice().sort((a, b) => a.sort - b.sort);
}

// 操作：draw（p 獎項序、s 名額序、w 工號）、void、unvoid、complete（p）
function applyOps_(ops, prizes, code, state) {
  const now = new Date();
  const oplog = oplogSheet_();
  oplog.getRange(oplog.getLastRow() + 1, 1, ops.length, OPLOG_HEADERS.length)
    .setValues(ops.map(o => [String(o.id), JSON.stringify(o), now, code, '套用']));

  const winners = sheet_(SHEET.WINNERS);
  const people = sheet_(SHEET.PEOPLE);
  const pLast = people.getLastRow();
  const ids = pLast >= 2 ? people.getRange(2, PC.ID, pLast - 1, 1).getDisplayValues().map(r => normId(r[0])) : [];
  const snap = {};
  const sLast = sheet_(SHEET.SNAPSHOT).getLastRow();
  if (sLast >= 2) sheet_(SHEET.SNAPSHOT).getRange(2, 1, sLast - 1, 3).getValues().forEach(r => { snap[String(r[0])] = r; });

  const setPrizeOfPerson = (id, text) => {
    const i = ids.indexOf(id);
    if (i >= 0) people.getRange(i + 2, PC.PRIZE).setValue(text);
  };
  const findWinnerRow = (prizeName, seq, id) => {
    const last = winners.getLastRow();
    if (last < 2) return -1;
    const rows = winners.getRange(2, 1, last - 1, 7).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (rows[i][1] === prizeName && Number(rows[i][2]) === seq && String(rows[i][4]) === id) return i + 2;
    }
    return -1;
  };
  const prizeSheet = sheet_(SHEET.PRIZES);
  let firstDraw = false;

  ops.forEach(o => {
    const prize = prizes[o.p];
    if (!prize) return;
    const seq = Number(o.s) + 1;
    if (o.t === 'draw') {
      const who = snap[String(o.w)];
      if (!who) { log_('recordDraw', false, null, String(o.w), code, '抽獎池裡沒有這個工號'); return; }
      winners.appendRow([new Date(Number(o.at) || now.getTime()), prize.name, seq, who[1], String(who[0]), who[2], '有效', o.r ? '補抽' : '']);
      setPrizeOfPerson(String(o.w), prize.name);
      if (prize.status !== '進行中' && prize.status !== '已完成') {
        prizeSheet.getRange(prize.row, 6).setValue('進行中');
        prize.status = '進行中';
      }
      firstDraw = true;
    } else if (o.t === 'void' || o.t === 'unvoid') {
      const row = findWinnerRow(prize.name, seq, String(o.w));
      if (row > 0) winners.getRange(row, 7).setValue(o.t === 'void' ? '作廢' : '有效');
      setPrizeOfPerson(String(o.w), o.t === 'void' ? prize.name + '（作廢）' : prize.name);
    } else if (o.t === 'complete') {
      prizeSheet.getRange(prize.row, 6).setValue('已完成');
      prize.status = '已完成';
    }
  });

  if (firstDraw && state === STATE.LOCKED) {
    sysSet_('系統狀態', STATE.DRAWING);
    invalidateConfig_();
    log_('開始抽獎', true, null, STATE.LOCKED + ' → ' + STATE.DRAWING, code);
  }
}

// ── 後台選單：強制回主頁、強制切換 ─────────────────────

function menuForceHome() {
  const ui = SpreadsheetApp.getUi();
  stageSheet_().getRange(ST.FORCE_HOME, 2).setValue(new Date());
  log_('強制前台回主頁', true, null, '');
  ui.alert('已送出', '前台下一次回報時（約 5 秒內）會回到主頁。\n前台斷線時收不到，控制台「最後回報時間」會變紅。', ui.ButtonSet.OK);
}

function menuForceSwitch() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('強制切換前台', '請輸入新前台畫面上顯示的 4 位數接手代號：', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const code = r.getResponseText().trim();
  if (!/^\d{4}$/.test(code)) { ui.alert('代號要是 4 位數字。'); return; }
  const vals = stageValues_();
  const pending = Number(vals[ST.PENDING - 1]) || 0;
  if (vals[ST.ACTIVE - 1] && String(vals[ST.ACTIVE - 1]) !== code && pending > 0) {
    const c = ui.alert('原前台還有待同步的結果',
      '原前台（' + vals[ST.ACTIVE - 1] + '）最後回報時有 ' + pending + ' 筆還沒寫進 Sheet。\n切換後這些結果會被拒絕，要靠台下看到的結果人工補登。\n\n確定要切換嗎？', ui.ButtonSet.YES_NO);
    if (c !== ui.Button.YES) return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    stageSheet_().getRange(ST.ACTIVE, 2).setValue(code);
  } finally {
    lock.releaseLock();
  }
  log_('強制切換前台', true, null, (vals[ST.ACTIVE - 1] || '（無）') + ' → ' + code, code);
  ui.alert('已切換', '代號 ' + code + ' 的前台會在幾秒內接手，並從 Sheet 恢復進度。', ui.ButtonSet.OK);
}
