/** 報到：工號對照表（放 CacheService）、lookupId、checkin。checkin 先記在 Script Properties 立刻回應，再由空著的請求批次寫進 Sheet。 */

const MAP_TTL = 21600;
const MAP_CHUNK = 30000;
const PENDING = 'p:';

// 對照表內容：工號 → [遮蔽姓名, 單位, 可否摸彩, 是否移除, 列號, 報到時間 HH:mm]
const R = { NAME: 0, UNIT: 1, ELIGIBLE: 2, REMOVED: 3, ROW: 4, TIME: 5 };

function apiLookupId_(p) {
  const state = getConfig_().state;
  if (state === STATE.PREP) return { ok: true, result: 'prep' };
  if (state === STATE.ENDED) return { ok: true, result: 'ended' };
  const id = normId(p.id);
  const rec = id ? findPerson_(id) : null;
  if (!rec || rec[R.REMOVED]) return { ok: true, result: 'notFound' };
  const time = checkinTimeOf_(id, rec);
  flushCheckins_();
  if (!time && state !== STATE.OPEN) return { ok: true, result: 'closed' };
  return {
    ok: true, result: 'found', id: id,
    name: rec[R.NAME], unit: rec[R.UNIT], eligible: !!rec[R.ELIGIBLE], time: time || null,
  };
}

function apiCheckin_(p) {
  const state = getConfig_().state;
  if (state === STATE.PREP) return { ok: true, result: 'prep' };
  if (state === STATE.ENDED) return { ok: true, result: 'ended' };
  const id = normId(p.id);
  const rec = id ? findPerson_(id) : null;
  if (!rec || rec[R.REMOVED]) return { ok: true, result: 'notFound' };
  const who = { id: id, name: rec[R.NAME], unit: rec[R.UNIT], eligible: !!rec[R.ELIGIBLE] };

  const already = checkinTimeOf_(id, rec);
  if (already) {
    flushCheckins_();
    return Object.assign({ ok: true, result: 'already', time: already }, who);
  }
  if (state !== STATE.OPEN) return { ok: true, result: 'closed' };

  const now = Date.now();
  PropertiesService.getScriptProperties().setProperty(PENDING + id, String(now));
  const time = Utilities.formatDate(new Date(now), 'Asia/Taipei', 'HH:mm');
  const cache = CacheService.getScriptCache();
  cache.putAll({ ['ci:' + id]: time, pend: '1' }, MAP_TTL);
  flushCheckins_();
  return Object.assign({ ok: true, result: 'done', time: time }, who);
}

// 報到時間：剛報到的看快取，尚未寫進 Sheet 的看待寫入清單，其餘看對照表
function checkinTimeOf_(id, rec) {
  const hit = CacheService.getScriptCache().get('ci:' + id);
  if (hit) return hit;
  const pending = PropertiesService.getScriptProperties().getProperty(PENDING + id);
  if (pending) return Utilities.formatDate(new Date(Number(pending)), 'Asia/Taipei', 'HH:mm');
  return rec[R.TIME];
}

// 有待寫入的報到、而且鎖是空的才寫；寫完若又有新的進來，最多再接著寫兩輪
function flushCheckins_() {
  const cache = CacheService.getScriptCache();
  if (!cache.get('pend')) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;
  try {
    for (let round = 0; round < 3 && cache.get('pend'); round++) flushLocked_();
  } finally {
    lock.releaseLock();
  }
}

// 呼叫前要先持有 Script Lock。依工號欄重新找列，不依賴記住的列號；已有報到時間的不覆寫
function flushLocked_() {
  const t0 = Date.now();
  CacheService.getScriptCache().remove('pend');
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const keys = Object.keys(all).filter(k => k.indexOf(PENDING) === 0);
  if (!keys.length) return 0;

  const sh = sheet_(SHEET.PEOPLE);
  const last = sh.getLastRow();
  const ids = last >= 2 ? sh.getRange(2, PC.ID, last - 1, 1).getDisplayValues() : [];
  const times = last >= 2 ? sh.getRange(2, PC.TIME, last - 1, 1).getValues() : [];
  const rowOf = {};
  ids.forEach((r, i) => {
    const id = normId(r[0]);
    if (id && !(id in rowOf)) rowOf[id] = i;
  });

  const missing = [];
  keys.forEach(k => {
    const id = k.slice(PENDING.length);
    const i = rowOf[id];
    if (i == null) missing.push(id);
    else if (times[i][0] === '') sh.getRange(i + 2, PC.TIME).setValue(new Date(Number(all[k])));
  });
  SpreadsheetApp.flush();
  keys.forEach(k => props.deleteProperty(k));
  if (missing.length) log_('報到寫入', false, Date.now() - t0, missing.join(','), '', '人員分頁找不到這些工號，報到時間沒有寫入');
  return keys.length;
}

function discardPendingCheckins_() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).filter(k => k.indexOf(PENDING) === 0).forEach(k => props.deleteProperty(k));
  CacheService.getScriptCache().remove('pend');
}

// 查不到時重讀人員分頁（現場新增的人），每 10 秒最多一次
function findPerson_(id) {
  const map = loadMap_();
  if (map[id]) return map[id];
  const cache = CacheService.getScriptCache();
  if (cache.get('ppl:retry')) return null;
  cache.put('ppl:retry', '1', 10);
  return rebuildMap_()[id] || null;
}

function loadMap_() {
  const cache = CacheService.getScriptCache();
  const meta = cache.getAll(['ppl:gen', 'ppl:dirty']);
  if (meta['ppl:gen'] && !meta['ppl:dirty']) {
    const map = readMapFromCache_(meta['ppl:gen']);
    if (map) return map;
  }
  return rebuildMap_();
}

function readMapFromCache_(gen) {
  const parts = String(gen).split(':');
  const keys = [];
  for (let i = 0; i < Number(parts[1]); i++) keys.push('ppl:' + parts[0] + ':' + i);
  const got = CacheService.getScriptCache().getAll(keys);
  if (!keys.every(k => got[k] != null)) return null;
  return JSON.parse(keys.map(k => got[k]).join(''));
}

// 搶鎖重建；等鎖期間別人已經重建好的話直接沿用
function rebuildMap_() {
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const cache = CacheService.getScriptCache();
    const meta = cache.getAll(['ppl:gen', 'ppl:dirty']);
    if (meta['ppl:gen'] && !meta['ppl:dirty'] && Number(meta['ppl:gen'].split(':')[0]) > t0) {
      const map = readMapFromCache_(meta['ppl:gen']);
      if (map) return map;
    }
    cache.remove('ppl:dirty');
    const map = buildMapFromSheet_();
    const gen = Date.now();
    const json = JSON.stringify(map);
    const chunks = {};
    let n = 0;
    for (let i = 0; i < json.length; i += MAP_CHUNK) chunks['ppl:' + gen + ':' + (n++)] = json.slice(i, i + MAP_CHUNK);
    cache.putAll(chunks, MAP_TTL);
    cache.put('ppl:gen', gen + ':' + n, MAP_TTL);
    const ciKeys = Object.keys(map).map(id => 'ci:' + id);
    for (let i = 0; i < ciKeys.length; i += 500) cache.removeAll(ciKeys.slice(i, i + 500));
    return map;
  } finally {
    lock.releaseLock();
  }
}

function buildMapFromSheet_() {
  const sh = sheet_(SHEET.PEOPLE);
  const last = sh.getLastRow();
  const map = {};
  if (last < 2) return map;
  const rows = sh.getRange(2, 1, last - 1, PEOPLE_HEADERS.length).getDisplayValues();
  rows.forEach((r, i) => {
    const id = normId(r[PC.ID - 1]);
    if (!id || map[id]) return;
    map[id] = [
      maskName(r[PC.NAME - 1]),
      String(r[PC.UNIT - 1]).trim(),
      String(r[PC.ELIGIBLE - 1]).trim() === '是' ? 1 : 0,
      String(r[PC.REG - 1]).trim() === '移除' ? 1 : 0,
      i + 2,
      hhmm_(r[PC.TIME - 1]),
    ];
  });
  return map;
}

function markPeopleDirty_() {
  CacheService.getScriptCache().put('ppl:dirty', '1', MAP_TTL);
}
