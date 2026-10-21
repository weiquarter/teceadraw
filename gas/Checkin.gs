/** 報到：工號對照表（放 CacheService）、lookupId、checkin。checkin 不搶鎖，只寫自己那一列。 */

const MAP_TTL = 21600;
const MAP_CHUNK = 30000;

// 對照表內容：工號 → [遮蔽姓名, 單位, 可否摸彩, 是否移除, 列號, 報到時間 HH:mm]
const R = { NAME: 0, UNIT: 1, ELIGIBLE: 2, REMOVED: 3, ROW: 4, TIME: 5 };

function apiLookupId_(p) {
  const state = getConfig_().state;
  if (state === STATE.PREP) return { ok: true, result: 'prep' };
  if (state === STATE.ENDED) return { ok: true, result: 'ended' };
  const id = normId(p.id);
  const rec = id ? findPerson_(id) : null;
  if (!rec || rec[R.REMOVED]) return { ok: true, result: 'notFound' };
  const time = CacheService.getScriptCache().get('ci:' + id) || rec[R.TIME];
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
  let rec = id ? findPerson_(id) : null;
  if (!rec) return { ok: true, result: 'notFound' };

  const sh = sheet_(SHEET.PEOPLE);
  let row = sh.getRange(rec[R.ROW], 1, 1, PEOPLE_HEADERS.length).getDisplayValues()[0];
  if (normId(row[PC.ID - 1]) !== id) {
    // 有人刪列或排序，列號跑掉了
    rec = rebuildMap_()[id];
    if (!rec) return { ok: true, result: 'notFound' };
    row = sh.getRange(rec[R.ROW], 1, 1, PEOPLE_HEADERS.length).getDisplayValues()[0];
    if (normId(row[PC.ID - 1]) !== id) throw new Error('對照表重建後列號仍不吻合');
  }

  if (String(row[PC.REG - 1]).trim() === '移除') return { ok: true, result: 'notFound' };
  const who = {
    id: id,
    name: maskName(row[PC.NAME - 1]),
    unit: String(row[PC.UNIT - 1]).trim(),
    eligible: String(row[PC.ELIGIBLE - 1]).trim() === '是',
  };
  if (String(row[PC.TIME - 1]).trim() !== '') {
    return Object.assign({ ok: true, result: 'already', time: hhmm_(row[PC.TIME - 1]) }, who);
  }
  if (state !== STATE.OPEN) return { ok: true, result: 'closed' };

  const now = new Date();
  sh.getRange(rec[R.ROW], PC.TIME).setValue(now);
  const time = Utilities.formatDate(now, 'Asia/Taipei', 'HH:mm');
  CacheService.getScriptCache().put('ci:' + id, time, MAP_TTL);
  return Object.assign({ ok: true, result: 'done', time: time }, who);
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
