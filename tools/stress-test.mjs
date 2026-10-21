// 報到壓力測試：在指定秒數內隨機送出 N 人的完整報到（config → lookupId → checkin），記錄逐筆耗時
// 用法：node tools/stress-test.mjs --e <部署作業ID> --csv test-people.csv [--n 300] [--window 30] [--out stress-result.csv]
// 需要 Node 18 以上；系統狀態要先切到「報到中」，而且這批工號還沒報到過

import { readFileSync, writeFileSync } from 'node:fs';

const args = Object.fromEntries(process.argv.slice(2).reduce((acc, v, i, a) => {
  if (v.startsWith('--')) acc.push([v.slice(2), a[i + 1]]);
  return acc;
}, []));
if (!args.e || !args.csv) {
  console.error('用法：node tools/stress-test.mjs --e <部署作業ID> --csv test-people.csv [--n 300] [--window 30]');
  process.exit(1);
}
const N = Number(args.n || 300);
const WINDOW = Number(args.window || 30) * 1000;
const OUT = args.out || 'stress-result.csv';
const API = `https://script.google.com/macros/s/${args.e}/exec`;

const ids = readFileSync(args.csv, 'utf8').replace(/^﻿/, '').trim().split(/\r?\n/).slice(1)
  .map(line => line.split(',')[3]).filter(Boolean);
if (ids.length < N) {
  console.error(`CSV 只有 ${ids.length} 個工號，不夠 ${N} 人`);
  process.exit(1);
}
for (let i = ids.length - 1; i > 0; i--) {
  const j = Math.floor(Math.random() * (i + 1));
  [ids[i], ids[j]] = [ids[j], ids[i]];
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// 與報到頁相同的重試規則
async function call(params, post) {
  const deadline = Date.now() + 30000;
  let wait = 1000 + Math.random() * 2000;
  let tries = 0;
  for (;;) {
    tries++;
    const left = deadline - Date.now();
    if (left < 1500) return { data: null, tries };
    try {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), Math.min(15000, left));
      const res = post
        ? await fetch(API, { method: 'POST', body: JSON.stringify(params), headers: { 'Content-Type': 'text/plain;charset=utf-8' }, signal: ctrl.signal })
        : await fetch(`${API}?${new URLSearchParams(params)}`, { signal: ctrl.signal });
      const data = await res.json();
      clearTimeout(timer);
      if (data && (data.ok || data.code !== 'error')) return { data, tries };
    } catch (e) { /* 重試 */ }
    if (Date.now() + wait > deadline) return { data: null, tries };
    await sleep(wait);
    wait = Math.min(wait * 1.6, 8000) * (0.8 + Math.random() * 0.4);
  }
}

async function oneUser(id, startAt, t0) {
  await sleep(startAt);
  const rec = { id, startAt, steps: {}, tries: 0, result: '', total: 0 };
  const begin = Date.now();
  for (const [name, params, post] of [
    ['config', { a: 'config' }, false],
    ['lookupId', { a: 'lookupId', id }, false],
    ['checkin', { a: 'checkin', id }, true],
  ]) {
    const s = Date.now();
    const { data, tries } = await call(params, post);
    rec.steps[name] = Date.now() - s;
    rec.tries += tries;
    if (!data) { rec.result = `${name}:失敗`; break; }
    if (name === 'lookupId' && data.result !== 'found') { rec.result = `lookupId:${data.result}`; break; }
    if (name === 'checkin') rec.result = data.result;
  }
  rec.total = Date.now() - begin;
  rec.finishedAt = Date.now() - t0;
  process.stdout.write(rec.result === 'done' ? '.' : 'x');
  return rec;
}

const pct = (arr, p) => {
  if (!arr.length) return 0;
  const s = [...arr].sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.floor(s.length * p))];
};

console.log(`送出 ${N} 人，分散在 ${WINDOW / 1000} 秒內…`);
const t0 = Date.now();
const recs = await Promise.all(ids.slice(0, N).map(id => oneUser(id, Math.random() * WINDOW, t0)));
console.log('\n');

const done = recs.filter(r => r.result === 'done');
const already = recs.filter(r => r.result === 'already');
const failed = recs.filter(r => r.result !== 'done' && r.result !== 'already');
const within60 = recs.filter(r => (r.result === 'done' || r.result === 'already') && r.finishedAt <= 60000);
const retried = recs.filter(r => r.tries > 3);

console.log(`成功報到（done）：${done.length}`);
console.log(`回報已報到（already）：${already.length}　← 第一次報到應該是 0；有數字代表重試時前一次其實已寫入，仍是單次寫入`);
console.log(`失敗：${failed.length}${failed.length ? '　' + [...new Set(failed.map(r => r.result))].join('、') : ''}`);
console.log(`開始後 60 秒內完成：${within60.length} / ${N}`);
console.log(`有重試的人：${retried.length}`);
for (const step of ['config', 'lookupId', 'checkin']) {
  const t = recs.map(r => r.steps[step]).filter(v => v != null);
  console.log(`${step.padEnd(9)} 中位數 ${pct(t, 0.5)} ms｜95% ${pct(t, 0.95)} ms｜最慢 ${Math.max(...t)} ms`);
}
const tot = recs.map(r => r.total);
console.log(`每人總耗時 中位數 ${pct(tot, 0.5)} ms｜95% ${pct(tot, 0.95)} ms｜最慢 ${Math.max(...tot)} ms`);

const csv = [['工號', '開始(ms)', '完成(ms)', 'config(ms)', 'lookupId(ms)', 'checkin(ms)', '總耗時(ms)', '呼叫次數', '結果']]
  .concat(recs.map(r => [r.id, Math.round(r.startAt), r.finishedAt, r.steps.config ?? '', r.steps.lookupId ?? '', r.steps.checkin ?? '', r.total, r.tries, r.result]));
writeFileSync(OUT, '﻿' + csv.map(r => r.join(',')).join('\n') + '\n');
console.log(`\n逐筆紀錄：${OUT}`);
const pass = done.length + already.length === N && within60.length === N;
console.log(pass ? '✅ 通過：全部在 60 秒內成功' : '❌ 未通過');
process.exit(pass ? 0 : 2);
