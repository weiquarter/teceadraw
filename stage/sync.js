// 前台同步模組：前台代號、呼叫 GAS、待同步操作的暫存與補寫、心跳。畫面程式只透過這裡跟後端溝通。
(function (global) {
  'use strict';

  const HEARTBEAT_MS = 5000;

  function createSync(opts) {
    const e = opts.e;
    const k = opts.k;
    const api = 'https://script.google.com/macros/s/' + encodeURIComponent(e) + '/exec';
    const ns = 'td:' + e + ':';
    const listeners = { status: [], beat: [] };

    const store = {
      get(key, fallback) {
        try { const v = localStorage.getItem(ns + key); return v == null ? fallback : JSON.parse(v); } catch (err) { return fallback; }
      },
      set(key, value) {
        try { localStorage.setItem(ns + key, JSON.stringify(value)); } catch (err) { /* 無痕或空間不足時只存在記憶體 */ }
      },
    };

    // 前台代號存在這台電腦，F5 後仍是同一台
    let code = store.get('code', '');
    if (!/^\d{4}$/.test(code)) {
      const buf = new Uint32Array(1);
      crypto.getRandomValues(buf);
      code = String(1000 + (buf[0] % 9000));
      store.set('code', code);
    }

    let pending = store.get('pending', []);
    let opSeq = store.get('opSeq', 0);
    let online = true;
    let replaced = false;
    let lastError = '';
    let flushing = false;
    let retryTimer = null;
    let retryWait = 2000;

    function emit(type, data) { listeners[type].forEach(fn => fn(data)); }
    function status() { return { code, online, replaced, pending: pending.length, lastError }; }
    function notify() { emit('status', status()); }

    async function call(params, post, timeoutMs) {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), timeoutMs || 20000);
      try {
        const body = Object.assign({ k, code }, params);
        const res = post
          ? await fetch(api, { method: 'POST', body: JSON.stringify(body), headers: { 'Content-Type': 'text/plain;charset=utf-8' }, signal: ctrl.signal })
          : await fetch(api + '?' + new URLSearchParams(body), { signal: ctrl.signal });
        const data = await res.json();
        if (data && data.code === 'error') throw new Error(data.message || 'GAS 錯誤');
        if (!online) { online = true; notify(); }
        return data;
      } catch (err) {
        if (online) { online = false; notify(); }
        throw err;
      } finally {
        clearTimeout(timer);
      }
    }

    // 新操作：先存本機，再嘗試補寫
    function enqueue(op) {
      op.id = code + '-' + (++opSeq);
      op.at = Date.now();
      store.set('opSeq', opSeq);
      pending.push(op);
      store.set('pending', pending);
      notify();
      flush();
      return op;
    }

    async function flush() {
      if (flushing || replaced || !pending.length) return;
      flushing = true;
      clearTimeout(retryTimer);
      const batch = pending.slice(0, 20);
      try {
        const r = await call({ a: 'recordDraw', v: opts.version(), ops: batch }, true, 30000);
        if (r.ok) {
          const acked = new Set(r.acked || []);
          pending = pending.filter(o => !acked.has(o.id));
          store.set('pending', pending);
          lastError = '';
          retryWait = 2000;
        } else if (r.code === 'notActive') {
          replaced = true;
          lastError = r.message;
        } else {
          lastError = r.message || r.code;
          emit('beat', { rejected: r });
        }
      } catch (err) {
        lastError = '';
      } finally {
        flushing = false;
        notify();
      }
      if (pending.length && !replaced) {
        // 還有下一批就馬上送；被後端拒絕（版本、獎項變更）或離線時放慢
        const delay = lastError ? 15000 : online ? 300 : retryWait;
        retryTimer = setTimeout(flush, delay);
        if (!online) retryWait = Math.min(retryWait * 1.6, 15000);
      }
    }

    let beatTimer = null;
    async function beat() {
      clearTimeout(beatTimer);
      try {
        const s = opts.report();
        const r = await call({ a: 'heartbeat', page: s.page, prize: s.prize, pending: pending.length }, false, 12000);
        if (r.ok && r.active === false) replaced = true;
        if (r.ok && r.active) replaced = false;
        emit('beat', r);
        notify();
        if (pending.length) flush();
      } catch (err) { /* 離線：下一輪再試 */ }
      beatTimer = setTimeout(beat, HEARTBEAT_MS);
    }

    async function load() {
      const r = await call({ a: 'stageLoad' }, false, 30000);
      if (r.ok && r.ready) store.set('lastLoad', r);
      return r;
    }

    async function asset(id) {
      for (let i = 0; i < 3; i++) {
        try {
          const r = await call({ a: 'stageAsset', id }, false, 45000);
          if (r.ok) return r;
          return null;
        } catch (err) {
          await new Promise(res => setTimeout(res, 1500 * (i + 1)));
        }
      }
      return null;
    }

    // 版本變更（解除鎖定後重新鎖定、清除資料）時，舊版本的待同步操作作廢
    function resetForVersion(v) {
      if (store.get('version', null) !== v) {
        pending = [];
        store.set('pending', pending);
        store.set('version', v);
        notify();
      }
    }

    addEventListener('online', () => { flush(); beat(); });

    return {
      get code() { return code; },
      status,
      pendingOps: () => pending.slice(),
      cachedLoad: () => store.get('lastLoad', null),
      remember: (key, value) => store.set('ui:' + key, value),
      recall: (key, fallback) => store.get('ui:' + key, fallback),
      on: (type, fn) => listeners[type].push(fn),
      load, asset, enqueue, flush, resetForVersion,
      start: () => beat(),
      beatNow: () => beat(),
    };
  }

  global.createSync = createSync;
})(window);
