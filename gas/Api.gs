/** Web API 進入點：依參數 a 分派到各動作，統一回傳 JSON；出錯時寫 Log 並回傳可重試的錯誤。 */

// 用函式包起來：GAS 依檔案順序執行頂層程式碼，直接寫成常數會拿不到其他檔案的函式
function routes_() {
  return {
    ping: apiPing_,
    config: apiConfig_,
    lookupId: apiLookupId_,
    checkin: apiCheckin_,
    heartbeat: apiHeartbeat_,
    stageLoad: apiStageLoad_,
    stageAsset: apiStageAsset_,
    recordDraw: apiRecordDraw_,
  };
}

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }
  return handle_(Object.assign({}, (e && e.parameter) || {}, body));
}

function handle_(p) {
  const t0 = Date.now();
  const fn = Object.prototype.hasOwnProperty.call(routes_(), p.a) ? routes_()[p.a] : null;
  let out;
  try {
    out = fn ? fn(p) : { ok: false, code: 'badAction', message: '未知的動作' };
  } catch (err) {
    log_(p.a || '?', false, Date.now() - t0, p.id, '', err);
    out = { ok: false, code: 'error', message: '系統忙碌，請稍後再試' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// 給「初始化」匿名檢查部署用，不含任何活動資料
function apiPing_() {
  return { ok: true, apiVersion: API_VERSION, build: BUILD, sheet: sheetHash_(), timeZone: Session.getScriptTimeZone() };
}

function apiConfig_() {
  const c = getConfig_();
  if (c.state === STATE.ENDED) {
    return { ok: true, apiVersion: API_VERSION, state: c.state, texts: { ended: c.texts.ended } };
  }
  return Object.assign({ ok: true, apiVersion: API_VERSION }, c);
}
