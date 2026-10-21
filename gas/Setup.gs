/** 活動前與進階：建立範本結構與控制台、初始化（資料夾＋部署檢查）、產生連結、資料檢查、模擬全員報到、清除資料、重新產生前台連結。 */

function settingRows_() {
  return [
    ['活動名稱', '活動名稱', '所有畫面的主標題'],
    ['副標題', '', '例如日期、地點'],
    ['配色主題', '草莓牛奶', '草莓牛奶／薄荷檸檬／薰衣草藍莓'],
    ['拉霸樣式', 'A 跳字', 'A 跳字／B 直向轉輪／C 翻牌'],
    ['字型', 'Huninn', 'Google Fonts 的字型名稱'],
    ['主色', '', '進階：色碼如 #F6B8C6，空白＝用配色主題'],
    ['輔色', '', '進階：色碼，空白＝用配色主題'],
    ['背景色', '', '進階：色碼，空白＝用配色主題'],
    ['文字色', '', '進階：色碼，空白＝用配色主題'],
    ['提示_報到成功', '報到成功', ''],
    ['提示_具抽獎資格', '具抽獎資格', ''],
    ['提示_無抽獎資格', '本次無抽獎資格', ''],
    ['提示_報到已截止', '報到已截止', ''],
    ['提示_查無工號', '查無此工號，請至報到處', ''],
    ['提示_系統準備中', '系統準備中，請稍後再試', ''],
    ['提示_活動已結束', '活動已結束', ''],
    ['提示_請截圖', '請截圖保存', ''],
    ['前端網址', DEFAULT_PAGES_BASE, '進階：fork 自架時才需要改'],
  ];
}

// 在 Apps Script 編輯器手動執行，建立範本 Sheet 的所有分頁；已存在的分頁不會動，控制台每次都重建
function setupTemplate() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  try {
    // 控制台的「最近 10 分鐘錯誤」靠 NOW() 每分鐘重算
    ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE_AND_MINUTE);
  } catch (err) {
    Logger.log('無法設定重新計算頻率，請手動到「檔案 › 設定 › 計算」改成「變更時和每分鐘」：' + err.message);
  }

  const created = [];
  const make = (name, headers, fn) => {
    if (ss.getSheetByName(name)) return;
    const sh = ss.insertSheet(name);
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3E7EA');
      sh.setFrozenRows(1);
    }
    if (fn) fn(sh);
    created.push(name);
  };

  make(SHEET.PEOPLE, PEOPLE_HEADERS, sh => {
    ensureRows_(sh, 2000);
    sh.getRange('D2:D').setNumberFormat('@');
    sh.getRange('H2:H').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sh.getRange('F2:F').setDataValidation(listRule_(['是', '否']));
    sh.getRange('G2:G').setDataValidation(listRule_(['既有', '現場新增', '移除']));
    [80, 100, 200, 110, 110, 90, 100, 170, 160].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  });
  make(SHEET.PRIZES, PRIZE_HEADERS, sh => {
    sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 10).setAllowInvalid(false).setHelpText('數量 1～10，超過請拆成兩個獎項').build());
    sh.getRange('E2:E').setDataValidation(listRule_(['一次抽完', '逐個抽出']));
    [60, 260, 70, 180, 110, 90].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  });
  make(SHEET.WINNERS, WINNER_HEADERS, sh => {
    sh.getRange('E2:E').setNumberFormat('@');
    sh.getRange('A2:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2="作廢"').setStrikethrough(true).setFontColor('#999999')
      .setRanges([sh.getRange('A2:H')]).build()]);
    [150, 220, 80, 200, 110, 110, 70, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  });
  make(SHEET.SETTINGS, ['項目', '值', '說明'], sh => {
    const rows = settingRows_();
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
    sh.getRange('B4').setDataValidation(listRule_(Object.keys(PALETTE_IDS)));
    sh.getRange('B5').setDataValidation(listRule_(['A 跳字', 'B 直向轉輪', 'C 翻牌']));
    [140, 320, 320].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  });
  make(SHEET.LOG, LOG_HEADERS, sh => {
    sh.getRange('A2:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sh.getRange('E2:E').setNumberFormat('@');
  });
  make(SHEET.SYSTEM, ['項目', '值'], sh => {
    sh.getRange(2, 1, 4, 2).setValues([['系統狀態', STATE.PREP], ['部署ID', ''], ['素材資料夾ID', ''], ['備份資料夾ID', '']]);
    sh.hideSheet();
  });
  [SHEET.STAGE, SHEET.SNAPSHOT, SHEET.OPLOG].forEach(name => make(name, null, sh => sh.hideSheet()));

  // 控制台公式引用其他分頁，要最後建，否則公式會指向還不存在的分頁而算錯；內容全是公式，每次重建以套用新版
  const old = ss.getSheetByName(SHEET.CONSOLE);
  const next = old ? old.getRange('B17').getValue() : '';
  if (old) ss.deleteSheet(old);
  const consoleSheet = ss.insertSheet(SHEET.CONSOLE, 0);
  buildConsole_(consoleSheet);
  if (next) consoleSheet.getRange('B17').setValue(next);
  created.push(SHEET.CONSOLE + '（重建）');

  ['工作表1', 'Sheet1'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(ss.getSheetByName(SHEET.CONSOLE));
  ss.moveActiveSheet(1);
  Logger.log(created.length ? '已建立：' + created.join('、') : '所有分頁都已存在，沒有變更');
}

function buildConsole_(sh) {
  const kv = (sheetName, key, fallback) =>
    "IFERROR(VLOOKUP(\"" + key + "\",'" + sheetName + "'!A:B,2,FALSE),\"" + fallback + "\")";
  const live = "'" + SHEET.PEOPLE + "'!";
  const log = "'" + SHEET.LOG + "'!";
  const prize = "'" + SHEET.PRIZES + "'!";
  const nextStep = 'SWITCH(B3,' + Object.keys(NEXT_STEP).map(k => '"' + k + '","' + NEXT_STEP[k] + '"').join(',') + ',"")';
  const stage = k => '=' + kv(SHEET.STAGE, k, '—');
  const rows = [
    ['摸彩系統 控制台', ''],
    ['', ''],
    ['目前狀態', '=' + kv(SHEET.SYSTEM, '系統狀態', '準備中')],
    ['下一步', '=' + nextStep],
    ['', ''],
    ['報到統計', ''],
    ['已報到', '=COUNTIFS(' + live + 'H2:H,"<>",' + live + 'D2:D,"<>",' + live + 'G2:G,"<>移除")'],
    ['總人數', '=COUNTIFS(' + live + 'D2:D,"<>",' + live + 'G2:G,"<>移除")'],
    ['有抽獎資格且已報到', '=COUNTIFS(' + live + 'H2:H,"<>",' + live + 'D2:D,"<>",' + live + 'F2:F,"是",' + live + 'G2:G,"<>移除")'],
    ['', ''],
    ['名單快照', ''],
    ['版本', '=IF(' + kv(SHEET.SYSTEM, '名單檢查碼', '') + '="","（尚未鎖定）",' + kv(SHEET.SYSTEM, '名單版本', '') + ')'],
    ['抽獎池人數', '=' + kv(SHEET.SYSTEM, '抽獎池人數', '')],
    ['檢查碼', '=' + kv(SHEET.SYSTEM, '名單檢查碼', '') + '&IF(' + kv(SHEET.SYSTEM, '獎項檢查碼', '') + '="",""," ／ 獎項 "&' + kv(SHEET.SYSTEM, '獎項檢查碼', '') + ')'],
    ['獎項', '=IF(' + kv(SHEET.SYSTEM, '獎項鎖定後變更', '') + '="是","⚠️ 獎項在鎖定後被修改，請解除鎖定後重新鎖定","")'],
    ['', ''],
    ['下一個要抽的獎項', ''],
    ['', '=IF(B17<>"","↑ 前台回主頁時生效；清空＝照排序","空白＝照排序：" & IFERROR(INDEX(SORT(FILTER({' + prize + 'A2:A,' + prize + 'B2:B},' + prize + 'B2:B<>"",' + prize + 'F2:F<>"已完成"),1,TRUE),1,2),"（沒有待抽的獎項）"))'],
    ['', ''],
    ['前台監看', ''],
    ['前台代號', stage('前台代號')],
    ['目前獎項', stage('目前獎項')],
    ['目前頁面', stage('目前頁面')],
    ['待同步筆數', stage('待同步筆數')],
    ['最後回報時間', stage('最後回報時間')],
    ['舊前台補寫被拒', '=IF(N(' + kv(SHEET.STAGE, '被拒絕的補寫', '0') + ')>0,' + kv(SHEET.STAGE, '被拒絕的補寫', '0') + '&" 筆，最後一次 "&TEXT(' + kv(SHEET.STAGE, '最後被拒絕時間', '') + ',"HH:mm:ss")&"，請以台下結果人工核對","—")'],
    ['錯誤監看', ''],
    ['最近 10 分鐘 GAS 錯誤', '=COUNTIFS(' + log + 'A2:A,">"&(NOW()-10/1440),' + log + 'C2:C,"失敗")'],
    ['', ''],
    ['提示', '第一次使用「摸彩系統」選單時會要求授權，屬正常現象。\n這一頁只有「下一個要抽的獎項」可以改，其他都會自動更新。'],
    ['', ''],
    ['即時得獎名單（目前獎項）', ''],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange('A33:E33').setValues([['名額序號', '單位', '工號', '姓名', '狀態']]).setFontWeight('bold').setBackground('#F3E7EA');
  sh.getRange('A34').setFormula("=IFERROR(FILTER({'" + SHEET.WINNERS + "'!C2:C,'" + SHEET.WINNERS + "'!D2:D,'" + SHEET.WINNERS + "'!E2:E,'" + SHEET.WINNERS + "'!F2:F,'" + SHEET.WINNERS + "'!G2:G},'" + SHEET.WINNERS + "'!B2:B=B22),\"（還沒有得獎人）\")");

  sh.getRange('G1').setValue('資料檢查結果').setFontSize(14).setFontWeight('bold');
  sh.getRange('G2').setValue('（還沒有檢查過。按「摸彩系統 › 活動前 › 資料檢查」）');
  sh.getRange('G4:I4').setValues([['等級', '項目', '說明']]).setFontWeight('bold').setBackground('#F3E7EA');

  sh.getRange('A1').setFontSize(18).setFontWeight('bold');
  sh.getRange('A3:B4').setFontSize(20).setFontWeight('bold');
  ['A6', 'A11', 'A17', 'A20', 'A27', 'A32'].forEach(a => sh.getRange(a).setFontWeight('bold').setFontSize(12));
  sh.getRange('B7:B9').setFontSize(16);
  sh.getRange('B2:B30').setHorizontalAlignment('left');
  sh.getRange('B15').setFontColor('#B5481C').setFontWeight('bold');
  sh.getRange('B18').setFontColor('#6E5A60');
  sh.getRange('B28').setFontColor('#B5481C');
  sh.getRange('B30').setWrap(true);
  sh.getRange('B25').setNumberFormat('hh:mm:ss');

  const b17 = sh.getRange('B17');
  b17.setBackground('#FFF3C4').setFontSize(14).setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(sheet_(SHEET.PRIZES).getRange('B2:B'), true).setAllowInvalid(false).build());

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($B$25),NOW()-$B$25>15/86400)')
      .setBackground('#F8C9C0').setFontColor('#8A1C0B').setRanges([sh.getRange('A25:B25')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B$28>0').setBackground('#F8C9C0').setRanges([sh.getRange('A28:B28')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B$26<>"—"').setBackground('#F8C9C0').setRanges([sh.getRange('A26:B26')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E34="作廢"').setStrikethrough(true).setFontColor('#999999').setRanges([sh.getRange('A34:E50')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('錯誤').setBackground('#F8C9C0').setRanges([sh.getRange('G5:G300')]).build(),
  ];
  sh.setConditionalFormatRules(rules);
  [230, 360, 110, 110, 90, 30, 70, 260, 420].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(4);

  // 只有「下一個要抽的獎項」可以改，其他格子改了會跳警告
  const p = sh.protect().setWarningOnly(true).setDescription('摸彩系統：控制台只有「下一個要抽的獎項」可以改');
  p.setUnprotectedRanges([b17]);
}

function listRule_(values) {
  return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
}

function ensureRows_(sh, n) {
  if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows());
}

// ── 初始化 ───────────────────────────────────────────

function menuInit() {
  const folders = ensureFolders_();
  const t = HtmlService.createTemplateFromFile('InitDialog');
  t.folders = folders;
  t.deployId = String(sysGet_('部署ID') || '');
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(640).setHeight(620), '初始化');
}

// 在 Sheet 所在的資料夾建立「素材」「備份」，已建立過就沿用
function ensureFolders_() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = file.getParents();
  if (!parents.hasNext()) throw new Error('找不到這份 Sheet 所在的資料夾');
  const parent = parents.next();
  const out = {};
  [['素材', '素材資料夾ID'], ['備份', '備份資料夾ID']].forEach(([name, key]) => {
    let folder = null;
    const id = sysGet_(key);
    if (id) { try { folder = DriveApp.getFolderById(id); } catch (e) { folder = null; } }
    if (!folder) {
      const it = parent.getFoldersByName(name);
      folder = it.hasNext() ? it.next() : parent.createFolder(name);
      sysSet_(key, folder.getId());
    }
    out[name] = folder.getUrl();
  });
  return out;
}

// 初始化對話框呼叫：檢查通過才把部署作業 ID 存起來
function verifyDeployment(input) {
  const m = String(input || '').trim().match(/\/s\/([\w-]+)\//) || String(input || '').trim().match(/^([\w-]{30,})$/);
  if (!m) return { ok: false, message: '看不懂貼上的內容。請貼「部署作業 ID」（AKfycb 開頭的一長串），或網頁應用程式網址。' };
  const r = checkDeployment_(m[1]);
  if (!r.ok) return r;
  sysSet_('部署ID', m[1]);
  log_('初始化', true, null, m[1]);
  return { ok: true, message: '檢查通過：匿名存取正常，部署屬於這份 Sheet。' + (r.warn ? '\n' + r.warn : '') };
}

// 以匿名方式打部署網址，確認存取權、對應的 Sheet、時區與程式版本
function checkDeployment_(id) {
  const url = 'https://script.google.com/macros/s/' + id + '/exec?a=ping';
  let res;
  try {
    res = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true });
  } catch (err) {
    return { ok: false, message: '連不到部署網址：' + err.message };
  }
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code === 404) return { ok: false, message: '找不到這個部署，可能已經被封存，或貼的不是「網頁應用程式」的部署作業 ID。' };
  let data = null;
  try { data = JSON.parse(text); } catch (e) { data = null; }
  if (!data) {
    if (/accounts\.google\.com|ServiceLogin|signin/i.test(text)) {
      return { ok: false, message: '匿名存取被導去登入頁。請到「管理部署作業」把「誰可以存取」改成「任何人」，執行身分選「我」。如果選單裡沒有「任何人」，代表 Workspace 管理員關閉了這個選項，需要請管理員開啟。' };
    }
    return { ok: false, message: '部署網址沒有回傳預期的資料（HTTP ' + code + '），可能已經被封存。請確認部署作業 ID。' };
  }
  if (!data.ok) return { ok: false, message: '部署的程式版本太舊，沒有檢查用的功能。請「管理部署作業 › 編輯現有部署 › 版本：新版本」後再試。' };
  if (data.sheet !== sheetHash_()) return { ok: false, message: '這個部署屬於另一份 Sheet。請在這份 Sheet 的「擴充功能 › Apps Script」裡部署，再貼上那個部署作業 ID。' };
  if (data.apiVersion !== API_VERSION) return { ok: false, message: '部署的程式版本跟目前程式碼不相容，請部署新版本。' };
  const warn = [];
  if (Number(data.build || 0) < BUILD) {
    warn.push('⚠️ 部署的版本比目前的程式舊。請到 Apps Script「部署 › 管理部署作業」，編輯現有的部署、版本選最新（不要新增部署，ID 會變）。');
  }
  if (data.timeZone !== 'Asia/Taipei') warn.push('⚠️ 專案時區是 ' + data.timeZone + '，請到「專案設定」改成 (GMT+08:00) 台北。');
  return { ok: true, warn: warn.join('\n') };
}

// ── 產生連結 ─────────────────────────────────────────

function menuLinks() {
  const id = sysGet_('部署ID');
  if (!id) {
    SpreadsheetApp.getUi().alert('請先完成「摸彩系統 › 活動前 › 初始化」。');
    return;
  }
  const base = pagesBase_();
  const t = HtmlService.createTemplateFromFile('LinksDialog');
  t.checkinUrl = base + '/checkin/?e=' + encodeURIComponent(id);
  t.stageUrl = base + '/stage/?e=' + encodeURIComponent(id) + '&k=' + encodeURIComponent(stageKey_());
  t.title = String(readKeyValue_(SHEET.SETTINGS)['活動名稱'] || '');
  const chk = checkDeployment_(id);
  t.warning = chk.ok ? (chk.warn || '') : '⚠️ ' + chk.message + '\n連結可能無法使用，請先重新執行「活動前 › 初始化」。';
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(620).setHeight(640), '產生連結');
}

function stageKey_() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('STAGE_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('STAGE_KEY', key);
  }
  return key;
}

// ── 清除資料 ─────────────────────────────────────────

function menuClearData() {
  const html = HtmlService.createHtmlOutputFromFile('ClearDialog').setWidth(520).setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(html, '清除資料');
}

// 清除對話框呼叫
function clearDataConfirmed(text, backup) {
  if (String(text || '').trim() !== '確定清除') return { ok: false, message: '輸入的文字不對，沒有清除任何資料。' };
  const t0 = Date.now();
  let backupName = '';
  if (backup) {
    ensureFolders_();
    const folder = DriveApp.getFolderById(sysGet_('備份資料夾ID'));
    const ss = SpreadsheetApp.getActive();
    backupName = ss.getName() + '_備份_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss');
    DriveApp.getFileById(ss.getId()).makeCopy(backupName, folder);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    clearBelowHeader_(SHEET.PEOPLE, PC.TIME, 2);
    clearBelowHeader_(SHEET.PRIZES, 6, 1);
    clearBelowHeader_(SHEET.WINNERS, 1, WINNER_HEADERS.length);
    [SHEET.STAGE, SHEET.SNAPSHOT, SHEET.OPLOG].forEach(n => sheet_(n).clearContents());
    discardPendingCheckins_();
    removeSystemProtections_();
    ['名單檢查碼', '獎項檢查碼', '抽獎池人數', '獎項鎖定後變更', '鎖定時間'].forEach(k => sysSet_(k, ''));
    sysSet_('系統狀態', STATE.PREP);
    SpreadsheetApp.flush();
    invalidateConfig_();
    markPeopleDirty_();
    log_('清除資料', true, Date.now() - t0, backupName ? '已備份：' + backupName : '未備份');
  } catch (err) {
    log_('清除資料', false, Date.now() - t0, '', '', err);
    throw err;
  } finally {
    lock.releaseLock();
  }
  return { ok: true, message: '已清除，系統回到「準備中」。' + (backupName ? '\n備份檔：' + backupName : '') };
}

// ── 模擬全員報到、重新產生前台連結 ────────────────────

function menuSimulateCheckin() {
  const ui = SpreadsheetApp.getUi();
  if (!checkState_(STATE.OPEN, '模擬全員報到')) return;
  const r = ui.alert('模擬全員報到', '會把所有「可否摸彩＝是」、不是「移除」、還沒報到的人填上目前時間。\n只用在彩排與壓測，正式活動不要按。', ui.ButtonSet.OK_CANCEL);
  if (r !== ui.Button.OK) return;
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let n = 0;
  try {
    for (let i = 0; i < 10 && flushLocked_() > 0; i++) { /* 先寫完真的報到 */ }
    const sh = sheet_(SHEET.PEOPLE);
    const last = sh.getLastRow();
    if (last >= 2) {
      const rows = sh.getRange(2, 1, last - 1, PEOPLE_HEADERS.length).getValues();
      const now = new Date();
      const times = rows.map(row => {
        const fill = row[PC.TIME - 1] === '' && String(row[PC.ID - 1]).trim() !== '' &&
          String(row[PC.ELIGIBLE - 1]).trim() === '是' && String(row[PC.REG - 1]).trim() !== '移除';
        if (fill) n++;
        return [fill ? now : row[PC.TIME - 1]];
      });
      sh.getRange(2, PC.TIME, times.length, 1).setValues(times);
    }
    markPeopleDirty_();
    log_('模擬全員報到', true, Date.now() - t0, n + ' 人');
  } finally {
    lock.releaseLock();
  }
  ui.alert('已模擬報到', '填入 ' + n + ' 人的報到時間。', ui.ButtonSet.OK);
}

function menuRegenStageKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('重新產生前台連結？', '舊的前台連結會立刻失效，正在使用的前台要改開新連結。', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().deleteProperty('STAGE_KEY');
  stageKey_();
  log_('重新產生前台連結', true, null, '');
  menuLinks();
}

// ── 資料檢查 ─────────────────────────────────────────

function menuDataCheck() {
  const issues = runChecks_({ people: true, prizes: true });
  writeCheckResults_('資料檢查', issues);
  log_('資料檢查', true, null, issues.length + ' 個項目');
  SpreadsheetApp.getUi().alert('資料檢查完成', checkSummary_(issues, '資料檢查') + '\n\n詳細內容在控制台右側的「資料檢查結果」。', SpreadsheetApp.getUi().ButtonSet.OK);
}

// 回傳問題清單：{ level: 錯誤／提醒, item, detail, blocksLock }
function runChecks_(opts) {
  const issues = [];
  const add = (level, item, detail, blocksLock) => issues.push({ level: level, item: item, detail: detail || '', blocksLock: !!blocksLock });
  const list = (arr, n) => arr.slice(0, n || 10).join('、') + (arr.length > (n || 10) ? ' 等 ' + arr.length + ' 筆' : '');
  const state = currentState_();
  const afterClose = [STATE.CLOSED, STATE.LOCKED, STATE.DRAWING, STATE.ENDED].indexOf(state) >= 0;

  const people = readPeople_();
  const active = people.filter(p => p.reg !== '移除' && p.id);
  const eligibleCount = active.filter(p => p.eligible === '是' && (!afterClose || p.checkedIn)).length;

  if (opts.people) {
    const seen = {};
    const dups = [];
    people.forEach(p => {
      if (!p.id) return;
      if (seen[p.id]) dups.push(p.id + '（第 ' + seen[p.id] + '、' + p.row + ' 列）');
      else seen[p.id] = p.row;
    });
    if (dups.length) add('錯誤', '工號重複', list(dups), true);

    const blank = people.filter(p => !p.unit || !p.id || !p.name || !p.eligible).map(p => '第 ' + p.row + ' 列');
    if (blank.length) add('錯誤', '單位、工號、姓名、可否摸彩有空白', list(blank));
    const badElig = people.filter(p => p.eligible && p.eligible !== '是' && p.eligible !== '否').map(p => '第 ' + p.row + ' 列（' + p.eligible + '）');
    if (badElig.length) add('錯誤', '「可否摸彩」只能填 是／否', list(badElig));
    const noReg = people.filter(p => !p.reg).map(p => '第 ' + p.row + ' 列');
    if (noReg.length) add('提醒', '登錄狀態空白（會當作「既有」）', list(noReg));
    const badReg = people.filter(p => p.reg && ['既有', '現場新增', '移除'].indexOf(p.reg) < 0).map(p => '第 ' + p.row + ' 列（' + p.reg + '）');
    if (badReg.length) add('錯誤', '登錄狀態只能填 既有／現場新增／移除', list(badReg));

    const numeric = people.filter(p => p.idIsNumber).map(p => '第 ' + p.row + ' 列');
    if (numeric.length) add('提醒', '工號是數字格式，前導 0 可能已經遺失', list(numeric) + '。請把工號欄設成純文字後重新貼上。');
    const lens = {};
    people.filter(p => /^\d+$/.test(p.id)).forEach(p => { lens[p.id.length] = (lens[p.id.length] || 0) + 1; });
    const common = Object.keys(lens).sort((a, b) => lens[b] - lens[a])[0];
    if (common) {
      const short = people.filter(p => /^\d+$/.test(p.id) && p.id.length < Number(common)).map(p => p.id + '（第 ' + p.row + ' 列）');
      if (short.length) add('提醒', '工號比多數人短，可能少了前導 0', '多數是 ' + common + ' 碼：' + list(short));
    }

    const byName = {};
    active.forEach(p => { const k = p.unit + '|' + p.name; (byName[k] = byName[k] || []).push(p); });
    const same = Object.keys(byName).filter(k => byName[k].length > 1).map(k => k.replace('|', '_') + '（' + byName[k].map(p => p.id).join('、') + '）');
    if (same.length) add('提醒', '同單位同名同姓', list(same) + '。前台會自動加上工號後四碼區分。');

    const rare = active.filter(p => /[㐀-䶿豈-﫿]|[\uD840-\uD87F][\uDC00-\uDFFF]/.test(p.name)).map(p => p.name + '（' + p.id + '）');
    if (rare.length) add('提醒', '罕用字（字型可能沒有收錄）', list(rare) + '。請在前台預覽確認顯示正常。');
  }

  if (opts.prizes) {
    const prizes = readPrizes_();
    if (!prizes.length) add('錯誤', '獎項分頁沒有任何獎項', '', true);
    const bad = [];
    const sorts = {};
    prizes.forEach(p => {
      const where = '第 ' + p.row + ' 列' + (p.name ? '（' + p.name + '）' : '');
      if (!p.name) bad.push(where + '：獎項名稱空白');
      if (!(p.qty >= 1 && p.qty <= 10 && Math.floor(p.qty) === p.qty)) bad.push(where + '：數量要是 1～10 的整數');
      if (PRIZE_MODES.indexOf(p.mode) < 0) bad.push(where + '：抽法要選 一次抽完／逐個抽出');
      if (typeof p.sort !== 'number') bad.push(where + '：排序要填數字');
      else if (sorts[p.sort]) bad.push(where + '：排序 ' + p.sort + ' 跟第 ' + sorts[p.sort] + ' 列重複');
      else sorts[p.sort] = p.row;
    });
    if (bad.length) add('錯誤', '獎項資料有誤', list(bad, 20), true);

    const total = prizes.reduce((a, p) => a + (p.qty > 0 ? p.qty : 0), 0);
    if (total > eligibleCount) {
      add('錯誤', '獎項總份數超過' + (afterClose ? '抽獎池' : '有抽獎資格的') + '人數',
        '獎項共 ' + total + ' 份，' + (afterClose ? '已報到且有資格的' : '有資格的') + '只有 ' + eligibleCount + ' 人', true);
    }

    const folderId = sysGet_('素材資料夾ID');
    const withPhoto = prizes.filter(p => p.photo);
    if (withPhoto.length && !folderId) add('提醒', '還沒初始化，無法檢查獎品照片', '請先執行「活動前 › 初始化」');
    if (withPhoto.length && folderId) {
      const files = [];
      const it = DriveApp.getFolderById(folderId).getFiles();
      while (it.hasNext()) files.push(it.next().getName());
      const have = {};
      files.forEach(f => { have[photoKey_(f)] = true; });
      const missing = withPhoto.filter(p => !have[photoKey_(p.photo)]).map(p => p.name + '：' + p.photo);
      if (missing.length) add('提醒', '找不到獎品照片（會改用只有文字的版面）', list(missing) + '。素材資料夾裡實際的檔名：' + (files.length ? list(files, 30) : '（資料夾是空的）'));
    }
  }
  return issues;
}

function photoKey_(name) {
  return String(name || '').trim().replace(/\.[^.]+$/, '').toLowerCase();
}

function readPeople_() {
  const sh = sheet_(SHEET.PEOPLE);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const range = sh.getRange(2, 1, last - 1, PEOPLE_HEADERS.length);
  const values = range.getValues();
  const shown = range.getDisplayValues();
  return shown.map((r, i) => ({
    row: i + 2,
    unit: String(r[PC.UNIT - 1]).trim(),
    id: normId(r[PC.ID - 1]),
    idIsNumber: typeof values[i][PC.ID - 1] === 'number',
    name: String(r[PC.NAME - 1]).trim(),
    eligible: String(r[PC.ELIGIBLE - 1]).trim(),
    reg: String(r[PC.REG - 1]).trim(),
    checkedIn: String(r[PC.TIME - 1]).trim() !== '',
  })).filter(p => p.unit || p.id || p.name || p.eligible || p.reg);
}

function readPrizes_() {
  const sh = sheet_(SHEET.PRIZES);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, PRIZE_HEADERS.length).getValues()
    .map((r, i) => ({
      row: i + 2,
      sort: r[0] === '' ? '' : (typeof r[0] === 'number' ? r[0] : (isNaN(Number(r[0])) ? r[0] : Number(r[0]))),
      name: String(r[1]).trim(),
      qty: Number(r[2]),
      photo: String(r[3]).trim(),
      mode: String(r[4]).trim(),
      status: String(r[5]).trim(),
    }))
    .filter(p => p.sort !== '' || p.name || p.photo || p.mode);
}

// 抽獎池：已報到、可否摸彩＝是、不是移除；同單位同名同姓加工號後四碼
function buildPool_() {
  const pool = readPeople_().filter(p => p.id && p.checkedIn && p.eligible === '是' && p.reg !== '移除');
  const count = {};
  pool.forEach(p => { const k = p.unit + '|' + p.name; count[k] = (count[k] || 0) + 1; });
  pool.forEach(p => {
    p.display = p.unit + '_' + p.name + (count[p.unit + '|' + p.name] > 1 ? '（' + p.id.slice(-4) + '）' : '');
  });
  return pool;
}

function poolHash_(pool) {
  return hash_(pool.map(p => p.id).sort().join(','), 4);
}

// 依排序計算，跟獎項列在分頁上的先後無關
function prizeHash_(prizes) {
  const sorted = prizes.slice().sort((a, b) => (a.sort > b.sort ? 1 : a.sort < b.sort ? -1 : 0));
  return hash_(JSON.stringify(sorted.map(p => [p.sort, p.name, p.qty, photoKey_(p.photo), p.mode])), 4);
}

function writeCheckResults_(title, issues) {
  const sh = sheet_(SHEET.CONSOLE);
  const last = Math.max(sh.getLastRow(), 5);
  sh.getRange(5, 7, last - 4, 3).clearContent();
  sh.getRange('G2').setValue(title + '｜' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MM/dd HH:mm') + '｜' + checkSummary_(issues, '').trim());
  const rows = issues.length ? issues.map(x => [x.level, x.item, x.detail]) : [['—', '沒有發現問題', '']];
  sh.getRange(5, 7, rows.length, 3).setValues(rows).setWrap(true).setVerticalAlignment('top');
}

function checkSummary_(issues, label) {
  const err = issues.filter(x => x.level === '錯誤').length;
  const warn = issues.filter(x => x.level === '提醒').length;
  if (!err && !warn) return (label ? label + '：' : '') + '沒有發現問題';
  return (label ? label + '：' : '') + err + ' 個錯誤、' + warn + ' 個提醒';
}

function clearBelowHeader_(name, col, numCols) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, col, last - 1, numCols).clearContent();
}
