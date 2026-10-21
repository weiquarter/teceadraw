/** 活動前與進階：建立範本結構、初始化（資料夾＋部署檢查）、產生連結、清除資料。 */

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

// 在 Apps Script 編輯器手動執行一次，建立範本 Sheet 的所有分頁；已存在的分頁不會動
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
  // 控制台公式引用其他分頁，要最後建，否則公式會指向還不存在的分頁而算錯
  make(SHEET.CONSOLE, null, buildConsole_);

  ['工作表1', 'Sheet1'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(ss.getSheetByName(SHEET.CONSOLE));
  ss.moveActiveSheet(1);
  Logger.log(created.length ? '已建立：' + created.join('、') : '所有分頁都已存在，沒有變更');
}

function buildConsole_(sh) {
  const st = "IFERROR(VLOOKUP(\"系統狀態\",'" + SHEET.SYSTEM + "'!A:B,2,FALSE),\"準備中\")";
  const nextStep = 'SWITCH(B3,' + Object.keys(NEXT_STEP).map(k => '"' + k + '","' + NEXT_STEP[k] + '"').join(',') + ',"")';
  const live = "'" + SHEET.PEOPLE + "'!";
  const rows = [
    ['摸彩系統 控制台', ''],
    ['', ''],
    ['目前狀態', '=' + st],
    ['下一步', '=' + nextStep],
    ['', ''],
    ['報到統計', ''],
    ['已報到', '=COUNTIFS(' + live + 'H2:H,"<>",' + live + 'D2:D,"<>",' + live + 'G2:G,"<>移除")'],
    ['總人數', '=COUNTIFS(' + live + 'D2:D,"<>",' + live + 'G2:G,"<>移除")'],
    ['有抽獎資格且已報到', '=COUNTIFS(' + live + 'H2:H,"<>",' + live + 'D2:D,"<>",' + live + 'F2:F,"是",' + live + 'G2:G,"<>移除")'],
    ['', ''],
    ['錯誤監看', ''],
    ['最近 10 分鐘 GAS 錯誤', '=COUNTIFS(\'' + SHEET.LOG + '\'!A2:A,">"&(NOW()-10/1440),\'' + SHEET.LOG + '\'!C2:C,"失敗")'],
    ['', ''],
    ['提示', '第一次使用「摸彩系統」選單時會要求授權，屬正常現象。'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange('A1').setFontSize(18).setFontWeight('bold');
  sh.getRange('A3:B4').setFontSize(20).setFontWeight('bold');
  sh.getRange('A6').setFontWeight('bold');
  sh.getRange('A11').setFontWeight('bold');
  sh.getRange('B7:B9').setFontSize(16).setHorizontalAlignment('left');
  sh.getRange('B12').setHorizontalAlignment('left');
  sh.getRange('B12').setFontColor('#B5481C');
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 420);
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

// 初始化對話框呼叫：以匿名方式打部署網址，確認存取權、對應的 Sheet 與時區
function verifyDeployment(input) {
  const m = String(input || '').trim().match(/\/s\/([\w-]+)\//) || String(input || '').trim().match(/^([\w-]{30,})$/);
  if (!m) return { ok: false, message: '看不懂貼上的內容。請貼「部署作業 ID」（AKfycb 開頭的一長串），或網頁應用程式網址。' };
  const id = m[1];
  const url = 'https://script.google.com/macros/s/' + id + '/exec?a=ping';
  let res;
  try {
    res = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true });
  } catch (err) {
    return { ok: false, message: '連不到部署網址：' + err.message };
  }
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code === 404) return { ok: false, message: '找不到這個部署。請確認貼的是「網頁應用程式」的部署作業 ID，而不是程式庫或專案 ID。' };
  let data = null;
  try { data = JSON.parse(text); } catch (e) { data = null; }
  if (!data) {
    if (/accounts\.google\.com|ServiceLogin|signin/i.test(text)) {
      return { ok: false, message: '匿名存取被導去登入頁。請到「管理部署作業」把「誰可以存取」改成「任何人」，執行身分選「我」。如果選單裡沒有「任何人」，代表 Workspace 管理員關閉了這個選項，需要請管理員開啟。' };
    }
    return { ok: false, message: '部署網址沒有回傳預期的資料（HTTP ' + code + '）。請確認部署的是這個專案的最新版本。' };
  }
  if (!data.ok) return { ok: false, message: '部署的程式版本太舊，沒有檢查用的功能。請「管理部署作業 › 編輯 › 版本：新版本」後再試。' };
  if (data.sheet !== sheetHash_()) return { ok: false, message: '這個部署屬於另一份 Sheet。請在這份 Sheet 的「擴充功能 › Apps Script」裡部署，再貼上那個部署作業 ID。' };
  if (data.apiVersion !== API_VERSION) return { ok: false, message: '部署的程式版本跟目前程式碼不一致，請部署新版本。' };
  sysSet_('部署ID', id);
  log_('初始化', true, null, id);
  const warn = data.timeZone !== 'Asia/Taipei' ? '\n⚠️ 專案時區是 ' + data.timeZone + '，請到「專案設定」改成 (GMT+08:00) 台北。' : '';
  return { ok: true, message: '檢查通過：匿名存取正常，部署屬於這份 Sheet。' + warn };
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

// 回傳是否真的清除了（給「開始報到」判斷要不要繼續）
function menuClearData() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('清除資料',
    '會清除：報到時間、中獎獎項、得獎名單、前台狀態、名單快照、操作紀錄，系統回到「準備中」，前台需要重新開啟。\n' +
    '不會清除：人員名單、獎項、設定、Log。\n\n請輸入「確定清除」：', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return false;
  if (r.getResponseText().trim() !== '確定清除') {
    ui.alert('輸入的文字不對，沒有清除任何資料。');
    return false;
  }
  const b = ui.alert('清除前先備份？',
    '按「是」會先把整份 Sheet 複製一份到「備份」資料夾（正式活動建議備份）。\n彩排時可以按「否」，避免備份堆積。', ui.ButtonSet.YES_NO_CANCEL);
  if (b !== ui.Button.YES && b !== ui.Button.NO) return false;

  const t0 = Date.now();
  let backupName = '';
  if (b === ui.Button.YES) {
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
    removeSystemProtections_();
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
  ui.alert('已清除', '系統已回到「準備中」。' + (backupName ? '\n備份檔：' + backupName : ''), ui.ButtonSet.OK);
  return true;
}

function clearBelowHeader_(name, col, numCols) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, col, last - 1, numCols).clearContent();
}
