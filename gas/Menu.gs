/** 「摸彩系統」選單、編輯時清快取，以及「活動當天」的狀態切換動作。 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('摸彩系統')
    .addSubMenu(ui.createMenu('活動前')
      .addItem('初始化', 'menuInit')
      .addItem('產生連結', 'menuLinks'))
    .addSubMenu(ui.createMenu('活動當天')
      .addItem('開始報到', 'menuStartCheckin')
      .addItem('截止報到', 'menuCloseCheckin'))
    .addSubMenu(ui.createMenu('進階')
      .addItem('清除資料', 'menuClearData'))
    .addToUi();
}

// 直接在 Sheet 打字時，讓 API 的快取知道要重讀
function onEdit(e) {
  const name = e && e.range && e.range.getSheet().getName();
  if (name === SHEET.PEOPLE) markPeopleDirty_();
  else if (name === SHEET.SETTINGS) invalidateConfig_();
}

function menuStartCheckin() {
  const ui = SpreadsheetApp.getUi();
  if (!sysGet_('部署ID')) {
    ui.alert('還沒完成部署設定', '請先由系統管理者執行「摸彩系統 › 活動前 › 初始化」。', ui.ButtonSet.OK);
    return;
  }
  if (!checkState_(STATE.PREP, '開始報到')) return;
  if (hasLeftoverData_()) {
    const r = ui.alert('偵測到之前留下的資料',
      '人員分頁已有報到時間，或得獎名單已有資料（可能是彩排留下的）。\n要先清除嗎？\n\n按「是」進入清除流程；按「否」保留資料直接開始報到。',
      ui.ButtonSet.YES_NO_CANCEL);
    if (r === ui.Button.YES) { if (!menuClearData()) return; }
    else if (r !== ui.Button.NO) return;
  }
  const ok = transition_(STATE.PREP, STATE.OPEN, '開始報到', () => {
    protectWarning_(SHEET.PEOPLE, '摸彩系統：報到中，請勿刪列或排序');
  });
  if (!ok) return;
  rebuildMap_();
  ui.alert('已開始報到', '報到頁已開放。\n提醒：人員分頁不刪列、不排序，新增一律加在最下面。', ui.ButtonSet.OK);
}

function menuCloseCheckin() {
  if (!checkState_(STATE.OPEN, '截止報到')) return;
  if (!transition_(STATE.OPEN, STATE.CLOSED, '截止報到')) return;
  SpreadsheetApp.getUi().alert('已截止報到', '已報到的人仍可查詢自己的報到狀態，未報到的人會看到「報到已截止」。\n下一步：' + NEXT_STEP[STATE.CLOSED], SpreadsheetApp.getUi().ButtonSet.OK);
}

// 狀態不對時用白話說明下一步，回傳是否可以繼續
function checkState_(expected, actionName) {
  const state = currentState_();
  if (state === expected) return true;
  SpreadsheetApp.getUi().alert('現在不能「' + actionName + '」',
    '目前狀態是「' + state + '」。\n下一步請按：' + NEXT_STEP[state], SpreadsheetApp.getUi().ButtonSet.OK);
  return false;
}

function transition_(from, to, actionName, extra) {
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let changed = '';
  try {
    if (currentState_() !== from) {
      changed = currentState_();
      return false;
    }
    if (extra) extra();
    sysSet_('系統狀態', to);
    SpreadsheetApp.flush();
    invalidateConfig_();
    log_(actionName, true, Date.now() - t0, from + ' → ' + to);
    return true;
  } catch (err) {
    log_(actionName, false, Date.now() - t0, '', '', err);
    throw err;
  } finally {
    lock.releaseLock();
    if (changed) SpreadsheetApp.getUi().alert('狀態已經被別人改成「' + changed + '」，這次沒有執行。');
  }
}

function protectWarning_(sheetName, description) {
  const sh = sheet_(sheetName);
  const exists = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).some(p => p.getDescription().indexOf('摸彩系統') === 0);
  if (!exists) sh.protect().setWarningOnly(true).setDescription(description);
}

function removeSystemProtections_() {
  [SHEET.PEOPLE, SHEET.PRIZES].forEach(name => {
    sheet_(name).getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .filter(p => p.getDescription().indexOf('摸彩系統') === 0)
      .forEach(p => p.remove());
  });
}

function hasLeftoverData_() {
  const people = sheet_(SHEET.PEOPLE);
  const last = people.getLastRow();
  if (last >= 2 && people.getRange(2, PC.TIME, last - 1, 1).getValues().some(r => r[0] !== '')) return true;
  return sheet_(SHEET.WINNERS).getLastRow() >= 2;
}
