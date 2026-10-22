/** 「摸彩系統」選單、編輯時清快取，以及「活動當天」的動作：開始／截止報到、鎖定／解除、整理得獎名單、結束活動。 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('摸彩系統')
    .addSubMenu(ui.createMenu('活動前')
      .addItem('初始化', 'menuInit')
      .addItem('產生連結', 'menuLinks')
      .addItem('資料檢查', 'menuDataCheck'))
    .addSubMenu(ui.createMenu('活動當天')
      .addItem('開始報到', 'menuStartCheckin')
      .addItem('截止報到', 'menuCloseCheckin')
      .addItem('鎖定名單', 'menuLock')
      .addItem('解除鎖定', 'menuUnlock')
      .addSeparator()
      .addItem('強制前台回主頁', 'menuForceHome')
      .addItem('強制切換前台', 'menuForceSwitch')
      .addSeparator()
      .addItem('整理得獎名單', 'menuSortWinners')
      .addItem('結束活動', 'menuEnd'))
    .addSubMenu(ui.createMenu('進階')
      .addItem('模擬全員報到', 'menuSimulateCheckin')
      .addItem('清除資料', 'menuClearData')
      .addItem('重新產生前台連結', 'menuRegenStageKey'))
    .addToUi();
}

// 直接在 Sheet 打字時，讓 API 的快取知道要重讀；鎖定後改獎項要標記
function onEdit(e) {
  const name = e && e.range && e.range.getSheet().getName();
  if (name === SHEET.PEOPLE) markPeopleDirty_();
  else if (name === SHEET.SETTINGS) invalidateConfig_();
  else if (name === SHEET.PRIZES) {
    const state = currentState_();
    // 跟鎖定時的檢查碼比對，改回原樣就會解除警告
    if (state === STATE.LOCKED || state === STATE.DRAWING) {
      sysSet_('獎項鎖定後變更', prizeHash_(readPrizes_()) === String(sysGet_('獎項檢查碼')) ? '' : '是');
    }
  }
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
      '人員分頁已有報到時間，或得獎名單已有資料（可能是彩排留下的）。\n要先清除嗎？\n\n' +
      '按「是」打開清除畫面，清除完成後請再按一次「開始報到」。\n按「否」保留資料，直接開始報到。',
      ui.ButtonSet.YES_NO_CANCEL);
    if (r === ui.Button.YES) { menuClearData(); return; }
    if (r !== ui.Button.NO) return;
  }
  const ok = transition_(STATE.PREP, STATE.OPEN, '開始報到', () => {
    protectWarning_(SHEET.PEOPLE, '摸彩系統：報到中，請勿刪列或排序');
  });
  if (!ok) return;
  rebuildMap_();
  ui.alert('已開始報到', '報到頁已開放。\n提醒：人員分頁不刪列、不排序，新增一律加在最下面。', ui.ButtonSet.OK);
}

function menuCloseCheckin() {
  const ui = SpreadsheetApp.getUi();
  if (!checkState_(STATE.OPEN, '截止報到')) return;
  if (!transition_(STATE.OPEN, STATE.CLOSED, '截止報到', () => {
    // 先把還沒寫進 Sheet 的報到全部寫完，再切狀態
    for (let i = 0; i < 10 && flushLocked_() > 0; i++) { /* 直到沒有待寫入 */ }
  })) return;
  const issues = runChecks_({ people: false, prizes: true });
  writeCheckResults_('截止報到時的獎項檢查', issues);
  ui.alert('已截止報到',
    '已報到的人仍可查詢自己的報到狀態，未報到的人會看到「報到已截止」。\n\n' +
    checkSummary_(issues, '獎項檢查') + '\n\n下一步：' + NEXT_STEP[STATE.CLOSED], ui.ButtonSet.OK);
}

function menuLock() {
  const ui = SpreadsheetApp.getUi();
  if (!checkState_(STATE.CLOSED, '鎖定名單')) return;
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  try {
    result = lockLocked_();
  } catch (err) {
    log_('鎖定名單', false, Date.now() - t0, '', '', err);
    throw err;
  } finally {
    lock.releaseLock();
  }
  writeCheckResults_('鎖定名單時的檢查', result.issues);
  if (!result.ok) {
    log_('鎖定名單', false, Date.now() - t0, '', '', result.reason);
    ui.alert('沒有鎖定', result.reason + '\n\n詳細內容在控制台右側的「資料檢查結果」。修正後再按一次「鎖定名單」。', ui.ButtonSet.OK);
    return;
  }
  log_('鎖定名單', true, Date.now() - t0, '版本 ' + result.version + '，抽獎池 ' + result.pool + ' 人');
  ui.alert('已鎖定名單',
    '名單版本：' + result.version + '\n抽獎池：' + result.pool + ' 人\n獎項：' + result.prizeCount + ' 項，共 ' + result.qty + ' 份\n\n' +
    '獎項與人員分頁已加上修改警告。前台請重新整理，確認畫面上的版本號跟控制台一致。', ui.ButtonSet.OK);
}

// 呼叫前要持有 Script Lock
function lockLocked_() {
  if (currentState_() !== STATE.CLOSED) return { ok: false, issues: [], reason: '狀態已經被別人改成「' + currentState_() + '」。' };
  for (let i = 0; i < 10 && flushLocked_() > 0; i++) { /* 寫完截止前最後幾筆 */ }
  const issues = runChecks_({ people: true, prizes: true });
  const blocking = issues.filter(x => x.blocksLock);
  if (blocking.length) {
    return { ok: false, issues: issues, reason: '有 ' + blocking.length + ' 個問題要先處理：\n・' + blocking.map(x => x.item).join('\n・') };
  }

  const pool = buildPool_();
  const prizes = readPrizes_();
  const snap = sheet_(SHEET.SNAPSHOT);
  snap.clearContents();
  snap.getRange(1, 1, 1, SNAPSHOT_HEADERS.length).setValues([SNAPSHOT_HEADERS]);
  snap.getRange('A:A').setNumberFormat('@');
  if (pool.length) snap.getRange(2, 1, pool.length, SNAPSHOT_HEADERS.length).setValues(pool.map(p => [p.id, p.unit, p.name, p.display]));

  const prizeSheet = sheet_(SHEET.PRIZES);
  prizes.forEach(p => prizeSheet.getRange(p.row, PZ.STATUS).setValue('未抽'));

  const version = Number(sysGet_('名單版本') || 0) + 1;
  sysSet_('名單版本', version);
  sysSet_('名單檢查碼', poolHash_(pool));
  sysSet_('獎項檢查碼', prizeHash_(prizes));
  sysSet_('抽獎池人數', pool.length);
  sysSet_('鎖定時間', new Date());
  sysSet_('獎項鎖定後變更', '');
  protectWarning_(SHEET.PRIZES, '摸彩系統：已鎖定名單，修改獎項要先解除鎖定');
  protectWarning_(SHEET.PEOPLE, '摸彩系統：報到中，請勿刪列或排序');
  sysSet_('系統狀態', STATE.LOCKED);
  SpreadsheetApp.flush();
  invalidateConfig_();
  return {
    ok: true, issues: issues, version: version, pool: pool.length,
    prizeCount: prizes.length, qty: prizes.reduce((a, p) => a + p.qty, 0),
  };
}

function menuUnlock() {
  const ui = SpreadsheetApp.getUi();
  const state = currentState_();
  if (state === STATE.DRAWING) {
    ui.alert('不能解除鎖定', '前台已經抽出第一位，名單不能再改。', ui.ButtonSet.OK);
    return;
  }
  if (!checkState_(STATE.LOCKED, '解除鎖定')) return;
  const ok = transition_(STATE.LOCKED, STATE.CLOSED, '解除鎖定', () => {
    sheet_(SHEET.PRIZES).getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .filter(p => p.getDescription().indexOf('摸彩系統') === 0)
      .forEach(p => p.remove());
    sheet_(SHEET.SNAPSHOT).clearContents();
    sysSet_('名單檢查碼', '');
    sysSet_('獎項檢查碼', '');
    sysSet_('抽獎池人數', '');
    sysSet_('獎項鎖定後變更', '');
    clearBelowHeader_(SHEET.PRIZES, PZ.STATUS, 1);
  });
  if (ok) ui.alert('已解除鎖定', '回到「已截止」。可以修改人員與獎項，改完再按「鎖定名單」。', ui.ButtonSet.OK);
}

function menuSortWinners() {
  const t0 = Date.now();
  const sh = sheet_(SHEET.WINNERS);
  const last = sh.getLastRow();
  if (last < 3) {
    SpreadsheetApp.getUi().alert('得獎名單少於兩筆，不需要整理。');
    return;
  }
  const order = {};
  readPrizes_().forEach(p => { order[p.name] = p.sort; });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const range = sh.getRange(2, 1, last - 1, WINNER_HEADERS.length);
    const rows = range.getValues();
    const key = r => [r[1] in order ? order[r[1]] : 9999, r[6] === '作廢' ? 1 : 0, Number(r[2]) || 0];
    rows.sort((a, b) => {
      const ka = key(a), kb = key(b);
      for (let i = 0; i < ka.length; i++) if (ka[i] !== kb[i]) return ka[i] - kb[i];
      return 0;
    });
    range.setValues(rows);
    log_('整理得獎名單', true, Date.now() - t0, rows.length + ' 筆');
  } finally {
    lock.releaseLock();
  }
  SpreadsheetApp.getUi().alert('已整理', '依獎項排序，作廢的排在各獎項最後。可以直接按 Ctrl+P 列印。', SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuEnd() {
  const ui = SpreadsheetApp.getUi();
  if (!checkState_(STATE.DRAWING, '結束活動')) return;
  const r = ui.alert('確定要結束活動？',
    '結束後：報到頁顯示「活動已結束」、所有公開 API 不再回傳資料、前台專屬連結失效。\n這個動作不能復原。', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const ok = transition_(STATE.DRAWING, STATE.ENDED, '結束活動', () => {
    PropertiesService.getScriptProperties().deleteProperty('STAGE_KEY');
  });
  if (ok) ui.alert('活動已結束', '列印前請先按「整理得獎名單」。', ui.ButtonSet.OK);
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
