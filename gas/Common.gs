/** 共用：常數、系統狀態、設定讀取、Log、工號正規化與姓名遮蔽。 */

const API_VERSION = 1;
// 每次推送新程式就加一，用來判斷部署是不是最新版本
const BUILD = 12;

const SHEET = {
  CONSOLE: '控制台',
  PEOPLE: '人員',
  PRIZES: '獎項',
  WINNERS: '得獎名單',
  SETTINGS: '設定',
  LOG: 'Log',
  SYSTEM: '_系統',
  STAGE: '_前台狀態',
  SNAPSHOT: '_名單快照',
  OPLOG: '_操作紀錄',
};

const STATE = {
  PREP: '準備中',
  OPEN: '報到中',
  CLOSED: '已截止',
  LOCKED: '已鎖定',
  DRAWING: '抽獎中',
  ENDED: '已結束',
};

const NEXT_STEP = {
  '準備中': '活動當天 › 開始報到',
  '報到中': '活動當天 › 截止報到',
  '已截止': '活動當天 › 鎖定名單',
  '已鎖定': '前台開始抽獎',
  '抽獎中': '全部抽完後：活動當天 › 結束活動',
  '已結束': '（活動已結束）',
};

const PEOPLE_HEADERS = ['縣市', '部門', '單位', '工號', '姓名', '可否摸彩', '登錄狀態', '報到時間', '中獎獎項'];
const PC = { UNIT: 3, ID: 4, NAME: 5, ELIGIBLE: 6, REG: 7, TIME: 8, PRIZE: 9 };

const PRIZE_HEADERS = ['排序', '獎項名稱', '數量', '照片檔名', '抽法', '狀態'];
const WINNER_HEADERS = ['時間', '獎項', '名額序號', '單位', '工號', '姓名', '狀態', '備註'];
const LOG_HEADERS = ['時間', 'API／動作', '結果', '耗時（毫秒）', '對象', '前台代號', '錯誤訊息'];

const SNAPSHOT_HEADERS = ['工號', '單位', '姓名', '顯示名稱'];
const STAGE_STATUS_KEYS = ['前台代號', '目前獎項', '目前頁面', '待同步筆數', '最後回報時間'];
const PRIZE_MODES = ['一次抽完', '逐個抽出'];

const PALETTE_IDS = { '草莓牛奶': 'strawberry', '薄荷檸檬': 'mint', '薰衣草藍莓': 'lavender' };

const TEXT_KEYS = {
  success: '提示_報到成功',
  eligible: '提示_具抽獎資格',
  ineligible: '提示_無抽獎資格',
  closed: '提示_報到已截止',
  notFound: '提示_查無工號',
  prep: '提示_系統準備中',
  ended: '提示_活動已結束',
  screenshot: '提示_請截圖',
};

const DEFAULT_PAGES_BASE = 'https://weiquarter.github.io/teceadraw';

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('找不到分頁「' + name + '」');
  return sh;
}

// 讀取「項目／值」兩欄式的分頁（設定、_系統）
function readKeyValue_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  const out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
    if (r[0] !== '') out[String(r[0]).trim()] = r[1];
  });
  return out;
}

function sysGet_(key) {
  return readKeyValue_(SHEET.SYSTEM)[key];
}

function sysSet_(key, value) {
  const sh = sheet_(SHEET.SYSTEM);
  const last = sh.getLastRow();
  const keys = last ? sh.getRange(1, 1, last, 1).getValues().map(r => String(r[0]).trim()) : [];
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 1, 2).setValue(value);
  else sh.appendRow([key, value]);
}

function currentState_() {
  return String(sysGet_('系統狀態') || STATE.PREP);
}

// 公開 API 用的設定，放 CacheService；設定分頁被編輯或狀態改變時清掉
function getConfig_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('cfg');
  if (hit) return JSON.parse(hit);
  const cfg = buildConfig_();
  cache.put('cfg', JSON.stringify(cfg), 600);
  return cfg;
}

function buildConfig_() {
  const s = readKeyValue_(SHEET.SETTINGS);
  const str = k => (s[k] == null ? '' : String(s[k]).trim());
  const hex = v => (/^#[0-9a-fA-F]{6}$/.test(v) ? v : '');
  const texts = {};
  Object.keys(TEXT_KEYS).forEach(k => { texts[k] = str(TEXT_KEYS[k]); });
  return {
    state: currentState_(),
    title: str('活動名稱'),
    subtitle: str('副標題'),
    theme: {
      palette: PALETTE_IDS[str('配色主題')] || 'strawberry',
      colors: {
        primary: hex(str('主色')),
        secondary: hex(str('輔色')),
        bg: hex(str('背景色')),
        ink: hex(str('文字色')),
      },
      font: str('字型') || 'Huninn',
      reel: str('拉霸樣式').charAt(0) || 'A',
    },
    texts: texts,
  };
}

function invalidateConfig_() {
  CacheService.getScriptCache().remove('cfg');
}

function pagesBase_() {
  const v = String(readKeyValue_(SHEET.SETTINGS)['前端網址'] || '').trim();
  return (v || DEFAULT_PAGES_BASE).replace(/\/+$/, '');
}

function log_(action, ok, ms, target, stageCode, err) {
  try {
    sheet_(SHEET.LOG).appendRow([
      new Date(), action, ok ? '成功' : '失敗', ms == null ? '' : ms,
      target == null ? '' : String(target), stageCode || '',
      err ? String(err.message || err) : '',
    ]);
  } catch (e) {
    console.error('寫入 Log 失敗', e);
  }
}

// 去空白、全形轉半形、英文轉大寫
function normId(s) {
  return String(s == null ? '' : s)
    .replace(/[！-～]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[\s　]/g, '')
    .toUpperCase();
}

// 王小明 → 王○明；王明 → 王○；歐陽靜文 → 歐○○文
function maskName(name) {
  const chars = Array.from(String(name || '').trim());
  if (chars.length <= 1) return chars.join('');
  if (chars.length === 2) return chars[0] + '○';
  return chars[0] + '○'.repeat(chars.length - 2) + chars[chars.length - 1];
}

// 從報到時間的顯示文字取出 HH:mm
function hhmm_(display) {
  const m = String(display || '').match(/(\d{1,2}):(\d{2})/);
  return m ? (m[1].length === 1 ? '0' + m[1] : m[1]) + ':' + m[2] : '';
}

function sheetHash_() {
  return hash_(SpreadsheetApp.getActive().getId(), 6);
}

// SHA-256 取前 n 個位元組的十六進位
function hash_(text, n) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return d.slice(0, n).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function include_(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}
