/**
 * Wrong Clicker - a clicker game that runs inside a real Google Sheet!!!!
 * Data lives in the "Players" sheet (one row per player) and the
 * "Leaderboard" sheet is just a QUERY formula over it, so boring(uhgnnn) Sheets
 * formulas, charts and formatting all work!!!
 * Passwords are hashed!! and kept in Script Properties (not in any cell).
 */

const PLAYERS = 'Players';
const BOARD = 'Leaderboard';
const START = { score: 0, mult: 1, cost: 15, earned: 0, auto: 0, autoCost: 100 };

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🕹️ Wrong Clicker')
    .addItem('Open game', 'openGame')
    .addSeparator()
    .addItem('Set up sheets (run once)', 'setup')
    .addItem('Reset ALL players', 'resetAll')
    .addToUi();
}

function openGame() {
  setup();
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Wrong Clicker');
  SpreadsheetApp.getUi().showSidebar(html);
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  let p = ss.getSheetByName(PLAYERS);
  if (!p) {
    p = ss.insertSheet(PLAYERS);
    p.getRange(1, 1, 1, 8).setValues([['Player', 'Score', 'Per click', 'Upgrade cost', 'Total earned', 'Auto per sec', 'Auto cost', 'Last seen']])
      .setFontWeight('bold').setBackground('#e8f0fe');
    p.setFrozenRows(1);
  }
  let b = ss.getSheetByName(BOARD);
  if (!b) {
    b = ss.insertSheet(BOARD);
    b.getRange('A1').setValue('🏆 Wrong Clicker Leaderboard').setFontWeight('bold').setFontSize(14);
    b.getRange('A3').setFormula(
      '=QUERY(Players!A1:E,"select A, E where A is not null order by E desc limit 10 label A \'Player\', E \'Total earned\'",1)');
    b.getRange('D3').setValue('Everyone has clicked').setFontWeight('bold');
    b.getRange('D4').setFormula('=SUM(Players!E2:E)');
    b.getRange('A3:B3').setFontWeight('bold').setBackground('#e8f0fe');
    b.setColumnWidth(1, 160);
    b.setColumnWidth(4, 160);
  }
}

function resetAll() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Delete ALL players and scores?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().deleteAllProperties();
  const p = SpreadsheetApp.getActive().getSheetByName(PLAYERS);
  if (p && p.getLastRow() > 1) p.getRange(2, 1, p.getLastRow() - 1, 8).clearContent();
}

function salt_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('salt');
  if (!s) { s = Utilities.getUuid(); props.setProperty('salt', s); }
  return s;
}

function hash_(name, pw) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    'wt:' + name.toLowerCase() + ':' + pw + ':' + salt_());
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function checkName_(name, pw) {
  if (!/^[A-Za-z0-9_]{2,16}$/.test(name || '')) throw new Error('Username: 2-16 letters, numbers or _');
  if (!pw || pw.length < 4) throw new Error('Password must be at least 4 characters');
}

function register(name, pw) {
  checkName_(name, pw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const key = name.toLowerCase();
    if (props.getProperty('u_' + key)) throw new Error('That username is taken');
    props.setProperty('u_' + key, JSON.stringify({ name: name, hash: hash_(name, pw) }));
    const s = Object.assign({ t: Date.now() }, START);
    props.setProperty('g_' + key, JSON.stringify(s));
    mirror_(name, s);
  } finally { lock.releaseLock(); }
  return login(name, pw);
}

function login(name, pw) {
  checkName_(name, pw);
  const props = PropertiesService.getScriptProperties();
  const key = name.toLowerCase();
  const u = JSON.parse(props.getProperty('u_' + key) || 'null');
  if (!u) throw new Error('No such account. Use Create account.');
  if (u.hash !== hash_(u.name, pw)) throw new Error('Wrong password');
  const token = Utilities.getUuid();
  props.setProperty('tok_' + key, token);
  return { name: u.name, token: token, state: act_(u.name, token, function () {}) };
}

function authed_(name, token) {
  const ok = name && token &&
    PropertiesService.getScriptProperties().getProperty('tok_' + name.toLowerCase()) === token;
  if (!ok) throw new Error('Please sign in again.');
}


function load_(key) {
  const s = JSON.parse(PropertiesService.getScriptProperties().getProperty('g_' + key));
  const secs = Math.min(Math.floor((Date.now() - s.t) / 1000), 3600);
  if (secs > 0) {
    const gain = s.auto * secs;
    s.score += gain; s.earned += gain; s.t += secs * 1000;
  }
  return s;
}

function act_(name, token, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    authed_(name, token);
    const key = name.toLowerCase();
    const s = load_(key);
    fn(s);
    const doMirror = !s.m || Date.now() - s.m > 8000;
    if (doMirror) s.m = Date.now();
    PropertiesService.getScriptProperties().setProperty('g_' + key, JSON.stringify(s));
    if (doMirror) mirror_(name, s);
    return { score: s.score, mult: s.mult, cost: s.cost, earned: s.earned, auto: s.auto, autoCost: s.autoCost };
  } finally { lock.releaseLock(); }
}

function click(name, token, n) {
  n = Math.max(0, Math.min(Math.floor(Number(n) || 0), 60));
  return act_(name, token, function (s) { const g = s.mult * n; s.score += g; s.earned += g; });
}

function buy(name, token, what) {
  return act_(name, token, function (s) {
    if (what === 'mult') {
      if (s.score < s.cost) throw new Error('Not enough score');
      s.score -= s.cost; s.mult += 1; s.cost = Math.round(s.cost * 1.6);
    } else if (what === 'auto') {
      if (s.score < s.autoCost) throw new Error('Not enough score');
      s.score -= s.autoCost; s.auto += 1; s.autoCost = Math.round(s.autoCost * 1.8);
    }
  });
}

function syncAll(name, token, ops) {
  ops = (ops || []).slice(0, 100);
  return act_(name, token, function (s) {
    ops.forEach(function (o) {
      if (o[0] === 'c') {
        const n = Math.max(0, Math.min(Math.floor(Number(o[1]) || 0), 200));
        const g = s.mult * n; s.score += g; s.earned += g;
      } else if (o[0] === 'mult' && s.score >= s.cost) {
        s.score -= s.cost; s.mult += 1; s.cost = Math.round(s.cost * 1.6);
      } else if (o[0] === 'auto' && s.score >= s.autoCost) {
        s.score -= s.autoCost; s.auto += 1; s.autoCost = Math.round(s.autoCost * 1.8);
      }
    });
  });
}

function getBoard() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const rows = [];
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('g_') !== 0) return;
    const u = JSON.parse(all['u_' + k.slice(2)] || 'null');
    if (u) rows.push({ name: u.name, earned: JSON.parse(all[k]).earned });
  });
  rows.sort(function (a, b) { return b.earned - a.earned; });
  return rows.slice(0, 10);
}

function mirror_(name, s) {
  const sh = SpreadsheetApp.getActive().getSheetByName(PLAYERS);
  if (!sh) return;
  const last = sh.getLastRow();
  const names = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  const i = names.indexOf(name);
  const row = i >= 0 ? i + 2 : last + 1;
  sh.getRange(row, 1, 1, 8).setValues([[name, s.score, s.mult, s.cost, s.earned, s.auto, s.autoCost, new Date()]]);
}
