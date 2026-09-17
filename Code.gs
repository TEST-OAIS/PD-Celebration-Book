/***********************************************************************
 * PD CELEBRATION BOOK — v6.11
 * v6.11 PDF FIX: no blank overflow page after photo pages;
 * contributors without a photo get ONE page (message only).
 * SETUP: run  setup  once.   DIAGNOSE: run  diagnose
 ***********************************************************************/

var PARENT_FOLDER_ID = '1f_pvdei04FWcW9qGySBeEjTeVZYccTwl';

var APP_NAME    = 'PD Celebration Book';
var MAX_MESSAGE = 1200;
var SS_NAME     = 'PD Celebration Book — Data';
var IMG_FOLDER  = 'Celebration Images';
var PDF_FOLDER  = 'Celebration PDFs';

var RETENTION_DAYS = 15;
var WARN_DAYS      = 3;
var MAX_ATTACH_MB  = 22;

var TWEMOJI_BASE = 'https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/';

var R_BLUE = '#0B41CD';
var R_DARK = '#022366';
var R_LITE = '#1482FA';
var R_TINT = '#E8EFFC';
var R_LINE = '#C3D6F7';
var R_INK  = '#131F30';
var R_MUTE = '#5E7291';
var R_GOLD = '#F2B705';

var BOOK_HEADERS = ['bookId','createdAt','ownerEmail','ownerName','deputies',
                    'occasion','honouree','coverMessage','welcomeMessage','theme',
                    'coverImageId','bgStyle','bgImageId','deadline','recipients',
                    'status','sentAt','warnedAt','deletedAt','pdfFileId','thankYouBy'];

var PAGE_HEADERS = ['pageId','bookId','createdAt','updatedAt','authorEmail',
                    'authorName','message','imageId','deleted'];

var THEME_HEX = {
  sunset : { main:'#C2410C', light:'#F97316', soft:'#FFEDD5', paper:'#FFFDF8', ink:'#3F2A1D', line:'#EFD9C4', muted:'#8A6A55' },
  golden : { main:'#A16207', light:'#EAB308', soft:'#FEF3C7', paper:'#FFFEF7', ink:'#3F3117', line:'#EFE2B8', muted:'#8A7440' },
  blossom: { main:'#BE185D', light:'#F472B6', soft:'#FCE7F3', paper:'#FFFAFC', ink:'#4A1D33', line:'#F3D5E3', muted:'#96637C' },
  sage   : { main:'#9A3412', light:'#84A98C', soft:'#EAE4D9', paper:'#FFFDFA', ink:'#3B2F26', line:'#E4D9C9', muted:'#7D6B5A' },
  roche  : { main:R_BLUE,    light:R_LITE,    soft:R_TINT,    paper:'#FCFDFF', ink:R_INK,     line:R_LINE,    muted:R_MUTE }
};
function themeHex_(id) { return THEME_HEX[id] || THEME_HEX.sunset; }


/* =================== SETUP / RESET =================== */

function setup() {
  if (PARENT_FOLDER_ID === 'PASTE_SHARED_DRIVE_FOLDER_ID_HERE' || !PARENT_FOLDER_ID)
    throw new Error('Paste your Shared Drive folder ID into PARENT_FOLDER_ID first.');

  var parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  var props  = PropertiesService.getScriptProperties();

  var ssId = props.getProperty('SS_ID'), ss = null;
  if (ssId) { try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(SS_NAME);
    try { DriveApp.getFileById(ss.getId()).moveTo(parent); } catch (e) {}
    props.setProperty('SS_ID', ss.getId());
  }

  ensureSheet_(ss, 'Books', BOOK_HEADERS);
  ensureSheet_(ss, 'Pages', PAGE_HEADERS);
  removeStraySheets_(ss);

  props.setProperty('IMG_FOLDER_ID', subFolder_(parent, IMG_FOLDER).getId());
  props.setProperty('PDF_FOLDER_ID', subFolder_(parent, PDF_FOLDER).getId());

  installTriggers_();
  var msg = 'SETUP COMPLETE\nSpreadsheet: ' + ss.getUrl();
  Logger.log(msg);
  return msg;
}

function subFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function resetData() {
  var ss = getSS_();
  resetSheet_(ss, 'Books', BOOK_HEADERS);
  resetSheet_(ss, 'Pages', PAGE_HEADERS);
  removeStraySheets_(ss);
  var n = 0;
  ['IMG_FOLDER_ID','PDF_FOLDER_ID'].forEach(function (key) {
    var id = PropertiesService.getScriptProperties().getProperty(key);
    if (!id) return;
    try {
      var files = DriveApp.getFolderById(id).getFiles();
      while (files.hasNext()) { files.next().setTrashed(true); n++; }
    } catch (e) {}
  });
  var msg = 'RESET COMPLETE — sheets rebuilt, ' + n + ' file(s) removed.';
  Logger.log(msg);
  return msg;
}

function resetSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  else { sh.clear(); try { sh.setFrozenRows(0); } catch (e) {} }
  writeHeaders_(sh, headers);
  return sh;
}

function removeStraySheets_(ss) {
  var keep = { 'Books': true, 'Pages': true };
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (keep[all[i].getName()]) continue;
    if (ss.getSheets().length <= 1) break;
    try { ss.deleteSheet(all[i]); } catch (e) {}
  }
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); writeHeaders_(sh, headers); return sh; }
  var lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  if (lastRow === 0 || lastCol === 0) { writeHeaders_(sh, headers); return sh; }
  var existing = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){return String(h).trim();});
  if (lastRow === 1) { sh.getRange(1,1,1,lastCol).clearContent(); writeHeaders_(sh, headers); return sh; }
  var missing = [];
  for (var i = 0; i < headers.length; i++) if (existing.indexOf(headers[i]) === -1) missing.push(headers[i]);
  if (missing.length) {
    sh.getRange(1, existing.length+1, 1, missing.length).setValues([missing]);
    sh.getRange(1,1,1, existing.length+missing.length).setFontWeight('bold');
  }
  return sh;
}

function writeHeaders_(sh, headers) {
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function installTriggers_() {
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    var f = all[i].getHandlerFunction();
    if (f === 'dailyMaintenance' || f === 'processDeliveries') ScriptApp.deleteTrigger(all[i]);
  }
  ScriptApp.newTrigger('processDeliveries').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('dailyMaintenance').timeBased().atHour(3).everyDays(1).create();
}


/* =================== STORAGE =================== */

function getSS_() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('Setup has not been run yet.');
  return SpreadsheetApp.openById(id);
}
function sheet_(name, headers) {
  var ss = getSS_(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); writeHeaders_(sh, headers); }
  return sh;
}
function headerMap_(sh) {
  var lastCol = sh.getLastColumn();
  if (!lastCol) return {};
  var row = sh.getRange(1,1,1,lastCol).getValues()[0], map = {};
  for (var c = 0; c < row.length; c++) {
    var k = String(row[c]).trim();
    if (k && map[k] === undefined) map[k] = c + 1;
  }
  return map;
}
function ensureHeaders_(sh, headers) {
  var map = headerMap_(sh), missing = [];
  for (var i = 0; i < headers.length; i++) if (!map[headers[i]]) missing.push(headers[i]);
  if (missing.length) {
    var start = sh.getLastColumn() + 1;
    sh.getRange(1,start,1,missing.length).setValues([missing]);
    sh.getRange(1,1,1,start+missing.length-1).setFontWeight('bold');
    map = headerMap_(sh);
  }
  return map;
}
function headersFor_(name) { return (name === 'Books') ? BOOK_HEADERS : PAGE_HEADERS; }

function readSheet_(name) {
  var sh = sheet_(name, headersFor_(name));
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  var vals = sh.getRange(1,1,lastRow,lastCol).getValues();
  var head = vals[0].map(function(h){return String(h).trim();});
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var o = { _row: r+1 }, blank = true;
    for (var c = 0; c < head.length; c++) {
      if (!head[c]) continue;
      o[head[c]] = vals[r][c];
      if (vals[r][c] !== '' && vals[r][c] !== null) blank = false;
    }
    if (!blank) out.push(o);
  }
  return out;
}
function appendRow_(name, headers, obj) {
  var sh = sheet_(name, headers), map = ensureHeaders_(sh, headers);
  var width = sh.getLastColumn(), row = new Array(width);
  for (var i = 0; i < width; i++) row[i] = '';
  for (var h in map) if (obj[h] !== undefined) row[map[h]-1] = obj[h];
  sh.appendRow(row);
}
function updateRow_(name, headers, rowIndex, obj) {
  var sh = sheet_(name, headers), map = ensureHeaders_(sh, headers);
  for (var h in obj) { var col = map[h]; if (col) sh.getRange(rowIndex, col).setValue(obj[h]); }
}
function deleteRow_(name, rowIndex) { sheet_(name, headersFor_(name)).deleteRow(rowIndex); }


/* =================== HELPERS =================== */

function folderProp_(key) {
  var id = PropertiesService.getScriptProperties().getProperty(key);
  if (!id) throw new Error('Setup has not been run yet.');
  return DriveApp.getFolderById(id);
}
function getImagesFolder_() { return folderProp_('IMG_FOLDER_ID'); }
function getPdfFolder_()    { return folderProp_('PDF_FOLDER_ID'); }

function getUserEmail_() {
  var em = '';
  try { em = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  if (!em) { try { em = Session.getEffectiveUser().getEmail() || ''; } catch (e) {} }
  return String(em).toLowerCase().trim();
}
function requireUser_() {
  var em = getUserEmail_();
  if (!em) throw new Error('We could not identify your Roche account. Please sign in and reload.');
  return em;
}
function baseUrl_() { return ScriptApp.getService().getUrl(); }
function uuid_() { return Utilities.getUuid().replace(/-/g,'').substring(0,12); }

function parseEmails_(v) {
  if (!v) return [];
  return String(v).split(/[\s,;]+/).map(function(s){return s.toLowerCase().trim();})
    .filter(function(s){return s.indexOf('@') > 0;});
}
function nameFromEmail_(email) {
  if (!email) return '';
  return email.split('@')[0].split(/[._-]+/).map(function(p){
    return p ? p.charAt(0).toUpperCase()+p.slice(1) : '';
  }).join(' ').trim();
}
function parseDeadline_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3], +m[4], +m[5], 0);
  var d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}
function imgUrl_(id) { return id ? 'https://drive.google.com/thumbnail?id='+id+'&sz=w1600' : ''; }

function saveImage_(dataUrl, hint) {
  if (!dataUrl) return '';
  var parts = String(dataUrl).split(',');
  if (parts.length < 2) return '';
  var ct = (parts[0].match(/:(.*?);/) || [])[1] || 'image/jpeg';
  var ext = {'image/gif':'gif','image/png':'png','image/webp':'webp','image/jpeg':'jpg'}[ct] || 'jpg';
  var blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), ct, (hint||'image')+'.'+ext);
  var file = getImagesFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return file.getId();
}
function imageMime_(id) {
  if (!id) return '';
  try { return DriveApp.getFileById(String(id)).getBlob().getContentType() || ''; }
  catch (e) { return ''; }
}
function trashImage_(id) {
  if (!id) return;
  try { DriveApp.getFileById(String(id)).setTrashed(true); } catch (e) {}
}

function findBook_(bookId) {
  if (!bookId) return null;
  var rows = readSheet_('Books');
  for (var i = 0; i < rows.length; i++)
    if (String(rows[i].bookId).trim() === String(bookId).trim()) return rows[i];
  return null;
}
function pagesFor_(bookId) {
  var rows = readSheet_('Pages'), out = [];
  for (var i = 0; i < rows.length; i++)
    if (String(rows[i].bookId).trim() === String(bookId).trim() &&
        String(rows[i].deleted).toUpperCase() !== 'TRUE') out.push(rows[i]);
  out.sort(function(a,b){ return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime(); });
  return out;
}

function roleFor_(book, email) {
  if (!book || !email) return 'none';
  var me = String(email).toLowerCase().trim();
  if (String(book.ownerEmail).toLowerCase().trim() === me) return 'owner';
  if (parseEmails_(book.deputies).indexOf(me) !== -1) return 'deputy';
  var pages = readSheet_('Pages');
  for (var i = 0; i < pages.length; i++)
    if (String(pages[i].bookId).trim() === String(book.bookId).trim() &&
        String(pages[i].authorEmail).toLowerCase().trim() === me &&
        String(pages[i].deleted).toUpperCase() !== 'TRUE') return 'contributor';
  return 'none';
}
function canManage_(role) { return role === 'owner' || role === 'deputy'; }
function requireManager_(book) {
  var email = requireUser_(), role = roleFor_(book, email);
  if (!canManage_(role))
    throw new Error('You do not have permission to manage this book. Signed in as ' + email + '.');
  return role;
}

function isClosed_(book) {
  var s = String(book.status);
  if (s === 'cancelled' || s === 'locked' || s === 'delivered') return true;
  var d = parseDeadline_(book.deadline);
  return d ? Date.now() > d.getTime() : false;
}
function statusOf_(book) {
  var s = String(book.status);
  if (s === 'cancelled') return 'cancelled';
  if (s === 'delivered' || book.sentAt) return 'delivered';
  var d = parseDeadline_(book.deadline);
  if (!d) return 'open';
  var now = Date.now();
  if (now > d.getTime()) return 'closed';
  if (d.getTime() - now < 3*24*3600*1000) return 'closing';
  return 'open';
}
function links_(bookId) {
  var b = baseUrl_();
  return {
    contribute: b + '?page=contribute&id=' + bookId,
    book:       b + '?page=book&id=' + bookId,
    manage:     b + '?page=manage&id=' + bookId,
    home:       b
  };
}
function deletionDateText_(book) {
  var d = parseDeadline_(book.deadline);
  if (!d) return '';
  var del = new Date(d.getTime() + RETENTION_DAYS*24*3600*1000);
  return Utilities.formatDate(del, Session.getScriptTimeZone(), 'd MMM yyyy');
}
function fmtDeadlineLong_(v) {
  var d = parseDeadline_(v);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "EEEE d MMMM yyyy 'at' HH:mm");
}
function escHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}


/* =================== EMOJI → IMAGES (PDF only) =================== */

var EMOJI_RE = /(?:\uD83C[\uDDE6-\uDDFF]\uD83C[\uDDE6-\uDDFF])|(?:[\uD800-\uDBFF][\uDC00-\uDFFF](?:\u200D[\uD800-\uDBFF][\uDC00-\uDFFF])*)|[\u2190-\u21FF\u2300-\u23FF\u2460-\u24FF\u25A0-\u27BF\u2B00-\u2BFF\u3030\u303D\u3297\u3299]\uFE0F?/g;

var _emojiMemo = {};

function emojiCodepoints_(seq) {
  var cps = [], i = 0;
  while (i < seq.length) {
    var c = seq.charCodeAt(i);
    if (c >= 0xD800 && c <= 0xDBFF && i + 1 < seq.length) {
      var c2 = seq.charCodeAt(i + 1);
      cps.push((((c - 0xD800) * 0x400) + (c2 - 0xDC00) + 0x10000).toString(16));
      i += 2;
    } else {
      if (c !== 0xFE0F) cps.push(c.toString(16));
      i += 1;
    }
  }
  return cps.join('-');
}

function emojiDataUri_(code) {
  if (_emojiMemo[code] !== undefined) return _emojiMemo[code];

  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache) {
    var hit = cache.get('emo_' + code);
    if (hit) { _emojiMemo[code] = hit; return hit; }
    if (cache.get('emoX_' + code)) { _emojiMemo[code] = ''; return ''; }
  }

  var uri = '';
  try {
    var res = UrlFetchApp.fetch(TWEMOJI_BASE + code + '.png',
      { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true });
    if (res.getResponseCode() === 200) {
      uri = 'data:image/png;base64,' + Utilities.base64Encode(res.getContent());
      if (cache) { try { cache.put('emo_' + code, uri, 21600); } catch (e) {} }
    } else {
      if (cache) { try { cache.put('emoX_' + code, '1', 21600); } catch (e) {} }
    }
  } catch (e) {
    if (cache) { try { cache.put('emoX_' + code, '1', 21600); } catch (e2) {} }
  }

  _emojiMemo[code] = uri;
  return uri;
}

function pdfText_(s, px) {
  var size = px || 15;
  var raw = String(s == null ? '' : s);
  var out = '', last = 0, m;
  EMOJI_RE.lastIndex = 0;
  while ((m = EMOJI_RE.exec(raw)) !== null) {
    out += escHtml_(raw.substring(last, m.index));
    var code = emojiCodepoints_(m[0]);
    var uri = code ? emojiDataUri_(code) : '';
    if (uri) {
      out += '<img src="' + uri + '" width="' + size + '" height="' + size +
             '" style="vertical-align:-2px">';
    }
    last = m.index + m[0].length;
  }
  out += escHtml_(raw.substring(last));
  return out;
}

function testEmojiFetch() {
  _emojiMemo = {};
  try { CacheService.getScriptCache().removeAll(['emo_1f389','emoX_1f389']); } catch (e) {}
  var uri = emojiDataUri_('1f389');
  var msg = uri
    ? 'EMOJI OK — the CDN is reachable, emoji will appear in PDFs.'
    : 'EMOJI UNAVAILABLE — the CDN could not be reached. Emoji will be removed from PDFs (never shown as boxes). They still work in the online book.';
  Logger.log(msg);
  return msg;
}


/* =================== ROUTER =================== */

function doGet(e) {
  if (!PropertiesService.getScriptProperties().getProperty('SS_ID')) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Segoe UI,sans-serif;padding:48px;max-width:620px;margin:auto">' +
      '<h2>Setup not complete</h2><p>Run <b>setup</b> in the Apps Script editor, then reload.</p></div>');
  }
  var page = (e && e.parameter && e.parameter.page) || 'home';
  var id   = (e && e.parameter && e.parameter.id)   || '';
  var file, title;
  if (page === 'book')            { file='Book';       title='Celebration Book'; }
  else if (page === 'contribute') { file='Contribute'; title='Add your page'; }
  else if (page === 'manage')     { file='Manage';     title='Manage book'; }
  else if (page === 'create')     { file='Create';     title='New book'; }
  else                            { file='Home';       title=APP_NAME; }

  var t = HtmlService.createTemplateFromFile(file);
  t.bookId = id; t.appName = APP_NAME;
  return t.evaluate().setTitle(title)
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }


/* =================== DASHBOARD =================== */

function apiDashboard() {
  var email = requireUser_();
  var books = readSheet_('Books'), pages = readSheet_('Pages');
  var counts = {}, mine = {};
  for (var i = 0; i < pages.length; i++) {
    var p = pages[i];
    if (String(p.deleted).toUpperCase() === 'TRUE') continue;
    var bid = String(p.bookId).trim();
    counts[bid] = (counts[bid]||0) + 1;
    if (String(p.authorEmail).toLowerCase().trim() === email) mine[bid] = true;
  }
  var managed = [], contributing = [];
  for (var b = 0; b < books.length; b++) {
    var bk = books[b], bid2 = String(bk.bookId).trim();
    if (!bid2 || bk.deletedAt) continue;
    var isOwner  = String(bk.ownerEmail).toLowerCase().trim() === email;
    var isDeputy = parseEmails_(bk.deputies).indexOf(email) !== -1;
    var didWrite = !!mine[bid2];
    if (!isOwner && !isDeputy && !didWrite) continue;
    var row = {
      bookId: bid2, occasion: bk.occasion, honouree: bk.honouree,
      theme: bk.theme||'roche', deadline: String(bk.deadline),
      status: statusOf_(bk), ownerName: bk.ownerName, ownerEmail: bk.ownerEmail,
      links: links_(bid2)
    };
    if (isOwner || isDeputy) {
      row.role = isOwner ? 'owner' : 'deputy';
      row.count = counts[bid2] || 0;
      row.deleteOn = deletionDateText_(bk);
      managed.push(row);
    } else {
      row.role = 'contributor';
      contributing.push(row);
    }
  }
  var byD = function(x,y){
    var a = parseDeadline_(x.deadline), c = parseDeadline_(y.deadline);
    return (a?a.getTime():0) - (c?c.getTime():0);
  };
  managed.sort(byD); contributing.sort(byD);
  return {
    email: email, name: nameFromEmail_(email),
    managed: managed, contributing: contributing,
    retentionDays: RETENTION_DAYS, createUrl: baseUrl_() + '?page=create'
  };
}


/* =================== CREATE =================== */

function apiCreateBook(p) {
  var email = requireUser_();
  if (!p.occasion) throw new Error('Please enter an occasion title.');
  if (!p.honouree) throw new Error('Please enter who you are celebrating.');
  if (!p.deadline) throw new Error('Please choose a closing date and time.');
  var d = parseDeadline_(p.deadline);
  if (!d) throw new Error('That closing date could not be read.');
  if (d.getTime() < Date.now()) throw new Error('The closing date must be in the future.');

  var bookId = uuid_();
  appendRow_('Books', BOOK_HEADERS, {
    bookId: bookId, createdAt: new Date().toISOString(),
    ownerEmail: email, ownerName: nameFromEmail_(email),
    deputies: parseEmails_(p.deputies).join(', '),
    occasion: p.occasion, honouree: p.honouree,
    coverMessage: p.coverMessage||'', welcomeMessage: p.welcomeMessage||'',
    theme: p.theme||'roche',
    coverImageId: p.coverImageData ? saveImage_(p.coverImageData,'cover-'+bookId) : '',
    bgStyle: p.bgStyle||'balloons',
    bgImageId: p.bgImageData ? saveImage_(p.bgImageData,'bg-'+bookId) : '',
    deadline: p.deadline,
    recipients: parseEmails_(p.recipients).join(', '),
    status: 'open', sentAt: '', warnedAt: '', deletedAt: '', pdfFileId: ''
  });

  var check = findBook_(bookId);
  if (!check || String(check.ownerEmail).toLowerCase().trim() !== email)
    throw new Error('Saved but owner not confirmed. Run resetData and try again.');

  return { bookId: bookId, links: links_(bookId) };
}

function apiGetTemplate(bookId) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  requireManager_(book);
  return {
    occasion: book.occasion, honouree: '',
    coverMessage: book.coverMessage, welcomeMessage: book.welcomeMessage,
    theme: book.theme||'roche', bgStyle: book.bgStyle||'balloons',
    deputies: book.deputies, recipients: ''
  };
}


/* =================== CONTRIBUTE =================== */

function apiGetContributeInfo(bookId) {
  var book = findBook_(bookId);
  if (!book || book.deletedAt) throw new Error('This celebration book could not be found.');
  var email = requireUser_(), mine = null;
  var all = readSheet_('Pages');
  for (var i = 0; i < all.length; i++) {
    if (String(all[i].bookId).trim() === String(bookId).trim() &&
        String(all[i].authorEmail).toLowerCase().trim() === email &&
        String(all[i].deleted).toUpperCase() !== 'TRUE') {
      mine = { pageId: all[i].pageId, authorName: all[i].authorName,
               message: all[i].message, imageUrl: imgUrl_(all[i].imageId),
               imageId: all[i].imageId, imageMime: imageMime_(all[i].imageId) };
      break;
    }
  }
  return {
    bookId: bookId, occasion: book.occasion, honouree: book.honouree,
    welcomeMessage: book.welcomeMessage, theme: book.theme||'roche',
    deadline: String(book.deadline), closed: isClosed_(book),
    cancelled: String(book.status) === 'cancelled',
    guessName: nameFromEmail_(email), maxMessage: MAX_MESSAGE,
    existing: mine, bookUrl: links_(bookId).book, homeUrl: baseUrl_()
  };
}

function apiSaveContribution(bookId, p) {
  var book = findBook_(bookId);
  if (!book || book.deletedAt) throw new Error('This celebration book could not be found.');
  if (isClosed_(book)) throw new Error('This book has closed — contributions are no longer accepted.');
  var email = requireUser_();
  var name = String(p.authorName||'').trim(), msg = String(p.message||'').trim();
  if (!name) throw new Error('Please enter your name.');
  if (!msg)  throw new Error('Please write a message.');
  if (msg.length > MAX_MESSAGE) throw new Error('Your message is too long for one page.');

  var all = readSheet_('Pages'), existing = null;
  for (var i = 0; i < all.length; i++)
    if (String(all[i].bookId).trim() === String(bookId).trim() &&
        String(all[i].authorEmail).toLowerCase().trim() === email &&
        String(all[i].deleted).toUpperCase() !== 'TRUE') { existing = all[i]; break; }

  var imageId = existing ? existing.imageId : '';
  if (p.removeImage) { trashImage_(imageId); imageId = ''; }
  if (p.imageData) { if (imageId) trashImage_(imageId); imageId = saveImage_(p.imageData,'page-'+bookId); }

  if (existing) {
    updateRow_('Pages', PAGE_HEADERS, existing._row, {
      updatedAt: new Date().toISOString(), authorName: name, message: msg, imageId: imageId });
    return { ok:true, updated:true };
  }
  appendRow_('Pages', PAGE_HEADERS, {
    pageId: uuid_(), bookId: bookId,
    createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    authorEmail: email, authorName: name, message: msg, imageId: imageId, deleted: 'FALSE' });
  return { ok:true, updated:false };
}

function apiDeleteMyContribution(bookId) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  if (isClosed_(book)) throw new Error('This book has closed.');
  var email = requireUser_();
  var all = readSheet_('Pages');
  for (var i = 0; i < all.length; i++)
    if (String(all[i].bookId).trim() === String(bookId).trim() &&
        String(all[i].authorEmail).toLowerCase().trim() === email &&
        String(all[i].deleted).toUpperCase() !== 'TRUE') {
      trashImage_(all[i].imageId);
      updateRow_('Pages', PAGE_HEADERS, all[i]._row,
        { deleted:'TRUE', imageId:'', updatedAt:new Date().toISOString() });
      return { ok:true };
    }
  throw new Error('You do not have a page in this book yet.');
}


/* =================== VIEW BOOK =================== */

function apiGetBook(bookId) {
  var book = findBook_(bookId);
  if (!book || book.deletedAt) throw new Error('This celebration book could not be found.');
  var viewer = getUserEmail_();
  var recipients = parseEmails_(book.recipients);
  var thanked = parseEmails_(book.thankYouBy);
  var pages = pagesFor_(bookId).map(function(p){
    return { authorName:p.authorName, message:p.message,
             imageUrl:imgUrl_(p.imageId), imageId:p.imageId, imageMime:imageMime_(p.imageId) };
  });
  return {
    occasion: book.occasion, honouree: book.honouree, coverMessage: book.coverMessage,
    theme: book.theme||'roche', bgStyle: book.bgStyle||'balloons',
    bgImageUrl: imgUrl_(book.bgImageId), coverImageUrl: imgUrl_(book.coverImageId),
    homeUrl: baseUrl_(), pages: pages,
    canThank: !!book.sentAt && !!viewer && recipients.indexOf(viewer) !== -1,
    thankYouSent: !!viewer && thanked.indexOf(viewer) !== -1,
    thankYouDefault: 'Thank you so much for contributing to my celebration book. Your message and wishes meant a lot to me.'
  };
}

function thankYouHtml_(book, message) {
  return '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:15px;line-height:1.65;color:' + R_INK + '">' +
    '<p>Hello,</p>' +
    '<p><b>' + escHtml_(book.honouree) + '</b> has sent a thank you to everyone who contributed to the celebration book.</p>' +
    '<div style="margin:22px 0;padding:18px 20px;background:' + R_TINT + ';border-left:4px solid ' + R_BLUE + ';font-style:italic">' +
      escHtml_(message).replace(/\n/g,'<br>') + '</div>' +
    '<p style="color:' + R_MUTE + ';font-size:13px">You are receiving this because you contributed a page to the book. Other contributors are hidden in BCC.</p></div>';
}

function apiSendThankYou(bookId, message) {
  var book = findBook_(bookId);
  if (!book || book.deletedAt) throw new Error('This celebration book could not be found.');
  var sender = requireUser_();
  if (!book.sentAt) throw new Error('The thank-you can be sent only after the book has been delivered.');
  if (parseEmails_(book.recipients).indexOf(sender) === -1)
    throw new Error('Only a recipient of this book can send the thank-you.');
  var already = parseEmails_(book.thankYouBy);
  if (already.indexOf(sender) !== -1) throw new Error('You have already sent a thank-you to the contributors.');
  var txt = String(message || '').trim();
  if (!txt) throw new Error('Please write a short thank-you message.');
  if (txt.length > 800) throw new Error('Please keep the thank-you message under 800 characters.');

  var pages = pagesFor_(bookId), seen = {}, targets = [];
  for (var i = 0; i < pages.length; i++) {
    var em = String(pages[i].authorEmail || '').toLowerCase().trim();
    if (!em || seen[em] || em === sender) continue;
    seen[em] = true;
    targets.push(em);
  }
  if (!targets.length) throw new Error('There are no contributor email addresses to thank.');

  MailApp.sendEmail({
    to: sender,
    bcc: targets.join(','),
    subject: 'A thank you from ' + (book.honouree || 'the celebration book recipient'),
    htmlBody: thankYouHtml_(book, txt),
    name: (book.honouree || APP_NAME) + ' (via ' + APP_NAME + ')',
    replyTo: sender
  });

  already.push(sender);
  var fresh = findBook_(bookId);
  if (fresh) updateRow_('Books', BOOK_HEADERS, fresh._row, { thankYouBy: already.join(', ') });
  return { ok:true, sent:targets.length };
}


/* =================== MANAGE =================== */

function apiGetManage(bookId) {
  var book = findBook_(bookId);
  if (!book || book.deletedAt) throw new Error('This celebration book could not be found.');
  var role = requireManager_(book);
  var pages = pagesFor_(bookId).map(function(p){
    return { pageId:p.pageId, authorName:p.authorName, authorEmail:p.authorEmail,
             message:p.message, imageUrl:imgUrl_(p.imageId) };
  });
  var pdfUrl = '';
  if (book.pdfFileId) { try { pdfUrl = DriveApp.getFileById(String(book.pdfFileId)).getUrl(); } catch (e) {} }

  return {
    bookId: bookId, role: role,
    occasion: book.occasion, honouree: book.honouree,
    coverMessage: book.coverMessage, welcomeMessage: book.welcomeMessage,
    theme: book.theme||'roche', bgStyle: book.bgStyle||'balloons',
    bgImageUrl: imgUrl_(book.bgImageId), coverImageUrl: imgUrl_(book.coverImageId),
    deadline: String(book.deadline),
    recipients: book.recipients, deputies: book.deputies,
    ownerName: book.ownerName, ownerEmail: book.ownerEmail,
    status: statusOf_(book), sentAt: String(book.sentAt||''),
    pdfUrl: pdfUrl, deleteOn: deletionDateText_(book), retentionDays: RETENTION_DAYS,
    links: links_(bookId), homeUrl: baseUrl_(), pages: pages
  };
}

function apiUpdateBook(bookId, fields) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  requireManager_(book);
  var upd = {};
  ['occasion','honouree','coverMessage','welcomeMessage','theme','bgStyle']
    .forEach(function(k){ if (fields[k] !== undefined) upd[k] = fields[k]; });
  if (fields.recipients !== undefined) upd.recipients = parseEmails_(fields.recipients).join(', ');
  if (fields.deputies   !== undefined) upd.deputies   = parseEmails_(fields.deputies).join(', ');
  if (fields.deadline) {
    if (!parseDeadline_(fields.deadline)) throw new Error('That closing date could not be read.');
    upd.deadline = fields.deadline;
  }
  if (fields.coverImageData) {
    trashImage_(book.coverImageId);
    upd.coverImageId = saveImage_(fields.coverImageData,'cover-'+bookId);
  }
  if (fields.bgImageData) {
    trashImage_(book.bgImageId);
    upd.bgImageId = saveImage_(fields.bgImageData,'bg-'+bookId);
  }
  if (fields.removeBgImage) { trashImage_(book.bgImageId); upd.bgImageId = ''; }
  updateRow_('Books', BOOK_HEADERS, book._row, upd);
  return { ok:true };
}

function apiRemovePage(bookId, pageId) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  requireManager_(book);
  var all = readSheet_('Pages');
  for (var i = 0; i < all.length; i++)
    if (String(all[i].pageId).trim() === String(pageId).trim()) {
      trashImage_(all[i].imageId);
      updateRow_('Pages', PAGE_HEADERS, all[i]._row,
        { deleted:'TRUE', imageId:'', updatedAt:new Date().toISOString() });
      return { ok:true };
    }
  throw new Error('That page could not be found.');
}

function apiCancelBook(bookId) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  requireManager_(book);
  if (book.sentAt) throw new Error('This book has already been delivered.');
  updateRow_('Books', BOOK_HEADERS, book._row, { status:'cancelled' });
  return { ok:true };
}

function apiReopenBook(bookId, newDeadline) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  requireManager_(book);
  var d = parseDeadline_(newDeadline);
  if (!d) throw new Error('Please choose a new closing date.');
  if (d.getTime() < Date.now()) throw new Error('The new closing date must be in the future.');
  updateRow_('Books', BOOK_HEADERS, book._row, { status:'open', deadline:newDeadline });
  return { ok:true };
}

function apiDeleteBook(bookId) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  requireManager_(book);
  purgeBook_(book);
  return { ok:true };
}

function purgeBook_(book) {
  var pages = readSheet_('Pages'), rows = [];
  for (var i = 0; i < pages.length; i++)
    if (String(pages[i].bookId).trim() === String(book.bookId).trim()) {
      trashImage_(pages[i].imageId); rows.push(pages[i]._row);
    }
  rows.sort(function(a,b){return b-a;});
  for (var r = 0; r < rows.length; r++) deleteRow_('Pages', rows[r]);
  trashImage_(book.coverImageId);
  trashImage_(book.bgImageId);
  trashImage_(book.pdfFileId);
  var fresh = findBook_(book.bookId);
  if (fresh) deleteRow_('Books', fresh._row);
}


/* =================== DELIVERY =================== */

function emailFestive_() {
  return '<div style="font-size:15px;letter-spacing:10px;line-height:1">' +
    '<span style="color:#BBD5FA">&#10022;</span>' +
    '<span style="color:#FFD966">&#9679;</span>' +
    '<span style="color:#FFFFFF">&#10022;</span>' +
    '<span style="color:#FFD966">&#9679;</span>' +
    '<span style="color:#BBD5FA">&#10022;</span></div>';
}

function deliveryHtml_(book) {
  var link = links_(book.bookId).book;
  var thankLink = link + '&thank=1';
  return '' +
  '<div style="margin:0;padding:0;background-color:#F4F6F9;">' +
  '<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#F4F6F9"><tr>' +
  '<td align="center" style="padding:28px 14px">' +
  '<table width="600" cellpadding="0" cellspacing="0" border="0" ' +
  'style="max-width:600px;font-family:Segoe UI,Arial,sans-serif">' +
    '<tr><td height="5" bgcolor="' + R_GOLD + '" style="border-radius:10px 10px 0 0"></td></tr>' +
    '<tr><td bgcolor="' + R_BLUE + '" align="center" style="padding:36px 30px 34px 30px">' +
      emailFestive_() +
      '<div style="height:20px"></div>' +
      '<div style="font-size:12px;letter-spacing:4px;text-transform:uppercase;color:#BBD5FA">' +
        escHtml_(book.occasion) + '</div>' +
      '<div style="height:14px"></div>' +
      '<div style="font-size:32px;font-weight:700;font-family:Georgia,serif;color:#FFFFFF;line-height:1.2">' +
        escHtml_(book.honouree) + '</div>' +
      (book.coverMessage
        ? '<div style="height:12px"></div><div style="font-size:15px;font-style:italic;color:#DCE8FC">' +
          escHtml_(book.coverMessage) + '</div>' : '') +
      '<div style="height:22px"></div>' + emailFestive_() +
    '</td></tr>' +
    '<tr><td bgcolor="#FFFFFF" style="padding:34px 34px 30px 34px;color:' + R_INK + ';' +
      'font-size:15px;line-height:1.68;border-left:1px solid #DBE0E6;border-right:1px solid #DBE0E6">' +
      '<p style="margin:0 0 16px 0">Your colleagues have put together a celebration book for you.</p>' +
      '<p style="margin:0 0 26px 0">Every page is a message written just for you. ' +
      'You can read it online, or keep the PDF attached to this email.</p>' +
      '<table cellpadding="0" cellspacing="0" border="0" align="center"><tr>' +
      '<td bgcolor="' + R_BLUE + '" style="border-radius:6px">' +
      '<a href="' + link + '" style="display:inline-block;padding:15px 38px;color:#FFFFFF;' +
      'text-decoration:none;font-weight:600;font-size:15px">Open the celebration book</a></td></tr></table>' +
      '<div style="height:12px"></div>' +
      '<div style="font-size:13px;color:' + R_MUTE + ';text-align:center"><b>PDF attached to this email</b></div>' +
      '<div style="height:22px"></div>' +
      '<table cellpadding="0" cellspacing="0" border="0" align="center"><tr>' +
      '<td bgcolor="#FFFFFF" style="border:1px solid ' + R_BLUE + ';border-radius:6px">' +
      '<a href="' + thankLink + '" style="display:inline-block;padding:13px 30px;color:' + R_BLUE + ';' +
      'text-decoration:none;font-weight:600;font-size:14px">Send a thank you to contributors</a></td></tr></table>' +
      '<div style="height:30px"></div>' +
      '<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="' + R_TINT + '">' +
      '<tr><td style="padding:16px 18px;font-size:13.5px;color:' + R_INK + '">' +
      '<b>The PDF is attached.</b> The online version stays available for ' + RETENTION_DAYS +
      ' days, so please keep the PDF as your copy.</td></tr></table>' +
    '</td></tr>' +
    '<tr><td height="6" bgcolor="' + R_BLUE + '" style="border-radius:0 0 10px 10px"></td></tr>' +
  '</table></td></tr></table></div>';
}

function deliverBook_(book) {
  var pages = pagesFor_(book.bookId);
  var recipients = parseEmails_(book.recipients);
  var managers = [String(book.ownerEmail).toLowerCase()]
                   .concat(parseEmails_(book.deputies))
                   .filter(function(x){ return x && x.indexOf('@') > 0; });

  if (!pages.length) {
    notifyManagers_(book, 'No contributions to send',
      'The celebration book reached its closing date with no contributions, so nothing was sent.');
    var f0 = findBook_(book.bookId);
    if (f0) updateRow_('Books', BOOK_HEADERS, f0._row, { status:'locked' });
    return { sent:false, reason:'no contributions' };
  }
  if (!recipients.length) {
    notifyManagers_(book, 'No recipients set',
      'The book has closed but no recipient email was set. Add one and use &ldquo;Send now&rdquo;.');
    var f1 = findBook_(book.bookId);
    if (f1) updateRow_('Books', BOOK_HEADERS, f1._row, { status:'locked' });
    return { sent:false, reason:'no recipients' };
  }

  var file = generatePdf_(book);
  var sizeMb = file.getSize() / 1048576;
  var attach = (sizeMb <= MAX_ATTACH_MB);

  var opts = {
    to: recipients.join(','),
    subject: book.occasion + ' — a celebration book for ' + book.honouree,
    htmlBody: deliveryHtml_(book),
    name: (book.ownerName || nameFromEmail_(book.ownerEmail)) + ' (via PD Celebration Book)',
    replyTo: String(book.ownerEmail),
    bcc: managers.join(',')
  };
  if (attach) opts.attachments = [file.getBlob()];
  MailApp.sendEmail(opts);

  var fresh = findBook_(book.bookId);
  if (fresh) updateRow_('Books', BOOK_HEADERS, fresh._row,
    { status:'delivered', sentAt:new Date().toISOString() });

  if (!attach) {
    notifyManagers_(book, 'Delivered — PDF too large to attach',
      'The book was delivered, but the PDF is ' + Math.round(sizeMb) + ' MB, above the ' +
      MAX_ATTACH_MB + ' MB limit, so only the link was sent.');
  }
  return { sent:true, recipients:recipients.length, attached:attach,
           sizeMb: Math.round(sizeMb*10)/10 };
}

function notifyManagers_(book, subject, htmlInner) {
  var to = [String(book.ownerEmail)].concat(parseEmails_(book.deputies))
             .filter(function(x){ return x && x.indexOf('@') > 0; }).join(',');
  if (!to) return;
  MailApp.sendEmail({
    to: to, subject: APP_NAME + ' — ' + subject,
    htmlBody: '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:' + R_INK +
              ';line-height:1.6"><p>' + htmlInner + '</p><p><a href="' + links_(book.bookId).manage +
              '" style="color:' + R_BLUE + '">Open the book</a></p></div>',
    name: APP_NAME
  });
}

function apiSendNow(bookId) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  requireManager_(book);
  if (book.sentAt) throw new Error('This book has already been delivered.');
  if (String(book.status) === 'cancelled') throw new Error('This book is cancelled. Reopen it first.');
  if (!parseEmails_(book.recipients).length)
    throw new Error('Please add at least one recipient email before sending.');
  return deliverBook_(book);
}


/* ===================================================================
 *  PDF
 *  RULES — do not "modernise":
 *   1. Colour via bgcolor="" on <td>, never CSS background-color.
 *   2. NEVER put CSS padding on a <td> that also has height="".
 *      That is what caused the blank overflow pages.
 *   3. Header + body + footer must sum EXACTLY to PDF_H.
 * =================================================================== */

var PDF_W       = 780;
var PDF_H       = 1000;
var PDF_HDR_BAR = 10;
var PDF_HDR_STR = 34;
var PDF_FTR_FST = 44;
var PDF_FTR_BAR = 8;
var PDF_HDR     = PDF_HDR_BAR + PDF_HDR_STR;
var PDF_FTR     = PDF_FTR_FST + PDF_FTR_BAR;
var PDF_BODY    = PDF_H - PDF_HDR - PDF_FTR;
var PDF_CARD_W  = 660;
var PDF_CARD_H  = 830;
var PDF_INSET_W = 560;

function imageDataUri_(fileId) {
  if (!fileId) return '';
  try {
    var blob = DriveApp.getFileById(String(fileId)).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) { return ''; }
}

function pdfImage_(fileId, maxW, maxH) {
  if (!fileId) return null;
  try {
    var blob = DriveApp.getFileById(String(fileId)).getBlob();
    var uri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    var w = 0, h = 0;
    try {
      var im = ImagesService.openImage(blob);
      w = im.getWidth(); h = im.getHeight();
    } catch (e2) {}

    if (!w || !h) {
      return { uri:uri, width:maxW, height:'', style:'max-width:'+maxW+'px;max-height:'+maxH+'px' };
    }
    var scale = Math.min(maxW / w, maxH / h, 1);
    return {
      uri: uri,
      width: Math.max(1, Math.round(w * scale)),
      height: Math.max(1, Math.round(h * scale)),
      style: ''
    };
  } catch (e) { return null; }
}

function festiveRow_(size, light) {
  var gold = light ? '#FFD966' : R_GOLD;
  var blue = light ? '#FFFFFF' : R_LITE;
  var pale = light ? '#BBD5FA' : '#8FC0F7';
  return '<div style="font-size:' + size + 'px;letter-spacing:9px;line-height:1">' +
    '<font color="' + pale + '">&#10022;</font><font color="' + gold + '">&#9679;</font>' +
    '<font color="' + blue + '">&#10022;</font><font color="' + gold + '">&#9679;</font>' +
    '<font color="' + pale + '">&#10022;</font></div>';
}
function balloonRow_(size) {
  return '<div style="font-size:' + size + 'px;letter-spacing:14px;line-height:1">' +
    '<font color="#FFD966">&#9679;</font><font color="#8FC0F7">&#9679;</font>' +
    '<font color="#FFFFFF">&#9679;</font><font color="#8FC0F7">&#9679;</font>' +
    '<font color="#FFD966">&#9679;</font></div>';
}

function pdfMessageFont_(message) {
  var n = String(message || '').length;
  if (n <= 280) return 22;
  if (n <= 520) return 20;
  if (n <= 820) return 18;
  return 16;
}

function pdfInset_(innerHtml, width) {
  return '<table width="' + (width || PDF_INSET_W) + '" cellpadding="0" cellspacing="0" ' +
         'border="0" align="center"><tr><td align="left">' + innerHtml + '</td></tr></table>';
}
function spacer_(px) { return '<div style="height:' + px + 'px;line-height:' + px + 'px">&nbsp;</div>'; }

function pdfPageHeader_(book) {
  return '<tr><td height="' + PDF_HDR_BAR + '" bgcolor="' + R_BLUE + '"></td></tr>' +
    '<tr><td height="' + PDF_HDR_STR + '" bgcolor="' + R_TINT + '" align="center" valign="middle" ' +
    'style="font-family:Helvetica,Arial,sans-serif;font-size:9px;letter-spacing:3px;text-transform:uppercase">' +
    '<font color="' + R_MUTE + '">' + pdfText_(book.occasion, 9) +
    ' &nbsp;&#8226;&nbsp; ' + pdfText_(book.honouree, 9) + '</font></td></tr>';
}

function pdfPageFooter_() {
  return '<tr><td height="' + PDF_FTR_FST + '" align="center" valign="middle" bgcolor="' + R_TINT + '">' +
    festiveRow_(11, false) + '</td></tr>' +
    '<tr><td height="' + PDF_FTR_BAR + '" bgcolor="' + R_BLUE + '"></td></tr>';
}

function pdfPageOpen_(cls, bg) {
  return '<div class="pg ' + cls + '"><table width="' + PDF_W + '" height="' + PDF_H + '" ' +
         'cellpadding="0" cellspacing="0" border="0" bgcolor="' + bg + '">';
}
function pdfPageClose_() { return '</table></div>'; }

function buildPdfHtml_(book, pages, opts) {
  opts = opts || {};
  var withImages = opts.withImages !== false;

  var css =
    '@page { size:8.5in 11in; margin:0; }' +
    'body { margin:0; padding:0; font-family:Georgia,"Times New Roman",serif; color:' + R_INK + '; }' +
    'table { border-collapse:collapse; }' +
    'img { display:block; }' +
    '.pg { page-break-inside:avoid; }' +
    '.brk { page-break-after:always; }' +
    '.last { page-break-after:auto; }';

  var h = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' + css + '</style></head><body>';

  /* ---- COVER ---- */
  var coverImg = withImages ? pdfImage_(book.coverImageId, 210, 210) : null;
  var coverMain = PDF_H - 60;

  var coverInner = balloonRow_(20) + spacer_(24);
  if (coverImg) {
    coverInner += '<table cellpadding="0" cellspacing="0" border="0" align="center"><tr>' +
      '<td bgcolor="#FFFFFF" style="padding:6px">' +
      '<img src="' + coverImg.uri + '" width="' + coverImg.width + '"' +
      (coverImg.height ? ' height="' + coverImg.height + '"' : '') +
      ' style="display:block;' + coverImg.style + '">' +
      '</td></tr></table>' + spacer_(28);
  }
  coverInner +=
    '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px;letter-spacing:6px;' +
    'text-transform:uppercase"><font color="#BBD5FA">' + pdfText_(book.occasion, 14) + '</font></div>' +
    spacer_(16) +
    '<table cellpadding="0" cellspacing="0" border="0" align="center"><tr>' +
    '<td width="90" height="2" bgcolor="' + R_GOLD + '"></td></tr></table>' +
    spacer_(24) +
    '<div style="font-size:42px;font-weight:bold;line-height:1.16">' +
    '<font color="#FFFFFF">' + pdfText_(book.honouree, 38) + '</font></div>';
  if (book.coverMessage) {
    coverInner += spacer_(18) + '<div style="font-size:17px;font-style:italic">' +
      '<font color="#DCE8FC">' + pdfText_(book.coverMessage, 17) + '</font></div>';
  }
  coverInner += spacer_(30) + festiveRow_(15, true);

  h += pdfPageOpen_('brk', R_BLUE) +
       '<tr><td height="' + coverMain + '" align="center" valign="middle" bgcolor="' + R_BLUE + '">' +
       '<table width="600" cellpadding="0" cellspacing="0" border="0" align="center">' +
       '<tr><td align="center">' + coverInner + '</td></tr></table>' +
       '</td></tr>' +
       '<tr><td height="60" align="center" valign="middle" bgcolor="' + R_BLUE + '" ' +
       'style="font-family:Helvetica,Arial,sans-serif;font-size:10px;letter-spacing:4px;text-transform:uppercase">' +
       '<font color="#8FB4F2">A PD Celebration Book</font></td></tr>' +
       pdfPageClose_();

  /* ---- CONTRIBUTIONS ----
     photo    -> 2 pages (photo, then message)
     no photo -> 1 page  (message only)                */
  var CARD_BAR   = 6;
  var CARD_MAIN  = PDF_CARD_H - CARD_BAR;
  var SIGN_H     = 120;
  var TEXT_H     = CARD_MAIN - SIGN_H;
  var IMG_MAX_W  = PDF_CARD_W - 80;
  var IMG_MAX_H  = CARD_MAIN - 90;

  for (var i = 0; i < pages.length; i++) {
    var p = pages[i];
    var pic = withImages ? pdfImage_(p.imageId, IMG_MAX_W, IMG_MAX_H) : null;

    if (pic) {
      h += pdfPageOpen_('brk', R_TINT) + pdfPageHeader_(book) +
           '<tr><td height="' + PDF_BODY + '" align="center" valign="middle" bgcolor="' + R_TINT + '">' +
           '<table width="' + PDF_CARD_W + '" height="' + PDF_CARD_H + '" cellpadding="0" cellspacing="0" ' +
           'border="0" bgcolor="#FFFFFF">' +
           '<tr><td height="' + CARD_BAR + '" bgcolor="' + R_BLUE + '"></td></tr>' +
           '<tr><td height="' + CARD_MAIN + '" align="center" valign="middle" bgcolor="#FFFFFF">' +
           '<img src="' + pic.uri + '" width="' + pic.width + '"' +
           (pic.height ? ' height="' + pic.height + '"' : '') +
           ' style="display:block;margin:0 auto;' + pic.style + '">' +
           '</td></tr></table>' +
           '</td></tr>' + pdfPageFooter_() + pdfPageClose_();
    }

    var fs = pdfMessageFont_(p.message);
    var textBlock =
      spacer_(44) +
      '<div style="font-family:Helvetica,Arial,sans-serif;font-size:10px;letter-spacing:3px;text-transform:uppercase">' +
      '<font color="' + R_MUTE + '">From ' + pdfText_(p.authorName, 10) + '</font></div>' +
      spacer_(26) +
      '<div style="font-size:' + fs + 'px;line-height:1.6;white-space:pre-wrap">' +
      '<font color="' + R_INK + '">' + pdfText_(p.message, fs) + '</font></div>';

    var signBlock =
      '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
      '<tr><td height="1" bgcolor="' + R_LINE + '"></td></tr></table>' +
      spacer_(14) +
      '<div align="right" style="font-size:19px;font-style:italic;font-weight:bold">' +
      '<font color="' + R_BLUE + '">&mdash; ' + pdfText_(p.authorName, 19) + '</font></div>' +
      spacer_(32);

    h += pdfPageOpen_('brk', R_TINT) + pdfPageHeader_(book) +
         '<tr><td height="' + PDF_BODY + '" align="center" valign="middle" bgcolor="' + R_TINT + '">' +
         '<table width="' + PDF_CARD_W + '" height="' + PDF_CARD_H + '" cellpadding="0" cellspacing="0" ' +
         'border="0" bgcolor="#FFFFFF">' +
         '<tr><td height="' + CARD_BAR + '" bgcolor="' + R_GOLD + '"></td></tr>' +
         '<tr><td height="' + TEXT_H + '" valign="top" bgcolor="#FFFFFF">' +
         pdfInset_(textBlock) + '</td></tr>' +
         '<tr><td height="' + SIGN_H + '" valign="bottom" bgcolor="#FFFFFF">' +
         pdfInset_(signBlock) + '</td></tr>' +
         '</table>' +
         '</td></tr>' + pdfPageFooter_() + pdfPageClose_();
  }

  /* ---- CLOSING ---- */
  var closeInner =
    balloonRow_(22) + spacer_(36) +
    '<div style="font-size:31px;font-weight:bold;line-height:1.4">' +
    '<font color="#FFFFFF">With warm wishes<br>from all of us</font></div>' +
    spacer_(24) +
    '<div style="font-size:17px;font-style:italic">' +
    '<font color="#DCE8FC">' + pdfText_(book.honouree, 17) + '</font></div>' +
    spacer_(32) +
    '<table cellpadding="0" cellspacing="0" border="0" align="center"><tr>' +
    '<td width="90" height="2" bgcolor="' + R_GOLD + '"></td></tr></table>' +
    spacer_(28) + festiveRow_(15, true);

  h += pdfPageOpen_('last', R_BLUE) +
       '<tr><td height="' + PDF_H + '" align="center" valign="middle" bgcolor="' + R_BLUE + '">' +
       '<table width="580" cellpadding="0" cellspacing="0" border="0" align="center">' +
       '<tr><td align="center">' + closeInner + '</td></tr></table>' +
       '</td></tr>' + pdfPageClose_();

  return h + '</body></html>';
}

function safeFileName_(book) {
  var raw = (book.occasion || 'Celebration') + ' - ' + (book.honouree || '');
  return raw.replace(/[\\/:*?"<>|]/g,'').replace(EMOJI_RE,'').trim().substring(0,90) + '.pdf';
}

function generatePdf_(book) {
  var pages = pagesFor_(book.bookId);
  var html = buildPdfHtml_(book, pages, { withImages:true });
  var blob;
  try { blob = Utilities.newBlob(html,'text/html','book.html').getAs('application/pdf'); }
  catch (e) {
    html = buildPdfHtml_(book, pages, { withImages:false });
    blob = Utilities.newBlob(html,'text/html','book.html').getAs('application/pdf');
  }
  blob.setName(safeFileName_(book));
  trashImage_(book.pdfFileId);
  var file = getPdfFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var fresh = findBook_(book.bookId);
  if (fresh) updateRow_('Books', BOOK_HEADERS, fresh._row, { pdfFileId: file.getId() });
  return file;
}

function apiGeneratePdf(bookId) {
  var book = findBook_(bookId);
  if (!book) throw new Error('Book not found.');
  requireManager_(book);
  var pages = pagesFor_(bookId);
  if (!pages.length) throw new Error('There are no contributions yet, so there is nothing to put in the PDF.');
  var file = generatePdf_(book);
  return { ok:true, url:file.getUrl(), name:file.getName(),
           sizeMb: Math.round(file.getSize()/104857.6)/10 };
}


/* =================== TRIGGERS =================== */

function processDeliveries() {
  var books = readSheet_('Books'), now = Date.now();
  for (var i = 0; i < books.length; i++) {
    var b = books[i];
    if (!b.bookId || b.deletedAt || b.sentAt) continue;
    var st = String(b.status);
    if (st === 'cancelled' || st === 'delivered' || st === 'locked') continue;
    var d = parseDeadline_(b.deadline);
    if (!d || now < d.getTime()) continue;
    try { deliverBook_(b); }
    catch (e) {
      Logger.log('Delivery failed for ' + b.bookId + ': ' + e.message);
      try {
        notifyManagers_(b, 'Delivery could not be completed',
          'The book reached its closing date but could not be sent.<br><br>Reason: ' +
          escHtml_(e.message) + '<br><br>Please use &ldquo;Send now&rdquo;.');
      } catch (e2) {}
      var fresh = findBook_(b.bookId);
      if (fresh) updateRow_('Books', BOOK_HEADERS, fresh._row, { status:'locked' });
    }
  }
}

function dailyMaintenance() {
  var books = readSheet_('Books'), now = Date.now();
  for (var i = 0; i < books.length; i++) {
    var b = books[i];
    if (!b.bookId) continue;
    var d = parseDeadline_(b.deadline);
    if (!d) continue;
    var deleteAt = d.getTime() + RETENTION_DAYS*24*3600*1000;
    var warnAt   = deleteAt - WARN_DAYS*24*3600*1000;
    if (now >= deleteAt) { purgeBook_(b); continue; }
    if (now >= warnAt && !b.warnedAt) {
      try { sendDeletionWarning_(b); } catch (e) {}
      var fresh = findBook_(b.bookId);
      if (fresh) updateRow_('Books', BOOK_HEADERS, fresh._row, { warnedAt:new Date().toISOString() });
    }
  }
}

function sendDeletionWarning_(book) {
  var when = deletionDateText_(book);
  notifyManagers_(book, 'Book will be deleted on ' + when,
    'The celebration book <b>' + escHtml_(book.occasion) + ' &mdash; ' + escHtml_(book.honouree) +
    '</b> will be permanently deleted on <b>' + when + '</b>.<br><br>' +
    'Please download the PDF now if you need a copy.');
}

function apiGetImageData(fileId) {
  if (!fileId) return '';
  var blob = DriveApp.getFileById(String(fileId)).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}


/* =================== DIAGNOSTICS =================== */

function diagnose() {
  var out = [];
  out.push('Signed in as: ' + (getUserEmail_() || '(unknown)'));
  var ss;
  try { ss = getSS_(); }
  catch (e) { out.push('ERROR: ' + e.message); Logger.log(out.join('\n')); return out.join('\n'); }
  out.push('Spreadsheet: ' + ss.getUrl());
  var trg = ScriptApp.getProjectTriggers().map(function(t){return t.getHandlerFunction();});
  out.push('Triggers: ' + (trg.length ? trg.join(', ') : 'NONE — run setup'));
  out.push('Emoji CDN: ' + (emojiDataUri_('1f389') ? 'reachable' : 'NOT reachable (emoji stripped from PDFs)'));
  readSheet_('Books').forEach(function(b){
    var pgs = pagesFor_(b.bookId), withPic = 0;
    pgs.forEach(function(p){ if (p.imageId) withPic++; });
    out.push('  ' + b.bookId + ' | ' + statusOf_(b) + ' | owner=' + b.ownerEmail +
             ' | pages=' + pgs.length + ' (' + withPic + ' with photo)' +
             ' | expected pdf pages=' + (2 + pgs.length + withPic));
  });
  var txt = out.join('\n');
  Logger.log(txt);
  return txt;
}
