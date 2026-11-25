/**
 * مستقبل ملفات المواقع على Google Drive
 * يستقبل تقارير الفنيين من تطبيق "ملفات المواقع" ويحفظها في Drive الخاص بك:
 *   ملفات المواقع / المنظمة / اسم الموقع / تاريخ_اسم الفني_رقم / (الصور والفيديو والفويس + تقرير.txt + report.json)
 *
 * الإعداد: Deploy > New deployment > Web app (من غير أي رمز سري)
 * Execute as: Me   |   Who has access: Anyone
 */
const ROOT_NAME = 'ملفات المواقع';

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    switch (b.action) {
      case 'ping': return out({ ok: true, v: 1, root: rootFolder().getName() });
      case 'init': return out(initReport(b));
      case 'text': return out(saveText(b));
      case 'file-init': return out(fileInit(b));
      case 'chunk': return out(chunk(b));
      case 'done': return out(logDone(b));
      case 'list': return out(listDone());
      default: return out({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() { return out({ ok: true, hello: 'sitefiles receiver' }); }

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function child(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function rootFolder() {
  const P = PropertiesService.getScriptProperties();
  const id = P.getProperty('root');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recreate */ } }
  const f = child(DriveApp.getRootFolder(), ROOT_NAME);
  P.setProperty('root', f.getId());
  return f;
}

function clean(s) { return String(s || '').replace(/[\\/:*?"<>|\n\r]/g, '_').slice(0, 80) || '-'; }

// يعيد نفس المجلد لو أُعيد إرسال نفس التقرير
function initReport(b) {
  const P = PropertiesService.getScriptProperties();
  const known = P.getProperty('rep:' + b.reportId);
  if (known) { try { DriveApp.getFolderById(known); return { ok: true, folder: known }; } catch (e) { /* recreate */ } }
  const base = b.org ? child(rootFolder(), clean(b.org)) : rootFolder();
  const siteDir = child(base, clean(b.site));
  const d = Utilities.formatDate(new Date(b.at || Date.now()), 'Africa/Cairo', 'yyyy-MM-dd_HH-mm');
  const folder = child(siteDir, d + '_' + clean(b.tech) + '_' + String(b.reportId).slice(0, 6));
  P.setProperty('rep:' + b.reportId, folder.getId());
  return { ok: true, folder: folder.getId() };
}

function saveText(b) {
  const folder = DriveApp.getFolderById(b.folder);
  const it = folder.getFilesByName(b.name);
  if (it.hasNext()) { it.next().setContent(b.content); return { ok: true }; }
  folder.createFile(b.name, b.content, b.mime || 'text/plain');
  return { ok: true };
}

// رفع متقطع (resumable) على Drive: كل جزء 3 ميجا، ويستكمل من حيث وقف
function fileInit(b) {
  const size = Number(b.size);
  if (size === 0) {
    DriveApp.getFolderById(b.folder).createFile(b.name, '', b.type || 'application/octet-stream');
    return { ok: true, done: true, next: 0 };
  }
  const P = PropertiesService.getScriptProperties();
  const old = P.getProperty('up:' + b.key);
  if (old) {
    const s = JSON.parse(old);
    return { ok: true, key: b.key, next: queryOffset(s.url, s.size) };
  }
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'X-Upload-Content-Type': b.type || 'application/octet-stream',
      'X-Upload-Content-Length': String(size)
    },
    payload: JSON.stringify({ name: b.name, parents: [b.folder] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: 'drive ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200) };
  }
  const h = res.getAllHeaders();
  const url = h['Location'] || h['location'];
  P.setProperty('up:' + b.key, JSON.stringify({ url: url, size: size }));
  return { ok: true, key: b.key, next: 0 };
}

function queryOffset(url, size) {
  const res = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: { 'Content-Range': 'bytes */' + size },
    payload: '',
    muteHttpExceptions: true
  });
  const c = res.getResponseCode();
  if (c === 200 || c === 201) return size;
  const h = res.getAllHeaders();
  const r = h['Range'] || h['range'];
  return r ? parseInt(String(r).split('-')[1], 10) + 1 : 0;
}

function chunk(b) {
  const P = PropertiesService.getScriptProperties();
  const raw = P.getProperty('up:' + b.key);
  if (!raw) return { ok: false, error: 'no session' };
  const s = JSON.parse(raw);
  const bytes = Utilities.base64Decode(b.data);
  const off = Number(b.offset), end = off + bytes.length - 1;
  const res = UrlFetchApp.fetch(s.url, {
    method: 'put',
    contentType: 'application/octet-stream',
    headers: { 'Content-Range': 'bytes ' + off + '-' + end + '/' + s.size },
    payload: bytes,
    muteHttpExceptions: true
  });
  const c = res.getResponseCode();
  if (c === 200 || c === 201) { P.deleteProperty('up:' + b.key); return { ok: true, done: true, next: s.size }; }
  if (c === 308) {
    const h = res.getAllHeaders();
    const r = h['Range'] || h['range'];
    const next = r ? parseInt(String(r).split('-')[1], 10) + 1 : 0;
    return { ok: next === end + 1, next: next };
  }
  return { ok: false, next: queryOffset(s.url, s.size), error: 'drive ' + c };
}


// سجل التقارير المستلمة (للمدير)
function logFile() {
  const root = rootFolder();
  const it = root.getFilesByName('السجل.json');
  return it.hasNext() ? it.next() : root.createFile('السجل.json', '[]', 'application/json');
}
function logDone(b) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const f = logFile();
    let a = []; try { a = JSON.parse(f.getBlob().getDataAsString()); } catch (e) {}
    a = a.filter(x => x.folder !== b.folder);
    a.push({ folder: b.folder, org: b.org, site: b.site, code: b.code, tech: b.tech, at: b.at, count: b.count, bytes: b.bytes, maint: b.maint, sentAt: Date.now() });
    if (a.length > 600) a = a.slice(-600);
    f.setContent(JSON.stringify(a));
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function listDone() {
  let a = []; try { a = JSON.parse(logFile().getBlob().getDataAsString()); } catch (e) {}
  return { ok: true, items: a.slice(-200).reverse() };
}
