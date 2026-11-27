// DataX Advisory — lisenziya serveri (Google Sheets + Apps Script)
// Müştərilər vərəqində: tarixi dəyiş — uzadılır; Status = "dayandırılıb" — dayandırılır.
const NC = 11;
const C = { phone: 0, name: 1, plan: 2, until: 3, status: 4, code: 5, dev: 6, seen: 7, ver: 8, reg: 9, note: 10 };
const HEAD = ['Telefon', 'Ad / şirkət', 'Plan', 'Bitmə tarixi', 'Status', 'Hesab kodu', 'Cihaz ID', 'Son giriş', 'Versiya', 'Qeydiyyat', 'Qeyd'];
const PRIV = `__PEM__`;   // gizli açar yalnız Google hesabındakı nüsxədədir

const book = () => SpreadsheetApp.getActive();
const sh = n => book().getSheetByName(n);
const TZ = () => book().getSpreadsheetTimeZone();
function setting(name, def) {
  const s = sh('Ayarlar'); if (!s || s.getLastRow() < 1) return def;
  const v = s.getRange(1, 1, s.getLastRow(), 2).getValues();
  for (const r of v) if (String(r[0]).trim() === name) return r[1] === '' ? def : r[1];
  return def;
}
function digits(p) { let d = String(p || '').replace(/\D/g, ''); if (d.indexOf('994') === 0 && d.length === 12) d = d.slice(3); if (d.charAt(0) === '0') d = d.slice(1); return /^\d{9}$/.test(d) ? '994' + d : ''; }
function pretty(d) { return d ? '+994 ' + d.slice(3, 5) + ' ' + d.slice(5, 8) + ' ' + d.slice(8, 10) + ' ' + d.slice(10, 12) : ''; }
function ymd(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ(), 'yyyy-MM-dd');
  const s = String(v || ''); let m = s.match(/(\d{4})-(\d{2})-(\d{2})/); if (m) return m[0];
  m = s.match(/(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/); return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
}
const b64u = bytes => Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
function sign(dev, name, until) {
  const json = JSON.stringify({ v: 1, d: dev, n: name || '', u: until, t: Utilities.formatDate(new Date(), TZ(), 'yyyy-MM-dd'), s: 1 });
  const sig = Utilities.computeRsaSha256Signature(json, PRIV, Utilities.Charset.UTF_8);
  return b64u(Utilities.newBlob(json).getBytes()) + '.' + b64u(sig);
}
function newCode(rows) {
  const A = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', have = new Set(rows.map(r => String(r[C.code]).toUpperCase()));
  for (;;) { let c = ''; for (let i = 0; i < 6; i++) c += A.charAt(Math.floor(Math.random() * A.length)); if (!have.has(c)) return c; }
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function doGet() { return out({ ok: true, srv: 'DataX', t: new Date().toISOString() }); }
function doPost(e) { let q = {}; try { q = JSON.parse(e.postData.contents || '{}'); } catch (x) {} try { return out(handle(q)); } catch (x) { return out({ ok: false, err: 'Server xətası: ' + x.message }); } }

function handle(q) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (x) { return { ok: false, err: 'Server məşğuldur, bir az sonra yenidən cəhd et.' }; }
  try {
    const s = sh('Müştərilər'), n = s.getLastRow(), rows = n > 1 ? s.getRange(2, 1, n - 1, NC).getValues() : [];
    const dev = String(q.device || '').trim(), contact = String(setting('Əlaqə', '+994 50 597 15 67'));
    if (!/^DX-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}$/.test(dev)) return { ok: false, err: 'Cihaz kodu düzgün deyil.' };
    const now = new Date(), ver = String(q.ver || '').slice(0, 20);
    const find = f => { for (let i = 0; i < rows.length; i++) if (f(rows[i])) return i; return -1; };
    const resp = i => {
      const r = rows[i], until = ymd(r[C.until]), stop = /dayand/i.test(String(r[C.status]));
      return { ok: true, status: stop ? 'dayandirilib' : 'aktiv', until, plan: String(r[C.plan] || ''), name: String(r[C.name] || ''), phone: pretty(digits(r[C.phone])), code: String(r[C.code] || ''), contact,
        key: !stop && until ? sign(dev, String(r[C.name] || ''), until) : '' };
    };
    const touch = i => { s.getRange(i + 2, C.seen + 1, 1, 2).setValues([[now, ver]]); if (q.upd != null) s.getRange(i + 2, NC + 1).setValue(String(q.upd).slice(0, 120)); };   // L sütunu: yeniləyicinin son işləməsi
    if (q.a === 'sync') {
      const i = find(r => String(r[C.dev]).trim() === dev);
      if (i < 0) return { ok: false, reason: 'yoxdur', contact };
      touch(i); logUsage(rows[i], q.usage); return resp(i);
    }
    if (q.a === 'login') {
      const ph = digits(q.phone); if (!ph) return { ok: false, err: 'Telefon nömrəsini düzgün yaz, məsələn: 050 123 45 67' };
      let i = find(r => digits(r[C.phone]) === ph);
      if (i < 0) {
        if (/^x/i.test(String(setting('Yeni qeydiyyat açıqdır', 'Bəli')))) return { ok: false, err: 'Yeni qeydiyyat hazırda bağlıdır. Əlaqə: ' + contact };
        const j = find(r => String(r[C.dev]).trim() === dev);
        if (j >= 0) return { ok: false, err: 'Bu kompüter artıq ' + pretty(digits(rows[j][C.phone])) + ' nömrəsi ilə qeydiyyatdadır. Həmin nömrə ilə daxil ol və ya əlaqə saxla: ' + contact };
        const days = Number(setting('Sınaq günləri', 7)) || 7, until = new Date(now.getTime() + days * 864e5);
        const row = ["'" + pretty(ph), String(q.name || '').trim().slice(0, 80), 'Sınaq', until, 'aktiv', newCode(rows), dev, now, ver, now, ''];
        s.appendRow(row); const R = s.getLastRow();
        s.getRange(R, C.until + 1).setNumberFormat('dd.MM.yyyy'); s.getRange(R, C.seen + 1).setNumberFormat('dd.MM.yyyy HH:mm'); s.getRange(R, C.reg + 1).setNumberFormat('dd.MM.yyyy');
        rows.push(row); i = rows.length - 1;
        const r0 = resp(i); r0.isNew = true; return r0;
      }
      const r = rows[i];
      const boss = /rəhbər/i.test(String(r[C.plan]));                              // rəhbər hesabı: istənilən kompüterdən kodsuz
      if (String(r[C.dev]).trim() !== dev) {
        const code = boss ? String(r[C.code]).toUpperCase() : String(q.code || '').replace(/\s/g, '').toUpperCase();
        if (!code) return { ok: false, need: 'code', err: 'Bu nömrə ilə hesab artıq var. Başqa kompüterdən daxil olmaq üçün hesab kodunu yaz (köhnə kompüterdə DataX-in aşağısında görünür) və ya əlaqə saxla: ' + contact };
        if (code !== String(r[C.code]).toUpperCase()) return { ok: false, need: 'code', err: 'Hesab kodu düzgün deyil.' };
        const j = find(x => String(x[C.dev]).trim() === dev); if (j >= 0 && j !== i) { s.getRange(j + 2, C.dev + 1).setValue(''); rows[j][C.dev] = ''; }
        s.getRange(i + 2, C.dev + 1).setValue(dev); rows[i][C.dev] = dev;          // köhnə kompüter növbəti yoxlamada çıxış edəcək
      }
      if (q.name && !String(r[C.name]).trim()) { s.getRange(i + 2, C.name + 1).setValue(String(q.name).trim().slice(0, 80)); rows[i][C.name] = String(q.name).trim(); }
      touch(i); return resp(i);
    }
    return { ok: false, err: 'naməlum əmr' };
  } finally { lock.releaseLock(); }
}
function logUsage(r, usage) {
  if (!usage || typeof usage !== 'object') return;
  const u = sh('İstifadə'); if (!u) return;
  const now = new Date(), add = Object.keys(usage).filter(k => +usage[k] > 0).slice(0, 30).map(k => [now, "'" + String(r[C.phone] || ''), String(r[C.name] || ''), String(k).slice(0, 40), +usage[k]]);
  if (add.length) { u.getRange(u.getLastRow() + 1, 1, add.length, 5).setValues(add); u.getRange(u.getLastRow() - add.length + 1, 1, add.length, 1).setNumberFormat('dd.MM.yyyy HH:mm'); }
}

// Bir dəfə işə salınır: vərəqləri, başlıqları, seçim siyahılarını qurur
function qurasdir() {
  const b = book(); b.setSpreadsheetTimeZone('Asia/Baku');
  let s = sh('Müştərilər') || b.getSheets()[0]; s.setName('Müştərilər');
  s.getRange(1, 1, 1, NC).setValues([HEAD]).setFontWeight('bold').setBackground('#0d2229').setFontColor('#ffffff');
  s.setFrozenRows(1); s.getRange('A:A').setNumberFormat('@'); s.getRange('F:G').setNumberFormat('@');
  s.getRange('D2:D').setNumberFormat('dd.MM.yyyy'); s.getRange('H2:H').setNumberFormat('dd.MM.yyyy HH:mm'); s.getRange('J2:J').setNumberFormat('dd.MM.yyyy');
  s.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Sınaq', 'Aylıq', 'İllik'], true).setAllowInvalid(true).build());
  s.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['aktiv', 'dayandırılıb'], true).setAllowInvalid(false).build());
  [150, 200, 80, 110, 110, 100, 170, 140, 70, 100, 220].forEach((w, i) => s.setColumnWidth(i + 1, w));
  let a = sh('Ayarlar') || b.insertSheet('Ayarlar');
  a.getRange(1, 1, 3, 2).setValues([['Sınaq günləri', 7], ['Əlaqə', '+994 50 597 15 67'], ['Yeni qeydiyyat açıqdır', 'Bəli']]); a.setColumnWidth(1, 200); a.setColumnWidth(2, 200);
  let u = sh('İstifadə') || b.insertSheet('İstifadə');
  u.getRange(1, 1, 1, 5).setValues([['Tarix', 'Telefon', 'Ad / şirkət', 'Funksiya', 'Neçə dəfə']]).setFontWeight('bold'); u.setFrozenRows(1);
}

// ---- Müddəti cədvəldən uzatmaq ----
// 1) "Plan" xanasında "Aylıq" və ya "İllik" seçəndə bitmə tarixi özü 1 ay / 1 il uzanır (müddət hələ bitməyibsə — köhnə tarixdən, bitibsə — bu gündən)
// 2) Menyu: DataX → Seçilmiş müştərini 1 ay / 1 il uzat (eyni plan təkrar ödəniləndə)
function uzat(s, row, months) {
  const cell = s.getRange(row, C.until + 1), old = cell.getValue(), today = new Date(); today.setHours(0, 0, 0, 0);
  const base = old instanceof Date && old > today ? new Date(old.getTime()) : today;
  base.setMonth(base.getMonth() + months);
  cell.setValue(base).setNumberFormat('dd.MM.yyyy');
  if (String(s.getRange(row, C.status + 1).getValue()).trim() !== 'aktiv') s.getRange(row, C.status + 1).setValue('aktiv');
  return Utilities.formatDate(base, TZ(), 'dd.MM.yyyy');
}
function onEdit(e) {
  try {
    const r = e.range, s = r.getSheet();
    if (s.getName() !== 'Müştərilər' || r.getRow() < 2 || r.getNumRows() !== 1 || r.getColumn() !== C.plan + 1) return;
    const p = String(e.value || '').toLowerCase(), mo = /ayl/.test(p) ? 1 : /ill/.test(p) ? 12 : 0;
    if (!mo) return;
    const d = uzat(s, r.getRow(), mo);
    e.source.toast(String(s.getRange(r.getRow(), C.name + 1).getValue()) + ': bitmə tarixi ' + d, 'DataX — uzadıldı', 6);
  } catch (x) {}
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('DataX').addItem('Seçilmiş müştərini 1 ay uzat', 'uzat1Ay').addItem('Seçilmiş müştərini 1 il uzat', 'uzat1Il').addToUi();
}
function uzatSecilen(mo) {
  const s = SpreadsheetApp.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (s.getName() !== 'Müştərilər') { ui.alert('Əvvəl "Müştərilər" vərəqində müştərinin sətrinə klik et.'); return; }
  const row = s.getActiveRange().getRow(); if (row < 2) { ui.alert('Müştərinin sətrinə klik et.'); return; }
  const name = s.getRange(row, C.name + 1).getValue() || s.getRange(row, C.phone + 1).getDisplayValue();
  if (ui.alert(name + ' — müddət ' + (mo === 12 ? '1 il' : '1 ay') + ' uzadılsın?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert(name + ': yeni bitmə tarixi ' + uzat(s, row, mo));
}
function uzat1Ay() { uzatSecilen(1); }
function uzat1Il() { uzatSecilen(12); }
