/**
 * SAFINAH — Sistem Automasi Finansial & Inventori Naskah
 * by Hikmah Pendita Resources · Versi 1.0 · 2026
 *
 * Fasa 1
 * Aliran Tunai · Rekod (Jualan / Duit Masuk / Duit Keluar) · Import Wasepme
 * Agihan Profit First · Katalog Kitab
 *
 * Hikmah Pendita Resources · Teraju Akal & Bina Budi
 */

const TZ = 'Asia/Kuala_Lumpur';
const SH = { TX: 'Transaksi', KAT: 'Katalog', PELAN: 'Pelan', TET: 'Tetapan', AGIH: 'Agihan' };

const TX_COLS = ['ID', 'Tarikh', 'Jenis', 'Kategori', 'Keterangan', 'Jumlah', 'Bank', 'Bank Ke',
  'Kos Kitab', 'Untung Sebenar', 'Kaedah Bayaran', 'Saluran', 'Sumber', 'No Rujukan',
  'Butiran Kitab', 'Resit', 'Status', 'Bendera', 'Direkod Oleh', 'Masa Rekod', 'Agihan'];
const KAT_COLS = ['Judul', 'Harga Jual', 'Harga LE', 'Berat', 'Kos', 'Dikemaskini'];
const AGIH_COLS = ['ID', 'Tarikh', 'Jualan Dikira', 'Untung Sebenar', 'Profit', 'Cukai', 'Gaji Pemilik', 'Belanja Operasi', 'Direkod Oleh'];
const PELAN_COLS = ['Bulan', 'Target Jualan', 'Target Untung Sebenar', 'Margin %', 'Profit %', 'Cukai %', 'Gaji %', 'Belanja Operasi %'];

const BANKS = ['Maybank', 'Hong Leong', 'OCBC', 'Bank Islam', 'Tunai']; // Tunai = duit tunai di tangan (belum bank-in)
const SALES_CATS = ['Jualan', 'Bayaran invois institusi', 'Deposit / pre-order'];
const IN_CATS = ['Bayaran invois institusi', 'Deposit / pre-order', 'Pulangan duit', 'Lain-lain (masuk)'];
const OUT_CATS = [
  { k: 'Beli kitab / kargo Mesir', bank: 'Maybank', grp: 'Duit stok' },
  { k: 'Gaji staf', bank: 'Hong Leong', grp: 'Belanja operasi' },
  { k: 'Pos & kurier', bank: 'Hong Leong', grp: 'Belanja operasi' },
  { k: 'Sewa gerai / event', bank: 'Hong Leong', grp: 'Belanja operasi' },
  { k: 'Bil elektrik, air, internet', bank: 'Hong Leong', grp: 'Belanja operasi' },
  { k: 'Kenderaan & minyak', bank: 'Hong Leong', grp: 'Belanja operasi' },
  { k: 'Bayaran loan', bank: 'Hong Leong', grp: 'Belanja operasi' },
  { k: 'Lain-lain (keluar)', bank: 'Hong Leong', grp: 'Belanja operasi' }
];

/* =========================================================
 * 1. PASANG SISTEM — jalankan sekali sahaja
 * ========================================================= */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SH.TX, TX_COLS);
  ensureSheet_(ss, SH.KAT, KAT_COLS);
  ensureSheet_(ss, SH.AGIH, AGIH_COLS);

  const pel = ensureSheet_(ss, SH.PELAN, PELAN_COLS);
  if (pel.getLastRow() < 2) {
    pel.getRange('A:A').setNumberFormat('@');
    pel.getRange(2, 1, 5, 8).setValues([
      ['2026-10', 14000, 5600, 40, 2, 2, 45, 51],
      ['2026-11', 15500, 6355, 41, 6, 4, 45, 45],
      ['2026-12', 17000, 7140, 42, 10, 6, 44, 40],
      ['2027-01', 16000, 6880, 43, 6, 5, 48, 41],
      ['2027-02', 18200, 8008, 44, 8, 6, 50, 36]
    ]);
  }

  const tet = ensureSheet_(ss, SH.TET, ['Tetapan', 'Nilai', 'Nota']);
  if (tet.getLastRow() < 2) {
    tet.getRange('B:B').setNumberFormat('@');
    const rows = [
      ['PIN Pemilik', '1234', 'TUKAR SEGERA. Nombor sahaja.'],
      ['PIN Staf', 'Staf 1:1111, Staf 2:2222', 'Format Nama:PIN, pisah dengan koma'],
      ['Rate LE', '11', 'Rate selamat LE bagi RM1'],
      ['Pos ringan', '10', '500 g ke bawah (RM)'],
      ['Pos berat', '20', 'Lebih 500 g hingga 1 kg (RM)'],
      ['Markup %', '80', ''],
      ['Tarikh baki awal', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), 'Format yyyy-mm-dd. Baki di bawah ialah baki sebenar pada tarikh ini.']
    ];
    rows.push(['Tarikh agihan terakhir', '', 'yyyy-mm-dd. Jualan SELEPAS tarikh ini dikira untuk agihan pertama dalam app (contoh 2026-10-10).']);
    BANKS.forEach(b => rows.push(['Baki awal ' + b, '0', b === 'Tunai' ? 'Duit tunai di tangan yang belum dimasukkan ke bank' : 'Baki sebenar dalam akaun pada tarikh baki awal']));
    rows.push(['Folder Resit ID', '', 'Diisi automatik']);
    tet.getRange(2, 1, rows.length, 3).setValues(rows);
    tet.setColumnWidth(1, 200); tet.setColumnWidth(2, 260); tet.setColumnWidth(3, 360);
  } else {
    // Tambah tetapan baru jika sheet Tetapan dibuat oleh versi lama
    const have = getSettings_();
    [['Tarikh agihan terakhir', '', 'yyyy-mm-dd. Jualan SELEPAS tarikh ini dikira untuk agihan pertama dalam app.'],
     ['Baki awal OCBC', '0', 'Satu akaun OCBC untuk Profit dan Cukai']].forEach(r => { if (have[r[0]] === undefined) tet.appendRow(r); });
  }
  SpreadsheetApp.getUi().alert('SAFINAH siap dipasang. Tukar PIN dan isi baki awal bank dalam sheet "Tetapan".');
}

function ensureSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#E0ECF7');
    sh.setFrozenRows(1);
  }
  return sh;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('SAFINAH')
    .addItem('Pasang / semak sistem', 'setup')
    .addToUi();
}

/* =========================================================
 * 2. WEB APP
 * ========================================================= */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SAFINAH · Hikmah Pendita Resources')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* =========================================================
 * 3. PEMBANTU
 * ========================================================= */
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function fmt_(d, p) { return d instanceof Date ? Utilities.formatDate(d, TZ, p || 'yyyy-MM-dd') : ''; }
function num_(v) { const n = Number(String(v).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : 0; }
function r2_(n) { return Math.round(n * 100) / 100; }
function newId_(p) { return (p || 'T') + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase(); }
function normTitle_(t) { return String(t || '').toUpperCase().replace(/[^A-Z0-9 ]/g, '').replace(/\s+/g, ' ').trim(); }
function parseDate_(s) {
  if (s instanceof Date) return s;
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/);
  if (!m) return new Date();
  return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 12), +(m[5] || 0));
}

function getSettings_() {
  const vals = sheet_(SH.TET).getDataRange().getValues();
  const s = {};
  vals.slice(1).forEach(r => { if (r[0]) s[String(r[0]).trim()] = r[1]; });
  return s;
}
function setSetting_(key, value) {
  const sh = sheet_(SH.TET);
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (vals[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  sh.appendRow([key, value, '']);
}

function checkPin_(pin) {
  const s = getSettings_();
  const p = String(pin || '').trim();
  if (!p) throw new Error('Sila masukkan PIN.');
  if (p === String(s['PIN Pemilik']).trim()) return { role: 'pemilik', name: 'Pemilik' };
  const list = String(s['PIN Staf'] || '').split(',');
  for (const item of list) {
    const parts = item.split(':');
    if (parts.length === 2 && parts[1].trim() === p) return { role: 'staf', name: parts[0].trim() };
  }
  throw new Error('PIN salah.');
}
function ownerOnly_(u) { if (u.role !== 'pemilik') throw new Error('Hanya pemilik boleh buka bahagian ini.'); }

function readTx_() {
  const sh = sheet_(SH.TX);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, TX_COLS.length).getValues();
  return vals.map((r, i) => {
    const o = { _row: i + 2 };
    TX_COLS.forEach((c, j) => { o[c] = r[j]; });
    return o;
  }).filter(o => o.ID);
}
function txOut_(t) {
  return {
    id: t.ID, tarikh: fmt_(t.Tarikh), jenis: t.Jenis, kategori: t.Kategori, keterangan: t.Keterangan,
    jumlah: num_(t.Jumlah), bank: t.Bank, bankKe: t['Bank Ke'],
    kos: t['Kos Kitab'] === '' ? null : num_(t['Kos Kitab']),
    untung: t['Untung Sebenar'] === '' ? null : num_(t['Untung Sebenar']),
    kaedah: t['Kaedah Bayaran'], saluran: t.Saluran, sumber: t.Sumber, ruj: String(t['No Rujukan'] || ''),
    resit: t.Resit, status: t.Status, bendera: t.Bendera, oleh: t['Direkod Oleh']
  };
}
function txRow_(o) { return TX_COLS.map(c => (o[c] === undefined || o[c] === null ? '' : o[c])); }
function isSale_(t) { return t.Jenis === 'Masuk' && SALES_CATS.indexOf(t.Kategori) > -1; }

function readKatalog_() {
  const sh = sheet_(SH.KAT);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, KAT_COLS.length).getValues()
    .map((r, i) => ({ _row: i + 2, judul: r[0], hargaJual: num_(r[1]), le: r[2] === '' ? '' : num_(r[2]), berat: r[3], kos: r[4] === '' ? 0 : num_(r[4]) }))
    .filter(k => k.judul);
}
function katMap_(kat) { const m = {}; kat.forEach(k => { m[normTitle_(k.judul)] = k; }); return m; }

function getPelan_(ym) {
  const vals = sheet_(SH.PELAN).getDataRange().getValues().slice(1);
  let row = vals.find(r => String(r[0]).trim() === ym || fmt_(r[0], 'yyyy-MM') === ym);
  if (!row && vals.length) row = vals[vals.length - 1];
  if (!row) return { bulan: ym, target: 0, targetUntung: 0, margin: 40, p: [5, 15, 50, 30] };
  return { bulan: ym, target: num_(row[1]), targetUntung: num_(row[2]), margin: num_(row[3]), p: [num_(row[4]), num_(row[5]), num_(row[6]), num_(row[7])] };
}

function balances_(txs, s) {
  const b = {};
  BANKS.forEach(k => { b[k] = num_(s['Baki awal ' + k]); });
  if (s['Baki awal OCBC'] === undefined) b.OCBC = num_(s['Baki awal OCBC Profit']) + num_(s['Baki awal OCBC Cukai']); // tetapan versi lama
  const sv = s['Tarikh baki awal'];
  const start = sv ? fmt_(parseDate_(sv instanceof Date ? fmt_(sv) : String(sv).slice(0, 10))) : '';
  txs.forEach(t => {
    if (start && fmt_(t.Tarikh) < start) return;
    const a = num_(t.Jumlah);
    if (t.Jenis === 'Masuk') b[t.Bank] = (b[t.Bank] || 0) + a;
    else if (t.Jenis === 'Keluar') b[t.Bank] = (b[t.Bank] || 0) - a;
    else if (t.Jenis === 'Pindahan') { b[t.Bank] = (b[t.Bank] || 0) - a; b[t['Bank Ke']] = (b[t['Bank Ke']] || 0) + a; }
  });
  Object.keys(b).forEach(k => { b[k] = r2_(b[k]); });
  return b;
}

function nextAgihan_(d) {
  const day = +fmt_(d, 'd');
  if (day < 10) return fmt_(d, 'yyyy-MM') + '-10';
  if (day < 25) return fmt_(d, 'yyyy-MM') + '-25';
  const n = new Date(d.getFullYear(), d.getMonth() + 1, 10);
  return fmt_(n);
}

function saveReceipt_(resit) {
  if (!resit || !resit.data) return '';
  const s = getSettings_();
  let folder;
  try { folder = s['Folder Resit ID'] ? DriveApp.getFolderById(s['Folder Resit ID']) : null; } catch (e) { folder = null; }
  if (!folder) { folder = DriveApp.createFolder('SAFINAH Resit'); setSetting_('Folder Resit ID', folder.getId()); }
  const blob = Utilities.newBlob(Utilities.base64Decode(resit.data), resit.mime || 'image/jpeg', resit.name || ('resit-' + Date.now() + '.jpg'));
  return folder.createFile(blob).getUrl();
}

/* =========================================================
 * 4. API — dipanggil dari app
 * ========================================================= */
function api_login(pin) {
  const u = checkPin_(pin);
  const kat = readKatalog_().map(k => ({ judul: k.judul, harga: k.hargaJual }));
  return { user: u, banks: BANKS, inCats: IN_CATS, outCats: OUT_CATS, katalog: kat, today: fmt_(new Date()) };
}

function api_dashboard(pin) {
  const u = checkPin_(pin);
  const now = new Date();
  const today = fmt_(now);
  const txs = readTx_();
  if (u.role === 'staf') {
    const mine = txs.filter(t => t['Direkod Oleh'] === u.name && fmt_(t.Tarikh) === today);
    return { role: 'staf', mine: mine.map(txOut_) };
  }
  const s = getSettings_();
  const ym = fmt_(now, 'yyyy-MM');
  const pelan = getPelan_(ym);
  const sales = txs.filter(t => isSale_(t) && fmt_(t.Tarikh, 'yyyy-MM') === ym);
  const salesSum = sales.reduce((a, t) => a + num_(t.Jumlah), 0);
  const known = sales.filter(t => t['Untung Sebenar'] !== '');
  const untungKnown = known.reduce((a, t) => a + num_(t['Untung Sebenar']), 0);
  const unknownSales = sales.filter(t => t['Untung Sebenar'] === '').reduce((a, t) => a + num_(t.Jumlah), 0);
  const weekAgo = new Date(now.getTime() - 7 * 86400000);
  const week = txs.filter(t => isSale_(t) && t.Tarikh >= weekAgo).reduce((a, t) => a + num_(t.Jumlah), 0);
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  return {
    role: 'pemilik', bulan: fmt_(now, 'MMMM yyyy'), pelan,
    jualan: r2_(salesSum), untungKnown: r2_(untungKnown), untungAnggaran: r2_(untungKnown + unknownSales * pelan.margin / 100),
    unknownSales: r2_(unknownSales), minggu: r2_(week), hariLagi: daysInMonth - now.getDate(),
    perluSemak: txs.filter(t => t.Status === 'Perlu semak').length,
    baki: balances_(txs, s), agihanSeterusnya: nextAgihan_(now)
  };
}

function api_listTx(pin, ym) {
  const u = checkPin_(pin); ownerOnly_(u);
  const month = ym || fmt_(new Date(), 'yyyy-MM');
  const txs = readTx_().filter(t => fmt_(t.Tarikh, 'yyyy-MM') === month)
    .sort((a, b) => (b.Tarikh - a.Tarikh) || (String(b['Masa Rekod']) > String(a['Masa Rekod']) ? 1 : -1));
  const s = getSettings_();
  return { bulan: month, rows: txs.map(txOut_), baki: balances_(readTx_(), s) };
}

/** Simpan rekod dari borang Rekod Baru */
function api_saveTx(pin, rec) {
  const u = checkPin_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const flags = [];
    const o = { ID: newId_('T'), Tarikh: parseDate_(rec.tarikh), Sumber: 'Manual', 'Direkod Oleh': u.name, 'Masa Rekod': new Date() };
    const jumlah = r2_(num_(rec.jumlah));
    if (rec.jenis === 'jualan') {
      const items = (rec.items || []).filter(i => num_(i.q) > 0);
      if (!items.length) throw new Error('Tiada kitab dipilih.');
      const km = katMap_(readKatalog_());
      let total = 0, kos = 0, kosOk = true;
      items.forEach(i => {
        total += num_(i.q) * num_(i.harga);
        const k = km[normTitle_(i.judul)];
        if (k && k.kos > 0) kos += k.kos * num_(i.q); else kosOk = false;
      });
      total = r2_(total - num_(rec.diskaun));
      Object.assign(o, {
        Jenis: 'Masuk', Kategori: 'Jualan', Jumlah: total,
        Bank: rec.kaedah === 'Tunai' ? 'Tunai' : 'Maybank', 'Kaedah Bayaran': rec.kaedah,
        Saluran: rec.saluran, Keterangan: (rec.saluran || '') + ' — ' + items.map(i => i.q + '× ' + i.judul).join(', '),
        'Butiran Kitab': JSON.stringify(items.map(i => ({ judul: i.judul, q: num_(i.q), harga: num_(i.harga) })))
      });
      if (kosOk) { o['Kos Kitab'] = r2_(kos); o['Untung Sebenar'] = r2_(total - kos); }
      else flags.push('Kos kitab tiada — isi di Katalog');
      if (rec.kaedah === 'Tunai') flags.push('Tunai belum masuk Maybank');
      if (rec.kaedah !== 'Tunai' && !(rec.resit && rec.resit.data)) flags.push('Tiada bukti bayaran');
    } else if (rec.jenis === 'masuk') {
      if (!(jumlah > 0)) throw new Error('Jumlah tidak sah.');
      if (IN_CATS.indexOf(rec.kategori) < 0) throw new Error('Pilih kategori.');
      Object.assign(o, { Jenis: 'Masuk', Kategori: rec.kategori, Jumlah: jumlah, Bank: rec.bank, Keterangan: rec.keterangan || rec.kategori });
      if (rec.bank !== 'Maybank') flags.push('Duit masuk sepatutnya ke Maybank');
    } else if (rec.jenis === 'keluar') {
      if (!(jumlah > 0)) throw new Error('Jumlah tidak sah.');
      const cat = OUT_CATS.find(c => c.k === rec.kategori);
      if (!cat) throw new Error('Pilih kategori.');
      Object.assign(o, { Jenis: 'Keluar', Kategori: rec.kategori, Jumlah: jumlah, Bank: rec.bank, Keterangan: rec.keterangan || rec.kategori, Saluran: cat.grp });
      if (rec.bank !== cat.bank) flags.push(cat.bank === 'Hong Leong' ? 'Bank salah — belanja operasi mesti dari Hong Leong' : 'Bank salah — duit stok dari Maybank sahaja');
      if (!(rec.resit && rec.resit.data)) flags.push('Tiada resit');
    } else if (rec.jenis === 'pindahan') {
      ownerOnly_(u);
      if (!(jumlah > 0)) throw new Error('Jumlah tidak sah.');
      Object.assign(o, { Jenis: 'Pindahan', Kategori: rec.kategori || 'Pindahan', Jumlah: jumlah, Bank: rec.bank, 'Bank Ke': rec.bankKe, Keterangan: rec.keterangan || ('Pindahan ' + rec.bank + ' → ' + rec.bankKe) });
    } else throw new Error('Jenis rekod tidak dikenali.');

    if (rec.resit && rec.resit.data) o.Resit = saveReceipt_(rec.resit);
    o.Bendera = flags.join('; ');
    o.Status = flags.length ? 'Perlu semak' : 'Lengkap';
    sheet_(SH.TX).appendRow(txRow_(o));
    return txOut_(o);
  } finally { lock.releaseLock(); }
}

/** Lengkapkan rekod: tambah resit / tukar kategori / tukar bank */
function api_completeTx(pin, id, patch) {
  const u = checkPin_(pin); ownerOnly_(u);
  const t = readTx_().find(x => x.ID === id);
  if (!t) throw new Error('Rekod tidak dijumpai.');
  const sh = sheet_(SH.TX);
  const col = c => TX_COLS.indexOf(c) + 1;
  let flags = String(t.Bendera || '').split('; ').filter(Boolean);
  if (patch.resit && patch.resit.data) {
    const url = saveReceipt_(patch.resit);
    sh.getRange(t._row, col('Resit')).setValue(url);
    flags = flags.filter(f => f !== 'Tiada resit' && f !== 'Tiada bukti bayaran');
  }
  if (patch.bank) {
    sh.getRange(t._row, col('Bank')).setValue(patch.bank);
    flags = flags.filter(f => f.indexOf('Bank salah') !== 0 && f.indexOf('Duit masuk sepatutnya') !== 0);
  }
  if (patch.selesai) flags = flags.filter(f => f !== 'Tunai belum masuk Maybank' && f.indexOf('Kos nampak') !== 0);
  sh.getRange(t._row, col('Bendera')).setValue(flags.join('; '));
  sh.getRange(t._row, col('Status')).setValue(flags.length ? 'Perlu semak' : 'Lengkap');
  return { ok: true, status: flags.length ? 'Perlu semak' : 'Lengkap', bendera: flags.join('; ') };
}

/* ---------- Import Wasepme ---------- */
function planImport_(orders, includePos, existing, km) {
  const rows = [];
  const sum = { import: 0, langkau: 0, ada: 0, jualan: 0, pos: 0, masuk: 0, tunai: 0, tiadaKos: 0, bayaran: {} };
  const baru = {};
  orders.forEach(o => {
    let d = 'import', sebab = '';
    if (/batal/i.test(o.status)) { d = 'langkau'; sebab = 'Dibatalkan'; }
    else if (existing[o.no]) { d = 'ada'; sebab = 'Sudah pernah diimport'; }
    else if (o.isPos && !includePos) { d = 'langkau'; sebab = 'POS — jualan kedai/event direkod di Rekod Jualan'; }
    const flags = [];
    let kos = '';
    if (d === 'import') {
      if (o.cost > 0) {
        kos = o.cost;
        if (o.subtotal > 0 && o.cost / o.subtotal < 0.25) flags.push('Kos nampak terlalu rendah (RM' + o.cost.toFixed(2) + ' untuk jualan RM' + o.subtotal.toFixed(2) + ')');
      } else {
        let k = 0, ok = o.items.length > 0;
        o.items.forEach(it => {
          const c = km[normTitle_(it.judul)];
          if (c && c.kos > 0) k += c.kos * it.q; else ok = false;
          if (!c && !baru[normTitle_(it.judul)]) baru[normTitle_(it.judul)] = { judul: it.judul, harga: it.q ? r2_(it.jumlah / it.q) : it.jumlah };
        });
        if (ok) kos = r2_(k); else flags.push('Kos kitab tiada — isi di Katalog');
      }
      if (/tunai/i.test(o.payment)) flags.push('Tunai belum masuk Maybank');
      const pos = r2_(Math.max(0, o.total - o.subtotal));
      sum.import++; sum.jualan += o.subtotal; sum.pos += pos; sum.masuk += o.subtotal + pos;
      if (/tunai/i.test(o.payment)) sum.tunai += o.subtotal;
      if (kos === '') sum.tiadaKos++;
      sum.bayaran[o.payment] = (sum.bayaran[o.payment] || 0) + 1;
    } else sum[d]++;
    rows.push(Object.assign({}, o, { keputusan: d, sebab: sebab, kos: kos, bendera: flags }));
  });
  ['jualan', 'pos', 'masuk', 'tunai'].forEach(k => { sum[k] = r2_(sum[k]); });
  return { rows: rows, ringkasan: sum, kitabBaru: Object.keys(baru).map(k => baru[k]) };
}

function existingWasepme_() {
  const m = {};
  readTx_().forEach(t => { if (t.Sumber === 'Wasepme' && t.Kategori === 'Jualan') m[String(t['No Rujukan'])] = true; });
  return m;
}

function api_previewImport(pin, orders, includePos) {
  const u = checkPin_(pin); ownerOnly_(u);
  return planImport_(orders, !!includePos, existingWasepme_(), katMap_(readKatalog_()));
}

function api_commitImport(pin, orders, includePos) {
  const u = checkPin_(pin); ownerOnly_(u);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const plan = planImport_(orders, !!includePos, existingWasepme_(), katMap_(readKatalog_()));
    const now = new Date();
    const out = [];
    plan.rows.filter(r => r.keputusan === 'import').forEach(r => {
      const tunai = /tunai/i.test(r.payment);
      const bank = tunai ? 'Tunai' : 'Maybank';
      const tarikh = parseDate_(r.tarikh);
      const items = r.items.map(i => ({ judul: i.judul, q: i.q, harga: i.q ? r2_(i.jumlah / i.q) : i.jumlah }));
      out.push(txRow_({
        ID: newId_('W'), Tarikh: tarikh, Jenis: 'Masuk', Kategori: 'Jualan',
        Keterangan: 'Wasepme #' + r.no + ' — ' + items.map(i => i.q + '× ' + i.judul).join(', '),
        Jumlah: r2_(r.subtotal), Bank: bank, 'Kos Kitab': r.kos,
        'Untung Sebenar': r.kos === '' ? '' : r2_(r.subtotal - r.kos),
        'Kaedah Bayaran': r.payment, Saluran: r.isPos ? 'Wasepme POS' : 'Wasepme online', Sumber: 'Wasepme',
        'No Rujukan': r.no, 'Butiran Kitab': JSON.stringify(items),
        Status: r.bendera.length ? 'Perlu semak' : 'Lengkap', Bendera: r.bendera.join('; '),
        'Direkod Oleh': 'Import Wasepme', 'Masa Rekod': now
      }));
      const pos = r2_(Math.max(0, r.total - r.subtotal));
      if (pos > 0) out.push(txRow_({
        ID: newId_('W'), Tarikh: tarikh, Jenis: 'Masuk', Kategori: 'Caj pos dikutip',
        Keterangan: 'Caj pos Wasepme #' + r.no, Jumlah: pos, Bank: bank, 'Kaedah Bayaran': r.payment,
        Saluran: 'Wasepme online', Sumber: 'Wasepme', 'No Rujukan': r.no, Status: 'Lengkap',
        'Direkod Oleh': 'Import Wasepme', 'Masa Rekod': now
      }));
    });
    if (out.length) {
      const sh = sheet_(SH.TX);
      sh.getRange(sh.getLastRow() + 1, 1, out.length, TX_COLS.length).setValues(out);
    }
    if (plan.kitabBaru.length) {
      const ksh = sheet_(SH.KAT);
      const rows = plan.kitabBaru.map(k => [k.judul, k.harga, '', '', '', now]);
      ksh.getRange(ksh.getLastRow() + 1, 1, rows.length, KAT_COLS.length).setValues(rows);
    }
    return { diimport: plan.ringkasan.import, jualan: plan.ringkasan.jualan, kitabBaru: plan.kitabBaru.length, tunai: plan.ringkasan.tunai };
  } finally { lock.releaseLock(); }
}

/* ---------- Agihan Profit First ---------- */
/** Jualan yang belum diagih — hanya sejak "Tarikh baki awal" (tarikh sistem mula digunakan) */
function pendingSales_(txs) {
  const s = getSettings_();
  const day = v => (v ? fmt_(parseDate_(String(v instanceof Date ? fmt_(v) : v).slice(0, 10))) : '');
  const lastAg = day(s['Tarikh agihan terakhir']);
  const start = day(s['Tarikh baki awal']);
  return txs.filter(t => {
    if (!isSale_(t) || t.Agihan) return false;
    const d = fmt_(t.Tarikh);
    if (lastAg) return d > lastAg;
    return !start || d >= start;
  });
}

function api_agihanInfo(pin) {
  const u = checkPin_(pin); ownerOnly_(u);
  const txs = readTx_();
  const now = new Date();
  const pelan = getPelan_(fmt_(now, 'yyyy-MM'));
  const p = pendingSales_(txs);
  const bank = p.filter(t => t.Bank !== 'Tunai');
  const sumOf = arr => r2_(arr.reduce((a, t) => a + num_(t.Jumlah), 0));
  const untungKnown = r2_(bank.filter(t => t['Untung Sebenar'] !== '').reduce((a, t) => a + num_(t['Untung Sebenar']), 0));
  const unknown = sumOf(bank.filter(t => t['Untung Sebenar'] === ''));
  return {
    pelan, jualan: sumOf(bank), bilangan: bank.length, untungKnown, jualanTanpaKos: unknown,
    untungCadangan: r2_(untungKnown + unknown * pelan.margin / 100),
    tunaiBelumBank: sumOf(p.filter(t => t.Bank === 'Tunai')),
    baki: balances_(txs, getSettings_()), seterusnya: nextAgihan_(now), ocbc: ocbcSplit_()
  };
}

/** OCBC satu akaun: pecahan Profit & Cukai dikira dari jumlah agihan dalam SAFINAH */
function ocbcSplit_() {
  const sh = sheet_(SH.AGIH);
  const last = sh.getLastRow();
  if (last < 2) return { profit: 0, cukai: 0 };
  const v = sh.getRange(2, 5, last - 1, 2).getValues();
  return { profit: r2_(v.reduce((a, r) => a + num_(r[0]), 0)), cukai: r2_(v.reduce((a, r) => a + num_(r[1]), 0)) };
}

function api_saveAgihan(pin, untung) {
  const u = checkPin_(pin); ownerOnly_(u);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const amt = r2_(num_(untung));
    if (!(amt > 0)) throw new Error('Untung sebenar tidak sah.');
    const now = new Date();
    const pelan = getPelan_(fmt_(now, 'yyyy-MM'));
    const id = newId_('A');
    const parts = [
      ['Profit', 'OCBC', pelan.p[0]], ['Cukai', 'OCBC', pelan.p[1]],
      ['Gaji pemilik', 'Bank Islam', pelan.p[2]], ['Belanja operasi', 'Hong Leong', pelan.p[3]]
    ].map(x => ({ nama: x[0], ke: x[1], pct: x[2], jumlah: r2_(amt * x[2] / 100) }));
    const txs = readTx_();
    const pending = pendingSales_(txs).filter(t => t.Bank !== 'Tunai');
    const jualan = r2_(pending.reduce((a, t) => a + num_(t.Jumlah), 0));
    const sh = sheet_(SH.TX);
    const rows = parts.filter(p => p.jumlah > 0).map(p => txRow_({
      ID: newId_('P'), Tarikh: now, Jenis: 'Pindahan', Kategori: 'Agihan Profit First',
      Keterangan: 'Agihan ' + p.nama + ' ' + p.pct + '%', Jumlah: p.jumlah, Bank: 'Maybank', 'Bank Ke': p.ke,
      Sumber: 'Agihan', 'No Rujukan': id, Status: 'Lengkap', 'Direkod Oleh': u.name, 'Masa Rekod': now, Agihan: id
    }));
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, TX_COLS.length).setValues(rows);
    const colA = TX_COLS.indexOf('Agihan') + 1;
    pending.forEach(t => sh.getRange(t._row, colA).setValue(id));
    sheet_(SH.AGIH).appendRow([id, now, jualan, amt].concat(parts.map(p => p.jumlah)).concat([u.name]));
    return { id, parts, jualan };
  } finally { lock.releaseLock(); }
}

/* ---------- Katalog ---------- */
function api_katalog(pin) {
  const u = checkPin_(pin); ownerOnly_(u);
  const s = getSettings_();
  return {
    rate: num_(s['Rate LE']) || 11, posRingan: num_(s['Pos ringan']) || 10, posBerat: num_(s['Pos berat']) || 20, markup: num_(s['Markup %']) || 80,
    rows: readKatalog_().map(k => ({ judul: k.judul, hargaJual: k.hargaJual, le: k.le, berat: k.berat, kos: k.kos }))
  };
}

function api_saveKatalog(pin, row) {
  const u = checkPin_(pin); ownerOnly_(u);
  const sh = sheet_(SH.KAT);
  const kat = readKatalog_();
  const key = normTitle_(row.judul);
  const found = kat.find(k => normTitle_(k.judul) === key);
  const vals = [row.judul, num_(row.hargaJual), row.le === '' ? '' : num_(row.le), row.berat || '', r2_(num_(row.kos)), new Date()];
  if (found) sh.getRange(found._row, 1, 1, KAT_COLS.length).setValues([vals]);
  else sh.appendRow(vals);
  const fixed = recalcCosts_();
  return { ok: true, dibetulkan: fixed };
}

/** Isi semula kos untuk jualan yang belum ada kos, bila katalog dikemas kini */
function recalcCosts_() {
  const km = katMap_(readKatalog_());
  const sh = sheet_(SH.TX);
  const col = c => TX_COLS.indexOf(c) + 1;
  let n = 0;
  readTx_().forEach(t => {
    if (t.Kategori !== 'Jualan' || t['Kos Kitab'] !== '' || !t['Butiran Kitab']) return;
    let items; try { items = JSON.parse(t['Butiran Kitab']); } catch (e) { return; }
    let kos = 0, ok = items.length > 0;
    items.forEach(i => { const k = km[normTitle_(i.judul)]; if (k && k.kos > 0) kos += k.kos * i.q; else ok = false; });
    if (!ok) return;
    kos = r2_(kos);
    const flags = String(t.Bendera || '').split('; ').filter(f => f && f.indexOf('Kos kitab tiada') !== 0);
    sh.getRange(t._row, col('Kos Kitab'), 1, 2).setValues([[kos, r2_(num_(t.Jumlah) - kos)]]);
    sh.getRange(t._row, col('Status'), 1, 2).setValues([[flags.length ? 'Perlu semak' : 'Lengkap', flags.join('; ')]]);
    n++;
  });
  return n;
}
