/**
 * haloeditin — Apps Script web app backend
 * Run setup() first; it prepares the database spreadsheet and Drive folder.
 */
const CFG = {
  sheets: {
    Users: ['Email', 'Nama', 'Role', 'Aktif', 'Dibuat', 'PasswordSalt', 'PasswordHash'],
    Clients: ['ClientID', 'Nama', 'Kontak', 'Email', 'Catatan', 'Dibuat'],
    Services: ['ServiceID', 'Nama', 'Deskripsi', 'HargaDasar', 'Aktif', 'Kategori', 'Satuan', 'HargaTiers', 'Catatan'],
    Orders: ['OrderID', 'Dibuat', 'ClientID', 'NamaKlien', 'Judul', 'JenisDokumen', 'Deadline', 'Status', 'Total', 'Catatan', 'FolderID'],
    OrderItems: ['ItemID', 'OrderID', 'ServiceID', 'NamaLayanan', 'Qty', 'Harga', 'Subtotal'],
    Tasks: ['TaskID', 'OrderID', 'EditorEmail', 'Status', 'Deadline', 'Catatan', 'Diperbarui'],
    Payments: ['PaymentID', 'OrderID', 'Tanggal', 'Nominal', 'Metode', 'Status', 'Catatan'],
    ActivityLog: ['Waktu', 'Email', 'Aksi', 'OrderID', 'Detail']
  },
  orderStatuses: ['Baru', 'Menunggu Pembayaran', 'Siap Dikerjakan', 'Dikerjakan', 'Pemeriksaan', 'Revisi', 'Selesai', 'Dibatalkan'],
  paymentStatuses: ['Belum Bayar', 'DP', 'Lunas', 'Refund'],
  taskStatuses: ['Belum Mulai', 'Dikerjakan', 'Menunggu Pemeriksaan', 'Revisi', 'Selesai']
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('haloeditin — Manajemen Jasa Dokumen')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Catalog values imported from the supplied Editin price list (1 June 2024).
const PRICE_TIERS = ['Reguler · 3 hari','Reguler · 2–3 hari','Reguler · 1–2 hari','Express · 2 hari','Express · 1 hari','Express · 12 jam','Express · 5–6 jam','Express · 3–4 jam','Express · <3 jam'];

function setup() {
  const settings = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActive();
  const existingDbId = settings.getProperty('DB_SPREADSHEET_ID');
  if (!ss && existingDbId) ss = SpreadsheetApp.openById(existingDbId);
  if (!ss) ss = SpreadsheetApp.create('haloeditin — Database');
  settings.setProperty('DB_SPREADSHEET_ID', ss.getId());
  Object.keys(CFG.sheets).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(CFG.sheets[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, CFG.sheets[name].length).setFontWeight('bold').setBackground('#222222').setFontColor('#ffffff');
    } else {
      const oldHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      CFG.sheets[name].forEach(h => { if (!oldHeaders.includes(h)) { sh.getRange(1, sh.getLastColumn() + 1).setValue(h).setFontWeight('bold').setBackground('#222222').setFontColor('#ffffff'); } });
    }
  });
  if (!settings.getProperty('ROOT_FOLDER_ID')) {
    const folder = DriveApp.createFolder('haloeditin — File Pesanan');
    settings.setProperty('ROOT_FOLDER_ID', folder.getId());
  }
  const email = Session.getActiveUser().getEmail();
  if (email && !findBy_('Users', 'Email', email)) {
    append_('Users', [email, email.split('@')[0], 'Owner', 'Ya', new Date()]);
  }
  return 'Setup selesai. Akun pelaksana setup terdaftar sebagai Owner.';
}

function getBootstrap(token) {
  const user = currentUser_(token);
  const data = {
    user,
    services: rows_('Services').filter(r => String(r.Aktif).toLowerCase() !== 'tidak').map(s => { try { s.PriceTiers = JSON.parse(s.HargaTiers || '[]'); } catch (_) { s.PriceTiers = []; } return s; }),
    orders: rows_('Orders').map(o => {
      o.payments = rows_('Payments').filter(p => p.OrderID === o.OrderID);
      o.items = rows_('OrderItems').filter(i => i.OrderID === o.OrderID);
      o.task = rows_('Tasks').find(t => t.OrderID === o.OrderID) || null;
      return o;
    }),
    clients: rows_('Clients'),
    tasks: rows_('Tasks'),
    payments: rows_('Payments'),
    users: user.Role === 'Editor' ? [] : rows_('Users').map(u => ({Email:u.Email,Nama:u.Nama,Role:u.Role,Aktif:u.Aktif})),
    statuses: { orders: CFG.orderStatuses, payments: CFG.paymentStatuses, tasks: CFG.taskStatuses }
  };
  if (user.Role === 'Editor') {
    data.orders = data.orders.filter(o => o.task && o.task.EditorEmail === user.Email);
    data.orders.forEach(o => { o.payments = []; });
    data.tasks = data.tasks.filter(t => t.EditorEmail === user.Email);
    data.payments = [];
    const visibleClients = new Set(data.orders.map(o => o.ClientID));
    data.clients = data.clients.filter(c => visibleClients.has(c.ClientID));
  }
  return data;
}

function saveOrder(input, token) {
  requireRole_(['Owner', 'Admin'], token);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const id = 'ORD-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0, 5).toUpperCase();
    const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
    const folder = root.createFolder(id + ' — ' + input.NamaKlien);
    ['01-File-Awal', '02-File-Kerja', '03-Hasil-Revisi', '04-Hasil-Final'].forEach(n => folder.createFolder(n));
    let client = input.Email ? findBy_('Clients', 'Email', input.Email) : null;
    if (!client) {
      const cid = 'CLI-' + Utilities.getUuid().slice(0, 7).toUpperCase();
      append_('Clients', [cid, input.NamaKlien, input.Kontak || '', input.Email || '', '', new Date()]);
      client = { ClientID: cid };
    }
    const items = input.items || [];
    const priced = items.map(i => {
      const svc = findBy_('Services', 'ServiceID', i.ServiceID);
      const tiers = svc ? JSON.parse(svc.HargaTiers || '[]') : [];
      const rawPrice = tiers[Number(i.TierIndex || 0)], price = typeof rawPrice === 'number' ? rawPrice : NaN;
      const qty = Number(i.Qty);
      if (!svc || !Number.isFinite(price) || price < 0 || !Number.isFinite(qty) || qty < 1) throw new Error('Layanan, jumlah, atau tarif tidak valid');
      return {i, svc, price, qty};
    });
    const total = priced.reduce((s, x) => s + Number(x.i.Qty) * x.price, 0);
    append_('Orders', [id, new Date(), client.ClientID, input.NamaKlien, input.Judul, input.JenisDokumen || '', input.Deadline || '', 'Baru', total, input.Catatan || '', folder.getId()]);
    priced.forEach(x => {
      const i=x.i, svc=x.svc, price=x.price;
      append_('OrderItems', ['ITM-' + Utilities.getUuid().slice(0, 8), id, i.ServiceID, svc.Nama, x.qty, price, x.qty * price]);
    });
    if (input.EditorEmail) append_('Tasks', ['TSK-' + Utilities.getUuid().slice(0, 8), id, input.EditorEmail, 'Belum Mulai', input.Deadline || '', '', new Date()]);
    log_('Pesanan dibuat', id, input.Judul, token);
    return { id, folderUrl: folder.getUrl() };
  } finally { lock.releaseLock(); }
}

function updateOrderStatus(orderId, status, token) {
  requireRole_(['Owner', 'Admin'], token);
  if (!CFG.orderStatuses.includes(status)) throw new Error('Status tidak valid');
  setCell_('Orders', 'OrderID', orderId, 'Status', status);
  log_('Status pesanan diubah', orderId, status, token);
  return true;
}

function updateTask(taskId, status, note, token) {
  const user = currentUser_(token);
  const task = findBy_('Tasks', 'TaskID', taskId);
  if (!task || (user.Role === 'Editor' && task.EditorEmail !== user.Email)) throw new Error('Akses tugas ditolak');
  if (!CFG.taskStatuses.includes(status)) throw new Error('Status tidak valid');
  setCell_('Tasks', 'TaskID', taskId, 'Status', status);
  setCell_('Tasks', 'TaskID', taskId, 'Catatan', note || '');
  setCell_('Tasks', 'TaskID', taskId, 'Diperbarui', new Date());
  const orderStatus = {
    'Dikerjakan': 'Dikerjakan',
    'Menunggu Pemeriksaan': 'Pemeriksaan',
    'Revisi': 'Revisi',
    'Selesai': 'Selesai'
  }[status];
  if (orderStatus) setCell_('Orders', 'OrderID', task.OrderID, 'Status', orderStatus);
  log_('Progres tugas diperbarui', task.OrderID, status, token);
  return true;
}

function addPayment(input, token) {
  requireRole_(['Owner', 'Admin'], token);
  const order = findBy_('Orders', 'OrderID', input.OrderID);
  if (!order) throw new Error('Pesanan tidak ditemukan');
  const amount = Number(input.Nominal);
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('Nominal pembayaran harus lebih dari nol');
  append_('Payments', ['PAY-' + Utilities.getUuid().slice(0, 8).toUpperCase(), input.OrderID, input.Tanggal || new Date(), amount, input.Metode || 'Transfer', input.Status || 'DP', input.Catatan || '']);
  const paid = rows_('Payments').filter(p => p.OrderID === input.OrderID && p.Status !== 'Refund').reduce((s, p) => s + Number(p.Nominal || 0), 0);
  setCell_('Orders', 'OrderID', input.OrderID, 'Status', paid >= Number(order.Total) ? 'Siap Dikerjakan' : 'Menunggu Pembayaran');
  log_('Pembayaran dicatat', input.OrderID, input.Nominal + ' — ' + input.Status, token);
  return true;
}

function saveService(input, token) {
  requireRole_(['Owner', 'Admin'], token);
  const id = input.ServiceID || 'SRV-' + Utilities.getUuid().slice(0, 7).toUpperCase();
  if (input.ServiceID) {
    ['Nama', 'Deskripsi', 'HargaDasar', 'Aktif'].forEach(k => setCell_('Services', 'ServiceID', id, k, input[k]));
  } else append_('Services', [id, input.Nama, input.Deskripsi || '', Number(input.HargaDasar || 0), 'Ya', input.Kategori || 'Lainnya', input.Satuan || 'Paket', JSON.stringify(input.PriceTiers || [Number(input.HargaDasar || 0)]), input.Catatan || '']);
  return id;
}

function saveUser(input, token) {
  const actor = requireRole_(['Owner'], token);
  const role = ['Owner', 'Admin', 'Editor'].includes(input.Role) ? input.Role : 'Editor';
  const email = String(input.Email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email tidak valid');
  const old = findBy_('Users', 'Email', email);
  if (!old && !input.Password) throw new Error('Kata sandi awal wajib diisi');
  if (email === actor.Email && (role !== 'Owner' || input.Aktif === 'Tidak')) throw new Error('Owner tidak dapat menonaktifkan atau menurunkan role akunnya sendiri');
  if (old) {
    setCell_('Users', 'Email', email, 'Nama', input.Nama);
    setCell_('Users', 'Email', email, 'Role', role);
    setCell_('Users', 'Email', email, 'Aktif', input.Aktif || 'Ya');
    if (input.Password) setPassword_(email, input.Password);
  } else {
    append_('Users', [email, input.Nama, role, 'Ya', new Date(), '', '']);
    if (input.Password) setPassword_(email, input.Password);
  }
  return true;
}

function uploadFile(orderId, name, mimeType, data, token) {
  const user = currentUser_(token);
  const order = findBy_('Orders', 'OrderID', orderId);
  if (!order) throw new Error('Pesanan tidak ditemukan');
  if (!name || !data) throw new Error('File tidak lengkap');
  const bytes = Utilities.base64Decode(data);
  if (bytes.length > 8 * 1024 * 1024) throw new Error('Ukuran file maksimal 8 MB');
  if (user.Role === 'Editor') {
    const task = rows_('Tasks').find(t => t.OrderID === orderId && t.EditorEmail === user.Email);
    if (!task) throw new Error('Akses file ditolak');
  } else requireRole_(['Owner', 'Admin']);
  const parent = DriveApp.getFolderById(order.FolderID);
  const folderName = user.Role === 'Editor' ? '02-File-Kerja' : '01-File-Awal';
  const it = parent.getFoldersByName(folderName);
  const folder = it.hasNext() ? it.next() : parent;
  const file = folder.createFile(Utilities.newBlob(bytes, mimeType, name));
  log_('File diunggah', orderId, name, token);
  return { name: file.getName(), url: file.getUrl() };
}

function getOrderFiles(orderId, token) {
  const user = currentUser_(token), order = findBy_('Orders', 'OrderID', orderId);
  if (!order) throw new Error('Pesanan tidak ditemukan');
  if (user.Role === 'Editor' && !rows_('Tasks').some(t => t.OrderID === orderId && t.EditorEmail === user.Email)) throw new Error('Akses file ditolak');
  const folder = DriveApp.getFolderById(order.FolderID), files = [];
  const walk = f => { const fi = f.getFiles(); while (fi.hasNext()) { const x = fi.next(); files.push({ name: x.getName(), url: x.getUrl() }); } const dirs = f.getFolders(); while (dirs.hasNext()) walk(dirs.next()); };
  walk(folder); return files;
}

function login(email, password) {
  email = String(email || '').trim().toLowerCase();
  const key = 'login-' + Utilities.base64EncodeWebSafe(email);
  const cache = CacheService.getScriptCache();
  const attempts = Number(cache.get(key) || 0);
  if (attempts >= 8) throw new Error('Terlalu banyak percobaan. Coba lagi 15 menit lagi.');
  const user = findBy_('Users', 'Email', email);
  if (!user || String(user.Aktif).toLowerCase() === 'tidak' || !user.PasswordHash || !password || passwordHash_(password, user.PasswordSalt) !== user.PasswordHash) {
    cache.put(key, String(attempts + 1), 900); throw new Error('Email atau kata sandi salah, atau akun belum memiliki akses.');
  }
  cache.remove(key); return createSession_(user);
}
function loginWithGoogle() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) throw new Error('Google tidak memberikan email akun. Gunakan email dan kata sandi yang dibuat Owner.');
  const user = findBy_('Users', 'Email', email);
  if (!user || String(user.Aktif).toLowerCase() === 'tidak') throw new Error('Akun Google ini belum diaktifkan oleh Owner.');
  return createSession_(user);
}
function createSession_(u) { const token=Utilities.getUuid()+Utilities.getUuid(); CacheService.getScriptCache().put('session-'+token, u.Email, 21600); return {token, user:{Email:u.Email,Nama:u.Nama,Role:u.Role}}; }
function logout(token) { if (token) CacheService.getScriptCache().remove('session-'+token); return true; }
function setPassword_(email, password) {
  if (String(password).length < 10) throw new Error('Kata sandi minimal 10 karakter');
  const salt=Utilities.getUuid()+Utilities.getUuid(); setCell_('Users','Email',email,'PasswordSalt',salt); setCell_('Users','Email',email,'PasswordHash',passwordHash_(password,salt));
}
function passwordHash_(password,salt) { let h=String(password)+'|'+String(salt); for(let i=0;i<12000;i++) h=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,h)); return h; }
function currentUser_(token) {
  const email = token && CacheService.getScriptCache().get('session-'+token);
  if (!email) throw new Error('Sesi masuk tidak berlaku. Silakan masuk kembali.');
  const user = findBy_('Users', 'Email', email);
  if (!user || String(user.Aktif).toLowerCase() === 'tidak') throw new Error('Akun belum diaktifkan. Hubungi owner aplikasi.');
  return { Email: email, Nama: user.Nama, Role: user.Role };
}
function requireRole_(roles, token) { const u = currentUser_(token); if (!roles.includes(u.Role)) throw new Error('Akses tidak diizinkan'); return u; }
function rows_(name) {
  const sh = database_().getSheetByName(name); if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  return values.slice(1).filter(r => r.some(v => v !== '')).map(r => Object.fromEntries(values[0].map((k, i) => [k, r[i] instanceof Date ? r[i].toISOString() : r[i]])));
}
function append_(name, values) { database_().getSheetByName(name).appendRow(values); }
function findBy_(name, key, value) { return rows_(name).find(r => String(r[key]).toLowerCase() === String(value).toLowerCase()) || null; }
function setCell_(name, key, value, col, newValue) {
  const sh = database_().getSheetByName(name), grid = sh.getDataRange().getValues();
  const ki = grid[0].indexOf(key), ci = grid[0].indexOf(col);
  for (let i = 1; i < grid.length; i++) if (String(grid[i][ki]) === String(value)) { sh.getRange(i + 1, ci + 1).setValue(newValue); return; }
  throw new Error('Data tidak ditemukan: ' + value);
}
function log_(action, orderId, detail, token) { const u = currentUser_(token); append_('ActivityLog', [new Date(), u.Email, action, orderId, detail]); }
function database_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActive();
  if (!active) throw new Error('Database belum disiapkan. Jalankan setup() terlebih dahulu.');
  return active;
}
function importServiceCatalog() {
  const email=String(Session.getActiveUser().getEmail()||'').trim().toLowerCase(), user=findBy_('Users','Email',email);
  if (!user || user.Role!=='Owner' || String(user.Aktif).toLowerCase()==='tidak') throw new Error('Fungsi impor hanya dapat dijalankan Owner.');
  const catalog = [{"category": "Jasa Pengeditan Skripsi", "name": "Penomoran Halaman Standar", "unit": "Halaman", "prices": [null, null, 100, null, null, 120, 135, 150, 165], "note": "-", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Penomoran Halaman Sedang", "unit": "Halaman", "prices": [null, null, 150, null, null, 180, 202.5, 225, 247.5], "note": "-", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Penomoran Halaman Rumit", "unit": "Halaman", "prices": [null, null, 250, null, null, 300, 337.5, 375, 412.5], "note": "-", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Daftar Isi, Tabel, Gambar & Lampiran Otomatis Proposal <50 halaman", "unit": "Dokumen", "prices": [null, null, 15000, null, null, 18000, 20250, 22500, 24750], "note": "Include Nomor Halaman", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Daftar Isi, Tabel, Gambar & Lampiran Otomatis Proposal 50 - 75 Halaman", "unit": "Dokumen", "prices": [null, null, 20000, null, null, 24000, 27000, 30000, 33000], "note": "Include Nomor Halaman", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Daftar Isi, Tabel, Gambar & Lampiran Otomatis Proposal 75 - 100 halaman", "unit": "Dokumen", "prices": [null, null, 25000, null, null, 30000, 33750, 37500, 41250], "note": "Include Nomor Halaman", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Daftar Isi, Tabel, Gambar & Lampiran Otomatis Proposal >100 halaman", "unit": "Dokumen", "prices": [null, null, 35000, null, null, 42000, 47250, 52500, 57750], "note": "Include Nomor Halaman", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Daftar Isi, Tabel, Gambar & Lampiran Otomatis Skripsi <100 halaman", "unit": "Dokumen", "prices": [null, null, 30000, null, null, 36000, 40500, 45000, 49500], "note": "Include Nomor Halaman", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Daftar Isi, Tabel, Gambar & Lampiran Otomatis Skripsi 100-150 halaman", "unit": "Dokumen", "prices": [null, null, 40000, null, null, 48000, 54000, 60000, 66000], "note": "Include Nomor Halaman", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Daftar Isi, Tabel, Gambar & Lampiran Otomatis Skripsi 150-200 halaman", "unit": "Dokumen", "prices": [null, null, 50000, null, null, 60000, 67500, 75000, 82500], "note": "Include Nomor Halaman", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Daftar Isi, Tabel, Gambar & Lampiran Otomatis Skripsi >200 halaman", "unit": "Dokumen", "prices": [null, null, 60000, null, null, 72000, 81000, 90000, 99000], "note": "Include Nomor Halaman", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit Kerapihan Level 1 (Format Penulisan)", "unit": "Halaman", "prices": [null, null, 250, null, null, 300, 337.5, 375, 412.5], "note": "-", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit Kerapihan Level 2 (+Kerapihan Tabel dan Gambar)", "unit": "Halaman", "prices": [null, null, 500, null, null, 600, 675, 750, 825], "note": "-", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit Kerapihan Level 3 (+Kerapihan Bagan, Struktur dan Grafik Standar)", "unit": "Halaman", "prices": [null, null, 2000, null, null, 2400, 2700, 3000, 3300], "note": "-", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit Kerapihan Level 4 (+Kerapihan Bagan, Struktur dan Grafik Rumit)", "unit": "Halaman", "prices": [null, null, 3000, null, null, 3600, 4050, 4500, 4950], "note": "-", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Bagan, Struktur, Grafik Mudah", "unit": "Bagian", "prices": [null, null, 7500, null, null, 9000, 10125, 11250, 12375], "note": "-", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Bagan, Struktur, Grafik Sedang", "unit": "Bagian", "prices": [null, null, 10000, null, null, 12000, 13500, 15000, 16500], "note": "-", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Pembuatan Bagan, Struktur, Grafik Rumit", "unit": "Bagian", "prices": [null, null, 15000, null, null, 18000, 20250, 22500, 24750], "note": "-", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit bagan, Struktur, Grafik Mudah", "unit": "Bagian", "prices": [null, null, 5000, null, null, 6000, 6750, 7500, 8250], "note": "", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit bagan, Struktur, Grafik Sedang", "unit": "Bagian", "prices": [null, null, 7500, null, null, 9000, 10125, 11250, 12375], "note": "", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit bagan, Struktur, Grafik Rumit", "unit": "Bagian", "prices": [null, null, 10000, null, null, 12000, 13500, 15000, 16500], "note": "", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit Daftar Isi Proposal", "unit": "Dokumen", "prices": [null, null, 10000, null, null, 12000, 13500, 15000, 16500], "note": "Hanya untuk yang Repeat Order Daftar Isi", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Edit Daftar Isi Skripsi", "unit": "Dokumen", "prices": [null, null, 15000, null, null, 18000, 20250, 22500, 24750], "note": "Hanya untuk yang Repeat Order Daftar Isi", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Perbaikan Typo Half Typing Otomatis (Kesalahan Pengetikan)", "unit": "Halaman", "prices": [null, null, 150, null, null, 180, 202.5, 225, 247.5], "note": "Pengecekan Otomatis", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Perbaikan Typo Half Typing Semi-Otomatis (Kesalahan Pengetikan)", "unit": "Halaman", "prices": [null, null, 500, null, null, 600, 675, 750, 825], "note": "Pengecekan otomatis perbaikan manual", "remark": "Untuk file yang sudah rapih"}, {"category": "Jasa Pengeditan Skripsi", "name": "Perbaikan Typo Full Typing (Kesalahan Penulisan, Cetak Miring bahasa asing)", "unit": "Halaman", "prices": [null, null, 750, null, null, 900, 1012.5, 1125, 1237.5], "note": "-", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Perbaikan Typo Full Typing (Kesalahan Penulisan, Cetak Miring bahasa asing, Tanda baca dll", "unit": "Halaman", "prices": [null, null, 1000, null, null, 1200, 1350, 1500, 1650], "note": "-", "remark": ""}, {"category": "Jasa Pengeditan Skripsi", "name": "Konversi File (Word, PDF, JPG, PNG, WEBP, Aviv dll)", "unit": "Dokumen", "prices": [null, null, 5000, null, null, 6000, 6750, 7500, 8250], "note": "-", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Pemisahan Dokumen Skripsi", "unit": "Dokumen", "prices": [null, null, 10000, null, null, 12000, 13500, 15000, 16500], "note": "Untuk 1 File Doc/PDF jika keduanya maka harga dikali 2", "remark": "-"}, {"category": "Jasa Pengeditan Skripsi", "name": "Watermark Dokumen", "unit": "Dokumen", "prices": [null, null, 250, null, null, 300, 337.5, 375, 412.5], "note": "", "remark": ""}, {"category": "Jasa Pembuatan Citasi", "name": "Pembuatan Bodynote Manual/Mendeley", "unit": "Bodynote", "prices": [null, null, 250, null, null, 300, 337.5, 375, 412.5], "note": "-", "remark": "-"}, {"category": "Jasa Pembuatan Citasi", "name": "Pembuatan Footnote Manual", "unit": "Footnote", "prices": [null, null, 500, null, null, 600, 675, 750, 825], "note": "", "remark": ""}, {"category": "Jasa Pembuatan Citasi", "name": "Pembuatan Footnote Mendeley", "unit": "Footnote", "prices": [null, null, 250, null, null, 300, 337.5, 375, 412.5], "note": "", "remark": ""}, {"category": "Jasa Pembuatan Citasi", "name": "Pembuatan Daftar Pustaka Manual Half Citation", "unit": "Sumber", "prices": [null, null, 1000, null, null, 1200, 1350, 1500, 1650], "note": "Informasi Sumber Pustaka dari Klien", "remark": ""}, {"category": "Jasa Pembuatan Citasi", "name": "Pembuatan Daftar Pustaka Manual Full Citation", "unit": "Sumber", "prices": [null, null, 2500, null, null, 3000, 3375, 3750, 4125], "note": "Informasi Sumber Pustaka dicari Oleh Tim Editin", "remark": "Jaminan Kelengkapan Data Pustaka"}, {"category": "Jasa Pembuatan Citasi", "name": "Pembuatan Daftar Pustaka Mendeley Half Citation Original", "unit": "Sumber", "prices": [null, null, 1000, null, null, 1200, 1350, 1500, 1650], "note": "Informasi Sumber Pustaka dari Klien", "remark": ""}, {"category": "Jasa Pembuatan Citasi", "name": "Pembuatan Daftar Pustaka Mendeley Half Citation + Kelengkapan", "unit": "Sumber", "prices": [null, null, 1500, null, null, 1800, 2025, 2250, 2475], "note": "Informasi Sumber Pustaka dari Klien hanya terdiri dari minimal 3 komponen pustaka", "remark": "Jaminan Kelengkapan Data Pustaka"}, {"category": "Jasa Pembuatan Citasi", "name": "Pembuatan Daftar Pustaka Mendeley Full Citation", "unit": "Sumber", "prices": [null, null, 2500, null, null, 3000, 3375, 3750, 4125], "note": "Informasi Sumber Pustaka dicari Oleh Tim Editin", "remark": "Jaminan Kelengkapan Data Pustaka"}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan A4 Spasi 1", "unit": "Halaman", "prices": [3500, null, null, 4025, 4550, 4900, 5250, 5600, 5950], "note": "-", "remark": "-"}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan A4 Spasi 1,5", "unit": "Halaman", "prices": [3000, null, null, 3450, 3900, 4200, 4500, 4800, 5100], "note": "-", "remark": "-"}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan A4 Spasi 2", "unit": "Halaman", "prices": [2500, null, null, 2875, 3250, 3500, 3750, 4000, 4250], "note": "-", "remark": "-"}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan F4 Spasi 1", "unit": "Halaman", "prices": [4000, null, null, 4600, 5200, 5600, 6000, 6400, 6800], "note": "-", "remark": "-"}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan F4 Spasi 1,5", "unit": "Halaman", "prices": [3500, null, null, 4025, 4550, 4900, 5250, 5600, 5950], "note": "-", "remark": "-"}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan F4 Spasi 2", "unit": "Halaman", "prices": [3000, null, null, 3450, 3900, 4200, 4500, 4800, 5100], "note": "-", "remark": "-"}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan Abstrak Indonesia & Inggris otomatis", "unit": "Dokumen", "prices": [25000, null, null, 28750, 32500, 35000, 37500, 40000, 42500], "note": "-", "remark": "-"}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan Abstrak Indonesia & Inggris Manual", "unit": "Dokumen", "prices": [45000, null, null, 51750, 58500, 63000, 67500, 72000, 76500], "note": "", "remark": ""}, {"category": "Jasa Pengetikan Umum", "name": "Pengetikan Abstrak Inggris Manual", "unit": "Dokumen", "prices": [25000, null, null, 28750, 32500, 35000, 37500, 40000, 42500], "note": "", "remark": ""}, {"category": "Jasa Pengetikan Umum", "name": "Input data Excel", "unit": "Cell", "prices": [20, null, null, 23, 26, 28, 30, 32, 34], "note": "-", "remark": "-"}, {"category": "Konversi Skripsi ke Artikel", "name": "Konversi Skripsi Ke Artikel Jurnal", "unit": "Dokumen", "prices": [75000, null, null, 90000, 112500, null, null, null, null], "note": "Skripsi <100 Halaman", "remark": ""}, {"category": "Konversi Skripsi ke Artikel", "name": "Konversi Skripsi Ke Artikel Jurnal", "unit": "Dokumen", "prices": [90000, null, null, 108000, 135000, null, null, null, null], "note": "Skripsi 100-125 Halaman", "remark": ""}, {"category": "Konversi Skripsi ke Artikel", "name": "Konversi Skripsi Ke Artikel Jurnal", "unit": "Dokumen", "prices": [105000, null, null, 126000, 157500, null, null, null, null], "note": "Skripsi 125-150 Halaman", "remark": ""}, {"category": "Konversi Skripsi ke Artikel", "name": "Konversi Skripsi Ke Artikel Jurnal", "unit": "Dokumen", "prices": [120000, null, null, 144000, 180000, null, null, null, null], "note": "Skripsi 150-175 Halaman", "remark": ""}, {"category": "Konversi Skripsi ke Artikel", "name": "Konversi Skripsi Ke Artikel Jurnal", "unit": "Dokumen", "prices": [135000, null, null, 162000, 202500, null, null, null, null], "note": "Skripsi 175-200 Halaman", "remark": ""}, {"category": "Konversi Skripsi ke Artikel", "name": "Konversi Skripsi Ke Artikel Jurnal", "unit": "Dokumen", "prices": [150000, null, null, 180000, 225000, null, null, null, null], "note": "Skripsi >200 Halaman", "remark": ""}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Proposal No Cover", "unit": "File", "prices": [60000, null, null, 75000, 90000, null, null, null, null], "note": "Materi Presentasi dari Klien", "remark": "Pakai Transisi + Rp 10.000"}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Skripsi No Cover", "unit": "File", "prices": [75000, null, null, 93750, 112500, null, null, null, null], "note": "Materi Presentasi dari Klien", "remark": ""}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Proposal Full Cover <75 Halaman", "unit": "File", "prices": [85000, null, null, 106250, 127500, null, null, null, null], "note": "Klien Hanya memberikan draft Proposal", "remark": "Pakai transisi + Rp 10.000"}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Proposal Full Cover >75 Halaman", "unit": "File", "prices": [100000, null, null, 125000, 150000, null, null, null, null], "note": "Klien Hanya memberikan draft Proposal", "remark": ""}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Skripsi Full Cover <100 Halaman", "unit": "File", "prices": [115000, null, null, 143750, 172500, null, null, null, null], "note": "Klien Hanya memberikan draft Proposal", "remark": ""}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Skripsi Full Cover 100-125 Halaman", "unit": "File", "prices": [135000, null, null, 168750, 202500, null, null, null, null], "note": "Klien Hanya memberikan draft Skripsi", "remark": "Pakai transisi+ Rp 10.000"}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Skripsi Full Cover >125-150 Halaman", "unit": "File", "prices": [150000, null, null, 187500, 225000, null, null, null, null], "note": "Klien Hanya memberikan draft Skripsi", "remark": "Pakai transisi+ Rp 10.000"}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Skripsi Full Cover >150-175 Halaman", "unit": "File", "prices": [165000, null, null, 206250, 247500, null, null, null, null], "note": "Klien Hanya memberikan draft Skripsi", "remark": ""}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Skripsi Full Cover >175-200 Halaman", "unit": "File", "prices": [180000, null, null, 225000, 270000, null, null, null, null], "note": "Klien Hanya memberikan draft Skripsi", "remark": ""}, {"category": "Jasa Pembuatan PPT", "name": "Pembuatan PPT Skripsi Full Cover >200 Halaman", "unit": "File", "prices": [200000, null, null, 250000, 300000, null, null, null, null], "note": "Klien Hanya memberikan draft Skripsi", "remark": ""}, {"category": "Jasa Pengetikan Transkrip Wawancara", "name": "Verbatim", "unit": "Menit", "prices": [4000, null, null, 4600, 5200, 5600, null, null, null], "note": "Persis dengan rekaman", "remark": "-"}, {"category": "Jasa Pengetikan Transkrip Wawancara", "name": "Substansial", "unit": "Menit", "prices": [3000, null, null, 3450, 3900, 4200, null, null, null], "note": "Ditulis intinya saja", "remark": "-"}, {"category": "Jasa Pembuatan CV & Berkas Lamaran", "name": "CV ATS", "unit": "File", "prices": [null, 20000, null, null, 26000, 28000, 30000, 32000, null], "note": "File berupa word yang dapat diedit", "remark": "-"}, {"category": "Jasa Pembuatan CV & Berkas Lamaran", "name": "CV Kreatif", "unit": "File", "prices": [null, 30000, null, null, 39000, 42000, 45000, 48000, null], "note": "File berupa gambar/pdf", "remark": "-"}, {"category": "Jasa Pembuatan CV & Berkas Lamaran", "name": "CV + Merge Berkas Lamaran", "unit": "File", "prices": [null, 35000, null, null, 45500, 49000, 52500, 56000, null], "note": "File berupa word dan pdf", "remark": "-"}, {"category": "Desain Grafis", "name": "Desain Poster", "unit": "Poster", "prices": [null, 25000, null, null, 32500, 35000, 37500, 40000, null], "note": "1 Poster", "remark": ""}, {"category": "Desain Grafis", "name": "Desain Buku Panduan", "unit": "Halaman", "prices": [null, 10000, null, null, 13000, 14000, 15000, 16000, null], "note": "-", "remark": ""}, {"category": "Desain Grafis", "name": "Watermark", "unit": "Halaman", "prices": [null, 250, null, null, 325, 350, 375, 400, null], "note": "-", "remark": ""}, {"category": "Jasa Translate Dokumen", "name": "Translate Dokumen Inggris > Indonesia", "unit": "Kata", "prices": [150, null, null, null, 195, 210, 225, 240, null], "note": "1 Kata", "remark": ""}, {"category": "Jasa Translate Dokumen", "name": "Translate Dokumen Indonesia > Inggris", "unit": "Kata", "prices": [200, null, null, null, 260, 280, 300, 320, null], "note": "1 kata", "remark": ""}];
  const sh = database_().getSheetByName('Services');
  const existing = new Set(rows_('Services').map(s => s.ServiceID));
  catalog.forEach((s,i) => { const id='EDITIN-'+String(i+1).padStart(3,'0'); if (!existing.has(id)) { const first=s.prices.find(x=>x!==null)||0; append_('Services',[id,s.name,s.note||s.remark,first,'Ya',s.category,s.unit,JSON.stringify(s.prices),[s.note,s.remark].filter(Boolean).filter(x=>x!=='-').join(' · ')]); } });
  return catalog.length+' layanan dari daftar 1 Juni 2024 tersedia.';
}


