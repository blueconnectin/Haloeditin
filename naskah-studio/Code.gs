/**
 * Naskah Studio — Apps Script web app backend
 * Bind this project to a Google Sheet, run setup(), then deploy as a web app.
 */
const CFG = {
  sheets: {
    Users: ['Email', 'Nama', 'Role', 'Aktif', 'Dibuat'],
    Clients: ['ClientID', 'Nama', 'Kontak', 'Email', 'Catatan', 'Dibuat'],
    Services: ['ServiceID', 'Nama', 'Deskripsi', 'HargaDasar', 'Aktif'],
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
    .setTitle('Naskah Studio — Manajemen Jasa Dokumen')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setup() {
  const settings = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActive();
  if (!ss) ss = SpreadsheetApp.create('Naskah Studio — Database');
  settings.setProperty('DB_SPREADSHEET_ID', ss.getId());
  Object.keys(CFG.sheets).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(CFG.sheets[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, CFG.sheets[name].length).setFontWeight('bold').setBackground('#222222').setFontColor('#ffffff');
    }
  });
  if (!settings.getProperty('ROOT_FOLDER_ID')) {
    const folder = DriveApp.createFolder('Naskah Studio — File Pesanan');
    settings.setProperty('ROOT_FOLDER_ID', folder.getId());
  }
  const email = Session.getActiveUser().getEmail();
  if (email && !findBy_('Users', 'Email', email)) {
    append_('Users', [email, email.split('@')[0], 'Owner', 'Ya', new Date()]);
  }
  return 'Setup selesai. Akun pelaksana setup terdaftar sebagai Owner.';
}

function getBootstrap() {
  const user = currentUser_();
  const data = {
    user,
    services: rows_('Services').filter(r => String(r.Aktif).toLowerCase() !== 'tidak'),
    orders: rows_('Orders').map(o => {
      o.payments = rows_('Payments').filter(p => p.OrderID === o.OrderID);
      o.items = rows_('OrderItems').filter(i => i.OrderID === o.OrderID);
      o.task = rows_('Tasks').find(t => t.OrderID === o.OrderID) || null;
      return o;
    }),
    clients: rows_('Clients'),
    tasks: rows_('Tasks'),
    users: user.Role === 'Editor' ? [] : rows_('Users').filter(u => String(u.Aktif).toLowerCase() !== 'tidak'),
    statuses: { orders: CFG.orderStatuses, payments: CFG.paymentStatuses, tasks: CFG.taskStatuses }
  };
  if (user.Role === 'Editor') {
    data.orders = data.orders.filter(o => o.task && o.task.EditorEmail === user.Email);
    data.tasks = data.tasks.filter(t => t.EditorEmail === user.Email);
  }
  return data;
}

function saveOrder(input) {
  requireRole_(['Owner', 'Admin']);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const id = 'ORD-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0, 5).toUpperCase();
    const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
    const folder = root.createFolder(id + ' — ' + input.NamaKlien);
    ['01-File-Awal', '02-File-Kerja', '03-Hasil-Revisi', '04-Hasil-Final'].forEach(n => folder.createFolder(n));
    let client = findBy_('Clients', 'Email', input.Email || '');
    if (!client) {
      const cid = 'CLI-' + Utilities.getUuid().slice(0, 7).toUpperCase();
      append_('Clients', [cid, input.NamaKlien, input.Kontak || '', input.Email || '', '', new Date()]);
      client = { ClientID: cid };
    }
    const items = input.items || [];
    const total = items.reduce((s, i) => s + Number(i.Qty || 1) * Number(i.Harga || 0), 0);
    append_('Orders', [id, new Date(), client.ClientID, input.NamaKlien, input.Judul, input.JenisDokumen || '', input.Deadline || '', 'Baru', total, input.Catatan || '', folder.getId()]);
    items.forEach(i => append_('OrderItems', ['ITM-' + Utilities.getUuid().slice(0, 8), id, i.ServiceID || '', i.NamaLayanan, Number(i.Qty || 1), Number(i.Harga || 0), Number(i.Qty || 1) * Number(i.Harga || 0)]));
    if (input.EditorEmail) append_('Tasks', ['TSK-' + Utilities.getUuid().slice(0, 8), id, input.EditorEmail, 'Belum Mulai', input.Deadline || '', '', new Date()]);
    log_('Pesanan dibuat', id, input.Judul);
    return { id, folderUrl: folder.getUrl() };
  } finally { lock.releaseLock(); }
}

function updateOrderStatus(orderId, status) {
  requireRole_(['Owner', 'Admin']);
  if (!CFG.orderStatuses.includes(status)) throw new Error('Status tidak valid');
  setCell_('Orders', 'OrderID', orderId, 'Status', status);
  log_('Status pesanan diubah', orderId, status);
  return true;
}

function updateTask(taskId, status, note) {
  const user = currentUser_();
  const task = findBy_('Tasks', 'TaskID', taskId);
  if (!task || (user.Role === 'Editor' && task.EditorEmail !== user.Email)) throw new Error('Akses tugas ditolak');
  if (!CFG.taskStatuses.includes(status)) throw new Error('Status tidak valid');
  setCell_('Tasks', 'TaskID', taskId, 'Status', status);
  setCell_('Tasks', 'TaskID', taskId, 'Catatan', note || '');
  setCell_('Tasks', 'TaskID', taskId, 'Diperbarui', new Date());
  log_('Progres tugas diperbarui', task.OrderID, status);
  return true;
}

function addPayment(input) {
  requireRole_(['Owner', 'Admin']);
  append_('Payments', ['PAY-' + Utilities.getUuid().slice(0, 8).toUpperCase(), input.OrderID, input.Tanggal || new Date(), Number(input.Nominal), input.Metode || 'Transfer', input.Status || 'DP', input.Catatan || '']);
  const paid = rows_('Payments').filter(p => p.OrderID === input.OrderID && p.Status !== 'Refund').reduce((s, p) => s + Number(p.Nominal || 0), 0);
  const order = findBy_('Orders', 'OrderID', input.OrderID);
  setCell_('Orders', 'OrderID', input.OrderID, 'Status', paid >= Number(order.Total) ? 'Siap Dikerjakan' : 'Menunggu Pembayaran');
  log_('Pembayaran dicatat', input.OrderID, input.Nominal + ' — ' + input.Status);
  return true;
}

function saveService(input) {
  requireRole_(['Owner', 'Admin']);
  const id = input.ServiceID || 'SRV-' + Utilities.getUuid().slice(0, 7).toUpperCase();
  if (input.ServiceID) {
    ['Nama', 'Deskripsi', 'HargaDasar', 'Aktif'].forEach(k => setCell_('Services', 'ServiceID', id, k, input[k]));
  } else append_('Services', [id, input.Nama, input.Deskripsi || '', Number(input.HargaDasar || 0), 'Ya']);
  return id;
}

function saveUser(input) {
  requireRole_(['Owner']);
  const role = ['Owner', 'Admin', 'Editor'].includes(input.Role) ? input.Role : 'Editor';
  const old = findBy_('Users', 'Email', input.Email);
  if (old) {
    setCell_('Users', 'Email', input.Email, 'Nama', input.Nama);
    setCell_('Users', 'Email', input.Email, 'Role', role);
    setCell_('Users', 'Email', input.Email, 'Aktif', input.Aktif || 'Ya');
  } else append_('Users', [input.Email, input.Nama, role, 'Ya', new Date()]);
  return true;
}

function uploadFile(orderId, name, mimeType, data) {
  const user = currentUser_();
  const order = findBy_('Orders', 'OrderID', orderId);
  if (!order) throw new Error('Pesanan tidak ditemukan');
  if (user.Role === 'Editor') {
    const task = rows_('Tasks').find(t => t.OrderID === orderId && t.EditorEmail === user.Email);
    if (!task) throw new Error('Akses file ditolak');
  } else requireRole_(['Owner', 'Admin']);
  const parent = DriveApp.getFolderById(order.FolderID);
  const folderName = user.Role === 'Editor' ? '02-File-Kerja' : '01-File-Awal';
  const it = parent.getFoldersByName(folderName);
  const folder = it.hasNext() ? it.next() : parent;
  const bytes = Utilities.base64Decode(data);
  const file = folder.createFile(Utilities.newBlob(bytes, mimeType, name));
  log_('File diunggah', orderId, name);
  return { name: file.getName(), url: file.getUrl() };
}

function getOrderFiles(orderId) {
  const user = currentUser_(), order = findBy_('Orders', 'OrderID', orderId);
  if (!order) throw new Error('Pesanan tidak ditemukan');
  if (user.Role === 'Editor' && !rows_('Tasks').some(t => t.OrderID === orderId && t.EditorEmail === user.Email)) throw new Error('Akses file ditolak');
  const folder = DriveApp.getFolderById(order.FolderID), files = [];
  const walk = f => { const fi = f.getFiles(); while (fi.hasNext()) { const x = fi.next(); files.push({ name: x.getName(), url: x.getUrl() }); } const dirs = f.getFolders(); while (dirs.hasNext()) walk(dirs.next()); };
  walk(folder); return files;
}

function currentUser_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Email akun Google tidak tersedia. Deploy web app untuk pengguna dalam domain Google Workspace.');
  const user = findBy_('Users', 'Email', email);
  if (!user || String(user.Aktif).toLowerCase() === 'tidak') throw new Error('Akun belum diaktifkan. Hubungi owner aplikasi.');
  return { Email: email, Nama: user.Nama, Role: user.Role };
}
function requireRole_(roles) { const u = currentUser_(); if (!roles.includes(u.Role)) throw new Error('Akses tidak diizinkan'); return u; }
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
function log_(action, orderId, detail) { const u = currentUser_(); append_('ActivityLog', [new Date(), u.Email, action, orderId, detail]); }
function database_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActive();
  if (!active) throw new Error('Database belum disiapkan. Jalankan setup() terlebih dahulu.');
  return active;
}

