/**
 * PORTAL RT DIGITAL — Google Apps Script Backend
 * Database: Google Sheets created automatically by setup()
 *
 * IMPORTANT:
 * 1) Run setup() once from Apps Script editor.
 * 2) Change the default admin password immediately after first login.
 * 3) Deploy as Web app: Execute as "Me"; Who has access: "Anyone".
 *
 * The same project can serve the full UI through doGet().
 * For GitHub/Blogger/other hosting, use the deployed Web App URL as an
 * iframe/embed or set APP_CONFIG.API_URL in Index.html for API integration.
 */

const APP = {
  NAME: 'Portal RT Digital',
  VERSION: '1.0.0',
  TZ: 'Asia/Jakarta',
  SESSION_HOURS: 12,
  DEFAULT_ADMIN_NIK: 'admin',
  DEFAULT_ADMIN_PASSWORD: 'Admin@12345',
  DEFAULT_RT: 'RT 05',
  DEFAULT_RW: '05',
  DEFAULT_VILLAGE: 'Desa Pegandekan',
  DEFAULT_DISTRICT: 'Kecamatan Contoh',
  DEFAULT_REGENCY: 'Kabupaten Contoh',
  DEFAULT_PROVINCE: 'Jawa Tengah',
  SHEETS: {
    USERS: 'Users',
    ANNOUNCEMENTS: 'Announcements',
    ACTIVITIES: 'Activities',
    RONDA: 'Ronda',
    ASSETS: 'Assets',
    BORROWINGS: 'Borrowings',
    FEEDBACK: 'Feedback',
    SOS: 'SOS',
    DOCUMENTS: 'Documents',
    DUES: 'Dues',
    FINANCE: 'Finance',
    PLANS: 'Plans',
    NOTIFICATIONS: 'Notifications',
    SETTINGS: 'Settings',
    AUDIT: 'AuditLog'
  }
};

const HEADERS = {
  Users: ['id','nik','name','phone','email','address','passwordHash','salt','role','status','avatar','createdAt','approvedAt','approvedBy','lastLogin'],
  Announcements: ['id','title','content','category','image','publishAt','status','createdBy','createdAt'],
  Activities: ['id','title','description','date','time','location','category','status','createdBy','createdAt'],
  Ronda: ['id','date','shift','groupName','members','location','status','notes','createdAt'],
  Assets: ['id','name','category','description','quantity','available','condition','location','rules','status','createdAt'],
  Borrowings: ['id','userId','assetId','assetName','startDate','endDate','purpose','quantity','documentUrl','status','reviewNote','createdAt','approvedAt','approvedBy'],
  Feedback: ['id','userId','name','category','title','message','imageUrl','status','reply','createdAt','updatedAt'],
  SOS: ['id','userId','name','phone','location','message','status','handledBy','handledAt','createdAt'],
  Documents: ['id','userId','name','type','url','note','status','createdAt'],
  Dues: ['id','userId','name','month','year','amount','status','paidAt','method','note','createdAt'],
  Finance: ['id','type','category','description','amount','date','method','reference','createdBy','createdAt'],
  Plans: ['id','type','title','description','date','time','location','budget','status','createdBy','createdAt'],
  Notifications: ['id','userId','title','message','type','read','link','createdAt'],
  Settings: ['key','value','updatedAt','updatedBy'],
  AuditLog: ['id','userId','actor','action','module','detail','createdAt']
};

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appUrl = ScriptApp.getService().getUrl() || '';
  return t.evaluate()
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** ---------- INITIALIZATION ---------- */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('DB_ID');
  let ss;

  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(APP.NAME + ' Database');
    props.setProperty('DB_ID', ss.getId());
  }

  Object.keys(HEADERS).forEach(name => ensureSheet_(ss, name, HEADERS[name]));

  seedSettings_(ss);
  seedAdmin_(ss);
  seedDemoData_(ss);

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    message: 'Setup selesai. Login admin: ' + APP.DEFAULT_ADMIN_NIK +
      ' / ' + APP.DEFAULT_ADMIN_PASSWORD + ' lalu segera ganti password.'
  };
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    if (current.slice(0,headers.length).join('|') !== headers.join('|')) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
  }
  return sh;
}

function seedSettings_(ss) {
  const sh = ss.getSheetByName(APP.SHEETS.SETTINGS);
  const rows = getRows_(sh);
  const existing = {};
  rows.forEach(r => existing[r.key] = true);

  const settings = {
    appName: APP.NAME,
    rt: APP.DEFAULT_RT,
    rw: APP.DEFAULT_RW,
    village: APP.DEFAULT_VILLAGE,
    district: APP.DEFAULT_DISTRICT,
    regency: APP.DEFAULT_REGENCY,
    province: APP.DEFAULT_PROVINCE,
    chairmanName: 'Bapak Andi Susanto',
    chairmanPhone: '0812-3456-7890',
    emergencyMessage: 'Gunakan SOS hanya untuk keadaan darurat.',
    logoDriveId: '1oL-IGjrcykR1hSppDwYAxVLANPsXzR54',
    backgroundDriveId: '1ot-sGhbWr9TnjSTed5F0ooE2D6NjNdSh'
  };

  const now = new Date();
  Object.keys(settings).forEach(k => {
    if (!existing[k]) sh.appendRow([k, String(settings[k]), now, 'setup']);
  });
}

function seedAdmin_(ss) {
  const sh = ss.getSheetByName(APP.SHEETS.USERS);
  const rows = getRows_(sh);
  if (rows.some(r => String(r.nik).toLowerCase() === APP.DEFAULT_ADMIN_NIK.toLowerCase())) return;

  const salt = Utilities.getUuid();
  sh.appendRow([
    uid_(), APP.DEFAULT_ADMIN_NIK, 'Ketua RT', '081234567890', '',
    'Kantor RT', hash_(APP.DEFAULT_ADMIN_PASSWORD, salt), salt,
    'ADMIN', 'APPROVED', '', now_(), now_(), 'SYSTEM', ''
  ]);
}

function seedDemoData_(ss) {
  const assets = ss.getSheetByName(APP.SHEETS.ASSETS);
  if (assets.getLastRow() === 1) {
    [
      [uid_(),'Kursi Plastik','Perlengkapan','Kursi kegiatan warga',50,50,'Baik','Gudang RT','Dikembalikan setelah dipakai','ACTIVE',now_()],
      [uid_(),'Tenda 4x6','Perlengkapan','Tenda kegiatan lingkungan',2,2,'Baik','Gudang RT','Jaga kebersihan','ACTIVE',now_()],
      [uid_(),'Sound System','Elektronik','Speaker + mixer + mic',1,1,'Baik','Sekretariat RT','Wajib dikembalikan lengkap','ACTIVE',now_()]
    ].forEach(r => assets.appendRow(r));
  }

  const ann = ss.getSheetByName(APP.SHEETS.ANNOUNCEMENTS);
  if (ann.getLastRow() === 1) {
    ann.appendRow([uid_(),'Kerja Bakti Lingkungan RT 05',
      'Mari menjaga kebersihan lingkungan bersama.','Kegiatan','',now_(),'PUBLISHED',
      'SYSTEM',now_()]);
    ann.appendRow([uid_(),'Pembayaran Iuran Bulanan',
      'Iuran bulan berjalan dapat dibayarkan melalui bendahara RT.','Keuangan','',
      now_(),'PUBLISHED','SYSTEM',now_()]);
  }

  const act = ss.getSheetByName(APP.SHEETS.ACTIVITIES);
  if (act.getLastRow() === 1) {
    const d = new Date(); d.setDate(d.getDate()+2);
    act.appendRow([uid_(),'Kerja Bakti Lingkungan','Kerja bakti rutin warga',d,'07:00 WIB',
      'Lingkungan RT 05','Kebersihan','ACTIVE','SYSTEM',now_()]);
    const d2 = new Date(); d2.setDate(d2.getDate()+5);
    act.appendRow([uid_(),'Posyandu Balita','Pelayanan kesehatan balita',d2,'08:00 WIB',
      'Balai Warga','Kesehatan','ACTIVE','SYSTEM',now_()]);
  }
}

/** ---------- AUTH ---------- */

function registerCitizen(data) {
  data = data || {};
  const nik = clean_(data.nik);
  const name = clean_(data.name);
  const phone = clean_(data.phone);
  const address = clean_(data.address);
  const password = String(data.password || '');

  if (!nik || !name || !phone || !address || password.length < 6) {
    throw new Error('NIK, nama, WhatsApp, alamat, dan password minimal 6 karakter wajib diisi.');
  }

  const sh = db_().getSheetByName(APP.SHEETS.USERS);
  const rows = getRows_(sh);
  if (rows.some(r => String(r.nik).toLowerCase() === nik.toLowerCase())) {
    throw new Error('NIK/username sudah terdaftar.');
  }

  const salt = Utilities.getUuid();
  const id = uid_();
  sh.appendRow([
    id, nik, name, phone, clean_(data.email), address,
    hash_(password, salt), salt, 'WARGA', 'PENDING', '',
    now_(), '', '', ''
  ]);

  audit_('', name, 'REGISTER', 'AUTH', 'Registrasi warga: '+nik);
  return { ok:true, id:id, status:'PENDING',
    message:'Pendaftaran berhasil. Akun menunggu persetujuan Ketua RT.' };
}

function login(data) {
  data = data || {};
  const nik = clean_(data.nik);
  const password = String(data.password || '');
  if (!nik || !password) throw new Error('NIK/username dan password wajib diisi.');

  const sh = db_().getSheetByName(APP.SHEETS.USERS);
  const row = getRows_(sh).find(r => String(r.nik).toLowerCase() === nik.toLowerCase());
  if (!row) throw new Error('Akun tidak ditemukan.');

  if (row.status !== 'APPROVED') {
    if (row.status === 'PENDING') throw new Error('Akun belum disetujui Ketua RT.');
    if (row.status === 'REJECTED') throw new Error('Pendaftaran akun ditolak. Hubungi Ketua RT.');
    throw new Error('Akun tidak aktif.');
  }

  if (hash_(password, row.salt) !== row.passwordHash) throw new Error('Password salah.');

  const token = Utilities.getUuid() + Utilities.getUuid();
  const user = sanitizeUser_(row);
  CacheService.getScriptCache().put('sess_'+token, JSON.stringify(user), APP.SESSION_HOURS*3600);

  updateUserLastLogin_(row.id);
  audit_(row.id, row.name, 'LOGIN', 'AUTH', 'Login berhasil');
  return { ok:true, token:token, user:user, expiresIn:APP.SESSION_HOURS*3600 };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_'+token);
  return {ok:true};
}

function getSession(token) {
  const user = session_(token, false);
  return {ok:!!user, user:user || null};
}

/** ---------- PUBLIC / DASHBOARD ---------- */

function getPublicData() {
  const ss = db_();
  const settings = settings_();
  const announcements = getRows_(ss.getSheetByName(APP.SHEETS.ANNOUNCEMENTS))
    .filter(r => r.status === 'PUBLISHED')
    .sort(byDateDesc_).slice(0,8);
  const activities = getRows_(ss.getSheetByName(APP.SHEETS.ACTIVITIES))
    .filter(r => r.status !== 'CANCELLED')
    .sort(byDateAsc_).slice(0,8);

  return {
    settings:settings,
    announcements:announcements.map(publicAnnouncement_),
    activities:activities.map(publicActivity_),
    plans:getPlansPublic_(),
    stats:publicStats_()
  };
}

function getDashboard(token) {
  const user = session_(token);
  const ss = db_();
  const stats = adminStats_();
  const base = {
    user:user,
    stats:stats,
    notifications:getNotifications_(user.id).slice(0,10),
    announcements:getRows_(ss.getSheetByName(APP.SHEETS.ANNOUNCEMENTS)).filter(r=>r.status==='PUBLISHED').sort(byDateDesc_).slice(0,6).map(publicAnnouncement_),
    activities:getRows_(ss.getSheetByName(APP.SHEETS.ACTIVITIES)).filter(r=>r.status!=='CANCELLED').sort(byDateAsc_).slice(0,6).map(publicActivity_)
  };

  if (user.role === 'ADMIN' || user.role === 'PENGURUS') {
    base.pendingUsers = getRows_(ss.getSheetByName(APP.SHEETS.USERS)).filter(r=>r.status==='PENDING').map(sanitizeUser_);
    base.recentBorrowings = getRows_(ss.getSheetByName(APP.SHEETS.BORROWINGS)).sort(byDateDesc_).slice(0,8);
    base.sos = getRows_(ss.getSheetByName(APP.SHEETS.SOS)).filter(r=>r.status==='OPEN').sort(byDateDesc_).slice(0,8);
    base.feedback = getRows_(ss.getSheetByName(APP.SHEETS.FEEDBACK)).sort(byDateDesc_).slice(0,8);
  }
  return base;
}

/** ---------- ADMIN: USERS ---------- */

function listUsers(token, filter) {
  requireRole_(token,['ADMIN','PENGURUS']);
  const rows = getRows_(db_().getSheetByName(APP.SHEETS.USERS)).map(sanitizeUser_);
  filter = filter || {};
  return rows.filter(r => !filter.status || r.status === filter.status)
             .filter(r => !filter.q || (r.nik+' '+r.name+' '+r.phone).toLowerCase().includes(String(filter.q).toLowerCase()));
}

function approveUser(token, userId, action, note) {
  const actor = requireRole_(token,['ADMIN','PENGURUS']);
  if (!['APPROVE','REJECT'].includes(action)) throw new Error('Aksi tidak valid.');

  const sh = db_().getSheetByName(APP.SHEETS.USERS);
  const rowIndex = findRowIndex_(sh,'id',userId);
  if (rowIndex < 2) throw new Error('Warga tidak ditemukan.');

  const row = rowObject_(sh,rowIndex);
  const status = action === 'APPROVE' ? 'APPROVED' : 'REJECTED';
  sh.getRange(rowIndex, indexOf_(sh,'status')).setValue(status);
  if (action === 'APPROVE') {
    sh.getRange(rowIndex, indexOf_(sh,'approvedAt')).setValue(now_());
    sh.getRange(rowIndex, indexOf_(sh,'approvedBy')).setValue(actor.name);
    notify_(userId,'Akun disetujui','Akun Portal RT Anda telah disetujui Ketua RT.','ACCOUNT');
  } else {
    notify_(userId,'Pendaftaran ditolak','Pendaftaran akun Anda ditolak. '+(note||''),'ACCOUNT');
  }
  audit_(actor.id,actor.name,action,'USERS',row.name+' ('+row.nik+')');
  return {ok:true,status:status};
}

function changeUserStatus(token,userId,status) {
  const actor = requireRole_(token,['ADMIN']);
  const allowed = ['APPROVED','SUSPENDED','REJECTED'];
  if (!allowed.includes(status)) throw new Error('Status tidak valid.');
  const sh=db_().getSheetByName(APP.SHEETS.USERS);
  const i=findRowIndex_(sh,'id',userId);
  if(i<2) throw new Error('Warga tidak ditemukan.');
  sh.getRange(i,indexOf_(sh,'status')).setValue(status);
  audit_(actor.id,actor.name,'STATUS_CHANGE','USERS',userId+' -> '+status);
  return {ok:true};
}

function resetPassword(token,userId,newPassword) {
  const actor=requireRole_(token,['ADMIN']);
  if(String(newPassword||'').length<6) throw new Error('Password minimal 6 karakter.');
  const sh=db_().getSheetByName(APP.SHEETS.USERS);
  const i=findRowIndex_(sh,'id',userId);
  if(i<2) throw new Error('Warga tidak ditemukan.');
  const salt=Utilities.getUuid();
  sh.getRange(i,indexOf_(sh,'salt')).setValue(salt);
  sh.getRange(i,indexOf_(sh,'passwordHash')).setValue(hash_(newPassword,salt));
  audit_(actor.id,actor.name,'RESET_PASSWORD','USERS',userId);
  return {ok:true};
}

/** ---------- PROFILE ---------- */

function updateProfile(token,data) {
  const user=session_(token);
  data=data||{};
  const sh=db_().getSheetByName(APP.SHEETS.USERS);
  const i=findRowIndex_(sh,'id',user.id);
  if(i<2) throw new Error('Akun tidak ditemukan.');

  ['name','phone','email','address','avatar'].forEach(k=>{
    if(Object.prototype.hasOwnProperty.call(data,k)) {
      sh.getRange(i,indexOf_(sh,k)).setValue(clean_(data[k]));
    }
  });
  audit_(user.id,user.name,'UPDATE','PROFILE','Profil diperbarui');
  return {ok:true,user:sanitizeUser_(rowObject_(sh,i))};
}

function changeOwnPassword(token,oldPassword,newPassword) {
  const user=session_(token);
  if(String(newPassword||'').length<6) throw new Error('Password baru minimal 6 karakter.');
  const sh=db_().getSheetByName(APP.SHEETS.USERS);
  const i=findRowIndex_(sh,'id',user.id);
  const row=rowObject_(sh,i);
  if(hash_(oldPassword,row.salt)!==row.passwordHash) throw new Error('Password lama salah.');
  const salt=Utilities.getUuid();
  sh.getRange(i,indexOf_(sh,'salt')).setValue(salt);
  sh.getRange(i,indexOf_(sh,'passwordHash')).setValue(hash_(newPassword,salt));
  return {ok:true};
}

/** ---------- ANNOUNCEMENTS ---------- */

function listAnnouncements(token) {
  const user=session_(token);
  const rows=getRows_(db_().getSheetByName(APP.SHEETS.ANNOUNCEMENTS)).sort(byDateDesc_);
  return (user.role==='ADMIN'||user.role==='PENGURUS') ? rows : rows.filter(r=>r.status==='PUBLISHED');
}
function saveAnnouncement(token,data) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  data=data||{};
  required_(data,['title','content']);
  const sh=db_().getSheetByName(APP.SHEETS.ANNOUNCEMENTS);
  const id=data.id || uid_();
  const payload=[id,clean_(data.title),clean_(data.content),clean_(data.category||'Umum'),
    clean_(data.image||''),data.publishAt||now_(),data.status||'PUBLISHED',
    actor.id,now_()];
  upsert_(sh,'id',id,payload);
  notifyAll_(clean_(data.title),clean_(data.content),'ANNOUNCEMENT');
  audit_(actor.id,actor.name,'SAVE','ANNOUNCEMENTS',clean_(data.title));
  return {ok:true,id:id};
}
function deleteAnnouncement(token,id) {
  const actor=requireRole_(token,['ADMIN']);
  deleteById_(db_().getSheetByName(APP.SHEETS.ANNOUNCEMENTS),id);
  audit_(actor.id,actor.name,'DELETE','ANNOUNCEMENTS',id);
  return {ok:true};
}

/** ---------- ACTIVITIES ---------- */

function listActivities(token) {
  session_(token);
  return getRows_(db_().getSheetByName(APP.SHEETS.ACTIVITIES)).sort(byDateAsc_);
}
function saveActivity(token,data) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  data=data||{}; required_(data,['title','date']);
  const sh=db_().getSheetByName(APP.SHEETS.ACTIVITIES);
  const id=data.id||uid_();
  upsert_(sh,'id',id,[id,clean_(data.title),clean_(data.description||''),data.date,
    clean_(data.time||''),clean_(data.location||''),clean_(data.category||'Umum'),
    clean_(data.status||'ACTIVE'),actor.id,now_()]);
  notifyAll_('Kegiatan: '+clean_(data.title),formatDate_(data.date)+' '+clean_(data.time||''),'ACTIVITY');
  audit_(actor.id,actor.name,'SAVE','ACTIVITIES',clean_(data.title));
  return {ok:true,id:id};
}
function deleteActivity(token,id) {
  const actor=requireRole_(token,['ADMIN']);
  deleteById_(db_().getSheetByName(APP.SHEETS.ACTIVITIES),id);
  audit_(actor.id,actor.name,'DELETE','ACTIVITIES',id);
  return {ok:true};
}

/** ---------- RONDA ---------- */

function listRonda(token) {
  session_(token);
  return getRows_(db_().getSheetByName(APP.SHEETS.RONDA)).sort(byDateAsc_);
}
function saveRonda(token,data) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  data=data||{}; required_(data,['date','shift','groupName']);
  const sh=db_().getSheetByName(APP.SHEETS.RONDA);
  const id=data.id||uid_();
  upsert_(sh,'id',id,[id,data.date,clean_(data.shift),clean_(data.groupName),
    clean_(data.members||''),clean_(data.location||'Pos Ronda'),clean_(data.status||'ACTIVE'),
    clean_(data.notes||''),now_()]);
  audit_(actor.id,actor.name,'SAVE','RONDA',data.date+' '+data.groupName);
  return {ok:true,id:id};
}

/** ---------- ASSETS ---------- */

function listAssets(token) {
  session_(token);
  return getRows_(db_().getSheetByName(APP.SHEETS.ASSETS)).filter(r=>r.status!=='DELETED');
}
function saveAsset(token,data) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  data=data||{}; required_(data,['name','quantity']);
  const sh=db_().getSheetByName(APP.SHEETS.ASSETS);
  const id=data.id||uid_();
  const q=Math.max(0,Number(data.quantity)||0);
  const available=Math.max(0,Number(data.available===''?q:data.available)||0);
  upsert_(sh,'id',id,[id,clean_(data.name),clean_(data.category||'Umum'),
    clean_(data.description||''),q,available,clean_(data.condition||'Baik'),
    clean_(data.location||''),clean_(data.rules||''),clean_(data.status||'ACTIVE'),now_()]);
  audit_(actor.id,actor.name,'SAVE','ASSETS',data.name);
  return {ok:true,id:id};
}
function deleteAsset(token,id) {
  const actor=requireRole_(token,['ADMIN']);
  const sh=db_().getSheetByName(APP.SHEETS.ASSETS);
  const i=findRowIndex_(sh,'id',id); if(i<2) throw new Error('Aset tidak ditemukan.');
  sh.getRange(i,indexOf_(sh,'status')).setValue('DELETED');
  audit_(actor.id,actor.name,'DELETE','ASSETS',id);
  return {ok:true};
}

function borrowAsset(token,data) {
  const user=session_(token);
  data=data||{}; required_(data,['assetId','startDate','endDate','purpose']);
  const ss=db_(), ash=ss.getSheetByName(APP.SHEETS.ASSETS);
  const ai=findRowIndex_(ash,'id',data.assetId); if(ai<2) throw new Error('Aset tidak ditemukan.');
  const asset=rowObject_(ash,ai);
  const qty=Math.max(1,Number(data.quantity)||1);
  if(Number(asset.available)<qty) throw new Error('Stok aset tidak mencukupi.');

  const sh=ss.getSheetByName(APP.SHEETS.BORROWINGS);
  const id=uid_();
  sh.appendRow([id,user.id,data.assetId,asset.name,data.startDate,data.endDate,
    clean_(data.purpose),qty,clean_(data.documentUrl||''),'PENDING','',
    now_(),'','']);
  notifyRole_('ADMIN','Peminjaman aset baru','Pengajuan '+asset.name+' dari '+user.name,'BORROWING');
  audit_(user.id,user.name,'CREATE','BORROWINGS',asset.name);
  return {ok:true,id:id,status:'PENDING'};
}

function listBorrowings(token) {
  const user=session_(token);
  const rows=getRows_(db_().getSheetByName(APP.SHEETS.BORROWINGS));
  return (user.role==='ADMIN'||user.role==='PENGURUS') ? rows.sort(byDateDesc_) :
    rows.filter(r=>r.userId===user.id).sort(byDateDesc_);
}

function approveBorrowing(token,id,action,note) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  const ss=db_(), sh=ss.getSheetByName(APP.SHEETS.BORROWINGS);
  const i=findRowIndex_(sh,'id',id); if(i<2) throw new Error('Pengajuan tidak ditemukan.');
  const row=rowObject_(sh,i);
  if(row.status!=='PENDING') throw new Error('Pengajuan sudah diproses.');

  const status=action==='APPROVE'?'APPROVED':action==='REJECT'?'REJECTED':'';
  if(!status) throw new Error('Aksi tidak valid.');

  if(status==='APPROVED'){
    const ash=ss.getSheetByName(APP.SHEETS.ASSETS);
    const ai=findRowIndex_(ash,'id',row.assetId);
    const asset=rowObject_(ash,ai);
    if(Number(asset.available)<Number(row.quantity)) throw new Error('Stok aset sudah tidak cukup.');
    ash.getRange(ai,indexOf_(ash,'available')).setValue(Number(asset.available)-Number(row.quantity));
    sh.getRange(i,indexOf_(sh,'approvedAt')).setValue(now_());
    sh.getRange(i,indexOf_(sh,'approvedBy')).setValue(actor.name);
  }
  sh.getRange(i,indexOf_(sh,'status')).setValue(status);
  sh.getRange(i,indexOf_(sh,'reviewNote')).setValue(clean_(note||''));
  notify_(row.userId,'Pengajuan aset '+status.toLowerCase(),
    row.assetName+' — '+(note||''),'BORROWING');
  audit_(actor.id,actor.name,status,'BORROWINGS',row.assetName);
  return {ok:true,status:status};
}

/** ---------- FEEDBACK ---------- */

function submitFeedback(token,data) {
  const user=session_(token);
  data=data||{}; required_(data,['category','title','message']);
  const sh=db_().getSheetByName(APP.SHEETS.FEEDBACK);
  const id=uid_();
  sh.appendRow([id,user.id,user.name,clean_(data.category),clean_(data.title),
    clean_(data.message),clean_(data.imageUrl||''),'OPEN','',now_(),now_()]);
  notifyRole_('ADMIN','Kritik & saran baru',user.name+': '+data.title,'FEEDBACK');
  audit_(user.id,user.name,'CREATE','FEEDBACK',data.title);
  return {ok:true,id:id};
}

function listFeedback(token) {
  const user=session_(token);
  const rows=getRows_(db_().getSheetByName(APP.SHEETS.FEEDBACK)).sort(byDateDesc_);
  return (user.role==='ADMIN'||user.role==='PENGURUS')?rows:rows.filter(r=>r.userId===user.id);
}

function replyFeedback(token,id,reply,status) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  const sh=db_().getSheetByName(APP.SHEETS.FEEDBACK);
  const i=findRowIndex_(sh,'id',id); if(i<2) throw new Error('Data tidak ditemukan.');
  const row=rowObject_(sh,i);
  sh.getRange(i,indexOf_(sh,'reply')).setValue(clean_(reply||''));
  sh.getRange(i,indexOf_(sh,'status')).setValue(clean_(status||'ANSWERED'));
  sh.getRange(i,indexOf_(sh,'updatedAt')).setValue(now_());
  notify_(row.userId,'Tanggapan Kritik & Saran',reply||'Pengurus telah menanggapi laporan Anda.','FEEDBACK');
  audit_(actor.id,actor.name,'REPLY','FEEDBACK',id);
  return {ok:true};
}

/** ---------- SOS ---------- */

function sendSOS(token,data) {
  const user=session_(token);
  data=data||{};
  const sh=db_().getSheetByName(APP.SHEETS.SOS);
  const id=uid_();
  sh.appendRow([id,user.id,user.name,user.phone,clean_(data.location||'Lokasi tidak dibagikan'),
    clean_(data.message||'Darurat membutuhkan bantuan'),'OPEN','','',now_()]);
  notifyRole_('ADMIN','DARURAT SOS — '+user.name,
    (data.message||'Butuh bantuan segera')+' | '+(data.location||'Lokasi tidak tersedia'),'SOS');
  audit_(user.id,user.name,'SOS','SOS',id);
  return {ok:true,id:id,status:'OPEN'};
}

function listSOS(token) {
  const user=session_(token);
  const rows=getRows_(db_().getSheetByName(APP.SHEETS.SOS)).sort(byDateDesc_);
  return (user.role==='ADMIN'||user.role==='PENGURUS')?rows:rows.filter(r=>r.userId===user.id);
}

function handleSOS(token,id,status,note) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  const sh=db_().getSheetByName(APP.SHEETS.SOS);
  const i=findRowIndex_(sh,'id',id); if(i<2) throw new Error('SOS tidak ditemukan.');
  const row=rowObject_(sh,i);
  sh.getRange(i,indexOf_(sh,'status')).setValue(status||'HANDLED');
  sh.getRange(i,indexOf_(sh,'handledBy')).setValue(actor.name);
  sh.getRange(i,indexOf_(sh,'handledAt')).setValue(now_());
  if(note) sh.getRange(i,indexOf_(sh,'message')).setValue(row.message+' | Tindakan: '+clean_(note));
  notify_(row.userId,'Status SOS diperbarui','Status darurat: '+(status||'HANDLED'),'SOS');
  audit_(actor.id,actor.name,'HANDLE','SOS',id);
  return {ok:true};
}

/** ---------- DOCUMENTS ---------- */

function uploadDocument(token,data) {
  const user=session_(token);
  data=data||{}; required_(data,['name','type','data']);
  const bytes=Utilities.base64Decode(String(data.data).split(',').pop());
  const blob=Utilities.newBlob(bytes,data.mimeType||'application/octet-stream',clean_(data.name));
  const folder=getOrCreateFolder_('Portal RT Digital - Dokumen');
  const file=folder.createFile(blob);
  const sh=db_().getSheetByName(APP.SHEETS.DOCUMENTS);
  const id=uid_();
  sh.appendRow([id,user.id,file.getName(),clean_(data.type),file.getUrl(),
    clean_(data.note||''),'ACTIVE',now_()]);
  audit_(user.id,user.name,'UPLOAD','DOCUMENTS',file.getName());
  return {ok:true,id:id,url:file.getUrl(),name:file.getName()};
}

function listDocuments(token) {
  const user=session_(token);
  const rows=getRows_(db_().getSheetByName(APP.SHEETS.DOCUMENTS));
  return (user.role==='ADMIN'||user.role==='PENGURUS')?rows:rows.filter(r=>r.userId===user.id);
}

/** ---------- DUES & FINANCE ---------- */

function listDues(token) {
  const user=session_(token);
  const rows=getRows_(db_().getSheetByName(APP.SHEETS.DUES)).sort(byDateDesc_);
  return (user.role==='ADMIN'||user.role==='PENGURUS')?rows:rows.filter(r=>r.userId===user.id);
}

function saveDue(token,data) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  data=data||{}; required_(data,['userId','month','year','amount']);
  const us=findUser_(data.userId); if(!us) throw new Error('Warga tidak ditemukan.');
  const sh=db_().getSheetByName(APP.SHEETS.DUES);
  const id=data.id||uid_();
  upsert_(sh,'id',id,[id,us.id,us.name,clean_(data.month),Number(data.year),
    Number(data.amount)||0,clean_(data.status||'UNPAID'),data.paidAt||'',
    clean_(data.method||''),clean_(data.note||''),now_()]);
  if(data.status==='PAID') notify_(us.id,'Iuran telah dicatat','Iuran '+data.month+' '+data.year+' telah dibayar.','DUES');
  audit_(actor.id,actor.name,'SAVE','DUES',us.name+' '+data.month);
  return {ok:true,id:id};
}

function saveFinance(token,data) {
  const actor=requireRole_(token,['ADMIN']);
  data=data||{}; required_(data,['type','category','description','amount','date']);
  const sh=db_().getSheetByName(APP.SHEETS.FINANCE);
  const id=data.id||uid_();
  upsert_(sh,'id',id,[id,clean_(data.type),clean_(data.category),clean_(data.description),
    Number(data.amount)||0,data.date,clean_(data.method||''),clean_(data.reference||''),
    actor.id,now_()]);
  audit_(actor.id,actor.name,'SAVE','FINANCE',data.description);
  return {ok:true,id:id};
}

function listFinance(token) {
  requireRole_(token,['ADMIN','PENGURUS']);
  return getRows_(db_().getSheetByName(APP.SHEETS.FINANCE)).sort(byDateDesc_);
}

/** ---------- PLANS ---------- */

function listPlans(token) {
  session_(token);
  return getRows_(db_().getSheetByName(APP.SHEETS.PLANS)).sort(byDateAsc_);
}

function savePlan(token,data) {
  const actor=requireRole_(token,['ADMIN','PENGURUS']);
  data=data||{}; required_(data,['type','title','date']);
  const sh=db_().getSheetByName(APP.SHEETS.PLANS);
  const id=data.id||uid_();
  upsert_(sh,'id',id,[id,clean_(data.type),clean_(data.title),clean_(data.description||''),
    data.date,clean_(data.time||''),clean_(data.location||''),Number(data.budget)||0,
    clean_(data.status||'PLANNED'),actor.id,now_()]);
  notifyAll_('Rencana: '+data.title,(data.date||'')+' '+(data.time||''),'PLAN');
  audit_(actor.id,actor.name,'SAVE','PLANS',data.title);
  return {ok:true,id:id};
}

function getPlansPublic_() {
  return getRows_(db_().getSheetByName(APP.SHEETS.PLANS)).sort(byDateAsc_).slice(0,10);
}

/** ---------- NOTIFICATIONS ---------- */

function getNotifications(token) {
  const user=session_(token);
  return getNotifications_(user.id);
}

function markNotificationRead(token,id) {
  const user=session_(token);
  const sh=db_().getSheetByName(APP.SHEETS.NOTIFICATIONS);
  const i=findRowIndex_(sh,'id',id); if(i<2) return {ok:true};
  const row=rowObject_(sh,i);
  if(row.userId!==user.id) throw new Error('Akses ditolak.');
  sh.getRange(i,indexOf_(sh,'read')).setValue('TRUE');
  return {ok:true};
}

/** ---------- SETTINGS / EXPORT ---------- */

function getSettings(token) {
  requireRole_(token,['ADMIN','PENGURUS']);
  return settings_();
}

function saveSettings(token,data) {
  const actor=requireRole_(token,['ADMIN']);
  const sh=db_().getSheetByName(APP.SHEETS.SETTINGS);
  data=data||{};
  Object.keys(data).forEach(k=>{
    if(k==='token') return;
    const i=findRowIndex_(sh,'key',k);
    if(i<2) sh.appendRow([k,String(data[k]),now_(),actor.name]);
    else {
      sh.getRange(i,indexOf_(sh,'value')).setValue(String(data[k]));
      sh.getRange(i,indexOf_(sh,'updatedAt')).setValue(now_());
      sh.getRange(i,indexOf_(sh,'updatedBy')).setValue(actor.name);
    }
  });
  audit_(actor.id,actor.name,'SAVE','SETTINGS','Pengaturan diperbarui');
  return {ok:true,settings:settings_()};
}

function exportModule(token,module) {
  requireRole_(token,['ADMIN','PENGURUS']);
  const map = {
    users:'Users', announcements:'Announcements', activities:'Activities',
    ronda:'Ronda', assets:'Assets', borrowings:'Borrowings', feedback:'Feedback',
    sos:'SOS', documents:'Documents', dues:'Dues', finance:'Finance', plans:'Plans'
  };
  const name=map[String(module||'').toLowerCase()];
  if(!name) throw new Error('Modul export tidak valid.');
  const rows=getRows_(db_().getSheetByName(name));
  const csv=toCsv_(rows);
  const blob=Utilities.newBlob(csv,'text/csv',name+'-'+Utilities.formatDate(new Date(),APP.TZ,'yyyyMMdd-HHmm')+'.csv');
  const folder=getOrCreateFolder_('Portal RT Digital - Export');
  const file=folder.createFile(blob);
  return {ok:true,url:file.getUrl(),name:file.getName()};
}

/** ---------- TRIGGER / REMINDER ---------- */

function createDailyTrigger() {
  const exists=ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='dailyReminder');
  if(!exists) ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(6).create();
  return {ok:true};
}

function dailyReminder() {
  const ss=db_();
  const tomorrow=new Date(); tomorrow.setDate(tomorrow.getDate()+1);
  const acts=getRows_(ss.getSheetByName(APP.SHEETS.ACTIVITIES));
  acts.filter(a=>sameDay_(a.date,tomorrow)).forEach(a=>{
    notifyAll_('Pengingat kegiatan: '+a.title,
      formatDate_(a.date)+' '+(a.time||'')+' — '+(a.location||''),'REMINDER');
  });

  const plans=getRows_(ss.getSheetByName(APP.SHEETS.PLANS));
  plans.filter(a=>sameDay_(a.date,tomorrow)).forEach(a=>{
    notifyAll_('Pengingat rencana: '+a.title,
      formatDate_(a.date)+' '+(a.time||'')+' — '+(a.location||''),'REMINDER');
  });
}

/** ---------- HELPERS ---------- */

function db_() {
  const id=PropertiesService.getScriptProperties().getProperty('DB_ID');
  if(!id) throw new Error('Database belum dibuat. Jalankan setup() terlebih dahulu.');
  return SpreadsheetApp.openById(id);
}

function session_(token, throwError) {
  if(!token) {
    if(throwError===false) return null;
    throw new Error('Sesi login tidak ditemukan.');
  }
  const raw=CacheService.getScriptCache().get('sess_'+token);
  if(!raw) {
    if(throwError===false) return null;
    throw new Error('Sesi berakhir. Silakan login kembali.');
  }
  return JSON.parse(raw);
}

function requireRole_(token,roles) {
  const user=session_(token);
  if(!roles.includes(user.role)) throw new Error('Anda tidak memiliki izin untuk tindakan ini.');
  return user;
}

function settings_() {
  const rows=getRows_(db_().getSheetByName(APP.SHEETS.SETTINGS));
  const out={};
  rows.forEach(r=>out[r.key]=r.value);
  return out;
}

function publicStats_() {
  const ss=db_();
  const users=getRows_(ss.getSheetByName(APP.SHEETS.USERS));
  const acts=getRows_(ss.getSheetByName(APP.SHEETS.ACTIVITIES));
  const assets=getRows_(ss.getSheetByName(APP.SHEETS.ASSETS)).filter(a=>a.status==='ACTIVE');
  return {
    residents:users.filter(u=>u.status==='APPROVED'&&u.role==='WARGA').length,
    activities:acts.filter(a=>a.status!=='CANCELLED').length,
    assets:assets.length
  };
}

function adminStats_() {
  const ss=db_();
  const users=getRows_(ss.getSheetByName(APP.SHEETS.USERS));
  const dues=getRows_(ss.getSheetByName(APP.SHEETS.DUES));
  const bor=getRows_(ss.getSheetByName(APP.SHEETS.BORROWINGS));
  const sos=getRows_(ss.getSheetByName(APP.SHEETS.SOS));
  const finance=getRows_(ss.getSheetByName(APP.SHEETS.FINANCE));
  return {
    residents:users.filter(u=>u.status==='APPROVED'&&u.role==='WARGA').length,
    pendingUsers:users.filter(u=>u.status==='PENDING').length,
    duesPaid:dues.filter(d=>d.status==='PAID').reduce((s,r)=>s+(Number(r.amount)||0),0),
    pendingBorrowings:bor.filter(b=>b.status==='PENDING').length,
    openSOS:sos.filter(s=>s.status==='OPEN').length,
    income:finance.filter(f=>String(f.type).toUpperCase()==='INCOME').reduce((s,r)=>s+(Number(r.amount)||0),0),
    expense:finance.filter(f=>String(f.type).toUpperCase()==='EXPENSE').reduce((s,r)=>s+(Number(r.amount)||0),0)
  };
}

function getRows_(sh) {
  if(!sh || sh.getLastRow()<2) return [];
  const values=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  return values.map(row=>{
    const o={}; headers.forEach((h,i)=>o[h]=normalizeValue_(row[i])); return o;
  });
}

function rowObject_(sh,rowIndex) {
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const vals=sh.getRange(rowIndex,1,1,sh.getLastColumn()).getValues()[0];
  const o={}; headers.forEach((h,i)=>o[h]=normalizeValue_(vals[i])); return o;
}

function normalizeValue_(v) {
  if(v instanceof Date) return Utilities.formatDate(v,APP.TZ,"yyyy-MM-dd'T'HH:mm:ss");
  return v;
}

function indexOf_(sh,key) {
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const i=headers.indexOf(key);
  if(i<0) throw new Error('Kolom tidak ditemukan: '+key);
  return i+1;
}

function findRowIndex_(sh,key,value) {
  const col=indexOf_(sh,key);
  const n=sh.getLastRow()-1; if(n<=0) return -1;
  const vals=sh.getRange(2,col,n,1).getValues();
  for(let i=0;i<vals.length;i++) if(String(vals[i][0])===String(value)) return i+2;
  return -1;
}

function upsert_(sh,key,value,row) {
  const i=findRowIndex_(sh,key,value);
  if(i<2) sh.appendRow(row);
  else sh.getRange(i,1,1,row.length).setValues([row]);
}

function deleteById_(sh,id) {
  const i=findRowIndex_(sh,'id',id);
  if(i<2) throw new Error('Data tidak ditemukan.');
  sh.deleteRow(i);
}

function sanitizeUser_(r) {
  return {
    id:r.id, nik:r.nik, name:r.name, phone:r.phone, email:r.email,
    address:r.address, role:r.role, status:r.status, avatar:r.avatar,
    createdAt:r.createdAt, approvedAt:r.approvedAt, lastLogin:r.lastLogin
  };
}

function findUser_(id) {
  const sh=db_().getSheetByName(APP.SHEETS.USERS);
  const i=findRowIndex_(sh,'id',id);
  return i<2?null:rowObject_(sh,i);
}

function updateUserLastLogin_(id) {
  const sh=db_().getSheetByName(APP.SHEETS.USERS);
  const i=findRowIndex_(sh,'id',id);
  if(i>=2) sh.getRange(i,indexOf_(sh,'lastLogin')).setValue(now_());
}

function notify_(userId,title,message,type) {
  db_().getSheetByName(APP.SHEETS.NOTIFICATIONS)
    .appendRow([uid_(),userId,clean_(title),clean_(message),clean_(type||'INFO'),'FALSE','',now_()]);
}
function notifyAll_(title,message,type) {
  const users=getRows_(db_().getSheetByName(APP.SHEETS.USERS))
    .filter(u=>u.status==='APPROVED');
  const sh=db_().getSheetByName(APP.SHEETS.NOTIFICATIONS);
  users.forEach(u=>sh.appendRow([uid_(),u.id,clean_(title),clean_(message),clean_(type||'INFO'),'FALSE','',now_()]));
}
function notifyRole_(role,title,message,type) {
  const users=getRows_(db_().getSheetByName(APP.SHEETS.USERS))
    .filter(u=>u.status==='APPROVED' && (u.role===role || (role==='ADMIN'&&u.role==='PENGURUS')));
  const sh=db_().getSheetByName(APP.SHEETS.NOTIFICATIONS);
  users.forEach(u=>sh.appendRow([uid_(),u.id,clean_(title),clean_(message),clean_(type||'INFO'),'FALSE','',now_()]));
}
function getNotifications_(userId) {
  return getRows_(db_().getSheetByName(APP.SHEETS.NOTIFICATIONS))
    .filter(n=>n.userId===userId).sort(byDateDesc_);
}

function audit_(userId,actor,action,module,detail) {
  try {
    db_().getSheetByName(APP.SHEETS.AUDIT)
      .appendRow([uid_(),userId||'',actor||'',action,module,clean_(detail||''),now_()]);
  } catch(e) {}
}

function toCsv_(rows) {
  if(!rows.length) return '';
  const keys=Object.keys(rows[0]);
  const esc=s=>'"'+String(s==null?'':s).replace(/"/g,'""')+'"';
  return [keys.map(esc).join(','),...rows.map(r=>keys.map(k=>esc(r[k])).join(','))].join('\n');
}

function getOrCreateFolder_(name) {
  const it=DriveApp.getFoldersByName(name);
  return it.hasNext()?it.next():DriveApp.createFolder(name);
}

function hash_(password,salt) {
  const bytes=Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt)+'::'+String(password),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}

function uid_() {
  return Utilities.getUuid();
}
function now_() {
  return new Date();
}
function clean_(v) {
  return String(v==null?'':v).trim();
}
function required_(obj,keys) {
  keys.forEach(k=>{ if(obj[k]===undefined || String(obj[k]).trim()==='') throw new Error(k+' wajib diisi.'); });
}
function byDateDesc_(a,b) { return String(b.createdAt||b.date||'').localeCompare(String(a.createdAt||a.date||'')); }
function byDateAsc_(a,b) { return String(a.date||a.createdAt||'').localeCompare(String(b.date||b.createdAt||'')); }
function formatDate_(d) {
  if(!d) return '';
  try { return Utilities.formatDate(new Date(d),APP.TZ,'dd MMM yyyy'); } catch(e) { return String(d); }
}
function sameDay_(a,b) {
  try { return Utilities.formatDate(new Date(a),APP.TZ,'yyyy-MM-dd')===Utilities.formatDate(b,APP.TZ,'yyyy-MM-dd'); }
  catch(e){ return false; }
}
function publicAnnouncement_(r) {
  return {id:r.id,title:r.title,content:r.content,category:r.category,image:r.image,publishAt:r.publishAt,status:r.status};
}
function publicActivity_(r) {
  return {id:r.id,title:r.title,description:r.description,date:r.date,time:r.time,location:r.location,category:r.category,status:r.status};
}
