/**
 * ============================================================
 *  ระบบลงทะเบียนชุมนุมออนไลน์ — โรงเรียนมหาชัยพิทยาคาร
 *  MHC Smart School Club Registration System
 *  เวอร์ชัน: v3.3.32 (29 ตุลาคม 2569)
 *  สถานศึกษา: โรงเรียนมหาชัยพิทยาคาร สพม.มหาสารคาม
 *  ผู้พัฒนาระบบ: ครูก้องนที อุ่นเจริญ | MHC Smart School v3.3.32
 * ============================================================
 */

const APP_CONFIG = {
  APP_NAME: 'ระบบลงทะเบียนชุมนุมออนไลน์ — โรงเรียนมหาชัยพิทยาคาร',
  SCHOOL_NAME: 'โรงเรียนมหาชัยพิทยาคาร',
  DEVELOPER_NAME: 'ครูก้องนที อุ่นเจริญ',
  VERSION: 'v3.3.32',
  CACHE_TTL: 30 // วินาทีสำหรับ CacheService ป้องกันโหลดเกินโควตาเมื่อ นร. เข้าพร้อมกัน
};

/**
 * Web App Entry Point: ให้บริการหน้าเว็บ HTML
 */
function doGet(e) {
  // รองรับ JSON API เผื่อเรียกทดสอบภายนอก (?api=getDatabase)
  if (e && e.parameter && e.parameter.api === 'getDatabase') {
    const data = getDatabaseFromSheet();
    return jsonOut_({ ok: true, data: data });
  }

  const htmlOutput = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return htmlOutput;
}

/**
 * Web App API Entry Point: รองรับการเรียกจาก GitHub Pages หรือภายนอกผ่าน HTTP POST
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (_) {
        payload = {};
      }
    }

    const fn = payload.fn;
    const args = payload.args || [];

    if (fn === 'getDatabaseFromSheet') {
      return jsonOut_({ ok: true, data: getDatabaseFromSheet() });
    } else if (fn === 'saveDatabaseToSheet') {
      const updated = saveDatabaseToSheet(args[0]);
      return jsonOut_({ ok: true, data: updated });
    } else if (fn === 'uploadClubCoverImage') {
      const url = uploadClubCoverImage(args[0], args[1]);
      return jsonOut_({ ok: true, data: url });
    }

    return jsonOut_({ ok: false, message: 'Forbidden or unknown function: ' + fn });
  } catch (err) {
    Logger.log('[doPost] Error: ' + (err.stack || err.message));
    return jsonOut_({ ok: false, message: err.message || 'Server error' });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ดึงข้อมูลระบบทั้งหมดจาก Google Sheets (Settings, Users, Classes, Clubs, Registrations)
 */
function getDatabaseFromSheet() {
  const cache = CacheService.getScriptCache();
  const cachedJson = cache.get('mhc_club_db');
  if (cachedJson) {
    try {
      return JSON.parse(cachedJson);
    } catch (_) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetsInitialized(ss);

  const db = {
    settings: readSettingsSheet(ss),
    users: readUsersSheet(ss),
    classes: readClassesSheet(ss),
    clubs: readClubsSheet(ss),
    registrations: readRegistrationsSheet(ss),
    students: readStudentsSheet(ss)
  };

  try {
    cache.put('mhc_club_db', JSON.stringify(db), APP_CONFIG.CACHE_TTL);
  } catch (err) {
    Logger.log('Cache put exceeded size limit: ' + err.message);
  }

  return db;
}

/**
 * บันทึกข้อมูลทั้งหมดกลับลง Google Sheets พร้อม LockService ป้องกัน Race Condition
 */
function saveDatabaseToSheet(db) {
  if (!db) throw new Error('ไม่พบข้อมูลที่จะบันทึก');

  const lock = LockService.getScriptLock();
  try {
    // รอคิวเขียนข้อมูลสูงสุด 15 วินาที
    lock.waitLock(15000);
  } catch (e) {
    throw new Error('ระบบกำลังมีผู้ใช้งานบันทึกข้อมูลหนาแน่น กรุณาลองใหม่อีกครั้ง');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheetsInitialized(ss);

    // 1. บันทึก Settings
    if (db.settings) {
      writeSettingsSheet(ss, db.settings);
    }

    // 2. บันทึก Users
    if (Array.isArray(db.users)) {
      writeUsersSheet(ss, db.users);
    }

    // 3. บันทึก Classes
    if (Array.isArray(db.classes)) {
      writeClassesSheet(ss, db.classes);
    }

    // 4. บันทึก Clubs
    if (Array.isArray(db.clubs)) {
      writeClubsSheet(ss, db.clubs);
    }

    // 5. บันทึก Registrations
    if (Array.isArray(db.registrations)) {
      writeRegistrationsSheet(ss, db.registrations);
    }

    // 6. บันทึก Students
    if (Array.isArray(db.students)) {
      writeStudentsSheet(ss, db.students);
    }

    // เคลียร์แคชเพื่อให้ทุกคนได้ข้อมูลล่าสุดทันที
    const cache = CacheService.getScriptCache();
    cache.remove('mhc_club_db');

    SpreadsheetApp.flush();
    return db;
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 *  HELPER: READ SHEET DATA
 * ============================================================ */

function readSettingsSheet(ss) {
  const sheet = ss.getSheetByName('Settings');
  const values = sheet.getDataRange().getValues();
  const settings = {
    startPeriod: '2026-10-01T08:00',
    endPeriod: '2026-11-30T16:30'
  };

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0]).trim();
    const val = values[i][1];
    if (key === 'startPeriod' && val) settings.startPeriod = String(val);
    if (key === 'endPeriod' && val) settings.endPeriod = String(val);
  }
  return settings;
}

function readUsersSheet(ss) {
  const sheet = ss.getSheetByName('Users');
  const values = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0] && !row[1]) continue;
    users.push({
      id: Number(row[0]) || (i),
      username: String(row[1] || ''),
      password: String(row[2] || ''),
      name: String(row[3] || ''),
      role: 'admin' // ระบบใช้เฉพาะผู้ดูแลระบบ (Admin)
    });
  }
  return users;
}

function readClassesSheet(ss) {
  const sheet = ss.getSheetByName('Classes');
  const values = sheet.getDataRange().getValues();
  const classes = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0] && !row[1]) continue;
    classes.push({
      id: Number(row[0]) || (i),
      name: String(row[1] || '')
    });
  }
  return classes;
}

function readClubsSheet(ss) {
  const sheet = ss.getSheetByName('Clubs');
  const values = sheet.getDataRange().getValues();
  const clubs = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0] && !row[1]) continue;
    clubs.push({
      id: Number(row[0]) || (i),
      name: String(row[1] || ''),
      cover: String(row[2] || ''),
      teachers: String(row[3] || ''), // รายชื่อครูผู้ดูแลชุมนุม (สามารถใส่ได้หลายคน)
      capacity: Number(row[4]) || 30
    });
  }
  return clubs;
}

function readRegistrationsSheet(ss) {
  const sheet = ss.getSheetByName('Registrations');
  const values = sheet.getDataRange().getValues();
  const registrations = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0] && !row[1]) continue;
    registrations.push({
      id: Number(row[0]) || (i),
      studentName: String(row[1] || ''),
      classId: Number(row[2]) || 0,
      clubId: Number(row[3]) || 0,
      timestamp: row[4] ? new Date(row[4]).toISOString() : new Date().toISOString()
    });
  }
  return registrations;
}

function readStudentsSheet(ss) {
  const sheet = ss.getSheetByName('Students');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const students = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0] && !row[1] && !row[2]) continue;
    students.push({
      id: Number(row[0]) || (i),
      studentCode: String(row[1] || ''),
      name: String(row[2] || ''),
      classId: Number(row[3]) || 0
    });
  }
  return students;
}

/* ============================================================
 *  HELPER: WRITE SHEET DATA (BATCH OPERATIONS)
 * ============================================================ */

function writeSettingsSheet(ss, settings) {
  const sheet = ss.getSheetByName('Settings');
  sheet.clearContents();
  const rows = [
    ['Key', 'Value'],
    ['startPeriod', settings.startPeriod || ''],
    ['endPeriod', settings.endPeriod || '']
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

function writeUsersSheet(ss, users) {
  const sheet = ss.getSheetByName('Users');
  sheet.clearContents();
  const rows = [['ID', 'Username', 'Password', 'Name', 'Role']];
  users.forEach(u => {
    rows.push([u.id, u.username, u.password, u.name, u.role]);
  });
  sheet.getRange(1, 1, rows.length, 5).setValues(rows);
}

function writeClassesSheet(ss, classes) {
  const sheet = ss.getSheetByName('Classes');
  sheet.clearContents();
  const rows = [['ID', 'ClassName']];
  classes.forEach(c => {
    rows.push([c.id, c.name]);
  });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

function writeClubsSheet(ss, clubs) {
  const sheet = ss.getSheetByName('Clubs');
  sheet.clearContents();
  const rows = [['ID', 'ClubName', 'CoverURL', 'Teachers', 'Capacity']];
  clubs.forEach(c => {
    const teachersVal = c.teachers ? (Array.isArray(c.teachers) ? c.teachers.join(', ') : String(c.teachers)) : '';
    rows.push([c.id, c.name, c.cover, teachersVal, c.capacity]);
  });
  sheet.getRange(1, 1, rows.length, 5).setValues(rows);
}

function writeRegistrationsSheet(ss, registrations) {
  const sheet = ss.getSheetByName('Registrations');
  sheet.clearContents();
  const rows = [['ID', 'StudentName', 'ClassID', 'ClubID', 'Timestamp']];
  registrations.forEach(r => {
    rows.push([r.id, r.studentName, r.classId, r.clubId, r.timestamp]);
  });
  sheet.getRange(1, 1, rows.length, 5).setValues(rows);
}

function writeStudentsSheet(ss, students) {
  let sheet = ss.getSheetByName('Students');
  if (!sheet) sheet = ss.insertSheet('Students');
  sheet.clearContents();
  const rows = [['ID', 'StudentCode', 'StudentName', 'ClassID']];
  students.forEach((s, idx) => {
    rows.push([s.id || (idx + 1), s.studentCode || '', s.name || '', s.classId || 0]);
  });
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
}

/* ============================================================
 *  HELPER: INITIALIZE DEFAULT SHEETS & DATA IF FRESH
 * ============================================================ */

function ensureSheetsInitialized(ss) {
  const required = ['Settings', 'Users', 'Classes', 'Clubs', 'Registrations', 'Students'];
  let createdAny = false;

  required.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      createdAny = true;
    }
  });

  if (createdAny) {
    // ตรวจสอบและลงข้อมูลเริ่มต้น
    const userSheet = ss.getSheetByName('Users');
    if (userSheet.getLastRow() <= 1) {
      writeUsersSheet(ss, [
        { id: 1, username: 'admin', password: '123', name: 'ครูก้องนที อุ่นเจริญ', role: 'admin' }
      ]);
    }

    const classSheet = ss.getSheetByName('Classes');
    if (classSheet.getLastRow() <= 1) {
      writeClassesSheet(ss, [
        { id: 101, name: 'ม.1/1' },
        { id: 102, name: 'ม.1/2' },
        { id: 201, name: 'ม.2/1' },
        { id: 202, name: 'ม.2/2' },
        { id: 301, name: 'ม.3/1' },
        { id: 302, name: 'ม.3/2' },
        { id: 401, name: 'ม.4/1' },
        { id: 501, name: 'ม.5/1' },
        { id: 601, name: 'ม.6/1' }
      ]);
    }

    const clubSheet = ss.getSheetByName('Clubs');
    if (clubSheet.getLastRow() <= 1) {
      writeClubsSheet(ss, [
        { id: 1, name: 'ชุมนุมหุ่นยนต์และปัญญาประดิษฐ์ (AI & Robotics)', cover: 'https://images.unsplash.com/photo-1485827404703-89b55fcc595e?auto=format&fit=crop&w=800&q=80', teachers: 'ครูก้องนที อุ่นเจริญ, ครูวชิราวุธ ขันติวงศ์', capacity: 30 },
        { id: 2, name: 'ชุมนุมดนตรีสากลและวงดุริยางค์', cover: 'https://images.unsplash.com/photo-1511671782779-c97d3d27a1d4?auto=format&fit=crop&w=800&q=80', teachers: 'ครูวชิราวุธ ขันติวงศ์, ครูนภัสสร มณีโชติ', capacity: 25 },
        { id: 3, name: 'ชุมนุมคณิตศาสตร์คิดเร็วและโอลิมปิกวิชาการ', cover: 'https://images.unsplash.com/photo-1635070041078-e363dbe005cb?auto=format&fit=crop&w=800&q=80', teachers: 'ครูวัลย์วิภา บุญสมบัติ', capacity: 35 },
        { id: 4, name: 'ชุมนุมกีฬาบาสเกตบอลและสันทนาการ', cover: 'https://images.unsplash.com/photo-1546519638-68e109498ffc?auto=format&fit=crop&w=800&q=80', teachers: 'ครูประเสริฐ ชัยนาท, ครูก้องนที อุ่นเจริญ', capacity: 30 }
      ]);
    }

    const setSheet = ss.getSheetByName('Settings');
    if (setSheet.getLastRow() <= 1) {
      writeSettingsSheet(ss, {
        startPeriod: '2026-10-01T08:00',
        endPeriod: '2026-11-30T16:30'
      });
    }

    const regSheet = ss.getSheetByName('Registrations');
    if (regSheet.getLastRow() === 0) {
      regSheet.appendRow(['ID', 'StudentName', 'ClassID', 'ClubID', 'Timestamp']);
    }

    const stuSheet = ss.getSheetByName('Students');
    if (stuSheet.getLastRow() <= 1) {
      writeStudentsSheet(ss, [
        { id: 1, studentCode: '10001', name: 'เด็กชายกิตติศักดิ์ ศรีวิชัย', classId: 101 },
        { id: 2, studentCode: '10002', name: 'เด็กหญิงพรพิมล ชื่นใจ', classId: 101 },
        { id: 3, studentCode: '10003', name: 'เด็กชายธนกฤต สุขเกษม', classId: 101 },
        { id: 4, studentCode: '10004', name: 'เด็กหญิงกัญญาณัฐ วงศ์วิจิตร', classId: 101 },
        { id: 5, studentCode: '10005', name: 'เด็กชายณัฐวุฒิ ทองหล่อ', classId: 102 },
        { id: 6, studentCode: '10006', name: 'เด็กหญิงอารียา สมบัติเจริญ', classId: 102 },
        { id: 7, studentCode: '20001', name: 'เด็กชายพงศกร พัฒนกิจ', classId: 201 },
        { id: 8, studentCode: '20002', name: 'เด็กหญิงกมลวรรณ บุญมี', classId: 201 },
        { id: 9, studentCode: '30001', name: 'เด็กชายนพรัตน์ แสนสุข', classId: 301 },
        { id: 10, studentCode: '40001', name: 'นายณัฐดนัย รัตนโชติ', classId: 401 }
      ]);
    }
  }
}

/* ============================================================
 *  HELPER: UPLOAD IMAGE TO GOOGLE DRIVE (MHC_Club_Covers)
 * ============================================================ */

/**
 * อัปโหลดรูปภาพหน้าปกชุมนุมขึ้น Google Drive โดยตรงในโฟลเดอร์ MHC_Club_Covers
 * @param {string} base64Data - Data URL ของรูปภาพ
 * @param {string} fileName - ชื่อไฟล์
 * @return {string} URL ของรูปภาพ (Google Drive direct CDN link)
 */
function uploadClubCoverImage(base64Data, fileName) {
  try {
    const splitIndex = base64Data.indexOf('base64,');
    if (splitIndex === -1) throw new Error('รูปแบบ Base64 ไม่ถูกต้อง');

    const contentTypeMatch = base64Data.match(/^data:(image\/[a-zA-Z0-9\+\.-]+);base64,/);
    const contentType = contentTypeMatch ? contentTypeMatch[1] : 'image/jpeg';
    const rawBase64 = base64Data.substring(splitIndex + 7);
    const bytes = Utilities.base64Decode(rawBase64);
    const safeName = (fileName || `club_${Date.now()}`).replace(/[^a-zA-Z0-9_\-\.]/g, '_');
    const blob = Utilities.newBlob(bytes, contentType, safeName.endsWith('.jpg') ? safeName : safeName + '.jpg');

    // ค้นหาหรือสร้างโฟลเดอร์เก็บภาพปกชุมนุม
    const folderName = 'MHC_Club_Covers';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return `https://lh3.googleusercontent.com/d/${file.getId()}`;
  } catch (err) {
    throw new Error('ไม่สามารถบันทึกภาพลง Google Drive ได้: ' + err.message);
  }
}
