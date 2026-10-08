/**
 * ============================================================================
 * LINOVHR - SMART ENTERPRISE ATTENDANCE & HR MANAGEMENT SYSTEM  (v2)
 * Backend Code (Google Apps Script)
 *
 * Fitur baru:
 *  - Geofencing (validasi radius GPS kantor di server)
 *  - Foto selfie presensi (disimpan ke Google Drive)
 *  - Shift kerja dinamis + jadwal per karyawan
 *  - Swap shift (persetujuan rekan + atasan langsung)
 *  - Idempotent sync untuk presensi offline (client_id)
 *
 * CARA PAKAI:
 *  1. Jalankan setupDatabase() SEKALI (aman dijalankan ulang: hanya menambah
 *     sheet/kolom/setting yang belum ada, tidak menghapus data).
 *  2. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 *     Setiap ada perubahan kode, buat "New version" pada deployment.
 *  3. Disarankan set timeZone "Asia/Jakarta" di appsscript.json.
 * ============================================================================
 */

const DB_SHEETS = {
  EMPLOYEES: "Employees",
  ATTENDANCE: "Attendance",
  LEAVES: "LeaveRequests",
  SETTINGS: "Settings",
  SHIFTS: "Shifts",
  SCHEDULES: "Schedules",
  SWAPS: "ShiftSwaps"
};

const SELFIE_FOLDER = "LinovHR_Selfies";

const HEADERS = {
  Employees: ["nik", "name", "email", "role", "position", "department", "password", "status", "created_at", "supervisor_nik"],
  Attendance: ["id", "nik", "name", "date", "clock_in", "clock_out", "location_in", "location_out", "late_minutes", "work_hours", "status", "notes",
    "shift_id", "shift_name", "lat_in", "lng_in", "dist_in", "photo_in", "lat_out", "lng_out", "dist_out", "photo_out", "client_id", "synced_offline"],
  LeaveRequests: ["id", "nik", "name", "type", "start_date", "end_date", "duration_days", "reason", "status", "created_at"],
  Settings: ["key", "value"],
  Shifts: ["shift_id", "name", "start", "end"],
  Schedules: ["id", "nik", "name", "date", "shift_id"],
  ShiftSwaps: ["id", "requester_nik", "requester_name", "target_nik", "target_name", "date", "requester_shift_id", "target_shift_id", "reason",
    "peer_status", "supervisor_nik", "supervisor_status", "status", "created_at", "decided_at", "decided_by"]
};

// Kolom yang HARUS disimpan sebagai teks (agar Sheets tidak mengubah "08:00" / "2026-10-08" / "123456")
const TEXT_HEADERS = ["nik", "password", "date", "start_date", "end_date", "clock_in", "clock_out", "start", "end", "value",
  "supervisor_nik", "requester_nik", "target_nik", "shift_id", "requester_shift_id", "target_shift_id", "id", "client_id", "key"];
const TIME_HEADERS = ["clock_in", "clock_out", "start", "end"];

const DEFAULT_SETTINGS = {
  jam_masuk: "08:00",
  jam_keluar: "17:00",
  toleransi_keterlambatan: "15",
  office_name: "Kantor Pusat",
  office_lat: "",
  office_lng: "",
  radius_m: "100",
  geofence_enabled: "FALSE",
  selfie_required: "TRUE"
};

/* ============================ SETUP / MIGRASI ============================ */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date().toISOString();

  ensureSheet_(ss, DB_SHEETS.EMPLOYEES, HEADERS.Employees, [
    ["ADM001", "Admin HRD", "admin@linovhr.com", "Admin", "HR Manager", "Human Resources", "admin123", "Active", now, ""],
    ["SPV001", "Rina Wijaya", "rina@linovhr.com", "Supervisor", "Team Lead", "IT Department", "123456", "Active", now, ""],
    ["EMP001", "Budi Santoso", "budi@linovhr.com", "Employee", "Software Engineer", "IT Department", "123456", "Active", now, "SPV001"],
    ["EMP002", "Siti Rahma", "siti@linovhr.com", "Employee", "UI/UX Designer", "Product Design", "123456", "Active", now, "SPV001"]
  ]);
  ensureSheet_(ss, DB_SHEETS.ATTENDANCE, HEADERS.Attendance, null);
  ensureSheet_(ss, DB_SHEETS.LEAVES, HEADERS.LeaveRequests, null);
  ensureSheet_(ss, DB_SHEETS.SETTINGS, HEADERS.Settings, null);
  ensureSheet_(ss, DB_SHEETS.SHIFTS, HEADERS.Shifts, [
    ["SH-PAGI", "Pagi", "08:00", "17:00"],
    ["SH-SIANG", "Siang", "13:00", "22:00"],
    ["SH-MALAM", "Malam", "22:00", "06:00"]
  ]);
  ensureSheet_(ss, DB_SHEETS.SCHEDULES, HEADERS.Schedules, null);
  ensureSheet_(ss, DB_SHEETS.SWAPS, HEADERS.ShiftSwaps, null);

  // Setting default (hanya menambah key yang belum ada)
  const sSheet = ss.getSheetByName(DB_SHEETS.SETTINGS);
  const existing = readRows_(sSheet).map(r => r.key);
  Object.keys(DEFAULT_SETTINGS).forEach(k => {
    if (existing.indexOf(k) === -1) appendObj_(sSheet, { key: k, value: DEFAULT_SETTINGS[k] });
  });

  Logger.log("Database initialized / migrated successfully!");
  return "Database Setup Complete";
}

function ensureSheet_(ss, name, headers, seedRows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const current = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(String);
    headers.forEach(h => {
      if (current.indexOf(h) === -1) {
        sheet.getRange(1, current.length + 1).setValue(h);
        current.push(h);
      }
    });
  }

  // Format kolom teks
  const hdr = getHeaders_(sheet);
  hdr.forEach((h, i) => {
    if (TEXT_HEADERS.indexOf(h) !== -1) sheet.getRange(1, i + 1, sheet.getMaxRows(), 1).setNumberFormat("@");
  });

  if (seedRows && sheet.getLastRow() <= 1) {
    // Susun nilai sesuai urutan kolom sebenarnya di sheet
    const rows = seedRows.map(r => hdr.map(h => {
      const i = headers.indexOf(h);
      return i >= 0 && r[i] !== undefined ? r[i] : "";
    }));
    sheet.getRange(2, 1, rows.length, hdr.length).setValues(rows);
  }
  return sheet;
}

/* ============================ HTTP HANDLERS ============================ */

function doGet(e) {
  return createJsonResponse({ status: "success", message: "LinovHR GAS Backend API v2 is Active and Ready." });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const payload = postData.payload || {};

    const routes = {
      login: handleLogin,
      getDashboard: handleGetDashboard,
      clockIn: handleClockIn,
      clockOut: handleClockOut,
      getEmployees: handleGetEmployees,
      getColleagues: handleGetColleagues,
      addEmployee: handleAddEmployee,
      updateEmployee: handleUpdateEmployee,
      deleteEmployee: handleDeleteEmployee,
      getAttendanceHistory: handleGetAttendanceHistory,
      getLeaveRequests: handleGetLeaveRequests,
      submitLeave: handleSubmitLeave,
      updateLeaveStatus: handleUpdateLeaveStatus,
      getSettings: handleGetSettings,
      updateSettings: handleUpdateSettings,
      getShifts: handleGetShifts,
      saveShift: handleSaveShift,
      deleteShift: handleDeleteShift,
      getSchedules: handleGetSchedules,
      assignSchedule: handleAssignSchedule,
      getSwapRequests: handleGetSwapRequests,
      createSwap: handleCreateSwap,
      respondSwapPeer: handleRespondSwapPeer,
      respondSwapSupervisor: handleRespondSwapSupervisor
    };

    if (!routes[action]) return fail_("Action not recognized: " + action);
    return routes[action](payload);
  } catch (error) {
    return fail_(error.toString());
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function ok_(data, message) {
  const o = { status: "success" };
  if (data !== undefined) o.data = data;
  if (message) o.message = message;
  return createJsonResponse(o);
}
function fail_(message) {
  return createJsonResponse({ status: "error", message: message });
}

/* ============================ SHEET HELPERS ============================ */

function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}
function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(String);
}
function tz_() {
  return Session.getScriptTimeZone() || "Asia/Jakarta";
}
function todayStr_() {
  return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd");
}
function norm_(h, v) {
  if (v instanceof Date) {
    if (/date/.test(h)) return Utilities.formatDate(v, tz_(), "yyyy-MM-dd");
    if (TIME_HEADERS.indexOf(h) !== -1) return Utilities.formatDate(v, tz_(), "HH:mm");
    return v.toISOString();
  }
  if (h === "password" || h === "nik" || h === "supervisor_nik" || h === "id" || h === "client_id") return String(v);
  return v;
}
function readRows_(sheet) {
  const v = sheet.getDataRange().getValues();
  if (v.length <= 1) return [];
  const h = v[0].map(String);
  return v.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => { o[k] = norm_(k, r[j]); });
    return o;
  });
}
function sheetToObjects(sheet) {
  return readRows_(sheet).map(o => { delete o._row; return o; });
}
function appendObj_(sheet, obj) {
  const h = getHeaders_(sheet);
  sheet.appendRow(h.map(k => (obj[k] === undefined ? "" : obj[k])));
}
function updateObj_(sheet, row, obj) {
  const h = getHeaders_(sheet);
  Object.keys(obj).forEach(k => {
    const c = h.indexOf(k);
    if (c >= 0) sheet.getRange(row, c + 1).setValue(obj[k]);
  });
}
function isTrue_(v) { return String(v).toUpperCase() === "TRUE"; }
function isOpen_(v) { return v === "" || v === "-" || v === "--:--" || v === undefined || v === null; }
function toMin_(t) { const p = String(t).split(":").map(Number); return p[0] * 60 + p[1]; }

/* ============================ SETTINGS ============================ */

function getSettingsObj() {
  const obj = {};
  Object.keys(DEFAULT_SETTINGS).forEach(k => { obj[k] = DEFAULT_SETTINGS[k]; });
  readRows_(getSheet(DB_SHEETS.SETTINGS)).forEach(r => {
    let val = r.value;
    if (val instanceof Date) val = Utilities.formatDate(val, tz_(), "HH:mm");
    obj[r.key] = String(val === undefined || val === null ? "" : val);
  });
  return obj;
}

function handleGetSettings() {
  return ok_(getSettingsObj());
}

function handleUpdateSettings(payload) {
  const sheet = getSheet(DB_SHEETS.SETTINGS);
  const allowed = Object.keys(DEFAULT_SETTINGS);
  const rows = readRows_(sheet);
  allowed.forEach(k => {
    if (payload[k] === undefined || payload[k] === null) return;
    const val = String(payload[k]);
    const found = rows.find(r => r.key === k);
    if (found) updateObj_(sheet, found._row, { value: val });
    else appendObj_(sheet, { key: k, value: val });
  });
  return ok_(undefined, "Pengaturan sistem berhasil diperbarui!");
}

/* ============================ GEOFENCE & SELFIE ============================ */

function haversine_(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = x => x * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function checkGeofence_(settings, lat, lng) {
  const hasPos = lat !== undefined && lat !== null && lat !== "" && lng !== undefined && lng !== null && lng !== "";
  const officeLat = parseFloat(settings.office_lat);
  const officeLng = parseFloat(settings.office_lng);
  const hasOffice = !isNaN(officeLat) && !isNaN(officeLng);
  const result = { ok: true, distance: "" };

  if (hasPos && hasOffice) {
    result.distance = Math.round(haversine_(Number(lat), Number(lng), officeLat, officeLng));
  }
  if (isTrue_(settings.geofence_enabled)) {
    if (!hasOffice) return { ok: false, message: "Geofencing aktif tetapi koordinat kantor belum diatur oleh Admin." };
    if (!hasPos) return { ok: false, message: "Lokasi GPS wajib untuk presensi." };
    const radius = Number(settings.radius_m || 100);
    if (result.distance > radius) {
      return { ok: false, message: "Anda berada di luar radius kantor (" + result.distance + " m dari titik kantor, maks " + radius + " m)." };
    }
  }
  return result;
}

function saveSelfie_(dataUrl, nik, dateStr, timeStr, kind) {
  if (!dataUrl) return "";
  try {
    const m = String(dataUrl).match(/^data:(image\/\w+);base64,(.+)$/);
    if (!m) return "";
    const bytes = Utilities.base64Decode(m[2]);
    const name = nik + "_" + dateStr + "_" + String(timeStr).replace(":", "") + "_" + kind + ".jpg";
    const blob = Utilities.newBlob(bytes, m[1], name);
    const it = DriveApp.getFoldersByName(SELFIE_FOLDER);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(SELFIE_FOLDER);
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) {}
    return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w600";
  } catch (err) {
    Logger.log("saveSelfie_ error: " + err);
    return "";
  }
}

/* ============================ SHIFT RESOLUTION ============================ */

function getShiftsList_() {
  return sheetToObjects(getSheet(DB_SHEETS.SHIFTS));
}

function getShiftForDate_(nik, dateStr, settings) {
  const sch = sheetToObjects(getSheet(DB_SHEETS.SCHEDULES)).find(s => s.nik === nik && s.date === dateStr);
  if (sch) {
    const sh = getShiftsList_().find(x => x.shift_id === sch.shift_id);
    if (sh) return { id: sh.shift_id, name: sh.name, start: sh.start, end: sh.end };
  }
  return { id: "DEFAULT", name: "Reguler", start: settings.jam_masuk, end: settings.jam_keluar };
}

/* ============================ AUTH ============================ */

function handleLogin(payload) {
  const nik = String(payload.nik || "").trim();
  const password = String(payload.password || "");
  const employees = sheetToObjects(getSheet(DB_SHEETS.EMPLOYEES));
  const user = employees.find(emp =>
    (String(emp.nik) === nik || String(emp.email).toLowerCase() === nik.toLowerCase()) &&
    String(emp.password) === password);

  if (!user) return fail_("NIK / Email atau Password salah!");
  if (user.status && user.status !== "Active") return fail_("Akun tidak aktif. Hubungi Admin HRD.");
  delete user.password;
  return ok_(user);
}

/* ============================ ATTENDANCE ============================ */

function handleClockIn(payload) {
  const nik = payload.nik, name = payload.name, dateStr = payload.dateStr, timeStr = payload.timeStr;
  if (!nik || !dateStr || !timeStr) return fail_("Data presensi tidak lengkap.");

  const settings = getSettingsObj();
  const attSheet = getSheet(DB_SHEETS.ATTENDANCE);
  const rows = readRows_(attSheet);

  // Idempotent untuk sinkronisasi offline
  if (payload.clientId && rows.some(a => a.client_id === payload.clientId)) {
    return ok_({ duplicate: true }, "Sudah tersinkron sebelumnya.");
  }

  const geo = checkGeofence_(settings, payload.lat, payload.lng);
  if (!geo.ok) return fail_(geo.message);
  if (isTrue_(settings.selfie_required) && !payload.photo) return fail_("Foto selfie wajib untuk presensi.");

  if (rows.some(a => a.nik === nik && a.date === dateStr)) {
    return fail_("Anda sudah melakukan Clock In hari ini!");
  }

  const shift = getShiftForDate_(nik, dateStr, settings);
  const tol = Number(settings.toleransi_keterlambatan || 15);
  const startMin = toMin_(shift.start);
  const inMin = toMin_(timeStr);

  let status = "Tepat Waktu";
  let lateMinutes = 0;
  if (inMin > startMin + tol) {
    status = "Terlambat";
    lateMinutes = inMin - startMin;
  }

  const photoUrl = saveSelfie_(payload.photo, nik, dateStr, timeStr, "in");
  const newId = "ATT-" + Date.now();

  appendObj_(attSheet, {
    id: newId, nik: nik, name: name, date: dateStr, clock_in: timeStr, clock_out: "--:--",
    location_in: payload.location || "", location_out: "-",
    late_minutes: lateMinutes, work_hours: 0, status: status,
    notes: payload.offline ? "Absen masuk (offline sync)" : "Absen masuk via Web/PWA",
    shift_id: shift.id, shift_name: shift.name,
    lat_in: payload.lat || "", lng_in: payload.lng || "", dist_in: geo.distance,
    photo_in: photoUrl, client_id: payload.clientId || "", synced_offline: payload.offline ? "YA" : ""
  });

  return ok_({ id: newId, date: dateStr, clockIn: timeStr, status: status, lateMinutes: lateMinutes },
    "Berhasil Clock In (" + status + ")");
}

function handleClockOut(payload) {
  const nik = payload.nik, dateStr = payload.dateStr, timeStr = payload.timeStr;
  if (!nik || !dateStr || !timeStr) return fail_("Data presensi tidak lengkap.");

  const settings = getSettingsObj();
  const attSheet = getSheet(DB_SHEETS.ATTENDANCE);
  const rows = readRows_(attSheet).filter(a => a.nik === nik);

  const open = rows.filter(a => isOpen_(a.clock_out)).pop();
  if (!open) {
    const done = rows.find(a => a.date === dateStr && a.clock_out === timeStr);
    if (payload.offline && done) return ok_({ duplicate: true }, "Sudah tersinkron sebelumnya.");
    return fail_("Anda belum Clock In (atau sudah Clock Out)!");
  }

  const geo = checkGeofence_(settings, payload.lat, payload.lng);
  if (!geo.ok) return fail_(geo.message);
  if (isTrue_(settings.selfie_required) && !payload.photo) return fail_("Foto selfie wajib untuk presensi.");

  let diff = toMin_(timeStr) - toMin_(open.clock_in);
  if (diff < 0) diff += 1440; // shift malam melewati tengah malam
  const workHours = (diff / 60).toFixed(1);
  const photoUrl = saveSelfie_(payload.photo, nik, dateStr, timeStr, "out");

  updateObj_(attSheet, open._row, {
    clock_out: timeStr, location_out: payload.location || "-", work_hours: workHours,
    lat_out: payload.lat || "", lng_out: payload.lng || "", dist_out: geo.distance, photo_out: photoUrl
  });

  return ok_({ clockOut: timeStr, workHours: workHours }, "Berhasil Clock Out! Durasi Kerja: " + workHours + " Jam");
}

function handleGetAttendanceHistory(payload) {
  const nik = payload.nik;
  const all = sheetToObjects(getSheet(DB_SHEETS.ATTENDANCE));
  return ok_(nik ? all.filter(a => a.nik === nik) : all);
}

/* ============================ EMPLOYEES ============================ */

function handleGetEmployees() {
  const list = sheetToObjects(getSheet(DB_SHEETS.EMPLOYEES)).map(e => { delete e.password; return e; });
  return ok_(list);
}

// Data minimal rekan kerja (untuk form swap shift) - tanpa password/email
function handleGetColleagues(payload) {
  const list = sheetToObjects(getSheet(DB_SHEETS.EMPLOYEES))
    .filter(e => e.status === "Active" && e.role !== "Admin" && e.nik !== payload.nik)
    .map(e => ({ nik: e.nik, name: e.name, department: e.department, position: e.position }));
  return ok_(list);
}

function handleAddEmployee(payload) {
  const sheet = getSheet(DB_SHEETS.EMPLOYEES);
  if (sheetToObjects(sheet).find(e => e.nik === payload.nik)) return fail_("NIK sudah terdaftar!");
  appendObj_(sheet, {
    nik: payload.nik, name: payload.name, email: payload.email,
    role: payload.role || "Employee", position: payload.position || "Staff",
    department: payload.department || "General", password: payload.password || "123456",
    status: "Active", created_at: new Date().toISOString(), supervisor_nik: payload.supervisor_nik || ""
  });
  return ok_(undefined, "Karyawan berhasil ditambahkan!");
}

function handleUpdateEmployee(payload) {
  const sheet = getSheet(DB_SHEETS.EMPLOYEES);
  const row = readRows_(sheet).find(e => e.nik === payload.nik);
  if (!row) return fail_("Karyawan tidak ditemukan!");
  const upd = {};
  ["name", "email", "role", "position", "department", "status", "supervisor_nik"].forEach(k => {
    if (payload[k] !== undefined) upd[k] = payload[k];
  });
  if (payload.password) upd.password = payload.password; // password hanya berubah jika diisi
  updateObj_(sheet, row._row, upd);
  return ok_(undefined, "Data karyawan berhasil diperbarui!");
}

function handleDeleteEmployee(payload) {
  const sheet = getSheet(DB_SHEETS.EMPLOYEES);
  const row = readRows_(sheet).find(e => e.nik === payload.nik);
  if (!row) return fail_("Karyawan tidak ditemukan!");
  sheet.deleteRow(row._row);
  return ok_(undefined, "Karyawan berhasil dihapus!");
}

/* ============================ LEAVES ============================ */

function handleSubmitLeave(payload) {
  appendObj_(getSheet(DB_SHEETS.LEAVES), {
    id: "LV-" + Date.now(), nik: payload.nik, name: payload.name, type: payload.type,
    start_date: payload.startDate, end_date: payload.endDate, duration_days: payload.duration,
    reason: payload.reason, status: "Pending", created_at: new Date().toISOString()
  });
  return ok_(undefined, "Pengajuan izin/cuti berhasil dikirim!");
}

function handleGetLeaveRequests(payload) {
  const all = sheetToObjects(getSheet(DB_SHEETS.LEAVES)).map(l => {
    l.duration = l.duration_days; // alias untuk frontend
    return l;
  });
  return ok_(payload.nik ? all.filter(l => l.nik === payload.nik) : all);
}

function handleUpdateLeaveStatus(payload) {
  const sheet = getSheet(DB_SHEETS.LEAVES);
  const row = readRows_(sheet).find(l => l.id === payload.id);
  if (!row) return fail_("Data pengajuan tidak ditemukan!");
  updateObj_(sheet, row._row, { status: payload.status });
  return ok_(undefined, "Status pengajuan telah diubah menjadi: " + payload.status);
}

/* ============================ SHIFTS & SCHEDULES ============================ */

function handleGetShifts() {
  return ok_(getShiftsList_());
}

function handleSaveShift(payload) {
  const sheet = getSheet(DB_SHEETS.SHIFTS);
  if (!payload.name || !payload.start || !payload.end) return fail_("Nama, jam mulai, dan jam selesai wajib diisi.");
  const existing = payload.shift_id ? readRows_(sheet).find(s => s.shift_id === payload.shift_id) : null;
  if (existing) {
    updateObj_(sheet, existing._row, { name: payload.name, start: payload.start, end: payload.end });
    return ok_(undefined, "Shift diperbarui!");
  }
  appendObj_(sheet, { shift_id: "SH-" + Date.now(), name: payload.name, start: payload.start, end: payload.end });
  return ok_(undefined, "Shift baru ditambahkan!");
}

function handleDeleteShift(payload) {
  const used = sheetToObjects(getSheet(DB_SHEETS.SCHEDULES)).some(s => s.shift_id === payload.shift_id);
  if (used) return fail_("Shift masih dipakai di jadwal karyawan, tidak bisa dihapus.");
  const sheet = getSheet(DB_SHEETS.SHIFTS);
  const row = readRows_(sheet).find(s => s.shift_id === payload.shift_id);
  if (!row) return fail_("Shift tidak ditemukan!");
  sheet.deleteRow(row._row);
  return ok_(undefined, "Shift dihapus!");
}

function handleGetSchedules(payload) {
  let list = sheetToObjects(getSheet(DB_SHEETS.SCHEDULES));
  if (payload.nik) list = list.filter(s => s.nik === payload.nik);
  if (payload.from) list = list.filter(s => s.date >= payload.from);
  if (payload.to) list = list.filter(s => s.date <= payload.to);
  list.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  return ok_(list);
}

function addDays_(dateStr, n) {
  const p = dateStr.split("-").map(Number);
  const d = new Date(p[0], p[1] - 1, p[2] + n);
  return Utilities.formatDate(d, tz_(), "yyyy-MM-dd");
}

function handleAssignSchedule(payload) {
  const niks = payload.niks || [];
  if (!niks.length || !payload.startDate || !payload.endDate || !payload.shiftId) return fail_("Lengkapi karyawan, tanggal, dan shift.");
  if (payload.endDate < payload.startDate) return fail_("Tanggal selesai harus setelah tanggal mulai.");

  const emps = sheetToObjects(getSheet(DB_SHEETS.EMPLOYEES));
  const sheet = getSheet(DB_SHEETS.SCHEDULES);
  const rows = readRows_(sheet);
  let count = 0, d = payload.startDate, guard = 0;

  while (d <= payload.endDate && guard < 62) {
    niks.forEach(nik => {
      const emp = emps.find(e => e.nik === nik);
      if (!emp) return;
      const ex = rows.find(r => r.nik === nik && r.date === d);
      if (ex) updateObj_(sheet, ex._row, { shift_id: payload.shiftId });
      else appendObj_(sheet, { id: "SC-" + Date.now() + "-" + (count++), nik: nik, name: emp.name, date: d, shift_id: payload.shiftId });
    });
    d = addDays_(d, 1);
    guard++;
  }
  return ok_(undefined, "Jadwal shift berhasil disimpan!");
}

/* ============================ SWAP SHIFT ============================ */

const SWAP_PENDING = ["Menunggu Rekan", "Menunggu Atasan"];

function findEmp_(nik) {
  return sheetToObjects(getSheet(DB_SHEETS.EMPLOYEES)).find(e => e.nik === nik);
}

function handleGetSwapRequests(payload) {
  const me = findEmp_(payload.nik);
  if (!me) return ok_([]);
  let list = sheetToObjects(getSheet(DB_SHEETS.SWAPS));
  if (me.role !== "Admin") {
    list = list.filter(s => s.requester_nik === me.nik || s.target_nik === me.nik || s.supervisor_nik === me.nik);
  }
  const shifts = getShiftsList_();
  const nameOf = id => { const s = shifts.find(x => x.shift_id === id); return s ? s.name + " (" + s.start + "-" + s.end + ")" : id; };
  list = list.map(s => { s.requester_shift_name = nameOf(s.requester_shift_id); s.target_shift_name = nameOf(s.target_shift_id); return s; });
  list.reverse();
  return ok_(list);
}

function handleCreateSwap(payload) {
  const req = findEmp_(payload.requester_nik);
  const tgt = findEmp_(payload.target_nik);
  if (!req || !tgt) return fail_("Data karyawan tidak ditemukan.");
  if (req.nik === tgt.nik) return fail_("Tidak bisa menukar shift dengan diri sendiri.");
  if (!payload.date || payload.date < todayStr_()) return fail_("Tanggal swap tidak boleh di masa lalu.");

  const schedules = sheetToObjects(getSheet(DB_SHEETS.SCHEDULES));
  const rs = schedules.find(s => s.nik === req.nik && s.date === payload.date);
  const ts = schedules.find(s => s.nik === tgt.nik && s.date === payload.date);
  if (!rs || !ts) return fail_("Kedua karyawan harus punya jadwal shift pada tanggal tersebut.");
  if (rs.shift_id === ts.shift_id) return fail_("Shift kedua karyawan sama, tidak perlu ditukar.");

  const dup = sheetToObjects(getSheet(DB_SHEETS.SWAPS)).some(s =>
    s.date === payload.date && SWAP_PENDING.indexOf(s.status) !== -1 &&
    (s.requester_nik === req.nik || s.target_nik === req.nik || s.requester_nik === tgt.nik || s.target_nik === tgt.nik));
  if (dup) return fail_("Sudah ada pengajuan swap yang masih berjalan untuk tanggal tersebut.");

  appendObj_(getSheet(DB_SHEETS.SWAPS), {
    id: "SW-" + Date.now(), requester_nik: req.nik, requester_name: req.name, target_nik: tgt.nik, target_name: tgt.name,
    date: payload.date, requester_shift_id: rs.shift_id, target_shift_id: ts.shift_id, reason: payload.reason || "",
    peer_status: "Pending", supervisor_nik: req.supervisor_nik || "", supervisor_status: "Pending",
    status: "Menunggu Rekan", created_at: new Date().toISOString()
  });
  return ok_(undefined, "Pengajuan tukar shift terkirim ke " + tgt.name + ".");
}

function handleRespondSwapPeer(payload) {
  const sheet = getSheet(DB_SHEETS.SWAPS);
  const sw = readRows_(sheet).find(s => s.id === payload.id);
  if (!sw) return fail_("Pengajuan tidak ditemukan.");
  if (sw.target_nik !== payload.nik) return fail_("Anda bukan rekan yang dituju.");
  if (sw.status !== "Menunggu Rekan") return fail_("Pengajuan sudah diproses.");

  if (payload.accept) {
    updateObj_(sheet, sw._row, { peer_status: "Accepted", status: "Menunggu Atasan" });
    return ok_(undefined, "Anda menyetujui. Menunggu persetujuan atasan.");
  }
  updateObj_(sheet, sw._row, { peer_status: "Rejected", status: "Ditolak", decided_at: new Date().toISOString(), decided_by: payload.nik });
  return ok_(undefined, "Pengajuan swap ditolak.");
}

function handleRespondSwapSupervisor(payload) {
  const sheet = getSheet(DB_SHEETS.SWAPS);
  const sw = readRows_(sheet).find(s => s.id === payload.id);
  if (!sw) return fail_("Pengajuan tidak ditemukan.");
  if (sw.status !== "Menunggu Atasan") return fail_("Pengajuan belum/sudah diproses.");

  const me = findEmp_(payload.nik);
  if (!me) return fail_("Pengguna tidak ditemukan.");
  const allowed = me.role === "Admin" || (sw.supervisor_nik && sw.supervisor_nik === me.nik);
  if (!allowed) return fail_("Hanya atasan langsung atau Admin yang dapat memproses pengajuan ini.");

  const now = new Date().toISOString();
  if (!payload.approve) {
    updateObj_(sheet, sw._row, { supervisor_status: "Rejected", status: "Ditolak", decided_at: now, decided_by: me.nik });
    return ok_(undefined, "Pengajuan swap ditolak.");
  }

  // Validasi ulang jadwal lalu tukar
  const sSheet = getSheet(DB_SHEETS.SCHEDULES);
  const rows = readRows_(sSheet);
  const rs = rows.find(s => s.nik === sw.requester_nik && s.date === sw.date);
  const ts = rows.find(s => s.nik === sw.target_nik && s.date === sw.date);
  if (!rs || !ts) return fail_("Jadwal salah satu karyawan sudah berubah. Pengajuan tidak dapat diproses.");

  updateObj_(sSheet, rs._row, { shift_id: ts.shift_id });
  updateObj_(sSheet, ts._row, { shift_id: rs.shift_id });
  updateObj_(sheet, sw._row, { supervisor_status: "Approved", status: "Disetujui", decided_at: now, decided_by: me.nik });
  return ok_(undefined, "Swap shift disetujui & jadwal sudah ditukar.");
}

/* ============================ DASHBOARD ============================ */

function handleGetDashboard() {
  const todayStr = todayStr_();
  const employees = sheetToObjects(getSheet(DB_SHEETS.EMPLOYEES)).filter(e => e.role !== "Admin");
  const attendance = sheetToObjects(getSheet(DB_SHEETS.ATTENDANCE));
  const leaves = sheetToObjects(getSheet(DB_SHEETS.LEAVES));
  const swaps = sheetToObjects(getSheet(DB_SHEETS.SWAPS));

  const todayAtt = attendance.filter(a => a.date === todayStr);

  const weekly = [];
  const labels = ["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"];
  for (let i = 6; i >= 0; i--) {
    const d = addDays_(todayStr, -i);
    const p = d.split("-").map(Number);
    weekly.push({ date: d, label: labels[new Date(p[0], p[1] - 1, p[2]).getDay()], count: attendance.filter(a => a.date === d).length });
  }

  return ok_({
    totalEmployees: employees.length,
    totalHadirToday: todayAtt.length,
    totalTerlambatToday: todayAtt.filter(a => a.status === "Terlambat").length,
    pendingLeaves: leaves.filter(l => l.status === "Pending").length,
    pendingSwaps: swaps.filter(s => s.status === "Menunggu Atasan").length,
    weekly: weekly,
    todayAtt: todayAtt,
    settings: getSettingsObj()
  });
}
