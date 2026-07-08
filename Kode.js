/* ==================================================
   SCRIPT PENGHUBUNG (CLIENT BRIDGE)
   Identifier Library: CoreSystem
   ================================================== */

function _sid() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

function _sidAgendaSafe_() {
  try {
    var sid1 = _sid();
    if (sid1) return sid1;
  } catch (e1) {}
  return "";
}

function _cbtSid_(sid) {
  sid = (sid || "").toString().trim();
  return sid || _sid();
}

function _getOrCreatePengaturanSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Pengaturan");
  if (!sh) {
    sh = ss.insertSheet("Pengaturan");
    sh.getRange(1, 1, 1, 2).setValues([["KEY", "VALUE"]]);
    sh.getRange(1, 1, 1, 2).setFontWeight("bold");
  }
  return sh;
}

function _getPengaturanValue_(key) {
  try {
    var sh = _getOrCreatePengaturanSheet_();
    var lr = sh.getLastRow();
    if (lr < 1) return "";
    var vals = sh.getRange(1, 1, lr, 2).getValues();
    var wanted = (key || "").toString().trim().toUpperCase();
    for (var i = 0; i < vals.length; i++) {
      var k = (vals[i][0] || "").toString().trim().toUpperCase();
      if (k === wanted) return (vals[i][1] || "").toString().trim();
    }
  } catch (e) {}
  return "";
}

function _parseQueryStringMap_(qs) {
  var out = {};
  var raw = (qs || "").toString().replace(/^\?/, "");
  if (!raw) return out;
  var pairs = raw.split("&");
  for (var i = 0; i < pairs.length; i++) {
    var part = pairs[i];
    if (!part) continue;
    var idx = part.indexOf("=");
    var key = idx >= 0 ? part.substring(0, idx) : part;
    var val = idx >= 0 ? part.substring(idx + 1) : "";
    try {
      key = decodeURIComponent((key || "").replace(/\+/g, " "));
      val = decodeURIComponent((val || "").replace(/\+/g, " "));
    } catch (e) {}
    if (key && typeof out[key] === "undefined") out[key] = val;
  }
  return out;
}

function _upsertPengaturan_(key, value) {
  var sh = _getOrCreatePengaturanSheet_();
  var lastRow = sh.getLastRow();
  var data = sh.getRange(1, 1, Math.max(lastRow, 1), 2).getValues();
  var row = -1;

  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim() === key) {
      row = i + 1;
      break;
    }
  }
  if (row === -1) {
    row = lastRow + 1;
    sh.getRange(row, 1).setValue(key);
  }
  sh.getRange(row, 2).setValue(value);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("ADMIN SEKOLAH")
    .addItem("Sync Semua Data", "SYNC_SEMUA_DATA")
    .addSeparator()
    .addItem("Setup RDM Sync", "SETUP_RDM_SYNC")
    .addItem("Tarik ID Siswa RDM", "TARIK_ID_SISWA_RDM")
    .addItem("Tarik ID Guru RDM", "TARIK_ID_GURU_RDM")
    .addItem("Tarik Mapel RDM", "TARIK_MAPEL_RDM")
    .addItem("Tarik Kelas RDM", "TARIK_KELAS_RDM")
    .addItem("Tarik Ajar RDM", "TARIK_AJAR_RDM")
    .addItem("Tarik Nilai Harian RDM", "TARIK_NILAI_HARIAN_RDM")
    .addSeparator()
    .addItem("Sync Data Siswa", "syncDataSiswaTotal")
    .addItem("Sync UID Siswa ESP32", "syncSiswaByUidFirebase")
    .addItem("Sync Data Guru", "syncDataGuruTotal")
    .addItem("Sync UID Guru ESP32", "syncGuruByUidFirebase")
    .addItem("Sync Aturan Jam", "uploadAturanJam")
    .addItem("Sync Pengumuman", "syncPengumuman")
    .addToUi();
}

// --- SETUP AWAL (TULIS KE SHEET PENGATURAN) ---
function SETUP_SEKOLAH() {
  // Paste langsung isi object `firebaseConfig` dari Firebase console ke sini.
  // Contoh:
  // const firebaseConfig = { ... };
  // Yang dipakai cukup bagian { ... }-nya saja.
  var firebaseConfig = {
    apiKey: "AIzaSyCg1faQr-Zt9C8Y8Ghj2RAAmxc-TEnrcpQ",
    authDomain: "smartschool-13a64.firebaseapp.com",
    databaseURL: "https://smartschool-13a64-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "smartschool-13a64",
    storageBucket: "smartschool-13a64.firebasestorage.app",
    messagingSenderId: "217024692178",
    appId: "1:217024692178:web:0f551456e70af99bd7520b"
  };

  _upsertPengaturan_(
    "FIREBASE_CONFIG_JSON",
    JSON.stringify(firebaseConfig)
  );

  Logger.log("SETUP_SEKOLAH selesai. FIREBASE_CONFIG_JSON tersimpan. KEY lain diisi manual lewat sheet Pengaturan.");
}

function SETUP_RDM_SYNC() {
  var res = CoreSystem.setupPengaturanRdmSync(_sid());
  var pesan = (res && res.pesan) ? res.pesan : "Setup RDM Sync selesai.";
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
  return res;
}

function TARIK_ID_SISWA_RDM() {
  var res = (typeof CoreSystem.tarikIdSiswaRdm === "function")
    ? CoreSystem.tarikIdSiswaRdm(_sid())
    : {
        status: "error",
        pesan: "CoreSystem.tarikIdSiswaRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
      };
  var pesan = (res && res.pesan) ? res.pesan : "Tarik ID Siswa RDM selesai.";
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
  return res;
}

function TARIK_ID_GURU_RDM() {
  var res = (typeof CoreSystem.tarikIdGuruRdm === "function")
    ? CoreSystem.tarikIdGuruRdm(_sid())
    : {
        status: "error",
        pesan: "CoreSystem.tarikIdGuruRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
      };
  var pesan = (res && res.pesan) ? res.pesan : "Tarik ID Guru RDM selesai.";
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
  return res;
}

function TARIK_MAPEL_RDM() {
  var res = (typeof CoreSystem.tarikMapelRdm === "function")
    ? CoreSystem.tarikMapelRdm(_sid())
    : {
        status: "error",
        pesan: "CoreSystem.tarikMapelRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
      };
  var pesan = (res && res.pesan) ? res.pesan : "Tarik Mapel RDM selesai.";
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
  return res;
}

function TARIK_KELAS_RDM() {
  var res = (typeof CoreSystem.tarikKelasRdm === "function")
    ? CoreSystem.tarikKelasRdm(_sid())
    : {
        status: "error",
        pesan: "CoreSystem.tarikKelasRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
      };
  var pesan = (res && res.pesan) ? res.pesan : "Tarik Kelas RDM selesai.";
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
  return res;
}

function SYNC_SEMUA_DATA() {
  var log = [];
  try { syncDataSiswaTotal(); log.push("Siswa: OK"); } catch (e) { log.push("Siswa: GAGAL - " + e); }
  try { syncDataGuruTotal(); log.push("Guru: OK"); } catch (e) { log.push("Guru: GAGAL - " + e); }
  try { uploadAturanJam(); log.push("Aturan Jam: OK"); } catch (e) { log.push("Aturan Jam: GAGAL - " + e); }
  try { uploadDataJadwal(); log.push("Jadwal: OK"); } catch (e) { log.push("Jadwal: GAGAL - " + e); }
  try { syncPengumuman(); log.push("Pengumuman: OK"); } catch (e) { log.push("Pengumuman: GAGAL - " + e); }
  try { uploadDataPelanggaran(); log.push("Pelanggaran: OK"); } catch (e) { log.push("Pelanggaran: GAGAL - " + e); }

  var hasil = "Hasil Sync Semua:\n\n" + log.join("\n");

  // Coba tampilkan alert kalau dari Spreadsheet UI
  try {
    SpreadsheetApp.getUi().alert(hasil);
  } catch (errUi) {
    // Kalau bukan konteks UI (webapp/trigger), fallback ke log
    Logger.log(hasil);
  }

  return { status: "sukses", pesan: hasil };
}


// --- PINTU GERBANG WEB APP ---
function _attachSidToEvent_(e) {
  e = e || {};
  e.parameter = e.parameter || {};
  var sid = _sid();
  var appUrl = ScriptApp.getService().getUrl();
  if (!e.parameter.ssid) e.parameter.ssid = sid;
  if (!e.parameter.sid) e.parameter.sid = sid;
  if (!e.parameter.appUrl) e.parameter.appUrl = appUrl;
  return e;
}

function _cleanupLegacyWebhookSheets_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var names = ["Webhook_Wrapper_Ingress", "Webhook_Wrapper_Response", "Webhook_Ingress"];
    for (var i = 0; i < names.length; i++) {
      var sh = ss.getSheetByName(names[i]);
      if (sh) ss.deleteSheet(sh);
    }
  } catch (err) {}
}

function doGet(e) {
  e = _attachSidToEvent_(e);
  var out = CoreSystem.renderApp(e);
  _cleanupLegacyWebhookSheets_();
  return out;
}
function doPost(e) {
  e = _attachSidToEvent_(e);
  var out = CoreSystem.doPost(e);
  _cleanupLegacyWebhookSheets_();
  return out;
}

// A. LOGIN & DASHBOARD
function prosesLogin(u, p, rememberMe) { return CoreSystem.prosesLogin(u, p, _sid(), rememberMe); }
function cekSesiLoginToken(token, sid) { return CoreSystem.cekSesiLoginToken(token, sid || _sid()); }
function kirimPasswordLupaLogin(payload) { return CoreSystem.kirimPasswordLupaLogin(payload, _sid()); }
function gantiPasswordGuru(nip, passwordLama, passwordBaru) { return CoreSystem.gantiPasswordGuru(nip, passwordLama, passwordBaru, _sid()); }
function gantiPasswordSiswa(nis, passwordLama, passwordBaru) { return CoreSystem.gantiPasswordSiswa(nis, passwordLama, passwordBaru, _sid()); }
function updateFotoSiswaPortal(nis, fileData, fileName, mimeType) { return CoreSystem.updateFotoSiswaPortal(nis, fileData, fileName, mimeType, _sid()); }
function updateFotoGuruPortal(nip, fileData, fileName, mimeType) { return CoreSystem.updateFotoGuruPortal(nip, fileData, fileName, mimeType, _sid()); }
function getDataDashboardSekolah() { return CoreSystem.getDataDashboardSekolah(_sid()); }
function getDataDashboardKepala() { return CoreSystem.getDataDashboardKepala(_sid()); }
function jalankanOtomatisAlpha() { return CoreSystem.jalankanOtomatisAlpha(_sid()); }

// B. SYNC DATA (ADMIN)
function syncDataSiswaTotal() {
  var res = CoreSystem.syncDataSiswaTotal(_sid());
  try { syncSiswaByUidFirebase(); } catch (e) { Logger.log("[ESP32] sync siswa_by_uid gagal: " + e); }
  return res;
}
function syncDataGuruTotal() {
  var res = CoreSystem.syncDataGuruTotal(_sid());
  try { syncGuruByUidFirebase(); } catch (e) { Logger.log("[ESP32] sync guru_by_uid gagal: " + e); }
  return res;
}
function uploadAturanJam() { return CoreSystem.uploadAturanJam(_sid()); }
function syncFirebase() { return CoreSystem.uploadAturanJam(_sid()); } // Alias
function uploadDataPelanggaran() { return CoreSystem.uploadDataPelanggaran(_sid()); }
function syncPelanggaran() { return CoreSystem.uploadDataPelanggaran(_sid()); } // Alias
function uploadDataJadwal() { return CoreSystem.uploadDataJadwal(_sid()); }

// C. CRUD SISWA
function getSemuaSiswa(mode) {
  if ((mode || "").toString() === "__ALUMNI__") return CoreSystem.getSemuaAlumni(_sid());
  return CoreSystem.getSemuaSiswa(_sid());
}

function _wrapperFirebaseSettings_() {
  var url = "";
  var secret = "";

  try {
    var cfgRaw = _getPengaturanValue_("FIREBASE_CONFIG_JSON");
    if (cfgRaw) {
      var cfg = JSON.parse(cfgRaw);
      url = (cfg && cfg.databaseURL ? cfg.databaseURL : "").toString().trim();
    }
  } catch (e) {}

  url = url || _getPengaturanValue_("FIREBASE_URL");
  secret = _getPengaturanValue_("FIREBASE_SECRET");

  url = (url || "").toString().trim();
  secret = (secret || "").toString().trim();
  if (!url) throw new Error("FIREBASE_URL / databaseURL belum diset di Pengaturan.");
  if (!secret) throw new Error("FIREBASE_SECRET belum diset di Pengaturan.");

  url = url.replace(/\/+$/, "");
  if (url.indexOf("https://") !== 0 && url.indexOf("http://") !== 0) {
    url = "https://" + url;
  }

  return { url: url, secret: secret };
}

function _wrapperFirebasePut_(path, payload) {
  var fb = _wrapperFirebaseSettings_();
  path = (path || "").toString();
  if (path.charAt(0) !== "/") path = "/" + path;

  var endpoint = fb.url + path + ".json?auth=" + encodeURIComponent(fb.secret);
  var resp = UrlFetchApp.fetch(endpoint, {
    method: "put",
    contentType: "application/json",
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("Firebase PUT " + path + " gagal HTTP " + code + ": " + resp.getContentText());
  }
  return code;
}

function _wrapperNormalizeUid_(uid) {
  uid = (uid || "").toString().trim().toUpperCase();
  return uid.replace(/[.#$\/\[\]]/g, "");
}

function syncSiswaByUidFirebase() {
  var siswa = getSemuaSiswa();
  var out = {};
  var count = 0;

  if (Array.isArray(siswa)) {
    siswa.forEach(function(s) {
      s = s || {};
      var uid = _wrapperNormalizeUid_(s.uid || s.UID || s.rfid || s.kartu);
      var nama = (s.nama || s.nama_lengkap || s.Nama || "").toString().trim();
      if (!uid || !nama) return;

      var kelas = (s.kelas || s.Kelas || "").toString().trim();
      var statusRaw = (s.status || s.Status || s.aktif || "").toString().trim().toLowerCase();
      var aktif = !(statusRaw === "nonaktif" || statusRaw === "tidak aktif" || statusRaw === "false" || statusRaw === "0");

      out[uid] = {
        nama: nama,
        kelas: kelas,
        aktif: aktif,
        nis: (s.nis || s.NIS || "").toString().trim()
      };
      count++;
    });
  }

  _wrapperFirebasePut_("/siswa_by_uid", out);
  Logger.log("[ESP32] siswa_by_uid synced: " + count);
  return { status: "sukses", count: count, path: "/siswa_by_uid" };
}

function syncGuruByUidFirebase() {
  var guru = getSemuaGuru();
  var out = {};
  var count = 0;

  if (Array.isArray(guru)) {
    guru.forEach(function(g) {
      g = g || {};
      var uid = _wrapperNormalizeUid_(g.uid || g.UID || g.rfid || g.kartu);
      var nama = (g.nama || g.nama_guru || g.nama_lengkap || g.Nama || "").toString().trim();
      if (!uid || !nama) return;

      var statusRaw = (g.status || g.Status || g.aktif || "").toString().trim().toLowerCase();
      var aktif = !(statusRaw === "nonaktif" || statusRaw === "tidak aktif" || statusRaw === "false" || statusRaw === "0");

      out[uid] = {
        nama: nama,
        aktif: aktif,
        nip: (g.nip || g.NIP || g.nuptk || g.NUPTK || "").toString().trim()
      };
      count++;
    });
  }

  _wrapperFirebasePut_("/guru_by_uid", out);
  Logger.log("[ESP32] guru_by_uid synced: " + count);
  return { status: "sukses", count: count, path: "/guru_by_uid" };
}

function getSemuaAlumni() { return CoreSystem.getSemuaAlumni(_sid()); }
function simpanAlumni(form) { return CoreSystem.simpanAlumni(form, _sid()); }
function hapusAlumni(nis) { return CoreSystem.hapusAlumni(nis, _sid()); }
function exportDataAlumniCSV() { return CoreSystem.exportDataAlumniCSV(_sid()); }
function generateDaftarAlumniPDF() { return CoreSystem.generateDaftarAlumniPDF(_sid()); }
function getPengaturanSKL() { return CoreSystem.getPengaturanSKL(_sid()); }
function getProfilSekolahBackend() { return CoreSystem.getProfilSekolahBackend(_sid()); }
function simpanProfilSekolahBackend(data) { return CoreSystem.simpanProfilSekolahBackend(data, _sid()); }
function sinkronkanProfilSekolahDariLegacy() { return CoreSystem.sinkronkanProfilSekolahDariLegacy(_sid()); }
function simpanPengaturanSKL(form) { return CoreSystem.simpanPengaturanSKL(form, _sid()); }
function getDaftarSKL() { return CoreSystem.getDaftarSKL(_sid()); }
function simpanDataSKL(form) { return CoreSystem.simpanDataSKL(form, _sid()); }
function getDataInputNilaiSKLGuru(nip, filter) { return CoreSystem.getDataInputNilaiSKLGuru(nip, filter || {}, _sid()); }
function simpanNilaiSKLGuru(payload) { return CoreSystem.simpanNilaiSKLGuru(payload || {}, _sid()); }
function generatePDFSKL(nis) { return CoreSystem.generatePDFSKL(nis, _sid()); }
function generatePDFSKLMassal(angkatan) { return CoreSystem.generatePDFSKLMassal(angkatan, _sid()); }
function verifikasiSKL(input) { return CoreSystem.verifikasiSKL(input, _sid()); }
function getDashboardBK() { return CoreSystem.getDashboardBK(_sid()); }
function getStatusDhuhaGuru(nip) { return CoreSystem.getStatusDhuhaGuru(nip, _sid()); }
function submitAbsenDhuhaGuru(form) { return CoreSystem.submitAbsenDhuhaGuru(form, _sid()); }
function getDashboardDhuhaGuru(filter) { return CoreSystem.getDashboardDhuhaGuru(filter || {}, _sid()); }
function prosesPersetujuanDhuhaGuru(form) { return CoreSystem.prosesPersetujuanDhuhaGuru(form, _sid()); }
function getLaporanBK(filter) { return CoreSystem.getLaporanBK(filter, _sid()); }
function generateLaporanBKPDF(filter) { return CoreSystem.generateLaporanBKPDF(filter, _sid()); }
function getSemuaMasterBK() { return CoreSystem.getSemuaMasterBK(_sid()); }
function simpanMasterBK(form) { return CoreSystem.simpanMasterBK(form, _sid()); }
function hapusMasterBK(id) { return CoreSystem.hapusMasterBK(id, _sid()); }
function getSemuaKasusBK(filter) { return CoreSystem.getSemuaKasusBK(filter, _sid()); }
function simpanKasusBK(form) { return CoreSystem.simpanKasusBK(form, _sid()); }
function hapusKasusBK(id) { return CoreSystem.hapusKasusBK(id, _sid()); }
function kirimPanggilanOrtuBK(id) { return CoreSystem.kirimPanggilanOrtuBK(id, _sid()); }
function generateSuratPanggilanBKPDF(id) { return CoreSystem.generateSuratPanggilanBKPDF(id, _sid()); }
function getTimelineKasusBK(id) { return CoreSystem.getTimelineKasusBK(id, _sid()); }
function simpanTimelineKasusBK(form) { return CoreSystem.simpanTimelineKasusBK(form, _sid()); }
function hapusTimelineKasusBK(kasusId, timelineId) { return CoreSystem.hapusTimelineKasusBK(kasusId, timelineId, _sid()); }
function getDashboardUKS() { return CoreSystem.getDashboardUKS(_sid()); }
function getSemuaMasterKeluhanUKS() { return CoreSystem.getSemuaMasterKeluhanUKS(_sid()); }
function simpanMasterKeluhanUKS(form) { return CoreSystem.simpanMasterKeluhanUKS(form, _sid()); }
function hapusMasterKeluhanUKS(id) { return CoreSystem.hapusMasterKeluhanUKS(id, _sid()); }
function getSemuaMasterTindakanUKS() { return CoreSystem.getSemuaMasterTindakanUKS(_sid()); }
function simpanMasterTindakanUKS(form) { return CoreSystem.simpanMasterTindakanUKS(form, _sid()); }
function hapusMasterTindakanUKS(id) { return CoreSystem.hapusMasterTindakanUKS(id, _sid()); }
function getSemuaKunjunganUKS(filter) { return CoreSystem.getSemuaKunjunganUKS(filter, _sid()); }
function simpanKunjunganUKS(form) { return CoreSystem.simpanKunjunganUKS(form, _sid()); }
function hapusKunjunganUKS(id) { return CoreSystem.hapusKunjunganUKS(id, _sid()); }
function getObatMenipisUKS() { return CoreSystem.getObatMenipisUKS(_sid()); }
function generateLaporanStokUKSPDF() { return CoreSystem.generateLaporanStokUKSPDF(_sid()); }
function generateSuratUKSPDF(id) { return CoreSystem.generateSuratUKSPDF(id, _sid()); }
function getSemuaObatUKS() { return CoreSystem.getSemuaObatUKS(_sid()); }
function simpanObatUKS(form) { return CoreSystem.simpanObatUKS(form, _sid()); }
function hapusObatUKS(id) { return CoreSystem.hapusObatUKS(id, _sid()); }
function getLaporanUKS(filter) { return CoreSystem.getLaporanUKS(filter, _sid()); }
function generateLaporanUKSPDF(filter) { return CoreSystem.generateLaporanUKSPDF(filter, _sid()); }
function getDashboardPerpus() { return CoreSystem.getDashboardPerpus(_sid()); }
function getSemuaKategoriPerpus() { return CoreSystem.getSemuaKategoriPerpus(_sid()); }
function simpanKategoriPerpus(form) { return CoreSystem.simpanKategoriPerpus(form, _sid()); }
function hapusKategoriPerpus(id) { return CoreSystem.hapusKategoriPerpus(id, _sid()); }
function getSemuaBukuPerpus(filter) { return CoreSystem.getSemuaBukuPerpus(filter, _sid()); }
function simpanBukuPerpus(form) { return CoreSystem.simpanBukuPerpus(form, _sid()); }
function hapusBukuPerpus(id) { return CoreSystem.hapusBukuPerpus(id, _sid()); }
function getSemuaPeminjamanPerpus(filter) { return CoreSystem.getSemuaPeminjamanPerpus(filter, _sid()); }
function simpanPeminjamanPerpus(form) { return CoreSystem.simpanPeminjamanPerpus(form, _sid()); }
function prosesPengembalianPerpus(form) { return CoreSystem.prosesPengembalianPerpus(form, _sid()); }
function getLaporanPerpus(filter) { return CoreSystem.getLaporanPerpus(filter, _sid()); }
function generateLaporanPerpusPDF(filter) { return CoreSystem.generateLaporanPerpusPDF(filter, _sid()); }
function kirimPengingatPerpusHariIni() { return CoreSystem.kirimPengingatPerpusHariIni(_sid()); }
function setupTriggerPengingatPerpusPagi() { return CoreSystem.setupTriggerPengingatPerpusPagi(_sid()); }
function hapusTriggerPengingatPerpusPagi() { return CoreSystem.hapusTriggerPengingatPerpusPagi(_sid()); }
function simpanSiswa(f, sid) {
  var res = CoreSystem.simpanSiswa(f, sid || _sid());
  try { syncSiswaByUidFirebase(); } catch (e) { Logger.log("[ESP32] sync siswa_by_uid gagal: " + e); }
  return res;
}
function hapusSiswa(n, sid) {
  var res = CoreSystem.hapusSiswa(n, sid || _sid());
  try { syncSiswaByUidFirebase(); } catch (e) { Logger.log("[ESP32] sync siswa_by_uid gagal: " + e); }
  return res;
}

// D. CRUD GURU
function getSemuaGuru() { return CoreSystem.getSemuaGuru(_sid()); }
function getSemuaGuruFallback() { return CoreSystem.getSemuaGuruFallback(_sid()); }
function simpanGuru(f, sid) {
  var res = CoreSystem.simpanGuru(f, sid || _sid());
  try { syncGuruByUidFirebase(); } catch (e) { Logger.log("[ESP32] sync guru_by_uid gagal: " + e); }
  return res;
}
function hapusGuru(n, sid) {
  var res = CoreSystem.hapusGuru(n, sid || _sid());
  try { syncGuruByUidFirebase(); } catch (e) { Logger.log("[ESP32] sync guru_by_uid gagal: " + e); }
  return res;
}

// E. ATURAN & PELANGGARAN
function getAturanJam() { return CoreSystem.getAturanJam(_sid()); }
function simpanAturanSatuan(f) { return CoreSystem.simpanAturanSatuan(f, _sid()); }
function getAturanJamGuru() { return CoreSystem.getAturanJamGuru(_sid()); }
function simpanAturanJamGuru(f) { return CoreSystem.simpanAturanJamGuru(f, _sid()); }
function getPengaturanAbsensiGuru() { return CoreSystem.getPengaturanAbsensiGuru(_sid()); }
function simpanPengaturanAbsensiGuru(form) { return CoreSystem.simpanPengaturanAbsensiGuru(form, _sid()); }
function getMasterJamSekolah() { return CoreSystem.getMasterJamSekolah(_sid()); }
function simpanMasterJamSekolah(f) { return CoreSystem.simpanMasterJamSekolah(f, _sid()); }
function hapusMasterJamSekolah(id) { return CoreSystem.hapusMasterJamSekolah(id, _sid()); }
function getAturanJamKelas() { return CoreSystem.getAturanJamKelas(_sid()); }
function simpanAturanJamKelas(f) { return CoreSystem.simpanAturanJamKelas(f, _sid()); }
function hapusAturanJamKelas(id) { return CoreSystem.hapusAturanJamKelas(id, _sid()); }
function getSemuaHariLibur(filter) { return CoreSystem.getSemuaHariLibur(filter, _sid()); }
function simpanHariLibur(form) { return CoreSystem.simpanHariLibur(form, _sid()); }
function hapusHariLibur(id) { return CoreSystem.hapusHariLibur(id, _sid()); }
function cekStatusHariLibur(tanggal) { return CoreSystem.cekStatusHariLibur(tanggal, _sid()); }
function getDaftarPelanggaran() { return CoreSystem.getDaftarPelanggaran(_sid()); }
function getDaftarKelasUnique() { return CoreSystem.getDaftarKelasUnique(_sid()); }
function getDaftarTahunMasukSiswaUnique() { return CoreSystem.getDaftarTahunMasukSiswaUnique(_sid()); }
function getSiswaByKelas(k) { return CoreSystem.getSiswaByKelas(k, _sid()); }
function simpanPelanggaran(f) { return CoreSystem.simpanPelanggaran(f, _sid()); }
function simpanMasterPelanggaran(form) { return CoreSystem.simpanMasterPelanggaran(form, _sid()); }
function updateMasterPelanggaran(form) { return CoreSystem.updateMasterPelanggaran(form, _sid()); }
function getPengaturanPortalSiswa() { return CoreSystem.getPengaturanPortalSiswa(_sid()); }
function simpanPengaturanPortalSiswa(form) { return CoreSystem.simpanPengaturanPortalSiswa(form, _sid()); }
function getPengaturanTahunAjaranAktif() { return CoreSystem.getPengaturanTahunAjaranAktif(_sid()); }
function simpanPengaturanTahunAjaranAktif(form) { return CoreSystem.simpanPengaturanTahunAjaranAktif(form, _sid()); }
function getPengaturanWaCommandLog() { return CoreSystem.getPengaturanWaCommandLog(_sid()); }
function simpanPengaturanWaCommandLog(form) { return CoreSystem.simpanPengaturanWaCommandLog(form, _sid()); }
function getPengaturanTemplateBeritaAcaraRapat() { return CoreSystem.getPengaturanTemplateBeritaAcaraRapat(_sid()); }
function simpanPengaturanTemplateBeritaAcaraRapat(form) { return CoreSystem.simpanPengaturanTemplateBeritaAcaraRapat(form, _sid()); }
function getPengaturanTemplateNotulenRapat() { return CoreSystem.getPengaturanTemplateNotulenRapat(_sid()); }
function simpanPengaturanTemplateNotulenRapat(form) { return CoreSystem.simpanPengaturanTemplateNotulenRapat(form, _sid()); }
function getPengaturanFolderExportNilai() { return CoreSystem.getPengaturanFolderExportNilai(_sid()); }
function simpanPengaturanFolderExportNilai(form) { return CoreSystem.simpanPengaturanFolderExportNilai(form, _sid()); }
function getPengaturanBobotNilai() { return CoreSystem.getPengaturanBobotNilai(_sid()); }
function simpanPengaturanBobotNilai(form) { return CoreSystem.simpanPengaturanBobotNilai(form, _sid()); }
function getPengaturanRaporDigital() { return CoreSystem.getPengaturanRaporDigital(_sid()); }
function simpanPengaturanRaporDigital(form) { return CoreSystem.simpanPengaturanRaporDigital(form, _sid()); }
function getPengaturanKartuPelajar() { return CoreSystem.getPengaturanKartuPelajar(_sid()); }
function simpanPengaturanKartuPelajar(form) { return CoreSystem.simpanPengaturanKartuPelajar(form, _sid()); }
function generateKartuPelajarDashboard() { return CoreSystem.generateKartuPelajarDashboard(_sid()); }
function resetProgressKartuPelajar() { return CoreSystem.resetProgressKartuPelajar(_sid()); }

function authorizeSlidesAccess() {
  var pres = SlidesApp.create("AUTH_SLIDES_ACCESS");
  try {
    pres = SlidesApp.openById(pres.getId());
  } catch (e) {}
  try {
    DriveApp.getFileById(pres.getId()).setTrashed(true);
  } catch (e2) {}
  return { status: "sukses", pesan: "Slides authorization OK." };
}
function getRaporDigitalSiswa(nis, filter) { return CoreSystem.getRaporDigitalSiswa(nis, filter || {}, _sid()); }
function getDataKenaikanKelas(kelasAsal, tahunMasuk) { return CoreSystem.getDataKenaikanKelas(kelasAsal, tahunMasuk, _sid()); }
function prosesKenaikanKelas(form) { return CoreSystem.prosesKenaikanKelas(form, _sid()); }
function getMasterSPMB() { return CoreSystem.getMasterSPMB(_sid()); }
function simpanMasterSPMB(form) { return CoreSystem.simpanMasterSPMB(form, _sid()); }
function getDataPendaftarSPMB(filter) { return CoreSystem.getDataPendaftarSPMB(filter, _sid()); }
function simpanPendaftarSPMB(form) { return CoreSystem.simpanPendaftarSPMB(form, _sid()); }
function updateStatusPendaftarSPMB(form) { return CoreSystem.updateStatusPendaftarSPMB(form, _sid()); }
function promosikanSPMBKeSiswa(form) { return CoreSystem.promosikanSPMBKeSiswa(form, _sid()); }
function promosikanMassalSPMBKeSiswa(form) { return CoreSystem.promosikanMassalSPMBKeSiswa(form, _sid()); }
function getMasterSPMBPublic() { return CoreSystem.getMasterSPMBPublic(_sid()); }
function submitSPMBPublic(form) { return CoreSystem.submitSPMBPublic(form, _sid()); }
function cekStatusSPMBPublic(form) { return CoreSystem.cekStatusSPMBPublic(form, _sid()); }
function updateBerkasSPMBPublic(form) { return CoreSystem.updateBerkasSPMBPublic(form, _sid()); }

// F. LAPORAN PDF & JURNAL
function getDaftarMapel() { return CoreSystem.getDaftarMapel(_sid()); }
function generateLaporanPDF(f) { return CoreSystem.generateLaporanPDF(f, _sid()); }
function generateLaporanGuruPDF(f) { return CoreSystem.generateLaporanGuruPDF(f, _sid()); }
function generateLaporanJurnalPDF(f) { return CoreSystem.generateLaporanJurnalPDF(f, _sid()); }
function generateLaporanJurnalRangePDF(filter) { return CoreSystem.generateLaporanJurnalRangePDF(filter, _sid()); }

// G. TABUNGAN (KEUANGAN)
function getRiwayatTabungan(n) { return CoreSystem.getRiwayatTabungan(n, _sid()); }
function prosesTransaksiTabungan(f) { return CoreSystem.prosesTransaksiTabungan(f, _sid()); }
function getRiwayatTabunganLengkap(nis) { return CoreSystem.getRiwayatTabunganLengkap(nis, _sid()); }
function updateRiwayatTabungan(nis, form) { return CoreSystem.updateRiwayatTabungan(nis, form, _sid()); }
function hapusRiwayatTabungan(nis, id) { return CoreSystem.hapusRiwayatTabungan(nis, id, _sid()); }

// H. PORTAL GURU & SISWA & IOT
function getKartuBaru() { return CoreSystem.getKartuBaru(_sid()); }
function getDataPortalSiswa(n) { return CoreSystem.getDataPortalSiswa(n, _sid()); }
function getPelanggaranSiswaPortal(nis) { return CoreSystem.getPelanggaranSiswaPortal(nis, _sid()); }
function getDataPortalGuru(n) { return CoreSystem.getDataPortalGuru(n, _sid()); }
function getIzinSiswaByKelasWali(k, limitHari) { return CoreSystem.getIzinSiswaByKelasWali(k, limitHari, _sid()); }
function getPelanggaranByKelasWali(k) { return CoreSystem.getPelanggaranByKelasWali(k, _sid()); }
function getDashboardWaliKelas(k) { return CoreSystem.getDashboardWaliKelas(k, _sid()); }
function kirimPengumumanKelasWali(k, pesan, pengirim) { return CoreSystem.kirimPengumumanKelasWali(k, pesan, pengirim, _sid()); }
function ajukanIzinSiswa(f) { return CoreSystem.ajukanIzinSiswa(f, _sid()); }
function ajukanIzinGuru(data) { return CoreSystem.ajukanIzinGuru(data, _sid()); }
function getDataAkademikSiswa(n) { return CoreSystem.getDataAkademikSiswa(n, _sid()); }
function getJadwalMengajarGuru(n) { return CoreSystem.getJadwalMengajarGuru(n, _sid()); }
function getMappingKelasMapel(n) { return CoreSystem.getMappingKelasMapel(n, _sid()); }
function getDataInputNilai(k, m, o) { return CoreSystem.getDataInputNilai(k, m, o || {}, _sid()); }
function simpanNilaiSemester(d) { return CoreSystem.simpanNilaiSemester(d, _sid()); }
function kirimNilaiKeWaliKelas(d) { return CoreSystem.kirimNilaiKeWaliKelas(d, _sid()); }
function batalKirimNilaiKeWaliKelas(d) { return CoreSystem.batalKirimNilaiKeWaliKelas(d || {}, _sid()); }
function getPengaturanRdmSync() { return CoreSystem.getPengaturanRdmSync(_sid()); }
function simpanPengaturanRdmSync(form) { return CoreSystem.simpanPengaturanRdmSync(form || {}, _sid()); }
function syncNilaiHarianKeRdm(form) { return CoreSystem.syncNilaiHarianKeRdm(form || {}, _sid()); }
function syncCatatanWaliKeRdm(form) { return CoreSystem.syncCatatanWaliKeRdm(form || {}, _sid()); }
function syncRekapAbsensiKeRdm(form) { return CoreSystem.syncRekapAbsensiKeRdm(form || {}, _sid()); }
function inspectRdmLock(form) { return CoreSystem.inspectRdmLock(form || {}, _sid()); }
function syncKirimNilaiGuruRdm(form) { return CoreSystem.syncKirimNilaiGuruRdm(form || {}, _sid()); }
function syncKunciNilaiKelasRdm(form) { return CoreSystem.syncKunciNilaiKelasRdm(form || {}, _sid()); }
function tarikIdSiswaRdm() {
  if (typeof CoreSystem.tarikIdSiswaRdm !== "function") {
    return {
      status: "error",
      pesan: "CoreSystem.tarikIdSiswaRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
    };
  }
  return CoreSystem.tarikIdSiswaRdm(_sid());
}
function tarikIdGuruRdm() {
  if (typeof CoreSystem.tarikIdGuruRdm !== "function") {
    return {
      status: "error",
      pesan: "CoreSystem.tarikIdGuruRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
    };
  }
  return CoreSystem.tarikIdGuruRdm(_sid());
}
function tarikMapelRdm() {
  if (typeof CoreSystem.tarikMapelRdm !== "function") {
    return {
      status: "error",
      pesan: "CoreSystem.tarikMapelRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
    };
  }
  return CoreSystem.tarikMapelRdm(_sid());
}
function tarikIdKelasRdm() {
  if (typeof CoreSystem.tarikIdKelasRdm !== "function") {
    return {
      status: "error",
      pesan: "CoreSystem.tarikIdKelasRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
    };
  }
  return CoreSystem.tarikIdKelasRdm(_sid());
}
function tarikAjarRdm() {
  if (typeof CoreSystem.tarikAjarRdm !== "function") {
    return {
      status: "error",
      pesan: "CoreSystem.tarikAjarRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
    };
  }
  return CoreSystem.tarikAjarRdm(_sid());
}
function tarikNilaiHarianRdm(filter) {
  if (typeof CoreSystem.tarikNilaiHarianRdm !== "function") {
    return {
      status: "error",
      pesan: "CoreSystem.tarikNilaiHarianRdm belum tersedia. Push/update library CoreSystem utama dulu, lalu refresh spreadsheet."
    };
  }
  return CoreSystem.tarikNilaiHarianRdm(_sid(), filter || {});
}
function TARIK_AJAR_RDM() {
  var res = tarikAjarRdm();
  var pesan = (res && res.pesan) ? res.pesan : "Tarik Ajar RDM selesai.";
  try { SpreadsheetApp.getUi().alert(pesan); } catch (e) { Logger.log(pesan); }
  return res;
}
function TARIK_NILAI_HARIAN_RDM() {
  var res = tarikNilaiHarianRdm({});
  var pesan = (res && res.pesan) ? res.pesan : "Tarik Nilai Harian RDM selesai.";
  try { SpreadsheetApp.getUi().alert(pesan); } catch (e) { Logger.log(pesan); }
  return res;
}
function getTopikNilaiMapel(k, m, o) { return CoreSystem.getTopikNilaiMapel(k, m, o || {}, _sid()); }
function simpanTopikNilaiMapel(form) { return CoreSystem.simpanTopikNilaiMapel(form, _sid()); }
function salinTopikNilaiMapelSemesterSebelumnya(form) { return CoreSystem.salinTopikNilaiMapelSemesterSebelumnya(form, _sid()); }
function getSemuaTemplateDeskripsiMapel() { return CoreSystem.getSemuaTemplateDeskripsiMapel(_sid()); }
function simpanTemplateDeskripsiMapel(form) { return CoreSystem.simpanTemplateDeskripsiMapel(form, _sid()); }
function hapusTemplateDeskripsiMapel(id) { return CoreSystem.hapusTemplateDeskripsiMapel(id, _sid()); }
function exportNilaiSemesterCSV(f) { return CoreSystem.exportNilaiSemesterCSV(f || {}, _sid()); }
function exportNilaiSumatifExcel(f) { return CoreSystem.exportNilaiSumatifExcel(f || {}, _sid()); }
function simpanJurnalMengajar(d) { return CoreSystem.simpanJurnalMengajar(d, _sid()); }
function getRiwayatJurnalGuru(n) { return CoreSystem.getRiwayatJurnalGuru(n, _sid()); }
function getSemuaMasterTahfidzGuru() { return CoreSystem.getSemuaMasterTahfidzGuru(_sid()); }
function simpanMasterTahfidzGuru(form) { return CoreSystem.simpanMasterTahfidzGuru(form || {}, _sid()); }
function hapusMasterTahfidzGuru(form) { return CoreSystem.hapusMasterTahfidzGuru(form || {}, _sid()); }
function getPengaturanTahfidz() { return CoreSystem.getPengaturanTahfidz(_sid()); }
function simpanPengaturanTahfidz(form) { return CoreSystem.simpanPengaturanTahfidz(form || {}, _sid()); }
function getSemuaPembinaTahfidzKelas() { return CoreSystem.getSemuaPembinaTahfidzKelas(_sid()); }
function simpanPembinaTahfidzKelas(form) { return CoreSystem.simpanPembinaTahfidzKelas(form || {}, _sid()); }
function hapusPembinaTahfidzKelas(form) { return CoreSystem.hapusPembinaTahfidzKelas(form || {}, _sid()); }
function getJadwalTahfidzGuru(n) { return CoreSystem.getJadwalTahfidzGuru(n, _sid()); }
function mulaiAbsensiTahfidzGuru(form) { return CoreSystem.mulaiAbsensiTahfidzGuru(form || {}, _sid()); }
function selesaiAbsensiTahfidzGuru(form) { return CoreSystem.selesaiAbsensiTahfidzGuru(form || {}, _sid()); }
function getJadwalEkskulGuru(n) { return CoreSystem.getJadwalEkskulGuru(n, _sid()); }
function loginPembinaEkskulEksternal(idPembina, pin) { return CoreSystem.loginPembinaEkskulEksternal(idPembina, pin, _sid()); }
function getJadwalEkskulPembinaEksternal(idPembina, filter) { return CoreSystem.getJadwalEkskulPembinaEksternal(idPembina, filter || {}, _sid()); }
function mulaiAbsensiEkskulEksternal(form) { return CoreSystem.mulaiAbsensiEkskulEksternal(form || {}, _sid()); }
function selesaiAbsensiEkskulEksternal(form) { return CoreSystem.selesaiAbsensiEkskulEksternal(form || {}, _sid()); }
function getRiwayatAbsensiPembinaEksternal(idPembina, filter) { return CoreSystem.getRiwayatAbsensiPembinaEksternal(idPembina, filter || {}, _sid()); }
function getRingkasanHonorPembinaEksternal(idPembina, filter) { return CoreSystem.getRingkasanHonorPembinaEksternal(idPembina, filter || {}, _sid()); }
function getAbsensiPesertaEkskulPembina(idPembina, form) { return CoreSystem.getAbsensiPesertaEkskulPembina(idPembina, form || {}, _sid()); }
function simpanAbsensiPesertaEkskulPembina(idPembina, form) { return CoreSystem.simpanAbsensiPesertaEkskulPembina(idPembina, form || {}, _sid()); }
function getRekapAbsensiPesertaEkskul(filter) { return CoreSystem.getRekapAbsensiPesertaEkskul(filter || {}, _sid()); }
function getJadwalEkskulEksternalAdmin(filter) { return CoreSystem.getJadwalEkskulEksternalAdmin(filter || {}, _sid()); }
function simpanAbsensiEkskulEksternalManual(form) { return CoreSystem.simpanAbsensiEkskulEksternalManual(form || {}, _sid()); }
function getRiwayatAbsensiEkskulEksternal(filter) { return CoreSystem.getRiwayatAbsensiEkskulEksternal(filter || {}, _sid()); }
function updateAbsensiEkskulEksternalManual(form) { return CoreSystem.updateAbsensiEkskulEksternalManual(form || {}, _sid()); }
function hapusAbsensiEkskulEksternalManual(id, tanggal) { return CoreSystem.hapusAbsensiEkskulEksternalManual(id, tanggal, _sid()); }
function mulaiAbsensiEkskulGuru(form) { return CoreSystem.mulaiAbsensiEkskulGuru(form, _sid()); }
function selesaiAbsensiEkskulGuru(form) { return CoreSystem.selesaiAbsensiEkskulGuru(form, _sid()); }
function terimaInputDariWeb(uid) { return CoreSystem.terimaInputDariWeb(uid, _sid()); }

// J. CBT (Computer-Based Test)
function uploadCbtImage(fileData, fileName, mimeType, authToken, authSid) { return CoreSystem.uploadCbtImage(fileData, fileName, mimeType, authToken || '', _cbtSid_(authSid)); }
function getBankSoalGuru(filter, authToken, authSid) { return CoreSystem.getBankSoalGuru(filter, authToken || '', _cbtSid_(authSid)); }
function simpanBankSoal(data, authToken, authSid) { return CoreSystem.simpanBankSoal(data, authToken || '', _cbtSid_(authSid)); }
function getStatusAiBankSoal(authToken, authSid) { return CoreSystem.getStatusAiBankSoal(authToken || '', _cbtSid_(authSid)); }
function generateBankSoalAiDariJurnal(filter, authToken, authSid) { return CoreSystem.generateBankSoalAiDariJurnal(filter || {}, authToken || '', _cbtSid_(authSid)); }
function simpanBankSoalAiDraft(items, authToken, authSid) { return CoreSystem.simpanBankSoalAiDraft(items || [], authToken || '', _cbtSid_(authSid)); }
function hapusBankSoal(id, authToken, authSid) { return CoreSystem.hapusBankSoal(id, authToken || '', _cbtSid_(authSid)); }
function getBankSoalTemplate() { return CoreSystem.getBankSoalTemplate(_sid()); }
function importBankSoalExcel(rows, nip, authToken, authSid) { return CoreSystem.importBankSoalExcel(rows, nip, authToken || '', _cbtSid_(authSid)); }
function importBankSoalExcelAdmin(rows, nipTarget, options) { return CoreSystem.importBankSoalExcelAdmin(rows, nipTarget, options || {}, _sid()); }
function getPengawasCbt(filter) { return CoreSystem.getPengawasCbt(filter || {}, _sid()); }
function simpanPengawasCbt(data) { return CoreSystem.simpanPengawasCbt(data, _sid()); }
function hapusPengawasCbt(idOrRow) { return CoreSystem.hapusPengawasCbt(idOrRow, _sid()); }
function getUjianGuru(filter, authToken, authSid) { return CoreSystem.getUjianGuru(filter, authToken || '', _cbtSid_(authSid)); }
function simpanUjian(data, authToken, authSid) { return CoreSystem.simpanUjian(data, authToken || '', _cbtSid_(authSid)); }
function hapusUjian(id, authToken, authSid) { return CoreSystem.hapusUjian(id, authToken || '', _cbtSid_(authSid)); }
function getJadwalGuruCbt(nip, authToken, authSid) { return CoreSystem.getJadwalGuruCbt(nip, authToken || '', _cbtSid_(authSid)); }
function getUjianSiswa(n, authToken, authSid) { return CoreSystem.getUjianSiswa(n, authToken || '', _cbtSid_(authSid)); }
function getRiwayatUjianSiswa(nis, authToken, authSid) { return CoreSystem.getRiwayatUjianSiswa(nis, authToken || '', _cbtSid_(authSid)); }
function mulaiUjianSiswa(nis, ujianId, token, preferredHasilId, authToken, authSid) { return CoreSystem.mulaiUjianSiswa(nis, ujianId, token, preferredHasilId || '', authToken || '', _cbtSid_(authSid)); }
function submitUjianSiswa(nis, ujianId, jawaban, authToken, authSid) { return CoreSystem.submitUjianSiswa(nis, ujianId, jawaban, authToken || '', _cbtSid_(authSid)); }
function getUjianPengawas(nip, authToken, authSid) { return CoreSystem.getUjianPengawas(nip, authToken || '', _cbtSid_(authSid)); }
function getHasilUjianGuru(ujianId, authToken, authSid) { return CoreSystem.getHasilUjianGuru(ujianId, authToken || '', _cbtSid_(authSid)); }
function getHasilUjianAdmin(filter) { return CoreSystem.getHasilUjianAdmin(filter, _sid()); }
function getReviewJawabanUjianAdmin(filter) { return CoreSystem.getReviewJawabanUjianAdmin(filter || {}, _sid()); }
function getReviewDataUjianAdmin(ujianId) { return CoreSystem.getReviewDataUjianAdmin(ujianId, _sid()); }
function getReviewDataUjianGuru(ujianId, authToken, authSid) { return CoreSystem.getReviewDataUjianGuru(ujianId, authToken || '', _cbtSid_(authSid)); }
function exportHasilUjianAdminPDF(filter) { return CoreSystem.exportHasilUjianAdminPDF(filter || {}, _sid()); }
function exportHasilUjianAdminExcel(filter) { return CoreSystem.exportHasilUjianAdminExcel(filter || {}, _sid()); }
function getDaftarUjianAdmin() { return CoreSystem.getDaftarUjianAdmin(_sid()); }
function updateDataUjianAdmin(form) { return CoreSystem.updateDataUjianAdmin(form || {}, _sid()); }
function getDaftarHadirUjian(ujianId, filterKelas) { return CoreSystem.getDaftarHadirUjian(ujianId, filterKelas, _sid()); }
function updateUjianProgress(hasilId, soalKe, tabSwitches, authToken, authSid) { return CoreSystem.updateUjianProgress(hasilId, soalKe, tabSwitches, authToken || '', _cbtSid_(authSid)); }
function simpanProgressJawaban(hasilId, jawabanJson, authToken, authSid) { return CoreSystem.simpanProgressJawaban(hasilId, jawabanJson, authToken || '', _cbtSid_(authSid)); }
function simpanProgressUjian(hasilId, soalKe, tabSwitches, jawabanJson, authToken, authSid) { return CoreSystem.simpanProgressUjian(hasilId, soalKe, tabSwitches, jawabanJson, authToken || '', _cbtSid_(authSid)); }
function getMonitorUjianGuru(ujianId, authToken, authSid) { return CoreSystem.getMonitorUjianGuru(ujianId, authToken || '', _cbtSid_(authSid)); }
function generateTokenUjian(ujianId, authToken, authSid) { return CoreSystem.generateTokenUjian(ujianId, authToken || '', _cbtSid_(authSid)); }
function generateTokenResume(ujianId, authToken, authSid) { return CoreSystem.toggleTokenAktif(ujianId, true, authToken || '', _cbtSid_(authSid)); } // Backward compat
function toggleTokenAktif(ujianId, aktif, authToken, authSid) { return CoreSystem.toggleTokenAktif(ujianId, aktif, authToken || '', _cbtSid_(authSid)); }
function studentExitUjian(hasilId, authToken, authSid, soalKe, tabSwitches) {
  if (typeof soalKe === 'undefined' && typeof tabSwitches === 'undefined') {
    return CoreSystem.studentExitUjian(hasilId, authToken || '', _cbtSid_(authSid));
  }
  return CoreSystem.studentExitUjian(hasilId, authToken || '', _cbtSid_(authSid), soalKe || 0, tabSwitches || 0);
}
function requestResumeApproval(nis, ujianId, authToken, authSid) { return CoreSystem.requestResumeApproval(nis, ujianId, authToken || '', _cbtSid_(authSid)); }
function approveResumeUjian(hasilId, approver, authToken, authSid) { return CoreSystem.approveResumeUjian(hasilId, approver, authToken || '', _cbtSid_(authSid)); }
function resetPesertaUjianKeResume(hasilId, approver, authToken, authSid) { return CoreSystem.resetPesertaUjianKeResume(hasilId, approver, authToken || '', _cbtSid_(authSid)); }
function cekPersetujuanResumeUjian(hasilId, authToken, authSid) { return CoreSystem.cekPersetujuanResumeUjian(hasilId, authToken || '', _cbtSid_(authSid)); }
function lanjutkanResumeUjian(nis, hasilId, authToken, authSid) { return CoreSystem.lanjutkanResumeUjian(nis, hasilId, authToken || '', _cbtSid_(authSid)); }

// I. TRIGGER (WAJIB ADA)
function prosesAntrianAbsensi() { return CoreSystem.prosesAntrianAbsensi(_sid()); }
function bersihkanUnknownAbsensiLama() { return CoreSystem.bersihkanUnknownAbsensiLama(_sid()); }
function hapusSemuaDitolakAbsensi() { return CoreSystem.hapusSemuaDitolakAbsensi(_sid()); }
function hapusSemuaUnknownAbsensi() { return CoreSystem.hapusSemuaUnknownAbsensi(_sid()); }
function debugUnknownAbsensiCleanup() { return CoreSystem.debugUnknownAbsensiCleanup(_sid()); }
function debugUnknownAbsensiCleanupText() { return CoreSystem.debugUnknownAbsensiCleanupText(_sid()); }
function kirimPengingatJurnalOtomatis() { return CoreSystem.kirimPengingatJurnalOtomatis(_sid()); }
function setupTriggerPengingatJurnalJam8() { return CoreSystem.setupTriggerPengingatJurnalJam8(_sid()); }
function hapusTriggerPengingatJurnalJam8() { return CoreSystem.hapusTriggerPengingatJurnalJam8(_sid()); }
function kirimPengingatJadwalPagiOtomatis() { return CoreSystem.kirimPengingatJadwalPagiOtomatis(_sid()); }
function setupTriggerPengingatJadwalPagi() { return CoreSystem.setupTriggerPengingatJadwalPagi(_sid()); }
function hapusTriggerPengingatJadwalPagi() { return CoreSystem.hapusTriggerPengingatJadwalPagi(_sid()); }

function _wrapperAbsensiSid_() {
  var sid = "";
  try { sid = (_sid() || "").toString().trim(); } catch (e) {}
  if (!sid) {
    try { sid = (PropertiesService.getScriptProperties().getProperty("WRAPPER_ABSENSI_TRIGGER_SID") || "").toString().trim(); } catch (e2) {}
  }
  return sid;
}

function prosesAntrianAbsensiOtomatis() {
  var sid = _wrapperAbsensiSid_();
  if (!sid) throw new Error("SID trigger absensi belum diset. Jalankan setup trigger dulu.");
  return CoreSystem.prosesAntrianAbsensi(sid);
}

function setupTriggerProsesAntrianAbsensi() {
  var sid = _wrapperAbsensiSid_();
  if (!sid) throw new Error("Spreadsheet tenant tidak terdeteksi.");
  PropertiesService.getScriptProperties().setProperty("WRAPPER_ABSENSI_TRIGGER_SID", sid);

  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === "prosesAntrianAbsensiOtomatis") {
      ScriptApp.deleteTrigger(all[i]);
    }
  }

  ScriptApp.newTrigger("prosesAntrianAbsensiOtomatis")
    .timeBased()
    .everyMinutes(1)
    .create();

  return {
    status: "sukses",
    pesan: "Trigger antrian absensi aktif (setiap 1 menit).",
    spreadsheetId: sid
  };
}

function statusTriggerProsesAntrianAbsensi() {
  var all = ScriptApp.getProjectTriggers();
  var jumlah = 0;
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === "prosesAntrianAbsensiOtomatis") jumlah++;
  }
  return {
    status: "sukses",
    aktif: jumlah > 0,
    jumlah: jumlah,
    spreadsheetId: _wrapperAbsensiSid_() || "-"
  };
}

function hapusTriggerProsesAntrianAbsensi() {
  var all = ScriptApp.getProjectTriggers();
  var count = 0;
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === "prosesAntrianAbsensiOtomatis") {
      ScriptApp.deleteTrigger(all[i]);
      count++;
    }
  }
  try { PropertiesService.getScriptProperties().deleteProperty("WRAPPER_ABSENSI_TRIGGER_SID"); } catch (e) {}
  return { status: "sukses", pesan: "Trigger antrian absensi dihapus: " + count };
}

// J. PENGUMUMAN
function syncPengumuman() { return CoreSystem.prosesSyncPengumuman(_sid()); }
function simpanPengumumanAdmin(form) { return CoreSystem.simpanPengumumanAdmin(form, _sid()); }
function getAllPengumumanAdmin() { return CoreSystem.getAllPengumumanAdmin(_sid()); }
function hapusPengumumanAdmin(id) { return CoreSystem.hapusPengumumanAdmin(id, _sid()); }
function getPengumumanSiswa() { return CoreSystem.getPengumumanSiswa(_sid()); }
function getPengumumanSiswaFilter(nis) { return CoreSystem.getPengumumanSiswaFilter(nis, _sid()); }
function getPengumumanGuruFilter() { return CoreSystem.getPengumumanGuruFilter(_sid()); }
function getSemuaAgendaRapat(filter) {
  // Samakan bentuk response dengan versi admin agar aman untuk frontend lama/baru.
  return getSemuaAgendaRapatAdmin(filter || {});
}
function _fmtTanggalAgendaWrap_(v) {
  if (!v && v !== 0) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, "Asia/Jakarta", "dd-MM-yyyy");
  }
  var s = v.toString().trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s.substring(8, 10) + "-" + s.substring(5, 7) + "-" + s.substring(0, 4);
  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) return s;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, "Asia/Jakarta", "dd-MM-yyyy");
  return s;
}
function _fmtJamAgendaWrap_(v) {
  if (!v && v !== 0) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, "Asia/Jakarta", "HH:mm");
  }
  var s = v.toString().trim();
  if (!s) return "";
  var m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return ("0" + Number(m[1])).slice(-2) + ":" + m[2];
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, "Asia/Jakarta", "HH:mm");
  return s;
}
function _ensureSheetAgendaWrap_() {
  var sid = _sidAgendaSafe_();
  if (!sid) throw new Error("SID tenant tidak ditemukan (Agenda Rapat).");
  var ss = SpreadsheetApp.openById(sid);
  var sh = ss.getSheetByName("AgendaRapat");
  if (!sh) sh = ss.insertSheet("AgendaRapat");
  var header = ["ID", "Tanggal", "Jam_Mulai", "Jam_Selesai", "Judul", "Lokasi", "Agenda", "Catatan", "Status", "Created_At", "Updated_At"];
  var cur = sh.getRange(1, 1, 1, header.length).getValues()[0];
  var need = false;
  for (var i = 0; i < header.length; i++) {
    if ((cur[i] || "").toString().trim() !== header[i]) { need = true; break; }
  }
  if (need) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function getSemuaAgendaRapatAdmin(filter) {
  try {
    var sid = _sidAgendaSafe_();
    if (!sid) return { status: "error", pesan: "SID tenant tidak ditemukan.", data: [], last_row: 0 };
    // Prioritas baca dari CoreSystem agar konsisten dengan renderApp library.
    try {
      var coreRes = CoreSystem.getSemuaAgendaRapatAdmin(filter || {}, sid);
      if (coreRes && typeof coreRes === "object") {
        if (!coreRes.spreadsheet_id) coreRes.spreadsheet_id = sid;
        return coreRes;
      }
    } catch (eCore) {}

    // Fallback lokal (jika Core gagal).
    var ss = SpreadsheetApp.openById(sid);
    var sh = ss.getSheetByName("AgendaRapat");
    if (!sh) return { status: "sukses", spreadsheet_id: sid, sheet: "AgendaRapat", last_row: 0, data: [] };
    var lr = sh.getLastRow();
    if (lr < 2) return { status: "sukses", spreadsheet_id: sid, sheet: "AgendaRapat", last_row: lr, data: [] };
    var vals = sh.getRange(2, 1, lr - 1, Math.min(11, sh.getLastColumn())).getValues();
    var out = vals.map(function(r, i) {
      return {
        id: r[0] || "",
        tanggal: _fmtTanggalAgendaWrap_(r[1] || ""),
        jam_mulai: _fmtJamAgendaWrap_(r[2] || ""),
        jam_selesai: _fmtJamAgendaWrap_(r[3] || ""),
        judul: r[4] || "",
        lokasi: r[5] || "",
        agenda: r[6] || "",
        catatan: r[7] || "",
        status: (r[8] || "AKTIF").toString().trim().toUpperCase(),
        created_at: r[9] || "",
        updated_at: r[10] || "",
        row_sheet: i + 2
      };
    });
    return { status: "sukses", spreadsheet_id: sid, sheet: "AgendaRapat", last_row: lr, data: out };
  } catch (e) {
    return { status: "error", pesan: e.toString(), data: [], last_row: 0 };
  }
}
function getAgendaRapatLocal(filter) {
  return getSemuaAgendaRapatAdmin(filter || {});
}
function simpanAgendaRapat(form) {
  var sid = _sidAgendaSafe_();
  if (!sid) return { status: "error", pesan: "SID tenant tidak ditemukan saat simpan agenda." };
  var res = CoreSystem.simpanAgendaRapat(form, sid);
  try {
    var f = form || {};
    if (!res || res.status !== "sukses") return res;
    var sh = _ensureSheetAgendaWrap_();
    var id = (res.id || f.id || "").toString().trim();
    if (!id) return res;
    var nowText = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    var lr = sh.getLastRow();
    var row = 0;
    if (lr >= 2) {
      var ids = sh.getRange(2, 1, lr - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if ((ids[i][0] || "").toString().trim() === id) { row = i + 2; break; }
      }
    }
    var tgl = (f.tanggal || "").toString().trim();
    var jm = (f.jam_mulai || "").toString().trim();
    var js = (f.jam_selesai || jm).toString().trim();
    var judul = (f.judul || "").toString().trim();
    var lokasi = (f.lokasi || "").toString().trim();
    var agenda = (f.agenda || "").toString().trim();
    var catatan = (f.catatan || "-").toString().trim();
    var status = (f.status || "AKTIF").toString().trim().toUpperCase();
    if (status !== "AKTIF" && status !== "NONAKTIF") status = "AKTIF";
    if (row >= 2) {
      var created = sh.getRange(row, 10).getValue() || nowText;
      sh.getRange(row, 1, 1, 11).setValues([[id, tgl, jm, js, judul, lokasi, agenda, catatan, status, created, nowText]]);
    } else {
      sh.appendRow([id, tgl, jm, js, judul, lokasi, agenda, catatan, status, nowText, nowText]);
    }
    res.last_row = sh.getLastRow();
    res.spreadsheet_id = sid || _sidAgendaSafe_() || "-";
  } catch (eSync) {}
  return res;
}
function hapusAgendaRapat(idOrRow) {
  var res = CoreSystem.hapusAgendaRapat(idOrRow, _sid());
  try {
    var sh = _ensureSheetAgendaWrap_();
    var lr = sh.getLastRow();
    if (lr < 2) return res;
    var row = Number(idOrRow || 0);
    if (!(row >= 2 && row <= lr)) {
      var id = (idOrRow || "").toString().trim();
      var ids = sh.getRange(2, 1, lr - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if ((ids[i][0] || "").toString().trim() === id) { row = i + 2; break; }
      }
    }
    if (row >= 2 && row <= sh.getLastRow()) sh.deleteRow(row);
  } catch (eDel) {}
  return res;
}
function kirimPengingatAgendaRapatHariH() { return CoreSystem.kirimPengingatAgendaRapatHariH(_sid()); }
function setupTriggerPengingatAgendaRapatPagi() { return CoreSystem.setupTriggerPengingatAgendaRapatPagi(_sid()); }
function hapusTriggerPengingatAgendaRapatPagi() { return CoreSystem.hapusTriggerPengingatAgendaRapatPagi(_sid()); }
function syncAgendaRapatSheetToFirebase() { return CoreSystem.syncAgendaRapatSheetToFirebase(_sid()); }
function getAgendaRapatGuruHariIni(nip) { return CoreSystem.getAgendaRapatGuruHariIni(nip, _sid()); }
function simpanKehadiranAgendaRapat(form) { return CoreSystem.simpanKehadiranAgendaRapat(form, _sid()); }
function getDaftarHadirAgendaRapat(agendaId) { return CoreSystem.getDaftarHadirAgendaRapat(agendaId, _sid()); }
function getDokumentasiAgendaRapat(agendaId) { return CoreSystem.getDokumentasiAgendaRapat(agendaId, _sid()); }
function simpanDokumentasiAgendaRapat(form) { return CoreSystem.simpanDokumentasiAgendaRapat(form, _sid()); }
function generateBeritaAcaraAgendaRapatPDF(agendaId) { return CoreSystem.generateBeritaAcaraAgendaRapatPDF(agendaId, _sid()); }
function generateNotulenAgendaRapatPDF(agendaId) { return CoreSystem.generateNotulenAgendaRapatPDF(agendaId, _sid()); }
function generateDaftarHadirAgendaRapatPDF(agendaId) { return CoreSystem.generateDaftarHadirAgendaRapatPDF(agendaId, _sid()); }

// Buku Tamu (Public Form)
function getDaftarTujuanBukuTamu() { return CoreSystem.getDaftarTujuanBukuTamu(_sid()); }
function getDataBukuTamuAdmin() {
  try {
    var sid = "";
    try { sid = _sid(); } catch (eSid) {}

    var ss = sid ? SpreadsheetApp.openById(sid) : null;
    if (!ss) throw new Error("Spreadsheet tidak terdeteksi.");

    var sh = ss.getSheetByName("BukuTamu") || ss.getSheetByName("Buku Tamu") || ss.getSheetByName("DataTamu") || ss.getSheetByName("Buku_Tamu");
    if (!sh) {
      var all = ss.getSheets();
      for (var i = 0; i < all.length; i++) {
        var nm = (all[i].getName() || "").toLowerCase();
        if (nm.indexOf("tamu") !== -1) { sh = all[i]; break; }
      }
    }
    if (!sh) return [];

    var lr = sh.getLastRow();
    if (lr < 2) return [];
    var lc = Math.max(sh.getLastColumn(), 17);
    var vals = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();

    function c(r, idx) { return idx < r.length ? r[idx] : ""; }

    function normalizeRow(r) {
      // Fallback: jika data tersimpan dalam 1 sel dengan pemisah TAB
      if ((c(r, 1) || "") === "" && (c(r, 0) || "").indexOf("	") !== -1) {
        var parts = (c(r, 0) || "").split("	");
        if (parts.length > 1) return parts;
      }
      return r;
    }

    var out = vals.map(function(raw) {
      var r = normalizeRow(raw);
      return {
        id: c(r, 0) || "",
        tanggal: c(r, 1) || "",
        jam_masuk: c(r, 2) || "",
        jam_keluar: c(r, 3) || "",
        nama_tamu: c(r, 4) || "",
        instansi: c(r, 5) || "",
        no_hp: c(r, 6) || "",
        keperluan: c(r, 7) || "",
        tujuan_nip: c(r, 8) || "",
        tujuan_nama: c(r, 9) || "",
        no_hp_tujuan: c(r, 10) || "",
        petugas: c(r, 11) || "",
        status: c(r, 12) || "",
        durasi_menit: Number(c(r, 13) || 0),
        catatan: c(r, 14) || "",
        created_at: c(r, 15) || "",
        updated_at: c(r, 16) || ""
      };
    }).filter(function(x) {
      return !!(x.id || x.nama_tamu || x.keperluan || x.tujuan_nama || x.tanggal || x.jam_masuk);
    });

    out.sort(function(a, b) {
      var ka = (a.tanggal || "") + " " + (a.jam_masuk || "");
      var kb = (b.tanggal || "") + " " + (b.jam_masuk || "");
      if (ka === kb) return 0;
      return ka < kb ? 1 : -1;
    });

    return out;
  } catch (e) {
    throw new Error("Buku Tamu error: " + e.toString());
  }
}
function submitBukuTamuPublic(form) { return CoreSystem.submitBukuTamuPublic(form, _sid()); }

function normalisasiBukuTamuAdmin() {
  try {
    var sid = "";
    try { sid = _sid(); } catch (eSid) {}
    var ss = sid ? SpreadsheetApp.openById(sid) : null;
    if (!ss) throw new Error("Spreadsheet tidak terdeteksi.");

    var sh = ss.getSheetByName("BukuTamu") || ss.getSheetByName("Buku Tamu") || ss.getSheetByName("DataTamu") || ss.getSheetByName("Buku_Tamu");
    if (!sh) {
      var all = ss.getSheets();
      for (var i = 0; i < all.length; i++) {
        var nm = (all[i].getName() || "").toLowerCase();
        if (nm.indexOf("tamu") !== -1) { sh = all[i]; break; }
      }
    }
    if (!sh) return { status: 'error', pesan: 'Sheet Buku Tamu tidak ditemukan.', total_diperbaiki: 0 };

    var header = [
      "ID","Tanggal","Jam_Masuk","Jam_Keluar","Nama_Tamu","Instansi","No_HP","Keperluan",
      "Tujuan_NIP","Tujuan_Nama","No_HP_Tujuan","Petugas","Status","Durasi_Menit","Catatan","Created_At","Updated_At"
    ];
    sh.getRange(1, 1, 1, header.length).setValues([header]);

    var lr = sh.getLastRow();
    if (lr < 2) return { status: 'sukses', pesan: 'Tidak ada data untuk dinormalisasi.', total_diperbaiki: 0 };

    var lc = Math.max(sh.getLastColumn(), 17);
    var vals = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();

    var fixed = 0;
    var out = vals.map(function(r) {
      var row = r.slice(0, 17);
      while (row.length < 17) row.push('');

      if ((row[1] || '') === '' && (row[0] || '').indexOf('	') !== -1) {
        var parts = (row[0] || '').split('	');
        if (parts.length > 1) {
          var norm = parts.slice(0, 17);
          while (norm.length < 17) norm.push('');
          row = norm;
          fixed++;
        }
      }
      return row;
    }).filter(function(r) {
      return (r[0] || r[4] || r[7] || r[9]);
    });

    sh.getRange(2, 1, Math.max(lr - 1, 1), 17).clearContent();
    if (out.length) sh.getRange(2, 1, out.length, 17).setValues(out);

    return { status: 'sukses', pesan: 'Normalisasi selesai.', total_diperbaiki: fixed, total_data: out.length };
  } catch (e) {
    return { status: 'error', pesan: 'Normalisasi gagal: ' + e.toString(), total_diperbaiki: 0 };
  }
}

function getDebugBukuTamuAdmin() {
  var info = {
    sid: "",
    spreadsheet_name: "",
    sheet_candidates: [],
    last_row: 0,
    error: ""
  };
  try { info.sid = _sid(); } catch (eSid) {}
  try {
    var ss = info.sid ? SpreadsheetApp.openById(info.sid) : null;
    if (!ss) throw new Error("Spreadsheet tidak terdeteksi");
    info.spreadsheet_name = ss.getName();
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      var nm = (all[i].getName() || "");
      if (nm.toLowerCase().indexOf("tamu") !== -1) info.sheet_candidates.push(nm);
    }
    var sh = ss.getSheetByName("BukuTamu") || ss.getSheetByName("Buku Tamu") || ss.getSheetByName("DataTamu") || ss.getSheetByName("Buku_Tamu");
    if (!sh && info.sheet_candidates.length) sh = ss.getSheetByName(info.sheet_candidates[0]);
    if (sh) info.last_row = sh.getLastRow();
  } catch (e) {
    info.error = String(e);
  }
  return info;
}

// K. MATA PELAJARAN
function getSemuaMapelMaster() { return CoreSystem.getSemuaMapelMaster(_sid()); }
function simpanMapelMaster(form) { return CoreSystem.simpanMapelMaster(form, _sid()); }
function hapusMapelMaster(kode) { return CoreSystem.hapusMapelMaster(kode, _sid()); }
function getSemuaPembinaEkskulEksternal() { return CoreSystem.getSemuaPembinaEkskulEksternal(_sid()); }
function simpanPembinaEkskulEksternal(form) { return CoreSystem.simpanPembinaEkskulEksternal(form, _sid()); }
function hapusPembinaEkskulEksternal(id) { return CoreSystem.hapusPembinaEkskulEksternal(id, _sid()); }
function getSemuaMasterEkskul() { return CoreSystem.getSemuaMasterEkskul(_sid()); }
function getPesertaEkskul(kode) { return CoreSystem.getPesertaEkskul(kode, _sid()); }
function simpanPesertaEkskul(form) { return CoreSystem.simpanPesertaEkskul(form, _sid()); }
function hapusPesertaEkskulItems(form) { return CoreSystem.hapusPesertaEkskulItems(form, _sid()); }
function simpanMasterEkskul(form) { return CoreSystem.simpanMasterEkskul(form, _sid()); }
function hapusMasterEkskul(kode) { return CoreSystem.hapusMasterEkskul(kode, _sid()); }

// L. JADWAL MENGAJAR
function getMasterDataJadwalMengajar() { return CoreSystem.getMasterDataJadwalMengajar(_sid()); }
function getSemuaJadwalMengajar() { return CoreSystem.getSemuaJadwalMengajar(_sid()); }
function simpanJadwalMengajar(form) { return CoreSystem.simpanJadwalMengajar(form, _sid()); }
function simpanJadwalMengajarBatch(items) { return CoreSystem.simpanJadwalMengajarBatch(items, _sid()); }
function hapusJadwalMengajar(row) { return CoreSystem.hapusJadwalMengajar(row, _sid()); }

// L2. JADWAL GURU PIKET
function getMasterDataGuruPiket() { return CoreSystem.getMasterDataGuruPiket(_sid()); }
function getSemuaJadwalGuruPiket() { return CoreSystem.getSemuaJadwalGuruPiket(_sid()); }
function simpanJadwalGuruPiket(form) { return CoreSystem.simpanJadwalGuruPiket(form, _sid()); }
function hapusJadwalGuruPiket(row) { return CoreSystem.hapusJadwalGuruPiket(row, _sid()); }


// M. TANGGUNGAN + PEMBUKUAN + PENGGAJIAN
function getSemuaTanggunganSiswaMaster() { return CoreSystem.getSemuaTanggunganSiswaMaster(_sid()); }
function simpanTanggunganSiswaMaster(form) { return CoreSystem.simpanTanggunganSiswaMaster(form, _sid()); }
function hapusTanggunganSiswaMaster(row) { return CoreSystem.hapusTanggunganSiswaMaster(row, _sid()); }
function getTanggunganSiswaByNis(nis) { return CoreSystem.getTanggunganSiswaByNis(nis, _sid()); }
function getTagihanTanggunganSiswa(nis) { return CoreSystem.getTagihanTanggunganSiswa(nis, _sid()); }
function bayarTanggunganSiswa(form) { return CoreSystem.bayarTanggunganSiswa(form, _sid()); }
function getSemuaSPPInfaqMaster() { return CoreSystem.getSemuaSPPInfaqMaster(_sid()); }
function simpanSPPInfaqMaster(form) { return CoreSystem.simpanSPPInfaqMaster(form, _sid()); }
function hapusSPPInfaqMaster(row) { return CoreSystem.hapusSPPInfaqMaster(row, _sid()); }
function getTagihanSPPInfaqSiswa(nis, periode) { return CoreSystem.getTagihanSPPInfaqSiswa(nis, periode, _sid()); }
function getDaftarTagihanSPPInfaqSiswaMulti(nis, periode) { return CoreSystem.getDaftarTagihanSPPInfaqSiswaMulti(nis, periode, _sid()); }
function getDaftarTagihanSPPInfaqByJumlahBulan(nis, periode, jumlahBulan) { return CoreSystem.getDaftarTagihanSPPInfaqByJumlahBulan(nis, periode, jumlahBulan, _sid()); }
function bayarSPPInfaqSiswa(form) { return CoreSystem.bayarSPPInfaqSiswa(form, _sid()); }
function bayarSPPInfaqMulti(form) { return CoreSystem.bayarSPPInfaqMulti(form, _sid()); }
function kirimPengingatSPPInfaqBelumLunas(periode) { return CoreSystem.kirimPengingatSPPInfaqBelumLunas(periode, _sid()); }
function getRingkasanSPPInfaq(periode) { return CoreSystem.getRingkasanSPPInfaq(periode, _sid()); }
function getLaporanSPPInfaq(filter) { return CoreSystem.getLaporanSPPInfaq(filter, _sid()); }
function generateLaporanSPPInfaqPDF(filter) { return CoreSystem.generateLaporanSPPInfaqPDF(filter, _sid()); }
function setupTriggerPengingatSPPInfaqPagi() { return CoreSystem.setupTriggerPengingatSPPInfaqPagi(_sid()); }
function hapusTriggerPengingatSPPInfaqPagi() { return CoreSystem.hapusTriggerPengingatSPPInfaqPagi(_sid()); }

function initPembukuanSekolah() { return CoreSystem.initPembukuanSekolah(_sid()); }
function syncPembukuanOtomatis() { return CoreSystem.syncPembukuanOtomatis(_sid()); }
function getDaftarAkunPembukuan() { return CoreSystem.getDaftarAkunPembukuan(_sid()); }
function simpanTransaksiManualPembukuan(form) { return CoreSystem.simpanTransaksiManualPembukuan(form, _sid()); }
function getTransaksiPembukuanByRow(rowSheet) { return CoreSystem.getTransaksiPembukuanByRow(rowSheet, _sid()); }
function updateTransaksiPembukuanByRow(form) { return CoreSystem.updateTransaksiPembukuanByRow(form, _sid()); }
function hapusTransaksiPembukuanByRow(rowSheet) { return CoreSystem.hapusTransaksiPembukuanByRow(rowSheet, _sid()); }
function getLaporanPembukuan(filter) { return CoreSystem.getLaporanPembukuan(filter, _sid()); }
function generateLaporanPembukuanPDF(filter) { return CoreSystem.generateLaporanPembukuanPDF(filter, _sid()); }

function syncMasterPenggajianDariDataGuru() { return CoreSystem.syncMasterPenggajianDariDataGuru(_sid()); }
function getMasterPenggajianGuru() { return CoreSystem.getMasterPenggajianGuru(_sid()); }
function simpanMasterPenggajianGuru(form) { return CoreSystem.simpanMasterPenggajianGuru(form, _sid()); }
function getPengaturanAcaraGaji() { return CoreSystem.getPengaturanAcaraGaji(_sid()); }
function simpanPengaturanAcaraGaji(form) { return CoreSystem.simpanPengaturanAcaraGaji(form, _sid()); }
function getPengaturanDendaDhuha() { return CoreSystem.getPengaturanDendaDhuha(_sid()); }
function simpanPengaturanDendaDhuha(form) { return CoreSystem.simpanPengaturanDendaDhuha(form, _sid()); }
function getPengaturanAbsenDhuha() { return CoreSystem.getPengaturanAbsenDhuha(_sid()); }
function simpanPengaturanAbsenDhuha(form) { return CoreSystem.simpanPengaturanAbsenDhuha(form, _sid()); }

// ESP32 Management
function getDaftarESP32() { return CoreSystem.getDaftarESP32(_sid()); }
function getStatusESP32(deviceId) { return CoreSystem.getStatusESP32(deviceId, _sid()); }
function simpanKonfigurasiESP32(deviceId, form) { return CoreSystem.simpanKonfigurasiESP32(deviceId, form, _sid()); }
function hapusESP32(deviceId) { return CoreSystem.hapusESP32(deviceId, _sid()); }
function autoDeteksiESP32() { return CoreSystem.autoDeteksiESP32(_sid()); }

function getPengaturanFiturDhuhaGuru() { return CoreSystem.getPengaturanFiturDhuhaGuru(_sid()); }
function simpanPengaturanFiturDhuhaGuru(form) { return CoreSystem.simpanPengaturanFiturDhuhaGuru(form, _sid()); }
function getPengaturanLabelDhuhaGuru() { return CoreSystem.getPengaturanLabelDhuhaGuru(_sid()); }
function simpanPengaturanLabelDhuhaGuru(form) { return CoreSystem.simpanPengaturanLabelDhuhaGuru(form, _sid()); }
function getPengaturanLabelIuranBulanan() { return CoreSystem.getPengaturanLabelIuranBulanan(_sid()); }
function simpanPengaturanLabelIuranBulanan(form) { return CoreSystem.simpanPengaturanLabelIuranBulanan(form, _sid()); }
function generateRekapHonorPembinaEkskul(filter) { return CoreSystem.generateRekapHonorPembinaEkskul(filter, _sid()); }
function getRekapHonorPembinaEkskul(filter) { return CoreSystem.getRekapHonorPembinaEkskul(filter, _sid()); }
function getDetailHonorPembinaEkskul(filter, pembinaId) { return CoreSystem.getDetailHonorPembinaEkskul(filter, pembinaId, _sid()); }
function bayarHonorPembinaEkskul(id, form) { return CoreSystem.bayarHonorPembinaEkskul(id, form, _sid()); }
function generateRekapGajiGuru(filter) { return CoreSystem.generateRekapGajiGuru(filter, _sid()); }
function getRekapGajiGuru(filter) { return CoreSystem.getRekapGajiGuru(filter, _sid()); }
function generateSlipGajiGuruPDF(filter) { return CoreSystem.generateSlipGajiGuruPDF(filter, _sid()); }
function generateRekapSesiNgajarKelasExcel(filter) { return CoreSystem.generateRekapSesiNgajarKelasExcel(filter, _sid()); }
function generateSuratAktifSiswaPDF(nis) { return CoreSystem.generateSuratAktifSiswaPDF(nis, _sid()); }
function bayarGajiGuru(id, form) { return CoreSystem.bayarGajiGuru(id, form, _sid()); }
function bayarSemuaGajiGuru(filter, form) { return CoreSystem.bayarSemuaGajiGuru(filter, form, _sid()); }

// N. API SCAN SECURITY
function setupApiScanKeamanan() { return CoreSystem.setupApiScanKeamanan(_sid(), 40); }
function simpanApiScanKeamanan(token, allowedSources) { return CoreSystem.simpanApiScanKeamanan(token, allowedSources, _sid()); }
function getStatusKeamananApiScan() { return CoreSystem.getStatusKeamananApiScan(_sid()); }

// O. AUTH / DIAGNOSA
function CEK_DASH_SIG() { return CoreSystem.getDashboardSignature(); }
function authorizeDokumenScope() {
  var doc = DocumentApp.create('AUTH_SCOPE_WRAPPER_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));
  var id = doc.getId();
  DriveApp.getFileById(id).setTrashed(true);
  return { status: 'sukses', pesan: 'Izin documents aktif di wrapper.', docId: id };
}
function FORCE_AUTH_SCRIPTAPP() {
  return ScriptApp.getProjectTriggers().length;
}

function CEK_ISI_SHEET_KLIEN() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pengaturan");
  if (!sheet) {
    Logger.log("Sheet Pengaturan tidak ditemukan.");
    return;
  }
  var data = sheet.getDataRange().getValues();
  Logger.log("=== CEK SHEET PENGATURAN ===");
  for (var i = 0; i < data.length; i++) {
    var k = (data[i][0] || "").toString().trim();
    if (k === "FIREBASE_URL" || k === "FIREBASE_SECRET" || k === "NAMA_SEKOLAH") {
      Logger.log("Baris " + (i + 1) + ": " + k + " = " + data[i][1]);
    }
  }
}

// Diagnosa cepat backend (opsional)
function CEK_HEALTH() {
  var out = {};
  try { out.sid = _sid(); } catch (e) { out.sid = "ERR: " + e; }
  try { out.dashboard = getDataDashboardSekolah(); } catch (e2) { out.dashboard = "ERR: " + e2; }
  try {
    var siswa = getSemuaSiswa();
    out.jumlahSiswa = (siswa && siswa.length) ? siswa.length : 0;
  } catch (e3) {
    out.jumlahSiswa = "ERR: " + e3;
  }
  Logger.log(JSON.stringify(out));
  return out;
}
function CEK_LIB() {
  Logger.log(JSON.stringify(CoreSystem.getBuildInfo()));
}
function FORCE_AUTH_BASIC() {
  return CoreSystem.FORCE_AUTH_BASIC();
}
function FORCE_AUTH_ALL() {
  return CoreSystem.FORCE_AUTH_ALL();
}

// Bridge khusus Portal Orang Tua agar google.script.run stabil
function apiPortalOrtu(method, payload, spreadsheetId) {
  return CoreSystem.apiPortalOrtu(method, payload || {}, spreadsheetId || _sid());
}
function getPelanggaranPortalOrtu(noOrtu, nis, spreadsheetId) {
  return CoreSystem.getPelanggaranPortalOrtu(noOrtu, nis, spreadsheetId || _sid());
}
function getKeuanganPortalOrtu(noOrtu, nis, periode, spreadsheetId) {
  return CoreSystem.getKeuanganPortalOrtu(noOrtu, nis, periode || "", spreadsheetId || _sid());
}

// ==================================================
//   AUTO-UPDATE WRAPPER
//   Fetch Kode.js terbaru dari GitHub & update script ini
// ==================================================

/**
 * Update wrapper ke versi terbaru dari GitHub.
 * Jalankan fungsi ini dari Apps Script editor (Run > updateWrapper).
 */
function updateWrapper() {
  var GITHUB_RAW = "https://raw.githubusercontent.com/misterjohn46/Smart-School-Wrapper/main/Kode.js";
  var scriptId = ScriptApp.getScriptId();
  
  try {
    // 1. Fetch kode terbaru
    var response = UrlFetchApp.fetch(GITHUB_RAW, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw new Error("Gagal fetch Kode.js: HTTP " + response.getResponseCode());
    }
    var newContent = response.getContentText();
    
    // 2. Ambil daftar file yang ada di project ini
    var url = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";
    var token = ScriptApp.getOAuthToken();
    
    var currentFiles = JSON.parse(UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    }).getContentText());
    
    // 3. Cari file Code.gs (atau Kode.gs)
    var sourceFileName = null;
    (currentFiles.files || []).forEach(function(f) {
      if (f.name && (f.name === "Code" || f.name.toLowerCase().indexOf("kode") !== -1)) {
        sourceFileName = f.name;
      }
    });
    if (!sourceFileName) sourceFileName = "Code";
    
    // 4. Update file via Apps Script API
    var updateUrl = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";
    var payload = {
      files: [{
        name: sourceFileName,
        type: "SERVER_JS",
        source: newContent
      }]
    };
    
    var result = UrlFetchApp.fetch(updateUrl, {
      method: "put",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (result.getResponseCode() === 200) {
      Logger.log("✅ updateWrapper BERHASIL! Kode.js telah diupdate ke versi terbaru.");
      Logger.log("File yang diupdate: " + sourceFileName);
      Logger.log("Restart / refresh Apps Script editor untuk melihat perubahan.");
    } else {
      Logger.log("❌ Gagal update: " + result.getContentText());
    }
    
  } catch (e) {
    Logger.log("❌ Error updateWrapper: " + e.toString());
    Logger.log("Pastikan scope script.projects sudah diaktifkan di appsscript.json");
  }
}

/**
 * Cek versi wrapper saat ini vs yang tersedia di GitHub.
 * Hanya logging perbandingan, tidak mengubah apa pun.
 */
function checkWrapperVersion() {
  var GITHUB_RAW = "https://raw.githubusercontent.com/misterjohn46/Smart-School-Wrapper/main/Kode.js";
  
  try {
    var response = UrlFetchApp.fetch(GITHUB_RAW, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      var remoteContent = response.getContentText();
      var remoteLines = remoteContent.split("\n").length;
      
      // Ambil versi lokal
      var scriptId = ScriptApp.getScriptId();
      var token = ScriptApp.getOAuthToken();
      var url = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";
      var currentFiles = JSON.parse(UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + token },
        muteHttpExceptions: true
      }).getContentText());
      
      var localLines = 0;
      (currentFiles.files || []).forEach(function(f) {
        if (f.source) localLines += f.source.split("\n").length;
      });
      
      Logger.log("📊 Perbandingan Wrapper:");
      Logger.log("  Lokal: ~" + localLines + " baris");
      Logger.log("  Remote (GitHub): ~" + remoteLines + " baris");
      
      if (Math.abs(localLines - remoteLines) > 5) {
        Logger.log("  ⚠️ Versi berbeda! Jalankan updateWrapper() untuk update.");
      } else {
        Logger.log("  ✅ Wrapper sudah up-to-date.");
      }
    }
  } catch (e) {
    Logger.log("❌ Gagal cek versi: " + e.toString());
  }
}
