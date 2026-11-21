/**
 * ============================================================================
 * BACKEND SINKRONISASI PORTAL ASAS GANJIL 2026/2027 · MTsN 3 SUBANG
 * ============================================================================
 * Cara Deploy:
 * 1. Buka https://script.google.com -> Proyek Baru -> Tempel seluruh kode ini.
 * 2. Klik "Terapkan" (Deploy) -> "Deployment baru" (New deployment).
 * 3. Pilih jenis: "Aplikasi web" (Web app).
 * 4. Jalankan sebagai: "Saya" (Me).
 * 5. Yang memiliki akses: "Siapa saja" (Anyone).
 * 6. Salin URL Web App dan pastikan sama dengan konstanta GAS_URL di index.html.
 */

const PROP_SCHEDULE_KEY = "ASAS_2026_SCHEDULE_JSON";
const PROP_PIN_KEY = "ASAS_2026_PROCTOR_PIN";
const DEFAULT_PROCTOR_PIN = "123456";

/**
 * Jalankan fungsi ini SEKALI di editor Google Apps Script (klik tombol Run)
 * untuk menginisialisasi database server & membuat sheet otomatis jika terhubung ke Spreadsheet.
 */
function setupServer() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_PIN_KEY)) {
    props.setProperty(PROP_PIN_KEY, DEFAULT_PROCTOR_PIN);
  }
  syncScheduleToSheet(
    props.getProperty(PROP_SCHEDULE_KEY) ? JSON.parse(props.getProperty(PROP_SCHEDULE_KEY)) : []
  );
  logActivityToSheet("SYSTEM", "Inisialisasi Server ASAS Berhasil");
  Logger.log("Server ASAS siap digunakan. PIN Proktor aktif: " + props.getProperty(PROP_PIN_KEY));
}

/**
 * Mengambil data jadwal & PIN saat halaman portal dibuka (HTTP GET)
 */
function doGet(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const rawSchedule = props.getProperty(PROP_SCHEDULE_KEY);
    const savedPin = props.getProperty(PROP_PIN_KEY) || DEFAULT_PROCTOR_PIN;

    const schedule = rawSchedule ? JSON.parse(rawSchedule) : null;

    return createJsonResponse({
      ok: true,
      app: "Server Portal ASAS Ganjil 2026/2027 - MTsN 3 Subang",
      schedule: schedule,
      proctorPin: savedPin,
      serverTime: Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss 'WIB'")
    });
  } catch (err) {
    return createJsonResponse({
      ok: false,
      error: "Gagal membaca data server: " + err.message
    });
  }
}

/**
 * Menyimpan perubahan jadwal atau pembaruan PIN dari Proktor (HTTP POST)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ ok: false, error: "Payload kosong." });
    }

    const data = JSON.parse(e.postData.contents);
    const props = PropertiesService.getScriptProperties();
    const currentPin = props.getProperty(PROP_PIN_KEY) || DEFAULT_PROCTOR_PIN;

    // Validasi PIN keamanan untuk aksi penyimpanan jadwal
    if (data.action === "saveSchedule") {
      if (String(data.pin || "") !== String(currentPin)) {
        return createJsonResponse({
          ok: false,
          error: "PIN Proktor tidak cocok dengan server."
        });
      }

      if (!Array.isArray(data.schedule)) {
        return createJsonResponse({
          ok: false,
          error: "Format data jadwal tidak valid."
        });
      }

      props.setProperty(PROP_SCHEDULE_KEY, JSON.stringify(data.schedule));
      syncScheduleToSheet(data.schedule);
      logActivityToSheet(data.proctorId || "proktor", "Simpan Jadwal");

      return createJsonResponse({
        ok: true,
        message: "Jadwal berhasil disimpan permanen di server Google Apps Script."
      });
    }

    // Pembaruan PIN Proktor sekaligus menyimpan jadwal terbaru
    if (data.action === "updatePin") {
      const newPin = String(data.pin || "").trim();
      if (!/^\d{4,8}$/.test(newPin)) {
        return createJsonResponse({
          ok: false,
          error: "PIN baru harus 4 hingga 8 digit angka."
        });
      }

      props.setProperty(PROP_PIN_KEY, newPin);
      if (Array.isArray(data.schedule)) {
        props.setProperty(PROP_SCHEDULE_KEY, JSON.stringify(data.schedule));
        syncScheduleToSheet(data.schedule);
      }
      logActivityToSheet(data.proctorId || "proktor", "Ubah PIN Proktor");

      return createJsonResponse({
        ok: true,
        message: "PIN Proktor berhasil diperbarui di server."
      });
    }

    return createJsonResponse({ ok: false, error: "Aksi tidak dikenali." });
  } catch (err) {
    return createJsonResponse({
      ok: false,
      error: "Kesalahan server GAS: " + err.message
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper mengembalikan respons JSON bersih
 */
function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menyalin tabel jadwal aktif ke sheet "Jadwal_ASAS" agar mudah dipantau panitia di Google Sheets
 */
function syncScheduleToSheet(schedule) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss || !Array.isArray(schedule)) return;

    let sheet = ss.getSheetByName("Jadwal_ASAS");
    if (!sheet) {
      sheet = ss.insertSheet("Jadwal_ASAS");
    }
    sheet.clearContents();
    sheet.appendRow(["Hari", "Tanggal", "Tingkat", "Sesi", "Waktu", "Mata Pelajaran", "Token", "Tautan Soal", "Status"]);

    const rows = [];
    schedule.forEach(function(day) {
      ["7", "8", "9"].forEach(function(grade) {
        const list = (day.sessions && day.sessions[grade]) || [];
        list.forEach(function(item) {
          if (item.break) return;
          rows.push([
            day.day || "",
            day.date || "",
            "Kelas " + grade,
            item.jam || "",
            item.time || "",
            item.subject || "",
            item.token || "",
            item.link || "",
            item.hidden ? "Disembunyikan" : "Aktif"
          ]);
        });
      });
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 9).setValues(rows);
    }
  } catch (_) {
    // Abaikan jika GAS tidak terikat ke Google Sheets
  }
}

/**
 * Mencatat riwayat perubahan ke sheet "Log_Proktor"
 */
function doGet(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const rawSchedule = props.getProperty(PROP_SCHEDULE_KEY);
    const savedPin = props.getProperty(PROP_PIN_KEY) || DEFAULT_PROCTOR_PIN;

    const schedule = rawSchedule ? JSON.parse(rawSchedule) : null;

    return createJsonResponse({
      ok: true,
      schedule: schedule,
      proctorPin: savedPin,
      serverTime: new Date().toISOString()
    });
  } catch (err) {
    return createJsonResponse({
      ok: false,
      error: "Gagal membaca data server: " + err.message
    });
  }
}

/**
 * Menyimpan perubahan jadwal atau pembaruan PIN dari Proktor (HTTP POST)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ ok: false, error: "Payload kosong." });
    }

    const data = JSON.parse(e.postData.contents);
    const props = PropertiesService.getScriptProperties();
    const currentPin = props.getProperty(PROP_PIN_KEY) || DEFAULT_PROCTOR_PIN;

    // Validasi PIN keamanan untuk aksi penyimpanan jadwal
    if (data.action === "saveSchedule") {
      if (String(data.pin || "") !== String(currentPin)) {
        return createJsonResponse({
          ok: false,
          error: "PIN Proktor tidak cocok dengan server."
        });
      }

      if (!Array.isArray(data.schedule)) {
        return createJsonResponse({
          ok: false,
          error: "Format data jadwal tidak valid."
        });
      }

      props.setProperty(PROP_SCHEDULE_KEY, JSON.stringify(data.schedule));
      logActivityToSheet(data.proctorId || "proktor", "Simpan Jadwal");

      return createJsonResponse({
        ok: true,
        message: "Jadwal berhasil disimpan permanen di Google Apps Script."
      });
    }

    // Pembaruan PIN Proktor sekaligus menyimpan jadwal terbaru
    if (data.action === "updatePin") {
      const newPin = String(data.pin || "").trim();
      if (!/^\d{4,8}$/.test(newPin)) {
        return createJsonResponse({
          ok: false,
          error: "PIN baru harus 4 hingga 8 digit angka."
        });
      }

      props.setProperty(PROP_PIN_KEY, newPin);
      if (Array.isArray(data.schedule)) {
        props.setProperty(PROP_SCHEDULE_KEY, JSON.stringify(data.schedule));
      }
      logActivityToSheet(data.proctorId || "proktor", "Ubah PIN Proktor");

      return createJsonResponse({
        ok: true,
        message: "PIN Proktor berhasil diperbarui di server."
      });
    }

    return createJsonResponse({ ok: false, error: "Aksi tidak dikenali." });
  } catch (err) {
    return createJsonResponse({
      ok: false,
      error: "Kesalahan server GAS: " + err.message
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper mengembalikan respons JSON bersih
 */
function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Opsional: Mencatat riwayat perubahan ke sheet "Log_Proktor" jika GAS terikat ke Google Sheets
 */
function logActivityToSheet(proctorId, activity) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    let sheet = ss.getSheetByName("Log_Proktor");
    if (!sheet) {
      sheet = ss.insertSheet("Log_Proktor");
      sheet.appendRow(["Waktu (WIB)", "ID Proktor", "Aktivitas"]);
    }
    const wibTime = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([wibTime, proctorId, activity]);
  } catch (_) {
    // Abaikan jika proyek GAS berdiri sendiri (standalone) tanpa Spreadsheet
  }
}
