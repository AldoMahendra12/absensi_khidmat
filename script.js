/* ══════════════════════════════════════════════════════════════
   script.js — Absensi Wali Santri v2.0
   Vanilla JavaScript — No Framework
   Sections:
     1.  CONFIG
     2.  STATE
     3.  NAVIGATION (SPA Router)
     4.  PERSISTENCE (localStorage)
     5.  EXCEL IMPORT (SheetJS)
     6.  QR SCANNER (html5-qrcode)
     7.  SCAN PROCESSOR
     8.  STATISTICS
     9.  RESULT DISPLAY
     10. GOOGLE SHEETS INTEGRATION
     11. ATTENDANCE LOG
     12. ADMIN TABLE (render, search, filter)
     13. EXPORT (Excel, CSV, Print)
     14. QR GENERATOR (generate, ZIP, Print A4)
     15. AUDIO & HAPTIC
     16. TOAST NOTIFICATIONS
     17. CONNECTION STATUS
     18. UTILITIES
     19. INIT
══════════════════════════════════════════════════════════════ */

/* ════════════════════════════════════════════════════════
   1. CONFIG
   ══ Ganti WEB_APP_URL dengan URL Google Apps Script Anda ══
════════════════════════════════════════════════════════ */
const CONFIG = {
  /** URL Google Apps Script Web App — ganti dengan milik Anda */
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbxPG4vgTo5868eR5mPS5S_S10vkY3h8xWzhErz-y-TshGz0y4Ur6SkM9pgE9GvrDRDW/exec',

  /** localStorage keys */
  LS_DB:         'absn_v2_database',
  LS_ATTENDANCE: 'absn_v2_attendance',

  /** Debounce antar scan QR yang sama (ms) */
  SCAN_DEBOUNCE_MS: 3000,

  /** FPS kamera scanner */
  SCANNER_FPS: 12,
};

/* ════════════════════════════════════════════════════════
   2. STATE
   Semua state aplikasi terpusat di sini
════════════════════════════════════════════════════════ */
const STATE = {
  /** Array semua wali santri dari Excel
   *  @type {{ id, kategori, namaSantri, namaBapak, namaIbu }[]} */
  database: [],

  /** Map kehadiran: { [id]: { ...person, waktuHadir, tanggal, timestampMs } }
   *  Digunakan sebagai Set + metadata */
  attendanceMap: {},

  /** Instance html5-qrcode */
  scanner: null,

  /** Apakah kamera sedang aktif */
  isScanning: false,

  /** ID terakhir yang di-scan (debounce) */
  lastScannedId: null,
  lastScanMs: 0,

  /** Filter admin table */
  adminFilter: { search: '', kategori: 'all', status: 'all' },

  /** Data QR yang sudah di-generate (untuk ZIP & Print) */
  generatedQRs: [],

  /** Halaman aktif */
  activePage: 'scanner',
};

/* ════════════════════════════════════════════════════════
   3. NAVIGATION — SPA Router
════════════════════════════════════════════════════════ */

/**
 * Pindah ke halaman tertentu.
 * @param {'scanner'|'admin'|'qrgen'} pageName
 */
function showPage(pageName) {
  if (STATE.activePage === pageName) return;

  /* Jika scanner aktif saat pindah halaman, matikan kamera */
  if (STATE.isScanning && pageName !== 'scanner') {
    stopScanner(true); /* silent — tanpa toast */
  }

  /* Sembunyikan semua page */
  document.querySelectorAll('.page').forEach(el => {
    el.classList.remove('page--active');
    el.hidden = true;
  });

  /* Tampilkan page target */
  const target = document.getElementById(`page-${pageName}`);
  if (target) {
    target.hidden = false;
    /* Pakai timeout minimal agar `display:flex` sudah terset sebelum animasi */
    requestAnimationFrame(() => target.classList.add('page--active'));
  }

  /* Update bottom nav */
  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.classList.remove('nav-btn--active');
    btn.setAttribute('aria-current', 'false');
  });
  const activeBtn = document.getElementById(`nav-${pageName}`);
  if (activeBtn) {
    activeBtn.classList.add('nav-btn--active');
    activeBtn.setAttribute('aria-current', 'page');
  }

  STATE.activePage = pageName;

  /* Saat masuk admin, render tabel */
  if (pageName === 'admin') renderAdminTable();

  /* Saat masuk qrgen, sinkronkan tampilan */
  if (pageName === 'qrgen') syncQRGenPage();
}

/* ════════════════════════════════════════════════════════
   4. PERSISTENCE — localStorage
════════════════════════════════════════════════════════ */

/** Simpan database & absensi ke localStorage */
function saveToStorage() {
  try {
    localStorage.setItem(CONFIG.LS_DB, JSON.stringify(STATE.database));
    localStorage.setItem(CONFIG.LS_ATTENDANCE, JSON.stringify(STATE.attendanceMap));
  } catch (e) {
    console.warn('[Storage] Gagal menyimpan:', e.message);
  }
}

/** Muat database & absensi dari localStorage */
function loadFromStorage() {
  try {
    const db  = localStorage.getItem(CONFIG.LS_DB);
    const att = localStorage.getItem(CONFIG.LS_ATTENDANCE);
    if (db)  STATE.database       = JSON.parse(db);
    if (att) STATE.attendanceMap  = JSON.parse(att);
  } catch (e) {
    console.warn('[Storage] Gagal memuat:', e.message);
    STATE.database = [];
    STATE.attendanceMap = {};
  }
}

/** Hapus seluruh database dan absensi */
function clearDatabase() {
  if (!confirm('Hapus semua data database dan absensi? Tindakan ini tidak dapat dibatalkan.')) return;

  STATE.database = [];
  STATE.attendanceMap = {};
  STATE.generatedQRs = [];
  saveToStorage();

  updateDBBanner();
  updateStats();
  renderLogList();
  renderAdminTable();
  syncQRGenPage();
  resetQRGrid();

  showToast('Database berhasil dihapus', 'info');
}

/* ════════════════════════════════════════════════════════
   5. EXCEL IMPORT — SheetJS
════════════════════════════════════════════════════════ */

/**
 * Handler upload Excel (dipanggil dari semua input file).
 * @param {File} file
 */
function handleExcelFile(file) {
  if (!file) return;

  /* Validasi ekstensi */
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) {
    showToast('Format file tidak valid. Gunakan file .xlsx atau .xls', 'error');
    return;
  }

  showToast('Membaca file Excel...', 'info');

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });

      const sheetNames = workbook.SheetNames;
      if (sheetNames.length < 1) {
        showToast('File Excel tidak memiliki sheet!', 'error');
        return;
      }

      let allData = [];

      /* Sheet pertama = Putra */
      if (sheetNames[0]) {
        const putraRows = parseExcelSheet(workbook.Sheets[sheetNames[0]], 'Putra');
        allData = [...allData, ...putraRows];
      }

      /* Sheet kedua = Putri (jika ada) */
      if (sheetNames[1]) {
        const putriRows = parseExcelSheet(workbook.Sheets[sheetNames[1]], 'Putri');
        allData = [...allData, ...putriRows];
      }

      if (allData.length === 0) {
        showToast('Tidak ada data yang terbaca dari Excel. Periksa format file.', 'error');
        return;
      }

      /* Simpan ke state */
      STATE.database = allData;
      STATE.attendanceMap = {}; /* Reset absensi saat import database baru */
      saveToStorage();

      /* Update semua UI */
      updateDBBanner();
      updateStats();
      renderLogList();
      renderAdminTable();
      syncQRGenPage();
      resetQRGrid();

      const putraCount = allData.filter(d => d.kategori === 'Putra').length;
      const putriCount = allData.filter(d => d.kategori === 'Putri').length;
      showToast(
        `✓ ${allData.length} data dimuat: ${putraCount} Putra, ${putriCount} Putri`,
        'success'
      );
    } catch (err) {
      console.error('[Excel Import]', err);
      showToast('Gagal membaca file Excel: ' + err.message, 'error');
    }
  };
  reader.onerror = () => showToast('Gagal membaca file.', 'error');
  reader.readAsArrayBuffer(file);
}

/**
 * Parse satu sheet Excel menjadi array data wali santri.
 * Setiap baris santri menghasilkan 1 atau 2 record wali:
 *   - Jika ada Nama Bapak → record dengan ID: PREFIX001B, jenisWali: 'Bapak'
 *   - Jika ada Nama Ibu   → record dengan ID: PREFIX001I, jenisWali: 'Ibu'
 *
 * Struktur kolom yang diharapkan:
 *   Kolom 0 (A): No (dilewati)
 *   Kolom 1 (B): Nama Santri / Nama Anak
 *   Kolom 2 (C): Nama Bapak
 *   Kolom 3 (D): Nama Ibu
 *
 * @param {object} worksheet   - SheetJS worksheet object
 * @param {'Putra'|'Putri'} kategori
 * @returns {object[]}
 */
function parseExcelSheet(worksheet, kategori) {
  /* sheet_to_json dengan header:1 mengembalikan array of arrays */
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    blankrows: false,
  });

  const prefix  = kategori === 'Putra' ? 'PUTRA' : 'PUTRI';
  const results = [];
  let counter   = 1;

  /* Baris pertama (index 0) adalah header — lewati */
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 2) continue;

    /* Ambil nilai kolom B (index 1) sebagai Nama Santri */
    const namaSantri = String(row[1] || '').trim();
    if (!namaSantri || namaSantri === '' || namaSantri === '-') continue;

    const namaBapak = String(row[2] || '').trim();
    const namaIbu   = String(row[3] || '').trim();
    const noSantri  = String(counter).padStart(3, '0');

    /* Setiap wali yang tercantum mendapatkan QR Code sendiri */
    const hasBapak = namaBapak && namaBapak !== '-' && namaBapak !== '';
    const hasIbu   = namaIbu   && namaIbu   !== '-' && namaIbu   !== '';

    if (hasBapak) {
      results.push({
        id:        `${prefix}${noSantri}B`,
        kategori,
        namaSantri,
        namaWali:  namaBapak,
        jenisWali: 'Bapak',
      });
    }

    if (hasIbu) {
      results.push({
        id:        `${prefix}${noSantri}I`,
        kategori,
        namaSantri,
        namaWali:  namaIbu,
        jenisWali: 'Ibu',
      });
    }

    /* Jika keduanya kosong, tetap catat dengan placeholder */
    if (!hasBapak && !hasIbu) {
      results.push({
        id:        `${prefix}${noSantri}B`,
        kategori,
        namaSantri,
        namaWali:  '-',
        jenisWali: 'Bapak',
      });
    }

    counter++;
  }

  return results;
}

/* ════════════════════════════════════════════════════════
   6. QR SCANNER — html5-qrcode
════════════════════════════════════════════════════════ */

/** Aktifkan kamera dan mulai scanning */
async function startScanner() {
  /* Periksa apakah database sudah ada */
  if (STATE.database.length === 0) {
    showToast('Upload data Excel terlebih dahulu sebelum scan!', 'warning');
    return;
  }

  if (STATE.isScanning) return;

  try {
    /* Buat instance baru jika belum ada */
    if (!STATE.scanner) {
      STATE.scanner = new Html5Qrcode('qr-reader', { verbose: false });
    }

    /* Konfigurasi scanner — qrbox adaptif terhadap ukuran viewport */
    const scannerConfig = {
      fps: CONFIG.SCANNER_FPS,
      qrbox: function(viewfinderWidth, viewfinderHeight) {
        const minEdge = Math.min(viewfinderWidth, viewfinderHeight);
        const size    = Math.floor(minEdge * 0.72);
        return { width: size, height: size };
      },
      aspectRatio: 1.0,
      disableFlip: false,
      experimentalFeatures: { useBarCodeDetectorIfSupported: true },
    };

    await STATE.scanner.start(
      { facingMode: 'environment' }, /* kamera belakang */
      scannerConfig,
      onScanSuccess,
      onScanError   /* error per-frame — diabaikan */
    );

    STATE.isScanning = true;
    setScannerUIActive(true);

    /* Sembunyikan idle overlay, tampilkan frame */
    el('scanner-idle').classList.add('hidden-idle');
    el('scan-frame').hidden = false;

  } catch (err) {
    console.error('[Scanner]', err);

    /* Pesan error yang lebih ramah untuk user */
    let msg = 'Gagal mengakses kamera.';
    if (err.name === 'NotAllowedError' || err.message?.includes('permission')) {
      msg = 'Izin kamera ditolak. Mohon izinkan akses kamera di pengaturan browser.';
    } else if (err.message?.includes('NotFoundError')) {
      msg = 'Kamera tidak ditemukan pada perangkat ini.';
    } else if (err.message) {
      msg = 'Error kamera: ' + err.message;
    }

    showToast(msg, 'error');
  }
}

/**
 * Matikan kamera.
 * @param {boolean} [silent=false] - jika true, tidak tampilkan toast
 */
async function stopScanner(silent = false) {
  if (!STATE.isScanning || !STATE.scanner) return;

  try {
    await STATE.scanner.stop();
    STATE.scanner = null;
    STATE.isScanning = false;
    STATE.lastScannedId = null;
    STATE.lastScanMs = 0;

    setScannerUIActive(false);
    el('scanner-idle').classList.remove('hidden-idle');
    el('scan-frame').hidden = true;

    if (!silent) showToast('Kamera dimatikan', 'info');
  } catch (err) {
    console.error('[Scanner stop]', err);
    STATE.isScanning = false;
    STATE.scanner = null;
  }
}

/** Reset hasil scan terakhir (bukan matikan kamera) */
function resetLastResult() {
  el('result-empty').hidden    = false;
  el('result-populated').hidden = true;
  STATE.lastScannedId = null;
  STATE.lastScanMs    = 0;
}

/**
 * Callback saat QR berhasil terbaca.
 * @param {string} decodedText - isi QR Code
 */
function onScanSuccess(decodedText) {
  const id  = decodedText.trim().toUpperCase();
  const now = Date.now();

  /* Debounce: abaikan scan berulang dalam waktu singkat */
  if (id === STATE.lastScannedId && (now - STATE.lastScanMs) < CONFIG.SCAN_DEBOUNCE_MS) {
    return;
  }

  STATE.lastScannedId = id;
  STATE.lastScanMs    = now;

  processID(id);
}

/** Error per-frame dari scanner — diabaikan (QR belum terbaca, normal) */
function onScanError() { /* intentionally empty */ }

/** Update UI tombol scanner & badge status */
function setScannerUIActive(active) {
  el('btn-start').disabled = active;
  el('btn-stop').disabled  = !active;

  const badge    = el('scanner-badge');
  const badgeDot = el('sbadge-dot');
  const badgeText = el('sbadge-text');

  if (active) {
    badge.classList.add('active');
    badgeDot.classList.add('active-dot');
    badgeText.textContent = 'Aktif';
  } else {
    badge.classList.remove('active');
    badgeDot.classList.remove('active-dot');
    badgeText.textContent = 'Nonaktif';
  }
}

/* ════════════════════════════════════════════════════════
   7. SCAN PROCESSOR
   Inti logika: cari di database, cek duplikat, catat hadir
════════════════════════════════════════════════════════ */

/**
 * Proses ID yang di-scan atau di-input manual.
 * @param {string} id - ID wali (e.g. "PUTRA001B" atau "PUTRA001I")
 */
function processID(id) {
  /* Normalisasi ID */
  const idClean = id.toUpperCase().trim();

  /* ── KASUS 1: ID tidak ditemukan di database ── */
  const person = STATE.database.find(p => p.id === idClean);
  if (!person) {
    showResultError(`ID "${idClean}" tidak ditemukan dalam database.`);
    showToast(`ID tidak ditemukan: ${idClean}`, 'error');
    return;
  }

  /* ── KASUS 2: Sudah pernah hadir (duplikat) ── */
  if (STATE.attendanceMap[idClean]) {
    const existing = STATE.attendanceMap[idClean];
    showResultDuplicate(person, existing);
    showToast(`⚠ QR sudah digunakan! ${person.namaWali} (hadir pukul ${existing.waktuHadir})`, 'warning');
    playWarningBeep();
    vibrateDevice();
    return;
  }

  /* ── KASUS 3: Berhasil — catat kehadiran ── */
  const now       = new Date();
  const waktuHadir = now.toLocaleTimeString('id-ID', {
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
  });
  const tanggal = now.toLocaleDateString('id-ID', {
    day: '2-digit', month: '2-digit', year: 'numeric',
  });

  const record = {
    ...person,
    waktuHadir,
    tanggal,
    timestampMs: now.getTime(),
  };

  STATE.attendanceMap[idClean] = record;
  saveToStorage();

  /* Perbarui UI */
  showResultSuccess(record);
  updateStats();
  addLogEntry(record);

  /* Feedback sensorik */
  playBeep();
  vibrateDevice();

  /* Toast */
  showToast(`✓ Absensi berhasil — ${person.namaWali} (${person.jenisWali} dari ${person.namaSantri})`, 'success');

  /* Kirim ke Google Sheets (async) */
  sendToSheets(record);
}

/** Proses input manual dari text field */
function processManualInput() {
  const input = el('manual-id-input');
  const val   = input.value.trim();

  if (!val) {
    showToast('Masukkan ID wali santri terlebih dahulu', 'warning');
    input.focus();
    return;
  }

  processID(val);
  input.value = '';
  input.focus();
}

/* ════════════════════════════════════════════════════════
   8. STATISTICS
   Update semua angka statistik & progress bar
════════════════════════════════════════════════════════ */

function updateStats() {
  const total      = STATE.database.length;
  const attended   = Object.values(STATE.attendanceMap);
  const hadirCount = attended.length;
  const belum      = total - hadirCount;
  const persen     = total > 0 ? Math.round((hadirCount / total) * 100) : 0;

  const putraHadir = attended.filter(a => a.kategori === 'Putra').length;
  const putriHadir = attended.filter(a => a.kategori === 'Putri').length;

  /* Animasi angka dengan "bump" efek */
  setStatValue('val-total',       total,                  '');
  setStatValue('val-hadir',       hadirCount,             '');
  setStatValue('val-belum',       belum,                  '');
  setStatValue('val-putra-hadir', putraHadir,             '');
  setStatValue('val-putri-hadir', putriHadir,             '');
  setStatValue('val-persen',      persen,                 '%');

  /* Progress bar */
  const fill = el('progress-fill');
  fill.style.width = persen + '%';

  const progressWrap = el('progress-wrap');
  progressWrap.setAttribute('aria-valuenow', persen);

  el('progress-label').textContent =
    `${hadirCount} dari ${total} wali santri telah hadir`;

  /* Sembunyikan dot pada progress bila 0 */
  fill.style.setProperty(
    '--show-dot',
    hadirCount > 0 ? 'block' : 'none'
  );
}

/**
 * Set nilai statistik dengan animasi bump.
 * @param {string} id  - element ID
 * @param {number} val - nilai baru
 * @param {string} suffix - suffix string (e.g. '%')
 */
function setStatValue(id, val, suffix) {
  const el_ = el(id);
  if (!el_) return;
  const newText = val + suffix;
  if (el_.textContent !== newText) {
    el_.textContent = newText;
    el_.classList.remove('bump');
    /* Force reflow untuk restart animasi */
    void el_.offsetWidth;
    el_.classList.add('bump');
  }
}

/* ════════════════════════════════════════════════════════
   9. RESULT DISPLAY
   Tampilkan hasil scan di card Hasil Scan
════════════════════════════════════════════════════════ */

/** Tampilkan hasil BERHASIL */
function showResultSuccess(record) {
  const now = new Date(record.timestampMs);

  /* Strip */
  const strip = el('result-strip');
  strip.className = 'result-strip'; /* reset */
  el('result-strip-icon').innerHTML = '✓';
  el('result-strip-title').textContent = 'Absensi Berhasil';
  el('result-strip-time').textContent = `Tercatat pada ${record.waktuHadir}`;
  el('result-strip-time').setAttribute('datetime', now.toISOString());

  /* Fields */
  el('res-id').textContent       = record.id;
  el('res-santri').textContent   = record.namaSantri;
  el('res-wali').textContent     = record.namaWali;
  el('res-jenis').textContent    = record.jenisWali;
  el('res-waktu').textContent    = record.waktuHadir;

  /* Jenis Wali badge (Bapak = biru, Ibu = pink) */
  const jenisBadge = el('res-jenis');
  jenisBadge.className = 'rf-badge' + (record.jenisWali === 'Ibu' ? ' putri' : '');

  /* Kategori badge */
  const katBadge = el('res-kategori');
  katBadge.textContent = record.kategori;
  katBadge.className   = 'rf-badge' + (record.kategori === 'Putri' ? ' putri' : '');

  /* Status badge */
  const statBadge = el('res-status');
  statBadge.textContent = 'Hadir ✓';
  statBadge.className   = 'rf-badge rf-badge--status';

  /* Sembunyikan sheets status */
  el('sheets-sending').hidden = true;
  el('sheets-ok').hidden      = true;
  el('sheets-err').hidden     = true;

  /* Tampilkan */
  el('result-empty').hidden     = true;
  el('result-populated').hidden = false;

  /* Scroll ke result card */
  el('result-card').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

/** Tampilkan hasil DUPLIKAT (sudah pernah hadir) */
function showResultDuplicate(person, existing) {
  const strip = el('result-strip');
  strip.className = 'result-strip strip--warning';
  el('result-strip-icon').innerHTML = '!';
  el('result-strip-title').textContent = 'QR Sudah Digunakan';
  el('result-strip-time').textContent  = `Sebelumnya hadir pukul ${existing.waktuHadir}`;
  el('result-strip-time').removeAttribute('datetime');

  el('res-id').textContent      = person.id;
  el('res-santri').textContent  = person.namaSantri;
  el('res-wali').textContent    = person.namaWali;
  el('res-jenis').textContent   = person.jenisWali;
  el('res-waktu').textContent   = existing.waktuHadir;

  const jenisBadge = el('res-jenis');
  jenisBadge.className = 'rf-badge' + (person.jenisWali === 'Ibu' ? ' putri' : '');

  const katBadge = el('res-kategori');
  katBadge.textContent = person.kategori;
  katBadge.className   = 'rf-badge' + (person.kategori === 'Putri' ? ' putri' : '');

  const statBadge = el('res-status');
  statBadge.textContent = 'Duplikat';
  statBadge.className   = 'rf-badge rf-badge--status duplicate';

  el('sheets-sending').hidden = true;
  el('sheets-ok').hidden      = true;
  el('sheets-err').hidden     = true;

  el('result-empty').hidden     = true;
  el('result-populated').hidden = false;

  el('result-card').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

/** Tampilkan hasil ERROR (tidak ditemukan) */
function showResultError(msg) {
  const strip = el('result-strip');
  strip.className = 'result-strip strip--error';
  el('result-strip-icon').innerHTML = '✕';
  el('result-strip-title').textContent = 'Data Tidak Ditemukan';
  el('result-strip-time').textContent  = msg;
  el('result-strip-time').removeAttribute('datetime');

  el('res-id').textContent      = '—';
  el('res-santri').textContent  = '—';
  el('res-wali').textContent    = '—';
  el('res-jenis').textContent   = '—';
  el('res-waktu').textContent   = '—';

  const jenisBadge = el('res-jenis');
  jenisBadge.className = 'rf-badge';

  const katBadge  = el('res-kategori');
  katBadge.textContent = '—';
  katBadge.className   = 'rf-badge';

  const statBadge = el('res-status');
  statBadge.textContent = 'Tidak Ditemukan';
  statBadge.className   = 'rf-badge rf-badge--status not-found';

  el('sheets-sending').hidden = true;
  el('sheets-ok').hidden      = true;
  el('sheets-err').hidden     = true;

  el('result-empty').hidden     = true;
  el('result-populated').hidden = false;

  el('result-card').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

/* ════════════════════════════════════════════════════════
   10. GOOGLE SHEETS INTEGRATION
   Kirim data kehadiran ke Google Apps Script Web App
════════════════════════════════════════════════════════ */

/**
 * Kirim record kehadiran ke Google Sheets via Apps Script.
 * @param {object} record - data kehadiran
 */
async function sendToSheets(record) {
  /* Jika URL belum dikonfigurasi, tampilkan info dan keluar */
  if (CONFIG.WEB_APP_URL === 'PASTE_GOOGLE_APPS_SCRIPT_URL_HERE') {
    console.info('[Sheets] URL belum dikonfigurasi. Skip.');
    return;
  }

  /* Tampilkan loader */
  el('sheets-sending').hidden = false;
  el('sheets-ok').hidden      = true;
  el('sheets-err').hidden     = true;

  const payload = {
    id:          record.id,
    kategori:    record.kategori,
    namaSantri:  record.namaSantri,
    namaWali:    record.namaWali,
    jenisWali:   record.jenisWali,
    waktuHadir:  record.waktuHadir,
    tanggal:     record.tanggal,
    status:      'Hadir',
  };

  try {
    /* mode: 'no-cors' diperlukan karena Google Apps Script tidak support CORS
       Konsekuensinya: kita tidak bisa membaca response body.
       Anggap sukses jika fetch tidak throw error. */
    await fetch(CONFIG.WEB_APP_URL, {
      method:  'POST',
      mode:    'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(payload),
    });

    el('sheets-sending').hidden = true;
    el('sheets-ok').hidden      = false;

  } catch (err) {
    console.error('[Sheets]', err);
    el('sheets-sending').hidden = true;
    el('sheets-err').hidden     = false;
    el('sheets-err-msg').textContent = 'Gagal kirim ke Sheets: ' + (err.message || 'Network error');
  }
}

/* ════════════════════════════════════════════════════════
   11. ATTENDANCE LOG
   List kehadiran di card Log
════════════════════════════════════════════════════════ */

/**
 * Tambah satu entri ke Log Kehadiran.
 * @param {object} record
 */
function addLogEntry(record) {
  const list    = el('log-list');
  const isEmpty = el('log-empty');

  isEmpty.hidden = true;

  /* Hitung nomor urut */
  const count = Object.keys(STATE.attendanceMap).length;

  const li = document.createElement('li');
  li.className = 'log-item';
  li.setAttribute('role', 'listitem');
  li.innerHTML = `
    <span class="log-item-no">${count}</span>
    <span class="log-item-dot ${record.jenisWali === 'Ibu' ? 'putri' : 'putra'}" aria-hidden="true"></span>
    <div class="log-item-info">
      <div class="log-item-name">${escHtml(record.namaWali)} <span style="font-size:0.7em;opacity:0.6">(${escHtml(record.jenisWali)})</span></div>
      <div class="log-item-id">${escHtml(record.id)} &bull; ${escHtml(record.namaSantri)}</div>
    </div>
    <span class="log-item-time">${escHtml(record.waktuHadir)}</span>
  `;

  /* Sisipkan di awal list (yang terbaru di atas) */
  list.insertBefore(li, list.firstChild);

  /* Update badge */
  el('log-badge').textContent = `${count} hadir`;
}

/** Re-render seluruh log dari attendanceMap */
function renderLogList() {
  const list  = el('log-list');
  const empty = el('log-empty');
  const badge = el('log-badge');

  list.innerHTML = '';

  const records = Object.values(STATE.attendanceMap)
    .sort((a, b) => b.timestampMs - a.timestampMs);

  if (records.length === 0) {
    empty.hidden = false;
    badge.textContent = '0 hadir';
    return;
  }

  empty.hidden = true;
  badge.textContent = `${records.length} hadir`;

  records.forEach((record, idx) => {
    const li = document.createElement('li');
    li.className = 'log-item';
    li.setAttribute('role', 'listitem');
    li.innerHTML = `
      <span class="log-item-no">${idx + 1}</span>
      <span class="log-item-dot ${record.jenisWali === 'Ibu' ? 'putri' : 'putra'}" aria-hidden="true"></span>
      <div class="log-item-info">
        <div class="log-item-name">${escHtml(record.namaWali)} <span style="font-size:0.7em;opacity:0.6">(${escHtml(record.jenisWali)})</span></div>
        <div class="log-item-id">${escHtml(record.id)} &bull; ${escHtml(record.namaSantri)}</div>
      </div>
      <span class="log-item-time">${escHtml(record.waktuHadir)}</span>
    `;
    list.appendChild(li);
  });
}

/* ════════════════════════════════════════════════════════
   12. ADMIN TABLE — Render, Search, Filter
════════════════════════════════════════════════════════ */

/** Set filter kategori */
function setKategoriFilter(btnEl, value) {
  STATE.adminFilter.kategori = value;
  /* Update aria-pressed dan visual */
  btnEl.closest('.chips-row').querySelectorAll('.chip').forEach(c => {
    c.classList.remove('chip--active');
    c.setAttribute('aria-pressed', 'false');
  });
  btnEl.classList.add('chip--active');
  btnEl.setAttribute('aria-pressed', 'true');
  filterAdminTable();
}

/** Set filter status */
function setStatusFilter(btnEl, value) {
  STATE.adminFilter.status = value;
  btnEl.closest('.chips-row').querySelectorAll('.chip').forEach(c => {
    c.classList.remove('chip--active');
    c.setAttribute('aria-pressed', 'false');
  });
  btnEl.classList.add('chip--active');
  btnEl.setAttribute('aria-pressed', 'true');
  filterAdminTable();
}

/** Update search filter dari input */
function filterAdminTable() {
  STATE.adminFilter.search = (el('admin-search').value || '').toLowerCase().trim();
  renderAdminTable();
}

/** Ambil data yang sudah difilter */
function getFilteredData() {
  const { search, kategori, status } = STATE.adminFilter;

  return STATE.database.filter(person => {
    const att     = STATE.attendanceMap[person.id];
    const isHadir = !!att;

    /* Filter kategori */
    if (kategori !== 'all' && person.kategori !== kategori) return false;

    /* Filter status */
    if (status === 'hadir' && !isHadir) return false;
    if (status === 'belum' && isHadir)  return false;

    /* Filter pencarian teks */
    if (search) {
      const haystack = [
        person.id, person.namaSantri, person.namaWali, person.jenisWali, person.kategori,
      ].join(' ').toLowerCase();
      if (!haystack.includes(search)) return false;
    }

    return true;
  });
}

/** Render tabel admin */
function renderAdminTable() {
  const tableEmpty = el('table-empty-state');
  const table      = el('admin-table');
  const tbody      = el('admin-tbody');
  const countBadge = el('admin-count');

  if (STATE.database.length === 0) {
    tableEmpty.hidden = false;
    table.hidden      = true;
    countBadge.textContent = '0 data';
    return;
  }

  const filtered = getFilteredData();
  countBadge.textContent = `${filtered.length} data`;

  tableEmpty.hidden = filtered.length === 0;
  table.hidden      = filtered.length === 0;

  if (filtered.length === 0) {
    tbody.innerHTML = '';
    return;
  }

  /* Build rows — menggunakan DocumentFragment untuk performa */
  const fragment = document.createDocumentFragment();

  filtered.forEach((person, idx) => {
    const att     = STATE.attendanceMap[person.id];
    const isHadir = !!att;

    const tr = document.createElement('tr');
    tr.className = isHadir ? 'row-hadir' : '';
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td class="mono-small">${escHtml(person.id)}</td>
      <td>${escHtml(person.namaSantri)}</td>
      <td>${escHtml(person.namaWali)}</td>
      <td>
        <span class="badge-kat badge-kat--${person.jenisWali === 'Ibu' ? 'putri' : 'putra'}">
          ${escHtml(person.jenisWali)}
        </span>
      </td>
      <td>
        <span class="badge-kat badge-kat--${person.kategori.toLowerCase()}">
          ${escHtml(person.kategori)}
        </span>
      </td>
      <td>${isHadir ? escHtml(att.waktuHadir) : '—'}</td>
      <td>
        <span class="badge-stat ${isHadir ? 'badge-stat--hadir' : 'badge-stat--belum'}">
          ${isHadir ? 'Hadir' : 'Belum Hadir'}
        </span>
      </td>
    `;
    fragment.appendChild(tr);
  });

  tbody.innerHTML = '';
  tbody.appendChild(fragment);
}

/* ════════════════════════════════════════════════════════
   13. EXPORT — Excel, CSV, Print
════════════════════════════════════════════════════════ */

/** Export tabel yang difilter ke file Excel (.xlsx) */
function exportExcel() {
  if (STATE.database.length === 0) {
    showToast('Tidak ada data untuk diekspor', 'warning');
    return;
  }

  const filtered = getFilteredData();
  if (filtered.length === 0) {
    showToast('Tidak ada data yang cocok dengan filter saat ini', 'warning');
    return;
  }

  const exportRows = filtered.map((person, idx) => {
    const att = STATE.attendanceMap[person.id];
    return {
      'No':          idx + 1,
      'ID':          person.id,
      'Nama Santri': person.namaSantri,
      'Nama Wali':   person.namaWali,
      'Jenis Wali':  person.jenisWali,
      'Kategori':    person.kategori,
      'Waktu Hadir': att ? att.waktuHadir : '-',
      'Tanggal':     att ? att.tanggal    : '-',
      'Status':      att ? 'Hadir'        : 'Belum Hadir',
    };
  });

  const ws = XLSX.utils.json_to_sheet(exportRows);

  /* Set column widths */
  ws['!cols'] = [
    { wch: 4 }, { wch: 12 }, { wch: 28 }, { wch: 28 },
    { wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 14 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Laporan Kehadiran');

  const dateStr = new Date().toLocaleDateString('id-ID').replace(/\//g, '-');
  XLSX.writeFile(wb, `Absensi-WaliSantri-${dateStr}.xlsx`);

  showToast(`✓ ${filtered.length} data diekspor ke Excel`, 'success');
}

/** Export tabel yang difilter ke file CSV */
function exportCSV() {
  if (STATE.database.length === 0) {
    showToast('Tidak ada data untuk diekspor', 'warning');
    return;
  }

  const filtered = getFilteredData();
  if (filtered.length === 0) {
    showToast('Tidak ada data yang cocok dengan filter', 'warning');
    return;
  }

  const headers = ['No','ID','Nama Santri','Nama Wali','Jenis Wali','Kategori','Waktu Hadir','Tanggal','Status'];
  const rows    = filtered.map((person, idx) => {
    const att = STATE.attendanceMap[person.id];
    return [
      idx + 1,
      person.id,
      person.namaSantri,
      person.namaWali,
      person.jenisWali,
      person.kategori,
      att ? att.waktuHadir : '-',
      att ? att.tanggal    : '-',
      att ? 'Hadir'        : 'Belum Hadir',
    ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(',');
  });

  const csvContent = [headers.map(h => `"${h}"`).join(','), ...rows].join('\r\n');
  const blob       = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });

  const dateStr = new Date().toLocaleDateString('id-ID').replace(/\//g, '-');
  const link    = document.createElement('a');
  link.href     = URL.createObjectURL(blob);
  link.download = `Absensi-WaliSantri-${dateStr}.csv`;
  link.click();
  URL.revokeObjectURL(link.href);

  showToast(`✓ ${filtered.length} data diekspor ke CSV`, 'success');
}

/** Print laporan kehadiran di jendela baru */
function printLaporan() {
  if (STATE.database.length === 0) {
    showToast('Tidak ada data untuk dicetak', 'warning');
    return;
  }

  const filtered = getFilteredData();
  const dateStr  = new Date().toLocaleDateString('id-ID', {
    day: 'numeric', month: 'long', year: 'numeric',
  });

  const hadirCount = filtered.filter(p => !!STATE.attendanceMap[p.id]).length;
  const belumCount = filtered.length - hadirCount;

  const rows = filtered.map((person, idx) => {
    const att = STATE.attendanceMap[person.id];
    const isHadir = !!att;
    return `
      <tr style="background:${isHadir ? '#f0fff4' : '#fff'}">
        <td>${idx + 1}</td>
        <td style="font-family:monospace;font-size:8pt">${escHtml(person.id)}</td>
        <td>${escHtml(person.namaSantri)}</td>
        <td>${escHtml(person.namaWali)}</td>
        <td style="text-align:center">
          <span style="font-size:7pt;font-weight:bold;padding:1pt 5pt;border-radius:3pt;
            background:${person.jenisWali==='Ibu'?'#fce7f3':'#dbeafe'};
            color:${person.jenisWali==='Ibu'?'#9d174d':'#1d4ed8'}">
            ${escHtml(person.jenisWali)}
          </span>
        </td>
        <td style="text-align:center">
          <span style="font-size:7pt;font-weight:bold;padding:1pt 5pt;border-radius:3pt;
            background:${person.kategori==='Putra'?'#dbeafe':'#fce7f3'};
            color:${person.kategori==='Putra'?'#1d4ed8':'#9d174d'}">
            ${escHtml(person.kategori)}
          </span>
        </td>
        <td style="text-align:center">${isHadir ? escHtml(att.waktuHadir) : '—'}</td>
        <td style="text-align:center">
          <span style="font-size:7pt;font-weight:bold;
            color:${isHadir?'#166534':'#6b7280'}">
            ${isHadir ? '✓ Hadir' : 'Belum'}
          </span>
        </td>
      </tr>
    `;
  }).join('');

  const pw = window.open('', '_blank');
  pw.document.write(`<!DOCTYPE html><html lang="id"><head>
    <meta charset="UTF-8">
    <title>Laporan Kehadiran — ${dateStr}</title>
    <style>
      @page { size: A4 landscape; margin: 12mm; }
      * { box-sizing:border-box; margin:0; padding:0; }
      body { font-family: Arial, sans-serif; font-size: 9pt; color: #1e293b; }
      h1 { font-size: 13pt; text-align:center; margin-bottom: 2mm; }
      .subtitle { text-align:center; font-size: 8pt; color: #64748b; margin-bottom: 4mm; }
      .summary { display:flex; gap:8mm; justify-content:center; margin-bottom:4mm; font-size:8pt; }
      .sum-item { padding:2mm 6mm; border:1px solid #e2e8f0; border-radius:3mm; }
      table { width:100%; border-collapse:collapse; font-size:8pt; }
      th { padding:4pt 6pt; background:#1e293b; color:white; text-align:left; font-size:7pt; letter-spacing:.04em; text-transform:uppercase; }
      td { padding:4pt 6pt; border-bottom:1px solid #e2e8f0; }
      tr:last-child td { border-bottom:none; }
      .footer { margin-top:5mm; text-align:right; font-size:7pt; color:#94a3b8; }
    </style></head><body>
    <h1>Laporan Kehadiran Wali Santri</h1>
    <p class="subtitle">Tanggal cetak: ${dateStr}</p>
    <div class="summary">
      <div class="sum-item">Total Terdaftar: <strong>${filtered.length}</strong></div>
      <div class="sum-item" style="color:#166634">Hadir: <strong>${hadirCount}</strong></div>
      <div class="sum-item" style="color:#991b1b">Belum Hadir: <strong>${belumCount}</strong></div>
      <div class="sum-item">Persentase: <strong>${filtered.length > 0 ? Math.round(hadirCount/filtered.length*100) : 0}%</strong></div>
    </div>
    <table>
      <thead><tr>
        <th>No</th><th>ID</th><th>Nama Santri</th><th>Nama Wali</th><th>Jenis Wali</th>
        <th>Kategori</th><th>Waktu Hadir</th><th>Status</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <div class="footer">Dicetak oleh Sistem Absensi Wali Santri &mdash; ${dateStr}</div>
    <script>window.onload=()=>{window.print();}<\/script>
  </body></html>`);
  pw.document.close();
}

/* ════════════════════════════════════════════════════════
   14. QR GENERATOR
   Generate gambar QR Code dari database, ZIP, Print A4
════════════════════════════════════════════════════════ */

/** Sinkronkan tampilan halaman QR Gen dengan state saat ini */
function syncQRGenPage() {
  const hasDB = STATE.database.length > 0;

  el('qrgen-no-db').hidden  = hasDB;
  el('qrgen-has-db').hidden = !hasDB;

  if (hasDB) {
    const putraCount = STATE.database.filter(d => d.kategori === 'Putra').length;
    const putriCount = STATE.database.filter(d => d.kategori === 'Putri').length;
    el('qs-putra').textContent = putraCount;
    el('qs-putri').textContent = putriCount;
    el('qs-total').textContent = STATE.database.length;
  }
}

/** Reset grid QR dan nonaktifkan tombol download/print */
function resetQRGrid() {
  el('qr-grid-card').hidden = true;
  el('qr-grid').innerHTML   = '';
  el('btn-dl-zip').disabled  = true;
  el('btn-print-qr').disabled = true;
  STATE.generatedQRs = [];
}

/**
 * Generate semua QR Code dari database.
 * Menggunakan async loop dengan yield setiap 10 item
 * agar browser tidak freeze saat proses banyak data.
 */
async function generateAllQR() {
  if (STATE.database.length === 0) {
    showToast('Database kosong. Upload Excel terlebih dahulu.', 'warning');
    return;
  }

  /* Reset state dan UI */
  STATE.generatedQRs = [];
  const grid    = el('qr-grid');
  const gridCard = el('qr-grid-card');
  const progWrap = el('gen-progress-wrap');
  const progFill = el('gen-progress-fill');
  const progLabel = el('gen-progress-label');

  grid.innerHTML = '';
  gridCard.hidden = false;
  progWrap.hidden = false;
  el('btn-gen-all').disabled   = true;
  el('btn-dl-zip').disabled    = true;
  el('btn-print-qr').disabled  = true;

  const total = STATE.database.length;

  try {
    for (let i = 0; i < total; i++) {
      const person = STATE.database[i];

      const dataURL = generateQRDataURL(person);

      /* Simpan untuk ZIP dan Print */
      STATE.generatedQRs.push({ person, dataURL });

      /* Buat card QR */
      const card = buildQRCard(person, dataURL, i);
      grid.appendChild(card);

      /* Update progress */
      const pct = Math.round(((i + 1) / total) * 100);
      progFill.style.width   = pct + '%';
      progLabel.textContent  = `Generating... ${pct}% (${i + 1}/${total})`;

      /* Yield ke browser setiap 8 item agar tidak freeze */
      if (i % 8 === 7) await sleep(0);
    }

    /* Selesai */
    progWrap.hidden = true;
    el('btn-dl-zip').disabled   = false;
    el('btn-print-qr').disabled = false;
    el('qr-count').textContent  = `${total} QR`;

    showToast(`✓ ${total} QR Code berhasil digenerate`, 'success');

  } catch (err) {
    console.error('[QR Gen]', err);
    showToast('Gagal generate QR: ' + err.message, 'error');
  } finally {
    el('btn-gen-all').disabled = false;
    progWrap.hidden = true;
  }
}

/**
 * Generate QR Code berkualitas tinggi sebagai PNG Data URL.
 * Output: canvas 480 x 630px berisi QR besar + nama santri + badge kategori.
 * Gambar ini langsung bisa didownload / dicetak tanpa pecah.
 *
 * @param {object} person - { id, namaSantri, namaBapak, kategori }
 * @returns {string} PNG data URL
 */
function generateQRDataURL(person) {
  try {
    /* ── 1. Buat data QR menggunakan qrcode-generator ── */
    const qr = qrcode(0, 'M');  /* typeNumber 0 = auto-size, level M */
    qr.addData(person.id);
    qr.make();
    const moduleCount = qr.getModuleCount();

    /* ── 2. Ukuran canvas ── */
    const QR_SIZE  = 420;  /* lebar & tinggi area QR dalam px */
    const PAD      = 36;   /* padding kiri-kanan-atas */
    const TEXT_H   = 160;  /* tinggi area teks di bawah QR */
    const TOP_BAR  = 10;   /* garis warna di atas */
    const W        = QR_SIZE + PAD * 2;   /* = 492 */
    const H        = TOP_BAR + QR_SIZE + PAD + TEXT_H; /* = 626 */

    const canvas   = document.createElement('canvas');
    canvas.width   = W;
    canvas.height  = H;
    const ctx      = canvas.getContext('2d');

    /* ── 3. Background putih ── */
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);

    /* ── 4. Border luar tipis ── */
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth   = 1.5;
    ctx.strokeRect(0.75, 0.75, W - 1.5, H - 1.5);

    /* Accent bar warna berdasarkan JENIS WALI (Bapak=biru, Ibu=pink) */
    const accent   = person.jenisWali === 'Ibu' ? '#ec4899' : '#3b82f6';
    ctx.fillStyle  = accent;
    ctx.fillRect(0, 0, W, TOP_BAR);

    /* ── 6. Gambar modul QR Code ── */
    const cellPx   = QR_SIZE / moduleCount;
    ctx.fillStyle  = '#000000';
    for (let row = 0; row < moduleCount; row++) {
      for (let col = 0; col < moduleCount; col++) {
        if (qr.isDark(row, col)) {
          ctx.fillRect(
            PAD + Math.floor(col * cellPx),
            TOP_BAR + PAD + Math.floor(row * cellPx),
            Math.ceil(cellPx),
            Math.ceil(cellPx)
          );
        }
      }
    }

    /* ── 7. Teks di bawah QR ── */
    const baseY = TOP_BAR + PAD + QR_SIZE + 20; /* Y awal area teks */
    const maxW  = W - PAD * 2;
    ctx.textAlign = 'center';

    /* ID (monospace, abu-abu) */
    ctx.fillStyle = '#3b82f6';
    ctx.font      = 'bold 20px "Courier New", monospace';
    ctx.fillText(person.id, W / 2, baseY, maxW);

    /* Garis pemisah */
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth   = 1;
    ctx.beginPath();
    ctx.moveTo(PAD, baseY + 12);
    ctx.lineTo(W - PAD, baseY + 12);
    ctx.stroke();

    /* Nama Santri (bold hitam, besar) */
    ctx.fillStyle = '#0f172a';
    ctx.font      = 'bold 22px Arial, sans-serif';
    /* Potong nama panjang agar tidak melebihi canvas */
    let namaDisplay = person.namaSantri;
    while (ctx.measureText(namaDisplay).width > maxW && namaDisplay.length > 4) {
      namaDisplay = namaDisplay.slice(0, -1);
    }
    if (namaDisplay !== person.namaSantri) namaDisplay += '...';
    ctx.fillText(namaDisplay, W / 2, baseY + 46, maxW);

    /* Nama Wali (abu-abu, lebih kecil) */
    ctx.fillStyle = '#64748b';
    ctx.font      = '16px Arial, sans-serif';
    let waliDisplay = person.jenisWali + ': ' + person.namaWali;
    while (ctx.measureText(waliDisplay).width > maxW && waliDisplay.length > 8) {
      waliDisplay = waliDisplay.slice(0, -1);
    }
    if (waliDisplay !== (person.jenisWali + ': ' + person.namaWali)) waliDisplay += '...';
    ctx.fillText(waliDisplay, W / 2, baseY + 76, maxW);

    /* Badge Jenis Wali (pill berwarna) */
    const bdgText  = person.jenisWali + ' • ' + person.kategori;
    const bdgW     = 80;
    const bdgH     = 30;
    const bdgX     = W / 2 - bdgW / 2;
    const bdgY     = baseY + 92;
    ctx.fillStyle  = accent;
    ctxRoundRect(ctx, bdgX, bdgY, bdgW, bdgH, 15);
    ctx.fill();
    ctx.fillStyle  = '#ffffff';
    ctx.font       = 'bold 15px Arial, sans-serif';
    ctx.fillText(bdgText, W / 2, bdgY + 20);

    /* Footer kecil */
    ctx.fillStyle  = '#cbd5e1';
    ctx.font       = '12px Arial, sans-serif';
    ctx.fillText('Absensi Wali Santri', W / 2, H - 10);

    return canvas.toDataURL('image/png');

  } catch (err) {
    console.error('[QRGen] Canvas error:', err);
    /* Fallback darurat: pakai createDataURL bawaan library */
    try {
      const qr2 = qrcode(4, 'M');
      qr2.addData(person.id || String(person));
      qr2.make();
      return qr2.createDataURL(8, 20);
    } catch (e2) { return ''; }
  }
}

/**
 * Gambar rounded rectangle ke canvas context.
 * (Polyfill untuk ctx.roundRect yang belum didukung semua browser)
 * @param {CanvasRenderingContext2D} ctx
 * @param {number} x
 * @param {number} y
 * @param {number} w
 * @param {number} h
 * @param {number} r - corner radius
 */
function ctxRoundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

/**
 * Buat element card QR untuk satu wali santri.
 * @param {object} person
 * @param {string} dataURL - PNG data URL dari QR code
 * @param {number} index   - index di STATE.generatedQRs
 * @returns {HTMLElement}
 */
function buildQRCard(person, dataURL, index) {
  const div = document.createElement('div');
  div.className = 'qr-item';
  div.setAttribute('role', 'listitem');
  div.innerHTML = `
    <div class="qr-img-wrap">
      <img src="${dataURL}"
        alt="QR Code untuk ${escHtml(person.id)} — ${escHtml(person.namaSantri)}"
        width="120" height="120" loading="lazy">
    </div>
    <div class="qr-info">
      <span class="qr-id">${escHtml(person.id)}</span>
      <span class="qr-nama">${escHtml(person.namaSantri)}</span>
      <span class="qr-wali">${escHtml(person.namaWali)}</span>
      <span class="qr-badge qr-badge--${person.jenisWali === 'Ibu' ? 'putri' : 'putra'}">${escHtml(person.jenisWali)}</span>
      <span class="qr-badge qr-badge--${person.kategori.toLowerCase()}">${escHtml(person.kategori)}</span>
    </div>
    <button class="btn btn-sm btn-ghost qr-dl-btn"
      onclick="downloadSingleQR(${index})"
      aria-label="Download QR Code untuk ${escHtml(person.id)}">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
        <polyline points="7 10 12 15 17 10"/>
        <line x1="12" y1="15" x2="12" y2="3"/>
      </svg>
      Download
    </button>
  `;
  return div;
}

/**
 * Download QR Code tunggal sebagai file PNG.
 * @param {number} index - index di STATE.generatedQRs
 */
function downloadSingleQR(index) {
  const item = STATE.generatedQRs[index];
  if (!item) return;

  const { person, dataURL } = item;
  const filename = `QR-${person.id}-${person.namaSantri.replace(/\s+/g, '-')}.png`;
  const link     = document.createElement('a');
  link.href      = dataURL;
  link.download  = filename;
  link.click();
}

/** Download semua QR Code dalam satu file ZIP */
async function downloadAllZIP() {
  if (STATE.generatedQRs.length === 0) {
    showToast('Generate QR Code terlebih dahulu', 'warning');
    return;
  }

  showToast('Menyiapkan file ZIP...', 'info');

  try {
    const zip        = new JSZip();
    const putraFolder = zip.folder('Putra');
    const putriFolder = zip.folder('Putri');

    for (const { person, dataURL } of STATE.generatedQRs) {
      /* Ambil bagian base64 setelah 'data:image/png;base64,' */
      const base64   = dataURL.split(',')[1];
      const filename = `${person.id}-${person.namaSantri.replace(/\s+/g, '-')}.png`;
      const folder   = person.kategori === 'Putra' ? putraFolder : putriFolder;
      folder.file(filename, base64, { base64: true });
    }

    const blob    = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    const dateStr = new Date().toLocaleDateString('id-ID').replace(/\//g, '-');
    saveAs(blob, `QR-WaliSantri-${dateStr}.zip`);

    showToast(`✓ ${STATE.generatedQRs.length} QR Code didownload sebagai ZIP`, 'success');

  } catch (err) {
    console.error('[ZIP]', err);
    showToast('Gagal membuat ZIP: ' + err.message, 'error');
  }
}

/** Cetak semua QR Code dalam format A4 (4 kolom) */
function printAllQR() {
  if (STATE.generatedQRs.length === 0) {
    showToast('Generate QR Code terlebih dahulu', 'warning');
    return;
  }

  const dateStr = new Date().toLocaleDateString('id-ID', {
    day: 'numeric', month: 'long', year: 'numeric',
  });

  const cards = STATE.generatedQRs.map(({ person, dataURL }) => `
    <div class="qr-card">
      <img src="${dataURL}" alt="QR ${escHtml(person.id)}" width="148" height="148">
      <div class="qr-id">${escHtml(person.id)}</div>
      <div class="qr-nama">${escHtml(person.namaSantri)}</div>
      <div class="qr-wali">${escHtml(person.jenisWali)}: ${escHtml(person.namaWali)}</div>
      <div class="qr-badge qr-badge--${person.jenisWali === 'Ibu' ? 'putri' : 'putra'}">${escHtml(person.jenisWali)} &bull; ${escHtml(person.kategori)}</div>
    </div>
  `).join('');

  const pw = window.open('', '_blank');
  pw.document.write(`<!DOCTYPE html><html lang="id"><head>
    <meta charset="UTF-8">
    <title>QR Code Wali Santri</title>
    <style>
      @page { size: A4; margin: 10mm; }
      * { box-sizing:border-box; margin:0; padding:0; }
      body { font-family: Arial, sans-serif; background: white; }
      h1 { text-align:center; font-size:13pt; margin-bottom:2mm; color:#1e293b; }
      .sub { text-align:center; font-size:8pt; color:#64748b; margin-bottom:5mm; }
      .grid { display:grid; grid-template-columns:repeat(4,1fr); gap:5mm; }
      .qr-card {
        border:1px solid #e2e8f0; border-radius:5mm; padding:3.5mm;
        text-align:center; page-break-inside:avoid;
      }
      .qr-card img { display:block; margin:0 auto 2mm; }
      .qr-id   { font-family:monospace; font-size:7.5pt; font-weight:bold; color:#3b82f6; }
      .qr-nama { font-size:7.5pt; font-weight:bold; color:#1e293b; margin:1.5mm 0 0.5mm; line-height:1.3; }
      .qr-bapak { font-size:6.5pt; color:#64748b; margin-bottom:1.5mm; }
      .qr-badge {
        display:inline-block; font-size:6.5pt; font-weight:bold;
        padding:1mm 3mm; border-radius:999px; color:white;
      }
      .qr-badge--putra { background:#3b82f6; }
      .qr-badge--putri { background:#ec4899; }
    </style></head><body>
    <h1>QR Code Absensi Wali Santri</h1>
    <p class="sub">Total: ${STATE.generatedQRs.length} QR Code &mdash; Dicetak: ${dateStr}</p>
    <div class="grid">${cards}</div>
    <script>window.onload=()=>{window.print();}<\/script>
  </body></html>`);
  pw.document.close();
}

/* ════════════════════════════════════════════════════════
   15. AUDIO & HAPTIC FEEDBACK
════════════════════════════════════════════════════════ */

/** AudioContext singleton */
let _audioCtx = null;

function getAudioCtx() {
  if (!_audioCtx) {
    try {
      _audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    } catch (e) { /* AudioContext tidak tersedia */ }
  }
  return _audioCtx;
}

/**
 * Putar suara beep saat absensi berhasil.
 * Dibuat via Web Audio API — tidak memerlukan file audio eksternal.
 */
function playBeep() {
  try {
    const ctx = getAudioCtx();
    if (!ctx) return;

    const osc   = ctx.createOscillator();
    const gain  = ctx.createGain();

    osc.connect(gain);
    gain.connect(ctx.destination);

    osc.type      = 'sine';
    osc.frequency.setValueAtTime(880, ctx.currentTime);
    osc.frequency.exponentialRampToValueAtTime(1100, ctx.currentTime + 0.08);

    gain.gain.setValueAtTime(0.4, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.25);

    osc.start(ctx.currentTime);
    osc.stop(ctx.currentTime + 0.25);
  } catch (e) { /* abaikan jika AudioContext tidak tersedia */ }
}

/**
 * Suara beep WARNING (duplikat scan) — nada lebih rendah & berbeda.
 */
function playWarningBeep() {
  try {
    const ctx  = getAudioCtx();
    if (!ctx) return;
    const osc  = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.type = 'square';
    osc.frequency.setValueAtTime(440, ctx.currentTime);
    osc.frequency.setValueAtTime(330, ctx.currentTime + 0.1);
    gain.gain.setValueAtTime(0.25, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
    osc.start(ctx.currentTime);
    osc.stop(ctx.currentTime + 0.35);
  } catch (e) { /* abaikan */ }
}

function vibrateDevice() {
  if ('vibrate' in navigator) {
    navigator.vibrate([80, 30, 80]);
  }
}

/* ════════════════════════════════════════════════════════
   16. TOAST NOTIFICATIONS
════════════════════════════════════════════════════════ */

const TOAST_ICONS = {
  success: '✓',
  error:   '✕',
  warning: '⚠',
  info:    'ℹ',
};

const TOAST_DURATION = {
  success: 3000,
  error:   4500,
  warning: 4000,
  info:    2500,
};

/**
 * Tampilkan toast notification.
 * @param {string} message
 * @param {'success'|'error'|'warning'|'info'} [type='info']
 */
function showToast(message, type = 'info') {
  const container = el('toast-container');

  const toast = document.createElement('div');
  toast.className = `toast toast--${type}`;
  toast.setAttribute('role', 'alert');
  toast.innerHTML = `
    <span class="toast-icon" aria-hidden="true">${TOAST_ICONS[type] || 'ℹ'}</span>
    <span class="toast-msg">${escHtml(message)}</span>
  `;

  container.appendChild(toast);

  /* Auto dismiss */
  const duration = TOAST_DURATION[type] || 3000;
  setTimeout(() => dismissToast(toast), duration);

  /* Click to dismiss */
  toast.addEventListener('click', () => dismissToast(toast));
}

function dismissToast(toastEl) {
  if (!toastEl.parentNode) return;
  toastEl.classList.add('toast-out');
  toastEl.addEventListener('animationend', () => toastEl.remove(), { once: true });
}

/* ════════════════════════════════════════════════════════
   17. SYNC FROM SHEETS
   Sinkronisasi 2 arah: baca data kehadiran dari Google Sheets
   setiap 1 detik agar semua perangkat menampilkan data yang sama.
════════════════════════════════════════════════════════ */

let _syncBusy     = false;  /* cegah request tumpuk */
let _syncTimer    = null;
let _syncEnabled  = false;

/**
 * Mulai sinkronisasi otomatis.
 * Request berikutnya baru dikirim SETELAH request sebelumnya selesai,
 * sehingga koneksi lambat tidak membuat request menumpuk.
 */
function startSync() {
  if (CONFIG.WEB_APP_URL === 'PASTE_GOOGLE_APPS_SCRIPT_URL_HERE') return;
  _syncEnabled = true;

  /* Tampilkan sync bar */
  const bar = el('sync-bar');
  if (bar) bar.hidden = false;

  /* Sinkron langsung saat pertama kali */
  syncFromSheets();

  /* Jadwal ulang setiap kali selesai (1 detik setelah selesai) */
  function scheduleNext() {
    if (!_syncEnabled) return;
    _syncTimer = setTimeout(async () => {
      await syncFromSheets();
      scheduleNext();
    }, 1000);
  }
  scheduleNext();
}

/** Hentikan sinkronisasi */
function stopSync() {
  _syncEnabled = false;
  if (_syncTimer) { clearTimeout(_syncTimer); _syncTimer = null; }
}

/**
 * Satu siklus sinkronisasi:
 * GET data dari Sheets → merge ke STATE.attendanceMap → update UI.
 */
async function syncFromSheets() {
  if (_syncBusy) return;
  if (STATE.database.length === 0) return; /* tunggu database dimuat dulu */
  if (!navigator.onLine) {
    setSyncUI('error', 'Offline — tidak dapat sinkron');
    return;
  }

  _syncBusy = true;
  setSyncUI('syncing', 'Menyinkronkan...');

  try {
    const resp = await fetch(CONFIG.WEB_APP_URL, {
      method: 'GET',
      cache:  'no-store',
    });

    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

    const json = await resp.json();
    if (json.status !== 'ok' || !Array.isArray(json.data)) {
      throw new Error('Format respons tidak valid');
    }

    let newCount = 0;

    for (const row of json.data) {
      const id = String(row.id || '').trim().toUpperCase();
      if (!id) continue;

      /* Sudah ada di lokal → skip */
      if (STATE.attendanceMap[id]) continue;

      /* Cari di database lokal */
      const person = STATE.database.find(p => p.id === id);
      if (!person) continue;

      /* Catat sebagai hadir (dari Sheets) */
      STATE.attendanceMap[id] = {
        ...person,
        waktuHadir:  String(row.waktuHadir || '—'),
        tanggal:     String(row.tanggal    || '—'),
        timestampMs: Date.now(),
        fromSync:    true,
      };
      newCount++;
    }

    if (newCount > 0) {
      saveToStorage();
      updateStats();
      renderLogList();
      if (STATE.activePage === 'admin') renderAdminTable();
      showToast(`↓ ${newCount} data baru dari perangkat lain`, 'info');
    }

    const now = new Date().toLocaleTimeString('id-ID', {
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    });
    setSyncUI('ok', `Sinkron ${now}`);

  } catch (err) {
    console.warn('[Sync]', err.message);
    setSyncUI('error', 'Gagal sinkron — coba lagi...');
  } finally {
    _syncBusy = false;
  }
}

/**
 * Update tampilan sync bar.
 * @param {'syncing'|'ok'|'error'} status
 * @param {string} text
 */
function setSyncUI(status, text) {
  const dot = el('sync-dot');
  const txt = el('sync-text');
  if (!dot || !txt) return;
  dot.className = 'sync-dot ' + status;
  txt.textContent = text;
}

/* ════════════════════════════════════════════════════════
   18. CONNECTION STATUS
════════════════════════════════════════════════════════ */

function updateConnectionStatus() {
  const online     = navigator.onLine;
  const dot        = el('conn-dot');
  const label      = el('conn-label');

  dot.className    = 'conn-dot ' + (online ? 'online' : 'offline');
  label.textContent = online ? 'Online' : 'Offline';
}

/* ════════════════════════════════════════════════════════
   18. DB BANNER UPDATE
════════════════════════════════════════════════════════ */

/** Perbarui banner status database */
function updateDBBanner() {
  const statusText = el('db-status-text');
  const clearBtn   = el('btn-clear-db');

  if (STATE.database.length === 0) {
    statusText.textContent = 'Database belum dimuat — Upload file Excel untuk memulai';
    statusText.classList.remove('loaded');
    clearBtn.hidden = true;
  } else {
    const putra = STATE.database.filter(d => d.kategori === 'Putra').length;
    const putri = STATE.database.filter(d => d.kategori === 'Putri').length;
    statusText.textContent = `${STATE.database.length} data dimuat: ${putra} Putra, ${putri} Putri`;
    statusText.classList.add('loaded');
    clearBtn.hidden = false;
  }
}

/* ════════════════════════════════════════════════════════
   19. UTILITIES
════════════════════════════════════════════════════════ */

/** Shorthand getElementById */
function el(id) { return document.getElementById(id); }

/**
 * Escape HTML untuk mencegah XSS.
 * @param {*} str
 * @returns {string}
 */
function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Sleep (yield ke browser event loop).
 * @param {number} ms
 * @returns {Promise<void>}
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/* ════════════════════════════════════════════════════════
   20. FILE INPUT LISTENERS
   Menghubungkan semua input[type=file] ke handler yang sama
════════════════════════════════════════════════════════ */

function bindFileInputs() {
  /* Input global (di banner) */
  el('excel-upload-global').addEventListener('change', function() {
    if (this.files[0]) handleExcelFile(this.files[0]);
    this.value = ''; /* Reset agar bisa upload file yang sama lagi */
  });

  /* Input di halaman QR Gen */
  el('excel-upload-qrgen').addEventListener('change', function() {
    if (this.files[0]) handleExcelFile(this.files[0]);
    this.value = '';
  });

  /* Drag & Drop di zona upload QR Gen */
  const dropZone = el('qrgen-drop-zone');
  if (dropZone) {
    dropZone.addEventListener('dragover', e => {
      e.preventDefault();
      dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', e => {
      e.preventDefault();
      dropZone.classList.remove('drag-over');
      const file = e.dataTransfer.files[0];
      if (file) handleExcelFile(file);
    });
    /* Keyboard activation */
    dropZone.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        el('excel-upload-qrgen').click();
      }
    });
  }
}

/* ════════════════════════════════════════════════════════
   21. INIT — Titik masuk aplikasi
════════════════════════════════════════════════════════ */

document.addEventListener('DOMContentLoaded', function() {

  /* ── Muat data dari localStorage ── */
  loadFromStorage();

  /* ── Footer year ── */
  el('footer-year').textContent = new Date().getFullYear();

  /* ── Koneksi online/offline ── */
  updateConnectionStatus();
  window.addEventListener('online',  updateConnectionStatus);
  window.addEventListener('offline', updateConnectionStatus);

  /* ── Hubungkan semua file input ── */
  bindFileInputs();

  /* ── Update UI awal ── */
  updateDBBanner();
  updateStats();
  renderLogList();
  renderAdminTable();
  syncQRGenPage();

  /* ── Mulai sinkronisasi otomatis dari Google Sheets ── */
  startSync();

  /* ── Pastikan hanya scanner page yang tampil pertama ── */
  document.querySelectorAll('.page').forEach(p => {
    if (p.id === 'page-scanner') {
      p.hidden = false;
      p.classList.add('page--active');
    } else {
      p.hidden = true;
      p.classList.remove('page--active');
    }
  });

  /* ── Log selamat datang ── */
  console.log(
    '%c Absensi Wali Santri v2.0 ',
    'background:#3b82f6;color:white;padding:4px 10px;border-radius:4px;font-weight:bold',
  );
  console.log('Database:', STATE.database.length, 'data');
  console.log('Absensi :', Object.keys(STATE.attendanceMap).length, 'hadir');

});

/* ════════════════════════════════════════════════════════
   ════ GOOGLE APPS SCRIPT — PANDUAN SETUP ════

   1. Buka script.google.com → New Project
   2. Paste kode berikut:

   function doPost(e) {
     var sheet = SpreadsheetApp.getActiveSpreadsheet()
                   .getSheetByName('Absensi') ||
                 SpreadsheetApp.getActiveSpreadsheet()
                   .insertSheet('Absensi');

     // Buat header jika sheet baru
     if (sheet.getLastRow() === 0) {
       sheet.appendRow([
         'No','ID','Kategori','Nama Santri','Nama Bapak',
         'Nama Ibu','Waktu Hadir','Tanggal','Status'
       ]);
     }

     var data = JSON.parse(e.postData.contents);
     sheet.appendRow([
       sheet.getLastRow(),
       data.id, data.kategori, data.namaSantri,
       data.namaBapak, data.namaIbu,
       data.waktuHadir, data.tanggal, data.status
     ]);

     return ContentService
       .createTextOutput(JSON.stringify({status:'OK'}))
       .setMimeType(ContentService.MimeType.JSON);
   }

   3. Deploy → New deployment → Web app
      - Execute as: Me
      - Who has access: Anyone
   4. Copy URL → Paste ke CONFIG.WEB_APP_URL di atas

════════════════════════════════════════════════════════ */
