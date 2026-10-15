/*************************************************************
 * 16_Rekap.gs
 * Rekap kehadiran per karyawan (Jadwal/Hadir/Alfa/Izin/Sakit/Cuti/OFF/%Kehadiran/Rank),
 * dihitung LANGSUNG dari grid Jadwal & Realisasi -- bukan dari sheet "Rekap" (yang bisa
 * stale/kosong kalau rumusnya belum diisi ulang untuk periode baru).
 *
 * Dua pintu masuk, penghitungnya sama (tallyAttendance_):
 *   getRekap_()            periode AKTIF: karyawan dicocokkan berdasarkan POSISI BARIS
 *                          (asumsi seluruh sistem, lihat checkRowAlignment_); %Kehadiran tidak
 *                          dibulatkan; Rank 1 = % PALING RENDAH ("Rank (Terendah)").
 *   computeRekapDariGrid_  periode ARSIP / tren: dicocokkan berdasarkan NAMA; %Kehadiran
 *                          dibulatkan 4 desimal; Rank 1 = % PALING TINGGI (nilai sama = rank sama).
 * Perbedaan rank/pembulatan itu perilaku warisan yang sengaja dipertahankan (lihat MIGRASI.md).
 *************************************************************/

/** Hitung jumlah hari terjadwal dan rincian status dari dua deret kode harian. */
function tallyAttendance_(shiftDays, realisasiDays) {
  const t = { jadwal: 0, hadir: 0, alfa: 0, izin: 0, sakit: 0, cuti: 0, off: 0 };

  shiftDays.forEach(v => {
    const code = cellCode_(v);
    if (code !== '' && code !== 'OFF') t.jadwal++;
  });

  realisasiDays.forEach(v => {
    const code = cellCode_(v);
    if (code === '1' || code === '2') t.hadir++; // "2" = Lembur, tetap dihitung Hadir
    else if (code === 'A') t.alfa++;
    else if (code === 'I') t.izin++;
    else if (code === 'S') t.sakit++;
    else if (code === 'C') t.cuti++;
    else if (code === 'OFF') t.off++;
  });
  return t;
}

/** Satu baris rekap (urutan kunci dipertahankan; "rank" ditambahkan kemudian). */
function makeRekapRow_(no, nama, jabatan, tally, persenKehadiran) {
  return {
    no: no,
    nama: nama,
    jabatan: jabatan,
    jadwal: tally.jadwal,
    hadir: tally.hadir,
    alfa: tally.alfa,
    izin: tally.izin,
    sakit: tally.sakit,
    cuti: tally.cuti,
    off: tally.off,
    persenKehadiran: persenKehadiran
  };
}

/** Rekap periode aktif (cocok per posisi baris). */
function getRekap_() {
  const jadwal = readGrid_(CONFIG.SHEET_JADWAL);
  const realisasi = readGrid_(CONFIG.SHEET_REALISASI);
  const dayCount = Math.min(jadwal.dayCount, realisasi.dayCount);
  const numEmployees = Math.min(jadwal.employees.length, realisasi.employees.length);

  const rows = [];
  for (let i = 0; i < numEmployees; i++) {
    const j = jadwal.employees[i];
    const r = realisasi.employees[i];
    const tally = tallyAttendance_(j.days.slice(0, dayCount), r.days.slice(0, dayCount));
    const persen = tally.jadwal > 0 ? (tally.hadir / tally.jadwal) : 0;

    rows.push(makeRekapRow_(r.no || j.no || (i + 1), r.nama || j.nama, r.jabatan || j.jabatan, tally, persen));
  }

  // Rank 1 = % kehadiran PALING RENDAH (sesuai label "Rank (Terendah)" di sheet Rekap lama)
  rows.slice().sort((a, b) => a.persenKehadiran - b.persenKehadiran)
    .forEach((row, idx) => { row.rank = idx + 1; });

  return rows;
}

/** Rekap dari dua grid (periode arsip, tren bulanan): dicocokkan per NAMA. */
function computeRekapDariGrid_(jadwalGrid, realisasiGrid) {
  const realisasiByName = {};
  realisasiGrid.employees.forEach(e => { realisasiByName[nameKey_(e.nama)] = e; });

  const rekap = jadwalGrid.employees.map(jEmp => {
    const rEmp = realisasiByName[nameKey_(jEmp.nama)];
    const tally = tallyAttendance_(jEmp.days, rEmp ? rEmp.days : []);
    const persen = tally.jadwal > 0 ? Math.round((tally.hadir / tally.jadwal) * 10000) / 10000 : 0;
    return makeRekapRow_(jEmp.no, jEmp.nama, jEmp.jabatan, tally, persen);
  });

  // Rank 1 = % Hadir tertinggi (mirip RANK() Google Sheets: nilai sama dapat rank sama).
  const sortedPct = rekap.map(r => r.persenKehadiran).sort((a, b) => b - a);
  rekap.forEach(r => { r.rank = sortedPct.indexOf(r.persenKehadiran) + 1; });

  return rekap;
}
