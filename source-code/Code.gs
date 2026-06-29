/**
 * Project: Rekap Otomatis Respons Google Form
 * Fungsi: Membuat ringkasan jumlah jawaban responden
 * secara otomatis di sheet bernama "Ringkasan".
 */

function buatRingkasan() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const namaSheetRingkasan = 'Ringkasan';

  // Mencari sheet respons utama.
  const sheetRespons = spreadsheet
    .getSheets()
    .find(sheet => sheet.getName() !== namaSheetRingkasan);

  if (!sheetRespons) {
    throw new Error('Sheet respons Google Form tidak ditemukan.');
  }

  const semuaData = sheetRespons.getDataRange().getDisplayValues();

  if (semuaData.length < 2) {
    throw new Error('Belum ada data responden yang dapat diringkas.');
  }

  const judulKolom = semuaData[0];

  // Menghapus baris yang benar-benar kosong.
  const dataResponden = semuaData.slice(1).filter(baris =>
    baris.some(nilai => nilai.toString().trim() !== '')
  );

  let sheetRingkasan =
    spreadsheet.getSheetByName(namaSheetRingkasan);

  if (!sheetRingkasan) {
    sheetRingkasan =
      spreadsheet.insertSheet(namaSheetRingkasan);
  }

  sheetRingkasan.clear();

  // Informasi utama.
  sheetRingkasan.getRange('A1:B4').setValues([
    ['RINGKASAN HASIL GOOGLE FORM', ''],
    ['Jumlah Responden', dataResponden.length],
    ['Terakhir Diperbarui', new Date()],
    ['Sumber Data', sheetRespons.getName()]
  ]);

  const hasilRingkasan = [
    ['Pertanyaan', 'Jawaban', 'Jumlah']
  ];

  /*
   * Kolom identitas tidak ikut dihitung agar nama,
   * NIM, kelas, dan email responden tidak ditampilkan.
   */
  const kolomIdentitas =
    /timestamp|stempel waktu|nama|nim|kelas|email/i;

  for (let kolom = 0; kolom < judulKolom.length; kolom++) {
    const pertanyaan = judulKolom[kolom];

    if (kolomIdentitas.test(pertanyaan)) {
      continue;
    }

    const jumlahJawaban = {};

    dataResponden.forEach(baris => {
      const jawaban = baris[kolom]
        ? baris[kolom].toString().trim()
        : '';

      if (jawaban !== '') {
        jumlahJawaban[jawaban] =
          (jumlahJawaban[jawaban] || 0) + 1;
      }
    });

    const daftarJawaban = Object.entries(jumlahJawaban)
      .sort((a, b) => b[1] - a[1]);

    daftarJawaban.forEach(([jawaban, jumlah]) => {
      hasilRingkasan.push([
        pertanyaan,
        jawaban,
        jumlah
      ]);
    });
  }

  sheetRingkasan
    .getRange(
      6,
      1,
      hasilRingkasan.length,
      hasilRingkasan[0].length
    )
    .setValues(hasilRingkasan);

  // Mengatur tampilan.
  sheetRingkasan.getRange('A1:B1').setFontWeight('bold');
  sheetRingkasan.getRange('A6:C6').setFontWeight('bold');
  sheetRingkasan.setFrozenRows(6);
  sheetRingkasan.autoResizeColumns(1, 3);
}

/**
 * Fungsi ini dijalankan otomatis setiap ada
 * responden yang mengirim Google Form.
 */
function saatFormDikirim(e) {
  buatRingkasan();
}