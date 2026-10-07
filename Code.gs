/**
 * Control Warehouse APM Depok - API data (Google Apps Script)
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me | Who has access: Anyone
 */
const SPREADSHEET_ID = '1HqAHIffe1AxYAFYFfAXGRjRjI02miOuu5CshhRmrPxE';

// Batas minimum stok per Kode Barang. Stok Akhir <= limit => "Harus segera di requestkan".
// Ubah angka di sini kapan saja untuk mengatur limit warning.
const LIMITS = {
  'H-001': 10,   // PATHCORD SC/UPC-SC/UPC 5M
  'H-002': 2,    // SPLITER 1/16 DENSIA
  'H-003': 1,    // ALTOS A1500
  'H-004': 1,    // ONT HUAWEI Echolife EG8021V5 (limit 1, sesuai kolom H sheet)
  'H-005': 30,   // ONT ZTE ZXHN F679D
  'H-008': 500,  // DropWire 1 Core (meter)
  'H-010': 20,   // ROSET INDOTEK
  'H-011': 20,   // PRECON 50M NEXTFIBER
  'H-012': 20,   // PRECON 100M
  'H-013': 20,   // PRECON 150M NEXTFIBER
  'H-014': 20,   // PRECON 80M ZTT
  'H-015': 10    // PRECON 200M
};

function doGet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const num = v => Number(v) || 0;
  const isDate = v => Object.prototype.toString.call(v) === '[object Date]';
  const fmt = v => isDate(v) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').trim();

  // Sheet "Rekap Barang": header di baris 4, data mulai baris 5 (kolom A-G)
  const rs = ss.getSheetByName('Rekap Barang');
  const rekap = rs.getRange(5, 1, Math.max(rs.getLastRow() - 4, 1), 7).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => {
      const kode = String(r[0]).trim();
      return {
        kode: kode, nama: String(r[1]).trim(),
        awal: num(r[2]), masuk: num(r[3]), keluar: num(r[4]), rusak: num(r[5]), akhir: num(r[6]),
        limit: LIMITS.hasOwnProperty(kode) ? LIMITS[kode] : null
      };
    });

  // Sheet "Barang Rusak": header di baris 2, data mulai baris 3
  const bs = ss.getSheetByName('Barang Rusak');
  const rusak = bs.getRange(3, 1, Math.max(bs.getLastRow() - 2, 1), 6).getValues()
    .filter(r => String(r[1]).trim() !== '')
    .map(r => ({
      tanggal: fmt(r[0]), kode: String(r[1]).trim(), nama: String(r[2]).trim(),
      jumlah: num(r[3]), ket: fmt(r[4]), ret: fmt(r[5])
    }));

  const out = { updated: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'), rekap: rekap, rusak: rusak };
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
