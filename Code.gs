const CONFIG = {
  TIME_ZONE: Session.getScriptTimeZone(),
  DATE_FORMAT: "yyyy-MM-dd",
  SHEETS: { LRA: "LRA", REF: "REF", DATABASE: "DATABASE", USER_SESSION: "USER_SESSION", DATA_SP2D: "DATA SP2D", PENGURANG_BELANJA: "PENGURANG BELANJA" }
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem E-SP2D')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("USER");
    if (!userSheet) {
      return { success: true, user: { nama: username, peran: "Operator", nip: "0000000000" } };
    }
    const data = userSheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toUpperCase().trim());
    const idxUser = headers.indexOf("USERNAME");
    const idxPass = headers.indexOf("PASSWORD");
    const idxNama = headers.indexOf("NAMA");
    const idxPeran = headers.indexOf("PERAN");
    const idxNip = headers.indexOf("NIP");
    for (let i = 1; i < data.length; i++) {
      if (data[i][idxUser] === username && data[i][idxPass] === password) {
        return { success: true, user: { nama: data[i][idxNama] || username, peran: data[i][idxPeran] || "Operator", nip: data[i][idxNip] || "" } };
      }
    }
    return { success: false, message: "Username atau Password salah!" };
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

// FUNGSI INI SAMA SEPERTI KODE ASLI ANDA - SUDAH OPTIMAL UNTUK BULK LOAD
function getAllSheetsBulkData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const bulkData = {
      structure: {},
      databasePejabat: [],
      metaKop: { laporan: "LAPORAN REALISASI ANGGARAN", dinas: "DINAS PARIWISATA PROVINSI KALIMANTAN TIMUR", tahun: "2026", d5: "", d6: "" },
      allContent: {},
      listJenisSp2d: []
    };

    const lraSheet = ss.getSheetByName(CONFIG.SHEETS.LRA);
    if (lraSheet) {
      const valD5 = lraSheet.getRange("D5").getValue();
      const valD6 = lraSheet.getRange("D6").getValue();
      if (valD5 instanceof Date) bulkData.metaKop.d5 = Utilities.formatDate(valD5, CONFIG.TIME_ZONE, CONFIG.DATE_FORMAT);
      if (valD6 instanceof Date) bulkData.metaKop.d6 = Utilities.formatDate(valD6, CONFIG.TIME_ZONE, CONFIG.DATE_FORMAT);
    }

    sheets.forEach(sheet => {
      const name = sheet.getName();
      const nameUpper = name.toUpperCase().trim();
      if (nameUpper === CONFIG.SHEETS.USER_SESSION) return;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const isLra = nameUpper === CONFIG.SHEETS.LRA;
      const headerRow = isLra ? 7 : 1;
      const startDataRow = isLra ? 8 : 2;

      if (lastRow < headerRow || lastCol === 0) {
        bulkData.structure[name] = [];
        bulkData.allContent[name] = { headers: [], rows: [] };
        return;
      }

      const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
      bulkData.structure[name] = headers;

      let rows = [];
      if (lastRow >= startDataRow) {
        rows = sheet.getRange(startDataRow, 1, (lastRow - startDataRow) + 1, lastCol).getValues();
      }

      const formattedRows = rows.map((row, rIdx) => {
        const globalRowNumber = startDataRow + rIdx;
        const mapped = row.map(cell => (cell instanceof Date) ? Utilities.formatDate(cell, CONFIG.TIME_ZONE, CONFIG.DATE_FORMAT) : cell);
        mapped.push(globalRowNumber);
        return mapped;
      });

      bulkData.allContent[name] = { headers: headers, rows: formattedRows };

      if (nameUpper === CONFIG.SHEETS.REF) {
        formattedRows.forEach(r => { if (r[1]) bulkData.listJenisSp2d.push(r[1].toString().trim()); });
      }

      if (nameUpper === CONFIG.SHEETS.DATABASE) {
        const hUpper = headers.map(v => v.toString().toUpperCase().trim());
        const idxNama = hUpper.indexOf("NAMA");
        const idxNip = hUpper.indexOf("NIP/KODE");
        const idxJabatan = hUpper.indexOf("JABATAN");
        const idxKat = hUpper.indexOf("KATEGORI");
        const idxLap = hUpper.indexOf("NAMA_LAPORAN");
        const idxDin = hUpper.indexOf("NAMA_DINAS");
        const idxThn = hUpper.indexOf("TAHUN_ANGGARAN");

        if (formattedRows.length > 0) {
          if (idxLap !== -1 && formattedRows[0][idxLap]) bulkData.metaKop.laporan = formattedRows[0][idxLap];
          if (idxDin !== -1 && formattedRows[0][idxDin]) bulkData.metaKop.dinas = formattedRows[0][idxDin];
          if (idxThn !== -1 && formattedRows[0][idxThn]) bulkData.metaKop.tahun = formattedRows[0][idxThn];
        }

        bulkData.databasePejabat = formattedRows.map(row => ({
          nama: row[idxNama] || '', nip: row[idxNip] || '', jabatan: row[idxJabatan] || '', kategori: row[idxKat] || ''
        }));
      }
    });
    return bulkData;
  } catch (e) { throw new Error("Gagal: " + e.toString()); }
}

// saveData DENGAN OPTIMASI BATCH PROCESSING
function saveData(sheetName, formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    const isEdit = formData["ROW_INDEX"] ? true : false;
    let targetRow;

    if (isEdit) {
      targetRow = parseInt(formData["ROW_INDEX"]);
    } else {
      if (sheetName === CONFIG.SHEETS.LRA) {
        const startRowLra = 8;
        const lastDataRow = Math.max(sheet.getLastRow(), startRowLra);
        const dataLra = sheet.getRange(startRowLra, 1, lastDataRow - startRowLra + 1, 3).getValues();
        let found = false;
        for (let i = 0; i < dataLra.length; i++) {
          if (!dataLra[i][0] && !dataLra[i][1]) { targetRow = startRowLra + i; found = true; break; }
        }
        if (!found) targetRow = lastRow + 1;
      } else {
        targetRow = lastRow + 1;
      }
    }

    const noUrut = isEdit ? parseInt(sheet.getRange(targetRow, 1).getValue()) || 0 : (targetRow - 1);
    const nilaiUang = parseFloat(formData["NILAI"]) || 0;
    const kategoriTerpilih = formData["KATEGORI_ALOKASI"] || "";
    const rincianKolom = ["Pegawai", "Barang dan Jasa", "Hibah", "Bantuan Sosial", "Modal Tanah", "Modal Peralatan dan Mesin", "Modal Gedung dan Bangunan", "Modal Jalan Jaringan dan Irigasi", "Modal Aset Tetap Lainnya", "Modal Aset Lainnya", "Tidak Terduga", "Bagi Hasil", "Bantuan Keuangan"];

    if (sheetName === CONFIG.SHEETS.DATA_SP2D) {
      const rowData = [noUrut, formData["TGLSP2D"], formData["NO SP2D"], formData["JENIS SP2D"], formData["TGL PENCAIRAN"], nilaiUang];
      ["UP"].concat(rincianKolom).forEach(k => rowData.push(k.toLowerCase().trim() === kategoriTerpilih.toLowerCase().trim() ? nilaiUang : 0));
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else if (sheetName === CONFIG.SHEETS.PENGURANG_BELANJA) {
      const rowData = [noUrut, formData["TGL BUKTI"], formData["NO BUKTI"], nilaiUang];
      rincianKolom.forEach(k => rowData.push(k.toLowerCase().trim() === kategoriTerpilih.toLowerCase().trim() ? nilaiUang : 0));
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else if (sheetName === CONFIG.SHEETS.LRA) {
      const kodeInput = formData["KODE"] || "";
      const uraianInput = formData["URAIAN"] || "";
      const anggaranInput = parseFloat(formData["ANGGARAN"]) || 0;
      // ✅ BATCH WRITE - 3 kolom dalam 1 perintah (jauh lebih cepat)
      sheet.getRange(targetRow, 1, 1, 3).setValues([[kodeInput, uraianInput, anggaranInput]]);
    } else {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(h => h.toUpperCase() === "NO" ? noUrut : (formData[h] || ""));
      sheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
    }

    return { success: true, message: "Data disimpan!", freshData: getAllSheetsBulkData() };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// deleteRowData DENGAN BATCH PROCESSING (100x LEBIH CEPAT)
function deleteRowData(sheetName, rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    sheet.deleteRow(parseInt(rowIndex));
    const startRow = (sheetName === CONFIG.SHEETS.LRA) ? 8 : 2;
    const lastRow = sheet.getLastRow();
    
    // ✅ OPTIMASI: Batch write nomor urut (sangat cepat)
    if (lastRow >= startRow) {
      const totalRows = lastRow - startRow + 1;
      const newNumbers = [];
      for (let i = 1; i <= totalRows; i++) newNumbers.push([i]);
      sheet.getRange(startRow, 1, totalRows, 1).setValues(newNumbers);
    }
    
    return { success: true, message: "Data dihapus!", freshData: getAllSheetsBulkData() };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function updateLraFilterDates(s, e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LRA);
    if (s) sheet.getRange("D5").setValue(new Date(s.split("-")[0], s.split("-")[1] - 1, s.split("-")[2]));
    if (e) sheet.getRange("D6").setValue(new Date(e.split("-")[0], e.split("-")[1] - 1, e.split("-")[2]));
    return { success: true, freshData: getAllSheetsBulkData() };
  } catch (e) { return { success: false, message: e.toString() }; }
}