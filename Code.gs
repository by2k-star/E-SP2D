// ==================================================
// KONFIGURASI GLOBAL
// ==================================================
const CONFIG = {
  TIME_ZONE: Session.getScriptTimeZone(),
  DATE_FORMAT: "yyyy-MM-dd",
  SHEETS: {
    LRA: "LRA",
    REF: "REF",
    DATABASE: "DATABASE",
    USER_SESSION: "USER_SESSION",
    DATA_SP2D: "DATA SP2D",
    PENGURANG_BELANJA: "PENGURANG BELANJA"
  }
};

// ==================================================
// WEB APP ENTRY POINT
// ==================================================
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem E-SP2D - Dinas Pariwisata')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================================================
// FUNGSI LOGIN
// ==================================================
function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("USER");
    
    // Mode Dummy jika sheet USER belum ada
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
        return { 
          success: true, 
          user: { 
            nama: data[i][idxNama] || username, 
            peran: data[i][idxPeran] || "Operator",
            nip: data[i][idxNip] || ""
          } 
        };
      }
    }
    return { success: false, message: "Username atau Password salah!" };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==================================================
// FUNGSI BARU: AMBIL DATA RINGAN UNTUK DASHBOARD AWAL
// ==================================================
function getDashboardEssentials() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    // Ambil daftar nama sheet untuk Sidebar
    const sheetNames = sheets.map(s => s.getName()).filter(n => n.toUpperCase() !== CONFIG.SHEETS.USER_SESSION);

    const essentials = {
      sheetNames: sheetNames, // PENTING: Untuk membangun menu sidebar
      metaKop: { laporan: "LAPORAN REALISASI ANGGARAN", dinas: "DINAS PARIWISATA", tahun: "2026", d5: "", d6: "" },
      databasePejabat: [],
      listJenisSp2d: [],
      structure: {} // Hanya berisi header (ringan)
    };

    // 1. Ambil Tanggal & Struktur dari setiap sheet
    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (name.toUpperCase() === CONFIG.SHEETS.USER_SESSION) return;
      
      const lastCol = sheet.getLastColumn();
      if (lastCol > 0) {
        const isLra = name.toUpperCase() === CONFIG.SHEETS.LRA;
        const headerRow = isLra ? 7 : 1;
        if (sheet.getLastRow() >= headerRow) {
           essentials.structure[name] = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
        }
      }

      if (name.toUpperCase() === CONFIG.SHEETS.LRA) {
        const valD5 = sheet.getRange("D5").getValue();
        const valD6 = sheet.getRange("D6").getValue();
        if (valD5 instanceof Date) essentials.metaKop.d5 = Utilities.formatDate(valD5, CONFIG.TIME_ZONE, CONFIG.DATE_FORMAT);
        if (valD6 instanceof Date) essentials.metaKop.d6 = Utilities.formatDate(valD6, CONFIG.TIME_ZONE, CONFIG.DATE_FORMAT);
      }
    });

    // 2. Ambil Data Pejabat dari DATABASE
    const dbSheet = ss.getSheetByName(CONFIG.SHEETS.DATABASE);
    if (dbSheet && dbSheet.getLastRow() > 1) {
      const data = dbSheet.getDataRange().getValues();
      const headers = data[0].map(v => v.toString().toUpperCase().trim());
      const idxNama = headers.indexOf("NAMA");
      const idxNip = headers.indexOf("NIP/KODE");
      const idxJabatan = headers.indexOf("JABATAN");
      const idxKat = headers.indexOf("KATEGORI");
      const idxLap = headers.indexOf("NAMA_LAPORAN");
      const idxDin = headers.indexOf("NAMA_DINAS");
      const idxThn = headers.indexOf("TAHUN_ANGGARAN");

      if (data.length > 1) {
        if (idxLap !== -1 && data[1][idxLap]) essentials.metaKop.laporan = data[1][idxLap];
        if (idxDin !== -1 && data[1][idxDin]) essentials.metaKop.dinas = data[1][idxDin];
        if (idxThn !== -1 && data[1][idxThn]) essentials.metaKop.tahun = data[1][idxThn];
      }
      essentials.databasePejabat = data.slice(1).map(row => ({
        nama: row[idxNama] || '', nip: row[idxNip] || '', jabatan: row[idxJabatan] || '', kategori: row[idxKat] || ''
      }));
    }

    // 3. Ambil Jenis SP2D dari REF
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF);
    if (refSheet && refSheet.getLastRow() > 1) {
      const refData = refSheet.getRange(2, 1, refSheet.getLastRow() - 1, 2).getValues();
      refData.forEach(r => { if (r[1]) essentials.listJenisSp2d.push(r[1].toString().trim()); });
    }

    return { success: true, data: essentials };
  } catch (e) {
    return { success: false, message: "Gagal memuat essentials: " + e.toString() };
  }
}

// ==================================================
// FUNGSI BARU: AMBIL DATA PER SHEET (LAZY LOADING)
// ==================================================
function getDataBySheetName(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) return { success: true, data: { headers: [], rows: [] } };

    const isLra = sheetName.toUpperCase() === CONFIG.SHEETS.LRA;
    const headerRow = isLra ? 7 : 1;
    const startDataRow = isLra ? 8 : 2;

    const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    let rows = [];
    if (lastRow >= startDataRow) {
      rows = sheet.getRange(startDataRow, 1, (lastRow - startDataRow) + 1, lastCol).getValues();
    }

    const formattedRows = rows.map((row, rIdx) => {
      const mapped = row.map(cell => (cell instanceof Date) ? Utilities.formatDate(cell, CONFIG.TIME_ZONE, CONFIG.DATE_FORMAT) : cell);
      mapped.push(startDataRow + rIdx); // Row Index
      return mapped;
    });

    return { success: true, data: { headers: headers, rows: formattedRows } };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==================================================
// FUNGSI BARU: AMBIL PAKET DATA KHUSUS LRA (LRA + SP2D + PENGURANG)
// ==================================================
function getLraDependencies() {
  try {
    const result = {};
    const sheetsNeeded = [CONFIG.SHEETS.LRA, CONFIG.SHEETS.DATA_SP2D, CONFIG.SHEETS.PENGURANG_BELANJA];
    
    sheetsNeeded.forEach(name => {
      const res = getDataBySheetName(name);
      if (res.success) {
        result[name] = res.data;
      }
    });
    
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==================================================
// FUNGSI SIMPAN DATA (OPTIMIZED)
// ==================================================
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
      sheet.getRange(targetRow, 1, 1, 3).setValues([[kodeInput, uraianInput, anggaranInput]]);
    } else {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(h => h.toUpperCase() === "NO" ? noUrut : (formData[h] || ""));
      sheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
    }

    // ✅ KEMBALIKAN HANYA DATA SHEET INI (BUKAN SEMUA)
    return { success: true, message: "Data disimpan!", freshData: getDataBySheetName(sheetName) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==================================================
// FUNGSI HAPUS DATA (OPTIMIZED)
// ==================================================
function deleteRowData(sheetName, rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    sheet.deleteRow(parseInt(rowIndex));
    
    const startRow = (sheetName === CONFIG.SHEETS.LRA) ? 8 : 2;
    const lastRow = sheet.getLastRow();
    
    if (lastRow >= startRow) {
      const totalRows = lastRow - startRow + 1;
      const newNumbers = [];
      for (let i = 1; i <= totalRows; i++) newNumbers.push([i]);
      sheet.getRange(startRow, 1, totalRows, 1).setValues(newNumbers);
    }

    // ✅ KEMBALIKAN HANYA DATA SHEET INI
    return { success: true, message: "Data dihapus!", freshData: getDataBySheetName(sheetName) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==================================================
// FUNGSI UPDATE FILTER TANGGAL
// ==================================================
function updateLraFilterDates(s, e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LRA);
    if(s) sheet.getRange("D5").setValue(new Date(s.split("-")[0], s.split("-")[1]-1, s.split("-")[2]));
    if(e) sheet.getRange("D6").setValue(new Date(e.split("-")[0], e.split("-")[1]-1, s.split("-")[2]));
    
    // ✅ KEMBALIKAN PAKET LRA (LRA + SP2D + PENGURANG)
    return { success: true, message: "Filter diterapkan!", freshData: getLraDependencies() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}