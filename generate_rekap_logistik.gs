function generateRekapLogistik() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const manifestSheet = ss.getSheetByName("MANIFEST");
  const manifestData = manifestSheet.getDataRange().getValues();

  const header = manifestData[0];
  const tujuanIdx = header.indexOf("Pel. Bongkar");
  const shipperIdx = header.indexOf("Shipper");
  const quantityIdx = header.indexOf("Quantity");
  const uangTambangIdx = header.indexOf("Uang Tambang");
  const voyageIdx = header.indexOf("Voyage");

  const data = manifestData.slice(1).filter(row => row[tujuanIdx] && row[shipperIdx]);

  const kapalIdx = header.indexOf("Kapal");
  const kapalNama = manifestData[1][kapalIdx] || "";
  const kapalNomor = kapalNama.match(/\d+/g)?.slice(-1)[0] || "-";  // Default "3" jika gagal

  // Ambil semua tujuan unik dalam urutan kemunculan
  const tujuanSet = new Set();
  data.forEach(row => {
    const tujuan = row[tujuanIdx];
    if (tujuan) tujuanSet.add(tujuan.toUpperCase());
  });
  const orderedTujuan = [...tujuanSet];  // Urutan kemunculan dipertahankan

  const normalizeName = name => {
    if (!name) return "";

    return name
      .toUpperCase()
      .replace(/\bP\s*\.?\s*T\b\.?/g, "PT.")  // Ubah semua variasi P T menjadi PT.
      .replace(/\bC\s*\.?\s*V\b\.?/g, "CV.")  // Ubah semua variasi C V menjadi CV.
      .replace(/\.+/g, ".")                   // Hapus titik ganda: ".." -> "."
      .replace(/\.([A-Z])/g, ". $1")          // Tambah spasi setelah titik sebelum huruf besar
      .replace(/\s{2,}/g, " ")                // Hilangkan spasi ganda
      .trim();
  };

  const shipperColorMap = {
    'CV. ALOIS GEMILANG': { bg: '#6495ED', font: '#ffffff', order: 1 },
    'CV. BANGSAHA JAYA': { bg: '#6495ED', font: '#ffffff', order: 1 },
    'PT. BINTARO JAYA LOGISTIK': { bg: '#6495ED', font: '#ffffff', order: 1 },
    'PT. PELITA TIGA PUTRA': { bg: '#6495ED', font: '#ffffff', order: 1 },
    'PT. BANGSAHA LOGISTIC MARITIM': { bg: '#6495ED', font: '#ffffff', order: 1 },
    'PT. EKASYA TANGGUH PRIMA': { bg: '#6495ED', font: '#ffffff', order: 1 },
    'CV. TAMAS BERKAH': { bg: '#6495ED', font: '#ffffff', order: 1 },
    'PT. SAMUDERA FALCINO ABAD': { bg: '#6495ED', font: '#ffffff', order: 1 },

    'PT. GENERASI ANAK PANAH': { bg: '#ffc000', font: '#000000', order: 2 },
    'CV. MANNA SEJAHTERA TRANS': { bg: '#ffc000', font: '#000000', order: 2 },
    'PT. SATU TUJUAN SEJAHTERA': { bg: '#ffc000', font: '#000000', order: 2 },
    'CV. GOSYEN SEJAHTERA ABADI': { bg: '#ffc000', font: '#000000', order: 2 },
    'PT. SAMUDERA INDONESIA TIMUR': { bg: '#ffc000', font: '#000000', order: 2 },
    'PT. SICEPAT LINTAS SAMUDRA': { bg: '#ffc000', font: '#000000', order: 2 },

    'CV. MAHA JAYA': { bg: '#bfbfbf', font: '#000000', order: 3 },
    'CV. MAKMUR SEJAHTERA': { bg: '#bfbfbf', font: '#000000', order: 3 },
    'PT. KEDIDI TRANS': { bg: '#bfbfbf', font: '#000000', order: 3 },
    'PT. SUNINDO TRANSJASA SEJAHTERA': { bg: '#ff0000', font: '#ffffff', order: 4 },
    'PT. BERDIKARI MITRA UTAMA': { bg: '#ff0000', font: '#ffffff', order: 4 },

    'CV. SURYA CEMERLANG GROUP': { bg: '#008b8b', font: '#ffffff', order: 5 },
    'PT. BERKAH WEDA LOGISTIK': { bg: '#008b8b', font: '#ffffff', order: 5 },
    'PT. ULTIMA TRANSPORTER INDONESIA': { bg: '#008b8b', font: '#ffffff', order: 5 },

    'PT. SARANA BANDAR LOGISTIK': { bg: '#00ffff', font: '#000000', order: 6 },
    'PT. SARANA BANDAR INDOTRADING SURABAYA': { bg: '#00ffff', font: '#000000', order: 6 }
  };

  const grouped = {};
  for (let row of data) {
    const tujuan = row[tujuanIdx];
    const rawShipper = row[shipperIdx];
    const shipper = normalizeName(rawShipper);
    const quantity = Number(row[quantityIdx]) || 0;
    const uangTambang = Number(row[uangTambangIdx]) || 0;
    const rawVoyage = row[voyageIdx];
    if (!rawVoyage) continue;

    const voyage = `VOY ${rawVoyage}`.toUpperCase();
    const penghasilanTambang = quantity * uangTambang;

    if (!grouped[voyage]) grouped[voyage] = {};
    if (!grouped[voyage][tujuan]) grouped[voyage][tujuan] = {};
    if (!grouped[voyage][tujuan][shipper]) grouped[voyage][tujuan][shipper] = { qty: 0, penghasilan: 0 };

    grouped[voyage][tujuan][shipper].qty += quantity;
    grouped[voyage][tujuan][shipper].penghasilan += penghasilanTambang;
  }

  for (let voyage in grouped) {
    const existingSheet = ss.getSheetByName(voyage);
    if (existingSheet) ss.deleteSheet(existingSheet);
    const sheet = ss.insertSheet(voyage);

    const tujuanMap = grouped[voyage];
    const tujuanList = orderedTujuan
      .filter(t => t in tujuanMap)
      .map(t => [t, tujuanMap[t]]);
    let colOffset = 1;

    for (let [tujuan, shipperMap] of tujuanList) {
      const requiredCols = colOffset + 3;
      const maxCols = sheet.getMaxColumns();
      if (maxCols < requiredCols) {
        sheet.insertColumnsAfter(maxCols, requiredCols - maxCols);
      }

      const sortedShippers = Object.entries(shipperMap).sort((a, b) => {
        const orderA = shipperColorMap[a[0]]?.order || 99;
        const orderB = shipperColorMap[b[0]]?.order || 99;
        if (orderA !== orderB) return orderA - orderB;
        return a[0].localeCompare(b[0]);
      });

      const title = `PEMBAGIAN ALOKASI LOGNUS ${kapalNomor} ${tujuan.toUpperCase()} ${voyage}`;
      const startRow = 1;

      // Title
      sheet.getRange(startRow, colOffset, 1, 4).merge()
        .setValue(title)
        .setFontWeight("bold")
        .setFontColor("#FFFFFF")
        .setBackground("#3B71D3")
        .setHorizontalAlignment("center")
        .setFontSize(11)
        .setBorder(true, true, true, true, true, true);

      // Header
      sheet.getRange(startRow + 1, colOffset, 1, 4).setValues([
        ["NO", "NAMA SHIPPER", "TERVALIDASI", "PENGHASILAN TAMBANG"]
      ])
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setFontColor("#FFFFFF")
        .setBackground("#3B71D3")
        .setBorder(true, true, true, true, true, true);

      // Data
      const dataRows = sortedShippers.map(([shipper, info], idx) => [
        idx + 1, shipper, info.qty, info.penghasilan
      ]);

      const dataStartRow = startRow + 2;
      if (dataRows.length > 0) {
        sheet.getRange(dataStartRow, colOffset, dataRows.length, 4).setValues(dataRows)
          .setHorizontalAlignment("center")
          .setBorder(true, true, true, true, true, true);

        // Format kolom "PENGHASILAN TAMBANG" (kolom ke-4)
        sheet.getRange(dataStartRow, colOffset + 3, dataRows.length, 1)
          .setNumberFormat('"Rp."#,##0')
          .setHorizontalAlignment("center");
      }

      dataRows.forEach((row, i) => {
        const shipper = row[1];
        const color = shipperColorMap[shipper] || { bg: "#FFFFFF", font: "#000000" };
        sheet.getRange(dataStartRow + i, colOffset, 1, 4).setBackground(color.bg).setFontColor(color.font);
      });

      const totalRow = dataStartRow + dataRows.length;
      sheet.getRange(totalRow, colOffset, 1, 2).merge().setValue("TOTAL")
        .setFontColor("#FFFFFF")
        .setBackground("#3B71D3")
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBorder(true, true, true, true, true, true);

      const totalQtyRange = sheet.getRange(dataStartRow, colOffset + 2, dataRows.length).getA1Notation();
      const totalUangRange = sheet.getRange(dataStartRow, colOffset + 3, dataRows.length).getA1Notation();

      sheet.getRange(totalRow, colOffset + 2).setFormula(`=SUM(${totalQtyRange})`)
        .setNumberFormat("#,##0")
        .setHorizontalAlignment("center")
        .setFontColor("#FFFFFF")
        .setBackground("#3B71D3")
        .setFontWeight("bold")
        .setBorder(true, true, true, true, true, true);

      sheet.getRange(totalRow, colOffset + 3).setFormula(`=SUM(${totalUangRange})`)
        .setNumberFormat('"Rp."#,##0')
        .setHorizontalAlignment("center")
        .setFontColor("#FFFFFF")
        .setBackground("#3B71D3")
        .setFontWeight("bold")
        .setBorder(true, true, true, true, true, true);

      // Tetapkan warna khusus hanya untuk kolom PENGHASILAN TAMBANG (kolom ke-4)
      sheet.getRange(dataStartRow, colOffset + 3, dataRows.length, 1)
        .setFontColor("#000000")
        .setBackground("#FFFFFF")
        .setNumberFormat('"Rp."#,##0');

      // Tetapkan warna khusus hanya untuk kolom TERVALIDASI (kolom ke-3)
      sheet.getRange(dataStartRow, colOffset + 2, dataRows.length, 1)
        .setFontColor("#000000")
        .setBackground("#FFFFFF")
        .setHorizontalAlignment("center")

      // Lebar kolom
      sheet.setColumnWidth(colOffset, 40);
      sheet.autoResizeColumn(colOffset + 1);
      sheet.autoResizeColumn(colOffset + 2);
      sheet.autoResizeColumn(colOffset + 3);

      colOffset += 5; // Tambahkan spasi antar blok
    }
  }
}
