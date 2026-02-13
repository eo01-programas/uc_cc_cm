import { UPC_COLUMNS } from "./upc_sticker.js";
import { fileToDataUrl, getFileExtension, groupBy } from "./input.js";

function applyUpcHeaderLayout(ws, options) {
  const widths = {
    A: 12,
    B: 16,
    C: 20,
    D: 8,
    E: 16,
    F: 20,
    G: 16,
    H: 12,
    I: 20,
    J: 16,
    K: 30,
  };
  Object.entries(widths).forEach(([col, width]) => {
    ws.getColumn(col).width = width;
  });

  for (let r = 5; r <= 9; r += 1) {
    ws.getRow(r).height = 40;
  }

  ws.mergeCells("K6:K7");

  ws.getCell("B2").value = "UPC STICKER PARA HANG TAG Y BOLSA";
  ws.getCell("B2").font = { bold: true, size: 16 };
  ws.getCell("B2").alignment = { horizontal: "left" };

  ws.getCell("B3").value = "USAR STICKER EN BLANCO COD: 327986 PARA ";
  ws.getCell("B3").font = { size: 11 };

  ws.getCell("G3").value = "38.5 MM";
  ws.getCell("G3").font = { bold: true, size: 11 };

  ws.getCell("B4").value = "IMPRIMIR (38 MM X 38.5 MM)";
  ws.getCell("B4").font = { size: 11 };

  ws.getCell("K4").value = "INFORMACIÓN FIJA";
  ws.getCell("K4").alignment = { horizontal: "center" };
  ws.getCell("K4").font = { size: 11 };

  ws.getCell("I6").value = "38 MM";
  ws.getCell("I6").font = { bold: true, size: 11 };

  ws.getCell("A10").value =
    "*Para todos los destinos: USA e INTERNACIONALES el UPC sticker irá en el Hang Tag y en ";
  ws.getCell("A10").font = { bold: true, size: 11 };
  ws.getCell("A11").value = "la bolsa.";
  ws.getCell("A11").font = { bold: true, size: 11 };
  ws.getCell("A12").value =
    "*Adicionalmente solo para los destinos INTERNACIONALES irá el UPC sticker también en la caja. ";
  ws.getCell("A12").font = { bold: true, size: 11 };

  if (options?.brazil) {
    ws.getCell("K5").value = "MADE IN PERU FEITO NO PERU".toUpperCase();
    ws.getCell("K6").value = "SIZE/TAMANHO:".toUpperCase();
    ws.getCell("K8").value = "COLOR/COR:".toUpperCase();
  } else {
    ws.getCell("K5").value = "MADE IN PERU FABRIQUÉ AU PÉROU".toUpperCase();
    ws.getCell("K6").value = "SIZE/TAILLE:".toUpperCase();
    ws.getCell("K8").value = "COLOR/COULEUR:".toUpperCase();
  }

  const borderStyle = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  ws.getCell("K4").border = borderStyle;
  ws.getCell("K5").border = borderStyle;
  // K6 is merged with K7. Setting border on K6 affects the merged range usually.
  ws.getCell("K6").border = borderStyle;
  ws.getCell("K7").border = borderStyle; // Explicitly set for merged cells sometimes needed
  ws.getCell("K8").border = borderStyle;

  ws.getCell("K5").font = { bold: true, size: 10 };
  ws.getCell("K5").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  ws.getCell("K6").font = { bold: true, size: 10 };
  ws.getCell("K6").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  ws.getCell("K8").font = { bold: true, size: 10 };
  ws.getCell("K8").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  for (let r = 1; r <= 20; r += 1) {
    const cell = ws.getCell(r, 11);
    const current = cell.font || {};
    cell.font = { ...current, size: 16 };
  }

  ws.views = [{ showGridLines: false }];
}

async function applyUpcImages(workbook, ws, image1File, image2File) {
  const tasks = [];
  if (image1File) {
    tasks.push(
      fileToDataUrl(image1File).then((dataUrl) => {
        const extension = getFileExtension(image1File);
        const id = workbook.addImage({ base64: dataUrl, extension });
        ws.addImage(id, {
          tl: { col: 1, row: 4 }, // Row 5
          ext: { width: 5 * 37.7952755906, height: 5 * 37.7952755906 },
        });
      }),
    );
  }
  if (image2File) {
    tasks.push(
      fileToDataUrl(image2File).then((dataUrl) => {
        const extension = getFileExtension(image2File);
        const id = workbook.addImage({ base64: dataUrl, extension });

        // Positioned slightly higher (row 4 instead of row 5) 
        // to be between F4 and H4 presumably covering F4 area visually if needed.
        // Or if user means F4-H4 area, tl row 3 (Row 4) is correct.
        // User said: "La imagen debe surbir un poquito mas arriba debe estar entre F4 y H4"
        // Row 4 is index 3. Original was row 4 (index 4 = Row 5).
        // So moving to index 3 (Row 4) moves it up.

        ws.addImage(id, {
          tl: { col: 5, row: 3 }, // F4
          ext: { width: 6.46 * 37.7952755906, height: 6 * 37.7952755906 },
        });
      }),
    );
  }
  await Promise.all(tasks);
}

function writeUpcData(ws, rows) {
  const headerRow = 14;
  const dataStartRow = 15;

  UPC_COLUMNS.forEach((col, idx) => {
    const cell = ws.getCell(headerRow, idx + 1);
    cell.value = col;
    cell.font = { bold: true };
    cell.alignment = { horizontal: "center", vertical: "center", wrapText: true };
  });

  rows.forEach((row, index) => {
    const rowNumber = dataStartRow + index;
    UPC_COLUMNS.forEach((col, idx) => {
      const cell = ws.getCell(rowNumber, idx + 1);
      const value = row[col] ?? "";
      cell.value = String(value ?? "").replace(/^'/, "");
      cell.alignment = { horizontal: "center", vertical: "center" };
      cell.numFmt = "@";
    });
  });

  const maxRow = dataStartRow + rows.length - 1;
  const maxCol = UPC_COLUMNS.length;

  if (rows.length) {
    ws.autoFilter = {
      from: { row: headerRow, column: 1 },
      to: { row: maxRow, column: maxCol },
    };
  }



  for (let colIdx = 1; colIdx <= maxCol; colIdx += 1) {
    const col = ws.getColumn(colIdx);
    let maxLen = 0;
    for (let rowIdx = headerRow; rowIdx <= maxRow; rowIdx += 1) {
      const val = ws.getCell(rowIdx, colIdx).value;
      if (val !== null && val !== undefined) {
        const len = String(val).length;
        if (len > maxLen) maxLen = len;
      }
    }
    col.width = Math.min(maxLen + 2, 60);
  }

  ws.getColumn(1).width = 20;
  ws.getColumn(11).width = 30;
}

export async function generateUpcWorkbook(rows, image1File, image2File, options, existingWorkbook) {
  const workbook = existingWorkbook || new ExcelJS.Workbook();
  const hasStyle = rows.some((row) => String(row["ESTILO"] ?? "").trim() !== "");
  const grouped = hasStyle
    ? groupBy(rows, (row) => String(row["ESTILO"] ?? "").trim() || "REPORTE")
    : new Map([["REPORTE", rows]]);

  const styleNames = Array.from(grouped.keys()).sort((a, b) => a.localeCompare(b, "es"));
  const suffix = options?.sheetSuffix ? ` (${options.sheetSuffix})` : "";

  for (const styleName of styleNames) {
    // Excel sheet limit is 31 chars. Reserve space for suffix.
    const maxBaseLen = 31 - suffix.length;
    const baseName = styleName.slice(0, maxBaseLen) || "REPORTE";
    const sheetName = `${baseName}${suffix}`;

    // Ensure unique sheet name output if collision (though styleNames are unique, collision might happen after truncation)
    let finalSheetName = sheetName;
    let counter = 1;
    while (workbook.getWorksheet(finalSheetName)) {
      finalSheetName = `${sheetName.slice(0, 31 - String(counter).length)}${counter}`;
      counter++;
    }

    const ws = workbook.addWorksheet(finalSheetName);
    applyUpcHeaderLayout(ws, options);
    await applyUpcImages(workbook, ws, image1File, image2File);
    const dataRows = grouped.get(styleName) || [];
    writeUpcData(ws, dataRows);
  }

  const isDefault = !options?.japan && !options?.canada && !options?.brazil;
  const isRegular = options?.regular || isDefault;

  // Logic for Indonesia sheet (only if Regular or Default)
  // We should check if we want to include Indonesia sheet logic for every option or just regular.
  // Usually Indonesia sheet is part of the "Regular" process.
  // If user selects multiple, we might want to include it only for the Regular pass to avoid duplicates, 
  // OR include it with suffix for each pass if the rows differ (e.g. filtered).
  // Current logic: if (isRegular).

  if (isRegular) {
    const indonesiaRows = rows.filter(
      (row) => String(row.DESTINO ?? "").trim().toUpperCase() === "INDONESIA",
    );
    if (indonesiaRows.length) {
      const indoSheetName = `INDONESIA${suffix}`;
      const ws = workbook.addWorksheet(indoSheetName);
      applyUpcHeaderLayout(ws, options);
      await applyUpcImages(workbook, ws, image1File, image2File);
      writeUpcData(ws, indonesiaRows);
    }
  }

  return workbook;
}
