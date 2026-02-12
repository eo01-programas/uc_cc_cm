import {
  CASE_CONTENT_COLUMNS,
} from "./case_content.js";
import {
  groupBy,
  fileToDataUrl,
  getFileExtension,
} from "./input.js";

function applyCaseContentHeaderLayout(ws) {
  const colWidths = [
    7, 13, 17, 12, 22.21875, 13, 12, 18, 13, 24, 16, 15,
  ];
  colWidths.forEach((width, idx) => {
    ws.getColumn(idx + 1).width = width;
  });

  const rowHeights = {
    2: 24,
    3: 24,
    4: 24,
    5: 28.05,
    6: 28.05,
    7: 28.05,
    8: 28.05,
    9: 28.05,
  };
  Object.entries(rowHeights).forEach(([row, height]) => {
    ws.getRow(Number(row)).height = height;
  });

  ws.mergeCells("D2:G2");
  ws.mergeCells("E4:F4");
  ws.mergeCells("I5:L5");
  ws.mergeCells("I6:K6");
  ws.mergeCells("I7:K7");

  ws.getCell("A2").value = "CASE CONTENT STICKER PARA CAJA.";
  ws.getCell("A2").font = { bold: true, size: 11 };
  ws.getCell("A2").alignment = { vertical: "middle" };

  ws.getCell("D2").value = "LAYOUT REFERENCIAL";
  ws.getCell("D2").font = { bold: true, size: 11 };
  ws.getCell("D2").alignment = { horizontal: "center", vertical: "middle" };

  ws.getCell("A3").value = "SÓLO DESTINO USA";
  ws.getCell("A3").font = { bold: true, size: 11 };
  ws.getCell("A3").alignment = { vertical: "middle" };

  ws.getCell("E4").value = "ANCHO: 9.8 CM";
  ws.getCell("E4").font = { bold: true, size: 11 };
  ws.getCell("E4").alignment = { horizontal: "center", vertical: "middle" };

  ws.getCell("A5").value = "USAR STICKER LANCO COD: 327984";
  ws.getCell("A5").font = { size: 11 };

  ws.getCell("I5").value = "INFORMACIÓN FIJA";
  ws.getCell("I5").font = { bold: true, size: 10 };
  ws.getCell("I5").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  ws.getCell("A6").value = "PARA IMPRIMIR TAMAÑO (9.8CM X 7.9 CM)";
  ws.getCell("A6").font = { size: 11 };

  ws.getCell("I6").value = "PO# BARCODE (POR ESTILO)";
  ws.getCell("I6").font = { bold: true, size: 10 };
  ws.getCell("I6").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  ws.getCell("G7").value = "ALTO:7.9 CM";
  ws.getCell("G7").font = { bold: true, size: 11 };
  ws.getCell("G7").alignment = { horizontal: "center", vertical: "middle" };

  ws.getCell("I7").value = "QTY BARCODE (CANT X CAJA/ESTILO)";
  ws.getCell("I7").font = { bold: true, size: 10 };
  ws.getCell("I7").alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  ws.views = [{ showGridLines: false }];
}

function applyCaseContentImage(workbook, ws, imageFile) {
  if (!imageFile) return;
  return fileToDataUrl(imageFile).then((dataUrl) => {
    const extension = getFileExtension(imageFile);
    const imageId = workbook.addImage({ base64: dataUrl, extension });
    ws.addImage(imageId, {
      tl: { col: 4, row: 4 },
      ext: { width: 227, height: 166 },
    });
  });
}

function setHeaderRow(ws, rowNumber) {
  CASE_CONTENT_COLUMNS.forEach((colName, idx) => {
    const cell = ws.getCell(rowNumber, idx + 1);
    cell.value = colName;
    cell.font = { bold: true, color: { argb: "000000" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF00" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
}

function applyColorRow(ws, rowNumber) {
  ws.getCell(rowNumber, 6).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "00B050" } };
  ws.mergeCells(`G${rowNumber}:H${rowNumber}`);
  ws.getCell(rowNumber, 7).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4472C4" } };
  ws.getCell(rowNumber, 9).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC000" } };
  ws.getCell(rowNumber, 10).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "7030A0" } };
  ws.getCell(6, 12).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "00B050" } };
}

function writeCaseContentData(ws, rows) {
  const headerRow = 10;
  const emptyRow = 11;
  const dataStartRow = 12;

  setHeaderRow(ws, headerRow);
  applyColorRow(ws, emptyRow);

  const textColumns = new Set(["SKX PO#", "STYLE/COLOR", "Case QTY", "US Size", "UPC Barcode"]);

  rows.forEach((row, index) => {
    const rowNumber = dataStartRow + index;
    CASE_CONTENT_COLUMNS.forEach((colName, colIndex) => {
      const cell = ws.getCell(rowNumber, colIndex + 1);
      const value = row[colName] ?? "";
      if (textColumns.has(colName)) {
        cell.value = String(value ?? "");
        cell.numFmt = "@";
      } else if (value !== "" && value !== null && value !== undefined) {
        const cleaned = String(value).trim();
        if (cleaned && !Number.isNaN(Number(cleaned))) {
          const num = Number(cleaned);
          cell.value = Number.isInteger(num) ? num : num;
        } else {
          cell.value = value;
        }
      } else {
        cell.value = value;
      }
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { bold: false };
    });
  });

  const lastDataRow = dataStartRow + rows.length - 1;
  const colCase = CASE_CONTENT_COLUMNS.indexOf("Case QTY") + 1;
  const colQty = CASE_CONTENT_COLUMNS.indexOf("QTY POR TALLA") + 1;
  const colResult = CASE_CONTENT_COLUMNS.indexOf("QTY DE STICKERS A IMPRIMIR") + 1;
  const caseLetter = colCase ? ws.getColumn(colCase).letter : null;
  const qtyLetter = colQty ? ws.getColumn(colQty).letter : null;

  if (caseLetter && qtyLetter && colResult && rows.length) {
    for (let r = dataStartRow; r <= lastDataRow; r += 1) {
      const cell = ws.getCell(r, colResult);
      cell.value = {
        formula: `IFERROR(ROUNDUP(${qtyLetter}${r}/VALUE(SUBSTITUTE(${caseLetter}${r},"Q","")),0)+3,0)`,
        result: 0,
      };
      cell.numFmt = "0";
      cell.alignment = { horizontal: "center", vertical: "middle" };
    }
  }

  const noteHeaderRow = (rows.length ? lastDataRow : headerRow) + 2;
  const noteTextRow = noteHeaderRow + 1;
  const noteHeaderCell = ws.getCell(noteHeaderRow, 1);
  noteHeaderCell.value = "Important notes:";
  noteHeaderCell.font = { bold: true, color: { argb: "000000" } };
  noteHeaderCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF00" } };

  const noteTextCell = ws.getCell(noteTextRow, 1);
  noteTextCell.value = "SKX PO#:, WIP Line #, UPC Barcode, CASE Qty:  Barcode format= code 128 ( con PO# 73 Barras)";
  noteTextCell.font = { size: 10 };
}

export async function generateCaseContentWorkbook(rows, imageFile) {
  const workbook = new ExcelJS.Workbook();
  const grouped = groupBy(rows, (row) => String(row["NOMBRE ESTILO"] ?? "").trim() || "REPORTE");
  const styleNames = Array.from(grouped.keys()).sort((a, b) => a.localeCompare(b, "es"));

  for (const styleName of styleNames) {
    const ws = workbook.addWorksheet(styleName.slice(0, 31) || "REPORTE");
    applyCaseContentHeaderLayout(ws);
    await applyCaseContentImage(workbook, ws, imageFile);

    const styleRows = grouped.get(styleName) || [];
    writeCaseContentData(ws, styleRows);
  }

  return workbook;
}
