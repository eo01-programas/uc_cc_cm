import {
  SIZE_CANONICAL,
  SIZE_MAP,
  normSize,
  normalizeText,
  normalizeToken,
  forwardFill,
  hasQty,
  innerJoin,
  dedupeByKeys,
  sizeIndex,
  sortByKeys,
} from "./input.js";

const CASE_COL_MAP = {
  STYLE: "NOMBRE ESTILO",
  ESTILOS: "NOMBRE ESTILO",
  "NOMBRE ESTILO": "NOMBRE ESTILO",
  OP: "PEDIDO PRODUCCION COFACO",
  "PEDIDO PRODUCCION COFACO": "PEDIDO PRODUCCION COFACO",
  PROTO: "PROTO COFACO",
  "PROTO COFACO": "PROTO COFACO",
  DESTINO: "DESTINO",
  "PO#": "PO#",
  PO: "PO#",
  "PO NO": "PO#",
  "PO NO.": "PO#",
  "DESCRIPCION COLOR": "NOMBRE COLOR",
  "NOMBRE COLOR": "NOMBRE COLOR",
  COLOR: "COLOR",
  CARTA: "CARTA",
  "COLR CODE": "COLOR CODE",
  "COLOR CODE": "COLOR CODE",
  COLUMNA1: "HOJA MARCACION",
  "HOJA DE MARCACION": "HOJA MARCACION",
  "HOJA DE MARCACION": "HOJA MARCACION",
  TOTAL: "UNITS/TALLA (PEDIDO)",
  "UNITS/TALLA (PEDIDO)": "UNITS/TALLA (PEDIDO)",
  "SKX PO#": "SKX PO#",
  "SKX PO": "SKX PO#",
  "WIP LINE NUMBER": "WIP LINE NUMBER",
  "WIP LINE NUMBER:": "WIP LINE NUMBER",
  LN: "WIP LINE NUMBER",
  "CASE QTY": "CASE QTY",
  CASEQTY: "CASE QTY",
};

function renameCaseColumns(rows) {
  return rows.map((row) => {
    const out = {};
    Object.entries(row).forEach(([key, value]) => {
      const cleanedKey = String(key).replace(/__dup\d+$/, "");
      const normKey = normalizeToken(cleanedKey);
      const canonical = CASE_COL_MAP[normKey] || String(cleanedKey).trim();
      if (out[canonical] === undefined || String(out[canonical]).trim() === "") {
        out[canonical] = value;
      }
    });
    return out;
  });
}

function normalizeRequired(rows) {
  rows.forEach((row) => {
    ["NOMBRE ESTILO", "DESTINO", "NOMBRE COLOR", "COLOR"].forEach((key) => {
      if (row[key] !== undefined) row[key] = normalizeText(row[key]);
    });
    if (row.SIZE !== undefined) row.SIZE = normSize(row.SIZE);
  });
}

function detectSizeColumns(columns) {
  return columns.filter((col) => {
    const key = String(col).trim().toUpperCase();
    return SIZE_CANONICAL.has(key) || SIZE_MAP[key];
  });
}

export function readCaseContentExcel(workbook) {
  const headerTerms = new Set([
    "STYLE",
    "ESTILOS",
    "OP",
    "PROTO",
    "DESTINO",
    "PO",
    "PO NO",
    "PO NO.",
    "PO#",
    "DESCRIPCION COLOR",
    "COLOR",
    "CARTA",
    "CODE",
    "COLR CODE",
    "COLOR CODE",
  ]);

  let best = null;
  workbook.SheetNames.forEach((sheet) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {
      header: 1,
      defval: "",
      raw: false,
    });
    const limit = Math.min(30, rows.length);
    for (let i = 0; i < limit; i += 1) {
      const cells = rows[i].map((cell) => normalizeToken(cell));
      const hits = cells.filter((cell) => headerTerms.has(cell));
      if (cells.includes("STYLE") || cells.includes("ESTILOS")) {
        if (!best || hits.length > best.hits) {
          best = { sheet, row: i, hits: hits.length };
        }
      }
    }
  });

  if (best) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[best.sheet], {
      header: 1,
      defval: "",
      raw: false,
    });
    return { sheet: best.sheet, rows, headerRow: best.row };
  }

  const fallbackSheet = workbook.SheetNames[0];
  const fallbackRows = XLSX.utils.sheet_to_json(workbook.Sheets[fallbackSheet], {
    header: 1,
    defval: "",
    raw: false,
  });

  let bestHeaderRow = 0;
  let maxTextScore = 0;
  const limit = Math.min(5, fallbackRows.length);
  for (let i = 0; i < limit; i += 1) {
    const row = fallbackRows[i];
    let textScore = 0;
    row.forEach((val) => {
      const raw = String(val ?? "").trim();
      if (raw.length > 2 && !raw.replace(/[.,]/g, "").match(/^\d+$/)) {
        textScore += 1;
      }
    });
    if (textScore > maxTextScore) {
      maxTextScore = textScore;
      bestHeaderRow = i;
    }
  }

  return { sheet: fallbackSheet, rows: fallbackRows, headerRow: bestHeaderRow };
}

export function prepareCaseContentExcel(rawObjects) {
  if (!rawObjects.length) {
    throw new Error("El Excel no contiene filas de datos.");
  }
  let rows = renameCaseColumns(rawObjects);

  const groupCol =
    rows.length && ("NOMBRE ESTILO" in rows[0] || "STYLE" in rows[0] || "ESTILOS" in rows[0])
      ? ["NOMBRE ESTILO", "STYLE", "ESTILOS"].find((col) => col in rows[0])
      : null;

  ["CASE QTY", "WIP LINE NUMBER", "TT"].forEach((col) => {
    if (rows.length && col in rows[0]) forwardFill(rows, col, groupCol);
  });

  rows = rows.map((row) => {
    const out = { ...row };

    if (!out["NOMBRE COLOR"]) {
      if (out.COLOR) out["NOMBRE COLOR"] = out.COLOR;
      else if (out.CARTA) out["NOMBRE COLOR"] = out.CARTA;
    }

    if (!out.COLOR && out["COLOR CODE"]) out.COLOR = out["COLOR CODE"];

    if (out.COLOR && out["COLOR CODE"]) {
      out.COLOR = out["COLOR CODE"];
    }

    return out;
  });

  if (!rows.length) return [];

  if (!rows[0]["WIP LINE NUMBER"] && rows[0]["HOJA MARCACION"]) {
    const sample = rows
      .map((row) => String(row["HOJA MARCACION"] ?? "").trim())
      .filter(Boolean)
      .slice(0, 20);
    if (sample.length) {
      const withDigits = sample.filter((value) => /\d/.test(value)).length;
      if (withDigits / sample.length >= 0.4) {
        rows = rows.map((row) => ({
          ...row,
          "WIP LINE NUMBER": row["HOJA MARCACION"],
        }));
      }
    }
  }

  if (!rows[0]["UPC Barcode"]) {
    const candidates = ["UPC CODE", "UPC", "UPC_BARCODE", "UPC BARCODE"];
    rows = rows.map((row) => {
      const out = { ...row };
      if (!out["UPC Barcode"]) {
        const found = candidates.find((c) => out[c]);
        if (found) out["UPC Barcode"] = out[found];
      }
      return out;
    });
  }

  const required = ["NOMBRE ESTILO", "DESTINO", "PO#", "NOMBRE COLOR", "COLOR"];
  const missing = required.filter((col) => !(col in rows[0]));
  if (missing.length) {
    throw new Error(`Faltan columnas requeridas en el Excel: ${missing.join(", ")}`);
  }

  rows = rows.map((row) => ({
    ...row,
    "PEDIDO PRODUCCION COFACO": row["PEDIDO PRODUCCION COFACO"] ?? "",
    "PROTO COFACO": row["PROTO COFACO"] ?? "",
  }));

  normalizeRequired(rows);

  const columns = Object.keys(rows[0]);
  const sizeColumns = detectSizeColumns(columns);
  const idVarsBase = [
    "NOMBRE ESTILO",
    "PEDIDO PRODUCCION COFACO",
    "PROTO COFACO",
    "DESTINO",
    "PO#",
    "NOMBRE COLOR",
    "COLOR",
  ];
  const extras = [
    "HOJA MARCACION",
    "UNITS/TALLA (PEDIDO)",
    "WIP LINE NUMBER",
    "CASE QTY",
    "TT",
  ].filter((col) => columns.includes(col));

  if (sizeColumns.length) {
    const longRows = [];
    rows.forEach((row) => {
      sizeColumns.forEach((sizeCol) => {
        const qty = row[sizeCol];
        if (!hasQty(qty)) return;
        const sizeKey = String(sizeCol).trim().toUpperCase();
        longRows.push({
          ...Object.fromEntries(idVarsBase.concat(extras).map((col) => [col, row[col] ?? ""])),
          SIZE: SIZE_MAP[sizeKey] || sizeKey,
          "QTY POR TALLA": qty,
        });
      });
    });
    longRows.forEach((row) => {
      row.SIZE = normSize(row.SIZE);
    });
    return longRows;
  }

  const out = rows.map((row) => ({
    ...Object.fromEntries(idVarsBase.concat(extras).map((col) => [col, row[col] ?? ""])),
    SIZE: "",
    "QTY POR TALLA": "",
  }));

  return out;
}

export function buildCaseContentData(pdfRows, excelRows) {
  const pdfData = pdfRows.map((row) => {
    const out = { ...row };
    ["STYLE", "COLOR CODE", "COLOR NAME", "SIZE"].forEach((key) => {
      if (out[key] !== undefined) out[key] = normalizeText(out[key]);
    });
    if (out.SIZE) out.SIZE = normSize(out.SIZE);
    return out;
  });

  const excelData = excelRows.map((row) => ({ ...row }));
  normalizeRequired(excelData);
  excelData.forEach((row) => {
    if (row.SIZE) row.SIZE = normSize(row.SIZE);
  });

  const filteredExcel = excelData.filter((row) => row.DESTINO === "USA");
  if (!filteredExcel.length) {
    throw new Error("No se encontraron filas con DESTINO = USA en el Excel.");
  }

  const excelHasSize = filteredExcel.some((row) => "SIZE" in row);
  let dfName = [];
  let dfCode = [];

  if (excelHasSize) {
    dfName = innerJoin(
      filteredExcel,
      pdfData,
      ["NOMBRE ESTILO", "NOMBRE COLOR", "SIZE"],
      ["STYLE", "COLOR NAME", "SIZE"],
    );
    dfCode = innerJoin(
      filteredExcel,
      pdfData,
      ["NOMBRE ESTILO", "COLOR", "SIZE"],
      ["STYLE", "COLOR CODE", "SIZE"],
    );
  } else {
    dfName = innerJoin(
      filteredExcel,
      pdfData,
      ["NOMBRE ESTILO", "NOMBRE COLOR"],
      ["STYLE", "COLOR NAME"],
    );
    dfCode = innerJoin(
      filteredExcel,
      pdfData,
      ["NOMBRE ESTILO", "COLOR"],
      ["STYLE", "COLOR CODE"],
    );
  }

  const subsetCols = [
    "NOMBRE ESTILO",
    "NOMBRE COLOR",
    "COLOR",
    "DESTINO",
    "PO#",
    "SIZE",
    "UPC CODE",
  ].filter((col) => dfName.concat(dfCode).some((row) => col in row));

  const merged = dedupeByKeys([...dfName, ...dfCode], subsetCols);

  if (!merged.length) {
    throw new Error("No hubo intersección entre PDFs y Excel (DESTINO=USA).");
  }

  const finalRows = merged.map((row) => {
    const proto = row["PROTO COFACO"] ?? "";
    const op = row["PEDIDO PRODUCCION COFACO"] ?? "";
    const poCliente = row["PO#"] ?? "";
    const usSize = row.SIZE ?? "";
    const qtyTalla = row["QTY POR TALLA"] ?? "";
    const unitsPedido = row.TT ?? row["UNITS/TALLA (PEDIDO)"] ?? "";
    const rawWip = String(row["WIP LINE NUMBER"] ?? "").trim();
    const wipNum = parseInt(rawWip, 10);
    const wip = !isNaN(wipNum) ? `N${String(wipNum).padStart(2, "0")}` : rawWip;
    const styleColor = row["STYLE COLOR"] ?? "";
    const upc = row["UPC CODE"] ?? "";

    const skx = String(poCliente ?? "").trim();
    const skxFinal = skx
      ? skx.toUpperCase().startsWith("P")
        ? skx
        : `P${skx}`
      : "";

    let caseQty = "";
    if (row["CASE QTY"] !== undefined) {
      const digits = String(row["CASE QTY"] ?? "")
        .trim()
        .replace(/[^0-9]/g, "");
      if (digits) caseQty = `Q${digits}`;
    }

    return {
      "NOMBRE ESTILO": row["NOMBRE ESTILO"] ?? "",
      PROTO: proto,
      OP: op,
      "PO(cliente)": poCliente,
      "UNITS/TALLA(pedido)": unitsPedido,
      "SKX PO#": skxFinal,
      "WIP Line Number": wip,
      "STYLE/COLOR": styleColor,
      "UPC Barcode": upc,
      "Case QTY": caseQty,
      "US Size": usSize,
      "QTY POR TALLA": qtyTalla,
      "QTY DE STICKERS A IMPRIMIR": "",
    };
  });

  const sorted = sortByKeys(finalRows, ["SKX PO#", "WIP Line Number", "STYLE/COLOR", "US Size"], {
    "US Size": (a, b) => sizeIndex(a["US Size"]) - sizeIndex(b["US Size"]),
  });

  return sorted;
}

export const CASE_CONTENT_COLUMNS = [
  "PROTO",
  "OP",
  "PO(cliente)",
  "UNITS/TALLA(pedido)",
  "SKX PO#",
  "WIP Line Number",
  "STYLE/COLOR",
  "UPC Barcode",
  "Case QTY",
  "US Size",
  "QTY POR TALLA",
  "QTY DE STICKERS A IMPRIMIR",
];
