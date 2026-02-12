import {
  SIZE_CANONICAL,
  SIZE_MAP,
  CANADA_SIZE_MAP,
  BRAZIL_SIZE_MAP,
  normSize,
  normalizeText,
  normalizeToken,
  removeAccents,
  hasQty,
  innerJoin,
  dedupeByKeys,
  sizeIndex,
  sortByKeys,
} from "./input.js";

function normCol(name) {
  const raw = removeAccents(name).toUpperCase();
  return raw.replace(/[^A-Z0-9#]+/g, " ").replace(/\s+/g, " ").trim();
}

function pickCol(colsNorm, ...candidates) {
  for (const cand of candidates) {
    if (colsNorm[cand]) return colsNorm[cand];
  }
  return null;
}

export function readUpcExcel(workbook) {
  const tokens = new Set([
    "STYLE",
    "ESTILOS",
    "OP",
    "RSV",
    "PROTO",
    "DESTINO",
    "PO",
    "PO NO",
    "PO NO.",
    "PO#",
    "DESCRIPCION COLOR",
    "DESCRIPCION DE COLOR",
    "COLOR",
    "CARTA",
    "CODE",
    "COLR CODE",
    "COLOR CODE",
    "LN",
  ]);

  let best = null;
  workbook.SheetNames.forEach((sheet) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {
      header: 1,
      defval: "",
      raw: false,
    });
    const limit = Math.min(40, rows.length);
    for (let i = 0; i < limit; i += 1) {
      const cells = rows[i].map((cell) => normalizeToken(cell));
      const hits = cells.filter((cell) => tokens.has(cell));
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

  return { sheet: fallbackSheet, rows: fallbackRows, headerRow: 1 };
}

export function prepareUpcExcel(rawObjects) {
  if (!rawObjects.length) {
    throw new Error("El Excel no contiene filas de datos.");
  }
  const colsNorm = {};
  Object.keys(rawObjects[0] || {}).forEach((col) => {
    const cleaned = String(col).replace(/__dup\d+$/, "");
    colsNorm[normCol(cleaned)] = col;
  });

  const renameMap = {};
  const styleCol = pickCol(colsNorm, "ESTILOS", "STYLE");
  if (styleCol) renameMap[styleCol] = "NOMBRE ESTILO";

  const opCol = pickCol(colsNorm, "OP");
  const rsvCol = pickCol(colsNorm, "RSV");
  if (opCol) renameMap[opCol] = "PEDIDO PRODUCCION COFACO";
  else if (rsvCol) renameMap[rsvCol] = "PEDIDO PRODUCCION COFACO";

  const protoCol = pickCol(colsNorm, "PROTO");
  if (protoCol) renameMap[protoCol] = "PROTO COFACO";

  const destinoCol = pickCol(colsNorm, "DESTINO");
  if (destinoCol) renameMap[destinoCol] = "DESTINO";

  const poCol = pickCol(colsNorm, "PO", "PO NO", "PO NO.", "PO#");
  if (poCol) renameMap[poCol] = "PO#";

  const lnCol = pickCol(colsNorm, "LN");
  if (lnCol) renameMap[lnCol] = "LN";

  let nombreColorCol = pickCol(colsNorm, "DESCRIPCION COLOR", "DESCRIPCION DE COLOR");
  if (!nombreColorCol) nombreColorCol = pickCol(colsNorm, "COLOR", "CARTA");
  if (nombreColorCol) renameMap[nombreColorCol] = "NOMBRE COLOR";

  let codigoColorCol = pickCol(colsNorm, "COLR CODE", "COLOR CODE");
  if (!codigoColorCol) codigoColorCol = pickCol(colsNorm, "CODE");
  if (!codigoColorCol) {
    const colorCol = colsNorm["COLOR"];
    if (colorCol && colorCol !== nombreColorCol) codigoColorCol = colorCol;
  }
  if (codigoColorCol) renameMap[codigoColorCol] = "COLOR";

  const rows = rawObjects.map((row) => {
    const out = {};
    Object.entries(row).forEach(([key, value]) => {
      const mapped = renameMap[key] || key;
      out[mapped] = value;
    });
    return out;
  });

  const required = [
    "NOMBRE ESTILO",
    "PEDIDO PRODUCCION COFACO",
    "PROTO COFACO",
    "DESTINO",
    "PO#",
    "NOMBRE COLOR",
  ];
  const missing = required.filter((col) => !(col in rows[0]));
  if (missing.length) {
    throw new Error(`Faltan columnas en el Excel: ${missing.join(", ")}`);
  }

  rows.forEach((row) => {
    if (!row.COLOR) row.COLOR = "";
    ["NOMBRE ESTILO", "DESTINO", "NOMBRE COLOR", "COLOR"].forEach((key) => {
      if (row[key] !== undefined) row[key] = normalizeText(row[key]);
    });
  });

  const columns = Object.keys(rows[0] || {});
  const sizeCols = columns.filter((col) => {
    const key = String(col).trim().toUpperCase();
    return SIZE_CANONICAL.has(key) || SIZE_MAP[key];
  });

  const idVars = [
    "NOMBRE ESTILO",
    "PEDIDO PRODUCCION COFACO",
    "PROTO COFACO",
    "DESTINO",
    "PO#",
    "NOMBRE COLOR",
    "COLOR",
  ];
  if ("LN" in rows[0]) idVars.push("LN");

  if (sizeCols.length) {
    const longRows = [];
    rows.forEach((row) => {
      sizeCols.forEach((sizeCol) => {
        if (!hasQty(row[sizeCol])) return;
        const sizeKey = String(sizeCol).trim().toUpperCase();
        const sizeValue = SIZE_MAP[sizeKey] || sizeKey;
        const base = {};
        idVars.forEach((key) => {
          base[key] = row[key] ?? "";
        });
        longRows.push({
          ...base,
          SIZE: normSize(sizeValue),
        });
      });
    });

    const sorted = sortByKeys(longRows, ["NOMBRE ESTILO", "DESTINO", "NOMBRE COLOR", "SIZE"], {
      SIZE: (a, b) => sizeIndex(a.SIZE) - sizeIndex(b.SIZE),
    });
    return sorted;
  }

  return rows.map((row) => {
    const out = {};
    idVars.forEach((key) => {
      out[key] = row[key] ?? "";
    });
    return out;
  });
}

export function buildUpcData(pdfRows, excelRows, options) {
  const pdfData = pdfRows.map((row) => {
    const out = { ...row };
    ["STYLE", "COLOR CODE", "COLOR NAME", "SIZE"].forEach((key) => {
      if (out[key] !== undefined) out[key] = normalizeText(out[key]);
    });
    if (out.SIZE) out.SIZE = normSize(out.SIZE);
    return out;
  });

  let excelData = excelRows.map((row) => ({ ...row }));
  excelData.forEach((row) => {
    if (row.SIZE) row.SIZE = normSize(row.SIZE);
  });

  if (options?.brazil) {
    excelData = excelData.filter(
      (row) => normalizeText(row.DESTINO ?? "") === "BRAZIL",
    );
    if (!excelData.length) {
      throw new Error("No se encontraron filas con DESTINO = BRAZIL en el Excel.");
    }
  }

  const excelHasSize = excelData.some((row) => "SIZE" in row);
  let dfName = [];
  let dfCode = [];

  if (excelHasSize) {
    dfName = innerJoin(
      excelData,
      pdfData,
      ["NOMBRE ESTILO", "NOMBRE COLOR", "SIZE"],
      ["STYLE", "COLOR NAME", "SIZE"],
    );
    dfCode = innerJoin(
      excelData,
      pdfData,
      ["NOMBRE ESTILO", "COLOR", "SIZE"],
      ["STYLE", "COLOR CODE", "SIZE"],
    );
  } else {
    dfName = innerJoin(
      excelData,
      pdfData,
      ["NOMBRE ESTILO", "NOMBRE COLOR"],
      ["STYLE", "COLOR NAME"],
    );
    dfCode = innerJoin(
      excelData,
      pdfData,
      ["NOMBRE ESTILO", "COLOR"],
      ["STYLE", "COLOR CODE"],
    );
  }

  const merged = dedupeByKeys(
    [...dfName, ...dfCode],
    [
      "NOMBRE ESTILO",
      "NOMBRE COLOR",
      "COLOR",
      "DESTINO",
      "PO#",
      "UPC CODE",
      "SIZE",
    ].filter((col) => dfName.concat(dfCode).some((row) => col in row)),
  );

  if (!merged.length) {
    throw new Error(
      "No hubo intersección entre PDFs y Excel con los criterios dados. Revisa STYLE/Color/Size.",
    );
  }

  if (options?.japan) {
    merged.forEach((row) => {
      const upc = String(row["UPC CODE"] ?? "");
      row["UPC CODE"] = upc.startsWith("0") ? upc : `0${upc}`;
    });
  }

  const columns = [
    "PROTO COFACO",
    "PEDIDO PRODUCCION COFACO",
    "DESTINO",
    "NOMBRE ESTILO",
    "NOMBRE COLOR",
    "PO#",
    "UPC CODE",
    "STYLE COLOR",
    "SIZE",
    "COLOR",
    "LN",
  ];

  merged.forEach((row) => {
    columns.forEach((col) => {
      if (!(col in row)) row[col] = "";
    });
  });

  let sorted = [];
  if (merged.some((row) => "SIZE" in row)) {
    sorted = sortByKeys(merged, [
      "PEDIDO PRODUCCION COFACO",
      "DESTINO",
      "PO#",
      "NOMBRE COLOR",
      "SIZE",
    ], {
      SIZE: (a, b) => sizeIndex(a.SIZE) - sizeIndex(b.SIZE),
    });
  } else {
    sorted = sortByKeys(merged, [
      "PEDIDO PRODUCCION COFACO",
      "DESTINO",
      "PO#",
      "NOMBRE COLOR",
    ]);
  }

  if (options?.canada) {
    sorted.forEach((row) => {
      if (row.SIZE) {
        const key = String(row.SIZE).toUpperCase().trim();
        row.SIZE = CANADA_SIZE_MAP[key] || row.SIZE;
      }
    });
  }

  if (options?.brazil) {
    sorted.forEach((row) => {
      if (row.SIZE) {
        const key = String(row.SIZE).toUpperCase().trim();
        row.SIZE = BRAZIL_SIZE_MAP[key] || row.SIZE;
      }
    });
  }

  const finalRows = sorted.map((row) => {
    const out = {};
    columns.forEach((col) => {
      out[col] = row[col] ?? "";
    });
    return out;
  });

  return { rows: finalRows, columns };
}

export const UPC_COLUMNS = [
  "PROTO COFACO",
  "PEDIDO PRODUCCION COFACO",
  "DESTINO",
  "NOMBRE ESTILO",
  "NOMBRE COLOR",
  "PO#",
  "UPC CODE",
  "STYLE COLOR",
  "SIZE",
  "COLOR",
  "LN",
];
