export const SIZE_MAP = {
  XXS: "XXS",
  XS: "XS",
  S: "S",
  M: "M",
  L: "L",
  XL: "XL",
  XXL: "2XL",
  "2XL": "2XL",
  XXXL: "3XL",
  "3XL": "3XL",
  XSS: "XS",
  SMALL: "S",
  MEDIUM: "M",
  LARGE: "L",
  "EXTRA SMALL": "XS",
  "EXTRA LARGE": "XL",
  "EXTRA EXTRA LARGE": "2XL",
  "EXTRA EXTRA EXTRA LARGE": "3XL",
  CHICO: "S",
  MEDIANO: "M",
  GRANDE: "L",
};

export const SIZE_ORDER = ["XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL"];
export const SIZE_CANONICAL = new Set(SIZE_ORDER);

export const CANADA_SIZE_MAP = {
  S: "S/P",
  M: "M/M",
  L: "L/G",
  XL: "XL/TG",
  "2XL": "2XL/TTG",
  "3XL": "3XL/TTTG",
};

export const BRAZIL_SIZE_MAP = {
  XS: "XS/PP",
  S: "S/P",
  M: "M/M",
  L: "L/G",
  XL: "XL/GG",
  "2XL": "XXL/XGG",
  XXL: "XXL/XGG",
};

export function normSize(value) {
  if (value === null || value === undefined) return value;
  const key = String(value).trim().toUpperCase();
  return SIZE_MAP[key] || key;
}

export function normalizeText(value) {
  return String(value ?? "").trim().toUpperCase();
}

export function removeAccents(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

export function normalizeToken(value) {
  const raw = removeAccents(value).toUpperCase();
  const cleaned = raw.replace(/[^A-Z0-9#]+/g, " ");
  return cleaned.replace(/\s+/g, " ").trim();
}

export function fileToArrayBuffer(file) {
  return file.arrayBuffer();
}

export function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

export function getFileExtension(file) {
  const name = file?.name || "";
  const parts = name.split(".");
  const ext = parts.length > 1 ? parts.pop().toLowerCase() : "png";
  if (ext === "jpg") return "jpeg";
  return ext;
}

export async function readWorkbook(file) {
  const buffer = await fileToArrayBuffer(file);
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

export function sheetToRows(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
}

export function rowsToObjects(rows, headerRow) {
  const header = rows[headerRow] || [];
  const seen = {};
  const columns = header.map((col, idx) => {
    const base = col ? String(col).trim() : `Col_${idx}`;
    if (!seen[base]) {
      seen[base] = 1;
      return base;
    }
    seen[base] += 1;
    return `${base}__dup${seen[base]}`;
  });
  const objects = [];
  for (let i = headerRow + 1; i < rows.length; i += 1) {
    const row = rows[i];
    const hasValue = row.some((cell) => String(cell ?? "").trim() !== "");
    if (!hasValue) continue;
    const obj = {};
    columns.forEach((col, idx) => {
      obj[col] = row[idx] ?? "";
    });
    objects.push(obj);
  }
  return { columns, objects };
}

export function innerJoin(leftRows, rightRows, leftKeys, rightKeys) {
  const rightMap = new Map();
  rightRows.forEach((row) => {
    const key = rightKeys.map((k) => String(row[k] ?? "")).join("||");
    if (!rightMap.has(key)) rightMap.set(key, []);
    rightMap.get(key).push(row);
  });

  const merged = [];
  leftRows.forEach((left) => {
    const key = leftKeys.map((k) => String(left[k] ?? "")).join("||");
    const matches = rightMap.get(key);
    if (!matches) return;
    matches.forEach((right) => {
      merged.push({ ...left, ...right });
    });
  });

  return merged;
}

export function dedupeByKeys(rows, keys) {
  const seen = new Set();
  const out = [];
  rows.forEach((row) => {
    const key = keys.map((k) => String(row[k] ?? "")).join("||");
    if (seen.has(key)) return;
    seen.add(key);
    out.push(row);
  });
  return out;
}

export function groupBy(rows, keyFn) {
  const map = new Map();
  rows.forEach((row) => {
    const key = keyFn(row);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  });
  return map;
}

export function sortByKeys(rows, keys, customComparators = {}) {
  const compareValue = (a, b, key) => {
    if (customComparators[key]) return customComparators[key](a, b);
    const av = String(a?.[key] ?? "");
    const bv = String(b?.[key] ?? "");
    return av.localeCompare(bv, "es", { numeric: true });
  };

  return [...rows].sort((a, b) => {
    for (const key of keys) {
      const res = compareValue(a, b, key);
      if (res !== 0) return res;
    }
    return 0;
  });
}

export function hasQty(value) {
  if (value === null || value === undefined) return false;
  const raw = String(value).trim();
  if (!raw || raw === "0" || raw === "0.0") return false;
  const num = Number(raw);
  if (Number.isNaN(num)) return true;
  return num > 0;
}

export function sizeIndex(value) {
  const key = String(value ?? "").toUpperCase();
  const idx = SIZE_ORDER.indexOf(key);
  return idx === -1 ? 999 : idx;
}

export function forwardFill(rows, column, groupColumn) {
  const lastByGroup = new Map();
  rows.forEach((row) => {
    const groupKey = groupColumn ? String(row[groupColumn] ?? "") : "__all__";
    const val = row[column];
    if (val !== null && val !== undefined && String(val).trim() !== "") {
      lastByGroup.set(groupKey, val);
    } else if (lastByGroup.has(groupKey)) {
      row[column] = lastByGroup.get(groupKey);
    }
  });
}

export async function extractPdfLines(file) {
  const data = new Uint8Array(await fileToArrayBuffer(file));
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const allLines = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent();
    const items = content.items
      .map((item) => ({
        str: item.str,
        x: item.transform[4],
        y: item.transform[5],
      }))
      .filter((item) => item.str && String(item.str).trim() !== "");

    items.sort((a, b) => {
      if (Math.abs(b.y - a.y) > 1) return b.y - a.y;
      return a.x - b.x;
    });

    let currentY = null;
    let line = [];
    const flushLine = () => {
      if (!line.length) return;
      const text = line.map((l) => l.str).join(" ").replace(/\s+/g, " ").trim();
      if (text) allLines.push(text);
      line = [];
    };

    items.forEach((item) => {
      if (currentY === null) {
        currentY = item.y;
        line.push(item);
        return;
      }
      if (Math.abs(item.y - currentY) <= 2) {
        line.push(item);
      } else {
        flushLine();
        currentY = item.y;
        line.push(item);
      }
    });
    flushLine();
  }

  return allLines;
}

export function detectPdfFormat(lines) {
  const text = lines.join("\n").replace(/\s*\|\s*/g, "|");
  if (text.includes("Division|")) return "Barras";
  if (text.includes("UPC REPORT")) return "Matricial";
  return "Desconocido";
}

export function extractDataBarras(lines) {
  const data = [];
  lines.forEach((line) => {
    const normalized = line.replace(/\s*\|\s*/g, "|");
    if (normalized.includes("Division|") && normalized.includes("Style|") && normalized.includes("UPC|")) return;
    if (!normalized.includes("|")) return;
    const parts = normalized.split("|").map((part) => part.trim());
    if (parts.length < 8) return;
    const [, style, upc, , colorCode, colorName, , size] = parts;
    const upcClean = String(upc ?? "").replace(/\D/g, "");
    if (!upcClean || !/^\d+$/.test(upcClean)) return;
    const styleValue = normalizeText(style);
    const colorCodeValue = normalizeText(colorCode);
    const colorNameValue = normalizeText(colorName);
    const sizeValue = normalizeText(size);
    data.push({
      STYLE: styleValue,
      "COLOR CODE": colorCodeValue,
      "COLOR NAME": colorNameValue,
      SIZE: sizeValue,
      "UPC CODE": upcClean,
      "STYLE COLOR": `${styleValue} ${colorCodeValue}`,
    });
  });
  return data;
}

export function extractDataMatricial(lines, options = {}) {
  const styleRe = options.stylePattern || /^([A-Z]{2}\d+[A-Z]?)\b/;
  const colorLineRe =
    options.colorPattern ||
    /^([A-Z0-9]{3,5})\s+([A-Z0-9/ .\-]+?)(?:\s+((?:\d{11,14}\s+)*\d{11,14}))?\s*$/;
  const sizeToken = /\*+\s*([A-Z0-9/]+)\s*\*+/g;
  const numbersOnly = /^(?:\d{11,14}\s+)*\d{11,14}$/;

  let styleActual = null;
  let sizes = [];
  const records = [];

  const extractSizes = (text) => Array.from(text.matchAll(sizeToken)).map((m) => m[1]);

  for (let i = 0; i < lines.length; i += 1) {
    const line = String(lines[i] ?? "").trim();
    if (!line || line.startsWith("-") || line.startsWith("*")) continue;

    const styleMatch = line.match(styleRe);
    if (styleMatch) {
      styleActual = styleMatch[1].toUpperCase();
      sizes = extractSizes(line);
      let j = i + 1;
      while (j < lines.length) {
        const next = String(lines[j] ?? "").trim();
        if (!next) break;
        if (styleRe.test(next) || colorLineRe.test(next)) break;
        const extraSizes = extractSizes(next);
        if (!extraSizes.length) break;
        sizes = sizes.concat(extraSizes);
        j += 1;
      }
      i = j - 1;
      continue;
    }

    const colorMatch = line.match(colorLineRe);
    if (colorMatch && styleActual && sizes.length) {
      const colorCode = colorMatch[1].toUpperCase();
      const colorName = colorMatch[2].trim().toUpperCase();
      let upcs = [];
      if (colorMatch[3]) upcs = upcs.concat(colorMatch[3].split(/\s+/));

      let k = i + 1;
      while (k < lines.length) {
        const next = String(lines[k] ?? "").trim();
        if (numbersOnly.test(next)) {
          upcs = upcs.concat(next.split(/\s+/));
          k += 1;
          continue;
        }
        if (colorLineRe.test(next) || styleRe.test(next)) break;
        if (!next) break;
        break;
      }

      const n = Math.min(sizes.length, upcs.length);
      for (let idx = 0; idx < n; idx += 1) {
        records.push({
          STYLE: styleActual,
          "COLOR CODE": colorCode,
          "COLOR NAME": colorName,
          SIZE: String(sizes[idx]).toUpperCase(),
          "UPC CODE": upcs[idx],
          "STYLE COLOR": `${styleActual} ${colorCode}`,
        });
      }
      i = k - 1;
    }
  }

  return records;
}

export async function extractPdfRecords(files, options = {}) {
  const all = [];
  for (const file of files) {
    const lines = await extractPdfLines(file);
    const format = detectPdfFormat(lines);
    let rows = [];
    if (format === "Barras") {
      rows = extractDataBarras(lines);
    } else {
      rows = extractDataMatricial(lines, options);
      if (!rows.length && format === "Desconocido") {
        rows = extractDataBarras(lines);
      }
    }
    all.push(...rows);
  }
  return all;
}
