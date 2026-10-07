// Core logic (port of app/core.py): sheet-name rules, source parsing, validation.

const TEMPLATE_FIELDS = {
  order_id: "ID Đơn hàng",
  receiver: "Tên người nhận",
  sku: "SKU gian hàng",
  price: "Đơn giá",
  qty: "Số lượng",
};

function norm(s) {
  return String(s ?? "").trim().toLowerCase();
}

function slugify(name) {
  let s = String(name).normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  s = s.replace(/[^0-9A-Za-z]+/g, "_").replace(/^_+|_+$/g, "");
  return s;
}

// Sheet names carry a category prefix that is not part of the customer name:
// 'Báo giá - Lukibum Aqua' -> 'Lukibum Aqua'
// 'Đơn hàng (Bum)'         -> 'Bum'
function stripSheetPrefix(name) {
  let n = String(name).trim();
  n = n.replace(/^Báo giá/i, "").replace(/^Đơn hàng/i, "");
  n = n.replace(/^[\s\-–—_()]+/, "");
  n = n.replace(/\s*\)$/, "");
  return n.trim();
}

// 'Báo giá - Lukibum Aqua' -> ['Aqua', 'Lukibum', 'Lukibum']
// 'Báo giá - Duy ZL'       -> ['KH',   'Duy ZL',  'Duy_ZL']
// 'Đơn hàng (Bum)'         -> ['KH',   'Bum',     'Bum']
function parseSheetName(name) {
  const n = stripSheetPrefix(name);
  let prefix, rest;
  if (n.includes("Aqua")) {
    prefix = "Aqua";
    rest = n.replace("Aqua", "").replace(/^[\s\-_]+|[\s\-_]+$/g, "");
  } else if (/CTV/i.test(n)) {
    prefix = "CTV";
    rest = n.replace(/CTV/i, "").replace(/^[\s\-_]+|[\s\-_]+$/g, "");
  } else {
    prefix = "KH";
    rest = n;
  }
  return [prefix, rest, slugify(rest) || "khach"];
}

function pad2(x) {
  return String(x).padStart(2, "0");
}

// Locate the header row containing 'BS - SKU' and 'BS - Số lượng'.
// Returns {row, cols} (0-based) or null.
function findHeaderRow(values) {
  for (let i = 0; i < Math.min(values.length, 15); i++) {
    const row = values[i] || [];
    const cells = {};
    row.forEach((c, j) => {
      if (c !== null && c !== undefined && String(c).trim() !== "") cells[norm(c)] = j;
    });
    const skuCol = cells["bs - sku"];
    const qtyCol = Object.entries(cells).find(
      ([h]) => h.includes("số lượng") && h.includes("bs")
    )?.[1];
    if (skuCol !== undefined && qtyCol !== undefined) {
      return { row: i, cols: cells, skuCol, qtyCol };
    }
  }
  return null;
}

function parseQty(raw) {
  if (raw === null || raw === undefined || String(raw).trim() === "") {
    return [null, "thiếu số lượng"];
  }
  const v = parseFloat(String(raw).replace(",", ".").trim());
  if (Number.isNaN(v)) return [null, `số lượng không hợp lệ: ${raw}`];
  if (v <= 0) return [null, `số lượng phải > 0 (nhận ${raw})`];
  return [Number.isInteger(v) ? v : v, null];
}

// Parse one sheet's values (array of rows) into an order preview structure.
function parseOrderSheet(sheetName, values, now = new Date()) {
  const [prefix, customer, slug] = parseSheetName(sheetName);
  const ts =
    `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}` +
    `_${pad2(now.getHours())}${pad2(now.getMinutes())}`;
  const orderId = `${prefix}_${slug}_${ts}`;
  const result = {
    sheet: sheetName,
    orderId,
    customerName: customer,
    hasHeader: false,
    rows: [],
    errors: [],
    skipped: 0,
    error: null,
  };
  const header = findHeaderRow(values);
  if (!header) {
    result.error = "Không tìm thấy cột 'BS - SKU' / 'BS - Số lượng' — bỏ qua sheet này";
    return result;
  }
  result.hasHeader = true;

  for (const raw of values.slice(header.row + 1)) {
    const row = raw || [];
    const sku = row[header.skuCol];
    const qtyRaw = row[header.qtyCol];
    if ((sku === null || sku === undefined || String(sku).trim() === "") &&
        (qtyRaw === null || qtyRaw === undefined || String(qtyRaw).trim() === "")) {
      continue;
    }
    const entry = { sku: null, qty: null, status: "ok", error: null };
    if (sku === null || sku === undefined || String(sku).trim() === "") {
      entry.status = "error";
      entry.error = "thiếu SKU";
    }
    const [qty, err] = parseQty(qtyRaw);
    if (err) {
      entry.status = "error";
      entry.error = err;
    } else {
      entry.qty = qty;
    }
    if (sku !== null && sku !== undefined && String(sku).trim() !== "") {
      entry.sku = String(sku).trim();
    }
    if (entry.status === "error") {
      result.skipped++;
      result.errors.push({
        sku: entry.sku,
        rawQty: qtyRaw === null || qtyRaw === undefined ? "" : String(qtyRaw),
        error: entry.error,
      });
    } else {
      result.rows.push(entry);
    }
  }
  return result;
}

// --- Template filling (xlsx-populate) ----------------------------------------

// Find header row + first empty row after it in an xlsx-populate sheet.
// Returns {startRow, cols} (1-based, xlsx-populate indexing) or throws.
function locateTemplateFields(sheet) {
  const maxCol = 200;
  for (let headerRow = 1; headerRow <= 5; headerRow++) {
    const cells = {};
    for (let c = 1; c <= maxCol; c++) {
      const v = sheet.cell(headerRow, c).value();
      if (v !== null && v !== undefined && String(v).trim() !== "") cells[norm(v)] = c;
    }
    const cols = {};
    const missing = [];
    for (const [field, header] of Object.entries(TEMPLATE_FIELDS)) {
      const col = cells[norm(header)];
      if (col === undefined) missing.push(header);
      else cols[field] = col;
    }
    if (missing.length) continue;
    let startRow = headerRow + 1;
    for (let r = headerRow + 1; r <= headerRow + 10; r++) {
      let empty = true;
      for (let c = 1; c <= maxCol; c++) {
        const v = sheet.cell(r, c).value();
        if (v !== null && v !== undefined && String(v).trim() !== "") {
          empty = false;
          break;
        }
      }
      if (empty) {
        startRow = r;
        break;
      }
    }
    return { startRow, cols };
  }
  throw new Error(
    "Không tìm thấy các cột " +
      Object.values(TEMPLATE_FIELDS).map((h) => `'${h}'`).join(", ") +
      " trong file template"
  );
}

// Append one order's rows at startRow. Returns next free row.
function appendOrder(sheet, order, startRow, cols) {
  let row = startRow;
  for (const entry of order.rows) {
    sheet.cell(row, cols.order_id).value(order.orderId);
    sheet.cell(row, cols.receiver).value(order.customerName);
    sheet.cell(row, cols.sku).value(entry.sku);
    sheet.cell(row, cols.price).value(0);
    sheet.cell(row, cols.qty).value(entry.qty);
    row++;
  }
  return row;
}

// --- CSV parsing (gviz output is standard quoted CSV) -------------------------

function parseCSV(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      field = "";
      rows.push(row);
      row = [];
    } else {
      field += ch;
    }
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}
