// --- BS Order Generator — static site logic -----------------------------------
// core.js provides: parseSheetName, parseOrderSheet, findHeaderRow,
// locateTemplateFields, appendOrder, parseCSV.

const $ = (id) => document.getElementById(id);

const IGNORE_SHEETS = new Set(["báo giá"]);
const LS_URL_KEY = "bs_sheet_url";

const state = {
  workbook: null,
  outputBlob: null,
  outputName: null,
  busy: false,
};

const STEPS = [
  { id: "template", label: "Nạp template BigSeller" },
  { id: "fetch", label: "Tải workbook từ Google Sheets (file lớn, có thể mất vài giây)" },
  { id: "scan", label: "Quét các sheet đơn hàng" },
  { id: "build", label: "Điền dữ liệu vào template" },
];

// --- Toast ---------------------------------------------------------------------

let toastTimer = null;
function toast(msg, isError = false) {
  const t = $("toast");
  t.textContent = msg;
  t.className = isError ? "error" : "";
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.hidden = true), 6000);
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
  );
}

// --- Progress steps --------------------------------------------------------------

function renderSteps() {
  const ol = $("step-list");
  ol.innerHTML = "";
  for (const s of STEPS) {
    const li = document.createElement("li");
    li.id = "step-" + s.id;
    li.className = "step pending";
    li.innerHTML = `<span class="step-icon">⏳</span> ${esc(s.label)}`;
    ol.append(li);
  }
}

function setStep(id, status, note) {
  const li = $("step-" + id);
  if (!li) return;
  li.className = "step " + status;
  const icons = { pending: "⏳", running: "", done: "✓", error: "✗" };
  const icon =
    status === "running"
      ? '<span class="spinner"></span>'
      : `<span class="step-icon">${icons[status]}</span>`;
  li.innerHTML = `${icon} ${esc(STEPS.find((s) => s.id === id).label)}${
    note ? ` <span class="muted">— ${esc(note)}</span>` : ""
  }`;
}

// --- Google Sheets ----------------------------------------------------------------

function extractSpreadsheetId(urlOrId) {
  const s = String(urlOrId).trim();
  const m = s.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  return m ? m[1] : s;
}

async function fetchGoogleWorkbook(id) {
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
  const res = await fetch(url);
  if (res.status === 404) throw new Error("Không tìm thấy sheet — kiểm tra lại link");
  if (res.status === 403 || res.status === 302) {
    throw new Error(
      'Sheet chưa được chia sẻ công khai. Bật: Chia sẻ → "Ai có link đều xem được"'
    );
  }
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return XLSX.read(await res.arrayBuffer(), { type: "array" });
}

function getSheetValues(sheetName) {
  const ws = state.workbook.Sheets[sheetName];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}

// --- Generate ----------------------------------------------------------------------

$("btn-generate").onclick = async () => {
  if (state.busy) return;
  const url = $("sheet-url").value.trim();
  if (!url) return toast("Dán link Google Sheet trước", true);
  localStorage.setItem(LS_URL_KEY, url);

  state.busy = true;
  $("btn-generate").disabled = true;
  $("btn-generate").classList.add("loading");
  $("btn-generate").innerHTML = '<span class="spinner"></span> Đang tạo file…';
  $("btn-download").disabled = true;
  $("card-progress").hidden = false;
  $("card-result").hidden = true;
  renderSteps();

  try {
    // 1. Load template
    setStep("template", "running");
    const tplRes = await fetch("template.xlsx");
    if (!tplRes.ok) throw new Error("Không tải được template.xlsx từ site");
    const templateDoc = await XlsxPopulate.fromDataAsync(await tplRes.arrayBuffer());
    setStep("template", "done");

    // 2. Fetch workbook from Google
    setStep("fetch", "running");
    state.workbook = await fetchGoogleWorkbook(extractSpreadsheetId(url));
    setStep("fetch", "done", `${state.workbook.SheetNames.length} sheet`);

    // 3. Scan order sheets (skip hidden sheets — SheetJS stores visibility in
    //    workbook.Workbook.Sheets[].Hidden: 0 visible, 1 hidden, 2 veryHidden)
    setStep("scan", "running");
    const now = new Date();
    const hiddenNames = new Set(
      (state.workbook.Workbook && state.workbook.Workbook.Sheets || [])
        .filter((s) => s.Hidden > 0)
        .map((s) => s.name)
    );
    const candidates = [];
    let skippedNoHeader = 0;
    let skippedHidden = 0;
    for (const name of state.workbook.SheetNames) {
      if (IGNORE_SHEETS.has(name.trim().toLowerCase())) continue;
      if (hiddenNames.has(name)) {
        skippedHidden++;
        continue;
      }
      if (findHeaderRow(getSheetValues(name)) === null) {
        skippedNoHeader++;
        continue;
      }
      candidates.push({ name, order: parseOrderSheet(name, getSheetValues(name), now) });
    }
    const ok = candidates.filter((c) => c.order.rows.length);
    const bad = candidates.filter((c) => !c.order.rows.length);
    if (!ok.length) {
      setStep("scan", "error", "không có sheet nào tạo được đơn");
      renderResult(ok, bad, skippedNoHeader, skippedHidden, candidates.length);
      toast("Không tạo được đơn nào — xem chi tiết bên dưới", true);
      return;
    }
    setStep("scan", "done", `${candidates.length} sheet đơn, ${ok.length} hợp lệ`);

    // 4. Build output file (rebuild from the clean template each run)
    setStep("build", "running");
    const doc = await XlsxPopulate.fromDataAsync(
      await templateDoc.outputAsync({ type: "arraybuffer" })
    );
    const sheet = doc.sheet(0);
    const { startRow, cols } = locateTemplateFields(sheet);
    let row = startRow;
    const seenIds = new Set();
    for (const c of ok) {
      let o = c.order;
      if (seenIds.has(o.orderId)) o = { ...o, orderId: o.orderId + "_2" };
      seenIds.add(o.orderId);
      row = appendOrder(sheet, o, row, cols);
    }
    const ts =
      `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}` +
      `_${pad2(now.getHours())}${pad2(now.getMinutes())}`;
    state.outputName = `BS-orders_${ts}.xlsx`;
    state.outputBlob = await doc.outputAsync({ type: "blob" });
    setStep("build", "done");

    $("btn-download").disabled = false;
    $("file-name").textContent = state.outputName;
    renderResult(ok, bad, skippedNoHeader, skippedHidden, candidates.length);
  } catch (e) {
    // mark the currently running step as failed
    const running = STEPS.find((s) => $("step-" + s.id).classList.contains("running"));
    if (running) setStep(running.id, "error", e.message);
    toast("Lỗi: " + e.message, true);
  } finally {
    state.busy = false;
    $("btn-generate").disabled = false;
    $("btn-generate").classList.remove("loading");
    $("btn-generate").textContent = "Tạo file xlsx cho BS";
  }
};

$("btn-download").onclick = () => {
  if (!state.outputBlob) return;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(state.outputBlob);
  a.download = state.outputName;
  a.click();
  URL.revokeObjectURL(a.href);
};

// --- Result rendering ----------------------------------------------------------------

function renderResult(ok, bad, skippedNoHeader, skippedHidden, totalCandidates) {
  $("card-result").hidden = false;
  $("result-summary").textContent =
    `Đã tạo được tổng cộng ${ok.length} đơn / ${totalCandidates} sheet.`;
  const notes = [];
  if (skippedHidden) notes.push(`${skippedHidden} sheet ẩn`);
  if (skippedNoHeader) notes.push(`${skippedNoHeader} sheet không có cột 'BS - SKU'`);
  $("result-note").textContent = notes.length
    ? `Đã bỏ qua ${notes.join(", ")}.`
    : "";

  const box = $("order-list");
  box.innerHTML = "";
  ok.forEach((c, i) => box.append(renderOrderCard(i + 1, c.order)));
  bad.forEach((c, i) => box.append(renderErrorCard(ok.length + i + 1, c)));
  $("card-result").scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function renderOrderCard(index, order) {
  const totalQty = order.rows.reduce((s, r) => s + r.qty, 0);
  const details = document.createElement("details");
  details.className = "order-card";
  const summary = document.createElement("summary");
  summary.innerHTML =
    `<b>Đơn ${index}: ${esc(order.customerName)}</b> — ` +
    `${order.rows.length} mặt hàng • ${totalQty} sản phẩm ` +
    `<span class="muted">(ID: ${esc(order.orderId)})</span>`;
  details.append(summary);

  const table = document.createElement("table");
  table.innerHTML = "<tr><th>SKU gian hàng</th><th>Số lượng</th></tr>";
  for (const r of order.rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${esc(r.sku)}</td><td>${esc(r.qty)}</td>`;
    table.append(tr);
  }
  details.append(table);

  if (order.errors.length) {
    const p = document.createElement("p");
    p.className = "warn";
    p.textContent = `${order.errors.length} dòng lỗi đã bỏ qua:`;
    details.append(p);
    const ul = document.createElement("ul");
    ul.className = "err-list";
    for (const e of order.errors.slice(0, 10)) {
      const li = document.createElement("li");
      li.textContent = `${e.sku || "(không có SKU)"} — ${e.error}`;
      ul.append(li);
    }
    if (order.errors.length > 10) {
      const li = document.createElement("li");
      li.textContent = `… và ${order.errors.length - 10} dòng khác`;
      ul.append(li);
    }
    details.append(ul);
  }
  return details;
}

function renderErrorCard(index, c) {
  const div = document.createElement("div");
  div.className = "order-card error-card";
  const reason = c.order.error || "không có dòng dữ liệu hợp lệ";
  const firstErrors = c.order.errors
    .slice(0, 5)
    .map((e) => `<li>${esc(`${e.sku || "(không có SKU)"} — ${e.error}`)}</li>`)
    .join("");
  div.innerHTML =
    `<b>Đơn ${index}: ${esc(c.name)}</b> — không tạo được đơn. ` +
    `<span class="err">${esc(reason)}</span>` +
    (firstErrors ? `<ul class="err-list">${firstErrors}</ul>` : "");
  return div;
}

// --- Init ------------------------------------------------------------------------------

(function init() {
  const saved = localStorage.getItem(LS_URL_KEY);
  if (saved) $("sheet-url").value = saved;
})();
