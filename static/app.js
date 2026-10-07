const state = {
  mode: "google", // google | upload
  spreadsheetId: "",
  fileId: "",
  sheets: [], // [{name, checked}]
  previews: {}, // sheetName -> parsed order
};

const $ = (id) => document.getElementById(id);

function toast(msg, isError = false) {
  const t = $("toast");
  t.textContent = msg;
  t.className = isError ? "error" : "";
  t.hidden = false;
  clearTimeout(t._timer);
  t._timer = setTimeout(() => (t.hidden = true), 5000);
}

async function api(path, opts = {}) {
  const res = await fetch(path, opts);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.detail || `HTTP ${res.status}`);
  return data;
}

// --- Auth / status -----------------------------------------------------------

async function refreshStatus() {
  const s = await api("/api/status");
  const pill = $("auth-status");
  if (!s.google_ready) {
    pill.textContent = "Google: chưa cấu hình credentials.json";
    pill.className = "pill warn";
  } else if (s.authenticated) {
    pill.textContent = "Google: đã kết nối";
    pill.className = "pill ok";
  } else {
    pill.textContent = "Google: cần đăng nhập";
    pill.className = "pill warn";
  }
  $("google-login-box").hidden = s.authenticated;
  $("google-url-box").hidden = !s.authenticated;
  $("btn-login").disabled = !s.google_ready;
  $("btn-logout").hidden = !s.authenticated;
  if (!s.google_ready) {
    $("google-status-text").textContent =
      "Chưa có credentials.json. Tạo OAuth client trên Google Cloud và đặt file vào thư mục tool (xem README).";
  }
}

$("btn-login").onclick = async () => {
  const { url } = await api("/api/auth/login");
  window.location.href = url;
};
$("btn-logout").onclick = async () => {
  await api("/api/auth/logout", { method: "POST" });
  refreshStatus();
};

// --- Tabs --------------------------------------------------------------------

$("tab-google").onclick = () => setMode("google");
$("tab-upload").onclick = () => setMode("upload");
function setMode(mode) {
  state.mode = mode;
  $("tab-google").classList.toggle("active", mode === "google");
  $("tab-upload").classList.toggle("active", mode === "upload");
  $("panel-google").hidden = mode !== "google";
  $("panel-upload").hidden = mode !== "upload";
}

// --- Load sheets -------------------------------------------------------------

$("btn-load-google").onclick = async () => {
  const url = $("sheet-url").value.trim();
  if (!url) return toast("Dán link hoặc ID của Google Sheet trước", true);
  try {
    const { sheets } = await api(`/api/sheets?spreadsheetId=${encodeURIComponent(url)}`);
    setSheets(sheets);
  } catch (e) {
    toast(e.message, true);
  }
};

$("btn-upload").onclick = async () => {
  const input = $("file-input");
  if (!input.files.length) return toast("Chọn file .xlsx hoặc .csv", true);
  const fd = new FormData();
  fd.append("file", input.files[0]);
  $("upload-status").textContent = "Đang tải lên…";
  try {
    const r = await api("/api/upload", { method: "POST", body: fd });
    state.fileId = r.fileId;
    $("upload-status").textContent = `Đã tải: ${r.filename}`;
    setSheets(r.sheets);
  } catch (e) {
    $("upload-status").textContent = "";
    toast(e.message, true);
  }
};

function setSheets(names) {
  state.sheets = names.map((n) => ({ name: n, checked: !isIgnored(n) }));
  renderSheetList();
  $("card-sheets").hidden = false;
  $("card-preview").hidden = true;
  $("card-generate").hidden = true;
}

function isIgnored(name) {
  return $("ignore-list")
    .value.split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean)
    .includes(name.trim().toLowerCase());
}

$("ignore-list").oninput = () => {
  state.sheets.forEach((s) => (s.checked = !isIgnored(s.name)));
  renderSheetList();
};

function renderSheetList() {
  const box = $("sheet-list");
  box.innerHTML = "";
  for (const s of state.sheets) {
    const label = document.createElement("label");
    label.className = "sheet-item";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = s.checked;
    cb.onchange = () => {
      s.checked = cb.checked;
      refreshPreview();
    };
    label.append(cb, document.createTextNode(" " + s.name));
    box.append(label);
  }
}

// --- Preview -----------------------------------------------------------------

async function refreshPreview() {
  const selected = state.sheets.filter((s) => s.checked);
  if (!selected.length) {
    $("card-preview").hidden = true;
    $("card-generate").hidden = true;
    return;
  }
  $("card-preview").hidden = false;
  $("card-generate").hidden = false;
  const box = $("preview-list");
  box.innerHTML = "<p class='muted'>Đang đọc dữ liệu…</p>";
  const qs =
    state.mode === "google"
      ? `spreadsheetId=${encodeURIComponent($("sheet-url").value.trim())}`
      : `fileId=${state.fileId}`;
  const results = await Promise.all(
    selected.map(async (s) => {
      try {
        return await api(`/api/preview?${qs}&sheet=${encodeURIComponent(s.name)}`);
      } catch (e) {
        return { sheet: s.name, error: e.message };
      }
    })
  );
  box.innerHTML = "";
  let okOrders = 0;
  results.forEach((p) => {
    state.previews[p.sheet] = p;
    box.append(renderPreviewCard(p));
    if (p.hasHeader && p.rows.length) okOrders++;
  });
  const skipped = results.filter((p) => p.skipped).reduce((a, p) => a + p.skipped, 0);
  $("generate-summary").textContent = `${okOrders} đơn hàng hợp lệ` +
    (skipped ? `, ${skipped} dòng lỗi sẽ bị bỏ qua` : "");
}

function renderPreviewCard(p) {
  const div = document.createElement("div");
  div.className = "order-card";
  if (p.error) {
    div.innerHTML = `<h3>${esc(p.sheet)}</h3><p class="err">${esc(p.error)}</p>`;
    return div;
  }
  const head = document.createElement("div");
  head.className = "order-head";
  head.innerHTML = `
    <h3>${esc(p.sheet)}</h3>
    <span>ID đơn: <b>${esc(p.orderId)}</b></span>
    <span>Người nhận: <b>${esc(p.customerName)}</b></span>`;
  div.append(head);
  if (!p.hasHeader) {
    div.append(el("p", "err", p.error));
    return div;
  }
  const table = document.createElement("table");
  table.innerHTML = `<tr><th>SKU gian hàng</th><th>Số lượng</th><th>Trạng thái</th></tr>`;
  p.rows.forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${esc(r.sku)}</td><td>${esc(r.qty)}</td><td>✓</td>`;
    table.append(tr);
  });
  (p.errors || []).forEach((e) => {
    const tr = document.createElement("tr");
    tr.className = "bad";
    tr.innerHTML = `<td>${esc(e.sku || "—")}</td><td>${esc(e.rawQty || "—")}</td><td>⚠ ${esc(e.error)} (bỏ qua)</td>`;
    table.append(tr);
  });
  div.append(table);
  return div;
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
  );
}
function el(tag, cls, text) {
  const e = document.createElement(tag);
  e.className = cls;
  e.textContent = text;
  return e;
}

// --- Generate ----------------------------------------------------------------

$("btn-generate").onclick = async () => {
  const selected = state.sheets.filter((s) => s.checked).map((s) => s.name);
  if (!selected.length) return toast("Chưa chọn sheet nào", true);
  const body =
    state.mode === "google"
      ? { spreadsheetId: $("sheet-url").value.trim(), sheets: selected }
      : { fileId: state.fileId, sheets: selected };
  $("generate-status").textContent = "Đang tạo file…";
  try {
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || `HTTP ${res.status}`);
    }
    const blob = await res.blob();
    const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename="([^"]+)"/);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = m ? m[1] : "output.xlsx";
    a.click();
    URL.revokeObjectURL(a.href);
    $("generate-status").textContent = "Đã tải file xuống.";
  } catch (e) {
    $("generate-status").textContent = "";
    toast(e.message, true);
  }
};

// --- Init --------------------------------------------------------------------

$("ignore-list").addEventListener("change", refreshPreview);
refreshStatus();
