"use strict";

const state = {
  storePath: "", xlsxPath: "", data: null, schema: null,
  tab: "摘要", dirty: false, selection: new Set(), anchor: null,
  filter: "", customer: "", job: "", rules: {},
  version: "?", update: null,
};

const ROW_H = 31, BUFFER = 20;

const TAB_HINTS = {
  "合同订单": "本页 = 合同头信息。总台数/设备型号自动算。新建订单先在这里加 JOB，再补付款条件/设备/发货/开票/回款。",
  "付款条件": "本页 = 结构化付款方式（款类/比例/账期/触发条件/说明），比例合计须 100%。合同页展示的“付款条件”文本在合同订单页改。",
  "设备台账": "本页 = 每台设备一行（含发货批次归属、质保时间、验收状态）。改设备属于哪个批次就在这里改“发货批次”列。",
  "发货批次": "本页只改“出荷日”和批次名（✎ 重命名批次）。台数/合计/覆盖番号为自动列；新增批次后要把设备台账里设备的“发货批次”改成同名。",
  "开票记录": "本页 = 每笔开票一行（款类/开票日/金额/覆盖批次/覆盖番号/账期/应收回款日/回款状态）。覆盖台数自动算。",
  "回款记录": "本页 = 每笔回款一行（款类/回款日/金额/覆盖批次/覆盖番号）。覆盖台数自动算。",
};

const $ = (id) => document.getElementById(id);

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function toast(msg, kind = "") {
  const t = $("toast");
  t.textContent = msg;
  t.className = "toast " + kind;
  clearTimeout(t._h);
  t._h = setTimeout(() => t.classList.add("hidden"), 2600);
}

function setStatus(msg, kind = "") {
  const s = $("status");
  s.textContent = msg;
  s.className = "status" + (kind ? " " + kind : "");
}

function fieldsOf(table) {
  return state.schema[table].map((f) => ({
    name: f[0], type: f[1], options: f[2] || [], derived: !!f[3],
  }));
}

function val(rec, field) {
  const v = rec[field];
  return v === null || v === undefined ? "" : v;
}

function renderAll() {
  const setBadge = (id, ok) => {
    const b = $(id); b.textContent = ok ? "✓" : "✗"; b.className = ok ? "ok" : "no";
  };
  setBadge("badgeStore", !!state.storePath);
  setBadge("badgeXlsx", !!state.xlsxPath);
  $("setStore").textContent = state.storePath || "（未选择）";
  $("setXlsx").textContent = state.xlsxPath || "（未选择）";
  $("setVersion").textContent = "v" + state.version;
  populateFilters();
  renderSummary();
  renderGrid();
  renderCounts();
}

function populateFilters() {
  if (!state.data) return;
  const custs = [...new Set(state.data["合同订单"]
    .map((c) => String(c["客户"] || "").trim()).filter(Boolean))].sort();
  $("filterCustomer").innerHTML = '<option value="">全部合同（客户）</option>' +
    custs.map((c) => `<option value="${esc(c)}" ${c === state.customer ? "selected" : ""}>${esc(c)}</option>`).join("");
  const jobs = state.data["合同订单"]
    .filter((c) => !state.customer || c["客户"] === state.customer)
    .map((c) => String(c["JOB No"])).sort();
  $("filterJob").innerHTML = '<option value="">全部 JOB</option>' +
    jobs.map((j) => `<option value="${esc(j)}" ${j === state.job ? "selected" : ""}>${esc(j)}</option>`).join("");
}

function renderSummary() {
  const cards = $("summaryCards");
  cards.innerHTML = "";
  if (!state.data) return;
  const counts = {};
  for (const t of Object.keys(state.schema)) counts[t] = state.data[t].length;
  const issues = validateAll();
  const errs = issues.filter((i) => i.severity === "error").length;
  const warns = issues.filter((i) => i.severity === "warning").length;
  const cardsHtml = [
    ["合同订单", counts["合同订单"]], ["付款条件", counts["付款条件"]],
    ["设备台账", counts["设备台账"]], ["发货批次", counts["发货批次"]],
    ["开票记录", counts["开票记录"]], ["回款记录", counts["回款记录"]],
    ["错误", errs], ["警告", warns],
  ].map(([k, v]) => `<div class="card"><div class="k">${k}</div><div class="v">${v}</div></div>`).join("");
  cards.innerHTML = cardsHtml;
  const ul = $("summaryIssues");
  ul.innerHTML = issues.slice(0, 200).map((i) =>
    `<li class="${i.severity}">${esc(i.msg)}</li>`).join("") ||
    '<li class="info">无问题</li>';
}

function renderCounts() {
  if (!state.data) return;
  const parts = [];
  for (const t of Object.keys(state.schema)) parts.push(`${t} ${state.data[t].length}`);
  $("counts").textContent = parts.join(" ｜ ") + (state.dirty ? " ｜ ● 未保存" : "");
}

function validateAll() {
  const issues = [];
  if (!state.data) return issues;
  const jobs = new Set(state.data["合同订单"].map((c) => String(c["JOB No"] || "")));
  for (const c of state.data["合同订单"]) {
    if (!/^\d{2}(BS|DS)\d{3}$/.test(String(c["JOB No"] || ""))) {
      issues.push({ severity: "error", msg: `合同 JOB No 格式不正确：${c["JOB No"]}` });
    }
  }
  const sums = {};
  for (const t of state.data["付款条件"]) {
    sums[t["JOB No"]] = (sums[t["JOB No"]] || 0) + Number(t["比例%"] || 0);
  }
  for (const [job, total] of Object.entries(sums)) {
    if (Math.abs(total - 100) > 0.01) {
      issues.push({ severity: "error", msg: `${job} 付款条件比例合计 ${total}% ≠ 100%` });
    }
  }
  const seen = new Set(), seenFull = new Set();
  for (const d of state.data["设备台账"]) {
    const job = String(d["JOB No"] || ""), ser = String(d["製造番号"] || ""), ki = String(d["機番"] || "");
    if (!ser) {
      issues.push({ severity: "error", msg: `${job} 设备台账存在空製造番号` });
      continue;
    }
    const k1 = job + "|" + ser, k2 = k1 + "|" + ki;
    if (seen.has(k1) && !(job === "23BS004" && (ser === "23BS004-058" || ser === "23BS004-061"))) {
      issues.push({ severity: "warning", msg: `製造番号重复：${job} ${ser}` });
    }
    seen.add(k1);
    if (seenFull.has(k2)) {
      issues.push({ severity: "error", msg: `设备业务键重复：${job} ${ser} ${ki}` });
    }
    seenFull.add(k2);
  }
  const devBatches = new Set(state.data["设备台账"].map((d) => d["JOB No"] + "|" + d["发货批次"]));
  for (const x of state.data["发货批次"]) {
    if (!devBatches.has(x["JOB No"] + "|" + x["发货批次"])) {
      issues.push({ severity: "warning", msg: `${x["JOB No"]} 批次 ${x["发货批次"]} 无设备，生成后不会显示（请补设备或删除）` });
    }
  }
  return issues;
}

function switchTab(tab) {
  state.tab = tab;
  $("panel-设置").classList.add("hidden");
  document.querySelectorAll(".tab").forEach((b) =>
    b.classList.toggle("active", b.dataset.tab === tab));
  $("panel-摘要").classList.toggle("hidden", tab !== "摘要");
  $("panel-表格").classList.toggle("hidden", tab === "摘要" || tab === "使用指南");
  $("panel-使用指南").classList.toggle("hidden", tab !== "使用指南");
  $("panel-预警规则").classList.toggle("hidden", tab !== "预警规则");
  document.querySelectorAll("[data-act=rename]").forEach((b) =>
    b.classList.toggle("hidden", tab !== "发货批次"));
  if (tab !== "摘要" && tab !== "使用指南" && tab !== "预警规则") renderGrid();
}

function renderRules() {
  const r = state.rules || {};
  $("rulesDays").value = r["临近天数"] ?? 30;
  $("rulesTol").value = r["金额容差"] ?? 0.01;
  $("rulesUnacc").checked = r["未验收待确认"] !== false;
  $("rulesInvBad").checked = r["发票异常待确认"] !== false;
}

function rulesFromForm() {
  return {
    "临近天数": Math.max(1, Math.round(Number($("rulesDays").value) || 30)),
    "金额容差": Math.max(0, Number($("rulesTol").value) || 0),
    "未验收待确认": $("rulesUnacc").checked,
    "发票异常待确认": $("rulesInvBad").checked,
  };
}

for (const id of ["rulesDays", "rulesTol", "rulesUnacc", "rulesInvBad"]) {
  $(id).addEventListener("input", () => {
    state.rules = rulesFromForm();
    state.dirty = true;
    renderCounts();
  });
}

$("btnRulesDefault").addEventListener("click", () => {
  state.rules = { "临近天数": 30, "金额容差": 0.01, "未验收待确认": true, "发票异常待确认": true };
  renderRules();
  state.dirty = true;
  renderCounts();
});

function visibleRows(table) {
  const q = state.filter.trim().toLowerCase();
  const jobOf = {};
  for (const c of state.data["合同订单"]) jobOf[c["JOB No"]] = c["客户"];
  return state.data[table].map((rec, i) => {
    const job = String(rec["JOB No"] || "");
    if (state.job && job !== state.job) return -1;
    if (state.customer && jobOf[job] !== state.customer) return -1;
    if (q && !Object.values(rec).join(" ").toLowerCase().includes(q)) return -1;
    return i;
  }).filter((i) => i >= 0);
}

function renderGrid() {
  const table = state.tab;
  if (table === "摘要" || !state.data) return;
  const fields = fieldsOf(table);
  const all = visibleRows(table);
  const total = all.length;
  const head = `<tr><th class="cb"></th>` +
    fields.map((f) => `<th title="${esc(f.name)}">${esc(f.name)}${f.derived ? " (自动)" : ""}</th>`).join("") + "</tr>";
  const wrap = $("gridWrap");
  const viewH = wrap.clientHeight || 600;
  const scrollTop = wrap.scrollTop || 0;
  let start = 0, end = total;
  if (total > 300) {
    start = Math.max(0, Math.floor(scrollTop / ROW_H) - BUFFER);
    end = Math.min(total, start + Math.ceil(viewH / ROW_H) + BUFFER * 2);
  }
  const rows = all.slice(start, end);
  const ncol = fields.length + 1;
  let html = `<tr class="spacer" style="height:${start * ROW_H}px"><td colspan="${ncol}"></td></tr>`;
  html += rows.map((ri) => {
    const rec = state.data[table][ri];
    const sel = state.selection.has(ri) ? " selected" : "";
    const warn = table === "发货批次" && !Number(rec["台数"] || 0) ? " warn-row" : "";
    const cells = fields.map((f) => {
      const v = val(rec, f.name);
      if (f.derived) {
        return `<td class="auto" title="自动计算"><span>${esc(v)}</span></td>`;
      }
      if (f.type === "select") {
        const opts = ["", ...f.options].map((o) =>
          `<option value="${esc(o)}" ${String(v) === String(o) ? "selected" : ""}>${esc(o || "（空）")}</option>`).join("");
        return `<td><select data-f="${esc(f.name)}">${opts}</select></td>`;
      }
      if (f.type === "date") {
        return `<td><input type="date" data-f="${esc(f.name)}" value="${esc(v)}"></td>`;
      }
      if (f.type === "number" || f.type === "int") {
        return `<td><input type="number" step="any" data-f="${esc(f.name)}" value="${esc(v)}"></td>`;
      }
      return `<td><input type="text" data-f="${esc(f.name)}" value="${esc(v)}"></td>`;
    }).join("");
    return `<tr data-ri="${ri}" class="${sel}${warn}">` +
      `<td class="cb"><input type="checkbox" ${sel ? "checked" : ""}></td>${cells}</tr>`;
  }).join("");
  html += `<tr class="spacer" style="height:${(total - end) * ROW_H}px"><td colspan="${ncol}"></td></tr>`;
  $("gridBody").innerHTML = html;
  $("gridHead").innerHTML = head;
  const hint = $("tabHint");
  hint.textContent = TAB_HINTS[table] || "";
}

function applySelectionClasses() {
  document.querySelectorAll("#gridBody tr[data-ri]").forEach((tr) => {
    const ri = Number(tr.dataset.ri);
    const sel = state.selection.has(ri);
    tr.classList.toggle("selected", sel);
    const cb = tr.querySelector("td.cb input");
    if (cb) cb.checked = sel;
  });
}

function updateRec(table, ri, field, raw) {
  const f = fieldsOf(table).find((x) => x.name === field);
  if (!f || f.derived) return;
  let v = raw;
  if (f.type === "number") v = raw === "" ? null : Number(raw);
  else if (f.type === "int") v = raw === "" ? null : Math.round(Number(raw));
  state.data[table][ri][field] = v;
  state.dirty = true;
  renderCounts();
}

$("gridBody").addEventListener("input", (e) => {
  const tr = e.target.closest("tr");
  if (!tr) return;
  const ri = Number(tr.dataset.ri);
  updateRec(state.tab, ri, e.target.dataset.f, e.target.value);
});

$("gridBody").addEventListener("change", (e) => {
  if (e.target.type === "checkbox" && e.target.closest("td.cb")) {
    const tr = e.target.closest("tr");
    const ri = Number(tr.dataset.ri);
    if (e.target.checked) state.selection.add(ri);
    else state.selection.delete(ri);
    tr.classList.toggle("selected", e.target.checked);
    return;
  }
  const tr = e.target.closest("tr");
  if (tr) {
    const ri = Number(tr.dataset.ri);
    updateRec(state.tab, ri, e.target.dataset.f, e.target.value);
  }
});

function selectRow(ri, extend) {
  if (extend && state.anchor !== null) {
    const a = state.anchor, b = ri;
    const [lo, hi] = a < b ? [a, b] : [b, a];
    state.selection = new Set();
    for (let i = lo; i <= hi; i++) state.selection.add(i);
  } else {
    state.anchor = ri;
    if (state.selection.has(ri)) state.selection.delete(ri);
    else state.selection.add(ri);
  }
  applySelectionClasses();
}

$("gridWrap").addEventListener("scroll", () => {
  if (state.tab === "摘要" || state._raf) return;
  state._raf = requestAnimationFrame(() => { state._raf = null; renderGrid(); });
});

$("gridBody").addEventListener("mousedown", (e) => {
  if (e.target.closest("td.cb") || e.target.tagName === "INPUT" || e.target.tagName === "SELECT") return;
  const tr = e.target.closest("tr");
  if (!tr) return;
  e.preventDefault();
  selectRow(Number(tr.dataset.ri), e.shiftKey || e.metaKey || e.ctrlKey);
});

document.querySelectorAll(".toolbar [data-act]").forEach((b) => {
  b.addEventListener("click", () => {
    const act = b.dataset.act;
    const table = state.tab;
    if (act === "add") {
      const rec = {};
      if (state.job) rec["JOB No"] = state.job;
      if (table === "合同订单" && state.customer) rec["客户"] = state.customer;
      state.data[table].push(rec);
      state.dirty = true;
      state.selection = new Set([state.data[table].length - 1]);
      renderGrid(); renderCounts();
      $("gridWrap").scrollTop = $("gridWrap").scrollHeight;
    } else if (act === "del") {
      const idxs = [...state.selection].sort((a, b) => b - a);
      for (const i of idxs) state.data[table].splice(i, 1);
      state.selection = new Set();
      state.dirty = true;
      renderGrid(); renderCounts();
    } else if (act === "dup") {
      const idxs = [...state.selection].sort((a, b) => a - b);
      const news = [];
      for (const i of idxs) news.push({ ...state.data[table][i] });
      const at = idxs.length ? idxs[idxs.length - 1] + 1 : state.data[table].length;
      state.data[table].splice(at, 0, ...news);
      state.selection = new Set(news.map((_, k) => at + k));
      state.dirty = true;
      renderGrid(); renderCounts();
      toast(`已重复 ${news.length} 行`);
    } else if (act === "copy") {
      copySelection(table);
    } else if (act === "rename" && table === "发货批次") {
      renameBatch();
    }
  });
});

function renameBatch() {
  const table = "发货批次";
  const idxs = [...state.selection];
  if (!idxs.length) { toast("请先选择要重命名的批次行", "error"); return; }
  const rec = state.data[table][idxs[0]];
  const job = String(rec["JOB No"] || "");
  const oldName = String(rec["发货批次"] || "");
  if (!job || !oldName) { toast("该行缺少 JOB No / 发货批次", "error"); return; }
  const newName = prompt(`将 ${job} 的批次「${oldName}」重命名为：`, oldName);
  if (newName === null || newName.trim() === "" || newName.trim() === oldName) return;
  const nw = newName.trim();
  let touched = 0;
  const fix = (r) => String(r["JOB No"] || "") === job;
  for (const d of state.data["设备台账"]) {
    if (fix(d) && String(d["发货批次"] || "") === oldName) { d["发货批次"] = nw; touched++; }
  }
  for (const x of state.data["发货批次"]) {
    if (fix(x) && String(x["发货批次"] || "") === oldName) { x["发货批次"] = nw; touched++; }
  }
  for (const t of ["开票记录", "回款记录"]) {
    for (const r of state.data[t]) {
      if (!fix(r)) continue;
      const bs = String(r["覆盖批次"] || "").split(";").map((b) => b.trim()).filter(Boolean);
      if (bs.includes(oldName)) {
        r["覆盖批次"] = bs.map((b) => (b === oldName ? nw : b)).join(";");
        touched++;
      }
    }
  }
  // 发货批次去重：同 (JOB, 批次) 保留有数据的那条（台数多/有出荷日的优先）
  const best = new Map();
  for (const x of state.data["发货批次"]) {
    const k = x["JOB No"] + "|" + x["发货批次"];
    const cur = best.get(k);
    if (!cur || Number(x["台数"] || 0) > Number(cur["台数"] || 0) ||
        (Number(x["台数"] || 0) === Number(cur["台数"] || 0) && x["出荷日"] && !cur["出荷日"])) {
      best.set(k, x);
    }
  }
  state.data["发货批次"] = [...best.values()];
  state.dirty = true;
  renderGrid(); renderCounts();
  toast(`已重命名批次 ${oldName} → ${nw}（同步 ${touched} 处）`, "ok");
}

function copySelection(table) {
  const idxs = [...state.selection].sort((a, b) => a - b);
  if (!idxs.length) { toast("请先选择行（点击行号区域）", "error"); return; }
  const fields = fieldsOf(table).map((f) => f.name);
  const lines = [fields.join("\t")];
  for (const i of idxs) {
    lines.push(fields.map((f) => String(val(state.data[table][i], f))).join("\t"));
  }
  navigator.clipboard.writeText(lines.join("\n")).then(() =>
    toast(`已复制 ${idxs.length} 行到剪贴板`)).catch(() => {
    const ta = document.createElement("textarea");
    ta.value = lines.join("\n");
    document.body.appendChild(ta); ta.select();
    document.execCommand("copy"); ta.remove();
    toast(`已复制 ${idxs.length} 行到剪贴板`);
  });
}

$("gridBody").addEventListener("paste", (e) => {
  const tr = e.target.closest("tr");
  if (!tr) return;
  e.preventDefault();
  const text = (e.clipboardData || window.clipboardData).getData("text/plain");
  if (!text.trim()) return;
  pasteGrid(state.tab, Number(tr.dataset.ri), e.target.closest("td")?.cellIndex ?? 1, text);
});

function pasteGrid(table, row0, col0, text) {
  const fields = fieldsOf(table);
  const lines = text.replace(/\r/g, "").split("\n").filter((l) => l.trim() !== "");
  if (!lines.length) return;
  const rows = lines.map((l) => l.split("\t"));
  let headerMap = null;
  const first = rows[0].map((x) => x.trim());
  if (first.some((x) => fields.some((f) => f.name === x)) &&
      first.filter((x) => fields.some((f) => f.name === x)).length >= 2) {
    headerMap = first.map((x) => fields.find((f) => f.name === x)?.name || null);
    rows.shift();
  }
  let r = row0;
  for (const row of rows) {
    if (!state.data[table][r]) {
      const rec = {};
      if (state.job) rec["JOB No"] = state.job;
      if (table === "合同订单" && state.customer) rec["客户"] = state.customer;
      state.data[table].push(rec);
      r = state.data[table].length - 1;
    }
    for (let c = 0; c < row.length; c++) {
      let field;
      if (headerMap) {
        if (!headerMap[c]) continue;
        field = fields.find((f) => f.name === headerMap[c]);
      } else {
        field = fields[col0 - 1 + c];
      }
      if (!field || field.derived) continue;
      state.data[table][r][field.name] = field.type === "number" ? Number(row[c] || 0)
        : field.type === "int" ? Math.round(Number(row[c] || 0)) : row[c];
    }
    r += 1;
  }
  state.dirty = true;
  renderGrid(); renderCounts();
  toast(`已粘贴 ${rows.length} 行`);
}

$("filterCustomer").addEventListener("change", (e) => {
  state.customer = e.target.value;
  state.job = "";
  populateFilters();
  renderGrid();
});

$("filterJob").addEventListener("change", (e) => {
  state.job = e.target.value;
  renderGrid();
});

$("btnClearFilter").addEventListener("click", () => {
  state.customer = "";
  state.job = "";
  state.filter = "";
  $("filterKeyword").value = "";
  populateFilters();
  renderGrid();
});

let _kwTimer = null;
$("filterKeyword").addEventListener("input", (e) => {
  clearTimeout(_kwTimer);
  _kwTimer = setTimeout(() => {
    state.filter = e.target.value;
    renderGrid();
  }, 180);
});

document.querySelectorAll(".tab").forEach((b) =>
  b.addEventListener("click", () => switchTab(b.dataset.tab)));

async function call(name, ...args) {
  if (!window.pywebview?.api?.[name]) throw new Error("后端不可用：" + name);
  return window.pywebview.api[name](...args);
}

$("btnSave").addEventListener("click", async () => {
  try {
    setStatus("保存中…");
    const r = await call("save_data", state.data, state.rules);
    state.dirty = false;
    setStatus("已保存：" + r.path);
    toast("数据已保存", "ok");
    renderCounts();
  } catch (err) { setStatus(String(err), "error"); toast(String(err), "error"); }
});

$("btnGenerate").addEventListener("click", async () => {
  try {
    setStatus("生成台账中…");
    const r = await call("generate", state.data, state.rules);
    state.dirty = false;
    const msg = `已生成：${r.out}（未回收 ${r.counts["未回收行数"]} 行，合计 ${Number(r.counts["未回收合计"]).toLocaleString()}）`;
    setStatus(msg);
    toast("台账已生成", "ok");
    renderCounts();
  } catch (err) { setStatus(String(err), "error"); toast(String(err), "error"); }
});

$("btnPickStore").addEventListener("click", async () => {
  const p = await call("pick_store");
  if (p) { await refresh(p, null); }
});

$("btnPickXlsx").addEventListener("click", async () => {
  const p = await call("pick_xlsx");
  if (p) { await refresh(null, p); }
});

$("btnSettings").addEventListener("click", () => {
  ["摘要", "表格", "使用指南", "预警规则"].forEach((p) => $("panel-" + p).classList.add("hidden"));
  $("panel-设置").classList.remove("hidden");
  document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
  $("setVersion").textContent = "v" + state.version;
});

$("btnCheckUpdate").addEventListener("click", async () => {
  $("updateInfo").textContent = "检查中…";
  $("btnDownloadUpdate").classList.add("hidden");
  try {
    const u = await call("check_update");
    state.update = u;
    if (u.error) { $("updateInfo").textContent = u.error; return; }
    if (u.has_update) {
      $("updateInfo").innerHTML = `最新 <b>v${u.latest}</b>（当前 v${u.current}）`;
      $("btnDownloadUpdate").classList.toggle("hidden", !u.asset_url);
    } else {
      $("updateInfo").textContent = `已是最新（v${u.current}）`;
    }
  } catch (e) { $("updateInfo").textContent = String(e); }
});

$("btnDownloadUpdate").addEventListener("click", async () => {
  const u = state.update;
  if (!u || !u.asset_url) {
    toast("未找到下载资产，打开发布页", "error");
    if (u && u.release_url) await call("open_path", u.release_url);
    return;
  }
  try {
    setStatus("下载更新中…");
    const dir = await call("download_update", u.asset_url, u.asset_name);
    setStatus("已下载并解压到：" + dir);
    toast("已下载，请退出本程序后用新版本替换", "ok");
    await call("open_path", dir);
  } catch (e) { setStatus(String(e), "error"); toast(String(e), "error"); }
});

$("btnOpenXlsx").addEventListener("click", async () => {
  if (state.xlsxPath) await call("open_path", state.xlsxPath);
});

async function refresh(store, xlsx) {
  const r = await call("load_state", store, xlsx);
  state.storePath = r.store_path;
  state.xlsxPath = r.xlsx_path;
  state.data = r.data;
  state.schema = r.schema;
  state.rules = r.rules || {};
  state.version = r.version || "?";
  state.dirty = false;
  renderAll();
  renderRules();
  switchTab("摘要");
  setStatus("就绪");
}

window.addEventListener("pywebviewready", async () => {
  try {
    const r = await call("load_state", null, null);
    state.storePath = r.store_path;
    state.xlsxPath = r.xlsx_path;
    state.data = r.data;
    state.schema = r.schema;
    state.rules = r.rules || {};
    state.dirty = false;
    renderAll();
    renderRules();
    switchTab("摘要");
    setStatus("就绪");
  } catch (err) {
    setStatus("加载失败：" + err, "error");
  }
});
