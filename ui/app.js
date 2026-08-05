"use strict";

const state = {
  storePath: "", xlsxPath: "", data: null, schema: null,
  tab: "摘要", dirty: false, selection: new Set(), anchor: null,
  filter: "", colFilters: {}, sortBy: {}, rules: {},
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
function s(v) { return v === null || v === undefined ? "" : String(v); }

function esc(x) {
  return String(x ?? "").replace(/[&<>"']/g, (c) => (
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
  renderSummary();
  renderGrid();
  renderCounts();
}

function currentContext() {
  // 从当前表的单值列筛选推导"新建行的上下文"（替代旧的客户/JOB 下拉）
  const cf = state.colFilters[state.tab] || {};
  const ctx = { job: "", customer: "" };
  const js = cf["JOB No"];
  if (js && js.size === 1) ctx.job = [...js][0];
  if (state.tab === "合同订单") {
    const cs = cf["客户"];
    if (cs && cs.size === 1) ctx.customer = [...cs][0];
  }
  return ctx;
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
  $("panel-表格").classList.toggle("hidden", tab === "摘要" || tab === "使用指南" || tab === "预警规则");
  $("panel-使用指南").classList.toggle("hidden", tab !== "使用指南");
  $("panel-预警规则").classList.toggle("hidden", tab !== "预警规则");
  document.querySelectorAll(".only-ship").forEach((b) =>
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
  const cf = state.colFilters[table] || {};
  const so = state.sortBy[table];
  const recs = state.data[table];
  const idxs = [];
  for (let i = 0; i < recs.length; i++) {
    const rec = recs[i];
    let keep = true;
    for (const f in cf) {
      const set = cf[f];
      if (set && !set.has(String(rec[f] ?? ""))) { keep = false; break; }
    }
    if (keep && q && !Object.values(rec).join(" ").toLowerCase().includes(q)) keep = false;
    if (keep) idxs.push(i);
  }
  if (so && so.dir) {
    const fld = fieldsOf(table).find((x) => x.name === so.field);
    const isNum = fld && (fld.type === "number" || fld.type === "int");
    const f = so.field, d = so.dir;
    idxs.sort((a, b) => {
      let cmp;
      if (isNum) cmp = (Number(recs[a][f]) || 0) - (Number(recs[b][f]) || 0);
      else cmp = String(recs[a][f] ?? "").localeCompare(String(recs[b][f] ?? ""), "zh");
      return d * cmp;
    });
  }
  return idxs;
}

function recomputeDerived() {
  // 镜像 core.derive() 的派生聚合（仅显示用；save 时后端 derive 仍权威落盘）
  const d = state.data;
  if (!d) return;
  const devBy = {};
  for (const dev of d["设备台账"]) {
    const k = s(dev["JOB No"]) + "" + s(dev["发货批次"]);
    (devBy[k] ||= []).push(dev);
  }
  for (const x of d["发货批次"]) {
    const devs = devBy[s(x["JOB No"]) + "" + s(x["发货批次"])] || [];
    x["台数"] = devs.length;
    x["未税合计"] = Math.round(devs.reduce((a, v) => a + (Number(v["未税单价"]) || 0), 0) * 1000) / 1000;
    x["含税合计"] = Math.round(x["未税合计"] * 1.13 * 100) / 100;
    x["覆盖製造番号"] = devs.map((v) => s(v["製造番号"])).filter(Boolean).join(";");
  }
  const cnt = {}, models = {};
  for (const dev of d["设备台账"]) {
    const j = s(dev["JOB No"]);
    cnt[j] = (cnt[j] || 0) + 1;
    (models[j] ||= new Set()).add(s(dev["设备型号"]));
  }
  for (const c of d["合同订单"]) {
    const j = s(c["JOB No"]);
    c["总台数"] = cnt[j] || 0;
    if (!s(c["设备型号"])) c["设备型号"] = [...(models[j] || new Set())].filter((m) => m).sort().join(";");
  }
  for (const t of ["开票记录", "回款记录"]) {
    for (const x of d[t]) {
      const set = new Set(s(x["覆盖製造番号"]).split(";").map((p) => p.trim()).filter(Boolean));
      x["覆盖台数"] = set.size;
    }
  }
}

function renderGrid() {
  const table = state.tab;
  if (table === "摘要" || !state.data) return;
  const fields = fieldsOf(table);
  const all = visibleRows(table);
  const total = all.length;
  const so = state.sortBy[state.tab];
  const cfT = state.colFilters[state.tab] || {};
  const head = `<tr><th class="cb"></th>` +
    fields.map((f) => {
      const ind = (so && so.field === f.name) ? (so.dir === 1 ? "▲" : "▼") : "";
      const act = cfT[f.name] ? " filt-on" : "";
      return `<th class="col-h${act}" data-field="${esc(f.name)}" title="${esc(f.name)}（点击排序）">` +
        `<span class="col-h-name">${esc(f.name)}${f.derived ? " (自动)" : ""}</span>` +
        `<span class="sort-ind">${ind}</span>` +
        `<span class="col-filt" data-field="${esc(f.name)}" title="筛选">▾</span></th>`;
    }).join("") + "</tr>";
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
        if (v === "") {
          return `<td class="date-cell empty" data-datefield="${esc(f.name)}"><span class="date-ph" title="点击设置日期">—</span></td>`;
        }
        return `<td class="date-cell"><input type="date" data-f="${esc(f.name)}" value="${esc(v)}"></td>`;
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
  recomputeDerived();
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
  if (e.target.closest("td.cb") || e.target.tagName === "INPUT" || e.target.tagName === "SELECT"
      || e.target.closest("td.date-cell.empty")) return;
  const tr = e.target.closest("tr");
  if (!tr) return;
  e.preventDefault();
  selectRow(Number(tr.dataset.ri), e.shiftKey || e.metaKey || e.ctrlKey);
});

// 空日期单元格：点击才挂日期选择器（macOS 原生 picker 会把空值显示成"今天"，造成"假数据"错觉）
$("gridBody").addEventListener("click", (e) => {
  const td = e.target.closest("td.date-cell.empty");
  if (!td) return;
  const field = td.dataset.datefield;
  td.classList.remove("empty");
  td.innerHTML = `<input type="date" data-f="${esc(field)}" value="">`;
  td.querySelector("input").focus();
});
$("gridBody").addEventListener("focusout", (e) => {
  const inp = e.target;
  if (inp.tagName === "INPUT" && inp.type === "date" && inp.value === "") {
    const td = inp.closest("td.date-cell");
    if (td && !td.classList.contains("empty")) {
      td.classList.add("empty");
      td.innerHTML = `<span class="date-ph" title="点击设置日期">—</span>`;
    }
  }
});

document.querySelectorAll(".toolbar [data-act]").forEach((b) => {
  b.addEventListener("click", () => {
    const act = b.dataset.act;
    const table = state.tab;
    if (act === "add") {
      const ctx = currentContext();
      const rec = {};
      if (ctx.job) rec["JOB No"] = ctx.job;
      if (table === "合同订单" && ctx.customer) rec["客户"] = ctx.customer;
      state.data[table].push(rec);
      state.dirty = true;
      state.selection = new Set([state.data[table].length - 1]);
      recomputeDerived(); renderGrid(); renderCounts();
      $("gridWrap").scrollTop = $("gridWrap").scrollHeight;
    } else if (act === "del") {
      const idxs = [...state.selection].sort((a, b) => b - a);
      for (const i of idxs) state.data[table].splice(i, 1);
      state.selection = new Set();
      state.dirty = true;
      recomputeDerived(); renderGrid(); renderCounts();
    } else if (act === "dup") {
      const idxs = [...state.selection].sort((a, b) => a - b);
      const news = [];
      for (const i of idxs) news.push({ ...state.data[table][i] });
      const at = idxs.length ? idxs[idxs.length - 1] + 1 : state.data[table].length;
      state.data[table].splice(at, 0, ...news);
      state.selection = new Set(news.map((_, k) => at + k));
      state.dirty = true;
      recomputeDerived(); renderGrid(); renderCounts();
      toast(`已重复 ${news.length} 行`);
    } else if (act === "copy") {
      copySelection(table);
    } else if (act === "export") {
      openExportModal();
    } else if (act === "rename" && table === "发货批次") {
      renameBatch();
    } else if (act === "split" && table === "发货批次") {
      splitBatch();
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
  recomputeDerived(); renderGrid(); renderCounts();
  toast(`已重命名批次 ${oldName} → ${nw}（同步 ${touched} 处）`, "ok");
}

function splitBatch() {
  const table = "发货批次";
  const idxs = [...state.selection];
  if (!idxs.length) { toast("请先选择要拆分的批次行", "error"); return; }
  const rec = state.data[table][idxs[0]];
  const job = s(rec["JOB No"]), src = s(rec["发货批次"]);
  if (!job || !src) { toast("该行缺少 JOB No / 发货批次", "error"); return; }
  const devIdxs = state.data["设备台账"]
    .map((d, i) => ({ d, i }))
    .filter((x) => s(x.d["JOB No"]) === job && s(x.d["发货批次"]) === src)
    .map((x) => x.i);
  if (!devIdxs.length) { toast(`批次「${src}」无设备`, "error"); return; }
  const groups = {};
  for (const i of devIdxs) {
    const m = s(state.data["设备台账"][i]["设备型号"]) || "（无型号）";
    (groups[m] ||= []).push(i);
  }
  const existing = new Set(state.data["发货批次"]
    .filter((b) => s(b["JOB No"]) === job).map((b) => s(b["发货批次"])));
  let n = 1; while (existing.has(String(n))) n++;
  $("splitTitle").textContent = `拆分 ${job} 批次「${src}」（共 ${devIdxs.length} 台）→ 勾选要移出的设备`;
  $("splitTarget").value = String(n);
  $("splitShipDate").value = s(rec["出荷日"]);
  $("splitDevList").innerHTML = Object.entries(groups).map(([m, list]) => {
    const rows = list.map((i) => {
      const d = state.data["设备台账"][i];
      const price = Number(d["未税单价"]) || 0;
      return `<label class="cf-item split-dev" data-idx="${i}"><input type="checkbox"><span class="cf-v">${esc(s(d["製造番号"]) || "(无番号)")}</span><span class="cf-c">${price.toLocaleString()}</span></label>`;
    }).join("");
    return `<div class="split-group"><label class="split-group-h"><input type="checkbox" class="split-grp-all"> <b>${esc(m)}</b>（${list.length} 台）</label><div class="split-group-body">${rows}</div></div>`;
  }).join("");
  state._split = { job, src, rec, devIdxs, groups };
  updateSplitPreview();
  $("splitModal").classList.remove("hidden");
}
function splitSelected() {
  return [...document.querySelectorAll("#splitDevList .split-dev input:checked")]
    .map((c) => Number(c.closest(".split-dev").dataset.idx));
}
function updateSplitPreview() {
  if (!state._split) return;
  const { devIdxs, groups } = state._split;
  const selset = new Set(splitSelected());
  const moved = [], remain = [];
  for (const [m, list] of Object.entries(groups)) {
    const mv = list.filter((i) => selset.has(i)).length;
    moved.push(`${m}:${mv}`); remain.push(`${m}:${list.length - mv}`);
  }
  const n = selset.size;
  $("splitPreview").textContent =
    `将移动 ${n} 台（${moved.join("，") || "无"}）；原批次剩 ${devIdxs.length - n} 台（${remain.join("，")}）`;
}
function closeSplit() { $("splitModal").classList.add("hidden"); state._split = null; }

$("splitDevList").addEventListener("change", (e) => {
  const grpAll = e.target.closest(".split-grp-all");
  if (grpAll) {
    grpAll.closest(".split-group").querySelectorAll(".split-dev input")
      .forEach((c) => { c.checked = grpAll.checked; });
  } else if (e.target.closest(".split-dev")) {
    const g = e.target.closest(".split-group");
    const boxes = g.querySelectorAll(".split-dev input");
    g.querySelector(".split-grp-all").checked = [...boxes].every((c) => c.checked);
  }
  updateSplitPreview();
});
$("splitOk").addEventListener("click", () => {
  const sp = state._split; if (!sp) return;
  const target = ($("splitTarget").value || "").trim();
  const shipDate = $("splitShipDate").value || "";
  if (!target) { toast("请填目标批次名", "error"); return; }
  if (target === sp.src) { toast("目标批次不能与源批次相同", "error"); return; }
  const sel = splitSelected();
  if (!sel.length) { toast("请勾选至少 1 台设备", "error"); return; }
  for (const i of sel) state.data["设备台账"][i]["发货批次"] = target;
  let targetRow = state.data["发货批次"].find((b) => s(b["JOB No"]) === sp.job && s(b["发货批次"]) === target);
  if (!targetRow) {
    state.data["发货批次"].push({ "JOB No": sp.job, "发货批次": target, "出荷日": shipDate });
  } else if (shipDate) {
    targetRow["出荷日"] = shipDate;
  }
  recomputeDerived();
  state.dirty = true;
  closeSplit();
  renderGrid(); renderCounts();
  toast(`已把 ${sel.length} 台设备移到批次「${target}」`, "ok");
});
$("splitCancel").addEventListener("click", closeSplit);
$("splitClose").addEventListener("click", closeSplit);
document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeSplit(); });

// 导出筛选结果到 Excel（可选字段）
function openExportModal() {
  const table = state.tab;
  const fields = fieldsOf(table);
  const n = visibleRows(table).length;
  $("exportTitle").textContent = `导出「${table}」筛选结果（${n} 行）`;
  $("exportFields").innerHTML = fields.map((f) =>
    `<label class="cf-item"><input type="checkbox" data-f="${esc(f.name)}" checked><span class="cf-v">${esc(f.name)}${f.derived ? " (自动)" : ""}</span></label>`
  ).join("");
  $("exportAll").checked = true;
  updateExportCount();
  $("exportModal").classList.remove("hidden");
}
function updateExportCount() {
  const n = visibleRows(state.tab).length;
  const nf = document.querySelectorAll("#exportFields input:checked").length;
  $("exportCount").textContent = `将导出 ${n} 行 × ${nf} 个字段  →  ~/Downloads/<表>_导出_<时间>.xlsx`;
}
function closeExport() { $("exportModal").classList.add("hidden"); }
function doExport() {
  const table = state.tab;
  const fields = [...document.querySelectorAll("#exportFields input:checked")].map((c) => c.dataset.f);
  if (!fields.length) { toast("请至少勾选一个字段", "error"); return; }
  const rows = visibleRows(table).map((i) => state.data[table][i]);
  if (!rows.length) { toast("当前筛选结果为空", "error"); return; }
  setStatus(`导出 ${rows.length} 行…`);
  call("export_xlsx", table, rows, fields).then((r) => {
    closeExport();
    setStatus(`已导出：${r.path}（${r.rows} 行 × ${r.fields} 列）`);
    toast(`已导出 ${r.rows} 行到 Downloads`, "ok");
    return call("open_path", r.path);
  }).catch((e) => { setStatus(String(e), "error"); toast(String(e), "error"); });
}
$("exportFields").addEventListener("change", () => {
  const boxes = document.querySelectorAll("#exportFields input");
  $("exportAll").checked = [...boxes].every((c) => c.checked);
  updateExportCount();
});
$("exportAll").addEventListener("change", (e) => {
  document.querySelectorAll("#exportFields input").forEach((c) => { c.checked = e.target.checked; });
  updateExportCount();
});
$("exportOk").addEventListener("click", doExport);
$("exportCancel").addEventListener("click", closeExport);
$("exportClose").addEventListener("click", closeExport);
document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeExport(); });

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
      const ctx = currentContext();
      const rec = {};
      if (ctx.job) rec["JOB No"] = ctx.job;
      if (table === "合同订单" && ctx.customer) rec["客户"] = ctx.customer;
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
  recomputeDerived(); renderGrid(); renderCounts();
  toast(`已粘贴 ${rows.length} 行`);
}

$("btnClearFilter").addEventListener("click", () => {
  state.colFilters[state.tab] = {};
  delete state.sortBy[state.tab];
  state.filter = "";
  $("filterKeyword").value = "";
  closeColFilter();
  renderGrid();
  toast("已清除本表筛选/排序");
});

// 表头：点列名=排序(升→降→取消)；点 ▾=打开该列筛选
$("gridHead").addEventListener("click", (e) => {
  const fb = e.target.closest(".col-filt");
  if (fb) { openColFilter(fb.dataset.field); return; }
  const th = e.target.closest("th.col-h");
  if (!th) return;
  const field = th.dataset.field;
  const cur = state.sortBy[state.tab];
  let dir;
  if (!cur || cur.field !== field) dir = 1;
  else if (cur.dir === 1) dir = -1;
  else dir = 0;
  if (dir) state.sortBy[state.tab] = { field, dir };
  else delete state.sortBy[state.tab];
  renderGrid();
});

let _cfAllVals = [];
function openColFilter(field) {
  const table = state.tab;
  const recs = state.data[table];
  const dist = new Map();
  for (const r of recs) {
    const k = String(r[field] ?? "");
    dist.set(k, (dist.get(k) || 0) + 1);
  }
  _cfAllVals = [...dist.keys()].sort((a, b) => a.localeCompare(b, "zh"));
  const allCount = _cfAllVals.length;
  state.colFilters[table] = state.colFilters[table] || {};
  const set = state.colFilters[table][field] || new Set(_cfAllVals);
  $("colfiltTitle").textContent = `${field}（${allCount} 个值）`;
  $("colfiltSearch").value = "";
  $("colfiltAll").checked = set.size >= allCount;
  $("colfiltList").innerHTML = _cfAllVals.map((v) => {
    const lbl = v === "" ? "（空）" : esc(v);
    return `<label class="cf-item"><input type="checkbox" data-v="${esc(v)}" ${set.has(v) ? "checked" : ""}><span class="cf-v">${lbl}</span></label>`;
  }).join("");
  const pop = $("colFilterPop");
  pop.dataset.field = field;
  pop.style.display = "block";
  const th = document.querySelector(`#gridHead th[data-field="${field}"]`);
  if (th) {
    const r = th.getBoundingClientRect();
    pop.style.top = Math.min(r.bottom + 2, window.innerHeight - 340) + "px";
    let left = r.left;
    if (left + 260 > window.innerWidth) left = window.innerWidth - 270;
    pop.style.left = Math.max(8, left) + "px";
  }
}
function cfApply(field, value, checked) {
  const table = state.tab;
  state.colFilters[table] = state.colFilters[table] || {};
  let set = state.colFilters[table][field];
  if (!set) { set = new Set(_cfAllVals); state.colFilters[table][field] = set; }
  if (checked) set.add(value); else set.delete(value);
  if (set.size >= _cfAllVals.length) delete state.colFilters[table][field];
  $("colfiltAll").checked = !state.colFilters[table] || !state.colFilters[table][field];
  renderGrid();
}
function closeColFilter() { $("colFilterPop").style.display = "none"; }

$("colfiltSearch").addEventListener("input", (e) => {
  const q = e.target.value.toLowerCase();
  document.querySelectorAll("#colfiltList .cf-item").forEach((el) => {
    el.style.display = el.querySelector(".cf-v").textContent.toLowerCase().includes(q) ? "" : "none";
  });
});
$("colfiltAll").addEventListener("change", (e) => {
  const field = $("colFilterPop").dataset.field;
  const table = state.tab;
  const checked = e.target.checked;
  state.colFilters[table] = state.colFilters[table] || {};
  if (checked) delete state.colFilters[table][field];
  else state.colFilters[table][field] = new Set();
  document.querySelectorAll("#colfiltList .cf-item input").forEach((c) => { c.checked = checked; });
  renderGrid();
});
$("colfiltList").addEventListener("change", (e) => {
  if (e.target.tagName !== "INPUT") return;
  cfApply($("colFilterPop").dataset.field, e.target.dataset.v, e.target.checked);
});
$("colfiltClear").addEventListener("click", () => {
  const field = $("colFilterPop").dataset.field;
  if (state.colFilters[state.tab]) delete state.colFilters[state.tab][field];
  closeColFilter();
  renderGrid();
});
$("colfiltClose").addEventListener("click", closeColFilter);
document.addEventListener("mousedown", (e) => {
  const pop = $("colFilterPop");
  if (!pop || pop.style.display === "none") return;
  if (e.target.closest("#colFilterPop") || e.target.closest(".col-filt")) return;
  closeColFilter();
});
document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeColFilter(); });

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
    if (u.current) { state.version = u.current; $("setVersion").textContent = "v" + u.current; }
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
