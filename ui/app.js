"use strict";

const state = {
  storePath: "", xlsxPath: "", data: null, schema: null,
  tab: "摘要", dirty: false, selection: new Set(), anchor: null,
  filter: "", colFilters: {}, sortBy: {}, rules: {}, jobFilter: null,
  version: "?", update: null, databaseInfo: null, pendingImport: null,
};

const ROW_H = 31, BUFFER = 20;

const TAB_HINTS = {
  "合同订单": "本页 = 合同头信息。总台数、设备型号和付款条件文本自动汇总；修改 JOB No 会同步更新六表关联。推荐用顶部“新建订单”。",
  "付款条件": "本页 = 结构化付款方式（款类/比例/账期/触发条件/说明），比例合计须 100%。说明会自动汇总到合同页的付款条件文本。",
  "设备台账": "本页 = 每台设备一行（含发货批次归属、质保时间、验收状态）。改设备属于哪个批次就在这里改“发货批次”列。",
  "发货批次": "本页改出荷日或批次名；直接改名与“重命名批次”都会同步设备及开票/回款覆盖批次。台数、合计和覆盖番号自动计算。",
  "开票记录": "本页 = 每笔开票一行（款类/开票日/金额/覆盖批次/覆盖番号/账期/应收回款日/回款状态）。覆盖台数自动算。",
  "回款记录": "本页 = 每笔回款一行（款类/回款日/金额/覆盖批次/覆盖番号）。覆盖台数自动算。",
};

const $ = (id) => document.getElementById(id);
function s(v) { return v === null || v === undefined ? "" : String(v); }
function coverageValues(value) {
  return s(value).split(/[;；\r\n]+/).map((x) => x.trim()).filter(Boolean);
}
let _localIdCounter = 0;
function newLocalId() {
  _localIdCounter += 1;
  return `local-${Date.now()}-${_localIdCounter}`;
}

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
  const setBadge = (id, ok, yes, no) => {
    const b = $(id); b.textContent = ok ? yes : no; b.className = ok ? "ok" : "no";
  };
  setBadge("badgeStore", !!state.storePath, "已连接", "未连接");
  setBadge("badgeXlsx", !!state.xlsxPath, "已选择", "未选择");
  $("setStore").textContent = state.storePath || "（未选择）";
  $("setXlsx").textContent = state.xlsxPath || "（未选择）";
  $("setVersion").textContent = "v" + state.version;
  const db = state.databaseInfo || {};
  $("setRevision").textContent = db.revision ?? "—";
  $("setIntegrity").textContent = db.integrity === "ok" ? "正常" : (db.integrity || "—");
  $("setLastSaved").textContent = db.last_saved_at ? db.last_saved_at.replace("T", " ") : "—";
  $("databaseSummary").textContent = db.integrity === "ok"
    ? `本地关系数据库完整性正常 · 修订 ${db.revision ?? 0} · 六表关联已启用外键保护`
    : "数据库尚未完成完整性检查";
  renderSummary();
  renderGrid();
  renderCounts();
}

function currentContext() {
  // 从当前表的单值列筛选推导"新建行的上下文"（替代旧的客户/JOB 下拉）
  const cf = state.colFilters[state.tab] || {};
  const ctx = { job: "", customer: "" };
  if (state.jobFilter && state.jobFilter.size === 1) ctx.job = [...state.jobFilter][0];
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
  ].map(([k, v]) => {
    const jump = state.schema[k] ? ` data-jump="${esc(k)}"` : "";
    const cls = k === "错误" && v ? " alert" : (k === "警告" && v ? " warn" : "");
    return `<div class="card${cls}"${jump}><div class="k">${k}</div><div class="v">${v}</div></div>`;
  }).join("");
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
  const revision = state.databaseInfo?.revision;
  $("counts").textContent = parts.join(" ｜ ") +
    (revision !== undefined ? ` ｜ DB r${revision}` : "") +
    (state.dirty ? " ｜ ● 未保存" : "");
  const saveState = $("saveState");
  if (saveState) {
    saveState.classList.toggle("dirty", state.dirty);
    saveState.classList.toggle("saved", !state.dirty);
    saveState.lastChild.textContent = state.dirty ? "有未保存更改" : "已保存";
  }
}

function validateAll() {
  const issues = [];
  if (!state.data) return issues;
  const jobs = new Set();
  for (const c of state.data["合同订单"]) {
    const job = String(c["JOB No"] || "");
    if (!/^\d{2}(BS|DS)\d{3}$/.test(job)) {
      issues.push({ severity: "error", msg: `合同 JOB No 格式不正确：${c["JOB No"]}` });
    }
    if (jobs.has(job)) issues.push({ severity: "error", msg: `合同 JOB No 重复：${job}` });
    jobs.add(job);
  }
  for (const table of ["付款条件", "设备台账", "发货批次", "开票记录", "回款记录"]) {
    for (const rec of state.data[table]) {
      const job = String(rec["JOB No"] || "");
      if (!job) issues.push({ severity: "error", msg: `${table} 存在缺少 JOB No 的记录` });
      else if (!jobs.has(job)) issues.push({ severity: "error", msg: `${table} 引用了不存在的 JOB No：${job}` });
    }
  }
  for (const table of Object.keys(state.schema)) {
    for (const field of fieldsOf(table)) {
      if (field.type !== "select" || !field.options?.length) continue;
      const allowed = new Set(field.options);
      for (const rec of state.data[table]) {
        const value = s(rec[field.name]);
        if (value && !allowed.has(value)) {
          issues.push({ severity: "error", msg: `${table} ${field.name}值不在允许列表中：${value}` });
        }
      }
    }
  }
  const sums = {};
  for (const t of state.data["付款条件"]) {
    if (!s(t["款类"])) continue; // 空占位行不参与比例校验
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
      issues.push({ severity: "warning", msg: `${job} 设备台账存在空製造番号（新建占位行待补）` });
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
  const shipmentKeys = new Set();
  for (const x of state.data["发货批次"]) {
    const key = s(x["JOB No"]) + "|" + s(x["发货批次"]);
    if (!s(x["发货批次"])) issues.push({ severity: "error", msg: `${s(x["JOB No"])} 发货批次名称不能为空` });
    if (shipmentKeys.has(key)) issues.push({ severity: "error", msg: `发货批次业务键重复：${s(x["JOB No"])} ${s(x["发货批次"])}` });
    shipmentKeys.add(key);
  }
  const devBatches = new Set(state.data["设备台账"].map((d) => d["JOB No"] + "|" + d["发货批次"]));
  for (const d of state.data["设备台账"]) {
    const batch = s(d["发货批次"]), key = s(d["JOB No"]) + "|" + batch;
    if (batch && !shipmentKeys.has(key)) {
      issues.push({ severity: "error", msg: `${s(d["JOB No"])} 设备引用了不存在的发货批次：${batch}` });
    }
  }
  const deviceSerials = new Set(state.data["设备台账"]
    .filter((d) => s(d["製造番号"]))
    .map((d) => s(d["JOB No"]) + "|" + s(d["製造番号"])));
  for (const table of ["开票记录", "回款记录"]) {
    for (const rec of state.data[table]) {
      const job = s(rec["JOB No"]);
      for (const batch of coverageValues(rec["覆盖批次"])) {
        if (!shipmentKeys.has(job + "|" + batch)) {
          issues.push({ severity: "error", msg: `${table} ${job} 覆盖了不存在的批次：${batch}` });
        }
      }
      for (const serial of coverageValues(rec["覆盖製造番号"])) {
        if (!deviceSerials.has(job + "|" + serial)) {
          issues.push({ severity: "error", msg: `${table} ${job} 覆盖了不存在的製造番号：${serial}` });
        }
      }
    }
  }
  for (const x of state.data["发货批次"]) {
    if (!devBatches.has(x["JOB No"] + "|" + x["发货批次"])) {
      issues.push({ severity: "warning", msg: `${x["JOB No"]} 批次 ${x["发货批次"]} 无设备，生成后不会显示（请补设备或删除）` });
    }
  }
  for (const table of Object.keys(state.schema)) {
    const dateFields = fieldsOf(table).filter((f) => f.type === "date").map((f) => f.name);
    for (const rec of state.data[table]) {
      for (const field of dateFields) {
        const value = s(rec[field]);
        if (!value) continue;
        const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        const dt = match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])) : null;
        if (!dt || dt.getFullYear() !== Number(match[1]) || dt.getMonth() !== Number(match[2]) - 1 || dt.getDate() !== Number(match[3])) {
          issues.push({ severity: "error", msg: `${s(rec["JOB No"])} ${field}=${value} 不是有效日期 YYYY-MM-DD` });
        }
      }
    }
  }
  return issues;
}

function switchTab(tab) {
  if (state.tab !== tab) {
    state.selection = new Set();
    state.anchor = null;
    closeColFilter();
  }
  state.tab = tab;
  $("panel-设置").classList.add("hidden");
  $("panel-跨表查询").classList.add("hidden");
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

$("summaryCards").addEventListener("click", (e) => {
  const card = e.target.closest(".card[data-jump]");
  if (card) switchTab(card.dataset.jump);
});

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
      if (f === "JOB No") continue; // JOB 走全局 jobFilter，跨表共享
      const set = cf[f];
      if (set && !set.has(String(rec[f] ?? ""))) { keep = false; break; }
    }
    if (keep && state.jobFilter && !state.jobFilter.has(String(rec["JOB No"] ?? ""))) keep = false;
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
  const termsByJob = {};
  for (const t of d["付款条件"]) (termsByJob[s(t["JOB No"])] ||= []).push(t);
  for (const dev of d["设备台账"]) {
    const j = s(dev["JOB No"]);
    cnt[j] = (cnt[j] || 0) + 1;
    (models[j] ||= new Set()).add(s(dev["设备型号"]));
    dev["验收状态"] = s(dev["质保开始日"]) ? "已验收" : "未验收";
  }
  for (const c of d["合同订单"]) {
    const j = s(c["JOB No"]);
    c["总台数"] = cnt[j] || 0;
    if (!s(c["设备型号"])) c["设备型号"] = [...(models[j] || new Set())].filter((m) => m).sort().join(";");
    const terms = termsByJob[j] || [];
    if (terms.length) {
      const desc = terms.map((t) => s(t["说明"]).trim()).filter(Boolean).join("；");
      if (desc) c["付款条件"] = desc;
    }
  }
  for (const t of ["开票记录", "回款记录"]) {
    for (const x of d[t]) {
      const set = new Set(coverageValues(x["覆盖製造番号"]));
      x["覆盖台数"] = set.size;
      if (t === "开票记录") {
        const m = s(x["开票日"]).match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) {
          const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]) + (parseInt(x["账期天数"], 10) || 0));
          x["应收回款日"] = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
        }
      }
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
      const act = (f.name === "JOB No" ? state.jobFilter : cfT[f.name]) ? " filt-on" : "";
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
  let h = TAB_HINTS[table] || "";
  if (state.jobFilter) h += `　［跨表 JOB 筛选：${[...state.jobFilter].map(esc).join("、")}］`;
  hint.textContent = h;
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
  const rec = state.data[table][ri];
  const previous = s(rec[field]);
  let v = raw;
  if (f.type === "number") v = raw === "" ? null : Number(raw);
  else if (f.type === "int") v = raw === "" ? null : Math.round(Number(raw));
  rec[field] = v;
  if (table === "合同订单" && field === "JOB No" && previous !== s(v)) {
    for (const child of ["付款条件", "设备台账", "发货批次", "开票记录", "回款记录"]) {
      for (const row of state.data[child]) {
        if (s(row["JOB No"]) === previous) row["JOB No"] = v;
      }
    }
    if (state.jobFilter?.has(previous)) {
      state.jobFilter.delete(previous);
      state.jobFilter.add(s(v));
    }
  }
  if (table === "发货批次" && field === "发货批次" && previous !== s(v)) {
    const job = s(rec["JOB No"]), next = s(v);
    for (const device of state.data["设备台账"]) {
      if (s(device["JOB No"]) === job && s(device["发货批次"]) === previous) device["发货批次"] = next;
    }
    for (const child of ["开票记录", "回款记录"]) {
      for (const row of state.data[child]) {
        if (s(row["JOB No"]) !== job) continue;
        const batches = coverageValues(row["覆盖批次"]);
        if (batches.includes(previous)) row["覆盖批次"] = batches.map((x) => x === previous ? next : x).join(";");
      }
    }
  }
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
    renderGrid();
    renderSummary();
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
      const rec = { "记录ID": newLocalId() };
      if (ctx.job) rec["JOB No"] = ctx.job;
      if (table === "合同订单" && ctx.customer) rec["客户"] = ctx.customer;
      state.data[table].push(rec);
      state.dirty = true;
      state.selection = new Set([state.data[table].length - 1]);
      recomputeDerived(); renderGrid(); renderCounts();
      $("gridWrap").scrollTop = $("gridWrap").scrollHeight;
    } else if (act === "del") {
      const idxs = [...state.selection].sort((a, b) => b - a);
      if (!idxs.length) { toast("请先选择要删除的行", "error"); return; }
      if (table === "合同订单") { tryCascadeDelete(idxs); return; }
      for (const i of idxs) state.data[table].splice(i, 1);
      state.selection = new Set();
      state.dirty = true;
      recomputeDerived(); renderGrid(); renderCounts();
    } else if (act === "dup") {
      const idxs = [...state.selection].sort((a, b) => a - b);
      if (!idxs.length) { toast("请先选择要重复的行", "error"); return; }
      const news = [];
      for (const i of idxs) news.push({ ...state.data[table][i], "记录ID": newLocalId() });
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

function tryCascadeDelete(idxs) {
  const jobs = [...new Set(idxs.map((i) => String(state.data["合同订单"][i]?.["JOB No"] || "")).filter(Boolean))];
  if (!jobs.length) {
    for (const i of idxs) state.data["合同订单"].splice(i, 1);
    state.selection = new Set();
    state.dirty = true;
    recomputeDerived(); renderGrid(); renderCounts(); renderSummary();
    toast("已删除无 JOB No 的合同草稿行");
    return true;
  }
  const list = jobs.join("、");
  if (!confirm("确定删除订单 " + list + "？\n这将删除该 JOB 在【全部六表】的所有数据（合同/付款条件/设备台账/发货批次/开票/回款）。")) return false;
  if (!confirm("再次确认：删除 " + list + " 的全部数据？保存数据库前仍可通过关闭程序放弃本次更改。")) return false;
  const jset = new Set(jobs);
  for (const t of Object.keys(state.schema)) {
    state.data[t] = state.data[t].filter((r) => !jset.has(String(r["JOB No"] || "")));
  }
  state.selection = new Set();
  state.jobFilter = null; // 删除的 JOB 可能正被筛，清筛选看全貌
  state.dirty = true;
  recomputeDerived(); renderGrid(); renderCounts(); renderSummary();
  toast("已删除 " + jobs.length + " 个订单（六表全部数据）");
  return true;
}

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
      const bs = coverageValues(r["覆盖批次"]);
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
    state.data["发货批次"].push({ "记录ID": newLocalId(), "JOB No": sp.job, "发货批次": target, "出荷日": shipDate });
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

// ============ 跨表查询导出（report builder）============
const PARENTS = {
  "设备台账": ["合同订单", "发货批次"],
  "发货批次": ["合同订单"],
  "付款条件": ["合同订单"],
  "开票记录": ["合同订单", "付款条件"],
  "回款记录": ["合同订单", "付款条件"],
  "合同订单": [],
};
const SHORT = { "合同订单": "合同", "发货批次": "批次", "付款条件": "条款", "设备台账": "设备", "开票记录": "开票", "回款记录": "回款" };
let _contractByJob = {}, _batchByJobBatch = {}, _termByJobKind = {};
function buildLookups() {
  _contractByJob = {};
  for (const c of state.data["合同订单"]) _contractByJob[s(c["JOB No"])] = c;
  _batchByJobBatch = {};
  for (const b of state.data["发货批次"]) _batchByJobBatch[s(b["JOB No"]) + "|" + s(b["发货批次"])] = b;
  _termByJobKind = {};
  for (const t of state.data["付款条件"]) _termByJobKind[s(t["JOB No"]) + "|" + s(t["款类"])] = t;
}
function getJoined(base, parent, rec) {
  if (parent === "合同订单") return _contractByJob[s(rec["JOB No"])];
  if (parent === "发货批次") return _batchByJobBatch[s(rec["JOB No"]) + "|" + s(rec["发货批次"])];
  if (parent === "付款条件") return _termByJobKind[s(rec["JOB No"]) + "|" + s(rec["款类"])];
  return null;
}
function qbTables(base) { return [base, ...(PARENTS[base] || [])]; }
function qbFieldList(base) {
  const out = [];
  for (const t of qbTables(base)) {
    const isBase = (t === base);
    for (const f of (state.schema[t] || [])) {
      const name = f[0], type = f[1];
      out.push({ table: t, name, type, label: isBase ? name : `${SHORT[t] || t}.${name}` });
    }
  }
  return out;
}
function openQueryBuilder() {
  if (!state.data) return;
  buildLookups();
  const tables = Object.keys(state.schema);
  $("qbBase").innerHTML = tables.map((t) => `<option value="${esc(t)}">${esc(t)}</option>`).join("");
  if (tables.includes("设备台账")) $("qbBase").value = "设备台账";
  renderQueryBuilder();
}
function renderQueryBuilder() {
  const base = $("qbBase").value;
  const fields = qbFieldList(base);
  $("qbFields").innerHTML = qbTables(base).map((t) => {
    const isBase = (t === base);
    const items = fields.filter((f) => f.table === t).map((f) =>
      `<label class="cf-item"><input type="checkbox" data-t="${esc(f.table)}" data-n="${esc(f.name)}" data-type="${f.type}" data-label="${esc(f.label)}"><span class="cf-v">${esc(f.label)}</span></label>`
    ).join("");
    return `<div class="qb-group"><div class="qb-group-h">${isBase ? esc(t) + "（基础表）" : "← " + esc(t)}</div><div class="qb-group-body">${items}</div></div>`;
  }).join("");
  $("qbFilters").innerHTML = "";
  $("qbPreview").innerHTML = "";
  $("qbCount").textContent = "";
}
function qbAddFilterRow() {
  const fields = qbFieldList($("qbBase").value);
  const opts = fields.map((f) =>
    `<option data-t="${esc(f.table)}" data-n="${esc(f.name)}" data-type="${f.type}">${esc(f.label)}</option>`).join("");
  const tr = document.createElement("div");
  tr.className = "qb-filter";
  tr.innerHTML = `<select class="qb-f-field">${opts}</select>
    <select class="qb-f-op">
      <option value="eq">等于</option><option value="ne">不等于</option>
      <option value="gt">大于</option><option value="lt">小于</option>
      <option value="empty">为空</option><option value="notempty">不为空</option>
    </select>
    <select class="qb-f-val"></select>
    <button class="qb-f-del" title="删除条件">✕</button>`;
  $("qbFilters").appendChild(tr);
  const f0 = fields[0];
  if (f0) populateValueSelect(tr.querySelector(".qb-f-val"), f0.table, f0.name);
}
function distinctValues(table, name) {
  const set = new Set();
  for (const r of (state.data[table] || [])) {
    const v = s(r[name]);
    if (v !== "") set.add(v);
  }
  const arr = [...set];
  if (arr.length && arr.every((v) => !isNaN(Number(v)))) arr.sort((a, b) => Number(a) - Number(b));
  else arr.sort((a, b) => a.localeCompare(b, "zh"));
  return arr;
}
function populateValueSelect(sel, table, name) {
  const vals = distinctValues(table, name);
  sel.innerHTML = `<option value="">（选值…）</option>` +
    vals.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
}
function readFilters() {
  return [...document.querySelectorAll("#qbFilters .qb-filter")].map((tr) => {
    const opt = tr.querySelector(".qb-f-field").selectedOptions[0];
    return {
      table: opt.dataset.t, name: opt.dataset.n, type: opt.dataset.type,
      op: tr.querySelector(".qb-f-op").value,
      value: tr.querySelector(".qb-f-val").value,
    };
  });
}
function _numOrStr(x) { const n = Number(x); return isNaN(n) ? s(x) : n; }
function matchFilter(value, flt) {
  const v = s(value);
  if (flt.op === "empty") return v === "";
  if (flt.op === "notempty") return v !== "";
  const target = s(flt.value);
  if (target === "") return true;  // 未选值 → 该条件不生效（避免空筛）
  if (flt.op === "eq") return v === target;
  if (flt.op === "ne") return v !== target;
  if (flt.op === "gt") return _numOrStr(v) > _numOrStr(target);
  if (flt.op === "lt") return _numOrStr(v) < _numOrStr(target);
  return true;
}
function runQuery() {
  const base = $("qbBase").value;
  const selFields = [...document.querySelectorAll("#qbFields input:checked")].map((c) => ({
    table: c.dataset.t, name: c.dataset.n, type: c.dataset.type, label: c.dataset.label,
  }));
  const filters = readFilters();
  const recs = state.data[base] || [];
  const rows = [];
  for (const r of recs) {
    const get = (table, name) => (table === base ? r[name] : (getJoined(base, table, r) || {})[name]);
    let keep = true;
    for (const flt of filters) {
      if (!matchFilter(get(flt.table, flt.name), flt)) { keep = false; break; }
    }
    if (keep) {
      const row = {};
      for (const f of selFields) row[f.label] = get(f.table, f.name);
      rows.push(row);
    }
  }
  return { fields: selFields, rows };
}
function qbDoPreview() {
  const { fields, rows } = runQuery();
  if (!fields.length) { toast("请至少勾选一个输出字段", "error"); return; }
  $("qbCount").textContent = `符合 ${rows.length} 行 × ${fields.length} 列`;
  const head = `<tr>${fields.map((f) => `<th>${esc(f.label)}</th>`).join("")}</tr>`;
  const body = rows.slice(0, 10).map((r) =>
    `<tr>${fields.map((f) => `<td>${esc(r[f.label])}</td>`).join("")}</tr>`).join("");
  $("qbPreview").innerHTML = `<thead>${head}</thead><tbody>${body}</tbody>`;
}
function qbDoExport() {
  const { fields, rows } = runQuery();
  if (!fields.length) { toast("请至少勾选一个输出字段", "error"); return; }
  if (!rows.length) { toast("查询结果为空", "error"); return; }
  const exportFields = fields.map((f) => ({ name: f.label, type: f.type }));
  setStatus(`导出 ${rows.length} 行…`);
  call("export_xlsx", $("qbBase").value, rows, exportFields).then((r) => {
    setStatus(`已导出：${r.path}（${r.rows} 行 × ${r.fields} 列）`);
    toast(`已导出 ${r.rows} 行到 Downloads`, "ok");
    return call("open_path", r.path);
  }).catch((e) => { setStatus(String(e), "error"); toast(String(e), "error"); });
}
$("btnQuery").addEventListener("click", () => {
  ["摘要", "表格", "使用指南", "预警规则", "设置"].forEach((p) => $("panel-" + p).classList.add("hidden"));
  $("panel-跨表查询").classList.remove("hidden");
  document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
  openQueryBuilder();
});
$("qbBase").addEventListener("change", renderQueryBuilder);
$("qbAddFilter").addEventListener("click", qbAddFilterRow);
$("qbRun").addEventListener("click", qbDoPreview);
$("qbExport").addEventListener("click", qbDoExport);
$("qbFilters").addEventListener("click", (e) => {
  const del = e.target.closest(".qb-f-del");
  if (del) del.closest(".qb-filter").remove();
});
$("qbFilters").addEventListener("change", (e) => {
  const fieldSel = e.target.closest(".qb-f-field");
  if (fieldSel) {
    const opt = fieldSel.selectedOptions[0];
    const tr = fieldSel.closest(".qb-filter");
    populateValueSelect(tr.querySelector(".qb-f-val"), opt.dataset.t, opt.dataset.n);
    return;
  }
  const opSel = e.target.closest(".qb-f-op");
  if (opSel) {
    const op = opSel.value;
    opSel.closest(".qb-filter").querySelector(".qb-f-val").style.display =
      (op === "empty" || op === "notempty") ? "none" : "";
  }
});

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
      const rec = { "记录ID": newLocalId() };
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
      if ((table === "合同订单" && field.name === "JOB No") ||
          (table === "发货批次" && field.name === "发货批次")) {
        updateRec(table, r, field.name, row[c]);
      } else {
        state.data[table][r][field.name] = field.type === "number"
          ? (row[c] === "" ? null : Number(row[c]))
          : field.type === "int" ? (row[c] === "" ? null : Math.round(Number(row[c]))) : row[c];
      }
    }
    r += 1;
  }
  state.dirty = true;
  recomputeDerived(); renderGrid(); renderCounts();
  toast(`已粘贴 ${rows.length} 行`);
}

$("btnClearFilter").addEventListener("click", () => {
  state.colFilters[state.tab] = {};
  state.jobFilter = null;
  delete state.sortBy[state.tab];
  state.filter = "";
  $("filterKeyword").value = "";
  closeColFilter();
  renderGrid();
  toast("已清除筛选/排序（含 JOB 跨表筛选）");
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
  const isJob = field === "JOB No";
  state.colFilters[table] = state.colFilters[table] || {};
  const set = isJob ? (state.jobFilter || new Set(_cfAllVals))
                    : (state.colFilters[table][field] || new Set(_cfAllVals));
  $("colfiltTitle").textContent = `${field}（${allCount} 个值${isJob ? "；此列跨表共享" : ""}）`;
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
  if (field === "JOB No") {
    let set = state.jobFilter ? new Set(state.jobFilter) : new Set(_cfAllVals);
    if (checked) set.add(value); else set.delete(value);
    state.jobFilter = set.size >= _cfAllVals.length ? null : set;
    $("colfiltAll").checked = !state.jobFilter;
    renderGrid();
    return;
  }
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
  const checked = e.target.checked;
  if (field === "JOB No") {
    state.jobFilter = checked ? null : new Set();
    document.querySelectorAll("#colfiltList .cf-item input").forEach((c) => { c.checked = checked; });
    renderGrid();
    return;
  }
  const table = state.tab;
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
  if (field === "JOB No") state.jobFilter = null;
  else if (state.colFilters[state.tab]) delete state.colFilters[state.tab][field];
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

function applyBackendState(r) {
  if (r.store_path || r.database_path) state.storePath = r.database_path || r.store_path;
  if (r.xlsx_path) state.xlsxPath = r.xlsx_path;
  if (r.data) state.data = r.data;
  if (r.schema) state.schema = r.schema;
  if (r.rules) state.rules = r.rules;
  if (r.version) state.version = r.version;
  if (r.database_info) state.databaseInfo = r.database_info;
  if (r.revision !== undefined) {
    state.databaseInfo = { ...(state.databaseInfo || {}), revision: r.revision,
      integrity: "ok", last_saved_at: new Date().toISOString().slice(0, 19) };
  }
}

async function saveCurrent(silent = false) {
  if (state._saving) return false;
  state._saving = true;
  try {
    if (!silent) setStatus("正在校验并保存数据库…");
    const r = await call("save_data", state.data, state.rules);
    applyBackendState(r);
    state.dirty = false;
    setStatus(`数据库已保存：${r.path}（修订 ${r.revision}）`);
    if (!silent) toast("数据库已通过校验并保存", "ok");
    renderAll();
    return true;
  } catch (err) {
    setStatus(String(err), "error");
    toast(String(err), "error");
    switchTab("摘要");
    renderSummary();
    return false;
  } finally {
    state._saving = false;
  }
}

$("btnSave").addEventListener("click", async () => {
  await saveCurrent(false);
});

$("btnGenerate").addEventListener("click", async () => {
  try {
    setStatus("生成台账中…");
    const r = await call("generate", state.data, state.rules);
    applyBackendState(r);
    state.dirty = false;
    const msg = `已生成：${r.out}（未回收 ${r.counts["未回收行数"]} 行，合计 ${Number(r.counts["未回收合计"]).toLocaleString()}）`;
    setStatus(msg);
    toast("台账已生成", "ok");
    renderAll();
  } catch (err) { setStatus(String(err), "error"); toast(String(err), "error"); }
});

$("btnPickStore").addEventListener("click", async () => {
  if (state.dirty && !confirm("当前有未保存更改。切换数据库会放弃这些更改，确定继续吗？")) return;
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

function diffValue(value) {
  if (value === null || value === undefined || value === "") return "（空）";
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
}

function closeImportPreview() {
  $("importModal").classList.add("hidden");
  $("importConfirm").checked = false;
  $("importApply").disabled = true;
  state.pendingImport = null;
}

function renderImportPreview(result) {
  state.pendingImport = result;
  const total = result.changes_total || 0;
  const blocked = result.stale || total === 0;
  const status = $("importStatus");
  status.classList.toggle("stale", result.stale);
  status.textContent = result.stale
    ? `此工作簿基于数据库修订 ${result.base_revision}，当前已是修订 ${result.current_revision}。为防止覆盖新数据，不能应用；请重新导出编辑副本。`
    : total ? `工作簿与数据库修订 ${result.current_revision} 匹配。请逐项核对后完成二次确认。`
      : "工作簿与当前数据库没有差异，无需导入。";
  const ac = result.action_counts || {};
  $("importCounts").innerHTML = [
    ["新增", ac["新增"] || 0, "add"], ["修改", ac["修改"] || 0, "change"],
    ["删除", ac["删除"] || 0, "delete"],
  ].map(([label, value, cls]) => `<div class="import-count ${cls}"><span>${label}</span><b>${value}</b></div>`).join("");
  $("importTableCounts").innerHTML = Object.entries(result.table_counts || {}).map(([table, counts]) => {
    const parts = ["新增", "修改", "删除"].filter((key) => counts[key]).map((key) => `${key}${counts[key]}`);
    return parts.length ? `<span>${esc(table)} · ${parts.join(" / ")}</span>` : "";
  }).join("");
  $("importFile").textContent = result.path || "";
  $("importChanges").innerHTML = (result.changes || []).map((change) => {
    const cls = change.action === "新增" ? "add" : (change.action === "删除" ? "delete" : "change");
    const fields = (change.fields || []).slice(0, 12).map((field) =>
      `<div><code>${esc(field.field)}${field.derived ? "（自动）" : ""}</code>：` +
      `${esc(diffValue(field.before))} → ${esc(diffValue(field.after))}</div>`).join("");
    return `<div class="import-change"><div class="import-change-head">` +
      `<span class="import-action ${cls}">${change.action}</span><b>${esc(change.table)}</b>` +
      `<span>${esc(change.label)}</span></div>` +
      (fields ? `<div class="import-change-fields">${fields}</div>` : "") + `</div>`;
  }).join("") || '<div class="import-change muted">没有变更</div>';
  if (result.changes_truncated) {
    $("importChanges").insertAdjacentHTML("beforeend", `<div class="import-change muted">仅展示前 500 项，共 ${total} 项；统计数量为完整结果。</div>`);
  }
  if (result.warnings?.length) {
    $("importChanges").insertAdjacentHTML("beforeend", result.warnings.slice(0, 20).map((warning) =>
      `<div class="import-change"><span class="import-action change">提示</span> ${esc(warning.msg)}</div>`).join(""));
  }
  $("importConfirm").checked = false;
  $("importConfirm").disabled = blocked;
  $("importConfirmLabel").classList.toggle("disabled", blocked);
  $("importApply").disabled = true;
  $("importModal").classList.remove("hidden");
}

$("btnExportEditable").addEventListener("click", async () => {
  try {
    setStatus("正在保存并生成可编辑副本…");
    const r = await call("export_editable", state.data, state.rules, state.dirty);
    applyBackendState(r);
    state.dirty = false;
    renderAll();
    setStatus(`已导出编辑副本：${r.path}（数据库修订 ${r.revision}）`);
    toast("六表编辑副本已导出到 Downloads", "ok");
    await call("open_path", r.path);
  } catch (err) { setStatus(String(err), "error"); toast(String(err), "error"); }
});

$("btnImportChanges").addEventListener("click", async () => {
  try {
    if (state.dirty) {
      if (!confirm("导入差异基于已保存数据库。当前有未保存更改，是否先保存再选择 Excel？")) return;
      if (!await saveCurrent(true)) return;
    }
    const path = await call("pick_import_xlsx");
    if (!path) return;
    setStatus("正在读取 Excel 并计算差异…");
    const result = await call("preview_import", path);
    renderImportPreview(result);
    setStatus(`差异预览完成：新增 ${result.action_counts["新增"]} / 修改 ${result.action_counts["修改"]} / 删除 ${result.action_counts["删除"]}`);
  } catch (err) { setStatus(String(err), "error"); toast(String(err), "error"); }
});

$("importConfirm").addEventListener("change", (e) => {
  $("importApply").disabled = !e.target.checked || !state.pendingImport || state.pendingImport.stale;
});
$("importApply").addEventListener("click", async () => {
  if (!state.pendingImport || !$("importConfirm").checked) return;
  try {
    $("importApply").disabled = true;
    setStatus("正在复核修订并应用 Excel 差异…");
    const r = await call("apply_import", state.pendingImport.token);
    applyBackendState(r);
    state.dirty = false;
    closeImportPreview();
    renderRules();
    renderAll();
    switchTab("摘要");
    setStatus(`Excel 变更已写入数据库（修订 ${r.revision}）`);
    toast("差异已通过二次确认并写入，原数据库已备份", "ok");
  } catch (err) {
    setStatus(String(err), "error");
    toast(String(err), "error");
    $("importApply").disabled = false;
  }
});
$("importCancel").addEventListener("click", closeImportPreview);
$("importClose").addEventListener("click", closeImportPreview);

async function refresh(store, xlsx) {
  const r = await call("load_state", store, xlsx);
  applyBackendState(r);
  state.dirty = false;
  renderAll();
  renderRules();
  switchTab("摘要");
  setStatus(r.migration ? `已从旧 JSON 迁移到本地数据库：${r.store_path}` : "就绪");
  if (r.migration) toast("JSON 已无损迁移；原文件保留不动", "ok");
}

window.addEventListener("pywebviewready", async () => {
  try {
    const r = await call("load_state", null, null);
    applyBackendState(r);
    state.dirty = false;
    renderAll();
    renderRules();
    switchTab("摘要");
    setStatus(r.migration ? `已从旧 JSON 迁移到本地数据库：${r.store_path}` : "就绪");
    if (r.migration) toast("JSON 已无损迁移；原文件保留不动", "ok");
  } catch (err) {
    setStatus("加载失败：" + err, "error");
  }
});

window.addEventListener("beforeunload", (e) => {
  if (!state.dirty) return;
  e.preventDefault();
  e.returnValue = "";
});

// ========== 新建订单 modal ==========
const PAYMENT_TEMPLATES = {
  A: [
    { kind: "预付款", ratio: 30, days: 0, trigger: "合同生效后", desc: "合同生效后付 30% 预付款" },
    { kind: "发货款", ratio: 30, days: 30, trigger: "货到签收后", desc: "货到签收后 30 天付 30% 发货款" },
    { kind: "验收款", ratio: 30, days: 60, trigger: "验收合格后", desc: "验收合格后 60 天付 30% 验收款" },
    { kind: "质保款", ratio: 10, days: 30, trigger: "质保期满后", desc: "质保期满后 30 天付 10% 尾款" },
  ],
  B: [
    { kind: "预付款", ratio: 50, days: 0, trigger: "合同生效后", desc: "合同生效后付 50% 预付款" },
    { kind: "发货款", ratio: 50, days: 30, trigger: "货到签收后", desc: "货到签收后 30 天付 50% 发货款" },
  ],
  C: [
    { kind: "预付款", ratio: 30, days: 0, trigger: "合同生效后", desc: "合同生效后付 30% 预付款" },
    { kind: "到货款", ratio: 40, days: 30, trigger: "货到签收后", desc: "货到签收后 30 天付 40% 到货款" },
    { kind: "验收款", ratio: 20, days: 60, trigger: "验收合格后", desc: "验收合格后 60 天付 20% 验收款" },
    { kind: "质保款", ratio: 10, days: 30, trigger: "质保期满后", desc: "质保期满后 30 天付 10% 尾款" },
  ],
};

function suggestJobNo(kind) {
  const yy = String(new Date().getFullYear()).slice(-2);
  const re = new RegExp("^" + yy + "(BS|DS)(\\d{3})$");
  let max = 0;
  for (const c of state.data["合同订单"]) {
    const m = String(c["JOB No"] || "").match(re);
    if (m && m[1] === kind) max = Math.max(max, parseInt(m[2], 10));
  }
  return yy + kind + String(max + 1).padStart(3, "0");
}

function buildNocDatalists() {
  const fill = (id, table, field) => {
    $(id).innerHTML = distinctValues(table, field).map((v) => `<option value="${esc(v)}">`).join("");
  };
  fill("dlCustomer", "合同订单", "客户");
  fill("dlAssist", "合同订单", "担当者");
  fill("dlModel", "设备台账", "设备型号");
  fill("dlShipMethod", "合同订单", "发货方式");
  fill("dlShipFrom", "合同订单", "发货地点");
  fill("dlShipTo", "合同订单", "送货地点");
  fill("dlPo", "设备台账", "PO No");
  fill("dlKind", "付款条件", "款类");
}

function nocDevRowHtml() {
  return `<tr>
    <td><input class="nd-model" list="dlModel" autocomplete="off" placeholder="型号"></td>
    <td><input class="nd-qty" type="number" min="1" step="1" value="1"></td>
    <td><input class="nd-price" type="number" step="any" placeholder="0"></td>
    <td><input class="nd-po" list="dlPo" autocomplete="off"></td>
    <td><select class="nd-free"><option>否</option><option>是</option></select></td>
    <td class="noc-del"><button type="button" class="nd-del">✕</button></td>
  </tr>`;
}

function nocUpdateSummary() {
  let qty = 0;
  const models = new Set();
  document.querySelectorAll("#nocDevRows tr").forEach((tr) => {
    const q = parseInt(tr.querySelector(".nd-qty").value, 10);
    const m = tr.querySelector(".nd-model").value.trim();
    if (q > 0) qty += q;
    if (m) models.add(m);
  });
  $("nocDevSummary").textContent = "合计 " + qty + " 台 / " + models.size + " 个型号";
}

function addNocDeviceRow() {
  const tr = document.createElement("tr");
  tr.innerHTML = nocDevRowHtml();
  $("nocDevRows").appendChild(tr);
}

function nocTermRowHtml() {
  return `<tr>
    <td><input class="nt-kind" list="dlKind" autocomplete="off" placeholder="款类"></td>
    <td><input class="nt-ratio" type="number" min="0" max="100" step="any" placeholder="0"></td>
    <td><input class="nt-days" type="number" min="0" step="1" placeholder="0"></td>
    <td class="noc-del"><button type="button" class="nt-del">✕</button></td>
  </tr>`;
}
function addNocTermRow(kind, ratio, days) {
  const tr = document.createElement("tr");
  tr.innerHTML = nocTermRowHtml();
  $("nocTermRows").appendChild(tr);
  if (kind !== undefined && kind !== "") {
    tr.querySelector(".nt-kind").value = kind;
    tr.querySelector(".nt-ratio").value = ratio;
    tr.querySelector(".nt-days").value = days;
  }
}
function fillNocTemplate(key) {
  const terms = PAYMENT_TEMPLATES[key];
  if (!terms) return;
  $("nocTermRows").innerHTML = "";
  for (const t of terms) addNocTermRow(t.kind, t.ratio, t.days);
  nocTermSummary();
}
function nocTermSummary() {
  let total = 0, n = 0;
  document.querySelectorAll("#nocTermRows tr").forEach((tr) => {
    const k = tr.querySelector(".nt-kind").value.trim();
    const r = Number(tr.querySelector(".nt-ratio").value) || 0;
    if (k) { total += r; n++; }
  });
  const sum = $("nocTermSummary");
  if (!n) { sum.textContent = ""; sum.style.color = ""; return; }
  const ok = Math.abs(total - 100) <= 0.01;
  sum.textContent = "合计 " + total + "% / " + n + " 条" + (ok ? "  ✓" : "  ⚠ 应为 100%");
  sum.style.color = ok ? "var(--ok)" : "var(--danger)";
}

function openNewOrder() {
  if (!state.data) return;
  buildNocDatalists();
  $("nocJobKind").value = "BS";
  $("nocJob").value = suggestJobNo("BS");
  $("nocCustomer").value = "";
  $("nocAssist").value = "朱方周";
  $("nocContent").value = "";
  $("nocCurrency").value = "RMB";
  $("nocShipMethod").value = "";
  $("nocShipFrom").value = "";
  $("nocShipTo").value = "";
  $("nocNote").value = "";
  $("nocBatch").value = "1";
  $("nocShipDate").value = "";
  $("nocTermRows").innerHTML = "";
  addNocTermRow();
  nocTermSummary();
  $("nocDevRows").innerHTML = "";
  addNocDeviceRow();
  nocUpdateSummary();
  $("nocErrors").classList.add("hidden");
  $("nocErrors").innerHTML = "";
  state._nocSubmitting = false;
  $("newOrderModal").classList.remove("hidden");
}

function closeNewOrder() { $("newOrderModal").classList.add("hidden"); }

function readNocForm() {
  const devices = [];
  document.querySelectorAll("#nocDevRows tr").forEach((tr) => {
    devices.push({
      model: tr.querySelector(".nd-model").value.trim(),
      qty: parseInt(tr.querySelector(".nd-qty").value, 10),
      price: tr.querySelector(".nd-price").value,
      po: tr.querySelector(".nd-po").value.trim(),
      free: tr.querySelector(".nd-free").value,
    });
  });
  return {
    job: $("nocJob").value.trim(),
    customer: $("nocCustomer").value.trim(),
    assist: $("nocAssist").value.trim(),
    content: $("nocContent").value.trim(),
    currency: $("nocCurrency").value,
    shipMethod: $("nocShipMethod").value.trim(),
    shipFrom: $("nocShipFrom").value.trim(),
    shipTo: $("nocShipTo").value.trim(),
    note: $("nocNote").value.trim(),
    batchName: $("nocBatch").value.trim() || "1",
    shipDate: $("nocShipDate").value,
    devices,
    terms: [...document.querySelectorAll("#nocTermRows tr")].map((tr) => ({
      kind: tr.querySelector(".nt-kind").value.trim(),
      ratio: Number(tr.querySelector(".nt-ratio").value) || 0,
      days: Number(tr.querySelector(".nt-days").value) || 0,
    })).filter((t) => t.kind),
  };
}

function validateNoc(f) {
  const errs = [];
  if (!/^\d{2}(BS|DS)\d{3}$/.test(f.job)) errs.push("JOB No 格式应为 YYBS### 或 YYDS###：" + (f.job || "（空）"));
  else if (state.data["合同订单"].some((c) => String(c["JOB No"]) === f.job)) errs.push("JOB No 已存在：" + f.job);
  if (!f.customer) errs.push("客户不能为空");
  const devs = f.devices.filter((d) => d.model && d.qty > 0);
  if (!devs.length) errs.push("至少添加 1 行设备（型号非空、台数>0）");
  f.devices.forEach((d, i) => {
    if (d.qty > 0 && !d.model) errs.push("第 " + (i + 1) + " 行设备缺少型号");
  });
  if (f.terms.length) {
    const sum = f.terms.reduce((a, t) => a + Number(t.ratio), 0);
    if (Math.abs(sum - 100) > 0.01) errs.push("付款条件比例合计 " + sum + "% ≠ 100%（或全部留空建空占位）");
  }
  return { ok: !errs.length, errors: errs };
}

function showNocErrors(errors) {
  const box = $("nocErrors");
  box.innerHTML = "<b>请修正以下 " + errors.length + " 处问题：</b><ul>" +
    errors.map((e) => "<li>" + esc(e) + "</li>").join("") + "</ul>";
  box.classList.remove("hidden");
}

function commitNewOrder(f) {
  const ts = Date.now();
  let rid = 0;
  const newId = () => "local-" + ts + "-" + (++rid);
  const job = f.job;
  const batch = f.batchName;

  state.data["合同订单"].push({
    "记录ID": newId(), "JOB No": job, "客户": f.customer, "担当者": f.assist,
    "订单内容": f.content, "发货方式": f.shipMethod, "发货地点": f.shipFrom,
    "送货地点": f.shipTo, "币种": f.currency || "RMB", "备注": f.note,
  });
  for (const d of f.devices) {
    if (!d.model || !(d.qty > 0)) continue;
    for (let i = 0; i < d.qty; i++) {
      state.data["设备台账"].push({
        "记录ID": newId(), "JOB No": job, "PO No": d.po, "设备型号": d.model,
        "未税单价": Number(d.price) || 0, "是否无偿": d.free, "发货批次": batch,
      });
    }
  }
  state.data["发货批次"].push({
    "记录ID": newId(), "JOB No": job, "发货批次": batch, "出荷日": f.shipDate,
  });
  if (f.terms.length) {
    for (const t of f.terms) {
      state.data["付款条件"].push({
        "记录ID": newId(), "JOB No": job, "款类": t.kind, "比例%": Number(t.ratio),
        "账期天数": Number(t.days),
      });
    }
  } else {
    state.data["付款条件"].push({ "记录ID": newId(), "JOB No": job }); // 空占位
  }
  // 开票/回款各建一条空占位：保证六表都有该 JOB 的记录（以 JOB No 索引），切表不空
  state.data["开票记录"].push({ "记录ID": newId(), "JOB No": job });
  state.data["回款记录"].push({ "记录ID": newId(), "JOB No": job });

  state.dirty = true;
  recomputeDerived();
  state.jobFilter = new Set([job]);
  switchTab("合同订单");
  renderCounts();
  closeNewOrder();
  const total = f.devices.reduce((a, d) => a + (d.qty > 0 ? d.qty : 0), 0);
  toast("已创建订单 " + job + "（" + total + " 台），已筛到该订单；点 ✕清除筛选 看全部", "ok");
}

$("btnNewOrder").addEventListener("click", openNewOrder);
$("nocClose").addEventListener("click", closeNewOrder);
$("nocCancel").addEventListener("click", closeNewOrder);
$("nocAddDev").addEventListener("click", () => { addNocDeviceRow(); nocUpdateSummary(); });
$("nocDevRows").addEventListener("input", nocUpdateSummary);
$("nocDevRows").addEventListener("click", (e) => {
  const del = e.target.closest(".nd-del");
  if (!del) return;
  if (document.querySelectorAll("#nocDevRows tr").length > 1) {
    del.closest("tr").remove();
    nocUpdateSummary();
  }
});
$("nocJobKind").addEventListener("change", (e) => { $("nocJob").value = suggestJobNo(e.target.value); });
$("nocAddTerm").addEventListener("click", () => { addNocTermRow(); nocTermSummary(); });
$("nocTermRows").addEventListener("input", nocTermSummary);
$("nocTermRows").addEventListener("click", (e) => {
  const del = e.target.closest(".nt-del");
  if (!del) return;
  del.closest("tr").remove();
  nocTermSummary();
});
document.querySelectorAll(".noc-tmpl").forEach((b) => {
  b.addEventListener("click", () => fillNocTemplate(b.dataset.tmpl));
});
$("nocOk").addEventListener("click", () => {
  if (state._nocSubmitting) return;
  const f = readNocForm();
  const v = validateNoc(f);
  if (!v.ok) { showNocErrors(v.errors); toast("请修正 " + v.errors.length + " 处问题", "error"); return; }
  state._nocSubmitting = true;
  commitNewOrder(f);
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !$("newOrderModal").classList.contains("hidden")) closeNewOrder();
});
