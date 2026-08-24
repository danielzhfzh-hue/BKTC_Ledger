"use strict";

const state = {
  storePath: "", xlsxPath: "", data: null, schema: null,
  tab: "摘要", dirty: false, selection: new Set(), anchor: null,
  filter: "", colFilters: {}, sortBy: {}, rules: {}, jobFilter: null,
  customerFilter: null,
  version: "?", update: null, databaseInfo: null, pendingImport: null,
};

const ROW_H = 31, BUFFER = 20;

const TAB_HINTS = {
  "合同订单": "本页 = 合同头信息。总台数、设备型号和付款条件文本自动汇总；修改 JOB No 会同步更新六表关联。",
  "付款条件": "点击“录入付款条件”后选择 JOB，一次维护完整付款条款；比例合计须 100%。客户自动带入，已被开票/回款引用的款类不可无关联地删除。",
  "设备台账": "点击“录入设备”使用独立表单；客户自动带入，发货批次只能选择该 JOB 已存在的批次。宽表保留用于查看和修改既有记录。",
  "发货批次": "点击“新增发货批次”填写批次并可勾选同一 JOB 的设备；台数、合计和覆盖番号自动计算。批次改名会同步关联表。",
  "开票记录": "点击“新增开票记录”；款类来自该 JOB 付款条件，覆盖范围只显示该 JOB 的批次/制造番号，覆盖台数和应收状态自动计算。",
  "回款记录": "点击“新增回款记录”；款类来自该 JOB 付款条件，可整批或逐台勾选，批次自动回写，覆盖台数和超期信息自动计算。",
};

const ENTRY_ACTION_LABELS = {
  "合同订单": "＋ 新建订单",
  "付款条件": "＋ 录入付款条件",
  "设备台账": "＋ 录入设备",
  "发货批次": "＋ 新增发货批次",
  "开票记录": "＋ 新增开票记录",
  "回款记录": "＋ 新增回款记录",
};

const $ = (id) => document.getElementById(id);
function s(v) { return v === null || v === undefined ? "" : String(v); }
function coverageValues(value) {
  return s(value).split(/[;；\r\n]+/).map((x) => x.trim()).filter(Boolean);
}
function replaceCoverageValue(value, previous, next) {
  if (!previous || previous === next) return value;
  let changed = false;
  const items = coverageValues(value).map((item) => {
    const pieces = item.split(/([→⇒⟶➡➝])/);
    for (let i = 0; i < pieces.length; i += 2) {
      if (pieces[i] === previous) { pieces[i] = next; changed = true; }
    }
    return pieces.join("");
  });
  return changed ? items.join(";") : value;
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
  if (document.activeElement !== $("setStoreInput")) {
    $("setStoreInput").value = state.storePath || "";
  }
  if (document.activeElement !== $("setXlsxInput")) {
    $("setXlsxInput").value = state.xlsxPath || "";
  }
  $("setVersion").textContent = "v" + state.version;
  const db = state.databaseInfo || {};
  $("setRevision").textContent = db.revision ?? "—";
  $("setIntegrity").textContent = db.integrity === "ok" ? "正常" : (db.integrity || "—");
  $("setLastSaved").textContent = db.last_saved_at ? db.last_saved_at.replace("T", " ") : "—";
  renderSummary();
  renderGrid();
  renderCounts();
}

function currentContext() {
  // 从当前表的单值列筛选推导"新建行的上下文"（替代旧的客户/JOB 下拉）
  const cf = state.colFilters[state.tab] || {};
  const ctx = { job: "", customer: "" };
  if (state.jobFilter && state.jobFilter.size === 1) ctx.job = [...state.jobFilter][0];
  if (state.customerFilter && state.customerFilter.size === 1) {
    ctx.customer = [...state.customerFilter][0];
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
  const termKinds = {};
  for (const term of state.data["付款条件"]) {
    const job = s(term["JOB No"]), kind = s(term["款类"]);
    if (kind) (termKinds[job] ||= new Set()).add(kind);
  }
  for (const table of ["开票记录", "回款记录"]) {
    for (const rec of state.data[table]) {
      const job = s(rec["JOB No"]), kind = s(rec["款类"]);
      if (!kind || (table === "开票记录" && kind === "全额")) continue;
      if (!termKinds[job]?.has(kind)) {
        issues.push({ severity: "error", msg: `${table} ${job} 款类“${kind}”不在该 JOB 的付款条件中` });
      }
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
  const jobsWithDevices = new Set(state.data["设备台账"].map((d) => s(d["JOB No"])));
  for (const table of ["开票记录", "回款记录"]) {
    for (const rec of state.data[table]) {
      const job = s(rec["JOB No"]);
      if (table === "回款记录") {
        const meaningful = s(rec["款类"]) || s(rec["回款日"])
          || ![null, undefined, "", 0].includes(rec["含税金额"]);
        if (meaningful && jobsWithDevices.has(job)
            && !s(rec["覆盖批次"]) && !s(rec["覆盖製造番号"])) {
          issues.push({ severity: "error", msg: `${job} 回款必须选择覆盖批次或设备，否则该笔金额不会参与未回收计算` });
        }
      }
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
  const addButton = $("btnAddRecord");
  if (addButton && ENTRY_ACTION_LABELS[tab]) addButton.textContent = ENTRY_ACTION_LABELS[tab];
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

function visibleRows(table, excludedField = "") {
  const q = state.filter.trim().toLowerCase();
  const cf = state.colFilters[table] || {};
  const so = state.sortBy[table];
  const recs = state.data[table];
  const idxs = [];
  for (let i = 0; i < recs.length; i++) {
    const rec = recs[i];
    let keep = true;
    for (const f in cf) {
      if (f === "JOB No" || f === "客户") continue;
      if (f === excludedField) continue;
      const set = cf[f];
      if (set && !set.has(String(rec[f] ?? ""))) { keep = false; break; }
    }
    if (keep && excludedField !== "JOB No" && state.jobFilter
        && !state.jobFilter.has(String(rec["JOB No"] ?? ""))) keep = false;
    if (keep && excludedField !== "客户" && state.customerFilter
        && !state.customerFilter.has(String(rec["客户"] ?? ""))) keep = false;
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

function coverageContains(text, serial) {
  return coverageValues(text).some((item) =>
    item === serial || item.split(/[→⇒⟶➡➝]/).includes(serial));
}

function coverageMatches(record, device) {
  const serialText = s(record["覆盖製造番号"]);
  const serial = s(device["製造番号"]);
  const batch = s(device["发货批次"]);
  const batches = new Set(coverageValues(record["覆盖批次"]));
  if (serialText && serial) return coverageContains(serialText, serial) && (!batches.size || batches.has(batch));
  return batches.size > 0 && batches.has(batch);
}

function matchingCoverageRecords(records, device) {
  const explicit = records.filter((r) => s(r["覆盖製造番号"]) && coverageMatches(r, device));
  return explicit.length ? explicit
    : records.filter((r) => !s(r["覆盖製造番号"]) && coverageMatches(r, device));
}

function isoDateValue(value) {
  const m = s(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const ms = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const dt = new Date(ms);
  if (dt.getUTCFullYear() !== Number(m[1]) || dt.getUTCMonth() !== Number(m[2]) - 1 || dt.getUTCDate() !== Number(m[3])) return null;
  return ms;
}

function isoFromMs(ms) {
  const dt = new Date(ms);
  return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth() + 1).padStart(2, "0")}-${String(dt.getUTCDate()).padStart(2, "0")}`;
}

function recomputeDerived() {
  // 镜像 core.derive()：前端即时显示，保存时后端再权威计算。
  const d = state.data;
  if (!d) return;
  const customerByJob = {};
  for (const c of d["合同订单"]) customerByJob[s(c["JOB No"])] = s(c["客户"]);
  for (const table of ["付款条件", "设备台账", "发货批次", "开票记录", "回款记录"]) {
    for (const rec of d[table]) rec["客户"] = customerByJob[s(rec["JOB No"])] || "";
  }

  const devBy = {}, devicesByJob = {};
  for (const dev of d["设备台账"]) {
    const job = s(dev["JOB No"]);
    (devBy[job + "\u0001" + s(dev["发货批次"])] ||= []).push(dev);
    (devicesByJob[job] ||= []).push(dev);
  }
  for (const x of d["发货批次"]) {
    const devs = devBy[s(x["JOB No"]) + "\u0001" + s(x["发货批次"])] || [];
    x["台数"] = devs.length;
    x["未税合计"] = Math.round(devs.reduce((a, v) => a + (Number(v["未税单价"]) || 0), 0) * 1000) / 1000;
    x["含税合计"] = Math.round(x["未税合计"] * 1.13 * 100) / 100;
    x["覆盖製造番号"] = devs.map((v) => s(v["製造番号"])).filter(Boolean).join(";");
  }

  const cnt = {}, models = {}, termsByJob = {};
  for (const t of d["付款条件"]) (termsByJob[s(t["JOB No"])] ||= []).push(t);
  for (const dev of d["设备台账"]) {
    const job = s(dev["JOB No"]);
    cnt[job] = (cnt[job] || 0) + 1;
    (models[job] ||= new Set()).add(s(dev["设备型号"]));
    dev["验收状态"] = s(dev["质保开始日"]) ? "已验收" : "未验收";
  }
  for (const c of d["合同订单"]) {
    const job = s(c["JOB No"]);
    c["总台数"] = cnt[job] || 0;
    if (!s(c["设备型号"])) c["设备型号"] = [...(models[job] || new Set())].filter(Boolean).sort().join(";");
    const desc = (termsByJob[job] || []).map((t) => s(t["说明"]).trim()).filter(Boolean).join("；");
    if (desc) c["付款条件"] = desc;
  }

  for (const table of ["开票记录", "回款记录"]) {
    for (const rec of d[table]) {
      const serials = new Set(coverageValues(rec["覆盖製造番号"]));
      const jobDevices = devicesByJob[s(rec["JOB No"])] || [];
      const hasCoverage = serials.size || coverageValues(rec["覆盖批次"]).length;
      rec["覆盖台数"] = hasCoverage
        ? jobDevices.filter((dev) => coverageMatches(rec, dev)).length
        : 0;
      if (serials.size) {
        const batches = [];
        let safeToRebuild = true;
        for (const token of serials) {
          const parts = new Set(token.split(/[→⇒⟶➡➝]/).filter(Boolean));
          const matches = jobDevices.filter((dev) =>
            s(dev["製造番号"]) === token || parts.has(s(dev["製造番号"])));
          if (matches.length !== 1) { safeToRebuild = false; break; }
          const batch = s(matches[0]["发货批次"]);
          if (batch && !batches.includes(batch)) batches.push(batch);
        }
        const blankSerialBatches = new Set(jobDevices
          .filter((dev) => !s(dev["製造番号"]))
          .map((dev) => s(dev["发货批次"])));
        const removesBlankSerialBatch = coverageValues(rec["覆盖批次"])
          .some((batch) => !batches.includes(batch) && blankSerialBatches.has(batch));
        if (removesBlankSerialBatch) safeToRebuild = false;
        if (safeToRebuild && batches.length) rec["覆盖批次"] = batches.join(";");
      }
      if (table === "开票记录") {
        const opened = isoDateValue(rec["开票日"]);
        rec["应收回款日"] = opened === null ? ""
          : isoFromMs(opened + (parseInt(rec["账期天数"], 10) || 0) * 86400000);
      }
    }
  }

  const invoicesByJob = {}, paymentsByJob = {};
  for (const inv of d["开票记录"]) (invoicesByJob[s(inv["JOB No"])] ||= []).push(inv);
  for (const pay of d["回款记录"]) (paymentsByJob[s(pay["JOB No"])] ||= []).push(pay);
  const coveredDevices = (rec, jobDevices) => jobDevices.filter((dev) =>
    s(dev["是否无偿"]) !== "是" && coverageMatches(rec, dev));
  const allocatedAmount = (rec, selectedDevices, jobDevices) => {
    const covered = coveredDevices(rec, jobDevices);
    const coveredSet = new Set(covered);
    const total = covered.reduce((sum, dev) => sum + (Number(dev["未税单价"]) || 0), 0);
    const selected = selectedDevices.reduce((sum, dev) =>
      sum + (coveredSet.has(dev) ? (Number(dev["未税单价"]) || 0) : 0), 0);
    return total ? (Number(rec["含税金额"]) || 0) * selected / total : 0;
  };
  const today = Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate());
  for (const inv of d["开票记录"]) {
    const meaningful = s(inv["款类"]) || s(inv["开票日"]) || Number(inv["含税金额"]) || s(inv["覆盖批次"]) || s(inv["覆盖製造番号"]);
    if (!meaningful) { inv["回款状态"] = ""; continue; }
    const job = s(inv["JOB No"]), kind = s(inv["款类"]), jobDevices = devicesByJob[job] || [];
    const candidates = (paymentsByJob[job] || []).filter((p) => kind === "全额" || s(p["款类"]) === kind);
    const relatedInvoices = (invoicesByJob[job] || []).filter((other) => s(other["款类"]) === kind);
    const invDevices = coveredDevices(inv, jobDevices);
    const paid = jobDevices.length
      ? candidates.reduce((sum, p) => sum + allocatedAmount(p, invDevices, jobDevices), 0)
      : candidates.reduce((sum, p) => sum + (Number(p["含税金额"]) || 0), 0);
    const invoiced = jobDevices.length
      ? relatedInvoices.reduce((sum, other) => sum + allocatedAmount(other, invDevices, jobDevices), 0)
      : relatedInvoices.reduce((sum, other) => sum + (Number(other["含税金额"]) || 0), 0);
    const amount = invoiced || Number(inv["含税金额"]) || 0;
    const due = isoDateValue(inv["应收回款日"]);
    inv["回款状态"] = amount > 0 && paid >= amount - 0.01 ? "已回款"
      : paid > 0.01 ? "部分回款" : due !== null && due < today ? "超期未回" : "未回款";
  }
  for (const pay of d["回款记录"]) {
    const job = s(pay["JOB No"]), kind = s(pay["款类"]), jobDevices = devicesByJob[job] || [];
    const candidates = (invoicesByJob[job] || []).filter((inv) => [kind, "全额"].includes(s(inv["款类"])));
    let relevant = candidates;
    if (jobDevices.length) {
      relevant = [];
      for (const dev of coveredDevices(pay, jobDevices)) {
        for (const inv of matchingCoverageRecords(candidates, dev)) if (!relevant.includes(inv)) relevant.push(inv);
      }
    }
    const dues = relevant.map((inv) => isoDateValue(inv["应收回款日"])).filter((v) => v !== null);
    const due = dues.length ? Math.min(...dues) : null;
    pay["对应应收回款日"] = due === null ? "" : isoFromMs(due);
    const paidOn = isoDateValue(pay["回款日"]);
    if (due !== null && paidOn !== null) {
      const days = Math.max(0, Math.round((paidOn - due) / 86400000));
      pay["是否超期"] = days ? "是" : "否";
      pay["超期天数"] = days;
    } else {
      pay["是否超期"] = "";
      pay["超期天数"] = null;
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
      const activeFilter = f.name === "JOB No" ? state.jobFilter
        : (f.name === "客户" ? state.customerFilter : cfT[f.name]);
      const act = activeFilter ? " filt-on" : "";
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
      if (["开票记录", "回款记录"].includes(table)
          && ["覆盖批次", "覆盖製造番号"].includes(f.name)) {
        const values = coverageValues(v);
        const label = f.name === "覆盖批次"
          ? (values.length ? values.join("、") : "选择批次")
          : (values.length ? `${values.length} 台设备` : "勾选设备");
        const pickerTitle = f.name === "覆盖批次" ? "按批次选择并带入整批设备" : "按单台设备逐一勾选";
        return `<td class="coverage-cell"><button type="button" class="coverage-picker" data-coverage-field="${esc(f.name)}" title="${pickerTitle}">${esc(label)} <span>▾</span></button></td>`;
      }
      if (f.type === "select") {
        let options = f.options;
        if (["开票记录", "回款记录"].includes(table) && f.name === "款类") {
          options = paymentKindsForJob(s(rec["JOB No"]));
          if (s(v) && !options.includes(s(v))) options = [...options, s(v)];
        }
        const opts = ["", ...options].map((o) =>
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
  if (state.customerFilter) h += `　［跨表客户筛选：${[...state.customerFilter].map(esc).join("、")}］`;
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

function updateRec(table, ri, field, raw, deferRecompute = false) {
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
        row["覆盖批次"] = replaceCoverageValue(row["覆盖批次"], previous, next);
      }
    }
  }
  if (table === "设备台账" && field === "製造番号" && previous !== s(v)) {
    const job = s(rec["JOB No"]), next = s(v);
    for (const child of ["开票记录", "回款记录"]) {
      for (const row of state.data[child]) {
        if (s(row["JOB No"]) === job) {
          row["覆盖製造番号"] = replaceCoverageValue(
            row["覆盖製造番号"], previous, next
          );
        }
      }
    }
  }
  if (table === "付款条件" && field === "款类" && previous !== s(v)) {
    const job = s(rec["JOB No"]), next = s(v);
    const oldKindStillExists = state.data["付款条件"].some((row) =>
      row !== rec && s(row["JOB No"]) === job && s(row["款类"]) === previous);
    if (!oldKindStillExists) {
      for (const child of ["开票记录", "回款记录"]) {
        for (const row of state.data[child]) {
          if (s(row["JOB No"]) === job && s(row["款类"]) === previous) row["款类"] = next;
        }
      }
    }
  }
  state.dirty = true;
  if (!deferRecompute) {
    recomputeDerived();
    renderCounts();
  }
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
      || e.target.closest("td.date-cell.empty") || e.target.closest(".coverage-picker")) return;
  const tr = e.target.closest("tr");
  if (!tr) return;
  e.preventDefault();
  selectRow(Number(tr.dataset.ri), e.shiftKey || e.metaKey || e.ctrlKey);
});

// 空日期单元格：点击才挂日期选择器（macOS 原生 picker 会把空值显示成"今天"，造成"假数据"错觉）
$("gridBody").addEventListener("click", (e) => {
  const picker = e.target.closest(".coverage-picker");
  if (picker) {
    const tr = picker.closest("tr[data-ri]");
    if (tr) openCoverageModal(state.tab, Number(tr.dataset.ri), picker.dataset.coverageField);
    return;
  }
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
      if (table === "合同订单") openNewOrder();
      else openEntryModal(table);
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
  state.customerFilter = null;
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

// 开票/回款覆盖范围：批次入口整批选择，製造番号入口逐台选择。
// draftRecord 用于独立录入窗口，应用覆盖范围前不写入正式表。
function openCoverageModal(table, ri, field, draftRecord = null) {
  if (!["开票记录", "回款记录"].includes(table)) return;
  const rec = draftRecord || state.data[table]?.[ri];
  if (!rec) return;
  const mode = field === "覆盖批次" ? "batch" : "serial";
  const job = s(rec["JOB No"]);
  const devices = state.data["设备台账"]
    .map((device, index) => ({ device, index }))
    .filter(({ device }) => s(device["JOB No"]) === job);
  const selectedSerials = new Set(coverageValues(rec["覆盖製造番号"]));
  const selectedBatches = new Set(coverageValues(rec["覆盖批次"]));
  const groups = new Map();
  for (const item of devices) {
    const batch = s(item.device["发货批次"]) || "（无批次）";
    if (!groups.has(batch)) groups.set(batch, []);
    groups.get(batch).push(item);
  }
  for (const shipment of state.data["发货批次"]) {
    if (s(shipment["JOB No"]) !== job) continue;
    const batch = s(shipment["发货批次"]);
    if (batch && !groups.has(batch)) groups.set(batch, []);
  }
  $("coverageTitle").textContent = mode === "batch"
    ? `${table === "回款记录" ? "回款" : "开票"}覆盖批次`
    : `${table === "回款记录" ? "回款" : "开票"}覆盖製造番号`;
  $("coverageContext").textContent = `${job || "（未填 JOB）"} · ${s(rec["客户"]) || "（无客户）"} · ${s(rec["款类"]) || "（未选款类）"} · ${mode === "batch" ? "勾选批次后带入该批全部设备" : "每个勾选框仅代表一台设备"}`;
  $("coverageSearch").value = "";
  $("coverageSelectVisible").classList.toggle("hidden", mode === "serial");
  $("coverageSelectVisible").textContent = "勾选当前批次";
  $("coverageManualBatchLabel").classList.toggle("hidden", mode !== "batch");
  $("coverageManualSerialLabel").classList.toggle("hidden", mode !== "serial");
  $("coverageSafety").textContent = mode === "batch"
    ? "每个批次会带入该批全部设备；保存前还会再做关联校验。"
    : "请按实际开票/回款范围逐台勾选；覆盖批次会自动生成。";
  $("coverageManualBatch").value = s(rec["覆盖批次"]);
  $("coverageManualSerial").value = s(rec["覆盖製造番号"]);
  $("coverageList").innerHTML = groups.size ? [...groups.entries()].map(([batch, items]) => {
    const selectable = items.filter(({ device }) => s(device["製造番号"]));
    const rows = items.map(({ device, index }) => {
      const serial = s(device["製造番号"]);
      const checked = serial && (coverageContains(rec["覆盖製造番号"], serial)
        || (!selectedSerials.size && selectedBatches.has(s(device["发货批次"]))));
      const search = [batch, serial, device["機番"], device["设备型号"], device["PO No"]].map(s).join(" ").toLowerCase();
      const rowTag = mode === "serial" ? "label" : "div";
      const selector = mode === "serial"
        ? `<input type="checkbox" class="coverage-device-check" data-index="${index}" ${checked ? "checked" : ""} ${serial ? "" : "disabled"}>`
        : '<span class="coverage-device-marker">•</span>';
      return `<${rowTag} class="coverage-device${mode === "batch" ? " coverage-device-readonly" : ""}" data-search="${esc(search)}">` +
        selector +
        `<span class="serial">${esc(serial || "（无番号，不可勾选）")}</span>` +
        `<span>${esc(s(device["设备型号"]) || "—")}</span>` +
        `<span class="meta">机番 ${esc(s(device["機番"]) || "—")} · ${Number(device["未税单价"] || 0).toLocaleString()}</span></${rowTag}>`;
    }).join("");
    const allChecked = selectedBatches.has(batch) || (selectable.length && selectable.every(({ device }) =>
      coverageContains(rec["覆盖製造番号"], s(device["製造番号"]))
      || (!selectedSerials.size && selectedBatches.has(s(device["发货批次"])))));
    const groupHead = mode === "batch"
      ? `<label class="coverage-group-head"><input type="checkbox" class="coverage-group-all" data-batch-key="${esc(batch)}" ${allChecked ? "checked" : ""} ${batch === "（无批次）" ? "disabled" : ""}><b>批次 ${esc(batch)}</b><span class="coverage-group-meta">${items.length} 台</span></label>`
      : `<div class="coverage-group-head"><b>批次 ${esc(batch)}</b><span class="coverage-group-meta">${items.length} 台 · 请逐台勾选</span></div>`;
    return `<div class="coverage-group">${groupHead}${rows}</div>`;
  }).join("") : '<div class="coverage-empty">该 JOB 尚无设备。请展开下方“手工输入”填批次，或先到设备台账建设备。</div>';
  state._coverage = {
    table, ri, job, mode, record: rec, draft: !!draftRecord,
    deviceIndexes: devices.map((x) => x.index),
    batchIndexes: new Map([...groups.entries()].map(([batch, items]) =>
      [batch, items.map(({ index }) => index)])),
  };
  updateCoverageSummary();
  $("coverageModal").classList.remove("hidden");
}

function coverageSelectedIndexes() {
  if (state._coverage?.mode === "batch") {
    return [...document.querySelectorAll("#coverageList .coverage-group-all:checked")]
      .flatMap((box) => state._coverage.batchIndexes.get(box.dataset.batchKey) || []);
  }
  return [...document.querySelectorAll("#coverageList .coverage-device-check:checked")]
    .map((box) => Number(box.dataset.index));
}

function coverageSelectedBatches() {
  if (state._coverage?.mode !== "batch") return [];
  return [...document.querySelectorAll("#coverageList .coverage-group-all:checked")]
    .map((box) => box.dataset.batchKey)
    .filter((batch) => batch && batch !== "（无批次）");
}

function updateCoverageGroupChecks() {
  if (state._coverage?.mode !== "serial") return;
  document.querySelectorAll("#coverageList .coverage-group").forEach((group) => {
    const boxes = [...group.querySelectorAll(".coverage-device-check:not(:disabled)")];
    const all = group.querySelector(".coverage-group-all");
    if (all) {
      all.checked = boxes.length > 0 && boxes.every((box) => box.checked);
      all.indeterminate = boxes.some((box) => box.checked) && !all.checked;
    }
  });
}

function updateCoverageSummary() {
  const indexes = coverageSelectedIndexes();
  const batches = new Set(indexes.map((index) => s(state.data["设备台账"][index]?.["发货批次"])));
  const selectedBatchCount = coverageSelectedBatches().length;
  $("coverageSummary").textContent = indexes.length || selectedBatchCount
    ? (state._coverage?.mode === "batch"
      ? `已选 ${selectedBatchCount} 个批次 · 将覆盖 ${indexes.length} 台设备`
      : `已逐台选择 ${indexes.length} 台 · 覆盖 ${[...batches].filter(Boolean).length} 个批次`)
    : (state._coverage?.mode === "batch"
      ? "尚未选择批次；无设备订单可使用下方手工输入。"
      : "尚未逐台勾选设备；历史番号可使用下方手工输入。");
}

function closeCoverageModal() {
  $("coverageModal").classList.add("hidden");
  state._coverage = null;
}

$("coverageList").addEventListener("change", (e) => {
  const groupAll = e.target.closest(".coverage-group-all");
  if (groupAll) {
    groupAll.closest(".coverage-group").querySelectorAll(".coverage-device-check:not(:disabled)")
      .forEach((box) => { box.checked = groupAll.checked; });
  }
  updateCoverageGroupChecks();
  updateCoverageSummary();
});
$("coverageSearch").addEventListener("input", (e) => {
  const query = e.target.value.trim().toLowerCase();
  document.querySelectorAll("#coverageList .coverage-device").forEach((row) => {
    row.classList.toggle("hidden-by-search", !!query && !row.dataset.search.includes(query));
  });
  document.querySelectorAll("#coverageList .coverage-group").forEach((group) => {
    group.style.display = [...group.querySelectorAll(".coverage-device")]
      .some((row) => !row.classList.contains("hidden-by-search")) ? "" : "none";
  });
});
$("coverageSelectVisible").addEventListener("click", () => {
  document.querySelectorAll("#coverageList .coverage-group:not([style*='display: none']) .coverage-group-all:not(:disabled)")
    .forEach((box) => { box.checked = true; });
  updateCoverageGroupChecks();
  updateCoverageSummary();
});
$("coverageClear").addEventListener("click", () => {
  document.querySelectorAll("#coverageList .coverage-device-check, #coverageList .coverage-group-all")
    .forEach((box) => { box.checked = false; });
  $("coverageManualBatch").value = "";
  $("coverageManualSerial").value = "";
  updateCoverageGroupChecks();
  updateCoverageSummary();
});
$("coverageApply").addEventListener("click", () => {
  const context = state._coverage;
  if (!context) return;
  const rec = context.record || state.data[context.table]?.[context.ri];
  if (!rec) return closeCoverageModal();
  const indexes = coverageSelectedIndexes();
  let serials, batches;
  const selectedBatches = coverageSelectedBatches();
  if (context.mode === "batch" && selectedBatches.length) {
    serials = [...new Set(indexes.map((index) => s(state.data["设备台账"][index]["製造番号"])).filter(Boolean))];
    batches = selectedBatches;
  } else if (indexes.length) {
    serials = [...new Set(indexes.map((index) => s(state.data["设备台账"][index]["製造番号"])).filter(Boolean))];
    batches = [...new Set(indexes.map((index) => s(state.data["设备台账"][index]["发货批次"])).filter(Boolean))];
  } else {
    serials = coverageValues($("coverageManualSerial").value);
    batches = coverageValues($("coverageManualBatch").value);
  }
  const meaningfulPayment = context.table === "回款记录"
    && (s(rec["款类"]) || s(rec["回款日"]) || Number(rec["含税金额"]));
  if (meaningfulPayment && context.deviceIndexes.length && !serials.length && !batches.length) {
    toast("该 JOB 有设备，回款必须勾选设备或填覆盖批次", "error");
    return;
  }
  rec["覆盖製造番号"] = serials.join(";");
  rec["覆盖批次"] = batches.join(";");
  if (context.draft) {
    closeCoverageModal();
    updateEntryCoveragePreview();
    toast(`已带入 ${batches.length} 个批次 / ${indexes.length} 台设备`, "ok");
    return;
  }
  state.dirty = true;
  recomputeDerived();
  closeCoverageModal();
  renderGrid();
  renderCounts();
  renderSummary();
  toast(`已应用 ${serials.length ? serials.length + " 台设备" : batches.length + " 个批次"}`, "ok");
});
$("coverageCancel").addEventListener("click", closeCoverageModal);
$("coverageClose").addEventListener("click", closeCoverageModal);
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !$("coverageModal").classList.contains("hidden")) {
    e.stopImmediatePropagation();
    closeCoverageModal();
  }
});

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
const VIRTUAL_QUERY_TABLES = ["未付款订单", "未回款明细"];
const VIRTUAL_QUERY_SCHEMAS = {
  "未付款订单": [
    { name: "客户", type: "text" }, { name: "JOB No", type: "text" },
    { name: "未回款项数", type: "int" }, { name: "未回收合计", type: "number" },
    { name: "预警等级", type: "text" }, { name: "未回款原因", type: "text" },
  ],
  "未回款明细": [
    { name: "客户", type: "text" }, { name: "JOB No", type: "text" },
    { name: "批次", type: "text" }, { name: "款类", type: "text" },
    { name: "预警等级", type: "text" }, { name: "开票日期", type: "date" },
    { name: "预定回收日期", type: "date" },
    { name: "未回收金额", type: "number" }, { name: "未回收原因", type: "text" },
  ],
};
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
let _unpaidRows = [];

function querySchema(table) {
  if (VIRTUAL_QUERY_SCHEMAS[table]) return VIRTUAL_QUERY_SCHEMAS[table];
  return (state.schema[table] || []).map((f) => ({ name: f[0], type: f[1] }));
}

function queryRows(table) {
  if (table === "未回款明细") return _unpaidRows;
  if (table === "未付款订单") {
    const groups = new Map();
    const severity = { "①已逾期": 0, "②临近": 1, "③未到期": 2, "④待确认": 3 };
    for (const row of _unpaidRows) {
      const job = s(row["JOB No"]);
      const key = `${job}\u0001${s(row["客户"])}`;
      const group = groups.get(key) || {
        "客户": row["客户"], "JOB No": job, "未回款项数": 0,
        "未回收合计": 0, "预警等级": new Set(), "未回款原因": new Set(),
      };
      group["未回款项数"] += 1;
      group["未回收合计"] += Number(row["未回收金额"]) || 0;
      if (s(row["预警等级"])) group["预警等级"].add(s(row["预警等级"]));
      if (s(row["未回收原因"])) group["未回款原因"].add(s(row["未回收原因"]));
      groups.set(key, group);
    }
    return [...groups.values()].map((group) => ({
      ...group,
      "未回收合计": Math.round(group["未回收合计"] * 100) / 100,
      "预警等级": [...group["预警等级"]].sort((a, b) => (severity[a] ?? 9) - (severity[b] ?? 9)).join(" / "),
      "未回款原因": [...group["未回款原因"]].join("；"),
    }));
  }
  return state.data[table] || [];
}

async function refreshUnpaidRows() {
  const rows = await call("get_unpaid_rows", state.data, state.rules);
  _unpaidRows = Array.isArray(rows) ? rows : (rows?.rows || []);
}
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
function qbTables(base) {
  return VIRTUAL_QUERY_SCHEMAS[base] ? [base] : [base, ...(PARENTS[base] || [])];
}
function qbFieldList(base) {
  const out = [];
  const baseNames = new Set(querySchema(base).map((field) => field.name));
  for (const t of qbTables(base)) {
    const isBase = (t === base);
    for (const f of querySchema(t)) {
      const name = f.name, type = f.type;
      if (!isBase && baseNames.has(name)) continue;
      out.push({ table: t, name, type, label: isBase ? name : `${SHORT[t] || t}.${name}` });
    }
  }
  return out;
}
function openQueryBuilder() {
  if (!state.data) return;
  buildLookups();
  const tables = [...VIRTUAL_QUERY_TABLES, ...Object.keys(state.schema)];
  $("qbBase").innerHTML = tables.map((t) => `<option value="${esc(t)}">${esc(t)}</option>`).join("");
  $("qbBase").value = "未付款订单";
  renderQueryBuilder();
  $("qbCount").textContent = "正在计算未回款口径…";
  refreshUnpaidRows().then(() => renderQueryBuilder()).catch((err) => {
    $("qbCount").textContent = "未回款数据加载失败";
    toast(String(err), "error");
  });
}
function renderQueryBuilder() {
  const base = $("qbBase").value;
  const fields = qbFieldList(base);
  const defaultChecked = VIRTUAL_QUERY_SCHEMAS[base] ? " checked" : "";
  $("qbFields").innerHTML = qbTables(base).map((t) => {
    const isBase = (t === base);
    const items = fields.filter((f) => f.table === t).map((f) =>
      `<label class="cf-item"><input type="checkbox"${defaultChecked} data-t="${esc(f.table)}" data-n="${esc(f.name)}" data-type="${f.type}" data-label="${esc(f.label)}"><span class="cf-v">${esc(f.label)}</span></label>`
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
      <option value="eq">等于</option><option value="contains">包含</option><option value="ne">不等于</option>
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
  for (const r of queryRows(table)) {
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
  if (flt.op === "contains") return v.toLocaleLowerCase().includes(target.toLocaleLowerCase());
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
  const recs = queryRows(base);
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
async function qbDoPreview() {
  if (VIRTUAL_QUERY_SCHEMAS[$("qbBase").value]) {
    try { await refreshUnpaidRows(); }
    catch (err) { toast(String(err), "error"); return; }
  }
  const { fields, rows } = runQuery();
  if (!fields.length) { toast("请至少勾选一个输出字段", "error"); return; }
  $("qbCount").textContent = `符合 ${rows.length} 行 × ${fields.length} 列`;
  const head = `<tr>${fields.map((f) => `<th>${esc(f.label)}</th>`).join("")}</tr>`;
  const body = rows.slice(0, 10).map((r) =>
    `<tr>${fields.map((f) => `<td>${esc(r[f.label])}</td>`).join("")}</tr>`).join("");
  $("qbPreview").innerHTML = `<thead>${head}</thead><tbody>${body}</tbody>`;
}
async function qbDoExport() {
  if (VIRTUAL_QUERY_SCHEMAS[$("qbBase").value]) {
    try { await refreshUnpaidRows(); }
    catch (err) { toast(String(err), "error"); return; }
  }
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
$("qbBase").addEventListener("change", () => {
  renderQueryBuilder();
  if (VIRTUAL_QUERY_SCHEMAS[$("qbBase").value]) {
    refreshUnpaidRows().then(() => renderQueryBuilder()).catch((err) => toast(String(err), "error"));
  }
});
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
      updateRec(table, r, field.name, row[c], true);
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
  state.customerFilter = null;
  delete state.sortBy[state.tab];
  state.filter = "";
  $("filterKeyword").value = "";
  closeColFilter();
  renderGrid();
  toast("已清除筛选/排序（含客户与 JOB 跨表筛选）");
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
  const isJob = field === "JOB No";
  const isCustomer = field === "客户";
  let recs = visibleRows(table, field).map((index) => state.data[table][index]);
  if (isJob || isCustomer) {
    recs = state.data["合同订单"].filter((rec) => {
      if (isJob && state.customerFilter && !state.customerFilter.has(s(rec["客户"]))) return false;
      if (isCustomer && state.jobFilter && !state.jobFilter.has(s(rec["JOB No"]))) return false;
      return true;
    });
  }
  const dist = new Map();
  for (const r of recs) {
    const k = String(r[field] ?? "");
    dist.set(k, (dist.get(k) || 0) + 1);
  }
  _cfAllVals = [...dist.keys()].sort((a, b) => a.localeCompare(b, "zh"));
  const allCount = _cfAllVals.length;
  state.colFilters[table] = state.colFilters[table] || {};
  const activeSet = isJob ? state.jobFilter
    : isCustomer ? state.customerFilter : state.colFilters[table][field];
  const set = activeSet
    ? new Set(_cfAllVals.filter((value) => activeSet.has(value)))
    : new Set(_cfAllVals);
  $("colfiltTitle").textContent = `${field}（${allCount} 个值${isJob || isCustomer ? "；此列跨表共享" : ""}）`;
  $("colfiltSearch").value = "";
  $("colfiltAll").checked = _cfAllVals.every((value) => set.has(value));
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
  if (field === "JOB No" || field === "客户") {
    const prop = field === "JOB No" ? "jobFilter" : "customerFilter";
    let set = state[prop]
      ? new Set(_cfAllVals.filter((candidate) => state[prop].has(candidate)))
      : new Set(_cfAllVals);
    if (checked) set.add(value); else set.delete(value);
    state[prop] = _cfAllVals.every((candidate) => set.has(candidate)) ? null : set;
    $("colfiltAll").checked = !state[prop];
    renderGrid();
    return;
  }
  state.colFilters[table] = state.colFilters[table] || {};
  let set = state.colFilters[table][field];
  set = set
    ? new Set(_cfAllVals.filter((candidate) => set.has(candidate)))
    : new Set(_cfAllVals);
  state.colFilters[table][field] = set;
  if (checked) set.add(value); else set.delete(value);
  if (_cfAllVals.every((candidate) => set.has(candidate))) delete state.colFilters[table][field];
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
  if (field === "JOB No" || field === "客户") {
    const prop = field === "JOB No" ? "jobFilter" : "customerFilter";
    state[prop] = checked ? null : new Set();
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
  else if (field === "客户") state.customerFilter = null;
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
  if (r.data) recomputeDerived();
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

async function refreshWithFeedback(store, xlsx) {
  try {
    await refresh(store, xlsx);
  } catch (err) {
    setStatus("路径切换失败：" + String(err), "error");
    toast("路径切换失败：" + String(err), "error");
  }
}

$("btnPickStore").addEventListener("click", async () => {
  if (state.dirty && !confirm("当前有未保存更改。切换数据库会放弃这些更改，确定继续吗？")) return;
  try {
    let p = await call("pick_store");
    if (!p) p = window.prompt("请输入 SQLite 或 JSON 文件的完整路径：", state.storePath || "");
    if (p) {
      $("setStoreInput").value = p;
      await refreshWithFeedback(p, null);
    }
  } catch (err) {
    setStatus("无法打开文件选择器，请粘贴完整路径：" + String(err), "error");
    toast("无法打开文件选择器，请使用下方输入框", "error");
  }
});

$("btnPickXlsx").addEventListener("click", async () => {
  try {
    let p = await call("pick_xlsx");
    if (!p) p = window.prompt("请输入 XLSX 文件的完整路径：", state.xlsxPath || "");
    if (p) {
      $("setXlsxInput").value = p;
      await refreshWithFeedback(null, p);
    }
  } catch (err) {
    setStatus("无法打开文件选择器，请粘贴完整路径：" + String(err), "error");
    toast("无法打开文件选择器，请使用下方输入框", "error");
  }
});

$("btnApplyStorePath").addEventListener("click", async () => {
  if (state.dirty && !confirm("当前有未保存更改。切换数据库会放弃这些更改，确定继续吗？")) return;
  const p = $("setStoreInput").value.trim();
  if (!p) return toast("请先输入数据库路径", "error");
  await refreshWithFeedback(p, null);
});

$("btnApplyXlsxPath").addEventListener("click", async () => {
  const p = $("setXlsxInput").value.trim();
  if (!p) return toast("请先输入 XLSX 路径", "error");
  await refreshWithFeedback(null, p);
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
  setStatus(r.config_warning || (r.migration
    ? `已从旧 JSON 迁移到本地数据库：${r.store_path}` : "就绪"),
    r.config_warning ? "error" : "");
  if (r.config_warning) toast(r.config_warning, "error");
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
    setStatus(r.config_warning || (r.migration
      ? `已从旧 JSON 迁移到本地数据库：${r.store_path}` : "就绪"),
      r.config_warning ? "error" : "");
    if (r.config_warning) toast(r.config_warning, "error");
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

// ========== 各业务表独立录入 modal ==========
const ENTRY_TITLES = {
  "付款条件": "录入付款条件",
  "设备台账": "录入设备",
  "发货批次": "新增发货批次",
  "开票记录": "新增开票记录",
  "回款记录": "新增回款记录",
};

function customerForJob(job) {
  return s(state.data["合同订单"].find((row) => s(row["JOB No"]) === s(job))?.["客户"]);
}

function entryJobOptions(selected) {
  const contracts = [...state.data["合同订单"]].sort((a, b) =>
    s(a["JOB No"]).localeCompare(s(b["JOB No"]), "zh"));
  return '<option value="">请选择 JOB No…</option>' + contracts.map((row) => {
    const job = s(row["JOB No"]), customer = s(row["客户"]);
    return `<option value="${esc(job)}" ${job === selected ? "selected" : ""}>${esc(job)}${customer ? " · " + esc(customer) : ""}</option>`;
  }).join("");
}

function entrySchemaOptions(table, field, selected, blankLabel = "请选择…") {
  const spec = fieldsOf(table).find((item) => item.name === field);
  const options = spec?.options || [];
  return `<option value="">${esc(blankLabel)}</option>` + options.filter(Boolean).map((value) =>
    `<option value="${esc(value)}" ${s(value) === s(selected) ? "selected" : ""}>${esc(value)}</option>`
  ).join("");
}

function entryJobSection(job, title = "① 关联订单") {
  return `<div class="entry-section">
    <h4>${esc(title)}</h4>
    <div class="entry-grid">
      <label>JOB No<select id="entryJob">${entryJobOptions(job)}</select></label>
      <label>客户（自动带入）<input id="entryCustomer" readonly value="${esc(customerForJob(job))}" placeholder="选择 JOB 后自动显示"></label>
    </div>
  </div>`;
}

function entryPaymentTermRow(rec = {}) {
  const rid = s(rec["记录ID"]);
  return `<tr data-rid="${esc(rid)}">
    <td><select class="et-kind">${entrySchemaOptions("付款条件", "款类", rec["款类"], "选择款类")}</select></td>
    <td><input class="et-ratio" type="number" min="0" max="100" step="any" value="${esc(val(rec, "比例%"))}"></td>
    <td><input class="et-days" type="number" min="0" step="1" value="${esc(val(rec, "账期天数"))}"></td>
    <td><select class="et-trigger">${entrySchemaOptions("付款条件", "触发条件", rec["触发条件"], "选择触发条件")}</select></td>
    <td><input class="et-desc" value="${esc(val(rec, "说明"))}" placeholder="条款说明"></td>
    <td class="entry-term-del"><button type="button" data-entry-act="del-term" title="删除本条">✕</button></td>
  </tr>`;
}

function renderPaymentTermsEntry(job) {
  const existing = job ? state.data["付款条件"].filter((row) => s(row["JOB No"]) === job) : [];
  const rows = existing.length ? existing : [{}];
  return entryJobSection(job) + `<div class="entry-section">
    <h4><span>② 完整付款条款</span><span>
      <button type="button" class="noc-mini" data-entry-act="add-term">＋ 添加条款</button>
      <button type="button" class="noc-mini" data-entry-template="A">A模板</button>
      <button type="button" class="noc-mini" data-entry-template="B">B模板</button>
      <button type="button" class="noc-mini" data-entry-template="C">C模板</button>
    </span></h4>
    ${job ? `<table class="entry-term-table"><thead><tr><th style="width:110px">款类</th><th style="width:72px">比例%</th><th style="width:80px">账期天数</th><th style="width:135px">触发条件</th><th>说明</th><th style="width:36px"></th></tr></thead><tbody id="entryTermRows">${rows.map(entryPaymentTermRow).join("")}</tbody></table>
      <div id="entryTermSummary" class="entry-help"></div>`
      : '<div class="entry-list-empty">请先选择 JOB No，再录入该订单的完整付款条件。</div>'}
    <p class="entry-help">提交会用这里的条款替换该 JOB 当前全部付款条件；合计必须为 100%。已被开票或回款引用的款类会受到关联保护。</p>
  </div>`;
}

function pendingDeviceOptions(job, selectedId) {
  const rows = state.data["设备台账"].filter((row) =>
    s(row["JOB No"]) === job && !s(row["製造番号"]));
  return '<option value="">新增一台设备</option>' + rows.map((row, index) => {
    const rid = s(row["记录ID"]);
    const label = `待完善 ${index + 1} · ${s(row["设备型号"]) || "无型号"} · 批次 ${s(row["发货批次"]) || "—"} · PO ${s(row["PO No"]) || "—"}`;
    return `<option value="${esc(rid)}" ${rid === selectedId ? "selected" : ""}>${esc(label)}</option>`;
  }).join("");
}

function renderDeviceEntry(job) {
  const targetId = s(state._entry.deviceTargetId);
  const target = state.data["设备台账"].find((row) =>
    s(row["JOB No"]) === job && s(row["记录ID"]) === targetId && !s(row["製造番号"])) || {};
  const batches = state.data["发货批次"].filter((row) => s(row["JOB No"]) === job);
  const batchOptions = '<option value="">请选择该 JOB 的批次…</option>' + batches.map((row) => {
    const batch = s(row["发货批次"]);
    return `<option value="${esc(batch)}" ${batch === s(target["发货批次"]) ? "selected" : ""}>${esc(batch)}${s(row["出荷日"]) ? " · " + esc(s(row["出荷日"])) : ""}</option>`;
  }).join("");
  return entryJobSection(job) + `<div class="entry-section">
    <h4>② 设备身份</h4>
    ${job ? `<div class="entry-grid">
      <label class="full">录入方式<select id="entryDeviceTarget">${pendingDeviceOptions(job, targetId)}</select></label>
      <label>PO No<input id="entryPo" list="dlPo" autocomplete="off" value="${esc(val(target, "PO No"))}"></label>
      <label>设备型号<input id="entryModel" list="dlModel" autocomplete="off" value="${esc(val(target, "设备型号"))}"></label>
      <label>製造番号<input id="entrySerial" autocomplete="off" value="${esc(val(target, "製造番号"))}" placeholder="未知时可稍后补录"></label>
      <label>機番<input id="entryMachine" autocomplete="off" value="${esc(val(target, "機番"))}"></label>
      <label>未税单价<input id="entryPrice" type="number" step="any" min="0" value="${esc(val(target, "未税单价"))}"></label>
      <label>是否无偿<select id="entryFree">${entrySchemaOptions("设备台账", "是否无偿", s(target["是否无偿"]) || "否", "请选择")}</select></label>
    </div>` : '<div class="entry-list-empty">请先选择 JOB No。</div>'}
  </div>
  <div class="entry-section">
    <h4>③ 发货与验收</h4>
    ${job ? `<div class="entry-grid three">
      <label>发货批次<select id="entryBatch">${batchOptions}</select></label>
      <label>送货单回收<select id="entryReceipt">${entrySchemaOptions("设备台账", "送货单回收", target["送货单回收"], "未回收")}</select></label>
      <label>质保期<input id="entryWarranty" value="${esc(val(target, "质保期"))}" placeholder="如 12个月"></label>
      <label>质保开始日<input id="entryWarrantyStart" type="date" value="${esc(val(target, "质保开始日"))}"></label>
      <label>质保结束日<input id="entryWarrantyEnd" type="date" value="${esc(val(target, "质保结束日"))}"></label>
      <label class="full">备注<textarea id="entryNote" rows="2">${esc(val(target, "备注"))}</textarea></label>
    </div>
    ${batches.length ? '<p class="entry-help">批次列表严格限定为当前 JOB；验收状态由质保开始日自动计算。</p>' : '<div class="entry-warning">该 JOB 尚无发货批次。请先到“发货批次”页新增批次，设备不能引用其他 JOB 的批次。</div>'}` : '<div class="entry-list-empty">选择 JOB 后显示该订单的批次。</div>'}
  </div>`;
}

function renderShipmentDevices(job) {
  const devices = state.data["设备台账"].filter((row) => s(row["JOB No"]) === job);
  if (!devices.length) return '<div class="entry-list-empty">该 JOB 暂无设备；可以先建立空批次，再通过“录入设备”选择该批次。</div>';
  return `<div class="entry-list">${devices.map((row) => `<label class="entry-device-option">
    <input type="checkbox" class="entry-ship-device" value="${esc(s(row["记录ID"]))}">
    <span class="serial">${esc(s(row["製造番号"]) || "（待编号设备）")}</span>
    <span>${esc(s(row["设备型号"]) || "—")}</span>
    <span class="meta">当前批次 ${esc(s(row["发货批次"]) || "—")} · 机番 ${esc(s(row["機番"]) || "—")}</span>
  </label>`).join("")}</div>`;
}

function renderShipmentEntry(job) {
  return entryJobSection(job) + `<div class="entry-section">
    <h4>② 批次信息</h4>
    ${job ? `<div class="entry-grid">
      <label>发货批次名<input id="entryShipmentName" autocomplete="off" placeholder="如 2"></label>
      <label>出荷日<input id="entryShipmentDate" type="date"></label>
    </div>` : '<div class="entry-list-empty">请先选择 JOB No。</div>'}
  </div><div class="entry-section">
    <h4>③ 将现有设备移入本批次（可选）</h4>
    ${job ? renderShipmentDevices(job) : '<div class="entry-list-empty">选择 JOB 后仅显示该订单的设备。</div>'}
    <p class="entry-help">列表不会出现其他 JOB 的设备。未勾选设备时仍可先建立空批次，之后在“录入设备”中引用。</p>
  </div>`;
}

function paymentKindsForJob(job) {
  return [...new Set(state.data["付款条件"]
    .filter((row) => s(row["JOB No"]) === job && s(row["款类"]))
    .map((row) => s(row["款类"])))];
}

function renderTransactionEntry(table, job) {
  const draft = state._entry.draft;
  draft["JOB No"] = job;
  draft["客户"] = customerForJob(job);
  const isInvoice = table === "开票记录";
  const kinds = paymentKindsForJob(job);
  const kindOptions = '<option value="">请选择付款条件中的款类…</option>' + kinds.map((kind) =>
    `<option value="${esc(kind)}" ${kind === s(draft["款类"]) ? "selected" : ""}>${esc(kind)}</option>`).join("");
  const dateField = isInvoice ? "开票日" : "回款日";
  return entryJobSection(job) + `<div class="entry-section">
    <h4>② ${isInvoice ? "开票" : "回款"}信息</h4>
    ${job ? `<div class="entry-grid three">
      <label>款类<select id="entryKind" ${kinds.length ? "" : "disabled"}>${kindOptions}</select></label>
      <label>${dateField}<input id="entryTransactionDate" type="date" value="${esc(val(draft, dateField))}"></label>
      <label>含税金额<input id="entryAmount" type="number" min="0" step="any" value="${esc(val(draft, "含税金额"))}" placeholder="手工输入金额"></label>
      ${isInvoice ? `<label>状态<input value="已开" readonly></label><label>账期天数<input id="entryInvoiceDays" type="number" min="0" step="1" value="${esc(val(draft, "账期天数"))}"></label>` : ""}
    </div>
    ${kinds.length ? '<p class="entry-help">款类只读取当前 JOB 的付款条件，不会混入其他订单。</p>' : '<div class="entry-warning">该 JOB 尚未录入有效付款条件。请先到“付款条件”页录入完整条款后再新增本记录。</div>'}` : '<div class="entry-list-empty">请先选择 JOB No。</div>'}
  </div><div class="entry-section">
    <h4>③ 覆盖范围</h4>
    <div class="entry-coverage-actions">
      <div id="entryCoverageSummary" class="entry-coverage-summary">尚未选择覆盖范围</div>
      <button type="button" data-entry-coverage="batch" ${job ? "" : "disabled"}>按批次勾选</button>
      <button type="button" data-entry-coverage="serial" ${job ? "" : "disabled"}>按制造番号逐台选择</button>
    </div>
    <p class="entry-help">批次和设备列表均严格限定为当前 JOB；逐台选择制造番号后会自动回写对应批次，覆盖台数自动计算。</p>
  </div>`;
}

function renderEntryBody() {
  const entry = state._entry;
  if (!entry) return;
  const job = s(entry.job);
  let html = "";
  if (entry.table === "付款条件") html = renderPaymentTermsEntry(job);
  else if (entry.table === "设备台账") html = renderDeviceEntry(job);
  else if (entry.table === "发货批次") html = renderShipmentEntry(job);
  else html = renderTransactionEntry(entry.table, job);
  $("entryBody").innerHTML = html;
  $("entryErrors").classList.add("hidden");
  $("entryErrors").innerHTML = "";
  if (entry.table === "付款条件" && job) updateEntryTermSummary();
  if (["开票记录", "回款记录"].includes(entry.table)) updateEntryCoveragePreview();
}

function openEntryModal(table) {
  if (!ENTRY_TITLES[table] || !state.data) return;
  buildNocDatalists();
  const contextJob = currentContext().job;
  state._entry = {
    table,
    job: state.data["合同订单"].some((row) => s(row["JOB No"]) === contextJob) ? contextJob : "",
    deviceTargetId: "",
    draft: { "记录ID": newLocalId() },
  };
  $("entryTitle").textContent = ENTRY_TITLES[table];
  $("entryOk").textContent = table === "付款条件" ? "保存完整条款" : "确认录入";
  $("entrySafety").textContent = ["开票记录", "回款记录"].includes(table)
    ? "JOB、客户、款类与覆盖范围会在提交时再次交叉校验。"
    : "客户及自动字段由关联数据生成；提交前会检查跨表引用。";
  renderEntryBody();
  $("entryModal").classList.remove("hidden");
}

function closeEntryModal() {
  $("entryModal").classList.add("hidden");
  state._entry = null;
}

function showEntryErrors(errors) {
  const box = $("entryErrors");
  box.innerHTML = `<b>请修正以下 ${errors.length} 处问题：</b><ul>${errors.map((error) => `<li>${esc(error)}</li>`).join("")}</ul>`;
  box.classList.remove("hidden");
  box.scrollIntoView({ block: "nearest" });
  toast(`请修正 ${errors.length} 处问题`, "error");
}

function updateEntryTermSummary() {
  const box = $("entryTermSummary");
  if (!box) return;
  const rows = [...document.querySelectorAll("#entryTermRows tr")];
  const meaningful = rows.filter((row) => row.querySelector(".et-kind").value);
  const total = meaningful.reduce((sum, row) => sum + (Number(row.querySelector(".et-ratio").value) || 0), 0);
  const ok = meaningful.length && Math.abs(total - 100) <= 0.01;
  box.textContent = meaningful.length ? `当前合计 ${total}% / ${meaningful.length} 条${ok ? "  ✓" : "  ⚠ 必须为 100%"}` : "尚未录入有效条款";
  box.style.color = ok ? "var(--ok)" : "var(--danger)";
}

function syncEntryTransactionDraft() {
  const entry = state._entry;
  if (!entry || !["开票记录", "回款记录"].includes(entry.table)) return;
  const draft = entry.draft;
  draft["JOB No"] = entry.job;
  draft["客户"] = customerForJob(entry.job);
  if ($("entryKind")) draft["款类"] = $("entryKind").value;
  if ($("entryTransactionDate")) draft[entry.table === "开票记录" ? "开票日" : "回款日"] = $("entryTransactionDate").value;
  if ($("entryAmount")) draft["含税金额"] = $("entryAmount").value === "" ? null : Number($("entryAmount").value);
  if (entry.table === "开票记录") {
    draft["状态"] = "已开";
    if ($("entryInvoiceDays")) draft["账期天数"] = $("entryInvoiceDays").value === "" ? null : Math.round(Number($("entryInvoiceDays").value));
  }
}

function updateEntryCoveragePreview() {
  const entry = state._entry, box = $("entryCoverageSummary");
  if (!entry || !box || !["开票记录", "回款记录"].includes(entry.table)) return;
  const draft = entry.draft;
  const batches = coverageValues(draft["覆盖批次"]);
  const serials = coverageValues(draft["覆盖製造番号"]);
  const devices = state.data["设备台账"].filter((row) => s(row["JOB No"]) === s(entry.job));
  const covered = devices.filter((device) => coverageMatches(draft, device));
  box.textContent = batches.length || serials.length
    ? `已选 ${batches.length} 个批次 · ${serials.length} 个制造番号 · 自动覆盖 ${covered.length} 台设备`
    : (devices.length ? `尚未选择 · 当前 JOB 共 ${devices.length} 台设备` : "该 JOB 暂无设备，可录入预付款/开票后稍后补覆盖范围");
}

function readEntryTerms() {
  return [...document.querySelectorAll("#entryTermRows tr")].map((row) => ({
    rid: row.dataset.rid || "",
    kind: row.querySelector(".et-kind").value,
    ratio: Number(row.querySelector(".et-ratio").value),
    days: Number(row.querySelector(".et-days").value),
    trigger: row.querySelector(".et-trigger").value,
    desc: row.querySelector(".et-desc").value.trim(),
  })).filter((row) => row.kind || row.ratio || row.days || row.trigger || row.desc);
}

function commitPaymentTermsEntry() {
  const job = s(state._entry.job), rows = readEntryTerms(), errors = [];
  const allowed = new Set(fieldsOf("付款条件").find((field) => field.name === "款类").options);
  if (!job || !customerForJob(job)) errors.push("请选择有效 JOB No");
  if (!rows.length) errors.push("至少录入一条付款条件");
  rows.forEach((row, index) => {
    if (!allowed.has(row.kind)) errors.push(`第 ${index + 1} 条款类无效`);
    if (!(row.ratio > 0 && row.ratio <= 100)) errors.push(`第 ${index + 1} 条比例必须大于 0 且不超过 100`);
    if (!Number.isInteger(row.days) || row.days < 0) errors.push(`第 ${index + 1} 条账期天数必须是非负整数`);
  });
  const total = rows.reduce((sum, row) => sum + (Number(row.ratio) || 0), 0);
  if (rows.length && Math.abs(total - 100) > 0.01) errors.push(`付款条件比例合计 ${total}% ≠ 100%`);

  const previous = state.data["付款条件"].filter((row) => s(row["JOB No"]) === job);
  const nextKinds = new Set(rows.map((row) => row.kind));
  const renames = new Map();
  for (const old of previous) {
    const oldKind = s(old["款类"]);
    if (!oldKind || nextKinds.has(oldKind)) continue;
    const replacement = rows.find((row) => row.rid && row.rid === s(old["记录ID"]) && row.kind !== oldKind);
    const referenced = ["开票记录", "回款记录"].some((table) => state.data[table].some((rec) =>
      s(rec["JOB No"]) === job && s(rec["款类"]) === oldKind));
    if (replacement) renames.set(oldKind, replacement.kind);
    else if (referenced) errors.push(`款类“${oldKind}”已被开票/回款引用，不能直接删除；请保留或改名`);
  }
  if (errors.length) return showEntryErrors(errors);

  const nextRows = rows.map((row) => ({
    "记录ID": row.rid || newLocalId(), "JOB No": job, "款类": row.kind,
    "比例%": row.ratio, "账期天数": row.days, "触发条件": row.trigger,
    "说明": row.desc,
  }));
  state.data["付款条件"] = state.data["付款条件"].filter((row) => s(row["JOB No"]) !== job).concat(nextRows);
  for (const [before, after] of renames) {
    for (const table of ["开票记录", "回款记录"]) {
      for (const rec of state.data[table]) {
        if (s(rec["JOB No"]) === job && s(rec["款类"]) === before) rec["款类"] = after;
      }
    }
  }
  finishEntryCommit("付款条件", job, nextRows.length, `已保存 ${job} 的 ${nextRows.length} 条付款条件`);
}

function commitDeviceEntry() {
  const job = s(state._entry.job), targetId = s(state._entry.deviceTargetId), errors = [];
  const batch = s($("entryBatch")?.value), model = s($("entryModel")?.value).trim();
  const serial = s($("entrySerial")?.value).trim(), machine = s($("entryMachine")?.value).trim();
  const priceRaw = $("entryPrice")?.value ?? "";
  if (!job || !customerForJob(job)) errors.push("请选择有效 JOB No");
  if (!model) errors.push("设备型号不能为空");
  if (!batch || !state.data["发货批次"].some((row) => s(row["JOB No"]) === job && s(row["发货批次"]) === batch)) {
    errors.push("必须选择当前 JOB 已存在的发货批次");
  }
  if (priceRaw !== "" && (!Number.isFinite(Number(priceRaw)) || Number(priceRaw) < 0)) errors.push("未税单价必须是非负数字");
  const start = s($("entryWarrantyStart")?.value), end = s($("entryWarrantyEnd")?.value);
  if (start && end && end < start) errors.push("质保结束日不能早于质保开始日");
  if (serial && state.data["设备台账"].some((row) => row !== state.data["设备台账"].find((x) => s(x["记录ID"]) === targetId)
      && s(row["JOB No"]) === job && s(row["製造番号"]) === serial && s(row["機番"]) === machine)) {
    errors.push(`设备业务键重复：${job} / ${serial} / ${machine || "空机番"}`);
  }
  if (errors.length) return showEntryErrors(errors);
  const record = {
    "JOB No": job, "PO No": s($("entryPo").value).trim(), "设备型号": model,
    "製造番号": serial, "機番": machine, "未税单价": priceRaw === "" ? null : Number(priceRaw),
    "是否无偿": $("entryFree").value || "否", "发货批次": batch,
    "送货单回收": $("entryReceipt").value, "质保开始日": start,
    "质保结束日": end, "质保期": s($("entryWarranty").value).trim(),
    "备注": s($("entryNote").value).trim(),
  };
  let target = targetId ? state.data["设备台账"].find((row) =>
    s(row["JOB No"]) === job && s(row["记录ID"]) === targetId && !s(row["製造番号"])) : null;
  if (target) Object.assign(target, record);
  else {
    target = { "记录ID": newLocalId(), ...record };
    state.data["设备台账"].push(target);
  }
  finishEntryCommit("设备台账", job, 1, targetId ? `已完善设备 ${serial || model}` : `已新增设备 ${serial || model}`,
    state.data["设备台账"].indexOf(target));
}

function commitShipmentEntry() {
  const job = s(state._entry.job), batch = s($("entryShipmentName")?.value).trim(), errors = [];
  if (!job || !customerForJob(job)) errors.push("请选择有效 JOB No");
  if (!batch) errors.push("发货批次名不能为空");
  if (state.data["发货批次"].some((row) => s(row["JOB No"]) === job && s(row["发货批次"]) === batch)) errors.push(`该 JOB 已存在批次“${batch}”`);
  const selectedIds = new Set([...document.querySelectorAll(".entry-ship-device:checked")].map((box) => box.value));
  const selected = state.data["设备台账"].filter((row) => selectedIds.has(s(row["记录ID"])));
  if (selected.some((row) => s(row["JOB No"]) !== job)) errors.push("设备选择中出现了其他 JOB 的记录，请重新打开录入窗口");
  if (errors.length) return showEntryErrors(errors);
  const shipment = { "记录ID": newLocalId(), "JOB No": job, "发货批次": batch, "出荷日": s($("entryShipmentDate").value) };
  state.data["发货批次"].push(shipment);
  for (const device of selected) device["发货批次"] = batch;
  finishEntryCommit("发货批次", job, 1, `已新增批次 ${batch}${selected.length ? `，并移入 ${selected.length} 台设备` : ""}`,
    state.data["发货批次"].indexOf(shipment));
}

function isBlankTransaction(table, rec) {
  const dateField = table === "开票记录" ? "开票日" : "回款日";
  return !s(rec["款类"]) && !s(rec[dateField]) && !Number(rec["含税金额"])
    && !s(rec["覆盖批次"]) && !s(rec["覆盖製造番号"]);
}

function commitTransactionEntry() {
  syncEntryTransactionDraft();
  const entry = state._entry, table = entry.table, draft = entry.draft;
  const job = s(entry.job), errors = [], isInvoice = table === "开票记录";
  const dateField = isInvoice ? "开票日" : "回款日";
  const kinds = new Set(paymentKindsForJob(job));
  if (!job || !customerForJob(job)) errors.push("请选择有效 JOB No");
  if (!s(draft["款类"]) || !kinds.has(s(draft["款类"]))) errors.push("款类必须来自当前 JOB 已录入的付款条件");
  if (!s(draft[dateField])) errors.push(`${dateField}不能为空`);
  if (!(Number(draft["含税金额"]) > 0)) errors.push("含税金额必须大于 0");
  if (isInvoice && (!Number.isInteger(Number(draft["账期天数"])) || Number(draft["账期天数"]) < 0)) errors.push("账期天数必须是非负整数");
  const batches = coverageValues(draft["覆盖批次"]), serials = coverageValues(draft["覆盖製造番号"]);
  const jobDevices = state.data["设备台账"].filter((row) => s(row["JOB No"]) === job);
  const validBatches = new Set(state.data["发货批次"].filter((row) => s(row["JOB No"]) === job).map((row) => s(row["发货批次"])));
  const validSerials = new Set(jobDevices.map((row) => s(row["製造番号"])).filter(Boolean));
  const badBatch = batches.find((batch) => !validBatches.has(batch));
  const badSerial = serials.find((serial) => !validSerials.has(serial));
  if (badBatch) errors.push(`覆盖批次“${badBatch}”不属于当前 JOB`);
  if (badSerial) errors.push(`製造番号“${badSerial}”不属于当前 JOB`);
  if (jobDevices.length && !batches.length && !serials.length) errors.push("该 JOB 已有设备，必须按批次或制造番号选择覆盖范围");
  if (errors.length) return showEntryErrors(errors);

  const record = {
    "JOB No": job, "款类": s(draft["款类"]), [dateField]: s(draft[dateField]),
    "含税金额": Number(draft["含税金额"]), "覆盖批次": batches.join(";"),
    "覆盖製造番号": serials.join(";"),
  };
  if (isInvoice) {
    record["状态"] = "已开";
    record["账期天数"] = Number(draft["账期天数"]);
  }
  let target = state.data[table].find((row) => s(row["JOB No"]) === job && isBlankTransaction(table, row));
  if (target) Object.assign(target, record);
  else {
    target = { "记录ID": newLocalId(), ...record };
    state.data[table].push(target);
  }
  finishEntryCommit(table, job, 1, `已新增${isInvoice ? "开票" : "回款"}记录：${s(record["款类"])} / ${Number(record["含税金额"]).toLocaleString()}`,
    state.data[table].indexOf(target));
}

function finishEntryCommit(table, job, count, message, selectedIndex = null) {
  state.dirty = true;
  recomputeDerived();
  state.jobFilter = new Set([job]);
  state.customerFilter = null;
  state.selection = selectedIndex === null ? new Set() : new Set([selectedIndex]);
  closeEntryModal();
  renderGrid();
  renderCounts();
  renderSummary();
  toast(message || `已录入 ${count} 条记录`, "ok");
}

function commitEntryModal() {
  if (!state._entry) return;
  if (state._entry.table === "付款条件") commitPaymentTermsEntry();
  else if (state._entry.table === "设备台账") commitDeviceEntry();
  else if (state._entry.table === "发货批次") commitShipmentEntry();
  else commitTransactionEntry();
}

$("entryBody").addEventListener("change", (event) => {
  const entry = state._entry;
  if (!entry) return;
  if (event.target.id === "entryJob") {
    entry.job = event.target.value;
    entry.deviceTargetId = "";
    entry.draft = { "记录ID": newLocalId(), "JOB No": entry.job };
    renderEntryBody();
    return;
  }
  if (event.target.id === "entryDeviceTarget") {
    entry.deviceTargetId = event.target.value;
    renderEntryBody();
    return;
  }
  if (event.target.id === "entryKind" && entry.table === "开票记录") {
    const term = state.data["付款条件"].find((row) =>
      s(row["JOB No"]) === s(entry.job) && s(row["款类"]) === event.target.value);
    if ($("entryInvoiceDays")) $("entryInvoiceDays").value = term?.["账期天数"] ?? 0;
  }
  syncEntryTransactionDraft();
});

$("entryBody").addEventListener("input", (event) => {
  if (event.target.closest("#entryTermRows")) updateEntryTermSummary();
  syncEntryTransactionDraft();
});

$("entryBody").addEventListener("click", (event) => {
  const action = event.target.closest("[data-entry-act]")?.dataset.entryAct;
  if (action === "add-term") {
    if (!$("entryTermRows")) return showEntryErrors(["请先选择 JOB No"]);
    $("entryTermRows").insertAdjacentHTML("beforeend", entryPaymentTermRow());
    updateEntryTermSummary();
    return;
  }
  if (action === "del-term") {
    event.target.closest("tr").remove();
    updateEntryTermSummary();
    return;
  }
  const templateKey = event.target.closest("[data-entry-template]")?.dataset.entryTemplate;
  if (templateKey) {
    if (!$("entryTermRows")) return showEntryErrors(["请先选择 JOB No"]);
    const terms = PAYMENT_TEMPLATES[templateKey] || [];
    $("entryTermRows").innerHTML = terms.map((term) => entryPaymentTermRow({
      "款类": term.kind, "比例%": term.ratio, "账期天数": term.days,
      "触发条件": term.trigger, "说明": term.desc,
    })).join("");
    updateEntryTermSummary();
    return;
  }
  const coverageMode = event.target.closest("[data-entry-coverage]")?.dataset.entryCoverage;
  if (coverageMode) {
    syncEntryTransactionDraft();
    if (!s(state._entry.job)) return showEntryErrors(["请先选择 JOB No"]);
    openCoverageModal(state._entry.table, -1, coverageMode === "batch" ? "覆盖批次" : "覆盖製造番号", state._entry.draft);
  }
});

$("entryClose").addEventListener("click", closeEntryModal);
$("entryCancel").addEventListener("click", closeEntryModal);
$("entryOk").addEventListener("click", commitEntryModal);
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && state._entry && !state._coverage) closeEntryModal();
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
    <td><select class="nt-kind">${entrySchemaOptions("付款条件", "款类", "", "选择款类")}</select></td>
    <td><input class="nt-ratio" type="number" min="0" max="100" step="any" placeholder="0"></td>
    <td><input class="nt-days" type="number" min="0" step="1" placeholder="0"></td>
    <td><select class="nt-trigger">${entrySchemaOptions("付款条件", "触发条件", "", "选择触发条件")}</select></td>
    <td><input class="nt-desc" autocomplete="off" placeholder="条款说明"></td>
    <td class="noc-del"><button type="button" class="nt-del">✕</button></td>
  </tr>`;
}
function addNocTermRow(kind, ratio, days, trigger, desc) {
  const tr = document.createElement("tr");
  tr.innerHTML = nocTermRowHtml();
  $("nocTermRows").appendChild(tr);
  if (kind !== undefined && kind !== "") {
    tr.querySelector(".nt-kind").value = kind;
    tr.querySelector(".nt-ratio").value = ratio;
    tr.querySelector(".nt-days").value = days;
    tr.querySelector(".nt-trigger").value = trigger || "";
    tr.querySelector(".nt-desc").value = desc || "";
  }
}
function fillNocTemplate(key) {
  const terms = PAYMENT_TEMPLATES[key];
  if (!terms) return;
  $("nocTermRows").innerHTML = "";
  for (const t of terms) addNocTermRow(t.kind, t.ratio, t.days, t.trigger, t.desc);
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
      trigger: tr.querySelector(".nt-trigger").value,
      desc: tr.querySelector(".nt-desc").value.trim(),
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
    const allowedKinds = new Set(fieldsOf("付款条件").find((field) => field.name === "款类").options);
    f.terms.forEach((term, index) => {
      if (!allowedKinds.has(term.kind)) errs.push("第 " + (index + 1) + " 条付款款类无效");
      if (!(term.ratio > 0 && term.ratio <= 100)) errs.push("第 " + (index + 1) + " 条付款比例必须大于 0 且不超过 100");
      if (!Number.isInteger(term.days) || term.days < 0) errs.push("第 " + (index + 1) + " 条账期天数必须是非负整数");
    });
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
        "账期天数": Number(t.days), "触发条件": t.trigger, "说明": t.desc,
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
