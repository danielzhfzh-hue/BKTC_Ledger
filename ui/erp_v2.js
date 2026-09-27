"use strict";

(function () {
  const oldSwitchTab = window.switchTab;
  const oldRenderAll = window.renderAll;
  let selectedJob = "";
  let erpWorkspace = null;
  let erpMasterData = null;
  let selectedDetailTab = "总览";
  let selectedFinanceBucket = "";
  let workspaceRequest = 0;

  function money(value, currency) {
    const n = Number(value) || 0;
    const symbol = currency === "USD" ? "$" : currency === "EUR" ? "€" : "¥";
    return symbol + n.toLocaleString(undefined, { maximumFractionDigits: 0 });
  }

  function tr(value) {
    return window.I18N?.t?.(value) || value;
  }

  function pct(n, d) {
    return d > 0 ? Math.max(0, Math.min(100, Math.round(n / d * 100))) : 0;
  }

  function snapshots() {
    return erpWorkspace?.orders || [];
  }

  function recordsFor(table, job) {
    return (state.data?.[table] || []).filter(function (row) {
      return s(row["JOB No"]) === job;
    });
  }

  function formatAmounts(amounts) {
    const entries = Object.entries(amounts || {}).filter(function (entry) {
      return Math.abs(Number(entry[1]) || 0) > 0.0001;
    });
    return entries.length
      ? entries.map(function (entry) { return money(entry[1], entry[0]); }).join(" / ")
      : money(0, "RMB");
  }

  function amountForCurrency(amounts, currency) {
    return Number((amounts || {})[currency] || 0);
  }

  async function refreshErpWorkspace() {
    if (!state.data) return;
    const request = ++workspaceRequest;
    try {
      const result = await call("get_erp_workspace", state.data, state.rules);
      if (request !== workspaceRequest) return;
      erpWorkspace = result;
      renderDashboard();
      renderOrderCenter();
      renderFinance();
      if (window.I18N?.apply) window.I18N.apply(document.body);
    } catch (error) {
      console.error("ERP workspace:", error);
      if (state.tab === "工作台" || state.tab === "订单中心" || state.tab === "财务驾驶舱") {
        toast(String(error), "error");
      }
    }
  }

  function progress(value) {
    return '<span class="erp-progress"><i><b style="width:' + value + '%"></b></i><span>' + value + '%</span></span>';
  }

  function renderRecent() {
    if (!$("erpRecentOrders")) return;
    const q = s($("erpDashboardFilter")?.value).toLowerCase();
    const rows = snapshots().filter(function (x) {
      return !q || [x.job, x.customer, x.content].join(" ").toLowerCase().includes(q);
    }).slice(0, 40);
    $("erpRecentOrders").innerHTML = rows.map(function (x) {
      return '<tr data-job="' + esc(x.job) + '"><td><b>' + esc(x.job) + '</b></td><td>' + esc(x.customer) +
        '</td><td>' + esc(x.content) + '</td><td>' + money(x.total_untaxed, x.currency) +
        '</td><td><span class="erp-stage-pill">' + esc(x.stage) + '</span></td><td>' + x.shipped_count + '/' + x.device_count +
        '</td><td>' + x.accepted_count + '/' + x.device_count + '</td><td>' + progress(x.invoice_pct) +
        '</td><td>' + progress(x.paid_pct) + '</td><td><span class="risk ' + x.risk_level + '">' + esc(x.risk_label) + '</span></td></tr>';
    }).join("") || '<tr><td colspan="10" class="muted">暂无订单</td></tr>';
  }

  function renderWorkList() {
    if (!$("erpWorkList")) return;
    const work = (erpWorkspace?.actions || []).filter(function (x) { return x.manual; });
    $("erpWorkList").innerHTML = work.length ? work.map(function (x) {
      const due = x.due_date ? esc(x.due_date) : "无截止日";
      return '<div class="erp-work-row" data-work-id="' + esc(x.source_id) + '">' +
        '<span class="risk ' + esc(x.priority) + '">' + (x.priority === "high" ? "高" : x.priority === "low" ? "低" : "中") + '</span>' +
        '<span class="job">' + esc(x.job || "—") + '</span>' +
        '<span><b>' + esc(x.title) + '</b><small class="due">' + due + (x.detail ? " · " + esc(x.detail) : "") + '</small></span>' +
        '<span><button data-work-postpone="' + esc(x.source_id) + '">延期</button> <button data-work-complete="' + esc(x.source_id) + '">完成</button></span></div>';
    }).join("") : '<div class="erp-empty-state" style="min-height:100px"><b>暂无手工待办</b><span>可为订单建立跟进任务。</span></div>';
  }

  function renderDashboard() {
    if (!$("erpKpis") || !erpWorkspace) return;
    const k = erpWorkspace.kpis || {};
    $("erpToday").textContent = new Date().toLocaleDateString(window.I18N?.locale?.() || "zh-CN", { year: "numeric", month: "long", day: "numeric", weekday: "short" });
    const kpis = [
      ["订单", k.orders || 0, "当前业务订单", ""],
      ["今日行动", k.action_count || 0, (k.high_action_count || 0) + " 项高优先级", k.high_action_count ? "danger" : ""],
      ["待发货", k.pending_ship_orders || 0, "尚未完成出货的订单", "warn"],
      ["待验收", (k.pending_accept_devices || 0) + " 台", "已发货未验收设备", "warn"],
      ["应收", formatAmounts(k.open_amounts), "当前未回收金额", ""],
      ["已逾期", formatAmounts(k.overdue_amounts), "需要优先跟进", Object.values(k.overdue_amounts || {}).some(Number) ? "danger" : ""]
    ];
    $("erpKpis").innerHTML = kpis.map(function (item) {
      return '<div class="erp-kpi ' + item[3] + '"><div class="label">' + item[0] + '</div><div class="value">' + item[1] +
        '</div><div class="sub">' + item[2] + '</div></div>';
    }).join("");

    const attention = (erpWorkspace.actions || []).filter(function (x) { return !x.manual && x.priority !== "low"; }).slice(0, 14);
    $("erpAttentionList").innerHTML = attention.length ? attention.map(function (x) {
      const amount = x.amount ? " · " + money(x.amount, x.currency) : "";
      const due = x.due_date ? " · " + esc(x.due_date) : "";
      return '<div class="erp-attention-row" data-job="' + esc(x.job) + '"><b>' + esc(x.job || "—") + '</b><span>' +
        esc(x.title) + amount + due + (x.detail ? '<small> · ' + esc(x.detail) + '</small>' : '') +
        '</span><span class="risk ' + esc(x.priority) + '">' + (x.priority === "high" ? "高风险" : "需关注") + '</span></div>';
    }).join("") : '<div class="erp-empty-state" style="min-height:150px"><b>暂无高优先级待办</b><span>当前订单状态正常。</span></div>';

    const stages = erpWorkspace.stage_counts || [];
    const maxCount = Math.max.apply(null, [1].concat(stages.map(function (x) { return x.count; })));
    $("erpStageBoard").innerHTML = stages.map(function (x) {
      return '<div class="erp-stage-row"><span>' + esc(x.stage) + '</span><span class="erp-stage-bar"><i style="width:' +
        Math.round(x.count / maxCount * 100) + '%"></i></span><b>' + x.count + '</b></div>';
    }).join("");
    renderWorkList();
    renderRecent();
  }

  const DETAIL_TABS = ["总览", "设备", "发货", "开票", "回款", "付款条件", "审计"];

  function detailTable(table, fields, job) {
    const rows = recordsFor(table, job);
    if (!rows.length) return '<div class="erp-empty-state" style="min-height:140px"><b>暂无记录</b></div>';
    return '<div class="erp-detail-table-wrap"><table class="erp-detail-table"><thead><tr>' +
      fields.map(function (field) { return '<th>' + esc(field) + '</th>'; }).join("") +
      '</tr></thead><tbody>' + rows.map(function (row) {
        return '<tr>' + fields.map(function (field) {
          const value = row[field];
          const text = value === null || value === undefined || value === "" ? "—" : String(value);
          return '<td>' + esc(text) + '</td>';
        }).join("") + '</tr>';
      }).join("") + '</tbody></table></div>';
  }

  function renderFlow(flow, currency) {
    const nodes = flow?.nodes || [];
    const edges = flow?.edges || [];
    if (!nodes.length) return '<span class="muted">暂无业务流数据</span>';
    const labels = { quote: "报价", order: "订单", shipment: "发货", invoice: "开票", payment: "回款" };
    const byId = Object.fromEntries(nodes.map(function (node) { return [node.id, node]; }));
    const card = function (node) {
      if (!node) return "";
      const meta = [node.date, node.count ? node.count + "台" : "", node.amount ? money(node.amount, currency || "RMB") : ""].filter(Boolean).join(" · ");
      return '<div class="erp-flow-node"><b>' + esc(labels[node.type] || node.type) + ' · ' + esc(node.label) +
        '</b><small>' + esc(meta || " ") + '</small></div>';
    };
    if (!edges.length) return '<div class="erp-flow-line">' + nodes.map(card).join('<div class="erp-flow-arrow">→</div>') + '</div>';
    return edges.map(function (edge) {
      return '<div class="erp-flow-line">' + card(byId[edge.from]) + '<div class="erp-flow-arrow">→</div>' +
        card(byId[edge.to]) + '</div>';
    }).join("");
  }

  function detailBody(x) {
    const job = x.job;
    if (selectedDetailTab === "设备") {
      return '<div class="erp-detail-section"><div class="erp-card-head"><h3>设备台账</h3><button data-erp-open="设备台账">深度编辑</button></div>' +
        detailTable("设备台账", ["设备型号", "製造番号", "機番", "PO No", "发货批次", "验收状态", "质保开始日", "质保结束日"], job) + '</div>';
    }
    if (selectedDetailTab === "发货") {
      return '<div class="erp-detail-section"><div class="erp-card-head"><h3>发货批次</h3><button data-erp-open="发货批次">深度编辑</button></div>' +
        detailTable("发货批次", ["发货批次", "出荷日", "台数", "未税合计", "含税合计", "覆盖製造番号"], job) + '</div>';
    }
    if (selectedDetailTab === "开票") {
      return '<div class="erp-detail-section"><div class="erp-card-head"><h3>开票记录</h3><button data-erp-open="开票记录">深度编辑</button></div>' +
        detailTable("开票记录", ["款类", "开票日", "含税金额", "应收回款日", "回款状态", "覆盖批次"], job) + '</div>';
    }
    if (selectedDetailTab === "回款") {
      return '<div class="erp-detail-section"><div class="erp-card-head"><h3>回款记录</h3><button data-erp-open="回款记录">深度编辑</button></div>' +
        detailTable("回款记录", ["款类", "回款日", "含税金额", "对应应收回款日", "是否超期", "覆盖批次"], job) + '</div>';
    }
    if (selectedDetailTab === "付款条件") {
      return '<div class="erp-detail-section"><div class="erp-card-head"><h3>付款条件</h3><button data-erp-open="付款条件">深度编辑</button></div>' +
        detailTable("付款条件", ["款类", "比例%", "账期天数", "触发条件", "说明"], job) + '</div>';
    }
    if (selectedDetailTab === "审计") {
      return '<div class="erp-detail-section"><h3>订单审计</h3><div id="erpOrderAudit" class="muted">正在读取审计记录…</div></div>';
    }

    const order = recordsFor("合同订单", job)[0] || {};
    const actions = (erpWorkspace?.actions || []).filter(function (a) { return a.job === job; }).slice(0, 8);
    return '<div class="erp-detail-metrics">' +
      '<div class="erp-detail-metric"><span>订单未税金额</span><b>' + money(x.total_untaxed, x.currency) + '</b></div>' +
      '<div class="erp-detail-metric"><span>发货进度</span><b>' + x.shipped_count + '/' + x.device_count + '</b></div>' +
      '<div class="erp-detail-metric"><span>验收进度</span><b>' + x.accepted_count + '/' + x.device_count + '</b></div>' +
      '<div class="erp-detail-metric"><span>已开票</span><b>' + money(x.invoice_amount, x.currency) + '</b></div>' +
      '<div class="erp-detail-metric"><span>已回款</span><b>' + money(x.paid_amount, x.currency) + '</b></div></div>' +
      (x.source_quote_no ? '<div class="erp-detail-section"><h3>来源报价</h3><b>' + esc(x.source_quote_no) + '</b><span class="muted"> → ' + esc(job) + '</span></div>' : '') +
      '<div class="erp-detail-section"><h3>Document Flow</h3><div class="erp-flow-lanes">' + renderFlow(x.document_flow, x.currency) + '</div></div>' +
      '<div class="erp-detail-section"><h3>当前行动</h3>' +
      (actions.length ? actions.map(function (a) {
        return '<div class="erp-attention-row" data-job="' + esc(job) + '"><span class="risk ' + esc(a.priority) + '">' +
          (a.priority === "high" ? "高" : a.priority === "low" ? "低" : "中") + '</span><span><b>' + esc(a.title) +
          '</b>' + (a.detail ? ' · ' + esc(a.detail) : '') + '</span><span>' + esc(a.due_date || "") + '</span></div>';
      }).join("") : '<span class="muted">暂无需要处理的异常。</span>') + '</div>' +
      '<div class="erp-detail-section"><h3>订单信息</h3><div class="erp-detail-actions">' +
      '<button data-erp-open="合同订单">编辑订单头</button><button data-add-work="' + esc(job) + '">＋ 新建待办</button>' +
      '<span class="muted">' + esc(s(order["担当者"])) + (s(order["送货地点"]) ? " · " + esc(s(order["送货地点"])) : "") + '</span></div></div>';
  }

  async function renderOrderAudit(job) {
    const box = $("erpOrderAudit");
    if (!box || selectedDetailTab !== "审计") return;
    try {
      const result = await call("get_audit_events", { job_no: job }, 100);
      const changes = [];
      for (const event of (result.events || [])) {
        for (const change of (event.changes || [])) {
          changes.push({ event: event, change: change });
          if (changes.length >= 40) break;
        }
        if (changes.length >= 40) break;
      }
      box.innerHTML = changes.length ? '<div class="erp-detail-table-wrap"><table class="erp-detail-table"><thead><tr><th>时间</th><th>操作人</th><th>对象</th><th>变更</th></tr></thead><tbody>' +
        changes.map(function (item) {
          const c = item.change, e = item.event;
          return '<tr><td>' + esc((e.created_at || "").replace("T", " ")) + '</td><td>' + esc(e.operator_name || "") +
            '</td><td>' + esc(c.table_name + (c.field_name ? " · " + c.field_name : "")) + '</td><td>' +
            esc(String(c.old_value ?? "—")) + ' → ' + esc(String(c.new_value ?? "—")) + '</td></tr>';
        }).join("") + '</tbody></table></div>' : '<span class="muted">暂无审计记录</span>';
    } catch (error) {
      box.textContent = String(error);
    }
  }

  function renderOrderDetail(job) {
    if (!$("erpOrderDetail") || !erpWorkspace) return;
    const x = snapshots().find(function (o) { return o.job === job; });
    if (!x) {
      $("erpOrderDetail").innerHTML = '<div class="erp-empty-state"><b>未找到订单</b></div>';
      return;
    }
    const life = ["受注", "生产", "出货", "验收", "开票", "回款"];
    const tabs = DETAIL_TABS.map(function (name) {
      const counts = {
        "设备": x.device_count,
        "发货": recordsFor("发货批次", job).length,
        "开票": recordsFor("开票记录", job).filter(function (r) { return s(r["款类"]) || Number(r["含税金额"]); }).length,
        "回款": recordsFor("回款记录", job).filter(function (r) { return s(r["款类"]) || Number(r["含税金额"]); }).length,
        "付款条件": recordsFor("付款条件", job).filter(function (r) { return s(r["款类"]); }).length
      };
      const count = counts[name];
      return '<button class="erp-detail-tab ' + (selectedDetailTab === name ? "active" : "") + '" data-detail-tab="' + esc(name) + '">' +
        esc(name) + (Number.isFinite(count) ? " " + count : "") + '</button>';
    }).join("");
    $("erpOrderDetail").innerHTML =
      '<div class="erp-detail-head"><div class="erp-detail-title"><div><h2>' + esc(x.job) + '</h2><div class="erp-detail-sub">' +
      esc(x.customer) + ' · ' + esc(x.content || "未填写订单内容") + ' · ' + x.device_count + ' 台</div></div><span class="risk ' +
      x.risk_level + '">' + esc(x.risk_label) + '</span></div></div>' +
      '<div class="erp-lifecycle">' + life.map(function (name, i) {
        const done = x.stage_index >= 6 || i < x.stage_index;
        const current = x.stage_index < 6 && i === x.stage_index;
        return '<div class="erp-life-step ' + (done ? "done" : current ? "current" : "") + '">' + name + '</div>';
      }).join("") + '</div><div class="erp-detail-tabs">' + tabs + '</div><div id="erpOrderDetailBody">' + detailBody(x) + '</div>';
    if (selectedDetailTab === "审计") renderOrderAudit(job);
  }

  function renderOrderCenter() {
    if (!$("erpOrderList") || !erpWorkspace) return;
    const orders = snapshots();
    const stageSel = $("erpOrderStageFilter");
    const stages = Array.from(new Set(orders.map(function (x) { return x.stage; })));
    const currentStage = s(stageSel?.value);
    if (stageSel) {
      stageSel.innerHTML = '<option value="">全部阶段</option>' + stages.map(function (x) { return '<option value="' + esc(x) + '">' + esc(x) + '</option>'; }).join("");
      stageSel.value = stages.includes(currentStage) ? currentStage : "";
    }
    const q = s($("erpOrderSearch")?.value).toLowerCase();
    const stageFilter = s(stageSel?.value);
    const visible = orders.filter(function (x) {
      return (!q || [x.job, x.customer, x.content].join(" ").toLowerCase().includes(q)) && (!stageFilter || x.stage === stageFilter);
    });
    if (!visible.some(function (x) { return x.job === selectedJob; })) selectedJob = visible[0]?.job || "";
    $("erpOrderList").innerHTML = visible.map(function (x) {
      return '<div class="erp-order-list-item ' + (x.job === selectedJob ? "active" : "") + '" data-job="' + esc(x.job) +
        '"><div class="top"><b>' + esc(x.job) + '</b><span class="erp-stage-pill">' + esc(x.stage) +
        '</span></div><div class="customer">' + esc(x.customer) + '</div><div class="meta"><span>' +
        esc(x.content || "—") + '</span><span>' + x.device_count + '台 · ' + money(x.total_untaxed, x.currency) + '</span></div></div>';
    }).join("") || '<div class="erp-empty-state"><b>没有匹配订单</b></div>';
    renderOrderDetail(selectedJob);
  }

  function openOrder(job) {
    selectedJob = job;
    window.switchTab("订单中心");
    renderOrderCenter();
  }

  function globalMatches(query) {
    const q = query.toLowerCase();
    if (!q || !state.data) return [];
    return snapshots().filter(function (x) {
      const deviceText = recordsFor("设备台账", x.job).map(function (d) {
        return s(d["PO No"]) + " " + s(d["製造番号"]) + " " + s(d["設備型号"]) + " " + s(d["设备型号"]);
      }).join(" ");
      return [x.job, x.customer, x.content, deviceText, x.source_quote_no].join(" ").toLowerCase().includes(q);
    }).slice(0, 12);
  }


  function sumBucketAmounts(buckets) {
    const totals = {};
    for (const bucket of buckets || []) {
      for (const [currency, amount] of Object.entries(bucket.amounts || {})) {
        totals[currency] = (totals[currency] || 0) + Number(amount || 0);
      }
    }
    return totals;
  }

  function bucketHtml(bucket, kind) {
    return '<div class="erp-bucket-row ' + (selectedFinanceBucket === kind + "::" + bucket.bucket ? "active" : "") +
      '" data-finance-kind="' + esc(kind) + '" data-finance-bucket="' + esc(bucket.bucket) + '">' +
      '<b>' + esc(bucket.bucket) + '</b><span>' + bucket.count + ' 笔</span><span class="erp-bucket-amounts">' +
      Object.entries(bucket.amounts || {}).map(function (entry) {
        return '<b>' + money(entry[1], entry[0]) + '</b>';
      }).join("") + '</span></div>';
  }

  function financeRowsForView() {
    const finance = erpWorkspace?.finance;
    if (!finance) return [];
    let rows = finance.rows || [];
    if (selectedFinanceBucket) {
      const [kind, bucketName] = selectedFinanceBucket.split("::");
      const bucket = (finance[kind] || []).find(function (x) { return x.bucket === bucketName; });
      rows = bucket?.rows || [];
    }
    const q = s($("erpFinanceFilter")?.value).toLowerCase();
    if (q) {
      rows = rows.filter(function (row) {
        return [row.job, row.customer, row["款类"], row["批次"], row["未回收原因"]].join(" ").toLowerCase().includes(q);
      });
    }
    return rows;
  }

  function renderFinance() {
    if (!$("erpFinanceKpis") || !erpWorkspace) return;
    const finance = erpWorkspace.finance || {};
    const due30 = sumBucketAmounts((finance.forecast || []).filter(function (x) {
      return x.bucket === "0-7天" || x.bucket === "8-30天";
    }));
    const confirm = (finance.aging || []).find(function (x) { return x.bucket === "待确认"; });
    const kpis = [
      ["应收总额", formatAmounts(finance.open_amounts), "当前未回收"],
      ["已逾期", formatAmounts(finance.overdue_amounts), "优先跟进"],
      ["30天内预计回款", formatAmounts(due30), "未来现金流"],
      ["待确认", (confirm?.count || 0) + " 笔", "缺少有效应收日期"],
      ["应收明细", (finance.rows || []).length + " 笔", "按 JOB / 款类下钻"]
    ];
    $("erpFinanceKpis").innerHTML = kpis.map(function (x, i) {
      return '<div class="erp-kpi ' + (i === 1 && Object.values(finance.overdue_amounts || {}).some(Number) ? "danger" : "") +
        '"><div class="label">' + x[0] + '</div><div class="value">' + x[1] + '</div><div class="sub">' + x[2] + '</div></div>';
    }).join("");
    $("erpAging").innerHTML = (finance.aging || []).map(function (x) { return bucketHtml(x, "aging"); }).join("");
    $("erpForecast").innerHTML = (finance.forecast || []).map(function (x) { return bucketHtml(x, "forecast"); }).join("");
    const rows = financeRowsForView();
    $("erpFinanceRows").innerHTML = rows.length ? rows.map(function (row) {
      const overdue = Number(row.days) < 0;
      return '<div class="erp-finance-row" data-job="' + esc(row.job) + '">' +
        '<b>' + esc(row.job) + '</b><span>' + esc(row.customer) + '</span><span>' + esc(row["款类"] || "—") +
        '</span><span>' + esc(row["批次"] || "—") + '</span><span class="amount">' + money(row.amount, row.currency) +
        '</span><span class="' + (overdue ? "overdue" : "") + '">' + esc(row.due_date || "待确认") +
        (row.days === null ? "" : " · " + (row.days < 0 ? "逾期" + Math.abs(row.days) + "天" : row.days + "天")) +
        '</span></div>';
    }).join("") : '<div class="erp-empty-state" style="min-height:160px"><b>当前筛选没有应收明细</b></div>';
  }

  async function loadMasterData(force) {
    if (erpMasterData && !force) { renderMasterData(); return; }
    try {
      erpMasterData = await call("list_master_data");
      renderMasterData();
      if (window.I18N?.apply) window.I18N.apply($("panel-主数据"));
    } catch (error) {
      toast(String(error), "error");
    }
  }

  function syncMasterDatalists() {
    if (!erpMasterData) return;
    const append = function (id, values) {
      const list = $(id);
      if (!list) return;
      const existing = new Set([...list.querySelectorAll("option")].map(function (o) { return o.value; }));
      for (const value of values.filter(Boolean)) {
        if (existing.has(value)) continue;
        list.insertAdjacentHTML("beforeend", '<option value="' + esc(value) + '">');
        existing.add(value);
      }
    };
    append("dlCustomer", (erpMasterData.customers || []).map(function (x) { return x.name; }));
    append("dlModel", (erpMasterData.models || []).map(function (x) { return x.model; }));
    append("dlAssist", (erpMasterData.customers || []).map(function (x) { return x.salesperson; }));
    append("dlShipTo", (erpMasterData.customers || []).map(function (x) { return x.ship_to; }));
  }

  function renderMasterData() {
    if (!erpMasterData || !$("erpMasterStats")) return;
    const customers = erpMasterData.customers || [];
    const models = erpMasterData.models || [];
    const templates = erpMasterData.payment_templates || [];
    const persistedCustomers = customers.filter(function (x) { return !x.inferred; }).length;
    const persistedModels = models.filter(function (x) { return !x.inferred; }).length;
    const stats = [
      ["客户", customers.length, persistedCustomers + " 已标准化"],
      ["设备型号", models.length, persistedModels + " 已标准化"],
      ["联系人", (erpMasterData.contacts || []).length, "客户联系人"],
      ["付款模板", templates.length, "可复用条款"]
    ];
    $("erpMasterStats").innerHTML = stats.map(function (x) {
      return '<div class="erp-kpi"><div class="label">' + x[0] + '</div><div class="value">' + x[1] +
        '</div><div class="sub">' + x[2] + '</div></div>';
    }).join("");
    $("erpCustomerMaster").innerHTML = customers.map(function (x) {
      return '<div class="erp-master-row ' + (x.inferred ? "inferred" : "") + '" data-master-kind="customer" data-master-id="' +
        esc(x.customer_id) + '"><span><b>' + esc(x.name) + '</b><small>' + esc([x.salesperson, x.currency, x.ship_to].filter(Boolean).join(" · ") || "尚未补充标准信息") +
        '</small></span><span class="tag">' + (x.inferred ? "历史值" : "已标准化") + '</span></div>';
    }).join("") || '<div class="erp-empty-state"><b>暂无客户</b></div>';
    $("erpContactMaster").innerHTML = (erpMasterData.contacts || []).map(function (x) {
      return '<div class="erp-master-row" data-master-kind="contact" data-master-id="' + esc(x.contact_id) +
        '"><span><b>' + esc(x.name) + '</b><small>' + esc([x.customer_name, x.title, x.email, x.phone].filter(Boolean).join(" · ")) +
        '</small></span><span class="tag">' + (Number(x.is_primary) ? "主要" : "联系人") + '</span></div>';
    }).join("") || '<div class="erp-empty-state"><b>暂无联系人</b></div>';
    $("erpModelMaster").innerHTML = models.map(function (x) {
      return '<div class="erp-master-row ' + (x.inferred ? "inferred" : "") + '" data-master-kind="model" data-master-id="' +
        esc(x.model_id) + '"><span><b>' + esc(x.model) + '</b><small>' + esc(x.description || x.unit || "台") +
        '</small></span><span class="tag">' + (x.inferred ? "历史值" : "已标准化") + '</span></div>';
    }).join("") || '<div class="erp-empty-state"><b>暂无型号</b></div>';
    $("erpPaymentMaster").innerHTML = templates.map(function (x) {
      return '<div class="erp-master-row" data-master-kind="payment_template" data-master-id="' + esc(x.template_id) +
        '"><span><b>' + esc(x.name) + '</b><small>' + esc(x.description || "付款条件模板") +
        '</small></span><span class="tag">' + ((x.items || []).length) + ' 条</span></div>';
    }).join("") || '<div class="erp-empty-state"><b>暂无付款条件模板</b></div>';
    syncMasterDatalists();
  }

  async function saveMaster(kind, existing) {
    existing = existing || {};
    let record = { ...existing };
    if (kind === "customer") {
      const name = prompt(tr("客户标准名称"), existing.name || "");
      if (name === null || !name.trim()) return;
      record = {
        ...record, name: name.trim(),
        currency: prompt(tr("默认币种"), existing.currency || "RMB") || "RMB",
        salesperson: prompt(tr("默认担当者"), existing.salesperson || "") || "",
        ship_to: prompt(tr("默认送货地点"), existing.ship_to || "") || "",
        address: prompt(tr("客户地址"), existing.address || "") || "",
        incoterm: prompt(tr("默认 Incoterm"), existing.incoterm || "") || "",
        active: 1
      };
    } else if (kind === "contact") {
      const name = prompt(tr("联系人姓名"), existing.name || "");
      if (name === null || !name.trim()) return;
      const customerName = prompt(tr("所属客户"), existing.customer_name || "");
      if (customerName === null) return;
      const customer = (erpMasterData?.customers || []).find(function (x) { return x.name === customerName; });
      record = {
        ...record, name: name.trim(), customer_name: customerName.trim(),
        customer_id: customer?.customer_id || "",
        title: prompt(tr("职务"), existing.title || "") || "",
        email: prompt("Email", existing.email || "") || "",
        phone: prompt(tr("电话"), existing.phone || "") || "",
        is_primary: confirm(tr("设为主要联系人？")) ? 1 : 0
      };
    } else if (kind === "model") {
      const model = prompt(tr("设备型号"), existing.model || "");
      if (model === null || !model.trim()) return;
      record = {
        ...record, model: model.trim(),
        description: prompt(tr("型号说明"), existing.description || "") || "",
        unit: prompt(tr("单位"), existing.unit || "台") || "台",
        active: 1
      };
    } else {
      const name = prompt(tr("付款条件模板名称"), existing.name || "");
      if (name === null || !name.trim()) return;
      const existingLines = (existing.items || []).map(function (item) {
        return [item.kind || "", item.ratio || 0, item.days || 0, item.trigger || "", item.description_zh || item.description || ""].join("|");
      }).join("；");
      const rawLines = prompt(tr("模板条款（款类|比例|账期天数|触发条件|说明；多条用分号分隔）"), existingLines);
      if (rawLines === null) return;
      const items = rawLines.split(/[；;\n]+/).map(function (line) {
        const parts = line.split("|").map(function (x) { return x.trim(); });
        return { kind: parts[0] || "", ratio: Number(parts[1]) || 0, days: Number(parts[2]) || 0,
          trigger: parts[3] || "", description_zh: parts[4] || "" };
      }).filter(function (item) { return item.kind; });
      record = {
        ...record, name: name.trim(),
        description: prompt(tr("模板说明"), existing.description || "") || "",
        currency: prompt(tr("适用币种（留空=不限）"), existing.currency || "") || "",
        items: items,
        active: 1
      };
    }
    try {
      const result = await call("save_master_record", kind, record);
      applyBackendState({ revision: result.revision });
      erpMasterData = result.master_data;
      renderMasterData();
      toast("主数据已保存", "ok");
    } catch (error) {
      toast(String(error), "error");
    }
  }

  function openWorkItem(job) {
    const select = $("workItemJob");
    const jobs = snapshots();
    select.innerHTML = '<option value="">不关联 JOB</option>' + jobs.map(function (x) {
      return '<option value="' + esc(x.job) + '">' + esc(x.job + " · " + x.customer) + '</option>';
    }).join("");
    select.value = job && jobs.some(function (x) { return x.job === job; }) ? job : "";
    $("workItemPriority").value = "medium";
    $("workItemTitle").value = "";
    $("workItemDue").value = "";
    $("workItemType").value = "follow_up";
    $("workItemNote").value = "";
    $("workItemModal").classList.remove("hidden");
    setTimeout(function () { $("workItemTitle").focus(); }, 0);
  }

  function closeWorkItem() {
    $("workItemModal").classList.add("hidden");
  }

  async function saveWorkItem() {
    const title = $("workItemTitle").value.trim();
    if (!title) return toast("待办标题不能为空", "error");
    try {
      const result = await call("save_work_item", {
        job_no: $("workItemJob").value,
        priority: $("workItemPriority").value,
        title: title,
        due_date: $("workItemDue").value,
        type: $("workItemType").value,
        note: $("workItemNote").value.trim(),
        status: "open"
      });
      applyBackendState({ revision: result.revision });
      closeWorkItem();
      await refreshErpWorkspace();
      toast("待办已创建", "ok");
    } catch (error) {
      toast(String(error), "error");
    }
  }

  window.switchTab = function (tab) {
    oldSwitchTab(tab);
    const erpViews = ["工作台", "订单中心", "财务驾驶舱", "主数据"];
    for (const view of erpViews) {
      const panel = $("panel-" + view);
      if (panel) panel.classList.toggle("hidden", tab !== view);
    }
    if (erpViews.includes(tab)) $("panel-表格").classList.add("hidden");
    document.querySelectorAll(".erp-nav-item, .erp-advanced button").forEach(function (b) {
      b.classList.toggle("active", b.dataset.view === tab);
    });
    if (tab === "工作台") renderDashboard();
    if (tab === "订单中心") renderOrderCenter();
    if (tab === "财务驾驶舱") renderFinance();
    if (tab === "主数据") loadMasterData(false);
  };

  window.renderAll = function () {
    oldRenderAll();
    refreshErpWorkspace();
    loadMasterData(false);
  };

  document.querySelectorAll(".erp-nav-item, .erp-advanced button").forEach(function (button) {
    button.addEventListener("click", function () {
      const view = button.dataset.view;
      if (view === "跨表查询") { $("btnQuery").click(); return; }
      if (view === "设置") { $("btnSettings").click(); return; }
      window.switchTab(view);
    });
  });

  $("erpQuickNew").addEventListener("click", openNewOrder);
  $("erpOrderNew").addEventListener("click", openNewOrder);
  $("erpOpenOrderCenter").addEventListener("click", function () { window.switchTab("订单中心"); });
  $("erpDashboardFilter").addEventListener("input", renderRecent);
  $("erpOrderSearch").addEventListener("input", renderOrderCenter);
  $("erpOrderStageFilter").addEventListener("change", renderOrderCenter);

  $("erpAddWorkItem").addEventListener("click", function () { openWorkItem(""); });
  $("workItemClose").addEventListener("click", closeWorkItem);
  $("workItemCancel").addEventListener("click", closeWorkItem);
  $("workItemSave").addEventListener("click", saveWorkItem);
  $("erpWorkList").addEventListener("click", async function (event) {
    const complete = event.target.closest("[data-work-complete]");
    const postpone = event.target.closest("[data-work-postpone]");
    if (!complete && !postpone) return;
    try {
      let result;
      if (complete) {
        result = await call("complete_work_item", complete.dataset.workComplete);
        toast("待办已完成", "ok");
      } else {
        const action = (erpWorkspace?.actions || []).find(function (x) {
          return x.manual && x.source_id === postpone.dataset.workPostpone;
        });
        if (!action?.work_item) return;
        const base = action.due_date ? new Date(action.due_date + "T00:00:00") : new Date();
        base.setDate(base.getDate() + 1);
        const suggested = [base.getFullYear(), String(base.getMonth() + 1).padStart(2, "0"),
          String(base.getDate()).padStart(2, "0")].join("-");
        const due = prompt(tr("新的截止日（YYYY-MM-DD）"), suggested);
        if (due === null || !due.trim()) return;
        result = await call("save_work_item", { ...action.work_item, due_date: due.trim(), status: "open" });
        toast("待办已延期", "ok");
      }
      applyBackendState({ revision: result.revision });
      await refreshErpWorkspace();
    } catch (error) {
      toast(String(error), "error");
    }
  });

  ["erpAging", "erpForecast"].forEach(function (id) {
    $(id).addEventListener("click", function (event) {
      const row = event.target.closest("[data-finance-bucket]");
      if (!row) return;
      const key = row.dataset.financeKind + "::" + row.dataset.financeBucket;
      selectedFinanceBucket = selectedFinanceBucket === key ? "" : key;
      renderFinance();
    });
  });
  $("erpFinanceFilter").addEventListener("input", renderFinance);
  $("erpFinanceRows").addEventListener("click", function (event) {
    const row = event.target.closest("[data-job]");
    if (row) openOrder(row.dataset.job);
  });

  $("erpAddCustomer").addEventListener("click", function () { saveMaster("customer", {}); });
  $("erpAddContact").addEventListener("click", function () { saveMaster("contact", {}); });
  $("erpAddModel").addEventListener("click", function () { saveMaster("model", {}); });
  $("erpAddPaymentTemplate").addEventListener("click", function () { saveMaster("payment_template", {}); });
  ["erpCustomerMaster", "erpContactMaster", "erpModelMaster", "erpPaymentMaster"].forEach(function (id) {
    $(id).addEventListener("click", function (event) {
      const row = event.target.closest("[data-master-kind]");
      if (!row || !erpMasterData) return;
      const kind = row.dataset.masterKind;
      const pools = {
        customer: erpMasterData.customers || [],
        contact: erpMasterData.contacts || [],
        model: erpMasterData.models || [],
        payment_template: erpMasterData.payment_templates || []
      };
      const idField = kind === "customer" ? "customer_id" : kind === "contact" ? "contact_id" : kind === "model" ? "model_id" : "template_id";
      const record = (pools[kind] || []).find(function (x) { return String(x[idField]) === row.dataset.masterId; }) || {};
      saveMaster(kind, record);
    });
  });

  document.addEventListener("keydown", function (event) {
    if (event.key === "Escape" && !$("workItemModal").classList.contains("hidden")) closeWorkItem();
  });

  $("erpOrderList").addEventListener("click", function (event) {
    const node = event.target.closest("[data-job]");
    if (!node) return;
    selectedJob = node.dataset.job;
    selectedDetailTab = "总览";
    renderOrderCenter();
  });
  $("erpRecentOrders").addEventListener("click", function (event) {
    const node = event.target.closest("[data-job]");
    if (node) openOrder(node.dataset.job);
  });
  $("erpAttentionList").addEventListener("click", function (event) {
    const node = event.target.closest("[data-job]");
    if (node) openOrder(node.dataset.job);
  });
  $("erpOrderDetail").addEventListener("click", function (event) {
    const tab = event.target.closest("[data-detail-tab]");
    if (tab) {
      selectedDetailTab = tab.dataset.detailTab;
      renderOrderDetail(selectedJob);
      return;
    }
    const addWork = event.target.closest("[data-add-work]");
    if (addWork) {
      openWorkItem(addWork.dataset.addWork);
      return;
    }
    const button = event.target.closest("[data-erp-open]");
    if (!button) return;
    state.jobFilter = new Set([selectedJob]);
    window.switchTab(button.dataset.erpOpen);
    renderGrid();
  });

  $("erpGlobalSearch").addEventListener("input", function (event) {
    const box = $("erpSearchResults");
    const query = event.target.value.trim();
    if (!query) { box.classList.add("hidden"); return; }
    const rows = globalMatches(query);
    box.innerHTML = rows.length ? rows.map(function (x) {
      return '<div class="erp-search-item" data-job="' + esc(x.job) + '"><b>' + esc(x.job) + '</b> · ' +
        esc(x.customer) + '<small>' + esc(x.content) + ' · ' + x.device_count + '台 · ' + esc(x.stage) + '</small></div>';
    }).join("") : '<div class="erp-search-item">没有匹配结果</div>';
    box.classList.remove("hidden");
  });
  $("erpSearchResults").addEventListener("click", function (event) {
    const node = event.target.closest("[data-job]");
    if (!node) return;
    $("erpSearchResults").classList.add("hidden");
    $("erpGlobalSearch").value = "";
    openOrder(node.dataset.job);
  });
  document.addEventListener("mousedown", function (event) {
    if (!event.target.closest(".erp-global-search") && $("erpSearchResults")) $("erpSearchResults").classList.add("hidden");
  });

  ["btnQuery", "btnSettings"].forEach(function (id) {
    $(id).addEventListener("click", function () {
      ["工作台", "订单中心", "财务驾驶舱", "主数据"].forEach(function (view) {
        $("panel-" + view)?.classList.add("hidden");
      });
      document.querySelectorAll(".erp-nav-item, .erp-advanced button").forEach(function (b) { b.classList.remove("active"); });
      const target = id === "btnQuery" ? "跨表查询" : "设置";
      const nav = document.querySelector('[data-view="' + target + '"]');
      if (nav) nav.classList.add("active");
    });
  });

  window.addEventListener("pywebviewready", function () {
    setTimeout(function () {
      if (state.data) {
        window.switchTab("工作台");
        refreshErpWorkspace();
        loadMasterData(false);
      }
    }, 0);
  });
})();
