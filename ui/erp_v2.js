"use strict";

(function () {
  const oldSwitchTab = window.switchTab;
  const oldRenderAll = window.renderAll;
  let selectedJob = "";

  function money(value, currency) {
    const n = Number(value) || 0;
    const symbol = currency === "USD" ? "$" : currency === "EUR" ? "€" : "¥";
    return symbol + n.toLocaleString(undefined, { maximumFractionDigits: 0 });
  }

  function pct(n, d) {
    return d > 0 ? Math.max(0, Math.min(100, Math.round(n / d * 100))) : 0;
  }

  function snapshots() {
    if (!state.data) return [];
    const devices = state.data["设备台账"] || [];
    const ships = state.data["发货批次"] || [];
    const invoices = state.data["开票记录"] || [];
    const payments = state.data["回款记录"] || [];
    const terms = state.data["付款条件"] || [];
    return (state.data["合同订单"] || []).map(function (order) {
      const job = s(order["JOB No"]);
      const ds = devices.filter(function (x) { return s(x["JOB No"]) === job; });
      const ss = ships.filter(function (x) { return s(x["JOB No"]) === job; });
      const ins = invoices.filter(function (x) { return s(x["JOB No"]) === job; });
      const pays = payments.filter(function (x) { return s(x["JOB No"]) === job; });
      const ts = terms.filter(function (x) { return s(x["JOB No"]) === job && s(x["款类"]); });
      const total = ds.reduce(function (a, x) { return a + (Number(x["未税单价"]) || 0); }, 0);
      const shipped = ds.filter(function (x) {
        const batch = s(x["发货批次"]);
        return batch && ss.some(function (y) { return s(y["发货批次"]) === batch && s(y["出荷日"]); });
      }).length;
      const accepted = ds.filter(function (x) {
        return s(x["验收状态"]) === "已验收" || !!s(x["质保开始日"]);
      }).length;
      const invoiceAmount = ins.reduce(function (a, x) { return a + (Number(x["含税金额"]) || 0); }, 0);
      const paidAmount = pays.reduce(function (a, x) { return a + (Number(x["含税金额"]) || 0); }, 0);
      const overdue = pays.some(function (x) { return s(x["是否超期"]) === "是"; }) ||
        ins.some(function (x) { return s(x["回款状态"]) === "超期未回"; });
      const pendingSerial = ds.filter(function (x) { return !s(x["製造番号"]); }).length;
      const shipPct = pct(shipped, ds.length);
      const acceptPct = pct(accepted, ds.length);
      const invoicePct = total > 0 ? Math.min(100, Math.round(invoiceAmount / (total * 1.13) * 100)) : (ins.length ? 100 : 0);
      const paidPct = invoiceAmount > 0 ? Math.min(100, Math.round(paidAmount / invoiceAmount * 100)) : 0;
      let stage = "已受注";
      let stageIndex = 0;
      if (ds.length && pendingSerial) { stage = "待生产"; stageIndex = 1; }
      if (shipped > 0 && shipped < ds.length) { stage = "部分出货"; stageIndex = 2; }
      if (ds.length && shipped === ds.length) { stage = "已出货"; stageIndex = 2; }
      if (accepted > 0 && accepted < ds.length) { stage = "部分验收"; stageIndex = 3; }
      if (ds.length && accepted === ds.length) { stage = "已验收"; stageIndex = 3; }
      if (invoiceAmount > 0) { stage = invoicePct >= 100 ? "已开票" : "部分开票"; stageIndex = 4; }
      if (paidAmount > 0) { stage = paidPct >= 99 ? "已结清" : "部分回款"; stageIndex = 5; }
      let risk = "正常";
      let riskLevel = "low";
      if (overdue) { risk = "逾期"; riskLevel = "high"; }
      else if (pendingSerial) { risk = pendingSerial + "台待编号"; riskLevel = "medium"; }
      else if (ins.some(function (x) { return ["未回款", "部分回款"].includes(s(x["回款状态"])); })) {
        risk = "应收跟进"; riskLevel = "medium";
      }
      return {
        order: order, job: job, customer: s(order["客户"]),
        content: s(order["订单内容"]) || s(order["设备型号"]),
        currency: s(order["币种"]) || "RMB",
        devices: ds, shipments: ss, invoices: ins, payments: pays, terms: ts,
        total: total, shipped: shipped, accepted: accepted,
        invoiceAmount: invoiceAmount, paidAmount: paidAmount,
        shipPct: shipPct, acceptPct: acceptPct, invoicePct: invoicePct, paidPct: paidPct,
        stage: stage, stageIndex: stageIndex, risk: risk, riskLevel: riskLevel,
        overdue: overdue, pendingSerial: pendingSerial
      };
    }).sort(function (a, b) { return b.job.localeCompare(a.job, "zh"); });
  }

  function progress(value) {
    return '<span class="erp-progress"><i><b style="width:' + value + '%"></b></i><span>' + value + '%</span></span>';
  }

  function renderRecent() {
    if (!$("erpRecentOrders") || !state.data) return;
    const q = s($("erpDashboardFilter") && $("erpDashboardFilter").value).toLowerCase();
    const rows = snapshots().filter(function (x) {
      return !q || [x.job, x.customer, x.content].join(" ").toLowerCase().includes(q);
    }).slice(0, 40);
    $("erpRecentOrders").innerHTML = rows.map(function (x) {
      return '<tr data-job="' + esc(x.job) + '"><td><b>' + esc(x.job) + '</b></td><td>' + esc(x.customer) +
        '</td><td>' + esc(x.content) + '</td><td>' + money(x.total, x.currency) +
        '</td><td><span class="erp-stage-pill">' + esc(x.stage) + '</span></td><td>' + x.shipped + '/' + x.devices.length +
        '</td><td>' + x.accepted + '/' + x.devices.length + '</td><td>' + progress(x.invoicePct) +
        '</td><td>' + progress(x.paidPct) + '</td><td><span class="risk ' + x.riskLevel + '">' + esc(x.risk) + '</span></td></tr>';
    }).join("");
  }

  function renderDashboard() {
    if (!$("erpKpis") || !state.data) return;
    const orders = snapshots();
    const overdue = orders.filter(function (x) { return x.overdue; });
    const receivable = orders.reduce(function (a, x) { return a + Math.max(0, x.invoiceAmount - x.paidAmount); }, 0);
    const pendingShip = orders.filter(function (x) { return x.devices.length && x.shipped < x.devices.length; });
    const pendingAccept = orders.reduce(function (a, x) { return a + Math.max(0, x.shipped - x.accepted); }, 0);
    const pendingSerial = orders.reduce(function (a, x) { return a + x.pendingSerial; }, 0);
    const issues = validateAll();
    $("erpToday").textContent = new Date().toLocaleDateString(window.I18N?.locale?.() || "zh-CN", { year: "numeric", month: "long", day: "numeric", weekday: "short" });
    const kpis = [
      ["订单", orders.length, "当前业务订单", ""],
      ["待发货", pendingShip.length, "尚未完成出货", "warn"],
      ["待验收", pendingAccept + " 台", "已发货未验收设备", "warn"],
      ["应收", money(receivable, "RMB"), "已开票尚未回款", ""],
      ["逾期", overdue.length + " 单", overdue.length ? money(overdue.reduce(function(a,x){return a+Math.max(0,x.invoiceAmount-x.paidAmount);},0),"RMB") : "无逾期", "danger"],
      ["数据待办", pendingSerial + issues.filter(function(x){return x.severity === "error";}).length, "待编号 + 校验错误", ""]
    ];
    $("erpKpis").innerHTML = kpis.map(function (k) {
      return '<div class="erp-kpi ' + k[3] + '"><div class="label">' + k[0] + '</div><div class="value">' + k[1] +
        '</div><div class="sub">' + k[2] + '</div></div>';
    }).join("");
    const attention = orders.filter(function (x) { return x.riskLevel !== "low"; }).slice(0, 12);
    $("erpAttentionList").innerHTML = attention.length ? attention.map(function (x) {
      return '<div class="erp-attention-row" data-job="' + esc(x.job) + '"><b>' + esc(x.job) + '</b><span>' +
        esc(x.customer) + ' · ' + esc(x.risk) + '</span><span class="risk ' + x.riskLevel + '">' +
        (x.riskLevel === "high" ? "高风险" : "需关注") + '</span></div>';
    }).join("") : '<div class="erp-empty-state"><b>暂无高优先级待办</b><span>当前订单状态正常。</span></div>';
    const names = ["已受注", "待生产", "履约/出货", "验收", "开票", "回款"];
    const counts = names.map(function (_, i) { return orders.filter(function (x) { return x.stageIndex === i; }).length; });
    const maxCount = Math.max.apply(null, [1].concat(counts));
    $("erpStageBoard").innerHTML = names.map(function (name, i) {
      return '<div class="erp-stage-row"><span>' + name + '</span><span class="erp-stage-bar"><i style="width:' +
        Math.round(counts[i] / maxCount * 100) + '%"></i></span><b>' + counts[i] + '</b></div>';
    }).join("");
    renderRecent();
  }

  function renderOrderDetail(job) {
    if (!$("erpOrderDetail") || !state.data) return;
    const x = snapshots().find(function (o) { return o.job === job; });
    if (!x) {
      $("erpOrderDetail").innerHTML = '<div class="erp-empty-state"><b>未找到订单</b></div>';
      return;
    }
    const life = ["受注", "生产", "出货", "验收", "开票", "回款"];
    const flow = [];
    flow.push('<div class="erp-flow-node"><b>订单 ' + esc(x.job) + '</b><small>' + esc(x.customer) + ' · ' + x.devices.length + '台</small></div>');
    x.shipments.forEach(function (sh) {
      flow.push('<div class="erp-flow-arrow">→</div>');
      flow.push('<div class="erp-flow-node"><b>发货 ' + esc(s(sh["发货批次"]) || "批次") + '</b><small>' +
        esc(s(sh["出荷日"]) || "日期未定") + ' · ' + (Number(sh["台数"]) || 0) + '台</small></div>');
    });
    if (x.invoices.length) {
      flow.push('<div class="erp-flow-arrow">→</div>');
      flow.push('<div class="erp-flow-node"><b>开票 ' + x.invoices.length + '笔</b><small>' + money(x.invoiceAmount, x.currency) + '</small></div>');
    }
    if (x.payments.length) {
      flow.push('<div class="erp-flow-arrow">→</div>');
      flow.push('<div class="erp-flow-node"><b>回款 ' + x.payments.length + '笔</b><small>' + money(x.paidAmount, x.currency) + '</small></div>');
    }
    $("erpOrderDetail").innerHTML =
      '<div class="erp-detail-head"><div class="erp-detail-title"><div><h2>' + esc(x.job) + '</h2><div class="erp-detail-sub">' +
      esc(x.customer) + ' · ' + esc(x.content || "未填写订单内容") + ' · ' + x.devices.length + ' 台</div></div><span class="risk ' +
      x.riskLevel + '">' + esc(x.risk) + '</span></div></div>' +
      '<div class="erp-lifecycle">' + life.map(function (name, i) {
        return '<div class="erp-life-step ' + (i < x.stageIndex ? "done" : i === x.stageIndex ? "current" : "") + '">' + name + '</div>';
      }).join("") + '</div>' +
      '<div class="erp-detail-metrics"><div class="erp-detail-metric"><span>订单未税金额</span><b>' + money(x.total, x.currency) +
      '</b></div><div class="erp-detail-metric"><span>发货进度</span><b>' + x.shipped + '/' + x.devices.length +
      '</b></div><div class="erp-detail-metric"><span>验收进度</span><b>' + x.accepted + '/' + x.devices.length +
      '</b></div><div class="erp-detail-metric"><span>已开票</span><b>' + money(x.invoiceAmount, x.currency) +
      '</b></div><div class="erp-detail-metric"><span>已回款</span><b>' + money(x.paidAmount, x.currency) + '</b></div></div>' +
      '<div class="erp-detail-section"><h3>Document Flow</h3><div class="erp-flow">' + flow.join("") + '</div></div>' +
      '<div class="erp-detail-section"><h3>快速进入业务明细</h3><div class="erp-detail-actions">' +
      '<button data-erp-open="合同订单">订单头</button><button data-erp-open="付款条件">付款条件 ' + x.terms.length +
      '</button><button data-erp-open="设备台账">设备 ' + x.devices.length + '</button><button data-erp-open="发货批次">发货 ' +
      x.shipments.length + '</button><button data-erp-open="开票记录">开票 ' + x.invoices.length +
      '</button><button data-erp-open="回款记录">回款 ' + x.payments.length + '</button></div></div>';
  }

  function renderOrderCenter() {
    if (!$("erpOrderList") || !state.data) return;
    const orders = snapshots();
    const stageSel = $("erpOrderStageFilter");
    const stages = Array.from(new Set(orders.map(function (x) { return x.stage; })));
    if (stageSel && stageSel.options.length <= 1) {
      stageSel.innerHTML = '<option value="">全部阶段</option>' + stages.map(function (x) { return '<option>' + esc(x) + '</option>'; }).join("");
    }
    const q = s($("erpOrderSearch") && $("erpOrderSearch").value).toLowerCase();
    const stageFilter = s(stageSel && stageSel.value);
    const visible = orders.filter(function (x) {
      return (!q || [x.job, x.customer, x.content].join(" ").toLowerCase().includes(q)) && (!stageFilter || x.stage === stageFilter);
    });
    if (!selectedJob && visible.length) selectedJob = visible[0].job;
    $("erpOrderList").innerHTML = visible.map(function (x) {
      return '<div class="erp-order-list-item ' + (x.job === selectedJob ? "active" : "") + '" data-job="' + esc(x.job) +
        '"><div class="top"><b>' + esc(x.job) + '</b><span class="erp-stage-pill">' + esc(x.stage) +
        '</span></div><div class="customer">' + esc(x.customer) + '</div><div class="meta"><span>' +
        esc(x.content || "—") + '</span><span>' + x.devices.length + '台 · ' + money(x.total, x.currency) + '</span></div></div>';
    }).join("");
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
      const deviceText = x.devices.map(function (d) {
        return s(d["PO No"]) + " " + s(d["製造番号"]) + " " + s(d["设备型号"]);
      }).join(" ");
      return [x.job, x.customer, x.content, deviceText].join(" ").toLowerCase().includes(q);
    }).slice(0, 12);
  }

  window.switchTab = function (tab) {
    oldSwitchTab(tab);
    if ($("panel-工作台")) $("panel-工作台").classList.toggle("hidden", tab !== "工作台");
    if ($("panel-订单中心")) $("panel-订单中心").classList.toggle("hidden", tab !== "订单中心");
    if (tab === "工作台" || tab === "订单中心") $("panel-表格").classList.add("hidden");
    document.querySelectorAll(".erp-nav-item, .erp-advanced button").forEach(function (b) {
      b.classList.toggle("active", b.dataset.view === tab);
    });
    if (tab === "工作台") renderDashboard();
    if (tab === "订单中心") renderOrderCenter();
  };

  window.renderAll = function () {
    oldRenderAll();
    renderDashboard();
    renderOrderCenter();
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

  $("erpOrderList").addEventListener("click", function (event) {
    const node = event.target.closest("[data-job]");
    if (!node) return;
    selectedJob = node.dataset.job;
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
        esc(x.customer) + '<small>' + esc(x.content) + ' · ' + x.devices.length + '台 · ' + esc(x.stage) + '</small></div>';
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
      $("panel-工作台").classList.add("hidden");
      $("panel-订单中心").classList.add("hidden");
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
        renderDashboard();
      }
    }, 0);
  });
})();
