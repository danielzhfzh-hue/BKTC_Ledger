import pathlib
import shutil
import subprocess
import unittest


class FrontendSaveTests(unittest.TestCase):
    @unittest.skipUnless(shutil.which("node"), "Node.js is required")
    def test_revision_only_save_refreshes_footer_without_replacing_data(self):
        source = (pathlib.Path(__file__).resolve().parents[1] / "ui/app.js").read_text()
        function = source[source.index("function applyBackendState(r)"):source.index("function saveSnapshot()")]
        script = r'''
const assert = require("node:assert/strict");
const data = {draft:"keep"};
const state = {data, schema:{}, databaseInfo:{revision:14}};
let displayedRevision;
const renderCounts = () => {displayedRevision = state.databaseInfo.revision;};
const recomputeDerived = () => {throw new Error("metadata update must not recompute drafts");};
''' + function + r'''
applyBackendState({revision:15});
assert.equal(displayedRevision,15);
assert.equal(state.data,data);
'''
        result = subprocess.run(["node", "-e", script], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    @unittest.skipUnless(shutil.which("node"), "Node.js is required")
    def test_unpaid_query_preserves_unknown_totals_and_exact_known_cents(self):
        source = (pathlib.Path(__file__).resolve().parents[1] / "ui/app.js").read_text()
        function = source[source.index("function queryRows(table)"):source.index("async function refreshUnpaidRows()")]
        script = r'''
const assert = require("node:assert/strict");
const s = value => String(value ?? "");
const _unpaidRows = [
  {"JOB No":"A", "客户":"GTX", "未回收金额":null},
  {"JOB No":"A", "客户":"GTX", "未回收金额":0.1},
  {"JOB No":"A", "客户":"GTX", "未回收金额":0.2},
  {"JOB No":"B", "客户":"GTX", "未回收金额":0},
];
''' + function + r'''
const rows = queryRows("未付款订单");
assert.equal(rows[0]["未回收合计"], null);
assert.equal(rows[0]["已知未回收合计"], 0.3);
assert.equal(rows[0]["金额待确认项数"], 1);
assert.equal(rows[1]["未回收合计"], 0);
assert.equal(rows[1]["金额待确认项数"], 0);
'''
        result = subprocess.run(["node", "-e", script], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    @unittest.skipUnless(shutil.which("node"), "Node.js is required")
    def test_workspace_ignores_old_response_after_edit(self):
        source = (pathlib.Path(__file__).resolve().parents[1] / "ui/erp_v2.js").read_text()
        functions = source[source.index("  async function refreshErpWorkspace()"):source.index("  function progress(")]
        script = r'''
const assert = require("node:assert/strict");
const state = {data: {value: "first"}, rules: {}, tab: "工作台"};
let workspaceRequest = 0, workspaceDirty = true, workspaceTimer = null, erpWorkspace = null;
const window = {}, document = {body: {}}, $ = () => null;
const clearTimeout = () => {}, setTimeout = () => 1, toast = () => {};
let renders = 0;
const renderDashboard = () => {renders++;}, renderOrderCenter = () => {}, renderFinance = () => {};
const pending = [], submitted = [];
const call = async (name, data) => {submitted.push(data); return await new Promise(resolve => pending.push(resolve));};
'''
        script += functions + r'''
(async () => {
  const first = refreshErpWorkspace();
  state.data.value = "second";
  window.invalidateErpWorkspace();
  const second = refreshErpWorkspace();
  assert.equal(submitted[0].value, "first");
  pending[1]({version: "new"}); await second;
  pending[0]({version: "old"}); await first;
  assert.equal(erpWorkspace.version, "new");
  assert.equal(renders, 1);
  assert.equal(workspaceDirty, false);
})().catch(error => {console.error(error); process.exitCode = 1;});
'''
        result = subprocess.run(["node", "-e", script], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    @unittest.skipUnless(shutil.which("node"), "Node.js is required")
    def test_finance_filter_recalculates_totals_and_buckets(self):
        source = (pathlib.Path(__file__).resolve().parents[1] / "ui/erp_v2.js").read_text()
        function = source[source.index("  function filteredFinance("):source.index("  function financeRowsForView(")]
        script = 'const assert = require("node:assert/strict"); const s = value => String(value || "");\n' + function + r'''
const rows = [{job: "A", customer: "GTX", amount: 0.1, currency: "RMB"},
              {job: "B", customer: "GTX", amount: 0.2, currency: "RMB"},
              {job: "C", customer: "Other", amount: 10, currency: "RMB"},
              {job: "D", customer: "GTX", amount: null, currency: "RMB"}];
const finance = {rows, aging: [{bucket: "待确认", rows}], forecast: [{bucket: "已逾期", rows}]};
const result = filteredFinance(finance, "gtx");
assert.equal(result.rows.length, 3);
assert.equal(result.open_amounts.RMB, 0.3);
assert.equal(result.overdue_amounts.RMB, 0.3);
assert.equal(result.aging[0].count, 3);
assert.equal(result.aging[0].amounts.RMB, 0.3);
assert.equal(finance.rows.length, 4);
assert.equal(filteredFinance(finance, "missing").rows.length, 0);
'''
        result = subprocess.run(["node", "-e", script], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    @unittest.skipUnless(shutil.which("node"), "Node.js is required")
    def test_edit_while_saving_is_retained(self):
        source = (pathlib.Path(__file__).resolve().parents[1] / "ui/app.js").read_text()
        functions = source[source.index("function applyBackendState(r)"):source.index('$("btnSave").addEventListener')]
        script = r'''
const assert = require("node:assert/strict");
const state = {data: {value: "first"}, rules: {}, dirty: true};
let finish, submitted;
const call = async (name, data) => {submitted = data; return await new Promise(resolve => finish = resolve);};
const pendingAuditContext = () => ({});
let cleared = false;
const clearPendingAuditActions = () => {cleared = true;};
const setStatus = () => {}, toast = () => {}, renderAll = () => {};
const switchTab = () => {}, renderSummary = () => {}, recomputeDerived = () => {};
'''
        script += functions + r'''
(async () => {
  const saving = saveCurrent(true);
  state.data.value = "second";
  assert.equal(submitted.value, "first", "request must own a snapshot");
  finish({data: {value: "first"}, revision: 2});
  await saving;
  assert.equal(state.data.value, "second");
  assert.equal(state.dirty, true);
  assert.equal(cleared, false);
  assert.equal(state.databaseInfo.revision, 2);
})().catch(error => {console.error(error); process.exitCode = 1;});
'''
        result = subprocess.run(["node", "-e", script], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
