import pathlib
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]


class ErpV2UiTests(unittest.TestCase):
    def test_dashboard_does_not_silently_drop_orders_after_first_forty(self):
        source = (ROOT / "ui" / "erp_v2.js").read_text(encoding="utf-8")
        recent = source[source.index("  function renderRecent()"):source.index("  function renderWorkList()")]
        self.assertNotIn(".slice(", recent)
        self.assertIn('$("erpRecentCount").textContent', recent)

    def test_cross_table_query_accepts_custom_values_and_shows_all_rows(self):
        source = (ROOT / "ui" / "app.js").read_text(encoding="utf-8")
        editor = source[source.index("function qbAddFilterRow()"):source.index("async function qbDoExport()")]
        self.assertIn('<input class="qb-f-val"', editor)
        self.assertIn('sel.setAttribute("list", suggestions.id)', editor)
        self.assertIn('const body = rows.map(', editor)
        self.assertNotIn('rows.slice(0, 10)', editor)

    def test_erp_modal_backdrop_reuses_centered_overlay_styles(self):
        css = (ROOT / "ui" / "style.css").read_text(encoding="utf-8")
        self.assertIn('.modal, .modal-backdrop { position: fixed; inset: 0;', css)

    def test_main_content_remains_scrollable_when_viewport_is_short(self):
        css = (ROOT / "ui" / "style.css").read_text(encoding="utf-8")
        self.assertRegex(css, r"main\s*\{[^}]*overflow:\s*auto")

    def test_editable_excel_preview_and_second_confirmation_are_wired(self):
        html = (ROOT / "ui" / "index.html").read_text(encoding="utf-8")
        source = (ROOT / "ui" / "app.js").read_text(encoding="utf-8")
        for marker in ('id="btnEditableExport"', 'id="btnEditableImport"'):
            self.assertIn(marker, html)
        for marker in ('call("preview_editable_import", path)',
                       'call("confirm_editable_import", preview.token, true',
                       'preview.changes || []', '第二次确认：',
                       'document.body.inert = true', '预览期间有新的修改'):
            self.assertIn(marker, source)
        for marker in ('选择文件期间数据已修改或数据库已切换',
                       '导出期间的新修改未包含在此工作簿中',
                       'state.storePath !== databasePath'):
            self.assertIn(marker, source)

    def test_master_editor_has_atomic_cancel_and_form_instead_of_prompts(self):
        source = (ROOT / "ui" / "erp_v2.js").read_text(encoding="utf-8")
        editor = source[source.index("  async function saveMaster("):source.index("  function openWorkItem(")]
        self.assertNotIn("prompt(", editor)
        self.assertNotIn("confirm(", editor)
        self.assertIn('overlay.setAttribute("aria-modal", "true")', editor)
        self.assertIn('button.addEventListener("click", () => close(null))', editor)
        self.assertIn('if (values === null) return;', editor)
        self.assertIn('if (saving) return;', editor)
        self.assertLess(editor.index('await persist(values);'), editor.index('close(values);'))
        self.assertIn('errorMessage.textContent = String(error);', editor)
        self.assertIn('Number.isFinite(ratio)', editor)
        self.assertIn('Number.isInteger(days)', editor)
        self.assertIn('overlay.querySelectorAll("button, input, textarea")', editor)
        self.assertIn('overlay.setAttribute("aria-busy", "true")', editor)
        self.assertIn('!overlay.contains(document.activeElement)', editor)
        self.assertIn('previousFocus?.focus()', editor)

    def test_v14_panels_and_work_item_modal_are_present(self):
        html = (ROOT / "ui" / "index.html").read_text(encoding="utf-8")
        for marker in (
            'id="panel-财务驾驶舱"',
            'id="panel-主数据"',
            'id="erpWorkList"',
            'id="workItemModal"',
            'id="erpContactMaster"',
            'id="nocMasterTermTemplate"',
        ):
            self.assertIn(marker, html)

    def test_summary_panel_id_is_unique_and_dashboard_sections_are_siblings(self):
        html = (ROOT / "ui" / "index.html").read_text(encoding="utf-8")
        self.assertEqual(html.count('id="panel-摘要"'), 1)
        self.assertEqual(html.count('class="erp-card erp-work-card"'), 1)
        self.assertEqual(html.count('class="erp-card erp-recent"'), 1)
        work_end = html.index('</section>', html.index('class="erp-card erp-work-card"'))
        recent_start = html.index('class="erp-card erp-recent"')
        self.assertLess(work_end, recent_start)

    def test_frontend_consumes_canonical_workspace_and_order_tabs(self):
        source = (ROOT / "ui" / "erp_v2.js").read_text(encoding="utf-8")
        for marker in (
            'call("get_erp_workspace", snapshot.data, snapshot.rules)',
            'const DETAIL_TABS = ["总览", "设备", "发货", "开票", "回款", "付款条件", "审计"]',
            'function renderFinance()',
            'function renderMasterData()',
            'data-work-postpone=',
            'call("complete_work_item"',
        ):
            self.assertIn(marker, source)

    def test_quote_conversion_and_master_template_are_wired(self):
        source = (ROOT / "ui" / "app.js").read_text(encoding="utf-8")
        for marker in (
            '"来源报价ID": f.sourceQuoteId',
            '"来源报价单号": f.sourceQuoteNo',
            'call("list_master_data")',
            'fillNocMasterTemplate',
        ):
            self.assertIn(marker, source)

    def test_new_ui_has_japanese_terms(self):
        source = (ROOT / "ui" / "i18n.js").read_text(encoding="utf-8")
        for marker in (
            '"财务驾驶舱": "財務コックピット"',
            '"主数据": "マスターデータ"',
            '"我的待办": "マイタスク"',
            '"客户主数据": "顧客マスター"',
            '"延期": "延期"',
        ):
            self.assertIn(marker, source)


    def test_v2_workbench_is_default_landing_page(self):
        source = (ROOT / "ui" / "app.js").read_text(encoding="utf-8")
        self.assertGreaterEqual(source.count('switchTab("工作台")'), 2)
        self.assertIn('switchTab("摘要")', source)  # 保存失败时仍回到校验摘要


if __name__ == "__main__":
    unittest.main()
