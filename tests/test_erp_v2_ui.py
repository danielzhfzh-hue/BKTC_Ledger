import pathlib
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]


class ErpV2UiTests(unittest.TestCase):
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

    def test_frontend_consumes_canonical_workspace_and_order_tabs(self):
        source = (ROOT / "ui" / "erp_v2.js").read_text(encoding="utf-8")
        for marker in (
            'call("get_erp_workspace", state.data, state.rules)',
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
