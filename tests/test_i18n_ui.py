import pathlib
import unittest


ROOT = pathlib.Path(__file__).resolve().parents[1]


class UiLanguageSwitchTests(unittest.TestCase):
    def test_language_switch_and_i18n_bundle_are_included(self):
        html = (ROOT / "ui" / "index.html").read_text(encoding="utf-8")
        self.assertIn('id="langSwitch"', html)
        self.assertIn('<option value="ja">日本語</option>', html)
        self.assertIn('<script src="i18n.js"></script>', html)
        self.assertLess(html.index('<script src="erp_v2.js"></script>'), html.index('<script src="i18n.js"></script>'))

    def test_japanese_dictionary_contains_core_erp_terms(self):
        source = (ROOT / "ui" / "i18n.js").read_text(encoding="utf-8")
        for expected in (
            '"订单中心": "受注センター"',
            '"工作台": "ダッシュボード"',
            '"付款条件": "支払条件"',
            '"开票记录": "請求記録"',
            '"回款记录": "入金記録"',
            '"已验收": "検収済み"',
            '"已回款": "入金済み"',
        ):
            self.assertIn(expected, source)

    def test_language_switch_is_display_only(self):
        source = (ROOT / "ui" / "i18n.js").read_text(encoding="utf-8")
        self.assertIn('localStorage.setItem(STORAGE_KEY, current)', source)
        self.assertNotIn('save_data', source)
        self.assertNotIn('database_path', source)


if __name__ == "__main__":
    unittest.main()
