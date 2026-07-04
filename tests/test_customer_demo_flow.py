import tempfile
import unittest
from pathlib import Path

import metrology_data_platform_v2_7 as app


class CustomerDemoFlowTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old_db_file = app.DB_FILE
        self.old_demo_data_dir = app.DEMO_DATA_DIR
        self.old_file_stable_wait = app.FILE_STABLE_WAIT_SECONDS
        app.DB_FILE = str(Path(self.tmp.name) / "customer_demo.db")
        app.DEMO_DATA_DIR = Path(self.tmp.name) / "demo_data"
        app.FILE_STABLE_WAIT_SECONDS = 0
        app.init_db()
        self.admin = {"username": "admin", "role": "admin"}

    def tearDown(self):
        app.DB_FILE = self.old_db_file
        app.DEMO_DATA_DIR = self.old_demo_data_dir
        app.FILE_STABLE_WAIT_SECONDS = self.old_file_stable_wait
        self.tmp.cleanup()

    def test_seed_demo_data_creates_customer_visible_flow(self):
        result = app.seed_demo_data(self.admin, "127.0.0.1")

        self.assertEqual(result["total"], 24)
        self.assertEqual(result["ms3_count"], 13)
        self.assertEqual(result["ms2_count"], 8)
        self.assertEqual(result["miss_count"], 3)
        self.assertEqual(result["inserted"], 24)
        self.assertTrue(Path(result["csv_path"]).exists())

        conn = app.get_conn()
        productions = conn.execute("SELECT COUNT(*) AS c FROM production_config WHERE production_code LIKE 'DEMO_WAFER_%'").fetchone()["c"]
        items = conn.execute("SELECT COUNT(*) AS c FROM measurement_item_config WHERE item_name='Demo CSV 三工序量测'").fetchone()["c"]
        jobs = conn.execute("SELECT COUNT(*) AS c FROM collect_job WHERE trigger_type='demo_seed'").fetchone()["c"]
        conn.close()
        self.assertEqual(productions, 2)
        self.assertEqual(items, 2)
        self.assertEqual(jobs, 2)

        dashboard_html = app.page_dashboard(self.admin)
        self.assertIn("Demo 数据已就绪", dashboard_html)
        self.assertIn("查看采集结果", dashboard_html)

    def test_templates_page_renders_for_demo_template(self):
        app.seed_demo_data(self.admin, "127.0.0.1")

        html = app.page_templates(self.admin)
        self.assertIn("5分钟客户演示 CSV 模板", html)
        self.assertIn("上传/粘贴模板并生成字段映射", html)


if __name__ == "__main__":
    unittest.main()
