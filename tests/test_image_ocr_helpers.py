import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import metrology_config_app_v2_3_pie_delete_process_guard as app


class ImageOcrHelperTests(unittest.TestCase):
    def test_invalid_image_config_json_reports_clear_error(self):
        with self.assertRaisesRegex(ValueError, "Invalid Image OCR config JSON"):
            app.parse_image_parse_config("{not-json", ["Rx"])

    def test_missing_roi_reports_clear_error(self):
        config = {"metrics": {"Rx": {"regex": r"Rx\s*=\s*(\d+)"}}}
        with self.assertRaisesRegex(ValueError, "missing roi"):
            app.parse_image_parse_config(json.dumps(config), ["Rx"])

    def test_missing_regex_reports_clear_error(self):
        config = {"metrics": {"Rx": {"roi": [0.1, 0.1, 0.2, 0.2]}}}
        with self.assertRaisesRegex(ValueError, "missing regex"):
            app.parse_image_parse_config(json.dumps(config), ["Rx"])

    def test_regex_value_extraction(self):
        value = app.extract_regex_value("Rx = -12.34 deg", r"Rx\s*=\s*([-+]?\d+(?:\.\d+)?)", "Rx")
        self.assertEqual(value, "-12.34")

    def test_directory_source_chooses_latest_supported_image(self):
        old_wait = app.FILE_STABLE_WAIT_SECONDS
        app.FILE_STABLE_WAIT_SECONDS = 0
        try:
            with tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                older = root / "older.png"
                newer = root / "newer.jpg"
                ignored = root / "ignored.txt"
                older.write_bytes(b"old")
                newer.write_bytes(b"new")
                ignored.write_bytes(b"text")
                os.utime(older, (1000, 1000))
                os.utime(newer, (2000, 2000))
                selected, data, _stat = app.find_stable_image_file(str(root), {"file_pattern": "*"})
                self.assertEqual(Path(selected).name, "newer.jpg")
                self.assertEqual(data, b"new")
        finally:
            app.FILE_STABLE_WAIT_SECONDS = old_wait

    def test_image_config_routes_by_filename_production_code(self):
        config = {
            "production_code_from_filename_regex": r"(?P<production_code>PROD_[A-Z]+)_",
            "metrics": {"Rx": {"roi": [0.1, 0.1, 0.2, 0.2], "regex": r"Rx=(\d+)"}},
        }
        parsed = app.parse_image_parse_config(json.dumps(config), ["Rx"])
        self.assertTrue(app.image_config_routes_by_production(parsed))

    def test_read_image_rows_scans_multiple_images_and_uses_filename_code_process(self):
        config = {
            "collect_mode": "all_stable",
            "production_code_from_filename_regex": r"(?P<production_code>PROD_[A-Z]+)_",
            "process_from_filename_regex": r"_(?P<process_step>STEP\d+)\.",
            "metrics": {"Rx": {"roi": [0.1, 0.1, 0.2, 0.2], "regex": r"Rx=(\d+)"}},
        }
        fake_images = [
            ("C:/images/PROD_A_STEP1.png", b"a", SimpleNamespace(st_mtime=1000)),
            ("C:/images/PROD_B_STEP2.png", b"b", SimpleNamespace(st_mtime=900)),
        ]
        ocr_results = [
            ({"Rx": "1.23"}, {"image_path": fake_images[0][0], "metrics": {}}, {}),
            ({"Rx": "4.56"}, {"image_path": fake_images[1][0], "metrics": {}}, {}),
        ]
        with mock.patch.object(app, "find_stable_image_files", return_value=fake_images), \
             mock.patch.object(app, "run_image_ocr_with_metadata", side_effect=ocr_results):
            fields, rows, label = app.read_image_rows(
                "C:/images", json.dumps(config), "ROUTER", "生产编号",
                "", "工序", ["Rx"]
            )
        self.assertIn("生产编号", fields)
        self.assertIn("工序", fields)
        self.assertEqual([r["生产编号"] for r in rows], ["PROD_A", "PROD_B"])
        self.assertEqual([r["工序"] for r in rows], ["STEP1", "STEP2"])
        self.assertEqual([r["Rx"] for r in rows], ["1.23", "4.56"])
        self.assertTrue(label.startswith("image:"))

    def test_read_image_rows_skips_bad_image_when_other_images_parse(self):
        config = {
            "collect_mode": "all_stable",
            "production_code_from_filename_regex": r"(?P<production_code>PROD_[A-Z]+)_",
            "metrics": {"Rx": {"roi": [0.1, 0.1, 0.2, 0.2], "regex": r"Rx=(\d+)"}},
        }
        fake_images = [
            ("C:/images/PROD_BAD_STEP1.png", b"bad", SimpleNamespace(st_mtime=1000)),
            ("C:/images/PROD_A_STEP2.png", b"good", SimpleNamespace(st_mtime=900)),
        ]
        ocr_results = [
            ValueError("Rx did not match"),
            ({"Rx": "1.23"}, {"image_path": fake_images[1][0], "metrics": {}}, {}),
        ]
        with mock.patch.object(app, "find_stable_image_files", return_value=fake_images), \
             mock.patch.object(app, "run_image_ocr_with_metadata", side_effect=ocr_results):
            fields, rows, _label = app.read_image_rows(
                "C:/images", json.dumps(config), "ROUTER", "生产编号",
                "AOI", "", ["Rx"]
            )
        self.assertIn("_image_parse_errors", fields)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["生产编号"], "PROD_A")
        self.assertEqual(rows[0]["Rx"], "1.23")
        self.assertEqual(rows[0]["_image_parse_errors"][0]["image_path"], fake_images[0][0])


class ImageOcrIntegrationTests(unittest.TestCase):
    @unittest.skipUnless(shutil.which(os.environ.get("MDCP_TESSERACT_CMD") or "tesseract"), "Tesseract is not installed")
    def test_synthetic_png_ocr_smoke(self):
        try:
            from PIL import Image, ImageDraw, ImageFont
            import pytesseract  # noqa: F401
            import cv2  # noqa: F401
            import numpy  # noqa: F401
        except ImportError as ex:
            self.skipTest(f"OCR dependency missing: {ex}")

        old_wait = app.FILE_STABLE_WAIT_SECONDS
        app.FILE_STABLE_WAIT_SECONDS = 0
        try:
            with tempfile.TemporaryDirectory() as tmp:
                image_path = Path(tmp) / "result_STEP01.png"
                img = Image.new("RGB", (900, 360), "white")
                draw = ImageDraw.Draw(img)
                try:
                    font = ImageFont.truetype("arial.ttf", 72)
                except Exception:
                    font = ImageFont.load_default()
                draw.text((30, 30), "Rx=1.23", fill="black", font=font)
                draw.text((30, 140), "Ry=4.56", fill="black", font=font)
                draw.text((30, 250), "Z=7.89", fill="black", font=font)
                img.save(image_path)

                config = {
                    "process_from_filename_regex": r"result_(?P<process_step>[^.]+)",
                    "ocr": {"lang": "eng", "psm": 7, "scale": 2.0, "threshold": True},
                    "metrics": {
                        "Rx": {"roi": [0.0, 0.00, 0.6, 0.28], "regex": r"Rx\s*[:=]?\s*([0-9.]+)"},
                        "Ry": {"roi": [0.0, 0.30, 0.6, 0.28], "regex": r"Ry\s*[:=]?\s*([0-9.]+)"},
                        "Z": {"roi": [0.0, 0.61, 0.6, 0.28], "regex": r"Z\s*[:=]?\s*([0-9.]+)"}
                    }
                }
                fields, rows, label = app.read_image_rows(
                    str(image_path), json.dumps(config), "PROD_A", "production_code",
                    "", "process_step", ["Rx", "Ry", "Z"]
                )
                self.assertIn("Rx", fields)
                self.assertEqual(rows[0]["production_code"], "PROD_A")
                self.assertEqual(rows[0]["process_step"], "STEP01")
                self.assertAlmostEqual(float(rows[0]["Rx"]), 1.23, places=2)
                self.assertAlmostEqual(float(rows[0]["Ry"]), 4.56, places=2)
                self.assertAlmostEqual(float(rows[0]["Z"]), 7.89, places=2)
                self.assertTrue(label.startswith("image:"))
        finally:
            app.FILE_STABLE_WAIT_SECONDS = old_wait


class ImageDynamicRoutingCollectTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old_db_file = app.DB_FILE
        app.DB_FILE = str(Path(self.tmp.name) / "image_dynamic_routing.db")
        app.init_db()
        self.conn = app.get_conn()
        self.cur = self.conn.cursor()
        self.router_id = self._insert_production("ROUTER")
        self.prod_a_id = self._insert_production("PROD_A")
        self.prod_b_id = self._insert_production("PROD_B")
        config = {
            "route_by_image_production_code": True,
            "production_code_from_filename_regex": r"(?P<production_code>PROD_[A-Z]+)_",
            "metrics": {"Rx": {"roi": [0.1, 0.1, 0.2, 0.2], "regex": r"Rx=(\d+)"}},
        }
        self.cur.execute(
            """
            INSERT INTO measurement_item_config (
                production_id, item_name, process_step, process_step_column, execution_time_text, equipment_name,
                data_source_type, data_source_path, image_parse_config_json, production_code_column,
                scan_frequency_seconds, enabled, updated_at
            ) VALUES (?, 'Shared camera', '', '工序', '', 'CAM-01', 'image', 'C:/images',
                      ?, '生产编号', 30, 1, ?)
            """,
            (self.router_id, json.dumps(config), app.now_str()),
        )
        self.item_id = self.cur.lastrowid
        self.cur.execute(
            """
            INSERT INTO metric_config (
                item_id, metric_name, source_column, data_type, unit, enabled, sort_order, created_at, updated_at
            ) VALUES (?, 'Rx', 'Rx', 'number', 'um', 1, 0, ?, ?)
            """,
            (self.item_id, app.now_str(), app.now_str()),
        )
        self.conn.commit()

    def tearDown(self):
        self.conn.close()
        app.DB_FILE = self.old_db_file
        self.tmp.cleanup()

    def _insert_production(self, code):
        self.cur.execute(
            "INSERT INTO production_config (production_code, production_name, updated_at) VALUES (?, ?, ?)",
            (code, code, app.now_str()),
        )
        return self.cur.lastrowid

    def test_collect_item_routes_image_rows_to_parsed_production_codes(self):
        rows = [
            {"生产编号": "PROD_A", "工序": "STEP1", "Rx": "1.23", "_source_path": "a.png"},
            {"生产编号": "PROD_B", "工序": "STEP2", "Rx": "4.56", "_source_path": "b.png"},
            {"生产编号": "PROD_X", "工序": "STEP3", "Rx": "7.89", "_source_path": "x.png"},
        ]
        with mock.patch.object(app, "read_source_rows", return_value=(["生产编号", "工序", "Rx"], rows, "image:mock")):
            result = app.collect_item(self.item_id)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["inserted"], 2)
        self.assertEqual(result["unknown_production_rows"], 1)
        saved = self.conn.execute(
            "SELECT production_code, process_step, metric_value_text FROM measurement_result ORDER BY production_code"
        ).fetchall()
        self.assertEqual(
            [(r["production_code"], r["process_step"], r["metric_value_text"]) for r in saved],
            [("PROD_A", "STEP1", "1.23"), ("PROD_B", "STEP2", "4.56")],
        )


if __name__ == "__main__":
    unittest.main()
