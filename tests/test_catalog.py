import unittest

import app


class MotorCatalogTests(unittest.TestCase):
    def test_all_source_files_loaded(self):
        self.assertFalse(app.SOURCE_ERRORS)
        self.assertEqual(len(app.RECORDS), 190)

    def test_family_class_power_dependencies(self):
        self.assertEqual(app.select_options({"family": "BAX", "efficiency": "IE2"})["powers"], app.select_options({"family": "BAX", "efficiency": "IE3"})["powers"])
        self.assertNotIn("0.12", app.select_options({"family": "BAX", "efficiency": "IE3"})["powers"])
        bax_small = app.select_options({"family": "BAX", "efficiency": "IE3", "power": "0.25"})
        bax_large = app.select_options({"family": "BAX", "efficiency": "IE3", "power": "3.7", "brakeType": "AC"})
        self.assertIn("71A", bax_small["frames"])
        self.assertIn("112M", bax_large["frames"])
        self.assertIn("B5", bax_large["record"]["values"]["Mounting"])
        self.assertIn("0.12", app.select_options({"family": "BMX", "efficiency": "IE3"})["powers"])
        self.assertIn("55", app.select_options({"family": "SMX", "efficiency": "IE2"})["powers"])

    def test_bax_brake_configuration_is_required(self):
        choices = app.select_options({"family": "BAX", "efficiency": "IE3", "power": "0.37"})
        self.assertEqual(choices["brakeTypes"], ["AC", "DC"])
        ac = app.select_options({"family": "BAX", "efficiency": "IE3", "power": "0.37", "brakeType": "AC"})
        dc = app.select_options({"family": "BAX", "efficiency": "IE3", "power": "0.37", "brakeType": "DC"})
        self.assertEqual(ac["record"]["values"]["Brake coil voltage"], "415")
        self.assertEqual(dc["record"]["values"]["Brake coil voltage"], "110")
        self.assertNotEqual(ac["record"]["values"]["Brake torque"], dc["record"]["values"]["Brake torque"])

    def test_brake_choice_flows_into_pdf(self):
        import pymupdf as fitz
        ac = app.select_options({"family": "BAX", "efficiency": "IE3", "power": "0.37", "brakeType": "AC"})["record"]
        dc = app.select_options({"family": "BAX", "efficiency": "IE3", "power": "0.37", "brakeType": "DC"})["record"]
        ac_text = fitz.open(stream=app.generate_pdf(ac), filetype="pdf")[0].get_text()
        dc_text = fitz.open(stream=app.generate_pdf(dc), filetype="pdf")[0].get_text()
        self.assertIn("415", ac_text)
        self.assertIn("14", ac_text)
        self.assertIn("110", dc_text)
        self.assertIn("9", dc_text)
        self.assertIn("(DC)", dc_text)
        dc_page = fitz.open(stream=app.generate_pdf(dc), filetype="pdf")[0]
        coil_values = [w[4] for w in dc_page.get_text("words") if 220 <= w[0] <= 260 and 565 <= w[1] <= 585]
        self.assertEqual(coil_values, ["110"])

    def test_workbook_conflicts_block_generation(self):
        bmx = app.select_options({"family": "BMX", "efficiency": "IE3", "power": "0.12"})
        self.assertIn("Type of terminal box", bmx["conflict"]["fields"])
        self.assertIsNone(bmx["record"])
        self.assertEqual(len(app.diagnostic_report()["conflicts"]), 2)

    def test_bax_ac_dc_choice_resolves_same_power_records(self):
        choices = app.select_options({"family": "BAX", "efficiency": "IE2", "power": "11"})
        self.assertEqual(choices["brakeTypes"], ["AC", "DC"])
        self.assertIsNone(choices["conflict"])
        for brake_type in ("AC", "DC"):
            selected = app.select_options({"family": "BAX", "efficiency": "IE2", "power": "11", "brakeType": brake_type})
            self.assertEqual(selected["recordCount"], 1)
            self.assertIsNone(selected["conflict"])
            self.assertEqual(selected["record"]["brakeType"], brake_type)
        ac = app.select_options({"family": "BAX", "efficiency": "IE2", "power": "11", "brakeType": "AC"})["record"]
        dc = app.select_options({"family": "BAX", "efficiency": "IE2", "power": "11", "brakeType": "DC"})["record"]
        import pymupdf as fitz
        ac_text=fitz.open(stream=app.generate_pdf(ac),filetype="pdf")[0].get_text()
        dc_text=fitz.open(stream=app.generate_pdf(dc),filetype="pdf")[0].get_text()
        self.assertIn("190", ac_text)
        self.assertIn("155", dc_text)

    def test_unavailable_selection_does_not_match(self):
        self.assertEqual(app.select_options({"family": "BAX", "efficiency": "IE3", "power": "999"})["recordCount"], 0)

    def test_generated_pdf_uses_selected_motor_values(self):
        record = app.select_options({"family": "BAX", "efficiency": "IE3", "power": "0.37", "brakeType": "AC"})["record"]
        import pymupdf as fitz
        doc = fitz.open(stream=app.generate_pdf(record), filetype="pdf")
        text = doc[0].get_text()
        for expected in ("0.37", "71B", "IE3", "1.8 / 1.1", "8.1", "9.1", "9.4", "14", "5.38", "DS-025-2028"):
            self.assertIn(expected, text)
        self.assertNotIn("3.7\nkW", text)

    def test_non_brake_motor_title_does_not_claim_a_brake(self):
        record = app.select_options({"family": "SMX", "efficiency": "IE2", "power": "55"})["record"]
        import pymupdf as fitz
        doc = fitz.open(stream=app.generate_pdf(record), filetype="pdf")
        self.assertIn("Data sheet - Three phase - Squirrel cage motors", doc[0].get_text())
        self.assertNotIn("Squirrel cage motors (Brake motor)", doc[0].get_text())


if __name__ == "__main__":
    unittest.main()
