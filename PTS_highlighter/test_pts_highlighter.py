"""Focused classification and workbook preservation checks."""

from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import openpyxl

from pts_highlighter import (
    COLOURS, ENGINEERING, RULES, classify, clean_date, process_workbook, select_sheet,
    sort_schedule,
    activity_code, validate_activity_ids,
    classification_formula, read_status_marks,
    STATUS_TITLE, LOOKUP_VERSION,
)


class ScheduleTests(unittest.TestCase):
    def test_shipment_is_not_needed(self):
        for family, discipline in (("Lot General Piping Hangers", "Hangers"),
                                   ("Lot General Piping Pipe Material", "Piping"),
                                   ("Lot General Piping Valves", "Valves")):
            with self.subTest(family=family), TemporaryDirectory() as directory:
                self.check_shipment_excluded(family, discipline, directory)

    def check_shipment_excluded(self, family, discipline, directory):
        result = classify(f"{family} - Shipment")
        self.assertFalse(result["disciplines"])
        self.assertFalse(result["ambiguous"])
        for milestone in ("Ready for Shipment", "DTS (T)", "DTS (S)"):
            self.assertEqual(classify(f"{family} - {milestone}")["disciplines"], (discipline,))
        source = Path(directory) / "schedule.xlsx"
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "Schedule"
        sheet.append(["Activity ID", "Activity Name", "Start", "Finish"])
        sheet.append(["SHIP", f"{family} - Shipment", None, datetime(2026, 10, 1)])
        sheet.append(["READY", f"{family} - Ready for Shipment", None, datetime(2026, 10, 2)])
        status = book.create_sheet("PTS Status")
        status["A1"], status["B2"], status["C2"] = STATUS_TITLE, sheet.title, LOOKUP_VERSION
        for column, name in enumerate(
                ("Activity Name", "Status", "Activity ID", "Piping", "Hangers", "Valves"), 1):
            status.cell(4, column, name)
        status["C5"] = "SHIP"
        status.cell(5, {"Piping": 4, "Hangers": 5, "Valves": 6}[discipline], "X")
        status.add_table(openpyxl.worksheet.table.Table(displayName="OldMarks", ref="A4:F5"))
        book.save(source)
        book.close()
        report = process_workbook(source)
        expected_counts = {"Piping": 0, "Hangers": 0, "Valves": 0, "ALL": 1}
        expected_counts[discipline] = 1
        self.assertEqual(report["counts"], expected_counts)
        self.assertEqual(report["excluded"], [])
        self.assertNotIn(f"{family} - Shipment", report["status"]["missing"])
        check = openpyxl.load_workbook(report["output"])
        self.assertNotIn("ship", read_status_marks(check, "Schedule"))
        self.assertTrue(check.active.row_dimensions[2].hidden)
        check.close()

    def test_activity_lookup_ids(self):
        validate_activity_ids(["ENG1", "ENG2", None, "", "001", 1])
        self.assertEqual(activity_code(123.0), "123")
        self.assertEqual(activity_code(None), "")
        for values in (["ENG1", "ENG1"], ["ENG1", "eng1"], [123, "123"]):
            with self.subTest(values=values):
                with self.assertRaisesRegex(ValueError, "Duplicate Activity ID"):
                    validate_activity_ids(values)

    def test_undated_headers_are_preserved_and_skipped(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "headers.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.title = "Schedule"
            sheet.append(["Activity ID", "Activity Name", "Start", "Finish"])
            activity = "Lot General Piping Engineering - Purchase Order"
            sheet.append(["SECTION", None, None, None])
            sheet.append(["SECTION", activity, "   ", None])
            sheet.append(["ENG1", activity, None, None])
            sheet.append(["ENG1", activity, "22.09.2026", None])
            sheet.append(["VAL1", "Lot General Piping Valves - Purchase Order",
                          None, datetime(2026, 10, 1)])
            book.save(source)
            report = process_workbook(source)
            self.assertTrue(report["validated"])
            self.assertEqual(report["skipped_undated_rows"], 3)
            self.assertEqual(report["counts"], {"Piping": 1, "Hangers": 1, "Valves": 1, "ALL": 2})
            check = openpyxl.load_workbook(report["output"])
            result = check.active
            preserved = [row for row in range(2, 7)
                         if result.cell(row, 3).value in (None, "   ")
                         and result.cell(row, 4).value is None]
            self.assertEqual(len(preserved), 3)
            for row in preserved:
                self.assertTrue(result.row_dimensions[row].hidden)
                for column in range(7, 11):
                    self.assertIsNone(result.cell(row, column).value)
            self.assertEqual(read_status_marks(check, "Schedule"),
                             {"eng1": ("Piping", "Hangers"), "val1": ("Valves",)})
            check.close()
            repeated = process_workbook(Path(report["output"]))
            self.assertEqual(repeated["counts"], report["counts"])
            sheet.append(["eng1", activity, None, datetime(2026, 11, 1)])
            book.save(source)
            book.close()
            with self.assertRaisesRegex(ValueError, "Duplicate Activity ID: eng1"):
                process_workbook(source, overwrite=True)

    def test_budapest_purchase_orders(self):
        engineering = classify("Lot General Piping Engineering - Purchase Order")
        self.assertEqual(engineering["milestone"], "Purchase Order")
        self.assertEqual(engineering["disciplines"], ("Piping", "Hangers"))
        hangers = classify("Lot General Piping Hangers - Purchase Order - Base on BOQ1")
        self.assertEqual(hangers["milestone"], "Purchase Order")
        self.assertEqual(hangers["disciplines"], ("Hangers",))
        for name in (
            "Lot General Piping Engineering - Purchase Order Cancelled",
            "Lot General Piping Hangers - Purchase Order - Base on BOQ1 - Cancelled",
            "Lot General Piping Hangers - PO adjustment base on BOQ2",
            "Lot General Piping Valves - Purchase Order - Base on BOQ1",
        ):
            with self.subTest(name=name):
                self.assertFalse(classify(name)["disciplines"])

    def test_editable_table_validation(self):
        book = openpyxl.Workbook()
        status = book.active
        status["A1"], status["B2"], status["C2"] = STATUS_TITLE, "Schedule", LOOKUP_VERSION
        for column, name in enumerate(("Activity ID", "Piping", "Hangers", "Valves", "ALL"), 1):
            status.cell(4, column, name)
        status["A5"], status["B5"] = "ID1", "X"
        status.add_table(openpyxl.worksheet.table.Table(displayName="Marks", ref="A4:E6"))
        self.assertEqual(read_status_marks(book, "Schedule"), {"id1": ("Piping",)})
        status["A6"] = "id1"
        with self.assertRaisesRegex(ValueError, "Duplicate Activity ID in PTS Status"):
            read_status_marks(book, "Schedule")
        status["A6"] = None
        status["B5"] = "Yes"
        with self.assertRaisesRegex(ValueError, "Invalid Piping mark"):
            read_status_marks(book, "Schedule")
        status["B5"] = None
        self.assertEqual(read_status_marks(book, "Schedule"), {"id1": ()})
        book.close()

    def test_every_approved_rule(self):
        for family, (milestones, disciplines) in RULES.items():
            for milestone in milestones:
                with self.subTest(family=family, milestone=milestone):
                    self.assertEqual(classify(f"{family} - {milestone}")["disciplines"], disciplines)

    def test_guardrails(self):
        for name in (
            "Lot General Piping Pipe Material",
            "Lot General Piping Pipe Material - Engineering & Manufacturing (by Supplier)",
            "Lot General Piping Pipe Material - TSD Documents complete / RFQ",
            "Lot General Piping Engineering - Detail Engineering - ISO and BOM for Erection - Issued",
            "Lot 5 - Erection General Piping - Contractor On Site Mobilisation",
            "Lot 6 - Erection General Piping - BOQ 1 (Doc's for RfQ) - Issued from Engineering",
            "Other Lot - Purchase Order",
            "Lot General Piping Valves - Purchase Order Cancelled",
        ):
            with self.subTest(name=name):
                self.assertFalse(classify(name)["disciplines"])

    def test_normalization_and_small_typo(self):
        result = classify("LOT General Piping Pipe Material: Tender Due Daate")
        self.assertEqual(result["disciplines"], ("Piping",))
        result = classify("Lot General Piping Pipe Material - Technical & Commercial Comparison Pre TCC")
        self.assertEqual(result["milestone"], "Technical & Commercial Comparison Pre-TCC")
        self.assertEqual(classify("Lot General Piping Hangers - DTS (S)")["milestone"], "DTS (S)")

    def test_dates(self):
        for text in ("22.09.2026 A", "*22.09.2026", "22/09/2026*", " 22-09-2026 "):
            self.assertEqual(clean_date(text), (datetime(2026, 9, 22), True))
        for text in ("31.02.2026 A", "not a date", "01.02.2026 / 02.02.2026"):
            self.assertEqual(clean_date(text), (text, False))
        self.assertEqual(clean_date(None), (None, True))

    def test_latest_visible(self):
        book = openpyxl.Workbook()
        book.active.title = "DD25Jul2026 Update"
        book.create_sheet("DD25Sep2026 Update")
        book.create_sheet("DD25Dec2026 Update").sheet_state = "hidden"
        self.assertEqual(select_sheet(book).title, "DD25Sep2026 Update")
        book.create_sheet("Unknown")
        with self.assertRaises(ValueError):
            select_sheet(book)

    def test_processing_and_repeat_run(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "schedule.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.title = "DD25Sep2026 Update"
            sheet.append(["Schedule"])
            sheet.append(["Activity ID", "Activity Name", "BL Project Start", "Start", "Finish", "Piping ", "Hanger", "All"])
            sheet.append(["ENG1", "Lot General Piping Engineering - " + ENGINEERING[0],
                          "baseline unchanged", "*22.09.2026 A", None, "x", None, '=IF(F3="x","X","")'])
            sheet.append(["BAD1", "Lot General Piping Pipe Material - Engineering & Manufacturing (by Supplier)",
                          None, "bad date", "23/09/2026*", "x"])
            sheet.append(["VAL1", "Lot General Piping Valves - Purchase Order",
                          None, "01/10/2026", "02/10/2026"])
            sheet.cell(3, 16373).fill = openpyxl.styles.PatternFill("solid", fgColor="FFEEEEEE")
            archive = book.create_sheet("Archive")
            archive.sheet_state = "hidden"
            archive["A1"] = "Preserved"
            book.save(source)
            report = process_workbook(source)
            self.assertTrue(report["validated"])
            self.assertEqual(report["counts"], {"Piping": 1, "Hangers": 1, "Valves": 1, "ALL": 2})
            self.assertEqual(len(report["unresolved_dates"]), 1)
            output = Path(report["output"])
            check = openpyxl.load_workbook(output)
            result = check.active
            self.assertEqual([result.cell(2, column).value for column in range(6, 12)],
                             ["Proposed date LS", "Remarks LS", "Piping", "Hangers", "Valves", "ALL"])
            for row in range(3, 6):
                self.assertIsNone(result.cell(row, 6).value)
                self.assertIsNone(result.cell(row, 7).value)
                self.assertEqual(result.cell(row, 6).number_format, "DD.MM.YYYY")
                self.assertEqual(result.cell(row, 7).number_format, "@")
                self.assertTrue(result.cell(row, 7).alignment.wrap_text)
            self.assertEqual(result.auto_filter.filterColumn[0].colId, 10)
            self.assertEqual([result.cell(row, 1).value for row in range(3, 6)],
                             ["BAD1", "VAL1", "ENG1"])
            self.assertEqual(result["C5"].value, "baseline unchanged")
            self.assertEqual(result["D5"].value, datetime(2026, 9, 22))
            self.assertEqual(result["H5"].value,
                             classification_formula(5, 1, report["status"]["table"], "Piping"))
            colour_rules = [rule for rules in result.conditional_formatting._cf_rules.values()
                            for rule in rules]
            self.assertEqual(len(colour_rules), 3)
            self.assertIn(COLOURS["Lot General Piping Engineering"],
                          [rule.dxf.fill.fgColor.rgb for rule in colour_rules])
            self.assertEqual(check["Archive"].sheet_state, "hidden")
            self.assertEqual(result.cell(5, 16376).fill.fgColor.rgb, "FFEEEEEE")
            status = check[report["status"]["worksheet"]]
            self.assertEqual(status["A4"].value, "NOT FOUND (27)")
            missing = report["status"]["missing"]
            self.assertIn("Lot General Piping Pipe Material - DTS (S)", missing)
            self.assertEqual([status.cell(row, 1).value for row in range(6, 6 + len(missing))], missing)
            self.assertEqual(report["status"]["found"], 2)
            calculator = report["calculator"]
            self.assertEqual(calculator["first_date"], "E9")
            self.assertEqual(calculator["second_date"], "E10")
            self.assertEqual(result["E11"].value, '=IF(COUNT(E9:E10)=2,(E10-E9)/7,"")')
            self.assertEqual(result["E9"].number_format, "DD.MM.YYYY")
            self.assertGreater(result["E9"].row, 5)
            self.assertFalse(result.row_dimensions[11].hidden)
            self.assertTrue(check.calculation.fullCalcOnLoad)
            self.assertEqual(check.calculation.calcMode, "auto")
            self.assertEqual(str(result.data_validations.dataValidation[0].sqref), "E9:E10")
            shift = calculator["date_shift"]
            self.assertEqual(shift["date"], "E15")
            self.assertEqual(shift["weeks"], "E16")
            self.assertEqual(shift["result"], "E17")
            self.assertIn('IF(ISNUMBER(E15),INT(E15),', result["E17"].value)
            self.assertIn('DATE(VALUE(RIGHT(', result["E17"].value)
            self.assertIn('"dd.mm.yyyy"', result["E17"].value)
            self.assertIn('"Invalid date"', result["E17"].value)
            self.assertEqual(result["E15"].number_format, "@")
            self.assertEqual(result["E17"].number_format, "DD.MM.YYYY")
            self.assertEqual(result.data_validations.dataValidation[1].formula1, "ISNUMBER(E16)")
            self.assertEqual(str(result.data_validations.dataValidation[1].sqref), "E16")
            self.assertEqual(result.data_validations.dataValidation[2].type, "custom")
            self.assertIn('ISTEXT(E15)', result.data_validations.dataValidation[2].formula1)
            self.assertEqual(str(result.data_validations.dataValidation[2].sqref), "E15")
            result["F5"] = datetime(2026, 11, 3)
            result["G5"] = "Proposed sequence\nAwaiting confirmation"
            check.save(output)
            check.close()
            with self.assertRaises(FileExistsError):
                process_workbook(source)
            second = process_workbook(output)
            self.assertEqual(second["counts"], report["counts"])
            self.assertEqual(second["status"], report["status"])
            self.assertEqual(second["calculator"], report["calculator"])
            repeated = openpyxl.load_workbook(second["output"])
            self.assertEqual(repeated.active["F5"].value, datetime(2026, 11, 3))
            self.assertEqual(repeated.active["G5"].value, "Proposed sequence\nAwaiting confirmation")
            self.assertEqual([repeated.active.cell(2, column).value for column in range(6, 12)],
                             ["Proposed date LS", "Remarks LS", "Piping", "Hangers", "Valves", "ALL"])
            repeated.close()

    def test_budapest_processing_and_exclusion_review(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "budapest.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.append(["Activity ID", "Activity Name", "Start", "Finish"])
            names = (
                ("ENGPO", "Lot General Piping Engineering - Purchase Order", 2),
                ("HANGPO", "Lot General Piping Hangers - Purchase Order - Base on BOQ1", 1),
                ("TENDER", "Lot General Piping Engineering - Tender Due Date", 3),
                ("OTHER", "Lot General Piping Additional Material - Purchase Order", 4),
                ("CANCEL", "Lot General Piping Engineering - Purchase Order Cancelled", 5),
                ("BOQ1", "Lot General Piping Engineering - BOM / BOQ-1 - Issue for RfQ GP - Piping Material, Valves, Hangers", 6),
                ("UNRELATED", "Other Lot - Purchase Order", 7),
                ("HEADING", "Lot General Piping Engineering", 8),
            )
            for identity, name, day in names:
                sheet.append([identity, name, None, datetime(2026, 10, day)])
            book.save(source)
            book.close()
            report = process_workbook(source)
            self.assertTrue(report["validated"])
            self.assertEqual(report["counts"], {"Piping": 1, "Hangers": 2, "Valves": 0, "ALL": 2})
            self.assertNotIn("Lot General Piping Hangers - Purchase Order", report["missing_milestones"])
            excluded = {item["activity_id"]: item for item in report["excluded"]}
            self.assertEqual(set(excluded), {"CANCEL", "BOQ1"})
            self.assertEqual(excluded["BOQ1"]["reason"], "Ambiguous milestone")
            self.assertEqual(excluded["BOQ1"]["source_row"], 7)
            check = openpyxl.load_workbook(report["output"])
            result = check.active
            self.assertEqual([result.cell(row, 1).value for row in (2, 3)], ["HANGPO", "ENGPO"])
            status = check[report["status"]["worksheet"]]
            first = report["status"]["header_row"] + 1
            last = report["status"]["last_row"]
            lookup_rows = {status.cell(row, 3).value: row for row in range(first, last + 1)}
            self.assertEqual(set(lookup_rows), {"ENGPO", "HANGPO", "CANCEL", "BOQ1"})
            self.assertEqual([status.cell(first - 1, column).value for column in range(1, 7)],
                             ["Activity Name", "Status", "Activity ID", "Piping", "Hangers", "Valves"])
            self.assertEqual(status.max_column, 6)
            self.assertIsNone(status.freeze_panes)
            self.assertIn(status.cell(first, 3).value, {"CANCEL", "BOQ1"})
            self.assertLess(status.row_dimensions[lookup_rows["ENGPO"]].height, 30)
            for item in report["excluded"]:
                offset = lookup_rows[item["activity_id"]]
                self.assertEqual(status.cell(offset, 1).value, item["activity_name"])
                self.assertIn('"NEEDS YOUR REVIEW"', status.cell(offset, 2).value)
                self.assertEqual(status.cell(offset, 3).hyperlink.target,
                                 f"#'Sheet'!A{item['row']}")
                self.assertEqual(result.cell(item["row"], 1).value, item["activity_id"])
                self.assertTrue(result.row_dimensions[item["row"]].hidden)
                self.assertIn('FALSE)', result.cell(item["row"], 7).value)
                self.assertIn('IFERROR', result.cell(item["row"], 7).value)
            review_row = lookup_rows["BOQ1"]
            status.cell(review_row, 4, "X")
            status.cell(review_row, 6, "X")
            status.cell(lookup_rows["ENGPO"], 4).value = None
            self.assertEqual(status.cell(review_row, 2).value,
                             f'=IF(COUNTIF(D{review_row}:F{review_row},"X")>0,"Manually approved","NEEDS YOUR REVIEW")')
            check.save(report["output"])
            check.close()
            regenerated = process_workbook(source, overwrite=True)
            self.assertEqual(regenerated["counts"], {"Piping": 1, "Hangers": 2, "Valves": 1, "ALL": 3})
            self.assertEqual([item["activity_id"] for item in regenerated["excluded"]], ["CANCEL"])
            accepted = openpyxl.load_workbook(regenerated["output"])
            marks = read_status_marks(accepted, "Sheet")
            self.assertEqual(marks["boq1"], ("Piping", "Valves"))
            self.assertEqual(marks["engpo"], ("Hangers",))
            self.assertFalse(accepted.active.row_dimensions[7].hidden)
            accepted.close()
            repeated = process_workbook(Path(regenerated["output"]))
            self.assertEqual(repeated["counts"], regenerated["counts"])

    def test_finish_sort_with_name_ties_and_undated_rows(self):
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.append(["Activity ID", "Activity Name", "Finish"])
        sheet.append(["LATE", "Alpha", datetime(2026, 10, 2)])
        sheet.append(["ZULU", "Zulu", datetime(2026, 10, 1)])
        sheet.append(["EMPTY", "Undated B", None])
        sheet.append(["ALPHA1", "alpha", datetime(2026, 10, 1)])
        sheet.append(["INVALID", "Undated A", "bad date"])
        sheet.append(["ALPHA2", "alpha", datetime(2026, 10, 1)])
        sheet.row_dimensions[5].height = 42
        sheet.row_dimensions[5].hidden = True
        sheet["A5"].hyperlink = "https://example.com"
        mapping = sort_schedule(sheet, 1, 7, {"Activity Name": 2, "Finish": 3})
        self.assertEqual([sheet.cell(row, 1).value for row in range(2, 8)],
                         ["ALPHA1", "ALPHA2", "ZULU", "LATE", "INVALID", "EMPTY"])
        self.assertEqual(mapping[5], 2)
        self.assertEqual(sheet.row_dimensions[2].height, 42)
        self.assertTrue(sheet.row_dimensions[2].hidden)
        self.assertEqual(sheet["A2"].hyperlink.ref, "A2")
        self.assertEqual(sheet["A1"].value, "Activity ID")


if __name__ == "__main__":
    unittest.main()