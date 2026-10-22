"""第 7 回の指摘: 素の文字のインライン記法のエスケープ、-r で重なるファイル名の location、--max-* の範囲."""

from __future__ import annotations

import contextlib
import io
import json
import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl
from devtools_common.cli import run_main

from doc_to_markdown import cli, lossaudit
from doc_to_markdown.model import Heading, ListItem, Paragraph, Table
from doc_to_markdown.render import render

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"


def run_cli(*args: str) -> tuple[int, str, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = run_main(cli.main, list(args))
    return rc, out.getvalue(), err.getvalue()


class PlainTextInlineEscapeTest(unittest.TestCase):
    def test_plain_table_cell_does_not_become_emphasis_or_link(self):
        md = render([Table(rows=[["式", "リンク"], ["2*3*4", "[x](javascript:alert(1))"], ["a_b_", "`code` ~~s~~"]])])
        self.assertIn("2\\*3\\*4", md)
        self.assertIn("\\[x\\](javascript:alert(1))", md)
        self.assertIn("a\\_b\\_", md)
        self.assertIn("\\`code\\` \\~\\~s\\~\\~", md)

    def test_plain_paragraph_heading_and_list(self):
        md = render([Heading(1, "*重要*"), Paragraph("[a](b) と **c**"), ListItem("_x_", 0)])
        self.assertIn("# \\*重要\\*", md)
        self.assertIn("\\[a\\](b) と \\*\\*c\\*\\*", md)
        self.assertIn("- \\_x\\_", md)

    def test_markup_text_keeps_intended_markup(self):
        md = render([Paragraph("**太字** と [リンク](https://example.com)", markup=True)])
        self.assertIn("**太字** と [リンク](https://example.com)", md)

    def test_escaped_symbols_count_as_one_char_in_loss_audit(self):
        # ロス検知の本文は、エスケープの \ を除いた元の文字に戻る（文字数が増えて見えない・減って見えない）
        text = "2*3*4 [x] a_b `c` ~d~"
        self.assertEqual(lossaudit.plain_text(render([Paragraph(text)])).strip(), text)


class XlsxCellEscapeEndToEndTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)

    def test_xlsx_cells_with_markdown_symbols(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "計算"
        ws.append(["項目", "式"])
        ws.append(["体積", "2*3*4"])
        ws.append(["参照", "[x](javascript:alert(1))"])
        src = self.tmp / "calc.xlsx"
        wb.save(src)
        rc, out, _ = run_cli(str(src), "-d", str(self.tmp / "out"), "-f", "json")
        self.assertEqual(rc, 0, out)
        md = (self.tmp / "out" / "calc.md").read_text(encoding="utf-8")
        self.assertIn("2\\*3\\*4", md)
        self.assertIn("\\[x\\](javascript:alert(1))", md)
        rules = {f["rule_id"] for f in json.loads(out)["findings"]}
        self.assertNotIn("loss.text", rules)
        self.assertNotIn("loss.raw_text", rules)


class DuplicateNameLocationTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)

    def test_recursive_failures_with_same_name_are_distinguishable(self):
        for sub in ("a", "b"):
            d = self.tmp / "in" / sub
            d.mkdir(parents=True)
            (d / "spec.docx").write_bytes(b"not a zip")
        shutil.copy(SAMPLES / "expense-list.xlsx", self.tmp / "in" / "a")
        out_dir = str(self.tmp / "out")
        rc, out, _ = run_cli(str(self.tmp / "in"), "-r", "-d", out_dir, "--continue-on-error", "-f", "json")
        self.assertEqual(rc, 1)
        errors = [f for f in json.loads(out)["findings"] if f["severity"] == "error"]
        self.assertEqual(sorted(f["location"] for f in errors), ["a/spec.docx", "b/spec.docx"])
        ok = [f for f in json.loads(out)["findings"] if f["rule_id"] == "convert.ok"]
        self.assertEqual([f["location"] for f in ok], ["expense-list.xlsx"])  # 重ならない名前は今までどおり

    def test_converter_findings_of_same_name_get_folder(self):
        for sub in ("a", "b"):
            d = self.tmp / "in" / sub
            d.mkdir(parents=True)
            shutil.copy(SAMPLES / "expense-list.xlsx", d / "list.xlsx")
        rc, out, _ = run_cli(str(self.tmp / "in"), "-r", "-d", str(self.tmp / "out"), "-f", "json")
        self.assertEqual(rc, 0)
        locs = {f["location"] for f in json.loads(out)["findings"] if f["location"]}
        self.assertEqual(locs, {"a/list.xlsx", "b/list.xlsx"})

    def test_dry_run_uses_the_same_locations(self):
        for sub in ("a", "b"):
            d = self.tmp / "in" / sub
            d.mkdir(parents=True)
            shutil.copy(SAMPLES / "expense-list.xlsx", d / "list.xlsx")
        rc, out, _ = run_cli(str(self.tmp / "in"), "-r", "--dry-run", "-f", "json")
        self.assertEqual(rc, 0)
        self.assertEqual(sorted(f["location"] for f in json.loads(out)["findings"]), ["a/list.xlsx", "b/list.xlsx"])


class MaxOptionRangeTest(unittest.TestCase):
    def test_zero_and_negative_are_usage_errors(self):
        for opt in ("--max-cells", "--max-unzipped-mb", "--max-zip-ratio"):
            for bad in ("0", "-5", "abc"):
                with self.subTest(opt=opt, bad=bad):
                    err = io.StringIO()
                    with contextlib.redirect_stderr(err), self.assertRaises(SystemExit) as cm:
                        cli.main(["x.xlsx", f"{opt}={bad}"])
                    self.assertEqual(cm.exception.code, 2)  # argparse の引数の誤り = 終了コード 2
                    self.assertIn(opt, err.getvalue())

    def test_positive_values_are_accepted(self):
        argv = ["x", "--max-cells", "1", "--max-unzipped-mb", "2", "--max-zip-ratio", "3"]
        args = cli.build_parser().parse_args(argv)
        self.assertEqual((args.max_cells, args.max_unzipped_mb, args.max_zip_ratio), (1, 2, 3))


if __name__ == "__main__":
    unittest.main()
