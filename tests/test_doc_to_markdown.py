"""doc-to-markdown のテスト.

ゴールデンテスト: samples/ の各文書の変換結果を tests/golden/*.md と比較する。
意図した変更で差分が出たら `UPDATE_GOLDEN=1 pytest` で期待値を更新し、差分を PR でレビューする。
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from devtools_common.report import Severity

from doc_to_markdown import cli, lossaudit
from doc_to_markdown.adapters import drawio_adapter
from doc_to_markdown.converter import convert_file, find_soffice
from doc_to_markdown.model import Heading, ListItem, Paragraph, SourceStats, Table
from doc_to_markdown.normalize import normalize
from doc_to_markdown.render import render

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"
GOLDEN = Path(__file__).resolve().parent / "golden"
SAMPLE_FILES = sorted(p for p in SAMPLES.iterdir() if p.suffix in {".docx", ".xlsx", ".pptx", ".pdf", ".drawio"})


def run_cli(*args: str) -> tuple[int, str, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        from devtools_common.cli import run_main

        rc = run_main(cli.main, list(args))
    return rc, out.getvalue(), err.getvalue()


class Tmp(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        # CI に draw.io CLI は無い前提でゴールデンを固定する
        patcher = mock.patch.object(drawio_adapter, "find_drawio_cli", return_value=None)
        patcher.start()
        self.addCleanup(patcher.stop)

    def tearDown(self):
        shutil.rmtree(self.tmp)


class GoldenTest(Tmp):
    def test_samples_match_golden(self):
        self.assertGreaterEqual(len(SAMPLE_FILES), 6)
        for src in SAMPLE_FILES:
            with self.subTest(src.name):
                result = convert_file(src, self.tmp)
                actual = result.output.read_text(encoding="utf-8")
                golden = GOLDEN / f"{src.stem}.md"
                if os.environ.get("UPDATE_GOLDEN"):
                    golden.parent.mkdir(exist_ok=True)
                    golden.write_text(actual, encoding="utf-8", newline="\n")
                self.assertEqual(actual, golden.read_text(encoding="utf-8"))
                losses = [f for f in result.findings if f.severity is Severity.WARNING]
                self.assertEqual(losses, [], f"unexpected loss warnings for {src.name}")

    def test_deterministic_output(self):
        a, b = self.tmp / "a", self.tmp / "b"
        for src in SAMPLE_FILES:
            convert_file(src, a)
            convert_file(src, b)
        files_a = sorted(p.relative_to(a) for p in a.rglob("*") if p.is_file())
        files_b = sorted(p.relative_to(b) for p in b.rglob("*") if p.is_file())
        self.assertEqual(files_a, files_b)
        for rel in files_a:
            self.assertEqual((a / rel).read_bytes(), (b / rel).read_bytes(), str(rel))

    def test_rerun_removes_stale_assets(self):
        src = SAMPLES / "expense-claim.docx"
        convert_file(src, self.tmp)
        stale = self.tmp / "expense-claim.assets" / "stale.png"
        stale.write_bytes(b"x")
        convert_file(src, self.tmp)
        self.assertFalse(stale.exists())


class AdapterDetailTest(Tmp):
    def test_docx_structure(self):
        md = convert_file(SAMPLES / "expense-claim.docx", self.tmp).output.read_text(encoding="utf-8")
        self.assertTrue(md.startswith("# 経費精算 申請書\n"))
        self.assertIn("## 1. 概要", md)  # Title があるので Heading 1 は H2
        self.assertIn("**出張旅費**", md)
        self.assertIn("  - 日当", md)  # List Bullet 2 は 2 段目
        self.assertIn("1. 申請書を作成する\n2. 上長が承認する", md)
        self.assertIn("![expense-claim-01](expense-claim.assets/expense-claim-01.png)", md)
        self.assertTrue((self.tmp / "expense-claim.assets" / "expense-claim-01.png").is_file())

    def test_xlsx_values(self):
        md = convert_file(SAMPLES / "expense-list.xlsx", self.tmp).output.read_text(encoding="utf-8")
        self.assertIn("| A-002 | 2026-04-02 | 宿泊費 | 8000 | 800 | 8900 | 佐藤 | D10 |", md)
        self.assertIn("## 部門マスタ", md)

    def test_drawio_compressed_equals_plain(self):
        plain = convert_file(SAMPLES / "approval-flow.drawio", self.tmp).output.read_text(encoding="utf-8")
        comp = convert_file(SAMPLES / "approval-flow-compressed.drawio", self.tmp).output.read_text(encoding="utf-8")
        body = lambda s: s.split("## ", 1)[1]  # noqa: E731
        self.assertEqual(body(plain), body(comp))
        self.assertIn("- 申請者 → 上長 （承認）（申請）", plain)
        self.assertTrue((self.tmp / "approval-flow.assets" / "approval-flow.drawio").is_file())

    @unittest.skipUnless(find_soffice(), "LibreOffice not installed")
    def test_via_pdf_docx_and_temp_cleanup(self):
        before = {p.name for p in Path(tempfile.gettempdir()).glob("doc2md-*")}
        r = convert_file(SAMPLES / "expense-claim.docx", self.tmp, via_pdf=True)
        after = {p.name for p in Path(tempfile.gettempdir()).glob("doc2md-*")}
        self.assertEqual(before, after, "temporary PDF directory must be removed")
        md = r.output.read_text(encoding="utf-8")
        self.assertTrue(r.via_pdf)
        self.assertIn("| A-001 | 交通費 | 1000 | 100 | 1100 |", md)  # 罫線なし表の検出
        self.assertIn("  - 日当", md)  # 字下げから箇条書きレベルを推定
        self.assertTrue(md.startswith("# expense-claim\n"))


class NormalizeRenderTest(unittest.TestCase):
    def test_heading_levels_do_not_skip(self):
        out = list(normalize([Heading(1, "A"), Heading(4, "B"), Heading(2, " "), Heading(3, "C")]))
        self.assertEqual([(h.level, h.text) for h in out], [(1, "A"), (2, "B"), (3, "C")])

    def test_table_is_rectangular_and_drops_empty(self):
        (t,) = list(normalize([Table([["a", "", "b"], ["1"], ["", "", ""]])]))
        self.assertEqual(t.rows, [["a", "b"], ["1", ""]])

    def test_render_list_numbering_and_escape(self):
        md = render(
            [
                ListItem("x", 0, True),
                ListItem("y", 1, True),
                ListItem("z", 0, True),
                Paragraph("p"),
                Table([["a|b", "c"], ["1", "2"]]),
            ]
        )
        self.assertIn("1. x\n  1. y\n2. z", md)
        self.assertIn("| a\\|b | c |", md)

    def test_list_level_capped(self):
        out = list(normalize([ListItem("a", 0), ListItem("b", 3)]))
        self.assertEqual(out[1].level, 1)


class LossAuditTest(unittest.TestCase):
    def test_detects_missing_table_and_text(self):
        src = SourceStats(tables=2, table_rows=4, text_chars=100)
        out = lossaudit.OutputStats()
        out.add("| a | b |\n| --- | --- |\n| 1 | 2 |", synthetic=False)
        rules = {f.rule_id for f in lossaudit.audit(src, out, location="x")}
        self.assertEqual(rules, {"loss.tables", "loss.table_rows", "loss.text"})

    def test_raw_text_check_catches_what_adapter_missed(self):
        out = lossaudit.OutputStats()
        out.add("本文だけ", synthetic=False)
        found = lossaudit.audit(SourceStats(text_chars=4), out, location="x", raw_text_chars=40)
        self.assertEqual([f.rule_id for f in found], ["loss.raw_text"])
        # 出力が多い（縦結合の繰り返し等）のは問題にしない
        self.assertEqual(lossaudit.audit(SourceStats(text_chars=1), out, location="x", raw_text_chars=2), [])

    def test_markup_not_counted_as_text(self):
        st = lossaudit.markdown_stats("## 見出し\n\n- **太字**\n\n![a](b.png)\n\n<!-- page 1 -->\n")
        self.assertEqual((st.headings, st.list_items, st.images, st.text_chars), (1, 1, 1, 5))


class CliTest(Tmp):
    def test_json_report_and_exit_code(self):
        rc, out, _ = run_cli(str(SAMPLES / "expense-list.xlsx"), "-d", str(self.tmp), "-f", "json")
        self.assertEqual(rc, 0)
        data = json.loads(out)
        import jsonschema
        from devtools_common.report import report_schema

        jsonschema.validate(data, report_schema())  # 出力形式の仕様（schema_version 1.x）に適合
        self.assertEqual(data["tool"], "doc2md")
        self.assertEqual(data["meta"]["converted"], 1)
        self.assertTrue((self.tmp / "expense-list.md").is_file())

    def test_unsupported_is_usage_error(self):
        bad = self.tmp / "a.txt"
        bad.write_text("x", encoding="utf-8")
        rc, _, err = run_cli(str(bad), "-d", str(self.tmp))
        self.assertEqual(rc, 2)
        self.assertIn("未対応", err)

    def test_missing_input(self):
        rc, _, _ = run_cli(str(self.tmp / "nope.docx"))
        self.assertEqual(rc, 2)

    def test_continue_on_error(self):
        broken = self.tmp / "in" / "broken.docx"
        broken.parent.mkdir()
        broken.write_bytes(b"not a zip")
        shutil.copy(SAMPLES / "expense-list.xlsx", broken.parent)
        rc, out, _ = run_cli(str(broken.parent), "-d", str(self.tmp / "out"), "--continue-on-error", "-f", "json")
        self.assertEqual(rc, 1)
        data = json.loads(out)
        self.assertEqual(data["summary"]["error"], 1)
        self.assertEqual(data["meta"]["converted"], 1)


if __name__ == "__main__":
    unittest.main()
