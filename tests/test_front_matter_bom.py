"""人が BOM 付き（Windows のメモ帳など）で保存し直した出力からも Front Matter の id・source を引き継ぐ."""

from __future__ import annotations

import contextlib
import io
import shutil
import tempfile
import unittest
from pathlib import Path

from devtools_common.cli import run_main

from doc_to_markdown import cli, llmstxt
from doc_to_markdown.converter import existing_id, existing_source

SAMPLES = Path(__file__).resolve().parents[1] / "samples"


def run_cli(*args: str) -> int:
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        return run_main(cli.main, list(args))


class FrontMatterBomTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)

    def test_reads_id_and_source_after_bom(self):
        md = self.tmp / "a.md"
        md.write_bytes('﻿---\r\ntitle: "a"\r\nid: "DOC-7"\r\nsource: "a.docx"\r\n---\r\n\r\n# a\r\n'.encode())
        self.assertEqual(existing_id(md), "DOC-7")
        self.assertEqual(existing_source(md), "a.docx")

    def test_reconvert_keeps_id_from_bom_file(self):
        src = SAMPLES / "expense-claim.docx"
        self.assertEqual(run_cli(str(src), "-d", str(self.tmp), "--front-matter"), 0)
        out = self.tmp / "expense-claim.md"
        edited = out.read_text(encoding="utf-8").replace('id: ""', 'id: "DOC-001"')
        out.write_text(edited, encoding="utf-8-sig")  # BOM 付きで保存し直す
        self.assertEqual(run_cli(str(src), "-d", str(self.tmp), "--front-matter"), 0)
        self.assertIn('id: "DOC-001"', out.read_text(encoding="utf-8"))

    def test_llms_txt_title_from_bom_file(self):
        md = self.tmp / "a.md"
        md.write_text('---\ntitle: "経費精算"\n---\n\n# 経費精算\n', encoding="utf-8-sig")
        llmstxt.write(self.tmp, [("a.docx", md)])
        full = (self.tmp / "llms-full.txt").read_text(encoding="utf-8")
        self.assertNotIn("﻿", full)
        self.assertNotIn("title:", full)


if __name__ == "__main__":
    unittest.main()
