"""ロス検知（本文の文字数）のしきい値を --loss-threshold で変えられる."""

from __future__ import annotations

import contextlib
import io
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from devtools_common.cli import run_main

from doc_to_markdown import cli, lossaudit
from doc_to_markdown.model import SourceStats

SAMPLES = Path(__file__).resolve().parents[1] / "samples"


def run_cli(*args: str) -> int:
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        return run_main(cli.main, list(args))


def _output(chars: int) -> lossaudit.OutputStats:
    out = lossaudit.OutputStats()
    out.structure.text_chars = chars
    return out


class LossThresholdTest(unittest.TestCase):
    def rules(self, findings) -> list[str]:
        return [f.rule_id for f in findings]

    def test_default_threshold_misses_small_loss(self):
        found = lossaudit.audit(SourceStats(text_chars=1000), _output(970), location="a", raw_text_chars=1000)
        self.assertEqual(self.rules(found), [])

    def test_stricter_threshold_detects_small_loss(self):
        found = lossaudit.audit(
            SourceStats(text_chars=1000), _output(970), location="a", raw_text_chars=1000, text_ratio_min=0.99
        )
        self.assertEqual(self.rules(found), ["loss.text", "loss.raw_text"])

    def test_cli_passes_threshold_to_audit(self):
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp)
        seen: list[float] = []
        real = lossaudit.audit

        def spy(*a, **kw):
            seen.append(kw["text_ratio_min"])
            return real(*a, **kw)

        with mock.patch.object(lossaudit, "audit", side_effect=spy):
            rc = run_cli(str(SAMPLES / "expense-claim.docx"), "-d", str(tmp), "--loss-threshold", "99.5")
        self.assertEqual(rc, 0)
        self.assertEqual(seen, [0.995])

    def test_cli_rejects_out_of_range(self):
        for bad in ("0", "101", "-5", "nan", "abc"):
            with self.subTest(bad=bad):
                err = io.StringIO()
                with contextlib.redirect_stderr(err), self.assertRaises(SystemExit) as cm:
                    cli.build_parser().parse_args([str(SAMPLES / "expense-claim.docx"), "--loss-threshold", bad])
                self.assertEqual(cm.exception.code, 2)  # 引数の誤り
                self.assertIn("--loss-threshold", err.getvalue())


if __name__ == "__main__":
    unittest.main()
