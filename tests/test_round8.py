"""第 8 回: LibreOffice が変換できないファイルの終了コード・--front-matter のヘルプ・終了コード・既定の OCR 言語."""

from __future__ import annotations

import contextlib
import io
import os
import re
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from devtools_common.cli import ExitCode, UsageError, run_main
from devtools_common.executables import TESSERACT, find_executable

from doc_to_markdown import cli, office
from doc_to_markdown.converter import convert_file, front_matter, is_broken_input
from doc_to_markdown.model import Options

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")
SPEC = (ROOT / "docs" / "OUTPUT_SPEC.md").read_text(encoding="utf-8")
REAL = ROOT / "samples" / "realworld"


class LibreOfficeFailureTest(unittest.TestCase):
    """以前は LibreOffice の RuntimeError を is_broken_input が拾わず、壊れた .doc が終了コード 3 だった."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)
        self.src = self.tmp / "broken.doc"
        self.src.write_bytes(b"not a word document")
        p = mock.patch.object(office, "find_soffice", return_value="/usr/bin/soffice")
        p.start()
        self.addCleanup(p.stop)

    def run_cli(self, *args: str) -> tuple[int, str]:
        err = io.StringIO()
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
            rc = run_main(cli.main, [*args])
        return rc, err.getvalue()

    def test_libreoffice_exit_nonzero_is_input_error(self):
        boom = subprocess.CalledProcessError(1, ["soffice"])
        with mock.patch.object(office.subprocess, "run", side_effect=boom):
            with self.assertRaises(UsageError):
                convert_file(self.src, self.tmp / "out")
            rc, err = self.run_cli(str(self.src), "-d", str(self.tmp / "out"))
        self.assertEqual(rc, ExitCode.USAGE, err)
        self.assertIn("broken.doc", err)

    def test_libreoffice_no_output_is_input_error(self):
        ok = subprocess.CompletedProcess(["soffice"], 0)
        with mock.patch.object(office.subprocess, "run", return_value=ok):
            rc, err = self.run_cli(str(self.src), "-d", str(self.tmp / "out"))
        self.assertEqual(rc, ExitCode.USAGE, err)

    def test_continue_on_error_records_input_error(self):
        ok = subprocess.CompletedProcess(["soffice"], 0)
        rep = self.tmp / "r.json"
        with mock.patch.object(office.subprocess, "run", return_value=ok):
            rc, _ = self.run_cli(
                str(self.src), "-d", str(self.tmp / "out"), "--continue-on-error", "-f", "json", "-o", str(rep)
            )
        self.assertEqual(rc, ExitCode.FINDINGS)
        self.assertIn("convert.input_error", rep.read_text(encoding="utf-8"))

    def test_timeout_stays_runtime_error(self):
        e = office.LibreOfficeInputError("x")
        self.assertTrue(is_broken_input(e))
        self.assertIsInstance(e, RuntimeError)  # RuntimeError で捕まえていた呼び出し側と互換
        timeout = subprocess.TimeoutExpired(["soffice"], 600)
        with mock.patch.object(office.subprocess, "run", side_effect=timeout):
            rc, _ = self.run_cli(str(self.src), "-d", str(self.tmp / "out"))
        self.assertEqual(rc, ExitCode.RUNTIME)


class HelpAndDocsTest(unittest.TestCase):
    def test_front_matter_help_lists_every_field(self):
        """--front-matter の --help に、実際に出力する Front Matter のキーがすべてある."""
        src = Path(tempfile.mkdtemp()) / "a.docx"
        src.write_bytes(b"x")
        try:
            keys = re.findall(r"^(\w+):", front_matter(src, "a"), re.M)
        finally:
            shutil.rmtree(src.parent)
        action = next(a for a in cli.build_parser()._actions if "--front-matter" in a.option_strings)
        self.assertEqual(keys, ["title", "id", "source", "source_sha256", "generator"])
        for k in keys:
            self.assertIn(k, action.help)
            self.assertIn(f"`{k}`", README)

    def test_dry_run_exit_code_is_documented(self):
        """OUTPUT_SPEC の dry-run の「error があれば終了コード 1」が README・--help の終了コードにもある."""
        self.assertIn("error があれば終了コード 1", SPEC)
        section = README.split("### 終了コード", 1)[1].split("\n## ", 1)[0]
        self.assertIn("plan.missing_tool", section)
        self.assertIn("plan.missing_tool", cli.__doc__ or "")


def _tesseract_langs() -> list[str] | None:
    exe = find_executable(TESSERACT)
    if not exe:
        return None
    r = subprocess.run([exe, "--list-langs"], capture_output=True, text=True, timeout=60)
    return r.stdout.split()


class DefaultOcrLangTest(unittest.TestCase):
    """既定の --ocr-lang（jpn+eng）で OCR できる. CI（環境変数 CI）で Tesseract があるのに jpn が無ければ失敗にする."""

    def test_default_ocr_lang(self):
        default = cli.build_parser().get_default("ocr_lang")
        self.assertEqual(default, Options().ocr_lang)
        langs = _tesseract_langs()
        if langs is None:
            self.skipTest("Tesseract が無い")
        missing = [x for x in default.split("+") if x not in langs]
        if missing:
            if os.environ.get("CI"):
                self.fail(f"CI に Tesseract の言語データ {missing} がありません（ci.yml の apt に tesseract-ocr-jpn）")
            self.skipTest(f"Tesseract の言語データ {missing} が無い")
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp)
        r = convert_file(REAL / "scanned.pdf", tmp / "out", options=Options())
        md = r.output.read_text(encoding="utf-8")
        self.assertIn("Expense Report", md)
        self.assertFalse(any("jpn が必要" in f.message for f in r.findings))


class CiInstallsJapaneseOcrTest(unittest.TestCase):
    def test_ci_installs_tesseract_jpn_everywhere_tesseract_is_installed(self):
        ci = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")
        lines = [ln for ln in ci.splitlines() if re.search(r"\btesseract-ocr\b", ln)]
        self.assertTrue(lines)
        for ln in lines:
            self.assertIn("tesseract-ocr-jpn", ln)


if __name__ == "__main__":
    unittest.main()
