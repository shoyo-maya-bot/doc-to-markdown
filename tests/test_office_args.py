"""LibreOffice に渡すファイルの引数（`-` で始まるファイル名がオプションとして読まれない. 第 6 回の指摘）."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from doc_to_markdown import office


class SofficeArgsTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)

    def test_dash_file_name_is_passed_as_absolute_path(self):
        (self.tmp / "-x.docx").write_bytes(b"")
        work = self.tmp / "work"
        work.mkdir()
        seen = {}

        def fake_run(cmd, **kwargs):
            seen["cmd"] = cmd
            (work / "lo-out" / "-x.pdf").write_bytes(b"%PDF")
            return subprocess.CompletedProcess(cmd, 0)

        cwd = os.getcwd()
        os.chdir(self.tmp)
        try:
            with (
                mock.patch.object(office, "find_soffice", return_value="soffice"),
                mock.patch.object(office.subprocess, "run", side_effect=fake_run),
            ):
                out = office.office_to_pdf(Path("-x.docx"), work)
        finally:
            os.chdir(cwd)
        arg = seen["cmd"][-1]
        self.assertFalse(arg.startswith("-"), arg)
        self.assertTrue(Path(arg).is_absolute())
        self.assertEqual(Path(arg), (self.tmp / "-x.docx").resolve())
        self.assertEqual(out.name, "-x.pdf")


if __name__ == "__main__":
    unittest.main()
