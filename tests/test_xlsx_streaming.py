"""大きなブックの省メモリ読み込みでも --header-row と --skip-hidden が効く."""

from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl

from doc_to_markdown.converter import convert_file
from doc_to_markdown.model import Options


class StreamingOptionsTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)

    def book(self, rows, *, hidden_rows=(), hidden_cols=()) -> Path:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "明細"
        for row in rows:
            ws.append(row)
        for r in hidden_rows:
            ws.row_dimensions[r].hidden = True
        for c in hidden_cols:
            ws.column_dimensions[c].hidden = True
        path = self.tmp / "big.xlsx"
        wb.save(path)
        return path

    def convert(self, path: Path, **opts):
        r = convert_file(path, self.tmp / "out", options=Options(xlsx_full_load_max_cells=1, **opts))
        return r, r.output.read_text(encoding="utf-8")

    def test_header_row_is_used(self):
        # 1〜3 行目はすべて文字の行. 自動推定だと 1 行目が見出しになる
        path = self.book(
            [["社内", "資料", "一覧"], ["作成", "総務", "部"], ["コード", "名称", "区分"], ["A1", "本社", "x"]]
        )
        _, md = self.convert(path, header_rows={"明細": 3})
        self.assertIn("社内 / 資料 / 一覧", md)
        self.assertIn("作成 / 総務 / 部", md)
        self.assertIn("| コード | 名称 | 区分 |\n| --- | --- | --- |\n| A1 | 本社 | x |", md)

    def test_header_row_star_applies_to_all_sheets(self):
        path = self.book([["題名", "", ""], ["コード", "名称", "区分"], ["A1", "本社", "x"]])
        _, md = self.convert(path, header_rows={"*": 2})
        self.assertIn("題名\n\n| コード | 名称 | 区分 |", md)

    def test_skip_hidden_rows_and_columns(self):
        path = self.book(
            [["コード", "秘密", "名称"], ["A1", "s1", "本社"], ["A2", "s2", "隠す支社"], ["A3", "s3", "支社"]],
            hidden_rows=[3],
            hidden_cols=["B"],
        )
        _, md = self.convert(path, skip_hidden=True)
        self.assertIn("| コード | 名称 |", md)
        self.assertIn("| A1 | 本社 |", md)
        self.assertIn("| A3 | 支社 |", md)
        self.assertNotIn("隠す支社", md)
        self.assertNotIn("秘密", md)

    def test_hidden_rows_are_reported_when_kept(self):
        path = self.book([["コード", "名称"], ["A1", "本社"], ["A2", "隠す支社"]], hidden_rows=[3])
        r, md = self.convert(path)
        self.assertIn("隠す支社", md)
        msgs = [f.message for f in r.findings if f.rule_id == "convert.warning"]
        self.assertTrue(any("非表示の行 1 行" in m for m in msgs), msgs)


if __name__ == "__main__":
    unittest.main()
