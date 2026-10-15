"""PDF のヘッダー・フッター（ページの端に繰り返し出る行）の除去で、本文を誤って消さない（第 6 回の指摘）.

- 数字だけが違う行（「Total 1200 yen」と「Total 980 yen」）は、ページ番号の歩みで増えて全ページに出るときだけ除く
- 除いた文字数はロス検知（loss.raw_text）の元の文字数から引かない
"""

from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from doc_to_markdown.adapters.pdf_adapter import _running_signatures
from doc_to_markdown.converter import convert_file
from doc_to_markdown.model import Options


def simple_pdf(path: Path, pages: list[list[tuple[float, float, float, str]]]) -> None:
    """(x, y, 文字サイズ, ASCII の文字) を並べただけの PDF を作る（Helvetica. y は下から）."""
    objs: list[bytes] = [b"", b""]  # 1: Catalog, 2: Pages（後で埋める）
    objs.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")  # 3
    kids = []
    for items in pages:
        stream = "".join(f"BT /F1 {s} Tf {x} {y} Td ({t}) Tj ET\n" for x, y, s, t in items).encode("ascii")
        objs.append(b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"endstream")
        content = len(objs)
        objs.append(
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
            b"/Resources << /Font << /F1 3 0 R >> >> /Contents %d 0 R >>" % content
        )
        kids.append(len(objs))
    objs[0] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objs[1] = b"<< /Type /Pages /Kids [%s] /Count %d >>" % (b" ".join(b"%d 0 R" % k for k in kids), len(kids))
    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for i, body in enumerate(objs, start=1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % i + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    out += b"".join(b"%010d 00000 n \n" % o for o in offsets)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
    path.write_bytes(bytes(out))


def page(n: int, total: int, footer_body: str | None) -> list[tuple[float, float, float, str]]:
    items = [(40, 815, 9, "Expense Report Rev 3"), (40, 700, 12, f"Section {n} body text about expenses.")]
    items += [(40, 680 - 16 * i, 12, f"Line {i} of the policy for page {n} describes the rule.") for i in range(4)]
    if footer_body:
        items.append((40, 45, 12, footer_body))
    items.append((500, 20, 9, f"Page {n} of {total}"))
    return items


class RunningSignaturesTest(unittest.TestCase):
    def test_page_numbers_on_all_pages(self):
        edges = {"Page#of#": {0: "Page1of3", 1: "Page2of3", 2: "Page3of3"}}
        self.assertEqual(_running_signatures(edges, 3), {"#Page#of#"})
        edges = {"-#-": {0: "-5-", 1: "-6-", 2: "-7-"}}  # 途中の番号から始まる
        self.assertEqual(_running_signatures(edges, 3), {"#-#-"})

    def test_numbers_that_do_not_follow_pages_are_kept(self):
        edges = {"Total#yen": {0: "Total1200yen", 1: "Total980yen", 2: "Total4500yen"}}
        self.assertEqual(_running_signatures(edges, 3), set())
        # 1 ページ目だけ同じ数字（2 と 3 の歩みが合わない）
        edges = {"No.#": {0: "No.1", 1: "No.1", 2: "No.3"}}
        self.assertEqual(_running_signatures(edges, 3), set())

    def test_varying_numbers_need_every_page(self):
        # 表紙にページ番号が無い: 数字違いの行は全ページに無いので除かない（誤って本文を消さない側に倒す）
        edges = {"Page#": {1: "Page2", 2: "Page3", 3: "Page4"}}
        self.assertEqual(_running_signatures(edges, 4), set())

    def test_same_text_on_half_of_pages(self):
        edges = {"社外秘": {0: "社外秘", 2: "社外秘"}, "目次": {0: "目次"}}
        self.assertEqual(_running_signatures(edges, 4), {"=社外秘"})


class PdfRunningHeadersTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)

    def convert(self, pages, **opts):
        src = self.tmp / "report.pdf"
        simple_pdf(src, pages)
        r = convert_file(src, self.tmp / "out", options=Options(**opts))
        return r, r.output.read_text(encoding="utf-8")

    def test_body_lines_with_different_numbers_are_kept(self):
        # 再現: ページの下端の「Total 1200 yen」などが、数字を # にすると同じ行なので、
        # ヘッダー・フッターとして消えていた
        totals = ["Total 1200 yen", "Total 980 yen", "Total 4500 yen"]
        r, md = self.convert([page(i + 1, 3, t) for i, t in enumerate(totals)])
        for t in totals:
            self.assertIn(t, md)
        self.assertNotIn("Page 1 of 3", md)  # ページ番号は除く
        self.assertNotIn("Expense Report Rev 3", md)  # 同じ文字のヘッダーは除く
        self.assertIn("Section 2 body text", md)

    def test_removed_chars_are_not_subtracted_from_loss_check(self):
        # 除いた分を元の文字数から引くと、誤って本文を消してもロス検知が効かない.
        # 引かないので、除いた分が多いと warning になる
        notice = (40, 30, 12, "Confidential - do not distribute outside the team")
        pages = [page(i + 1, 3, None) + [notice] for i in range(3)]
        r, md = self.convert(pages)
        self.assertNotIn("Confidential", md)
        rules = [f.rule_id for f in r.findings]
        self.assertIn("loss.raw_text", rules)
        note = next(f.message for f in r.findings if f.rule_id == "convert.note" and "ヘッダー" in f.message)
        self.assertRegex(note, r"\d+ 字を除去")
        # --keep-running-headers なら除かないのでロスにならない
        r2, md2 = self.convert(pages, remove_running_headers=False)
        self.assertIn("Confidential", md2)
        self.assertNotIn("loss.raw_text", [f.rule_id for f in r2.findings])


if __name__ == "__main__":
    unittest.main()
