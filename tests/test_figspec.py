"""figspec のテスト: 定義の検査、draw.io の生成、doc2md での読み戻し（往復）."""

from __future__ import annotations

import contextlib
import io
import json
import shutil
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

from devtools_common.cli import run_main

from doc_to_markdown import figspec
from doc_to_markdown.converter import convert_file

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "samples" / "figspec" / "approval-flow.yaml"


def run(*args: str) -> tuple[int, dict]:
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
        rc = run_main(figspec.main, [*args, "-f", "json"])
    return rc, json.loads(out.getvalue()) if out.getvalue() else {}


def rules(data: dict) -> set[str]:
    return {f["rule_id"] for f in data.get("findings", [])}


class FigspecTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def spec(self, data: dict) -> Path:
        p = self.tmp / "s.json"
        p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return p

    def test_sample_builds_and_round_trips_through_doc2md(self):
        out = self.tmp / "flow.drawio"
        rc, data = run("build", str(SAMPLE), "-d", str(out))
        self.assertEqual(rc, 0)
        self.assertEqual(rules(data), {"fig.built"})
        root = ET.parse(out).getroot()
        self.assertEqual(len([c for c in root.iter("mxCell") if c.get("vertex") == "1"]), 6)
        edges = [c for c in root.iter("mxCell") if c.get("edge") == "1"]
        self.assertTrue(all("orthogonalEdgeStyle" in c.get("style") for c in edges))
        md = (convert_file(out, self.tmp / "md").output).read_text(encoding="utf-8")
        self.assertIn("- 金額は 10 万円以上？ → 部長承認（はい）", md)

    def test_output_is_deterministic(self):
        spec, _ = figspec.parse(json.loads(json.dumps({"nodes": [{"id": "a"}, {"id": "b"}], "edges": []})))
        again, _ = figspec.parse({"nodes": [{"id": "a"}, {"id": "b"}], "edges": []})
        self.assertEqual(figspec.to_drawio(spec), figspec.to_drawio(again))

    def test_dangling_edge_and_self_loop_are_errors(self):
        p = self.spec(
            {"nodes": [{"id": "a", "label": "A"}], "edges": [{"from": "a", "to": "x"}, {"from": "a", "to": "a"}]}
        )
        rc, data = run("build", str(p))
        self.assertEqual(rc, 1)
        self.assertTrue({"fig.dangling_edge", "fig.self_loop"} <= rules(data))
        self.assertFalse((self.tmp / "s.drawio").exists())

    def test_warnings(self):
        p = self.spec(
            {
                "nodes": [
                    {"id": "q", "label": "OK?", "shape": "decision"},
                    {"id": "y", "label": "Y"},
                    {"id": "lonely", "label": ""},
                    {"id": "o1", "label": "1", "x": 0, "y": 0},
                    {"id": "o2", "label": "2", "x": 10, "y": 10},
                ],
                "edges": [{"from": "q", "to": "y"}, {"from": "q", "to": "y", "label": "x"}, {"from": "o1", "to": "o2"}],
            }
        )
        rc, data = run("check", str(p))
        self.assertEqual(rc, 0)
        self.assertEqual(
            rules(data),
            {"fig.decision_label", "fig.duplicate_edge", "fig.empty_label", "fig.isolated_node", "fig.overlap"},
        )
        rc, _ = run("build", str(p), "--strict")
        self.assertEqual(rc, 1)

    def test_bad_definition_is_usage_error(self):
        for bad in (
            {"nodes": [{"label": "no id"}]},
            {"nodes": [{"id": "a", "shape": "star"}]},
            {"direction": "RL", "nodes": []},
        ):
            with self.subTest(bad=bad):
                rc, _ = run("check", str(self.spec(bad)))
                self.assertEqual(rc, 2)

    def test_cycle_is_laid_out_without_overlap(self):
        spec, findings = figspec.parse(
            {
                "nodes": [{"id": i, "label": i} for i in "abc"],
                "edges": [
                    {"from": "a", "to": "b"},
                    {"from": "b", "to": "c"},
                    {"from": "c", "to": "b", "label": "差し戻し"},
                ],
            }
        )
        self.assertNotIn("fig.overlap", {f.rule_id for f in findings})

    def test_xml_special_characters_are_escaped(self):
        spec, _ = figspec.parse({"nodes": [{"id": "a", "label": 'A & <B> "C"'}], "edges": []})
        cell = [c for c in ET.fromstring(figspec.to_drawio(spec)).iter("mxCell") if c.get("vertex")][0]
        self.assertEqual(cell.get("value"), 'A & <B> "C"')


if __name__ == "__main__":
    unittest.main()
