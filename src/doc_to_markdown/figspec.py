"""figspec — 図の定義（YAML / JSON）から、編集できる draw.io ファイルを作る.

AI や人が「箱と矢印」を構造化データで書き、それを draw.io（.drawio）に変換する。画像ではなく
draw.io の形式で出すので、生成後に人が draw.io で手直しでき、Git で差分も追える。

    figspec check flow.yaml                 # 定義の誤り・線のルール違反を報告
    figspec build flow.yaml -o flow.drawio  # draw.io を作る（--svg で draw.io があれば svg も書き出す）

定義の例:

    title: 承認フロー
    direction: LR            # LR（左→右）/ TB（上→下）。座標を省略したときの自動配置の向き
    nodes:
      - {id: apply, label: 申請}
      - {id: check, label: 金額は 10 万円以上？, shape: decision}
      - {id: boss, label: 部長承認}
      - {id: done, label: 完了, shape: terminal}
    edges:
      - {from: apply, to: check}
      - {from: check, to: boss, label: はい}
      - {from: check, to: done, label: いいえ}
      - {from: boss, to: done}

線のルール（読みやすい図にするための約束）:
- 線は直交（縦横のみ）。斜め線は使わない
- 端点は必ずノード（宙に浮いた線を作らない）
- 自己ループ・同じ向きの重複線は作らない
- 判断（decision）から出る線にはラベル（はい／いいえ 等）を付ける
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from collections import defaultdict, deque
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape

from devtools_common import cli as common_cli
from devtools_common import report
from devtools_common.config import load_structured
from devtools_common.report import Finding, Severity

TOOL = "figspec"

SHAPES = {
    "process": "rounded=1;whiteSpace=wrap;html=1;",
    "decision": "rhombus;whiteSpace=wrap;html=1;",
    "terminal": "ellipse;whiteSpace=wrap;html=1;",
    "data": "shape=parallelogram;perimeter=parallelogramPerimeter;whiteSpace=wrap;html=1;",
    "database": "shape=cylinder3;whiteSpace=wrap;html=1;boundedLbl=1;size=12;",
    "document": "shape=document;whiteSpace=wrap;html=1;boundedLbl=1;",
    "actor": "shape=umlActor;verticalLabelPosition=bottom;verticalAlign=top;html=1;",
    "note": "shape=note;whiteSpace=wrap;html=1;size=14;fillColor=#fff2cc;strokeColor=#d6b656;",
}
DEFAULT_SIZE = {"decision": (140, 80), "terminal": (120, 50), "actor": (30, 60)}
EDGE_STYLE = "edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;endArrow=block;"
GAP_X, GAP_Y = 80, 60


@dataclass
class Node:
    id: str
    label: str
    shape: str = "process"
    x: float | None = None
    y: float | None = None
    w: float = 140
    h: float = 60


@dataclass
class Edge:
    src: str
    dst: str
    label: str = ""
    dashed: bool = False


@dataclass
class Spec:
    title: str
    direction: str
    nodes: list[Node]
    edges: list[Edge]


def parse(data: dict) -> tuple[Spec, list[Finding]]:
    """定義を読み込む. 形式の誤りは UsageError、図としての問題は Finding で返す."""
    if not isinstance(data, dict) or not isinstance(data.get("nodes"), list):
        raise common_cli.UsageError("nodes（ノードの並び）が必要です")
    direction = str(data.get("direction", "LR")).upper()
    if direction not in ("LR", "TB"):
        raise common_cli.UsageError("direction は LR か TB です")
    nodes: list[Node] = []
    for i, n in enumerate(data["nodes"]):
        if not isinstance(n, dict) or not n.get("id"):
            raise common_cli.UsageError(f"nodes[{i}] に id がありません")
        shape = n.get("shape", "process")
        if shape not in SHAPES:
            raise common_cli.UsageError(f"nodes[{i}] の shape '{shape}' は使えません（{', '.join(SHAPES)}）")
        w, h = DEFAULT_SIZE.get(shape, (140, 60))
        nodes.append(
            Node(
                id=str(n["id"]),
                label=str(n.get("label", "")),
                shape=shape,
                x=n.get("x"),
                y=n.get("y"),
                w=float(n.get("w", w)),
                h=float(n.get("h", h)),
            )
        )
    edges = []
    for i, e in enumerate(data.get("edges") or []):
        if not isinstance(e, dict) or "from" not in e or "to" not in e:
            raise common_cli.UsageError(f"edges[{i}] に from / to がありません")
        edges.append(Edge(str(e["from"]), str(e["to"]), str(e.get("label", "")), bool(e.get("dashed", False))))
    spec = Spec(str(data.get("title", "")), direction, nodes, edges)
    return spec, check(spec)


def check(spec: Spec) -> list[Finding]:
    findings: list[Finding] = []
    ids = [n.id for n in spec.nodes]
    seen: set[str] = set()
    for nid in ids:
        if nid in seen:
            findings.append(Finding("fig.duplicate_id", Severity.ERROR, "ノード ID が重複しています", location=nid))
        seen.add(nid)
    by_id = {n.id: n for n in spec.nodes}
    for n in spec.nodes:
        if not n.label.strip() and n.shape != "actor":
            findings.append(Finding("fig.empty_label", Severity.WARNING, "ラベルが空です", location=n.id))
    pairs: set[tuple[str, str]] = set()
    for e in spec.edges:
        loc = f"{e.src}->{e.dst}"
        for end in (e.src, e.dst):
            if end not in by_id:
                findings.append(
                    Finding(
                        "fig.dangling_edge",
                        Severity.ERROR,
                        f"線の端点 {end} がノードにありません（宙に浮いた線）",
                        location=loc,
                    )
                )
        if e.src == e.dst:
            findings.append(Finding("fig.self_loop", Severity.ERROR, "自己ループの線は使わない", location=loc))
        if (e.src, e.dst) in pairs:
            findings.append(
                Finding("fig.duplicate_edge", Severity.WARNING, "同じ向きの線が重複しています", location=loc)
            )
        pairs.add((e.src, e.dst))
        src = by_id.get(e.src)
        if src is not None and src.shape == "decision" and not e.label.strip():
            findings.append(
                Finding(
                    "fig.decision_label",
                    Severity.WARNING,
                    "判断から出る線にラベル（はい／いいえ 等）がありません",
                    location=loc,
                )
            )
    connected = {e.src for e in spec.edges} | {e.dst for e in spec.edges}
    if spec.edges:
        for n in spec.nodes:
            if n.id not in connected and n.shape != "note":
                findings.append(
                    Finding("fig.isolated_node", Severity.WARNING, "どの線ともつながっていません", location=n.id)
                )
    findings += _overlaps(layout(spec))
    return findings


def _overlaps(nodes: list[Node]) -> list[Finding]:
    out = []
    for i, a in enumerate(nodes):
        for b in nodes[i + 1 :]:
            if a.x < b.x + b.w and b.x < a.x + a.w and a.y < b.y + b.h and b.y < a.y + a.h:
                out.append(Finding("fig.overlap", Severity.WARNING, f"{b.id} と重なっています", location=a.id))
    return out


def layout(spec: Spec) -> list[Node]:
    """座標の無いノードを、線の向きに沿って層ごとに並べる（座標が書いてあるノードはそのまま）."""
    ids = [n.id for n in spec.nodes]
    out_edges: dict[str, list[str]] = defaultdict(list)
    indeg = {i: 0 for i in ids}
    for e in spec.edges:
        if e.src in indeg and e.dst in indeg and e.src != e.dst:
            out_edges[e.src].append(e.dst)
            indeg[e.dst] += 1
    rank = {i: 0 for i in ids}
    queue = deque(i for i in ids if indeg[i] == 0)
    remaining = dict(indeg)
    visited: set[str] = set()
    while queue:
        cur = queue.popleft()
        visited.add(cur)
        for nxt in out_edges[cur]:
            rank[nxt] = max(rank[nxt], rank[cur] + 1)
            remaining[nxt] -= 1
            if remaining[nxt] == 0:
                queue.append(nxt)
    # 循環に含まれるノード（戻り線がある図）は、到達済みの最大層の次に置く
    top = max((rank[i] for i in visited), default=0)
    for i in ids:
        if i not in visited:
            top += 1
            rank[i] = top
    layers: dict[int, list[Node]] = defaultdict(list)
    for n in spec.nodes:
        layers[rank[n.id]].append(n)
    placed: list[Node] = []
    for r in sorted(layers):
        offset = 40.0
        for n in layers[r]:
            if n.x is None or n.y is None:
                if spec.direction == "LR":
                    n.x, n.y = 40 + r * (180 + GAP_X) + (180 - n.w) / 2, offset + (60 - n.h) / 2
                else:
                    n.x, n.y = offset + (180 - n.w) / 2, 40 + r * (80 + GAP_Y) + (80 - n.h) / 2
            offset += (60 + GAP_Y) if spec.direction == "LR" else (180 + GAP_X)
            placed.append(n)
    return placed


def to_drawio(spec: Spec) -> str:
    """draw.io（mxfile、非圧縮）の XML を返す. 同じ定義からは常に同じ出力になる."""
    nodes = layout(spec)
    cells = ['<mxCell id="0"/>', '<mxCell id="1" parent="0"/>']
    for n in nodes:
        cells.append(
            f'<mxCell id="n-{escape(n.id, {chr(34): "&quot;"})}" value="{escape(n.label, {chr(34): "&quot;"})}" '
            f'style="{SHAPES[n.shape]}" vertex="1" parent="1">'
            f'<mxGeometry x="{n.x:g}" y="{n.y:g}" width="{n.w:g}" height="{n.h:g}" as="geometry"/></mxCell>'
        )
    known = {n.id for n in nodes}
    for i, e in enumerate(spec.edges, 1):
        if e.src not in known or e.dst not in known:
            continue
        style = EDGE_STYLE + ("dashed=1;" if e.dashed else "")
        cells.append(
            f'<mxCell id="e-{i}" value="{escape(e.label, {chr(34): "&quot;"})}" style="{style}" edge="1" parent="1" '
            f'source="n-{escape(e.src, {chr(34): "&quot;"})}" target="n-{escape(e.dst, {chr(34): "&quot;"})}">'
            '<mxGeometry relative="1" as="geometry"/></mxCell>'
        )
    title = escape(spec.title or "図", {chr(34): "&quot;"})
    body = "\n        ".join(cells)
    return (
        '<mxfile host="figspec">\n'
        f'  <diagram id="figspec" name="{title}">\n'
        '    <mxGraphModel grid="1" gridSize="10" guides="1" arrows="1" connect="1" page="1" pageScale="1">\n'
        f"      <root>\n        {body}\n      </root>\n"
        "    </mxGraphModel>\n  </diagram>\n</mxfile>\n"
    )


def export_svg(drawio: Path) -> Path | None:
    from devtools_common.executables import DRAWIO, find_executable

    exe = find_executable(DRAWIO)
    if not exe:
        return None
    svg = drawio.with_suffix(".svg")
    subprocess.run([exe, "--export", "--format", "svg", "--output", str(svg), str(drawio)], check=True, timeout=120)
    return svg


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog=TOOL, description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="command", required=True)
    c = sub.add_parser("check", help="定義の誤り・線のルール違反を報告する")
    c.add_argument("spec", type=Path)
    common_cli.add_common_args(c)
    b = sub.add_parser("build", help="draw.io ファイルを作る")
    b.add_argument("spec", type=Path)
    b.add_argument("-d", "--drawio", type=Path, help="出力先（既定: 定義と同じ名前の .drawio）")
    b.add_argument("--svg", action="store_true", help="draw.io Desktop があれば svg も書き出す")
    b.add_argument("--strict", action="store_true", help="警告があっても作らない（終了コード 1）")
    common_cli.add_common_args(b)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    common_cli.configure_from_args(args)
    if not args.spec.is_file():
        raise common_cli.UsageError(f"定義ファイルが見つかりません: {args.spec}")
    data = load_structured(args.spec) if args.spec.suffix != ".json" else json.loads(args.spec.read_text("utf-8"))
    spec, findings = parse(data)
    has_error = report.has_errors(findings)
    if args.command == "build" and not has_error and not (args.strict and findings):
        out = args.drawio or args.spec.with_suffix(".drawio")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(to_drawio(spec), encoding="utf-8", newline="\n")
        findings.append(Finding("fig.built", Severity.INFO, f"→ {out.as_posix()}", location=args.spec.name))
        if args.svg:
            svg = export_svg(out)
            findings.append(
                Finding("fig.svg", Severity.INFO, f"→ {svg.as_posix()}", location=args.spec.name)
                if svg
                else Finding("fig.svg", Severity.WARNING, "draw.io Desktop が無いため svg は作りませんでした")
            )
    text = report.render(
        findings,
        args.format,
        title="figspec レポート",
        meta={"nodes": len(spec.nodes), "edges": len(spec.edges)},
        tool=TOOL,
    )
    common_cli.write_output(text, args.output)
    if has_error or (args.command == "build" and args.strict and any(f.severity is Severity.WARNING for f in findings)):
        return common_cli.ExitCode.FINDINGS
    return common_cli.ExitCode.OK


def entrypoint() -> None:
    sys.exit(common_cli.run_main(main))


if __name__ == "__main__":
    entrypoint()
