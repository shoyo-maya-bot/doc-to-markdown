#!/usr/bin/env python3
"""CI と同じ品質ゲートを、CI と同じ順番でローカル実行する（push 前の確認用）.

    python scripts/verify_gates.py            # 全ゲート
    python scripts/verify_gates.py --quick    # lint・フォーマット・テストだけ
    python scripts/verify_gates.py --fix      # lint の自動修正とフォーマットを先に適用してから確認

ゲート（CI の python-ci.yml と同じ）:
  1. lint        ruff check
  2. format      ruff format --check
  3. test        pytest
  4. sast        bandit（src/ の High 以上。未導入なら skip）
  5. deps        pip-audit（依存ライブラリの既知脆弱性。未導入・オフラインなら skip）
  6. trace       spec-trace check traceability/（traceability/ があり spec-trace が入っている場合）

終了コード: 0 すべて PASS（skip を含む）/ 1 FAIL あり。
**ゲートを通すためにテストやチェックを甘くしない**（アサーション削除・除外追加・無効化は禁止）。
"""

from __future__ import annotations

import argparse
import importlib.util
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _module(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def gates(quick: bool) -> list[tuple[str, list[str] | None, str]]:
    py = sys.executable
    out: list[tuple[str, list[str] | None, str]] = [
        ("lint", [py, "-m", "ruff", "check", "."], ""),
        ("format", [py, "-m", "ruff", "format", "--check", "."], ""),
        ("test", [py, "-m", "pytest", "-q"], ""),
    ]
    if quick:
        return out
    src = "src" if (ROOT / "src").is_dir() else "."
    out.append(
        ("sast", [py, "-m", "bandit", "-q", "-r", src, "-ll"], "")
        if _module("bandit")
        else ("sast", None, "bandit 未導入（pip install bandit）")
    )
    out.append(
        ("deps", [py, "-m", "pip_audit", "--skip-editable", "--progress-spinner", "off"], "")
        if _module("pip_audit")
        else ("deps", None, "pip-audit 未導入（pip install pip-audit）")
    )
    if (ROOT / "traceability").is_dir():
        exe = shutil.which("spec-trace")
        out.append(
            ("trace", [exe, "check", "traceability", "--root", "."], "")
            if exe
            else ("trace", None, "spec-trace 未導入")
        )
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--quick", action="store_true", help="lint・フォーマット・テストだけ")
    ap.add_argument("--fix", action="store_true", help="ruff の自動修正とフォーマットを先に適用する")
    args = ap.parse_args(argv)
    if not _module("ruff") or not _module("pytest"):
        print("ruff と pytest が必要です: pip install -e '.[dev]'（または setup の --dev）", file=sys.stderr)
        return 1
    if args.fix:
        subprocess.run([sys.executable, "-m", "ruff", "check", "--fix", "."], cwd=ROOT)
        subprocess.run([sys.executable, "-m", "ruff", "format", "."], cwd=ROOT)

    results: list[tuple[str, str, str]] = []
    for name, cmd, note in gates(args.quick):
        if cmd is None:
            results.append((name, "SKIP", note))
            continue
        print(f"\n=== {name}: {' '.join(Path(c).name if i == 0 else c for i, c in enumerate(cmd))}", flush=True)
        rc = subprocess.run(cmd, cwd=ROOT).returncode
        results.append((name, "PASS" if rc == 0 else "FAIL", "" if rc == 0 else f"終了コード {rc}"))

    print("\n| ゲート | 結果 | 補足 |\n|---|---|---|")
    for name, status, note in results:
        print(f"| {name} | {status} | {note} |")
    failed = [n for n, s, _ in results if s == "FAIL"]
    if failed:
        print(f"\nFAIL: {', '.join(failed)}。原因を直してください（チェックを甘くして通さない）。")
        return 1
    print("\nすべてのゲートを通過しました（SKIP は CI で実行されます）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
