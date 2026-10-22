"""ドキュメント（README・docs/OUTPUT_SPEC.md・setup スクリプト・CHANGELOG）と実装・設定のずれを機械で見つける."""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

from devtools_common.cli import ExitCode

from doc_to_markdown import cli, lossaudit

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")
SPEC = (ROOT / "docs" / "OUTPUT_SPEC.md").read_text(encoding="utf-8")
PYPROJECT = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))


def _min_python() -> tuple[int, int]:
    m = re.fullmatch(r">=(\d+)\.(\d+)", PYPROJECT["project"]["requires-python"])
    assert m, PYPROJECT["project"]["requires-python"]
    return int(m.group(1)), int(m.group(2))


def test_python_version_is_the_same_everywhere():
    major, minor = _min_python()
    want = f"{major}.{minor}"
    assert f"Python {want} 以上が必要です" in README
    assert f"sys.version_info >= ({major}, {minor})" in (ROOT / "setup.sh").read_text(encoding="utf-8")
    assert f"sys.version_info >= ({major}, {minor})" in (ROOT / "scripts" / "setup.ps1").read_text(encoding="utf-8")
    assert PYPROJECT["tool"]["ruff"]["target-version"] == f"py{major}{minor}"
    ci = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")
    versions = re.search(r"python-versions:\s*'(\[.*?\])'", ci).group(1)
    assert min(tuple(map(int, v.split("."))) for v in re.findall(r'"(\d+\.\d+)"', versions)) == (major, minor)


def test_version_is_a_dev_release_while_unreleased_has_entries():
    version = PYPROJECT["project"]["version"]
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    released = re.findall(r"^## \[(\d+\.\d+\.\d+)\]", changelog, re.M)
    assert version not in released, f"{version} は CHANGELOG でリリース済み. 次の .dev 版にする"
    unreleased = changelog.split("## [Unreleased]", 1)[1].split("\n## [", 1)[0]
    if unreleased.strip():
        assert ".dev" in version, "Unreleased に変更があるので、版は開発版（.devN）にする"


def test_readme_limits_table_matches_parser_defaults():
    parser = cli.build_parser()
    defaults = {a.option_strings[-1]: a.default for a in parser._actions if a.option_strings}
    rows = re.findall(r"^\| `(--[a-z-]+) [A-Z]+` \| ([0-9.]+) \|", README, re.M)
    assert {r[0] for r in rows} == {"--max-cells", "--max-unzipped-mb", "--max-zip-ratio", "--loss-threshold"}
    for opt, value in rows:
        assert float(value) == float(defaults[opt]), opt


def test_output_spec_lists_every_rule_id_of_doc2md():
    src = ROOT / "src" / "doc_to_markdown"
    files = [p for p in src.rglob("*.py") if p.name != "figspec.py"]  # figspec は別のツール（rule_id は fig.*）
    used = set()
    for p in files:
        used |= set(re.findall(r'"((?:convert|loss|plan)\.[a-z_]+)"', p.read_text(encoding="utf-8")))
    used |= {f"loss.{name}" for name, _ in lossaudit.STRUCTURE_FIELDS}
    documented = set(re.findall(r"`((?:convert|loss|plan)\.[a-z_]+)`", SPEC))
    assert used - documented == set(), "OUTPUT_SPEC.md に無い rule_id"
    assert documented - used == set(), "OUTPUT_SPEC.md にあるがコードで使っていない rule_id"


def test_output_spec_exit_codes_match_common_exit_codes():
    section = SPEC.split("### 終了コード", 1)[1].split("\n### ", 1)[0]
    documented = {int(c) for c in re.findall(r"^\| (\d+) \|", section, re.M)}
    assert documented == {int(c) for c in ExitCode}
