"""このリポジトリを仮想環境にインストールする（setup.cmd / setup.sh から呼ばれる）.

    python scripts/install_deps.py [--dev]

- 隣のフォルダに devtools-common があれば、それを editable で入れ、pyproject.toml の
  `devtools-common @ git+…` 依存の代わりに使う（GitHub の認証が無くてもセットアップできる）
- 無ければ通常どおり pip install（devtools-common は GitHub から取得）
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
COMMON = REPO.parent / "devtools-common"
COMMON_NAME = "devtools-common"


def pip(*args: str) -> None:
    cmd = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", *args]
    print("$ pip install", " ".join(args), flush=True)
    subprocess.run(cmd, check=True)


def dependencies(extras: list[str]) -> list[str]:
    text = (REPO / "pyproject.toml").read_text(encoding="utf-8")
    try:
        import tomllib

        project = tomllib.loads(text)["project"]
        deps = list(project.get("dependencies", []))
        for e in extras:
            deps += project.get("optional-dependencies", {}).get(e, [])
    except ModuleNotFoundError:  # Python 3.10
        m = re.search(r"^dependencies\s*=\s*\[(.*?)\]", text, re.S | re.M)
        deps = re.findall(r'"([^"]+)"', m.group(1)) if m else []
        for e in extras:
            m = re.search(rf"^{e}\s*=\s*\[(.*?)\]", text, re.S | re.M)
            deps += re.findall(r'"([^"]+)"', m.group(1)) if m else []
    return deps


def main() -> int:
    dev = "--dev" in sys.argv[1:]
    extras = ["dev"] if dev else []
    is_common = REPO.name == COMMON_NAME or (REPO / "src" / "devtools_common").is_dir()
    if not is_common and (COMMON / "pyproject.toml").is_file():
        print(f"== 隣の {COMMON_NAME} を使います: {COMMON}")
        pip("-e", str(COMMON))
        third_party = [d for d in dependencies(extras) if not d.replace("_", "-").startswith(COMMON_NAME)]
        if third_party:
            pip(*third_party)
        pip("--no-deps", "-e", str(REPO))
    else:
        pip("-e", f"{REPO}[dev]" if dev else str(REPO))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except subprocess.CalledProcessError:
        print(
            "\nインストールに失敗しました。ネットワーク（社内プロキシ）と、devtools-common を GitHub から"
            "取得できるか確認してください。取得できない場合は、このフォルダと同じ場所に devtools-common を"
            " clone してから再実行してください。",
            file=sys.stderr,
        )
        sys.exit(1)
