"""doc-to-markdown — 各種ドキュメント（Office/PDF/図）を Markdown + 抽出図に変換する単機能ツール.

検証（項目間の整合性チェック）は責務外。field-validation ツールへ別プロセスで渡す。
"""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("doc-to-markdown")
except PackageNotFoundError:
    __version__ = "0.0.0+local"
