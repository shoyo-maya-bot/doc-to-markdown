"""抽出図・画像の書き出し（決定的な連番ファイル名）."""

from __future__ import annotations

import hashlib
from pathlib import Path

_EXT_BY_TYPE = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/svg+xml": "svg",
    "image/x-emf": "emf",
    "image/x-wmf": "wmf",
    "image/tiff": "tiff",
}


class AssetWriter:
    """`<out>/<stem>.assets/<stem>-NN.<ext>` に書き出し、Markdown からの相対パスを返す.

    同一内容の画像は 1 回だけ書き出す（重複排除）。番号は出現順なので同入力なら同出力になる。
    """

    def __init__(self, out_dir: Path, stem: str):
        self.dir_name = f"{stem}.assets"
        self.dir = out_dir / self.dir_name
        self.stem = stem
        self._seen: dict[str, str] = {}
        self.count = 0

    def write(self, data: bytes, *, content_type: str = "", ext: str = "") -> str:
        digest = hashlib.sha256(data).hexdigest()
        if digest in self._seen:
            return self._seen[digest]
        ext = (ext or _EXT_BY_TYPE.get(content_type, "bin")).lstrip(".").lower()
        self.count += 1
        name = f"{self.stem}-{self.count:02d}.{ext}"
        self.dir.mkdir(parents=True, exist_ok=True)
        (self.dir / name).write_bytes(data)
        rel = f"{self.dir_name}/{name}"
        self._seen[digest] = rel
        return rel

    def copy_file(self, src: Path, name: str) -> str:
        self.dir.mkdir(parents=True, exist_ok=True)
        (self.dir / name).write_bytes(src.read_bytes())
        return f"{self.dir_name}/{name}"
