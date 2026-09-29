"""形式ごとの変換アダプタ.

各アダプタは `convert(path: Path, assets: AssetWriter, opts: Options) -> Document` を実装する。
新しい形式を足すときはモジュールを追加し、ADAPTERS に拡張子を登録する。
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from ..assets import AssetWriter
from ..model import Document, Options
from . import docx_adapter, drawio_adapter, pdf_adapter, pptx_adapter, xlsx_adapter

Adapter = Callable[[Path, AssetWriter, Options], Document]

ADAPTERS: dict[str, Adapter] = {
    ".docx": docx_adapter.convert,
    ".xlsx": xlsx_adapter.convert,
    ".xlsm": xlsx_adapter.convert,
    ".pptx": pptx_adapter.convert,
    ".pdf": pdf_adapter.convert,
    ".drawio": drawio_adapter.convert,
}

# 直接読めない旧形式・他形式は LibreOffice で PDF 化してから pdf_adapter に渡す
VIA_PDF_ONLY = {".doc", ".xls", ".ppt", ".odt", ".ods", ".odp", ".rtf"}
OFFICE = {".docx", ".xlsx", ".xlsm", ".pptx"} | VIA_PDF_ONLY

SUPPORTED = set(ADAPTERS) | VIA_PDF_ONLY
