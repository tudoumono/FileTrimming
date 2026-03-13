"""追跡メタデータ

各ファイルの処理結果を追跡するためのメタデータ。
中間成果物 JSON に埋め込み、最終 Markdown には含めない
（Dify は YAML front matter を認識しないため）。
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime
from enum import Enum
from typing import Any


class ProcessStatus(str, Enum):
    SUCCESS = "success"
    SKIPPED = "skipped"
    ERROR = "error"
    WARNING = "warning"


@dataclass
class FileMetadata:
    """1ファイルの追跡情報"""
    source_path: str          # 元ファイルの相対パス (input/ からの相対)
    source_ext: str           # 元の拡張子 (.doc, .rtf, .docx 等)
    source_size_bytes: int = 0
    normalized_from: str = "" # 正規化元の拡張子 (例: .rtf → .docx なら ".rtf")
    doc_role_guess: str = ""  # "spec_body", "change_history", "mixed", "unknown"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class StepResult:
    """1ステップの処理結果（ログ用）"""
    file_path: str
    step: str                 # "normalize", "extract", "transform"
    status: ProcessStatus
    message: str = ""
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())
    duration_sec: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["status"] = self.status.value
        return d


@dataclass
class ExtractedFileRecord:
    """Step2 出力の JSON ファイルに含める全情報"""
    metadata: FileMetadata
    document: dict[str, Any]  # IntermediateDocument.to_dict() の結果

    def to_dict(self) -> dict[str, Any]:
        return {
            "metadata": self.metadata.to_dict(),
            "document": self.document,
        }
