"""Small in-memory logger that can also be written back to Excel."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class LogEntry:
    timestamp: str
    level: str
    message: str


@dataclass
class ProcessingLogger:
    entries: list[LogEntry] = field(default_factory=list)

    def _add(self, level: str, message: str) -> None:
        entry = LogEntry(
            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            level=level,
            message=message,
        )
        self.entries.append(entry)
        print(f"[{entry.timestamp}] [{entry.level}] {entry.message}")

    def info(self, message: str) -> None:
        self._add("INFO", message)

    def warning(self, message: str) -> None:
        self._add("WARNING", message)

    def error(self, message: str) -> None:
        self._add("ERROR", message)

