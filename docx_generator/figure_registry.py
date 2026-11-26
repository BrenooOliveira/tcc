"""Leitura e conferência do figure_registry.md da monografia."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


_FILE_RE = re.compile(
    r'`([^`\n]+\.(?:png|jpe?g|svg|webp|gif|mermaid))`',
    re.IGNORECASE,
)


@dataclass
class FigureRecord:
    arquivo: str
    titulo: str = ''
    secao: str = ''
    citacao: str = ''
    alt: str = ''
    status: str = ''


class FigureRegistry:
    """Indexa arquivos citados no registro e resolve caminhos locais."""

    def __init__(self, path: str | Path | None):
        self.path = Path(path).resolve() if path else None
        self.records: list[FigureRecord] = []
        self.assets: set[str] = set()
        if self.path and self.path.exists():
            self._parse(self.path.read_text(encoding='utf-8'))

    @property
    def loaded(self) -> bool:
        return bool(self.path and self.path.exists())

    def contains(self, url: str) -> bool:
        if not url or url.startswith('http'):
            return True
        normalized = url.replace('\\', '/').lstrip('./')
        name = Path(normalized).name.lower()
        for asset in self.assets:
            asset_name = Path(asset).name.lower()
            if name == asset_name or normalized.endswith(asset) or asset.endswith(name):
                return True
        return False

    def resolve(self, arquivo: str, markdown_dir: Path) -> Path | None:
        raw = arquivo.strip().strip('`')
        if not raw or raw in {'—', '-', '–'}:
            return None
        name = Path(raw).name
        roots = [markdown_dir, markdown_dir.parent]
        candidates = []
        for root in roots:
            candidates.append(root / raw)
            candidates.append(root / 'assets' / name)
            candidates.append(root / 'assets' / raw)
        seen = set()
        for candidate in candidates:
            key = str(candidate)
            if key in seen:
                continue
            seen.add(key)
            if candidate.exists():
                return candidate.resolve()
        return None

    def _parse(self, text: str):
        for match in _FILE_RE.findall(text):
            self.assets.add(match.strip().replace('\\', '/'))

        current_title = ''
        bucket: dict[str, str] = {}

        def flush():
            arquivo = bucket.get('arquivo', '').strip()
            if not arquivo or arquivo in {'—', '-', '–'}:
                bucket.clear()
                return
            record = FigureRecord(
                arquivo=arquivo,
                titulo=current_title,
                secao=bucket.get('secao', ''),
                citacao=bucket.get('citacao', ''),
                alt=bucket.get('alt', ''),
                status=bucket.get('status', ''),
            )
            self.records.append(record)
            self.assets.add(arquivo.replace('\\', '/'))
            bucket.clear()

        for line in text.splitlines():
            heading = re.match(r'^###\s+(.+)$', line.strip())
            if heading:
                flush()
                current_title = heading.group(1).strip()
                continue
            cells = _split_row(line)
            if len(cells) < 2:
                continue
            key = _field_key(cells[0])
            value = _clean_cell(cells[1])
            if key:
                bucket[key] = value
        flush()


def _split_row(line: str) -> list[str]:
    stripped = line.strip()
    if not stripped.startswith('|'):
        return []
    return [cell.strip() for cell in stripped.strip('|').split('|')]


def _field_key(label: str) -> str:
    plain = re.sub(r'\*+', '', label).strip().lower()
    if plain.startswith('arquivo'):
        return 'arquivo'
    if plain.startswith('seção') or plain.startswith('secao'):
        return 'secao'
    if 'abnt' in plain or plain.startswith('citação') or plain.startswith('citacao'):
        return 'citacao'
    if plain.startswith('alt'):
        return 'alt'
    if plain.startswith('status'):
        return 'status'
    return ''


def _clean_cell(value: str) -> str:
    value = value.strip()
    if value.startswith('`') and value.endswith('`') and len(value) >= 2:
        return value[1:-1].strip()
    return re.sub(r'`([^`]*)`', r'\1', value).strip()
