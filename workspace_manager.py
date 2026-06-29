from __future__ import annotations

import json
import re
import shutil
import uuid
from pathlib import Path


SUPPORTED_MATERIALS = {".doc", ".docx", ".pdf", ".xls", ".xlsx", ".jpg", ".jpeg", ".png"}
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png"}


def format_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def unique_destination(folder: Path, name: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    candidate = folder / name
    stem, suffix = candidate.stem, candidate.suffix
    number = 1
    while candidate.exists():
        candidate = folder / f"{stem}_{number}{suffix}"
        number += 1
    return candidate


def copy_without_overwrite(source: Path, folder: Path) -> Path:
    destination = unique_destination(folder, source.name)
    shutil.copy2(source, destination)
    if source.suffix.lower() in IMAGE_SUFFIXES:
        sidecar = source.with_name(source.name + ".records.json")
        if sidecar.exists():
            shutil.copy2(sidecar, destination.with_name(destination.name + ".records.json"))
    return destination


def material_item(path: Path, *, original_path: Path | None = None, status: str = "待解析") -> dict:
    exists = path.exists()
    return {
        "name": path.name,
        "type": path.suffix.lower().lstrip(".").upper() or "未知",
        "size": format_size(path.stat().st_size) if exists else "-",
        "status": status if exists else "路径不存在",
        "path": str(path.resolve()) if exists else str(path),
        "original_path": str((original_path or path).resolve()) if (original_path or path).exists() else str(original_path or path),
        "supported": path.suffix.lower() in SUPPORTED_MATERIALS,
    }


def import_material_files(sources: list[Path], destination: Path, existing: list[dict]) -> tuple[list[dict], list[str]]:
    known = {str(Path(item.get("original_path", item["path"])).resolve()).lower() for item in existing}
    added: list[dict] = []
    skipped: list[str] = []
    for source in sources:
        source = source.resolve()
        key = str(source).lower()
        if key in known:
            skipped.append(f"{source.name}：已在材料列表中")
            continue
        known.add(key)
        if source.suffix.lower() not in SUPPORTED_MATERIALS:
            added.append(material_item(source, original_path=source, status="不支持"))
            continue
        copied = copy_without_overwrite(source, destination)
        added.append(material_item(copied, original_path=source))
    return added, skipped


def collect_folder_files(folder: Path) -> list[Path]:
    return sorted(path for path in folder.rglob("*") if path.is_file() and not path.name.startswith("~$"))


def prepare_session(workspace: Path, target_month: str, materials: list[dict]) -> Path:
    session = workspace / "sessions" / target_month / uuid.uuid4().hex
    session.mkdir(parents=True, exist_ok=False)
    for item in materials:
        source = Path(item["path"])
        if not item.get("supported", source.suffix.lower() in SUPPORTED_MATERIALS) or not source.exists():
            continue
        copy_without_overwrite(source, session)
    return session


def template_month(name: str) -> str:
    match = re.search(r"(20\d{2})年\s*(\d{1,2})月", name)
    return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}" if match else "未识别"


def load_settings(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(path: Path, settings: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)
