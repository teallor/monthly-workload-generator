from __future__ import annotations

from pathlib import Path

from .excel_parser import parse_excel
from .image_parser import parse_image
from .pdf_parser import parse_pdf
from .word_parser import parse_doc, parse_docx

SUPPORTED = {".doc", ".docx", ".pdf", ".xls", ".xlsx", ".jpg", ".jpeg", ".png"}


def parse_sources(folder: Path, year: int, template: Path, output_dir: Path, enable_ocr: bool, teacher_name: str):
    records, warnings, scanned = [], [], []
    handlers = {".doc": parse_doc, ".docx": parse_docx, ".pdf": parse_pdf, ".xls": parse_excel, ".xlsx": parse_excel}
    template_resolved = template.resolve()
    output_resolved = output_dir.resolve()
    for path in sorted(folder.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in SUPPORTED or path.name.startswith("~$"):
            continue
        resolved = path.resolve()
        if resolved == template_resolved or (output_resolved != folder.resolve() and output_resolved in resolved.parents):
            continue
        if path.suffix.lower() in {".xls", ".xlsx"} and "工作量表" in path.name:
            continue
        scanned.append(path)
        try:
            if path.suffix.lower() in {".jpg", ".jpeg", ".png"}:
                parsed, warning = parse_image(path, year, teacher_name, enable_ocr)
                records.extend(parsed)
                if warning:
                    warnings.append(warning)
            else:
                records.extend(handlers[path.suffix.lower()](path, year))
        except Exception as exc:
            warnings.append(f"{path.name}: 解析失败：{exc}")
    return records, warnings, scanned
