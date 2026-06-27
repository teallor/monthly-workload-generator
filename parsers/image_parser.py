from __future__ import annotations

import shutil
import re
import json
from pathlib import Path

from rules import hours_from_time
from .common import CourseRecord


def parse_image(path: Path, year: int, teacher_name: str, enabled: bool = False):
    sidecar = Path(str(path) + ".records.json")
    if sidecar.exists():
        data = json.loads(sidecar.read_text(encoding="utf-8"))
        records = []
        for item in data.get("records", []):
            records.append(CourseRecord(
                source_file=path.name,
                date=item.get("date", ""), start_time=item.get("start_time", ""), end_time=item.get("end_time", ""),
                course_name=item.get("course_name", ""), teacher=item.get("teacher", ""), hours=item.get("hours"),
                project=item.get("project", path.stem), audience=item.get("audience", ""),
                location=item.get("location", ""), context=item.get("context", "照片课表人工校核记录"),
                confidence=float(item.get("confidence", 0.99)),
                needs_confirmation=bool(item.get("needs_confirmation", False)),
                confirmation_note=item.get("confirmation_note", ""),
            ))
        return records, f"{path.name}: 已读取人工校核侧车记录 {sidecar.name}，共 {len(records)} 条。"
    # OCR 是可选能力。没有本地 Tesseract 时保留清晰告警，不阻断 Word/PDF 主流程。
    if not enabled:
        return [], f"{path.name}: enable_ocr=false，图片未自动提取。"
    if not shutil.which("tesseract"):
        return [], f"{path.name}: 已启用OCR但未安装 Tesseract，图片未自动提取。"
    try:
        import pytesseract
        from PIL import Image
        text = pytesseract.image_to_string(Image.open(path), lang="chi_sim+eng")
    except Exception as exc:
        return [], f"{path.name}: OCR失败：{exc}"
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]
    records = []
    for index, line in enumerate(lines):
        if teacher_name not in line:
            continue
        context = " | ".join(lines[max(0, index - 8): index + 2])
        date_match = re.search(r"(?:(20\d{2})年)?\s*(\d{1,2})月\s*(\d{1,2})日", context)
        times = re.findall(r"\d{1,2}:\d{2}", context)
        course = lines[index - 1] if index else ""
        records.append(CourseRecord(
            source_file=path.name,
            date=(f"{int(date_match.group(1) or year):04d}-{int(date_match.group(2)):02d}-{int(date_match.group(3)):02d}" if date_match else ""),
            start_time=times[-2] if len(times) >= 2 else "", end_time=times[-1] if len(times) >= 2 else "",
            course_name=course, teacher=teacher_name,
            hours=hours_from_time("-".join(times[-2:])) if len(times) >= 2 else 8,
            project=path.stem, context=context, confidence=0.45, needs_confirmation=True,
        ))
    warning = f"{path.name}: OCR完成，识别到 {len(records)} 条低置信度候选，写入前必须人工确认。"
    return records, warning
