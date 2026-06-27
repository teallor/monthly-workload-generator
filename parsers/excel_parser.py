from __future__ import annotations

import json
import re
import subprocess
import tempfile
from pathlib import Path

from rules import hours_from_time
from .common import CourseRecord


def _clean(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _date(value: str, year: int) -> str:
    text = _clean(value)
    full = re.search(r"(20\d{2})[年/\-.](\d{1,2})[月/\-.](\d{1,2})", text)
    if full:
        return f"{int(full.group(1)):04d}-{int(full.group(2)):02d}-{int(full.group(3)):02d}"
    short = re.search(r"(\d{1,2})月(\d{1,2})日", text)
    if short:
        return f"{year:04d}-{int(short.group(1)):02d}-{int(short.group(2)):02d}"
    return ""


def _extract(path: Path) -> list[dict]:
    bridge = Path(__file__).resolve().parents[1] / "office_bridge.ps1"
    with tempfile.TemporaryDirectory(prefix="workload_excel_") as folder:
        output = Path(folder) / "cells.json"
        subprocess.run([
            "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(bridge),
            "-Action", "extract-excel", "-Path", str(path.resolve()), "-Output", str(output),
        ], check=True, capture_output=True, text=True, encoding="utf-8", errors="replace")
        return json.loads(output.read_text(encoding="utf-8-sig"))


def parse_excel(path: Path, year: int) -> list[CourseRecord]:
    records: list[CourseRecord] = []
    for sheet in _extract(path):
        rows = sheet.get("rows", [])
        header_index = None
        for i, row in enumerate(rows):
            labels = [_clean(x) for x in row]
            if any("日期" in x for x in labels) and any("课程" in x or "项目" in x for x in labels) and any("教师" in x or "培训师" in x or "授课老师" in x for x in labels):
                header_index = i
                break
        if header_index is None:
            continue
        headers = [_clean(x) for x in rows[header_index]]
        def col(*words):
            return next((i for i, h in enumerate(headers) if any(w in h for w in words)), None)
        date_col, time_col = col("日期"), col("时间")
        course_col, teacher_col = col("课程", "项目"), col("教师", "培训师", "授课老师")
        hours_col, location_col = col("课时"), col("地点", "场所")
        audience_col = col("对象", "班级", "单位")
        inherited_date = ""
        for row in rows[header_index + 1:]:
            values = [_clean(x) for x in row]
            raw_date = values[date_col] if date_col is not None and date_col < len(values) else ""
            inherited_date = _date(raw_date, year) or inherited_date
            course = values[course_col] if course_col is not None and course_col < len(values) else ""
            teacher = values[teacher_col] if teacher_col is not None and teacher_col < len(values) else ""
            if not course or not teacher:
                continue
            time_text = values[time_col] if time_col is not None and time_col < len(values) else ""
            found = re.findall(r"\d{1,2}:\d{2}", time_text)
            explicit = values[hours_col] if hours_col is not None and hours_col < len(values) else ""
            try:
                hours = float(explicit)
            except ValueError:
                hours = hours_from_time(time_text)
            records.append(CourseRecord(
                source_file=path.name, date=inherited_date,
                start_time=found[0] if found else time_text, end_time=found[1] if len(found) > 1 else "",
                course_name=course, teacher=teacher, hours=int(hours) if float(hours).is_integer() else hours,
                project=path.stem,
                audience=values[audience_col] if audience_col is not None and audience_col < len(values) else "",
                location=values[location_col] if location_col is not None and location_col < len(values) else "",
                context=" | ".join(values), confidence=0.88,
            ))
    return records
