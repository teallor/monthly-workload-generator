from __future__ import annotations

import re
from pathlib import Path

import pdfplumber

from rules import hours_from_time
from .common import CourseRecord


def _clean(value) -> str:
    return re.sub(r"\s+", "", str(value or "")).strip()


def _date(value: str, year: int) -> str:
    match = re.search(r"(\d{1,2})月(\d{1,2})日", _clean(value))
    return f"{year:04d}-{int(match.group(1)):02d}-{int(match.group(2)):02d}" if match else ""


def parse_pdf(path: Path, year: int) -> list[CourseRecord]:
    records: list[CourseRecord] = []
    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        years = re.findall(r"(20\d{2})\s*年", text + " " + path.name)
        document_year = int(years[0]) if years else year
        project_match = re.search(r"《([^》]+)》", text)
        project = project_match.group(1) if project_match else path.stem
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table:
                    continue
                headers = [_clean(cell) for cell in table[0]]
                if not any("日期" in h for h in headers) or not any("培训师" in h or "教师" in h for h in headers):
                    continue
                def index_of(*words):
                    return next((i for i, h in enumerate(headers) if any(w in h for w in words)), None)
                date_col = index_of("日期")
                time_col = index_of("时间")
                course_col = index_of("项目", "课程")
                teacher_col = index_of("培训师", "授课老师", "教师")
                location_col = index_of("场所", "地点")
                hours_col = index_of("课时")
                if None in (date_col, time_col, course_col, teacher_col):
                    continue
                inherited_date = inherited_time = ""
                for row in table[1:]:
                    values = [_clean(cell) for cell in row]
                    if date_col < len(values) and _date(values[date_col], document_year):
                        inherited_date = _date(values[date_col], document_year)
                    if time_col < len(values) and values[time_col]:
                        inherited_time = values[time_col]
                    course = values[course_col] if course_col < len(values) else ""
                    teacher = values[teacher_col] if teacher_col < len(values) else ""
                    if not course or not teacher:
                        continue
                    found_times = re.findall(r"\d{1,2}:\d{2}", inherited_time)
                    explicit = values[hours_col] if hours_col is not None and hours_col < len(values) else ""
                    try:
                        hours = float(explicit)
                    except ValueError:
                        hours = hours_from_time(inherited_time)
                    records.append(CourseRecord(
                        source_file=path.name, date=inherited_date,
                        start_time=found_times[0] if found_times else inherited_time,
                        end_time=found_times[1] if len(found_times) > 1 else "",
                        course_name=course, teacher=teacher, hours=int(hours) if float(hours).is_integer() else hours,
                        project=project,
                        location=values[location_col] if location_col is not None and location_col < len(values) else "",
                        context=" | ".join(values), confidence=0.96,
                    ))
    return records
