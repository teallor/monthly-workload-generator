from __future__ import annotations

from dataclasses import asdict, dataclass


@dataclass
class CourseRecord:
    source_file: str
    date: str
    start_time: str
    end_time: str
    course_name: str
    teacher: str
    hours: float | int | None
    project: str
    audience: str = ""
    location: str = ""
    context: str = ""
    category: str = ""
    subcategory: str = ""
    confidence: float = 0.9
    needs_confirmation: bool = False
    is_delivery: bool = False
    is_live: bool = False
    target_row: int | None = None
    target_cells: str = ""
    included: bool = True
    exclusion_reason: str = ""
    status: str = "待分类"
    weekday: str = ""
    write_module: str = ""
    teacher_match_type: str = ""
    confirmation_note: str = ""

    def to_dict(self) -> dict:
        return asdict(self)
