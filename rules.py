from __future__ import annotations

import math
import re
from datetime import datetime

ASSESSMENT_WORDS = (
    "等级工", "司机", "技能竞赛", "技能等级认定", "职业技能等级",
    "高级工", "中级工", "初级工", "技师", "高级技师", "应知", "应会",
)
THEORY_WORDS = ("理论", "理论知识", "应知", "知识辅导", "总复习")
PRACTICE_WORDS = (
    "实训", "实操", "操作", "模拟器", "排故", "维护保养", "吊具维护", "应会", "设备操作",
)


def classify(project: str, course: str) -> tuple[str, str, bool]:
    text = f"{project} {course}"
    if not any(word in text for word in ASSESSMENT_WORDS):
        return "培训", "", False
    theory = any(word in text for word in THEORY_WORDS)
    practice = any(word in text for word in PRACTICE_WORDS)
    if theory and not practice:
        return "考核", "理论", False
    if practice and not theory:
        return "考核", "实训", False
    return "考核", "需确认", True


def assessment_group_key(project: str, subcategory: str) -> str:
    return f"{project.strip()}|{subcategory.strip()}"


def assessment_class_name(project: str, subcategory: str) -> str:
    period = re.search(r"第?\s*(\d+)\s*期", project)
    level = re.search(r"(初级工|中级工|高级工|技师|高级技师)", project)
    if period and "电动港机装卸机械司机" in project:
        level_text = level.group(1).replace("工", "") if level else ""
        return f"{period.group(1)}期电司{level_text}{subcategory}"
    return f"{project}{subcategory}"


def hours_from_time(time_text: str) -> float:
    """45 分钟为 1 课时；仅写全天时计 8 课时。"""
    if "全天" in time_text or not re.search(r"\d{1,2}:\d{2}", time_text):
        return 8
    times = re.findall(r"(\d{1,2}):(\d{2})", time_text)
    if len(times) < 2:
        return 8
    start = datetime(2000, 1, 1, int(times[0][0]), int(times[0][1]))
    end = datetime(2000, 1, 1, int(times[1][0]), int(times[1][1]))
    minutes = (end - start).total_seconds() / 60
    if minutes < 0:
        minutes += 24 * 60
    value = minutes / 45
    return int(value) if value.is_integer() else round(value, 2)
