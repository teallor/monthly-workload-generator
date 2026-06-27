from __future__ import annotations

import csv
import json
from datetime import date
from pathlib import Path

from rules import assessment_group_key

WEEKDAYS = "一二三四五六日"


def _col_letter(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result

PREVIEW_FIELDS = [
    "目标月份", "日期", "星期", "开始时间", "结束时间", "课程名称", "教师", "课时",
    "来源文件", "项目或班级", "培训对象", "地点", "分类", "写入模块", "子分类",
    "教师命中方式", "预计写入单元格", "置信度", "是否需要人工确认", "确认说明", "原文上下文",
]
EXCLUDED_FIELDS = PREVIEW_FIELDS[:9] + ["状态", "排除原因"] + PREVIEW_FIELDS[9:]


def _weekday(value: str) -> str:
    try:
        parsed = date.fromisoformat(value)
        return "星期" + WEEKDAYS[parsed.weekday()]
    except ValueError:
        return ""


def _row(record, target_month: str) -> dict:
    return {
        "目标月份": target_month,
        "日期": record.date,
        "星期": _weekday(record.date),
        "开始时间": record.start_time,
        "结束时间": record.end_time,
        "课程名称": record.course_name,
        "教师": record.teacher,
        "课时": "需确认" if record.hours is None else record.hours,
        "来源文件": record.source_file,
        "项目或班级": record.project,
        "培训对象": record.audience,
        "地点": record.location,
        "分类": record.category,
        "写入模块": record.write_module or record.category,
        "子分类": record.subcategory,
        "教师命中方式": record.teacher_match_type,
        "预计写入单元格": record.target_cells,
        "置信度": f"{record.confidence:.0%}",
        "是否需要人工确认": "是" if record.needs_confirmation else "否",
        "确认说明": record.confirmation_note,
        "状态": record.status,
        "排除原因": record.exclusion_reason,
        "原文上下文": record.context,
    }


def assign_targets(records: list, layout: dict) -> list[str]:
    warnings: list[str] = []
    start, end = layout["data_start_row"], layout["data_end_row"]
    capacity = end - start + 1
    training = [r for r in records if r.status == "待写入" and r.category == "培训"]
    assessment = [r for r in records if r.status == "待写入" and r.category == "考核"]
    for group, begin_col, end_col, label in (
        (training, _col_letter(layout["training_start_col"]), _col_letter(layout["training_end_col"]), "培训"),
    ):
        for index, record in enumerate(group):
            if index < capacity:
                row = start + index
                record.target_row = row
                record.target_cells = f"{begin_col}{row}:{end_col}{row}"
            else:
                record.needs_confirmation = True
                record.target_cells = "容量不足"
        if len(group) > capacity:
            warnings.append(
                f"教学模块{label}空白行不足，当前需要写入 {len(group)} 条，模板仅有 {capacity} 条空白行，"
                "是否允许复制格式扩展行？"
            )
    assessment_rows: dict[str, dict] = {}
    assessment_slots = layout.get("assessment_slots", [])
    for record in assessment:
        key = assessment_group_key(record.project, record.subcategory)
        if key not in assessment_rows:
            slot_index = len(assessment_rows)
            assessment_rows[key] = assessment_slots[slot_index] if slot_index < len(assessment_slots) else {}
        slot = assessment_rows[key]
        if slot:
            row = int(slot["start_row"])
            end_row = int(slot["end_row"])
            record.target_row = row
            record.target_cells = (
                f"{_col_letter(layout['assessment_start_col'])}{row}:"
                f"{_col_letter(layout['assessment_end_col'])}{end_row}"
            )
        else:
            record.needs_confirmation = True
            record.target_cells = "容量不足"
    if len(assessment_rows) > len(assessment_slots):
        warnings.append(
            f"教学模块考核空白槽位不足，当前需要写入 {len(assessment_rows)} 个考核汇总组，"
            f"模板仅有 {len(assessment_slots)} 个槽位，"
            "是否允许复制格式扩展行？"
        )
    return warnings


def write_previews(output_dir: Path, target_month: str, records: list, log_lines: list[str], template: Path):
    output_dir.mkdir(parents=True, exist_ok=True)
    preview_path = output_dir / "待写入预览表.csv"
    excluded_path = output_dir / "被排除课程.csv"
    log_path = output_dir / "解析日志.txt"
    snapshot_path = output_dir / "解析结果.json"
    included = [_row(r, target_month) for r in records if r.status == "待写入"]
    excluded = [_row(r, target_month) for r in records if r.status != "待写入"]
    with preview_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=PREVIEW_FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(included)
    with excluded_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=EXCLUDED_FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(excluded)
    log_path.write_text("\n".join(log_lines) + "\n", encoding="utf-8")
    snapshot_path.write_text(json.dumps({
        "target_month": target_month,
        "template": str(template.resolve()),
        "records": [record.to_dict() for record in records],
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    return preview_path, excluded_path, log_path, snapshot_path
