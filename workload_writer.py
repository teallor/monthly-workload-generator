from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
from dataclasses import replace
from pathlib import Path
from rules import assessment_class_name, assessment_group_key


def compact_class_name(value: str) -> str:
    """Remove date/organization prefixes when a class name would otherwise clip in the fixed template."""
    text = re.sub(r"^20\d{2}年", "", value).strip()
    text = re.sub(r"^.*?(?:集团|公司)(?=.{4,}$)", "", text, count=1).strip()
    return text or value


def _bridge(*arguments: str):
    script = Path(__file__).resolve().parent / "office_bridge.ps1"
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), *arguments],
        check=False, capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    if result.returncode:
        detail = (result.stderr or result.stdout or "Office COM bridge failed").strip()
        raise RuntimeError(detail)
    return result


def locate_template(folder: Path, explicit: Path | None, keyword: str, teacher: str) -> Path:
    if explicit:
        if explicit.is_absolute():
            path = explicit
        elif (Path.cwd() / explicit).exists():
            path = Path.cwd() / explicit
        else:
            path = folder / explicit
        path = path.resolve()
        if not path.exists() or path.suffix.lower() not in {".xls", ".xlsx"}:
            raise RuntimeError(f"指定模板不存在或不是 .xls/.xlsx：{path}")
        return path
    candidates = sorted(
        p.resolve() for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"}
        and keyword in p.name and teacher in p.name and not p.name.startswith("~$")
    )
    if not candidates:
        raise RuntimeError(f"未找到文件名同时包含“{keyword}”和“{teacher}”的 .xls/.xlsx 模板")
    if len(candidates) > 1:
        listing = "\n".join(f"- {p.name}" for p in candidates)
        raise RuntimeError(f"找到多个模板，请使用 --template 明确选择：\n{listing}")
    return candidates[0]


def inspect_template(path: Path) -> dict:
    result = _bridge("-Action", "inspect", "-Path", str(path.resolve()))
    try:
        data = json.loads(result.stdout.strip().splitlines()[-1])
    except Exception as exc:
        raise RuntimeError("无法读取模板定位结果") from exc
    required = {
        "sheet", "title", "title_cell", "header_row", "data_start_row", "data_end_row",
        "training_start_col", "training_end_col", "assessment_start_col", "assessment_end_col",
        "assessment_slots",
    }
    if not required.issubset(data) or data["data_end_row"] < data["data_start_row"]:
        raise RuntimeError("无法稳定识别教学模块、月份单元格或可用行")
    return data


def build_target_title(source_title: str, year: int, month: int) -> str:
    replacement = f"{year}年{month}月"
    updated, count = re.subn(r"20\d{2}年\s*\d{1,2}月", replacement, source_title, count=1)
    if not count:
        raise RuntimeError("模板标题中找不到年份月份，无法安全修改表头")
    return updated


def build_output_name(template_name: str, year: int, month: int) -> str:
    replacement = f"{year}年{month}月"
    updated, count = re.subn(r"20\d{2}年\s*\d{1,2}月", replacement, template_name, count=1)
    if count:
        return updated
    path = Path(template_name)
    return f"{path.stem}（{replacement}）{path.suffix}"


def write_workbook(template: Path, output: Path, title: str, records: list, layout: dict) -> dict:
    capacity = layout["data_end_row"] - layout["data_start_row"] + 1
    training_count = sum(r.category == "培训" for r in records)
    assessment_count = len({assessment_group_key(r.project, r.subcategory) for r in records if r.category == "考核"})
    if training_count > capacity or assessment_count > len(layout["assessment_slots"]):
        required = max(training_count, assessment_count)
        raise RuntimeError(
            f"教学模块空白槽位不足，当前需要培训 {training_count} 条、考核 {assessment_count} 组；"
            f"模板仅有培训 {capacity} 行、考核 {len(layout['assessment_slots'])} 个槽位，"
            "是否允许复制格式扩展行？"
        )
    output.parent.mkdir(parents=True, exist_ok=True)
    if any(record.hours is None or record.needs_confirmation for record in records):
        raise RuntimeError("存在课时或分类未确认的课程，禁止写入 Excel")
    grouped_records = [record for record in records if record.category == "培训"]
    assessment_groups = {}
    for record in (r for r in records if r.category == "考核"):
        key = assessment_group_key(record.project, record.subcategory)
        if key not in assessment_groups:
            assessment_groups[key] = replace(record, hours=0, needs_confirmation=False)
        assessment_groups[key].hours += record.hours
    grouped_records.extend(assessment_groups.values())
    with tempfile.TemporaryDirectory(prefix="workload_write_") as folder:
        payload = Path(folder) / "payload.json"
        record_payload = []
        for record in grouped_records:
            item = record.to_dict()
            item["is_training"] = record.category == "培训"
            item["sheet_class_name"] = (
                assessment_class_name(record.project, record.subcategory)
                if record.category == "考核" else compact_class_name(record.project)
            )
            record_payload.append(item)
        payload.write_text(json.dumps({
            "title": title, "layout": layout, "records": record_payload,
        }, ensure_ascii=False), encoding="utf-8-sig")
        _bridge(
            "-Action", "write", "-Path", str(template.resolve()), "-Output", str(output.resolve()),
            "-Payload", str(payload),
        )
    return {
        "output": str(output),
        "changed": [layout["title_cell"], f"教学数据行{layout['data_start_row']}:{layout['data_end_row']}"],
        "untouched": "除月份标题及教学模块外未写入",
    }
