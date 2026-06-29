from __future__ import annotations

from dataclasses import asdict
from datetime import date
from pathlib import Path

from main import (
    classify_and_filter,
    load_preview_snapshot,
    parse_target_month,
)
from parsers import parse_sources
from preview import assign_targets, write_previews
from rules import assessment_class_name, assessment_group_key
from workload_writer import (
    build_output_name,
    build_target_title,
    inspect_template,
    locate_template,
    unique_output_path,
    write_workbook,
)


def resolve_path(value: str | Path, base: Path) -> Path:
    path = Path(value).expanduser()
    return (base / path).resolve() if not path.is_absolute() else path.resolve()


def build_assessment_summary(records: list) -> list[dict]:
    groups: dict[str, dict] = {}
    for record in records:
        if record.status != "待写入" or record.category != "考核":
            continue
        key = assessment_group_key(record.project, record.subcategory)
        if key not in groups:
            groups[key] = {
                "班级/项目": assessment_class_name(record.project, record.subcategory),
                "子分类": record.subcategory,
                "课时": 0,
                "明细条数": 0,
                "明细日期": [],
                "预计写入单元格": record.target_cells,
            }
        groups[key]["课时"] += record.hours or 0
        groups[key]["明细条数"] += 1
        if record.date and record.date not in groups[key]["明细日期"]:
            groups[key]["明细日期"].append(record.date)
    for group in groups.values():
        formatted = []
        for raw in group["明细日期"]:
            try:
                parsed = date.fromisoformat(raw)
                formatted.append(f"{parsed.month}月{parsed.day}日")
            except ValueError:
                formatted.append(raw)
        group["明细日期"] = "、".join(formatted)
    return list(groups.values())


def generate_preview(
    *,
    base_dir: Path,
    input_dir: str | Path,
    output_dir: str | Path,
    target_month_raw: str,
    teacher_name: str,
    teacher_aliases: list[str],
    template_path: str | Path | None = None,
    template_keyword: str = "工作量表",
    enable_ocr: bool = False,
    extra_records: list | None = None,
) -> dict:
    input_path = resolve_path(input_dir, base_dir)
    output_path = resolve_path(output_dir, base_dir)
    explicit = resolve_path(template_path, base_dir) if template_path else None
    template = locate_template(input_path, explicit, template_keyword, teacher_name)
    year, month, inference = parse_target_month(target_month_raw, template)
    target_month = f"{year:04d}-{month:02d}"
    month_output_dir = output_path / target_month
    layout = inspect_template(template)
    records, parser_warnings, scanned = parse_sources(
        input_path, year, month, template, output_path, enable_ocr, teacher_name
    )
    if extra_records:
        records.extend(extra_records)
        parser_warnings.append(f"已读取界面手工补录课程 {len(extra_records)} 条。")
    records = classify_and_filter(records, teacher_name, teacher_aliases, year, month)
    warnings = parser_warnings + assign_targets(records, layout)
    ocr_log_lines = []
    if enable_ocr:
        debug_dir = output_path / "ocr_debug"
        ocr_log_lines.append(f"OCR调试目录: {debug_dir}")
        for text_file in sorted(debug_dir.glob("*_ocr文本.txt")):
            try:
                raw_text = text_file.read_text(encoding="utf-8").strip()
                ocr_log_lines.extend([f"OCR原始文本 [{text_file.name}]:", raw_text])
            except OSError as exc:
                ocr_log_lines.append(f"OCR文本读取失败 [{text_file.name}]: {exc}")
    log_lines = [
        f"目标月份: {target_month}",
        f"月份推断: {inference}",
        f"教师: {teacher_name}",
        f"教师别名: {', '.join(teacher_aliases) or '无'}",
        f"模板: {template}",
        f"工作表: {layout['sheet']}",
        f"月份单元格: {layout['title_cell']}",
        f"教学数据行: {layout['data_start_row']}:{layout['data_end_row']}",
        f"输入目录: {input_path}",
        f"扫描依据文件数: {len(scanned)}",
        *[f"扫描: {path.name}" for path in scanned],
        f"识别课程总数: {len(records)}",
        f"待写入: {sum(record.status == '待写入' for record in records)}",
        f"培训: {sum(record.status == '待写入' and record.category == '培训' for record in records)}",
        f"考核: {sum(record.status == '待写入' and record.category == '考核' for record in records)}",
        f"参考: {sum(record.status == '参考' for record in records)}",
        f"排除: {sum(record.status == '排除' for record in records)}",
        *ocr_log_lines,
        *[f"需确认: {record.date} {record.course_name} - {record.confirmation_note or '请人工检查'}"
          for record in records if record.status == "待写入" and record.needs_confirmation],
        *[f"警告: {warning}" for warning in warnings],
    ]
    paths = write_previews(month_output_dir, target_month, records, log_lines, template)
    included = [record for record in records if record.status == "待写入"]
    excluded = [record for record in records if record.status != "待写入"]
    return {
        "target_month": target_month,
        "target_month_raw": target_month_raw,
        "teacher_name": teacher_name,
        "teacher_aliases": teacher_aliases,
        "enable_ocr": enable_ocr,
        "template": template,
        "input_dir": input_path,
        "output_dir": output_path,
        "month_output_dir": month_output_dir,
        "layout": layout,
        "records": records,
        "included": included,
        "excluded": excluded,
        "warnings": warnings,
        "scanned": scanned,
        "log_lines": log_lines,
        "paths": {
            "preview": paths[0],
            "excluded": paths[1],
            "log": paths[2],
            "snapshot": paths[3],
        },
        "assessment_summary": build_assessment_summary(records),
    }


def generate_excel(
    *,
    base_dir: Path,
    input_dir: str | Path,
    output_dir: str | Path,
    target_month_raw: str,
    teacher_name: str,
    template_path: str | Path | None = None,
    template_keyword: str = "工作量表",
) -> dict:
    input_path = resolve_path(input_dir, base_dir)
    output_path = resolve_path(output_dir, base_dir)
    explicit = resolve_path(template_path, base_dir) if template_path else None
    template = locate_template(input_path, explicit, template_keyword, teacher_name)
    year, month, _ = parse_target_month(target_month_raw, template)
    target_month = f"{year:04d}-{month:02d}"
    month_output_dir = output_path / target_month
    layout = inspect_template(template)
    included = load_preview_snapshot(month_output_dir, target_month, template)
    unresolved = [record for record in included if record.needs_confirmation or record.hours is None]
    if unresolved:
        details = "\n".join(
            f"- {record.date} {record.course_name}: {record.confirmation_note or '需要人工确认'}"
            for record in unresolved
        )
        raise RuntimeError("存在未确认课程，禁止写入 Excel：\n" + details)
    output = unique_output_path(month_output_dir / build_output_name(template.name, year, month))
    title = build_target_title(layout["title"], year, month)
    result = write_workbook(template, output, title, included, layout)
    result.update({
        "target_month": target_month,
        "template": str(template),
        "output": str(output),
        "records": [asdict(record) for record in included],
        "training_count": sum(record.category == "培训" for record in included),
        "assessment_count": sum(record.category == "考核" for record in included),
        "assessment_summary": build_assessment_summary(included),
    })
    return result
