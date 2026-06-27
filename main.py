from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import date
from pathlib import Path

from parsers import parse_sources
from parsers.common import CourseRecord
from preview import assign_targets, write_previews
from rules import classify
from workload_writer import (
    build_output_name, build_target_title, inspect_template, locate_template, write_workbook,
)


def match_teacher(raw_teacher: str, teacher_name: str, teacher_aliases=None) -> tuple[bool, str, str]:
    raw = re.sub(r"\s+", "", raw_teacher or "")
    full = re.sub(r"\s+", "", teacher_name)
    aliases = {re.sub(r"\s+", "", alias) for alias in (teacher_aliases or [full[0]]) if alias}
    if raw == full:
        return True, "全名命中", teacher_name
    if raw in aliases:
        return True, "简称命中", teacher_name
    tokens = [token for token in re.split(r"[、/,，]+", raw) if token]
    if len(tokens) > 1 and any(token == full or token in aliases for token in tokens):
        canonical = [teacher_name if token == full or token in aliases else token for token in tokens]
        return True, "多教师包含命中", "、".join(canonical)
    if full and full in raw:
        return True, "多教师包含命中", raw_teacher
    return False, "未命中", raw_teacher


def load_config(path: Path) -> dict:
    defaults = {
        "teacher_name": "黄佳豪", "teacher_aliases": ["黄"], "target_month": "", "template_keyword": "工作量表",
        "input_dir": ".", "output_dir": "output", "enable_ocr": False,
        "require_confirm_before_write": True,
    }
    if path.exists():
        defaults.update(json.loads(path.read_text(encoding="utf-8")))
    return defaults


def parse_target_month(raw: str, template: Path) -> tuple[int, int, str]:
    value = (raw or "").strip()
    match = re.fullmatch(r"(\d{4})[-/]([01]?\d)", value)
    if not match:
        match = re.fullmatch(r"(\d{4})年\s*([01]?\d)月", value)
    inference = "命令行或配置明确给出年份"
    if match:
        year, month = int(match.group(1)), int(match.group(2))
    else:
        month_only = re.fullmatch(r"([01]?\d)月", value)
        if not month_only:
            raise RuntimeError("目标月份格式应为 2026-07、2026年7月或7月")
        month = int(month_only.group(1))
        year_match = re.search(r"(20\d{2})年", template.name)
        if year_match:
            year = int(year_match.group(1))
            inference = f"仅输入月份；年份由模板文件名推断为 {year}"
        else:
            year = date.today().year
            inference = f"仅输入月份；模板无年份，按当前年份推断为 {year}"
    if not 1 <= month <= 12:
        raise RuntimeError("目标月份必须在1至12之间")
    return year, month, inference


def classify_and_filter(records, teacher_name: str, teacher_aliases: list[str], year: int, month: int):
    target = f"{year:04d}-{month:02d}"
    for record in records:
        record.category, record.subcategory, uncertain = classify(record.project, record.course_name)
        matched, match_type, canonical_teacher = match_teacher(record.teacher, teacher_name, teacher_aliases)
        record.teacher_match_type = match_type
        if matched:
            record.teacher = canonical_teacher
        record.write_module = record.category
        if record.category == "考核" and "总复习" in record.course_name:
            record.subcategory = "理论"
            record.hours = 8
        if record.category == "考核" and record.subcategory == "实训" and record.hours is None:
            record.needs_confirmation = True
            record.confirmation_note = record.confirmation_note or "实训课时缺失，需确认是否按一天8课时或半天4课时计算。"
        record.needs_confirmation = record.needs_confirmation or uncertain
        if not record.date:
            record.status, record.exclusion_reason = "排除", "无法判断日期"
        elif not matched:
            record.status, record.exclusion_reason = "排除", f"教师不是{teacher_name}"
        elif not record.date.startswith(target):
            record.status, record.exclusion_reason = "参考", "非目标月份，仅作规则参考"
        elif uncertain:
            record.status, record.exclusion_reason = "排除", "分类或日期需要人工确认"
        else:
            record.status, record.exclusion_reason = "待写入", ""
    return sorted(records, key=lambda r: (r.date or "9999", r.start_time, r.course_name))


def print_summary(records, warnings):
    groups = {name: [r for r in records if r.status == name] for name in ("待写入", "参考", "排除")}
    for name in ("待写入", "参考", "排除"):
        print(f"\n{name}课程（{len(groups[name])}条）：")
        for r in groups[name]:
            target = r.target_cells or "-"
            reason = f"；{r.exclusion_reason}" if r.exclusion_reason else ""
            hours = "需确认" if r.hours is None else f"{r.hours}课时"
            print(f"- {r.date or '日期未知'} {r.start_time}-{r.end_time} | {r.course_name} | {r.teacher} | "
                  f"{hours} | {r.category}-{r.subcategory or '空'} | {target}{reason}")
            if r.confirmation_note:
                print(f"  确认说明：{r.confirmation_note}")
    if warnings:
        print("\n提示：")
        for warning in warnings:
            print(f"- {warning}")


def load_preview_snapshot(month_dir: Path, target_month: str, template: Path) -> list[CourseRecord]:
    preview_path = month_dir / "待写入预览表.csv"
    snapshot_path = month_dir / "解析结果.json"
    required = [preview_path, month_dir / "被排除课程.csv", month_dir / "解析日志.txt", snapshot_path]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise RuntimeError(
            "缺少对应月份的 preview 文件，请先运行 --preview：\n" + "\n".join(f"- {path}" for path in missing)
        )
    snapshot = json.loads(snapshot_path.read_text(encoding="utf-8"))
    if snapshot.get("target_month") != target_month:
        raise RuntimeError("解析结果月份与本次 --target-month 不一致，已停止")
    if Path(snapshot.get("template", "")).resolve() != template.resolve():
        raise RuntimeError("preview 使用的模板与本次 --template 不一致，已停止")
    with preview_path.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    if any(row.get("目标月份") != target_month for row in rows):
        raise RuntimeError("待写入预览表中存在其他月份记录，已停止")
    records = [CourseRecord(**item) for item in snapshot.get("records", [])]
    included = [record for record in records if record.status == "待写入"]
    if len(rows) != len(included):
        raise RuntimeError("待写入预览表与同月份解析快照数量不一致，请重新运行 --preview")
    return included


def main():
    parser = argparse.ArgumentParser(description="按任意目标月份自动解析并预览/填入工作量表")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--preview", action="store_true", help="只生成预览，不修改Excel")
    mode.add_argument("--write", action="store_true", help="确认后生成目标月份Excel")
    parser.add_argument("--target-month", help="目标月份：2026-07、2026年7月或7月")
    parser.add_argument("--template", type=Path, help="工作量表模板路径")
    parser.add_argument("--input-dir", type=Path, help="课程依据文件夹")
    parser.add_argument("--output-dir", type=Path, help="预览及最终文件输出目录")
    parser.add_argument("--config", type=Path, default=Path("config.json"), help="配置文件")
    args = parser.parse_args()

    config = load_config(args.config.resolve())
    base = Path.cwd()
    input_dir = (args.input_dir or Path(config["input_dir"]))
    input_dir = (base / input_dir).resolve() if not input_dir.is_absolute() else input_dir.resolve()
    output_dir = (args.output_dir or Path(config["output_dir"]))
    output_dir = (base / output_dir).resolve() if not output_dir.is_absolute() else output_dir.resolve()
    template = locate_template(input_dir, args.template, config["template_keyword"], config["teacher_name"])
    raw_month = args.target_month or config.get("target_month", "")
    if not raw_month:
        raise SystemExit("未提供目标月份。请使用 --target-month，或在 config.json 中设置 target_month。")
    year, month, inference = parse_target_month(raw_month, template)
    target_month = f"{year:04d}-{month:02d}"
    month_output_dir = output_dir / target_month
    layout = inspect_template(template)
    if args.write:
        try:
            included = load_preview_snapshot(month_output_dir, target_month, template)
        except RuntimeError as exc:
            raise SystemExit(str(exc)) from exc
        capacity = layout["data_end_row"] - layout["data_start_row"] + 1
        from rules import assessment_group_key
        training_count = sum(r.category == "培训" for r in included)
        assessment_count = len({assessment_group_key(r.project, r.subcategory) for r in included if r.category == "考核"})
        assessment_capacity = len(layout["assessment_slots"])
        if training_count > capacity or assessment_count > assessment_capacity:
            raise SystemExit(
                f"教学模块空白槽位不足，当前需要培训 {training_count} 条、考核 {assessment_count} 组；"
                f"模板仅有培训 {capacity} 行、考核 {assessment_capacity} 个槽位，是否允许复制格式扩展行？"
            )
        unresolved = [r for r in included if r.needs_confirmation or r.hours is None]
        if unresolved:
            details = "\n".join(
                f"- {r.date} {r.course_name}: {r.confirmation_note or '需要人工确认'}" for r in unresolved
            )
            raise SystemExit("存在未确认课程，禁止写入 Excel。请先更新 preview 对应解析记录：\n" + details)
        output = month_output_dir / build_output_name(template.name, year, month)
        title = build_target_title(layout["title"], year, month)
        print("\n最终写入前确认")
        print(f"目标月份：{target_month}")
        print(f"模板文件：{template}")
        print(f"输出文件：{output}")
        print(f"即将写入课程数量：{len(included)}")
        print("即将写入课程清单：")
        for index, record in enumerate(included, 1):
            print(f"  {index}. {record.date} {record.start_time}-{record.end_time} | "
                  f"{record.course_name} | {record.hours}课时 | {record.category}-{record.subcategory or '空'}")
        print(f"是否修改表头月份：是（{layout['title_cell']} -> {year}年{month}月）")
        print("是否只写教学模块：是")
        if input("确认继续？请输入 y：").strip().lower() != "y":
            print("已取消，未生成 Excel。")
            return
        result = write_workbook(template, output, title, included, layout)
        print(f"已生成：{result['output']}\n改动：{', '.join(result['changed'])}\n未改区域：{result['untouched']}")
        return

    records, parser_warnings, scanned = parse_sources(
        input_dir, year, template, output_dir, bool(config.get("enable_ocr", False)), config["teacher_name"]
    )
    records = classify_and_filter(records, config["teacher_name"], config.get("teacher_aliases", []), year, month)
    capacity_warnings = assign_targets(records, layout)
    warnings = parser_warnings + capacity_warnings
    log_lines = [
        f"目标月份: {target_month}", f"月份推断: {inference}", f"教师: {config['teacher_name']}",
        f"模板: {template}", f"工作表: {layout['sheet']}", f"月份单元格: {layout['title_cell']}",
        f"教学数据行: {layout['data_start_row']}:{layout['data_end_row']}",
        f"输入目录: {input_dir}", f"扫描依据文件数: {len(scanned)}",
        *[f"扫描: {p.name}" for p in scanned],
        f"待写入: {sum(r.status == '待写入' for r in records)}",
        f"参考: {sum(r.status == '参考' for r in records)}",
        f"排除: {sum(r.status == '排除' for r in records)}",
        *[f"警告: {warning}" for warning in warnings],
    ]
    paths = write_previews(month_output_dir, target_month, records, log_lines, template)
    print(f"目标月份：{target_month}（{inference}）\n模板：{template.name}\n教学数据行："
          f"{layout['data_start_row']}:{layout['data_end_row']}")
    print_summary(records, warnings)
    print(f"\n预览：{paths[0]}\n排除：{paths[1]}\n日志：{paths[2]}")
    return


if __name__ == "__main__":
    main()
