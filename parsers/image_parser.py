from __future__ import annotations

from pathlib import Path

from .ocr_parser import run_rapidocr


def parse_image(path: Path, year: int, target_month: int, teacher_name: str, enabled: bool = False,
                debug_dir: Path | None = None):
    if not enabled:
        return [], (
            f"{path.name}: 照片课表未自动识别文字。"
            "如需写入照片中的课程，请使用手工补录或安装 OCR。"
        )
    try:
        records, payload = run_rapidocr(
            path, year, target_month, debug_dir or path.parent / "ocr_debug"
        )
    except Exception as exc:
        return [], f"{path.name}: OCR失败，可手工补录：{exc}"
    debug = payload["debug_files"]
    if not records:
        return [], (
            f"{path.name}: 图片 OCR 未能识别出有效课程，请使用手工补录。"
            f" OCR原始文本：{debug['text']}；OCR结果：{debug['json']}；最佳方向图片：{debug['image']}"
        )
    return records, (
        f"{path.name}: RapidOCR 已解析 {len(records)} 条月历课程，最佳方向 {payload['best_angle']}°。"
        f" OCR原始文本：{debug['text']}；OCR结果：{debug['json']}；最佳方向图片：{debug['image']}"
    )
