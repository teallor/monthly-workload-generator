from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

from .common import CourseRecord


OCR_KEYWORDS = (
    "期", "电动港机", "装卸机械司机", "中级工", "培训课程表",
    "实训", "总复习", "黄", "王", "月份", "星期",
)


def ocr_self_check() -> dict:
    checks = {}
    errors = []
    try:
        import rapidocr_onnxruntime
        from rapidocr_onnxruntime import RapidOCR
        checks["RapidOCR"] = f"可用（{getattr(rapidocr_onnxruntime, '__version__', '版本未知')}）"
        package_dir = Path(rapidocr_onnxruntime.__file__).resolve().parent
        model_files = sorted(package_dir.rglob("*.onnx"))
        checks["OCR模型"] = f"可访问（{len(model_files)}个ONNX模型）" if model_files else "未找到ONNX模型"
        if not model_files:
            errors.append("未找到 RapidOCR ONNX 模型文件")
    except Exception as exc:
        RapidOCR = None
        checks["RapidOCR"] = f"不可用：{exc}"
        checks["OCR模型"] = "无法检查"
        errors.append(f"RapidOCR 导入失败：{exc}")
    try:
        import onnxruntime
        providers = ", ".join(onnxruntime.get_available_providers())
        checks["ONNX Runtime"] = f"可用（{onnxruntime.__version__}；{providers}）"
    except Exception as exc:
        checks["ONNX Runtime"] = f"不可用：{exc}"
        errors.append(f"ONNX Runtime 导入失败：{exc}")
    try:
        import cv2
        checks["OpenCV"] = f"可用（{cv2.__version__}）"
    except Exception as exc:
        checks["OpenCV"] = f"不可用：{exc}"
        errors.append(f"OpenCV 导入失败：{exc}")
    try:
        import numpy as np
        checks["NumPy"] = f"可用（{np.__version__}）"
    except Exception as exc:
        np = None
        checks["NumPy"] = f"不可用：{exc}"
        errors.append(f"NumPy 导入失败：{exc}")
    try:
        import PIL
        checks["Pillow"] = f"可用（{PIL.__version__}）"
    except Exception as exc:
        checks["Pillow"] = f"不可用：{exc}"
        errors.append(f"Pillow 导入失败：{exc}")
    if RapidOCR is not None and np is not None and not errors:
        try:
            engine = RapidOCR()
            blank = np.full((80, 260, 3), 255, dtype=np.uint8)
            engine(blank)
            checks["运行测试"] = "通过（模型会话可创建并执行）"
        except Exception as exc:
            checks["运行测试"] = f"失败：{exc}"
            errors.append(f"OCR 模型运行失败：{exc}")
    else:
        checks["运行测试"] = "未执行"
    return {"available": not errors, "checks": checks, "errors": errors}


@dataclass
class OcrItem:
    box: list[list[float]]
    text: str
    score: float

    @property
    def cx(self) -> float:
        return sum(point[0] for point in self.box) / len(self.box)

    @property
    def cy(self) -> float:
        return sum(point[1] for point in self.box) / len(self.box)

    def to_dict(self) -> dict:
        return {"box": self.box, "text": self.text, "score": self.score, "center": [self.cx, self.cy]}


def _read_image(path: Path):
    import cv2
    import numpy as np

    data = np.fromfile(str(path), dtype=np.uint8)
    image = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if image is None:
        raise RuntimeError("无法读取图片像素")
    return image


def _write_png(path: Path, image) -> None:
    import cv2

    path.parent.mkdir(parents=True, exist_ok=True)
    success, encoded = cv2.imencode(".png", image)
    if not success:
        raise RuntimeError(f"无法保存 OCR 调试图片：{path}")
    encoded.tofile(str(path))


def _rotate(image, angle: int):
    import cv2

    rotations = {
        0: image,
        90: cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE),
        180: cv2.rotate(image, cv2.ROTATE_180),
        270: cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE),
    }
    return rotations[angle]


def _preprocess(image):
    import cv2

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    mask = gray < 248
    points = cv2.findNonZero(mask.astype("uint8"))
    if points is not None:
        x, y, width, height = cv2.boundingRect(points)
        padding = max(12, int(min(gray.shape) * 0.01))
        x0, y0 = max(0, x - padding), max(0, y - padding)
        x1, y1 = min(gray.shape[1], x + width + padding), min(gray.shape[0], y + height + padding)
        gray = gray[y0:y1, x0:x1]
    max_dimension = max(gray.shape)
    if max_dimension < 2600:
        scale = min(2.0, 2600 / max_dimension)
        gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(gray)
    gray = cv2.fastNlMeansDenoising(gray, None, 7, 7, 21)
    return cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 11
    )


def _normalize_result(result) -> list[OcrItem]:
    items: list[OcrItem] = []
    for raw in result or []:
        if not isinstance(raw, (list, tuple)) or len(raw) < 3:
            continue
        box, text, score = raw[0], str(raw[1]).strip(), float(raw[2])
        if text and box:
            items.append(OcrItem(
                box=[[float(point[0]), float(point[1])] for point in box], text=text, score=score,
            ))
    return items


def _ocr(engine, image) -> list[OcrItem]:
    result, _ = engine(image)
    return _normalize_result(result)


def _is_schedule_title(text: str) -> bool:
    return bool(re.search(r"第?\s*\d{3,4}\s*期", text) or "培训课程表" in text)


def _orientation_score(items: list[OcrItem], width: int, height: int, target_month: int) -> float:
    text = "\n".join(item.text for item in items)
    keyword_score = sum(3 if keyword in {"实训", "总复习"} else 1
                        for keyword in OCR_KEYWORDS if keyword in text)
    score = float(keyword_score)
    titles = [item for item in items if _is_schedule_title(item.text)]
    score += len(titles) * 8
    score += sum(18 for item in titles if item.cy < height * 0.32)
    month_pattern = re.compile(rf"(?<!\d){target_month}\s*月份?")
    score += sum(6 for item in items if month_pattern.search(item.text))
    score += sum(8 for item in items if "上课注意事项" in item.text and item.cy > height * 0.70)
    score += min(len(items), 150) / 30
    return score


def _best_orientation(path: Path, target_month: int):
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError as exc:
        raise RuntimeError(
            "未安装 RapidOCR。请运行：pip install rapidocr-onnxruntime opencv-python pillow numpy"
        ) from exc

    source = _read_image(path)
    engine = RapidOCR()
    candidates = []
    for angle in (0, 90, 180, 270):
        oriented = _rotate(source, angle)
        items = _ocr(engine, oriented)
        height, width = oriented.shape[:2]
        candidates.append({
            "angle": angle, "image": oriented, "items": items,
            "score": _orientation_score(items, width, height, target_month), "variant": "原图增强前",
        })
    best = max(candidates, key=lambda candidate: candidate["score"])

    processed = _preprocess(best["image"])
    processed_items = _ocr(engine, processed)
    processed_height, processed_width = processed.shape[:2]
    processed_score = _orientation_score(
        processed_items, processed_width, processed_height, target_month
    )
    if processed_score > best["score"] + 2:
        best = {
            **best, "items": processed_items, "score": processed_score,
            "ocr_image": processed, "variant": "灰度放大+对比度增强+去噪+自适应二值化",
        }
    else:
        best["ocr_image"] = best["image"]
    best["orientation_candidates"] = [
        {"angle": candidate["angle"], "score": candidate["score"], "items": len(candidate["items"])}
        for candidate in candidates
    ]
    return best


def _clean(text: str) -> str:
    return re.sub(r"\s+", "", text).replace("（", "(").replace("）", ")")


def _month_number(text: str) -> int | None:
    match = re.search(r"([1-9]|1[0-2])\s*月份", text)
    return int(match.group(1)) if match else None


def _course_candidate(text: str) -> tuple[int, str] | None:
    cleaned = _clean(text)
    date_match = re.match(r"(\d{1,2})", cleaned)
    if not date_match:
        return None
    day = int(date_match.group(1))
    if not 1 <= day <= 31:
        return None
    remainder = cleaned[date_match.end():]
    if re.search(r"总.{0,2}[复发].{0,2}习|总复|总发习", remainder):
        return day, "总复习4"
    if "训" in remainder and len(remainder) <= 8 and "总" not in remainder:
        return day, "实训"
    if any(token in remainder for token in ("实川", "买训", "实洲", "次训", "实性")):
        return day, "实训"
    return None


def _teacher_near(course: OcrItem, items: list[OcrItem], width: int, height: int) -> tuple[str, float, str] | None:
    nearby = []
    for item in items:
        cleaned = _clean(item.text)
        if "黄" not in cleaned:
            continue
        dx, dy = abs(item.cx - course.cx), abs(item.cy - course.cy)
        if dx <= width * 0.065 and dy <= height * 0.045:
            nearby.append((dx + dy * 1.3, item))
    if not nearby:
        return None
    teacher_item = min(nearby, key=lambda pair: pair[0])[1]
    teacher_text = _clean(teacher_item.text)
    teacher = "黄、王" if "王" in teacher_text else "黄"
    return teacher, teacher_item.score, teacher_item.text


def parse_calendar_records(path: Path, year: int, items: list[OcrItem], width: int, height: int,
                           angle: int) -> list[CourseRecord]:
    titles = [item for item in items if _is_schedule_title(item.text)]
    titles.sort(key=lambda item: item.cx)
    if not titles:
        return []
    boundaries = [0.0]
    boundaries.extend((left.cx + right.cx) / 2 for left, right in zip(titles, titles[1:]))
    boundaries.append(float(width))
    records: list[CourseRecord] = []
    seen: set[tuple] = set()
    for title_index, title in enumerate(titles):
        x0, x1 = boundaries[title_index], boundaries[title_index + 1]
        region = [item for item in items if x0 <= item.cx < x1]
        months = [(number, item) for item in region if (number := _month_number(item.text))]
        months.sort(key=lambda pair: pair[1].cy)
        if not months:
            continue
        remarks = [item.cy for item in region if "备注" in item.text and item.cy > title.cy]
        region_bottom = min(remarks) - height * 0.008 if remarks else height * 0.94
        for month_index, (month, month_item) in enumerate(months):
            if month_index:
                start_y = (months[month_index - 1][1].cy + month_item.cy) / 2
            else:
                start_y = max(title.cy + height * 0.04, month_item.cy - height * 0.09)
            if month_index + 1 < len(months):
                end_y = (month_item.cy + months[month_index + 1][1].cy) / 2
            else:
                end_y = region_bottom
            block = [item for item in region if start_y <= item.cy < end_y]
            for item in block:
                candidate = _course_candidate(item.text)
                if not candidate:
                    continue
                day, course_type = candidate
                teacher_match = _teacher_near(item, block, x1 - x0, height)
                if not teacher_match:
                    continue
                teacher, teacher_score, teacher_raw = teacher_match
                course_name = "实训（黄、王）" if course_type == "实训" and teacher == "黄、王" else (
                    "实训（黄）" if course_type == "实训" else "总复习4"
                )
                key = (title.text, month, day, course_type, teacher)
                if key in seen:
                    continue
                seen.add(key)
                confidence = min(title.score, item.score, teacher_score)
                records.append(CourseRecord(
                    source_file=path.name,
                    date=f"{year:04d}-{month:02d}-{day:02d}", start_time="", end_time="",
                    course_name=course_name, teacher=teacher, hours=8,
                    project=title.text, context=(
                        f"RapidOCR方向={angle}° | OCR课程={item.text} | OCR教师={teacher_raw} | "
                        f"月份={month}月"
                    ), confidence=confidence,
                    needs_confirmation=confidence < 0.65,
                    confirmation_note=("OCR置信度较低，请核对照片。" if confidence < 0.65 else ""),
                ))
    return sorted(records, key=lambda record: (record.date, record.project, record.course_name))


def run_rapidocr(path: Path, year: int, target_month: int,
                 debug_dir: Path) -> tuple[list[CourseRecord], dict]:
    best = _best_orientation(path, target_month)
    items = best["items"]
    ocr_image = best["ocr_image"]
    height, width = ocr_image.shape[:2]
    records = parse_calendar_records(path, year, items, width, height, best["angle"])

    debug_dir.mkdir(parents=True, exist_ok=True)
    image_path = debug_dir / f"{path.stem}_最佳方向.png"
    text_path = debug_dir / f"{path.stem}_ocr文本.txt"
    json_path = debug_dir / f"{path.stem}_ocr结果.json"
    _write_png(image_path, best["image"])
    text_path.write_text("\n".join(item.text for item in items) + "\n", encoding="utf-8")
    payload = {
        "source_file": path.name,
        "engine": "rapidocr-onnxruntime",
        "target_month": target_month,
        "best_angle": best["angle"],
        "best_variant": best["variant"],
        "orientation_score": best["score"],
        "orientation_candidates": best["orientation_candidates"],
        "ocr_items": [item.to_dict() for item in items],
        "parsed_records": [record.to_dict() for record in records],
        "debug_files": {"image": str(image_path), "text": str(text_path), "json": str(json_path)},
    }
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return records, payload
