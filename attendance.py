"""
考勤统计核心逻辑
- 读取打卡明细 Excel（read_only=True）+ SHA-256 校验
- 表头探测（扫描前 12 行）+ 列别名映射
- 按可配置班次判断迟到、早退、缺卡、加班
- 输出逐日明细（detail）+ 按人汇总（person_summary）
"""
import hashlib
from collections import defaultdict
from datetime import date, datetime, time, timedelta
from pathlib import Path

import openpyxl

MAX_HEADER_SCAN_ROWS = 12

DEFAULT_SHIFT = {
    "start": "08:30",
    "end": "17:30",
    "ot_threshold_min": 30,
}

ATTENDANCE_ALIASES = {
    "att_date":   {"日期", "打卡日期", "考勤日期", "出勤日期"},
    "emp_id":     {"工号", "员工编号", "职工编号", "员工工号"},
    "name":       {"姓名", "员工姓名"},
    "dept":       {"部门", "部门名称", "所在部门"},
    "clock_in":   {"上班打卡", "上班时间", "签到时间", "上班签到"},
    "clock_out":  {"下班打卡", "下班时间", "签退时间", "下班签退"},
}


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def normalize_id(raw) -> str | None:
    if raw is None:
        return None
    s = str(raw).strip()
    s = "".join(
        chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else c
        for c in s
    )
    s = s.upper()
    return s if s else None


def _build_alias_map(aliases: dict[str, set]) -> dict[str, str]:
    m: dict[str, str] = {}
    for field, names in aliases.items():
        for n in names:
            m[n.lower()] = field
    return m


def _detect_header(all_rows: list[list], alias_map: dict[str, str]) -> tuple[int, dict[str, int]]:
    best_idx, best_score, best_col_map = 0, 0, {}
    for i, row in enumerate(all_rows[:MAX_HEADER_SCAN_ROWS]):
        col_map: dict[str, int] = {}
        for j, cell in enumerate(row):
            if cell is None:
                continue
            key = str(cell).strip().lower()
            if key in alias_map and alias_map[key] not in col_map:
                col_map[alias_map[key]] = j
        if len(col_map) > best_score:
            best_score, best_idx, best_col_map = len(col_map), i, col_map
    if best_score == 0:
        raise ValueError("HEADER_NOT_FOUND|未找到有效表头行")
    return best_idx, best_col_map


def _parse_time(raw) -> time | None:
    if raw is None:
        return None
    if isinstance(raw, time):
        return raw
    if isinstance(raw, datetime):
        return raw.time()
    s = str(raw).strip()
    if not s:
        return None
    for fmt in ("%H:%M:%S", "%H:%M", "%H:%M:%S.%f"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            continue
    return None


def _parse_date(raw) -> date | None:
    if raw is None:
        return None
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    s = str(raw).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _time_to_minutes(t: time) -> int:
    return t.hour * 60 + t.minute


def _parse_shift_time(s: str) -> time:
    parts = s.strip().split(":")
    return time(int(parts[0]), int(parts[1]))


def analyze(file_path: Path, shift_config: dict | None = None) -> dict:
    path = Path(file_path)
    sha_before = sha256_of(path)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = [[cell.value for cell in row] for row in ws.rows]
    wb.close()

    sha_after = sha256_of(path)

    cfg = {**DEFAULT_SHIFT, **(shift_config or {})}
    shift_start = _parse_shift_time(cfg["start"])
    shift_end = _parse_shift_time(cfg["end"])
    ot_threshold = int(cfg["ot_threshold_min"])

    shift_start_min = _time_to_minutes(shift_start)
    shift_end_min = _time_to_minutes(shift_end)

    alias_map = _build_alias_map(ATTENDANCE_ALIASES)
    header_idx, col_map = _detect_header(all_rows, alias_map)

    required = {"att_date", "emp_id"}
    missing = required - set(col_map.keys())
    if missing:
        raise ValueError(
            f"MISSING_REQUIRED_COLUMN|缺少必要列（{', '.join(missing)}）"
        )

    # 解析打卡记录
    detail_rows = []
    for row in all_rows[header_idx + 1:]:
        if all(c is None for c in row):
            continue
        record: dict = {}
        for field, idx in col_map.items():
            if idx < len(row):
                record[field] = row[idx]

        emp_id = normalize_id(record.get("emp_id"))
        name = str(record.get("name") or "").strip() or None
        dept = str(record.get("dept") or "").strip() or None
        att_date = _parse_date(record.get("att_date"))
        clock_in = _parse_time(record.get("clock_in"))
        clock_out = _parse_time(record.get("clock_out"))

        if not emp_id or not att_date:
            continue

        cin_min = _time_to_minutes(clock_in) if clock_in else None
        cout_min = _time_to_minutes(clock_out) if clock_out else None

        is_missing_in = clock_in is None
        is_missing_out = clock_out is None
        is_late = False
        late_min = 0
        is_early = False
        early_min = 0
        is_overtime = False
        ot_min = 0

        if not is_missing_in and cin_min > shift_start_min:
            is_late = True
            late_min = cin_min - shift_start_min

        if not is_missing_out and cout_min < shift_end_min:
            is_early = True
            early_min = shift_end_min - cout_min

        if not is_missing_out and cout_min > shift_end_min + ot_threshold:
            is_overtime = True
            ot_min = cout_min - shift_end_min - ot_threshold

        # category 取最严重
        if is_missing_in or is_missing_out:
            category = "missing_punch"
        elif is_late:
            category = "late"
        elif is_early:
            category = "early"
        elif is_overtime:
            category = "overtime"
        else:
            category = "normal"

        detail_rows.append({
            "row_type": "detail",
            "match_key": emp_id,
            "name": name,
            "dept": dept,
            "att_date": att_date.isoformat() if att_date else None,
            "clock_in": clock_in.strftime("%H:%M") if clock_in else None,
            "clock_out": clock_out.strftime("%H:%M") if clock_out else None,
            "is_late": is_late,
            "late_min": late_min,
            "is_early": is_early,
            "early_min": early_min,
            "is_missing_in": is_missing_in,
            "is_missing_out": is_missing_out,
            "is_overtime": is_overtime,
            "ot_min": ot_min,
            "category": category,
        })

    # 按人汇总
    person_data = defaultdict(lambda: {
        "name": None, "dept": None,
        "days": 0, "late_count": 0, "early_count": 0,
        "missing_count": 0, "ot_total_min": 0,
    })
    for r in detail_rows:
        key = r["match_key"]
        p = person_data[key]
        if not p["name"]:
            p["name"] = r["name"]
        if not p["dept"]:
            p["dept"] = r["dept"]
        p["days"] += 1
        if r["is_late"]:
            p["late_count"] += 1
        if r["is_early"]:
            p["early_count"] += 1
        if r["is_missing_in"] or r["is_missing_out"]:
            p["missing_count"] += 1
        if r["is_overtime"]:
            p["ot_total_min"] += r["ot_min"]

    summary_rows = []
    for emp_id in sorted(person_data.keys()):
        p = person_data[emp_id]
        has_issue = p["late_count"] + p["early_count"] + p["missing_count"] > 0
        cat = "has_issue" if has_issue else "normal"
        ot_hours = round(p["ot_total_min"] / 60, 1)
        summary_rows.append({
            "row_type": "person_summary",
            "match_key": emp_id,
            "name": p["name"],
            "dept": p["dept"],
            "days": p["days"],
            "late_count": p["late_count"],
            "early_count": p["early_count"],
            "missing_count": p["missing_count"],
            "ot_total_min": p["ot_total_min"],
            "ot_hours": ot_hours,
            "category": cat,
        })

    all_rows_out = summary_rows + detail_rows

    summary = {
        "total_records": len(detail_rows),
        "total_persons": len(person_data),
        "late": sum(1 for r in detail_rows if r["is_late"]),
        "early": sum(1 for r in detail_rows if r["is_early"]),
        "missing_punch": sum(1 for r in detail_rows if r["is_missing_in"] or r["is_missing_out"]),
        "overtime": sum(1 for r in detail_rows if r["is_overtime"]),
        "normal": sum(1 for r in detail_rows if r["category"] == "normal"),
        "shift_start": cfg["start"],
        "shift_end": cfg["end"],
        "ot_threshold_min": ot_threshold,
    }

    return {
        "summary": summary,
        "rows": all_rows_out,
        "files": [
            {
                "side": "attendance",
                "sha256_before": sha_before,
                "sha256_after": sha_after,
                "readonly_verified": sha_before == sha_after,
            },
        ],
    }
