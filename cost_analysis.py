"""
人员成本分析核心逻辑
- 读取在职人员表（read_only=True）+ SHA-256 校验
- 表头探测（前 12 行）+ 列别名映射
- 按部门 / 职级汇总人数、总成本、人均成本、占比
- 全链路 Decimal
"""
import hashlib
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl

MAX_HEADER_SCAN_ROWS = 12

COST_ALIASES = {
    "emp_id":         {"工号", "员工编号", "职工编号", "员工工号"},
    "name":           {"姓名", "员工姓名"},
    "dept":           {"部门", "部门名称", "所在部门"},
    "level":          {"职级", "岗位级别", "级别", "职位级别"},
    "monthly_salary": {"月薪", "月工资", "基本工资", "月薪合计"},
}

COST_SCOPE = "月薪合计，不含奖金和社保"


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


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


def _parse_salary(raw) -> Decimal | None:
    if raw is None:
        return None
    s = str(raw).strip().replace(",", "").replace("，", "")
    for sym in ("￥", "¥", "$", "CNY"):
        s = s.replace(sym, "")
    for suffix in ("元", "万元"):
        if s.endswith(suffix):
            s = s[: -len(suffix)]
    if not s:
        return None
    try:
        return Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return None


def analyze(file_path: Path) -> dict:
    path = Path(file_path)
    sha_before = sha256_of(path)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = [[cell.value for cell in row] for row in ws.rows]
    wb.close()

    sha_after = sha256_of(path)

    alias_map = _build_alias_map(COST_ALIASES)
    header_idx, col_map = _detect_header(all_rows, alias_map)

    required = {"dept", "monthly_salary"}
    missing = required - set(col_map.keys())
    if missing:
        raise ValueError(f"MISSING_REQUIRED_COLUMN|缺少必要列（{', '.join(missing)}）")

    persons = []
    for row in all_rows[header_idx + 1:]:
        if all(c is None for c in row):
            continue
        record: dict = {}
        for field, idx in col_map.items():
            if idx < len(row):
                record[field] = row[idx]

        dept = str(record.get("dept") or "").strip()
        if not dept:
            continue
        salary = _parse_salary(record.get("monthly_salary"))
        if salary is None:
            continue

        persons.append({
            "emp_id": str(record.get("emp_id") or "").strip(),
            "name": str(record.get("name") or "").strip(),
            "dept": dept,
            "level": str(record.get("level") or "").strip() or "未标注",
            "monthly_salary": salary,
        })

    if not persons:
        raise ValueError("NO_DATA|未找到有效人员数据")

    grand_total = sum(p["monthly_salary"] for p in persons)

    dept_agg = defaultdict(lambda: {"headcount": 0, "total_cost": Decimal("0")})
    dept_level_agg = defaultdict(lambda: {"headcount": 0, "total_cost": Decimal("0")})
    levels_set = set()

    for p in persons:
        d = p["dept"]
        l = p["level"]
        levels_set.add(l)
        dept_agg[d]["headcount"] += 1
        dept_agg[d]["total_cost"] += p["monthly_salary"]
        dept_level_agg[(d, l)]["headcount"] += 1
        dept_level_agg[(d, l)]["total_cost"] += p["monthly_salary"]

    rows = []

    for dept in sorted(dept_agg.keys()):
        a = dept_agg[dept]
        avg = (a["total_cost"] / a["headcount"]).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        pct = (a["total_cost"] / grand_total * 100).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP) if grand_total else Decimal("0")
        rows.append({
            "row_type": "dept_summary",
            "dept": dept,
            "headcount": a["headcount"],
            "total_cost": str(a["total_cost"]),
            "avg_cost": str(avg),
            "pct": str(pct),
            "category": "dept_summary",
        })

    for (dept, level) in sorted(dept_level_agg.keys()):
        a = dept_level_agg[(dept, level)]
        avg = (a["total_cost"] / a["headcount"]).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        pct = (a["total_cost"] / grand_total * 100).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP) if grand_total else Decimal("0")
        rows.append({
            "row_type": "dept_level",
            "dept": dept,
            "level": level,
            "headcount": a["headcount"],
            "total_cost": str(a["total_cost"]),
            "avg_cost": str(avg),
            "pct": str(pct),
            "category": "dept_level",
        })

    for p in persons:
        rows.append({
            "row_type": "detail",
            "match_key": p["emp_id"],
            "name": p["name"],
            "dept": p["dept"],
            "level": p["level"],
            "monthly_salary": str(p["monthly_salary"]),
            "category": "detail",
        })

    summary = {
        "total_persons": len(persons),
        "total_cost": str(grand_total),
        "avg_cost": str((grand_total / len(persons)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)),
        "dept_count": len(dept_agg),
        "level_count": len(levels_set),
        "cost_scope": COST_SCOPE,
    }

    return {
        "summary": summary,
        "rows": rows,
        "files": [
            {
                "side": "staff",
                "sha256_before": sha_before,
                "sha256_after": sha_after,
                "readonly_verified": sha_before == sha_after,
            },
        ],
    }
