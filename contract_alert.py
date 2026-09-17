"""
合同到期预警核心逻辑
- 读取员工合同台账 Excel（read_only=True）+ SHA-256 校验
- 表头行探测（扫描前 12 行）
- 列别名映射 → 标准字段名
- 日期解析 → 计算剩余天数 → 分级预警
"""
import hashlib
from datetime import date, datetime
from pathlib import Path

import openpyxl

MAX_HEADER_SCAN_ROWS = 12

CONTRACT_ALIASES = {
    "emp_id":        {"工号", "员工编号", "职工编号", "员工工号"},
    "name":          {"姓名", "员工姓名"},
    "dept":          {"部门", "部门名称", "所在部门"},
    "contract_type": {"合同类型", "合同性质", "用工形式"},
    "start_date":    {"合同开始日期", "合同起始日", "合同起始日期", "入职日期", "起始日期"},
    "end_date":      {"合同结束日期", "合同到期日", "合同终止日期", "合同到期日期", "到期日期", "终止日期"},
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
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"DATE_PARSE_ERROR|无法解析日期：{raw!r}")


def _classify(days_remaining: int | None) -> str:
    if days_remaining is None:
        return "normal"
    if days_remaining < 0:
        return "expired"
    if days_remaining <= 30:
        return "within_30"
    if days_remaining <= 60:
        return "within_60"
    if days_remaining <= 90:
        return "within_90"
    return "normal"


def analyze(file_path: Path, ref_date: date | None = None) -> dict:
    """
    分析合同台账，返回：
    {
      summary: {expired, within_30, within_60, within_90, normal},
      rows: [{match_key, name, dept, contract_type, start_date, end_date, days_remaining, category}],
      files: [{side, sha256_before, sha256_after, readonly_verified}],
    }
    """
    today = ref_date or date.today()
    path = Path(file_path)

    sha_before = sha256_of(path)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = [[cell.value for cell in row] for row in ws.rows]
    wb.close()

    sha_after = sha256_of(path)

    alias_map = _build_alias_map(CONTRACT_ALIASES)
    header_idx, col_map = _detect_header(all_rows, alias_map)

    required = {"emp_id", "end_date"}
    missing = required - set(col_map.keys())
    if missing:
        all_headers = [str(v).strip() for v in all_rows[header_idx] if v is not None]
        known = {n for names in CONTRACT_ALIASES.values() for n in names}
        raise ValueError(
            f"MISSING_REQUIRED_COLUMN|缺少必要列（{', '.join(missing)}）"
            f"|found={all_headers}|known={sorted(known)}"
        )

    results = []
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
        contract_type = str(record.get("contract_type") or "").strip() or None

        start_date = _parse_date(record.get("start_date"))
        end_date = _parse_date(record.get("end_date"))

        is_open_ended = contract_type and "无固定" in contract_type

        if is_open_ended or end_date is None:
            days_remaining = None
            category = "normal"
        else:
            days_remaining = (end_date - today).days
            category = _classify(days_remaining)

        results.append({
            "match_key": emp_id,
            "name": name,
            "dept": dept,
            "contract_type": contract_type,
            "start_date": start_date.isoformat() if start_date else None,
            "end_date": end_date.isoformat() if end_date else None,
            "days_remaining": days_remaining,
            "category": category,
        })

    results.sort(key=lambda r: (
        r["days_remaining"] if r["days_remaining"] is not None else 99999,
    ))

    summary = {
        "expired": sum(1 for r in results if r["category"] == "expired"),
        "within_30": sum(1 for r in results if r["category"] == "within_30"),
        "within_60": sum(1 for r in results if r["category"] == "within_60"),
        "within_90": sum(1 for r in results if r["category"] == "within_90"),
        "normal": sum(1 for r in results if r["category"] == "normal"),
    }

    return {
        "summary": summary,
        "rows": results,
        "files": [
            {
                "side": "contract",
                "sha256_before": sha_before,
                "sha256_after": sha_after,
                "readonly_verified": sha_before == sha_after,
            },
        ],
    }
