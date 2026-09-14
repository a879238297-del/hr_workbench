"""
薪资对账核心逻辑
- 读取 Excel（read_only=True）+ SHA-256 校验
- 表头行探测（扫描前 MAX_HEADER_SCAN_ROWS 行）
- 列别名映射 → 标准字段名
- 工号归一（去空格、全角转半角、大写）
- 金额 Decimal（不用 float）
- 双表匹配，输出 4 类结果
"""
import hashlib
import unicodedata
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path

import openpyxl

MAX_HEADER_SCAN_ROWS = 12

# 标准字段 → 允许的列名别名（均已小写）
PAYROLL_ALIASES = {
    "emp_id":  {"工号", "员工编号", "职工编号", "员工工号"},
    "name":    {"姓名", "员工姓名"},
    "dept":    {"部门", "部门名称", "所在部门"},
    # 注意：不包含"应发合计"（税前总额），只取税后实发
    "amount":  {"实发合计", "实发工资", "实发金额", "实发"},
}

BANK_ALIASES = {
    "emp_id":  {"工号", "员工编号", "职工编号"},
    "name":    {"姓名", "员工姓名"},
    "amount":  {"实发金额", "代发金额", "实发合计", "实发工资", "实发"},
}


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def normalize_id(raw) -> str | None:
    """去空格、全角转半角、大写；None/空字符串返回 None"""
    if raw is None:
        return None
    s = str(raw).strip()
    # 全角转半角
    s = "".join(
        chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else c
        for c in s
    )
    s = s.upper()
    return s if s else None


def normalize_amount(raw, scale: int = 2) -> Decimal:
    """金额归一：去千位符、货币符号、单位后缀，转 Decimal"""
    if raw is None:
        raise ValueError(f"金额为空")
    s = str(raw).strip().replace(",", "").replace("，", "")
    for sym in ("￥", "¥", "$", "CNY"):
        s = s.replace(sym, "")
    for suffix in ("元", "万元"):
        if s.endswith(suffix):
            s = s[: -len(suffix)]
    try:
        d = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"无法解析金额：{raw!r}")
    return d.quantize(Decimal(10) ** -scale, rounding=ROUND_HALF_UP)


def _build_alias_map(aliases: dict[str, set]) -> dict[str, str]:
    """别名 → 标准字段名（key 均转小写）"""
    m: dict[str, str] = {}
    for field, names in aliases.items():
        for n in names:
            m[n.lower()] = field
    return m


def read_excel(path: Path, aliases: dict[str, set]) -> tuple[list[dict], str, str]:
    """
    读取 Excel，返回 (rows, sha256_before, sha256_after)
    rows: [{emp_id, name, dept?, amount_raw, amount}]
    两次 SHA-256 相同证明文件未被写回
    """
    sha_before = sha256_of(path)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = [[cell.value for cell in row] for row in ws.rows]
    wb.close()

    sha_after = sha256_of(path)

    alias_map = _build_alias_map(aliases)
    header_idx, col_map = _detect_header(all_rows, alias_map)

    required = {"emp_id", "amount"}
    missing = required - set(col_map.keys())
    if missing:
        all_headers = [str(v).strip() for v in all_rows[header_idx] if v is not None]
        known = {n for names in aliases.values() for n in names}
        raise ValueError(
            f"MISSING_REQUIRED_COLUMN|缺少必要列（{', '.join(missing)}）"
            f"|found={all_headers}|known={sorted(known)}"
        )

    rows = []
    for row in all_rows[header_idx + 1:]:
        if all(c is None for c in row):
            continue
        record: dict = {}
        for field, idx in col_map.items():
            if idx < len(row):
                record[field] = row[idx]
        record["emp_id"] = normalize_id(record.get("emp_id"))
        try:
            record["amount"] = normalize_amount(record.get("amount"))
        except ValueError as e:
            record["amount"] = None
            record["_amount_error"] = str(e)
        rows.append(record)

    return rows, sha_before, sha_after


def _detect_header(all_rows: list[list], alias_map: dict[str, str]) -> tuple[int, dict[str, int]]:
    """找最佳表头行，返回 (row_idx, {field: col_idx})"""
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


def reconcile(
    payroll_path: Path,
    bank_path: Path,
    scale: int = 2,
    tolerance: str = "0",
) -> dict:
    """
    执行薪资对账，返回：
    {
      summary: {match, amount_diff, payroll_only, bank_only},
      rows: [{match_key, name, dept, category, payroll_amount, bank_amount, delta}],
      files: [{side, sha256_before, sha256_after, readonly_verified}],
    }
    """
    tol = Decimal(tolerance)

    p_rows, p_sha_before, p_sha_after = read_excel(payroll_path, PAYROLL_ALIASES)
    b_rows, b_sha_before, b_sha_after = read_excel(bankpath := bank_path, BANK_ALIASES)

    # 按工号分组（空工号的行单独处理）
    def group_by_key(rows):
        keyed: dict[str, list] = {}
        empty: list = []
        for r in rows:
            k = r.get("emp_id")
            if k:
                keyed.setdefault(k, []).append(r)
            else:
                empty.append(r)
        return keyed, empty

    p_keyed, p_empty = group_by_key(p_rows)
    b_keyed, b_empty = group_by_key(b_rows)

    # 重复工号检测
    for side, keyed in [("payroll", p_keyed), ("bank", b_keyed)]:
        dups = {k for k, v in keyed.items() if len(v) > 1}
        if dups:
            raise ValueError(f"DUPLICATE_KEY|{side} 侧存在重复工号：{sorted(dups)}")

    results = []
    all_keys = sorted(set(p_keyed) | set(b_keyed))

    for key in all_keys:
        p = p_keyed.get(key, [None])[0] if p_keyed.get(key) else None
        b = b_keyed.get(key, [None])[0] if b_keyed.get(key) else None

        if p and b:
            pa = p.get("amount")
            ba = b.get("amount")
            if pa is not None and ba is not None and abs(ba - pa) <= tol:
                cat = "match"
                delta = Decimal("0")
            else:
                cat = "amount_diff"
                delta = (ba - pa) if (pa is not None and ba is not None) else None
        elif p:
            cat = "payroll_only"
            delta = None
        else:
            cat = "bank_only"
            delta = None

        pa_str = str(p["amount"]) if p and p.get("amount") is not None else None
        ba_str = str(b["amount"]) if b and b.get("amount") is not None else None

        if delta is not None and cat == "match":
            delta_str = None
        elif delta is not None:
            delta_str = f"{delta:+.{scale}f}"
        else:
            delta_str = None

        results.append({
            "match_key": key,
            "name": (p or b).get("name"),
            "dept": (p or {}).get("dept"),
            "category": cat,
            "payroll_amount": pa_str,
            "bank_amount": ba_str,
            "delta": delta_str,
        })

    # 空工号行归入单边缺失
    for r in p_empty:
        results.append({
            "match_key": None,
            "name": r.get("name"),
            "dept": r.get("dept"),
            "category": "payroll_only",
            "payroll_amount": str(r["amount"]) if r.get("amount") is not None else None,
            "bank_amount": None,
            "delta": None,
        })
    for r in b_empty:
        results.append({
            "match_key": None,
            "name": r.get("name"),
            "dept": r.get("dept"),
            "category": "bank_only",
            "payroll_amount": None,
            "bank_amount": str(r["amount"]) if r.get("amount") is not None else None,
            "delta": None,
        })

    summary = {
        "match": sum(1 for r in results if r["category"] == "match"),
        "amount_diff": sum(1 for r in results if r["category"] == "amount_diff"),
        "payroll_only": sum(1 for r in results if r["category"] == "payroll_only"),
        "bank_only": sum(1 for r in results if r["category"] == "bank_only"),
    }

    return {
        "summary": summary,
        "rows": results,
        "files": [
            {
                "side": "payroll",
                "sha256_before": p_sha_before,
                "sha256_after": p_sha_after,
                "readonly_verified": p_sha_before == p_sha_after,
            },
            {
                "side": "bank",
                "sha256_before": b_sha_before,
                "sha256_after": b_sha_after,
                "readonly_verified": b_sha_before == b_sha_after,
            },
        ],
    }
