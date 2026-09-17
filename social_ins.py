"""
社保公积金核对核心逻辑
- 读取工资表 + 社保申报表（read_only=True）+ SHA-256 校验
- 表头探测（前 12 行）+ 列别名映射
- 工号归一 → 按工号匹配
- 金额 Decimal 比较基数和金额
- 输出 5 类结果：match / base_diff / amount_diff / payroll_only / social_only
"""
import hashlib
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path

import openpyxl

MAX_HEADER_SCAN_ROWS = 12

PAYROLL_SI_ALIASES = {
    "emp_id":      {"工号", "员工编号", "职工编号", "员工工号"},
    "name":        {"姓名", "员工姓名"},
    "dept":        {"部门", "部门名称", "所在部门"},
    "si_base":     {"社保基数", "社保缴费基数", "社保缴纳基数"},
    "hf_base":     {"公积金基数", "公积金缴费基数", "住房公积金基数"},
    "si_person":   {"社保个人", "社保个人合计", "个人社保"},
    "si_company":  {"社保单位", "社保单位合计", "单位社保"},
    "hf_person":   {"公积金个人", "个人公积金"},
    "hf_company":  {"公积金单位", "单位公积金"},
}

SOCIAL_ALIASES = {
    "emp_id":      {"工号", "员工编号", "职工编号", "员工工号"},
    "name":        {"姓名", "员工姓名"},
    "si_base":     {"社保基数", "社保缴费基数", "社保缴纳基数", "社保申报基数"},
    "hf_base":     {"公积金基数", "公积金缴费基数", "住房公积金基数", "公积金申报基数"},
    "si_person":   {"社保个人", "社保个人合计", "个人社保"},
    "si_company":  {"社保单位", "社保单位合计", "单位社保"},
    "hf_person":   {"公积金个人", "个人公积金"},
    "hf_company":  {"公积金单位", "单位公积金"},
}

COMPARE_FIELDS = ["si_base", "hf_base", "si_person", "si_company", "hf_person", "hf_company"]
BASE_FIELDS = {"si_base", "hf_base"}
AMOUNT_FIELDS = {"si_person", "si_company", "hf_person", "hf_company"}

FIELD_LABELS = {
    "si_base": "社保基数",
    "hf_base": "公积金基数",
    "si_person": "社保个人",
    "si_company": "社保单位",
    "hf_person": "公积金个人",
    "hf_company": "公积金单位",
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


def _normalize_amount(raw) -> Decimal | None:
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
    except InvalidOperation:
        return None


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


def _read_table(path: Path, aliases: dict[str, set]) -> tuple[list[dict], str, str]:
    sha_before = sha256_of(path)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = [[cell.value for cell in row] for row in ws.rows]
    wb.close()
    sha_after = sha256_of(path)

    alias_map = _build_alias_map(aliases)
    header_idx, col_map = _detect_header(all_rows, alias_map)

    if "emp_id" not in col_map:
        raise ValueError("MISSING_REQUIRED_COLUMN|缺少工号列")

    rows = []
    for row in all_rows[header_idx + 1:]:
        if all(c is None for c in row):
            continue
        record: dict = {}
        for field, idx in col_map.items():
            if idx < len(row):
                record[field] = row[idx]
        record["emp_id"] = normalize_id(record.get("emp_id"))
        for f in COMPARE_FIELDS:
            record[f] = _normalize_amount(record.get(f))
        rows.append(record)

    return rows, sha_before, sha_after


def reconcile(payroll_path: Path, social_path: Path) -> dict:
    p_rows, p_sha_before, p_sha_after = _read_table(payroll_path, PAYROLL_SI_ALIASES)
    s_rows, s_sha_before, s_sha_after = _read_table(social_path, SOCIAL_ALIASES)

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
    s_keyed, s_empty = group_by_key(s_rows)

    for side, keyed in [("payroll", p_keyed), ("social", s_keyed)]:
        dups = {k for k, v in keyed.items() if len(v) > 1}
        if dups:
            raise ValueError(f"DUPLICATE_KEY|{side} 侧存在重复工号：{sorted(dups)}")

    results = []
    all_keys = sorted(set(p_keyed) | set(s_keyed))

    for key in all_keys:
        p = p_keyed[key][0] if key in p_keyed else None
        s = s_keyed[key][0] if key in s_keyed else None

        row = {"match_key": key}

        if p and s:
            row["name"] = p.get("name") or s.get("name")
            row["dept"] = p.get("dept")

            diff_fields = []
            has_base_diff = False
            has_amount_diff = False

            for f in COMPARE_FIELDS:
                pv = p.get(f)
                sv = s.get(f)
                row["payroll_" + f] = str(pv) if pv is not None else None
                row["social_" + f] = str(sv) if sv is not None else None
                if pv is not None and sv is not None:
                    delta = sv - pv
                    row[f + "_delta"] = str(delta) if delta != 0 else None
                    if delta != 0:
                        diff_fields.append(FIELD_LABELS[f])
                        if f in BASE_FIELDS:
                            has_base_diff = True
                        else:
                            has_amount_diff = True
                else:
                    row[f + "_delta"] = None

            if has_base_diff:
                row["category"] = "base_diff"
            elif has_amount_diff:
                row["category"] = "amount_diff"
            else:
                row["category"] = "match"
            row["diff_fields"] = diff_fields

        elif p:
            row["name"] = p.get("name")
            row["dept"] = p.get("dept")
            row["category"] = "payroll_only"
            row["diff_fields"] = []
            for f in COMPARE_FIELDS:
                pv = p.get(f)
                row["payroll_" + f] = str(pv) if pv is not None else None
                row["social_" + f] = None
                row[f + "_delta"] = None

        else:
            row["name"] = s.get("name")
            row["dept"] = None
            row["category"] = "social_only"
            row["diff_fields"] = []
            for f in COMPARE_FIELDS:
                sv = s.get(f)
                row["payroll_" + f] = None
                row["social_" + f] = str(sv) if sv is not None else None
                row[f + "_delta"] = None

        results.append(row)

    for r in p_empty:
        row = {
            "match_key": None, "name": r.get("name"), "dept": r.get("dept"),
            "category": "payroll_only", "diff_fields": [],
        }
        for f in COMPARE_FIELDS:
            pv = r.get(f)
            row["payroll_" + f] = str(pv) if pv is not None else None
            row["social_" + f] = None
            row[f + "_delta"] = None
        results.append(row)

    for r in s_empty:
        row = {
            "match_key": None, "name": r.get("name"), "dept": None,
            "category": "social_only", "diff_fields": [],
        }
        for f in COMPARE_FIELDS:
            sv = r.get(f)
            row["payroll_" + f] = None
            row["social_" + f] = str(sv) if sv is not None else None
            row[f + "_delta"] = None
        results.append(row)

    summary = {
        "match": sum(1 for r in results if r["category"] == "match"),
        "base_diff": sum(1 for r in results if r["category"] == "base_diff"),
        "amount_diff": sum(1 for r in results if r["category"] == "amount_diff"),
        "payroll_only": sum(1 for r in results if r["category"] == "payroll_only"),
        "social_only": sum(1 for r in results if r["category"] == "social_only"),
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
                "side": "social",
                "sha256_before": s_sha_before,
                "sha256_after": s_sha_after,
                "readonly_verified": s_sha_before == s_sha_after,
            },
        ],
    }
