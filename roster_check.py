"""
花名册校验核心逻辑
- 读取员工花名册 Excel（read_only=True）+ SHA-256 校验
- 表头探测（扫描前 12 行）+ 列别名映射
- 三类检查：必填字段缺失、工号/证件号重复、性别与证件号矛盾
- 身份证号脱敏输出
"""
import hashlib
import re
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import openpyxl

MAX_HEADER_SCAN_ROWS = 12

ROSTER_ALIASES = {
    "emp_id":     {"工号", "员工编号", "职工编号", "员工工号"},
    "name":       {"姓名", "员工姓名"},
    "gender":     {"性别"},
    "birth_date": {"出生日期", "出生年月", "生日"},
    "id_number":  {"身份证号", "身份证号码", "证件号码", "证件号"},
    "dept":       {"部门", "部门名称", "所在部门"},
    "hire_date":  {"入职日期", "入职时间", "报到日期"},
}

REQUIRED_FIELDS = {"emp_id", "name", "gender", "birth_date", "id_number", "dept", "hire_date"}

FIELD_LABELS = {
    "emp_id": "工号", "name": "姓名", "gender": "性别",
    "birth_date": "出生日期", "id_number": "身份证号",
    "dept": "部门", "hire_date": "入职日期",
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
    return None


def _mask_id_number(id_num: str) -> str:
    if not id_num or len(id_num) < 7:
        return id_num or ""
    return id_num[:3] + "*" * (len(id_num) - 7) + id_num[-4:]


def _normalize_gender(raw) -> str | None:
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("男", "M", "m", "male", "Male"):
        return "男"
    if s in ("女", "F", "f", "female", "Female"):
        return "女"
    return s if s else None


def validate(file_path: Path) -> dict:
    path = Path(file_path)
    sha_before = sha256_of(path)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = [[cell.value for cell in row] for row in ws.rows]
    wb.close()

    sha_after = sha256_of(path)

    alias_map = _build_alias_map(ROSTER_ALIASES)
    header_idx, col_map = _detect_header(all_rows, alias_map)

    # 解析所有数据行
    records = []
    for row in all_rows[header_idx + 1:]:
        if all(c is None for c in row):
            continue
        record: dict = {}
        for field, idx in col_map.items():
            if idx < len(row):
                record[field] = row[idx]
        records.append(record)

    # 第一遍：标准化 + 缺失检查
    rows_out = []
    for rec in records:
        emp_id = normalize_id(rec.get("emp_id"))
        name = str(rec.get("name") or "").strip() or None
        gender = _normalize_gender(rec.get("gender"))
        birth_date = _parse_date(rec.get("birth_date"))
        id_number = str(rec.get("id_number") or "").strip() or None
        dept = str(rec.get("dept") or "").strip() or None
        hire_date = _parse_date(rec.get("hire_date"))

        parsed = {
            "emp_id": emp_id,
            "name": name,
            "gender": gender,
            "birth_date": birth_date.isoformat() if birth_date else None,
            "id_number": id_number,
            "id_number_masked": _mask_id_number(id_number) if id_number else None,
            "dept": dept,
            "hire_date": hire_date.isoformat() if hire_date else None,
        }

        # 缺失检查
        missing = []
        field_vals = {
            "emp_id": emp_id, "name": name, "gender": gender,
            "birth_date": birth_date, "id_number": id_number,
            "dept": dept, "hire_date": hire_date,
        }
        for f in REQUIRED_FIELDS:
            if f in col_map and field_vals.get(f) is None:
                missing.append(FIELD_LABELS.get(f, f))
            elif f not in col_map:
                missing.append(FIELD_LABELS.get(f, f))

        parsed["_missing"] = missing
        parsed["_birth_date_obj"] = birth_date
        parsed["_id_number_raw"] = id_number
        parsed["_gender"] = gender
        rows_out.append(parsed)

    # 第二遍：重复检查
    emp_id_counts = Counter(r["emp_id"] for r in rows_out if r["emp_id"])
    id_num_counts = Counter(r["id_number"] for r in rows_out if r["id_number"])
    dup_emp_ids = {k for k, v in emp_id_counts.items() if v > 1}
    dup_id_nums = {k for k, v in id_num_counts.items() if v > 1}

    # 第三遍：矛盾检查 + 构建最终结果
    results = []
    for r in rows_out:
        issues = []

        # 缺失
        if r["_missing"]:
            issues.append({
                "type": "missing_field",
                "detail": "缺少：" + "、".join(r["_missing"]),
            })

        # 重复
        dup_types = []
        if r["emp_id"] and r["emp_id"] in dup_emp_ids:
            dup_types.append("工号")
        if r["id_number"] and r["id_number"] in dup_id_nums:
            dup_types.append("身份证号")
        if dup_types:
            issues.append({
                "type": "duplicate",
                "detail": "重复：" + "、".join(dup_types),
            })

        # 矛盾（仅18位身份证号）
        id_num = r["_id_number_raw"]
        if id_num and len(id_num) == 18 and re.match(r'^\d{17}[\dXx]$', id_num):
            # 性别矛盾
            gender = r["_gender"]
            if gender:
                id_gender_digit = int(id_num[16])
                id_gender = "男" if id_gender_digit % 2 == 1 else "女"
                if gender != id_gender:
                    issues.append({
                        "type": "contradiction",
                        "detail": "性别矛盾：填写[" + gender + "]，证件号显示[" + id_gender + "]",
                    })

            # 生日矛盾
            birth = r["_birth_date_obj"]
            if birth:
                id_birth_str = id_num[6:14]
                try:
                    id_birth = datetime.strptime(id_birth_str, "%Y%m%d").date()
                    if birth != id_birth:
                        issues.append({
                            "type": "contradiction",
                            "detail": "生日矛盾：填写[" + birth.isoformat() + "]，证件号显示[" + id_birth.isoformat() + "]",
                        })
                except ValueError:
                    pass

        # 确定 category（取最严重的）
        if not issues:
            category = "clean"
            issue_summary = ""
        else:
            type_priority = {"missing_field": 0, "duplicate": 1, "contradiction": 2}
            issues.sort(key=lambda x: type_priority.get(x["type"], 99))
            category = issues[0]["type"]
            issue_summary = "；".join(i["detail"] for i in issues)

        row_out = {
            "match_key": r["emp_id"],
            "name": r["name"],
            "gender": r["gender"],
            "birth_date": r["birth_date"],
            "id_number": r["id_number"],
            "id_number_masked": r["id_number_masked"],
            "dept": r["dept"],
            "hire_date": r["hire_date"],
            "category": category,
            "issues": issues,
            "issue_summary": issue_summary,
        }
        results.append(row_out)

    summary = {
        "total": len(results),
        "missing_field": sum(1 for r in results if any(i["type"] == "missing_field" for i in r["issues"])),
        "duplicate": sum(1 for r in results if any(i["type"] == "duplicate" for i in r["issues"])),
        "contradiction": sum(1 for r in results if any(i["type"] == "contradiction" for i in r["issues"])),
        "clean": sum(1 for r in results if r["category"] == "clean"),
    }

    return {
        "summary": summary,
        "rows": results,
        "files": [
            {
                "side": "roster",
                "sha256_before": sha_before,
                "sha256_after": sha_after,
                "readonly_verified": sha_before == sha_after,
            },
        ],
    }
