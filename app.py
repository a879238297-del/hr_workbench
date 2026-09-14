"""
HR 智能工作台 · Flask 后端
启动: python app.py
访问: http://localhost:8080
"""
import io
import json
import os
import shutil
import sqlite3
import uuid
from datetime import datetime, timedelta, timezone
from functools import wraps
from pathlib import Path

import jwt
import openpyxl
from flask import Flask, g, jsonify, request, send_file, send_from_directory

from reconcile import reconcile

# ── 配置 ────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"
WEB_DIR = BASE_DIR / "web"

DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

DB_PATH = DATA_DIR / "hr_workbench.db"
SECRET_KEY = os.getenv("SECRET_KEY", "dev-secret-change-in-production")
JWT_EXPIRE_HOURS = 8
MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

# ── 数据库 ───────────────────────────────────────────────────────────────────
SCHEMA = """
CREATE TABLE IF NOT EXISTS app_user (
    id INTEGER PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password TEXT NOT NULL,
    display_name TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS job (
    id TEXT PRIMARY KEY,
    module_key TEXT NOT NULL DEFAULT 'payroll_recon',
    period TEXT,
    status TEXT NOT NULL DEFAULT 'done',
    summary TEXT,
    created_by TEXT NOT NULL,
    created_at TEXT NOT NULL,
    finished_at TEXT
);

CREATE TABLE IF NOT EXISTS job_file (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id TEXT NOT NULL,
    side TEXT NOT NULL,
    original_name TEXT NOT NULL,
    sha256_before TEXT NOT NULL,
    sha256_after TEXT NOT NULL,
    readonly_verified INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS job_row (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id TEXT NOT NULL,
    row_index INTEGER NOT NULL,
    match_key TEXT,
    category TEXT NOT NULL,
    payload TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_job_row ON job_row(job_id, category);

CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL,
    action TEXT NOT NULL,
    target_id TEXT,
    created_at TEXT NOT NULL
);
"""

SEED_USERS = [
    ("hr01",    "hr123456",    "李静（HR）"),
    ("fin01",   "fin123456",   "王芳（财务）"),
    ("admin",   "admin888",    "陈工（管理员）"),
]


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL")
        g.db.execute("PRAGMA foreign_keys=ON")
    return g.db


@app.teardown_appcontext
def close_db(_):
    db = g.pop("db", None)
    if db:
        db.close()


def init_db():
    con = sqlite3.connect(DB_PATH)
    con.executescript(SCHEMA)
    for username, password, display_name in SEED_USERS:
        con.execute(
            "INSERT OR IGNORE INTO app_user(username, password, display_name) VALUES(?,?,?)",
            (username, password, display_name),
        )
    con.commit()
    con.close()


# ── JWT ─────────────────────────────────────────────────────────────────────
def make_token(username: str) -> str:
    payload = {
        "sub": username,
        "exp": datetime.now(timezone.utc) + timedelta(hours=JWT_EXPIRE_HOURS),
    }
    return jwt.encode(payload, SECRET_KEY, algorithm="HS256")


def verify_token(token: str) -> str | None:
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=["HS256"])
        return payload["sub"]
    except jwt.PyJWTError:
        return None


def require_auth(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        auth = request.headers.get("Authorization", "")
        if not auth.startswith("Bearer "):
            return jsonify({"code": "UNAUTHENTICATED", "message": "请先登录"}), 401
        username = verify_token(auth[7:])
        if not username:
            return jsonify({"code": "UNAUTHENTICATED", "message": "Token 已过期，请重新登录"}), 401
        g.current_user = username
        return f(*args, **kwargs)
    return wrapper


# ── 工具 ─────────────────────────────────────────────────────────────────────
def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def err(code: str, message: str, status: int = 400, detail: dict | None = None):
    body = {"code": code, "message": message}
    if detail:
        body["detail"] = detail
    return jsonify(body), status


# ── 路由：静态文件 ───────────────────────────────────────────────────────────
@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_static(path):
    if path.startswith("api/"):
        return err("NOT_FOUND", "接口不存在", 404)
    target = WEB_DIR / (path or "index.html")
    if target.is_file():
        return send_from_directory(WEB_DIR, path or "index.html")
    return send_from_directory(WEB_DIR, "index.html")


# ── 路由：认证 ───────────────────────────────────────────────────────────────
@app.route("/api/auth/login", methods=["POST"])
def login():
    body = request.get_json(force=True, silent=True) or {}
    username = str(body.get("username", "")).strip()
    password = str(body.get("password", "")).strip()
    if not username or not password:
        return err("INVALID_PARAMS", "用户名和密码不能为空")

    db = get_db()
    row = db.execute(
        "SELECT * FROM app_user WHERE username=? AND is_active=1", (username,)
    ).fetchone()
    if not row or row["password"] != password:
        return err("AUTH_FAILED", "用户名或密码错误", 401)

    token = make_token(username)
    db.execute(
        "INSERT INTO audit_log(username, action, created_at) VALUES(?,?,?)",
        (username, "login", now_iso()),
    )
    db.commit()
    return jsonify({
        "access_token": token,
        "token_type": "bearer",
        "expires_in": JWT_EXPIRE_HOURS * 3600,
        "username": username,
        "display_name": row["display_name"],
    })


# ── 路由：模块列表（前端导航用） ────────────────────────────────────────────
MODULES = [
    {"key": "payroll_recon", "name": "薪资对账", "phase": 1, "enabled": True},
    {"key": "contract_alert", "name": "合同到期预警", "phase": 2, "enabled": False},
    {"key": "roster_check", "name": "花名册校验", "phase": 2, "enabled": False},
    {"key": "attendance", "name": "考勤统计", "phase": 2, "enabled": False},
    {"key": "social_ins", "name": "社保公积金核对", "phase": 2, "enabled": False},
    {"key": "tax_check", "name": "个税代扣核对", "phase": 2, "enabled": False},
    {"key": "recruit", "name": "招聘漏斗分析", "phase": 2, "enabled": False},
    {"key": "cost_analysis", "name": "人员成本分析", "phase": 2, "enabled": False},
    {"key": "performance", "name": "绩效考核汇总", "phase": 2, "enabled": False},
    {"key": "training", "name": "培训学时统计", "phase": 2, "enabled": False},
    {"key": "turnover", "name": "离职率分析", "phase": 2, "enabled": False},
    {"key": "overtime", "name": "加班费核算", "phase": 2, "enabled": False},
]


@app.route("/api/modules")
@require_auth
def list_modules():
    return jsonify(MODULES)


# ── 路由：作业 ───────────────────────────────────────────────────────────────
@app.route("/api/jobs", methods=["POST"])
@require_auth
def create_job():
    payroll_file = request.files.get("payroll_file")
    bank_file = request.files.get("bank_file")
    period = request.form.get("period", "").strip()

    if not payroll_file or not bank_file:
        return err("MISSING_FILE", "请同时上传工资表和银行代发明细")

    for f in [payroll_file, bank_file]:
        if not f.filename.lower().endswith((".xlsx", ".xls")):
            return err("FILE_TYPE_MISMATCH", f"仅支持 Excel 文件（.xlsx / .xls），收到：{f.filename}")

    job_id = str(uuid.uuid4())
    job_dir = UPLOAD_DIR / job_id
    job_dir.mkdir()

    payroll_path = job_dir / f"payroll_{payroll_file.filename}"
    bank_path = job_dir / f"bank_{bank_file.filename}"
    payroll_file.save(payroll_path)
    bank_file.save(bank_path)

    try:
        result = reconcile(payroll_path, bank_path)
    except ValueError as e:
        parts = str(e).split("|")
        code = parts[0] if len(parts) > 1 else "RECONCILE_ERROR"
        msg = parts[1] if len(parts) > 1 else str(e)
        detail = {}
        for p in parts[2:]:
            if "=" in p:
                k, v = p.split("=", 1)
                try:
                    detail[k] = json.loads(v)
                except Exception:
                    detail[k] = v
        shutil.rmtree(job_dir, ignore_errors=True)
        return err(code, msg, 422, detail or None)

    db = get_db()
    created_at = now_iso()
    db.execute(
        "INSERT INTO job(id, period, summary, created_by, created_at, finished_at)"
        " VALUES(?,?,?,?,?,?)",
        (job_id, period or None, json.dumps(result["summary"], ensure_ascii=False),
         g.current_user, created_at, created_at),
    )
    for f_info in result["files"]:
        f_meta = payroll_file if f_info["side"] == "payroll" else bank_file
        db.execute(
            "INSERT INTO job_file(job_id, side, original_name, sha256_before, sha256_after, readonly_verified)"
            " VALUES(?,?,?,?,?,?)",
            (job_id, f_info["side"], f_meta.filename,
             f_info["sha256_before"], f_info["sha256_after"],
             1 if f_info["readonly_verified"] else 0),
        )
    for i, row in enumerate(result["rows"]):
        db.execute(
            "INSERT INTO job_row(job_id, row_index, match_key, category, payload)"
            " VALUES(?,?,?,?,?)",
            (job_id, i, row.get("match_key"), row["category"],
             json.dumps(row, ensure_ascii=False)),
        )
    db.execute(
        "INSERT INTO audit_log(username, action, target_id, created_at) VALUES(?,?,?,?)",
        (g.current_user, "create_job", job_id, created_at),
    )
    db.commit()

    return jsonify(_job_response(job_id, period, result["summary"], created_at, result["files"],
                                 payroll_file.filename, bank_file.filename)), 201


@app.route("/api/jobs", methods=["GET"])
@require_auth
def list_jobs():
    db = get_db()
    page = max(1, int(request.args.get("page", 1)))
    page_size = min(50, int(request.args.get("page_size", 20)))
    offset = (page - 1) * page_size

    rows = db.execute(
        "SELECT * FROM job ORDER BY created_at DESC LIMIT ? OFFSET ?",
        (page_size, offset),
    ).fetchall()
    total = db.execute("SELECT COUNT(*) FROM job").fetchone()[0]

    items = []
    for r in rows:
        files = db.execute("SELECT * FROM job_file WHERE job_id=?", (r["id"],)).fetchall()
        items.append({
            "id": r["id"],
            "period": r["period"],
            "status": r["status"],
            "summary": json.loads(r["summary"]) if r["summary"] else None,
            "created_by": r["created_by"],
            "created_at": r["created_at"],
            "files": [dict(f) for f in files],
        })
    return jsonify({"items": items, "total": total, "page": page, "page_size": page_size})


@app.route("/api/jobs/<job_id>", methods=["GET"])
@require_auth
def get_job(job_id):
    db = get_db()
    job = db.execute("SELECT * FROM job WHERE id=?", (job_id,)).fetchone()
    if not job:
        return err("JOB_NOT_FOUND", f"作业 {job_id} 不存在", 404)
    files = db.execute("SELECT * FROM job_file WHERE job_id=?", (job_id,)).fetchall()
    return jsonify({
        "id": job["id"],
        "period": job["period"],
        "status": job["status"],
        "summary": json.loads(job["summary"]) if job["summary"] else None,
        "created_by": job["created_by"],
        "created_at": job["created_at"],
        "files": [dict(f) for f in files],
    })


@app.route("/api/jobs/<job_id>/rows", methods=["GET"])
@require_auth
def get_job_rows(job_id):
    db = get_db()
    if not db.execute("SELECT 1 FROM job WHERE id=?", (job_id,)).fetchone():
        return err("JOB_NOT_FOUND", f"作业 {job_id} 不存在", 404)

    category = request.args.getlist("category")  # 可多值
    keyword = request.args.get("keyword", "").strip()
    page = max(1, int(request.args.get("page", 1)))
    page_size = min(200, int(request.args.get("page_size", 100)))

    q = "SELECT payload FROM job_row WHERE job_id=?"
    params: list = [job_id]
    if category:
        placeholders = ",".join("?" * len(category))
        q += f" AND category IN ({placeholders})"
        params.extend(category)
    q += " ORDER BY row_index"

    all_rows = [json.loads(r["payload"]) for r in db.execute(q, params).fetchall()]

    if keyword:
        kw = keyword.lower()
        all_rows = [
            r for r in all_rows
            if kw in (r.get("match_key") or "").lower()
            or kw in (r.get("name") or "").lower()
        ]

    total = len(all_rows)
    offset = (page - 1) * page_size
    items = all_rows[offset: offset + page_size]
    return jsonify({"items": items, "total": total, "page": page, "page_size": page_size})


@app.route("/api/jobs/<job_id>/export", methods=["GET"])
@require_auth
def export_job(job_id):
    db = get_db()
    job = db.execute("SELECT * FROM job WHERE id=?", (job_id,)).fetchone()
    if not job:
        return err("JOB_NOT_FOUND", f"作业 {job_id} 不存在", 404)

    rows = [
        json.loads(r["payload"])
        for r in db.execute(
            "SELECT payload FROM job_row WHERE job_id=? ORDER BY row_index", (job_id,)
        ).fetchall()
    ]
    summary = json.loads(job["summary"]) if job["summary"] else {}

    wb = openpyxl.Workbook()
    _write_detail_sheet(wb.active, rows, summary)
    wb.active.title = "完整对账明细"

    ws2 = wb.create_sheet("差异清单")
    diff_rows = [r for r in rows if r["category"] != "match"]
    _write_diff_sheet(ws2, diff_rows)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    period = job["period"] or "未知期间"
    filename = f"薪资对账_{period}_{job['created_at'][:10]}.xlsx"

    db.execute(
        "INSERT INTO audit_log(username, action, target_id, created_at) VALUES(?,?,?,?)",
        (g.current_user, "export_result", job_id, now_iso()),
    )
    db.commit()

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


# ── Excel 导出辅助 ────────────────────────────────────────────────────────────
CATEGORY_LABELS = {
    "match": "一致",
    "amount_diff": "金额不符",
    "payroll_only": "工资表有银行无",
    "bank_only": "银行有工资表无",
}

DETAIL_HEADERS = ["工号", "姓名", "部门", "对账结果", "工资表实发金额", "银行代发金额", "差额（银行-工资表）"]
DIFF_HEADERS = DETAIL_HEADERS + ["差异说明"]


def _write_detail_sheet(ws, rows: list[dict], summary: dict):
    ws.append(DETAIL_HEADERS)
    for r in rows:
        ws.append([
            r.get("match_key") or "",
            r.get("name") or "",
            r.get("dept") or "",
            CATEGORY_LABELS.get(r["category"], r["category"]),
            r.get("payroll_amount") or "—",
            r.get("bank_amount") or "—",
            r.get("delta") or "—",
        ])


def _write_diff_sheet(ws, rows: list[dict]):
    ws.append(DIFF_HEADERS)
    for r in rows:
        cat = r["category"]
        delta = r.get("delta")
        if cat == "amount_diff" and delta:
            try:
                d = float(delta)
                desc = f"银行{'少' if d < 0 else '多'}发 {abs(d):.2f} 元"
            except Exception:
                desc = delta
        elif cat == "payroll_only":
            desc = "工资表有记录，银行无代发"
        else:
            desc = "银行有代发，工资表无记录"

        ws.append([
            r.get("match_key") or "",
            r.get("name") or "",
            r.get("dept") or "",
            CATEGORY_LABELS.get(cat, cat),
            r.get("payroll_amount") or "—",
            r.get("bank_amount") or "—",
            r.get("delta") or "—",
            desc,
        ])


def _job_response(job_id, period, summary, created_at, files_info, p_name, b_name):
    return {
        "id": job_id,
        "period": period or None,
        "status": "done",
        "summary": summary,
        "created_at": created_at,
        "files": [
            {
                "side": f["side"],
                "original_name": p_name if f["side"] == "payroll" else b_name,
                "sha256_before": f["sha256_before"],
                "sha256_after": f["sha256_after"],
                "readonly_verified": f["readonly_verified"],
            }
            for f in files_info
        ],
    }


# ── 启动 ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    init_db()
    print("=" * 50)
    print("HR 智能工作台 启动中...")
    print(f"访问地址: http://localhost:8080")
    print("内置账号: hr01/hr123456  fin01/fin123456  admin/admin888")
    print("=" * 50)
    app.run(host="0.0.0.0", port=8080, debug=False)
