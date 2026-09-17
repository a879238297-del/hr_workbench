# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

华辰精密制造 HR 智能工作台（一期）—— 薪资对账模块。内网运行的 Web 应用，用于将工资表与银行代发明细进行自动对账，输出四类结果（一致、金额不符、工资表有银行无、银行有工资表无）。

## 启动与运行

```bash
# 安装依赖
pip install -r requirements.txt

# 启动服务（端口 8080）
python app.py

# Windows 双击启动
start.bat
```

访问 http://localhost:8080，内置账号：hr01/hr123456、fin01/fin123456、admin/admin888

## 技术栈

- **后端**：Python 3.11 + Flask（注意：技术方案文档写的是 FastAPI，实际实现用的 Flask）
- **数据库**：SQLite（`data/hr_workbench.db`），直接用 `sqlite3` 模块，无 ORM
- **前端**：单文件 `web/index.html`，Vanilla HTML/CSS/JS，零外部依赖，无构建步骤
- **Excel**：openpyxl（`read_only=True` 读取，写入模式导出）
- **认证**：JWT（PyJWT），8 小时过期
- **金额精度**：全链路 `decimal.Decimal`，禁止 float

## 架构

```
app.py          Flask 主应用：路由、JWT 认证、DB 初始化、Excel 导出
reconcile.py    薪资对账核心逻辑：Excel 读取、表头探测、工号归一、金额匹配
web/index.html  前端单页应用（CSS+JS 全内联）
data/           SQLite 数据库存放目录
uploads/        上传文件按 job_id 子目录存放
tests/fixtures/ 测试用 Excel 样例文件
doc/            需求文档、技术方案（doc/HR智能工作台技术方案.md 是技术基准文档）
```

### 数据流

1. 前端上传两个 Excel（工资表 + 银行代发明细）→ `POST /api/jobs`
2. `app.py` 保存文件到 `uploads/{job_id}/`，调用 `reconcile.reconcile()`
3. `reconcile.py`：SHA-256 校验文件完整性 → 探测表头行（前 12 行）→ 列别名映射 → 工号归一（去空格、全角转半角、大写）→ 金额 Decimal 解析 → 双表匹配
4. 结果写入 `job`、`job_file`、`job_row` 表，返回前端渲染

### 数据库表

- `app_user`：用户账号（明文密码，一期内网简化）
- `job`：对账作业记录
- `job_file`：作业关联的文件元信息（含 SHA-256）
- `job_row`：逐行对账结果（payload 为 JSON）
- `audit_log`：操作审计日志

### API 路由

- `POST /api/auth/login` — 登录
- `GET /api/modules` — 模块列表
- `POST /api/jobs` — 创建对账作业（multipart 上传）
- `GET /api/jobs` — 作业列表（分页）
- `GET /api/jobs/<id>` — 作业详情
- `GET /api/jobs/<id>/rows` — 对账明细行（分页、分类筛选、关键字搜索）
- `GET /api/jobs/<id>/export` — 导出 Excel

## 硬约束

1. **文件只读**：原始 Excel 只读不回写（`read_only=True` + SHA-256 前后校验证明未修改）
2. **金额禁用 float**：全链路 `decimal.Decimal`，包括解析、比较、输出
3. **内网运行**：不依赖外部 CDN、API 或网络资源
4. **表头探测**：表头行位置不固定，扫描前 12 行匹配别名
5. **列名别名**：同一字段有多种写法（如"实发合计"/"实发工资"/"实发金额"），通过 `PAYROLL_ALIASES` / `BANK_ALIASES` 字典映射

## reconcile.py 关键设计

- `normalize_id()`：工号归一（去空格、全角→半角、大写）
- `normalize_amount()`：金额归一（去千位符、货币符号、单位后缀，转 Decimal）
- `_detect_header()`：表头行探测，返回最佳匹配行及列映射
- 重复工号直接报错阻断（`DUPLICATE_KEY`），不静默合并
- 空工号行归入单边缺失
- 错误通过 `ValueError` 抛出，格式：`CODE|消息|key=value` 管道分隔

## 前端注意事项

- 单文件 `index.html` 内联所有 CSS 和 JS，不拆分
- 页面状态通过 `.hidden` CSS 类切换，无路由库
- 全局状态对象 `S`（token、currentJobId、分页等）
- XSS 防护：`esc()` 函数转义用户内容
- 步骤条三步：上传与核对 → 结果与差异 → 历史作业
