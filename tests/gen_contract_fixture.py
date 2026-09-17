"""生成合同台账脱敏样例 Excel，日期相对于运行日计算，覆盖所有预警分类。"""
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl

OUT = Path(__file__).parent / "fixtures" / "合同台账_脱敏样例.xlsx"

today = date.today()

ROWS = [
    # 已过期 2 条
    ("GH0001", "张伟",  "技术部", "固定期限", today - timedelta(days=730), today - timedelta(days=10)),
    ("GH0002", "王芳",  "技术部", "固定期限", today - timedelta(days=365), today - timedelta(days=45)),
    # 30天内到期 3 条
    ("GH0003", "李娜",  "技术部", "固定期限", today - timedelta(days=700), today + timedelta(days=5)),
    ("GH0004", "刘洋",  "技术部", "固定期限", today - timedelta(days=365), today + timedelta(days=18)),
    ("GH0005", "陈晨",  "销售部", "固定期限", today - timedelta(days=730), today + timedelta(days=28)),
    # 60天内到期 2 条
    ("GH0006", "赵静",  "销售部", "固定期限", today - timedelta(days=700), today + timedelta(days=35)),
    ("GH0007", "赵敏",  "技术部", "固定期限", today - timedelta(days=365), today + timedelta(days=55)),
    # 90天内到期 2 条
    ("GH0008", "孙涛",  "运营部", "固定期限", today - timedelta(days=365), today + timedelta(days=65)),
    ("GH0009", "周丽",  "人事部", "固定期限", today - timedelta(days=730), today + timedelta(days=85)),
    # 正常 10 条
    ("GH0010", "吴强",  "财务部", "固定期限", today - timedelta(days=365), today + timedelta(days=120)),
    ("GH0011", "郑超",  "技术部", "固定期限", today - timedelta(days=300), today + timedelta(days=180)),
    ("GH0012", "冯雪",  "销售部", "固定期限", today - timedelta(days=200), today + timedelta(days=250)),
    ("GH0013", "褚磊",  "运营部", "固定期限", today - timedelta(days=365), today + timedelta(days=365)),
    ("GH0014", "卫萍",  "人事部", "固定期限", today - timedelta(days=100), today + timedelta(days=630)),
    ("GH0015", "蒋明",  "技术部", "固定期限", today - timedelta(days=730), today + timedelta(days=365)),
    ("GH0016", "沈婷",  "财务部", "固定期限", today - timedelta(days=180), today + timedelta(days=550)),
    ("GH0017", "韩刚",  "销售部", "固定期限", today - timedelta(days=365), today + timedelta(days=400)),
    ("GH0018", "杨帆",  "技术部", "固定期限", today - timedelta(days=500), today + timedelta(days=230)),
    ("GH0019", "朱琳",  "运营部", "固定期限", today - timedelta(days=365), today + timedelta(days=180)),
    # 无固定期限 1 条
    ("GH0020", "秦伟",  "技术部", "无固定期限", today - timedelta(days=1000), None),
]

HEADERS = ["工号", "姓名", "部门", "合同类型", "合同开始日期", "合同结束日期"]


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "合同台账"
    ws.append(HEADERS)
    for emp_id, name, dept, ctype, start, end in ROWS:
        ws.append([emp_id, name, dept, ctype, start, end])
    wb.save(OUT)
    print(f"已生成 {OUT}，共 {len(ROWS)} 行数据")


if __name__ == "__main__":
    main()
