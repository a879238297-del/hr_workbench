"""生成打卡明细脱敏样例 Excel，5人×22天，覆盖各类异常。"""
import random
from datetime import date, time, timedelta
from pathlib import Path

import openpyxl

OUT = Path(__file__).parent / "fixtures" / "打卡明细_脱敏样例.xlsx"

HEADERS = ["日期", "工号", "姓名", "部门", "上班打卡", "下班打卡"]

EMPLOYEES = [
    ("GH0001", "张伟", "技术部"),
    ("GH0002", "王芳", "技术部"),
    ("GH0003", "李娜", "销售部"),
    ("GH0004", "刘洋", "运营部"),
    ("GH0005", "陈晨", "财务部"),
]

random.seed(42)

START_DATE = date(2026, 9, 1)


def workdays(start, count):
    days = []
    d = start
    while len(days) < count:
        if d.weekday() < 5:
            days.append(d)
        d += timedelta(days=1)
    return days


DAYS = workdays(START_DATE, 22)


def make_time(h, m):
    return time(h, m)


def gen_rows():
    rows = []
    for emp_id, name, dept in EMPLOYEES:
        for d in DAYS:
            r = random.random()
            if r < 0.08:
                # 缺卡（上班）
                clock_in = None
                clock_out = make_time(17, random.randint(25, 40))
            elif r < 0.14:
                # 缺卡（下班）
                clock_in = make_time(8, random.randint(20, 30))
                clock_out = None
            elif r < 0.24:
                # 迟到
                clock_in = make_time(8, random.randint(35, 59))
                clock_out = make_time(17, random.randint(30, 45))
            elif r < 0.30:
                # 早退
                clock_in = make_time(8, random.randint(15, 29))
                clock_out = make_time(16, random.randint(30, 59))
            elif r < 0.38:
                # 加班
                clock_in = make_time(8, random.randint(15, 28))
                clock_out = make_time(random.choice([19, 20, 21]), random.randint(0, 50))
            else:
                # 正常
                clock_in = make_time(8, random.randint(5, 28))
                clock_out = make_time(17, random.randint(30, 45))
            rows.append((d, emp_id, name, dept, clock_in, clock_out))
    return rows


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "打卡明细"
    ws.append(HEADERS)
    rows = gen_rows()
    for d, emp_id, name, dept, cin, cout in rows:
        ws.append([d, emp_id, name, dept, cin, cout])
    wb.save(OUT)
    print(f"已生成 {OUT}，共 {len(rows)} 行数据")


if __name__ == "__main__":
    main()
