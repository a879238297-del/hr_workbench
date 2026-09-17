"""生成员工花名册脱敏样例 Excel，覆盖所有校验分类。"""
from datetime import date, timedelta
from pathlib import Path

import openpyxl

OUT = Path(__file__).parent / "fixtures" / "员工花名册_脱敏样例.xlsx"

HEADERS = ["工号", "姓名", "性别", "出生日期", "身份证号", "部门", "入职日期"]

# 虚假身份证号生成辅助（地区码 110101 = 北京东城，校验码简化处理）
def fake_id(birth_str, gender_male=True, seq=1):
    """生成格式正确的虚假18位身份证号。birth_str: YYYYMMDD"""
    area = "110101"
    # 顺序码2位 + 第17位（奇数=男，偶数=女）
    seq2 = seq % 100
    if gender_male:
        digit17 = (seq2 % 5) * 2 + 1  # 1,3,5,7,9
    else:
        digit17 = (seq2 % 5) * 2 + 2  # 2,4,6,8,0
        if digit17 == 10:
            digit17 = 0
    base = area + birth_str + f"{seq2:02d}" + str(digit17)
    # 简化校验码计算
    weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    check_chars = "10X98765432"
    total = sum(int(base[i]) * weights[i] for i in range(17))
    check = check_chars[total % 11]
    return base + check


ROWS = [
    # 正常记录 15 条
    ("GH0001", "张伟",  "男", date(1990, 3, 15), fake_id("19900315", True, 1),  "技术部", date(2018, 6, 1)),
    ("GH0002", "王芳",  "女", date(1992, 7, 20), fake_id("19920720", False, 1), "技术部", date(2019, 3, 1)),
    ("GH0003", "李娜",  "女", date(1988, 11, 5), fake_id("19881105", False, 2), "销售部", date(2017, 9, 15)),
    ("GH0004", "刘洋",  "男", date(1995, 1, 10), fake_id("19950110", True, 2),  "销售部", date(2020, 7, 1)),
    ("GH0005", "陈晨",  "男", date(1991, 5, 25), fake_id("19910525", True, 3),  "运营部", date(2019, 1, 1)),
    ("GH0006", "赵静",  "女", date(1993, 8, 30), fake_id("19930830", False, 3), "运营部", date(2020, 4, 1)),
    ("GH0007", "孙涛",  "男", date(1987, 12, 1), fake_id("19871201", True, 4),  "人事部", date(2016, 5, 1)),
    ("GH0008", "周丽",  "女", date(1994, 4, 18), fake_id("19940418", False, 4), "人事部", date(2021, 2, 1)),
    ("GH0009", "吴强",  "男", date(1989, 9, 22), fake_id("19890922", True, 5),  "财务部", date(2018, 10, 1)),
    ("GH0010", "郑超",  "男", date(1996, 2, 14), fake_id("19960214", True, 6),  "技术部", date(2022, 1, 1)),
    ("GH0011", "冯雪",  "女", date(1990, 6, 8),  fake_id("19900608", False, 5), "销售部", date(2019, 8, 1)),
    ("GH0012", "褚磊",  "男", date(1985, 10, 3), fake_id("19851003", True, 7),  "运营部", date(2015, 3, 1)),
    ("GH0013", "卫萍",  "女", date(1997, 3, 28), fake_id("19970328", False, 6), "财务部", date(2022, 6, 1)),
    ("GH0014", "蒋明",  "男", date(1992, 11, 11),fake_id("19921111", True, 8),  "技术部", date(2020, 9, 1)),
    ("GH0015", "沈婷",  "女", date(1991, 7, 7),  fake_id("19910707", False, 7), "人事部", date(2019, 5, 1)),

    # 缺失字段 3 条
    ("GH0016", None,     "男", date(1993, 1, 1),  fake_id("19930101", True, 9),  "技术部", date(2021, 1, 1)),   # 缺姓名
    ("GH0017", "韩刚",  "男", date(1990, 5, 5),  None,                           "销售部", date(2020, 3, 1)),   # 缺身份证
    ("GH0018", "杨帆",  "女", date(1994, 8, 20), fake_id("19940820", False, 8),  "运营部", None),              # 缺入职日期

    # 工号重复 2 条（GH0001 重复）
    ("GH0001", "朱琳",  "女", date(1995, 12, 1), fake_id("19951201", False, 9),  "财务部", date(2022, 11, 1)),

    # 身份证重复 2 条（和 GH0002 的身份证相同）
    ("GH0020", "秦伟",  "男", date(1992, 7, 20), fake_id("19920720", False, 1),  "技术部", date(2023, 1, 1)),

    # 性别矛盾 2 条（身份证17位与性别列不符）
    ("GH0021", "许峰",  "女", date(1991, 4, 10), fake_id("19910410", True, 10),  "销售部", date(2020, 6, 1)),   # 证件号=男，性别列=女
    ("GH0022", "何芳",  "男", date(1993, 9, 15), fake_id("19930915", False, 10), "运营部", date(2021, 8, 1)),   # 证件号=女，性别列=男

    # 生日矛盾 1 条（身份证中生日与出生日期列不同）
    ("GH0023", "吕敏",  "女", date(1990, 3, 15), fake_id("19920601", False, 11), "人事部", date(2019, 12, 1)),  # 出生日期=1990-03-15，证件中=1992-06-01

    # 额外正常 2 条凑满 25
    ("GH0024", "施华",  "男", date(1988, 2, 28), fake_id("19880228", True, 11),  "技术部", date(2017, 4, 1)),
    ("GH0025", "曹洁",  "女", date(1996, 10, 10),fake_id("19961010", False, 12), "财务部", date(2023, 3, 1)),
]


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "员工花名册"
    ws.append(HEADERS)
    for row in ROWS:
        ws.append(list(row))
    wb.save(OUT)
    print(f"已生成 {OUT}，共 {len(ROWS)} 行数据")


if __name__ == "__main__":
    main()
