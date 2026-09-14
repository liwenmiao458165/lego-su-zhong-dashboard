# -*- coding: utf-8 -*-
"""苏中区域门店本周重点工作与行动计划（9/14-9/20）xlsx 生成脚本"""
import subprocess, sys

try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from datetime import date


def xl_color(css_hex: str) -> str:
    value = css_hex.removeprefix("#").upper()
    if len(value) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + value


XL_HEADER = xl_color("#4472C4")      # 商务蓝表头
XL_TITLE = xl_color("#2F5597")       # 深蓝标题
XL_LIGHT = xl_color("#D9E2F3")       # 浅蓝
XL_BORDER = xl_color("#BFBFBF")
XL_GREEN_BG = xl_color("#C6EFCE")
XL_GREEN_FT = xl_color("#006100")
XL_RED_BG = xl_color("#FFC7CE")
XL_RED_FT = xl_color("#9C0006")
XL_YELLOW_BG = xl_color("#FFEB9C")
XL_YELLOW_FT = xl_color("#9C6500")

thin_side = Side(style="thin", color=XL_BORDER)
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

font_title = Font(name="PingFang SC", size=14, bold=True, color="FFFFFFFF")
font_sub = Font(name="PingFang SC", size=10, color="FF595959")
font_header = Font(name="PingFang SC", size=11, bold=True, color="FFFFFFFF")
font_body = Font(name="PingFang SC", size=10.5, color="FF1A1A1A")

fill_title = PatternFill("solid", fgColor=XL_TITLE)
fill_header = PatternFill("solid", fgColor=XL_HEADER)

align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_r = Alignment(horizontal="right", vertical="center")

wb = openpyxl.Workbook()

# ============ Sheet1 本周行动计划 ============
ws = wb.active
ws.title = "本周行动计划"

# 锚点：第1行标题、第2行口径说明、第3行表头、第4-24行数据（21 条）
ws.merge_cells("A1:G1")
ws["A1"] = "苏中区域门店本周重点工作与行动计划（9/14-9/20）"
ws["A1"].font = font_title
ws["A1"].fill = fill_title
ws["A1"].alignment = align_c
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:G2")
ws["A2"] = "数据口径：第 37 周（9/7-9/13）周分析，截止 9/13；75 折限制促销 9/18 截止，长尾款口径 20 款；淮安新亚陈方请假中（相关事项由代班负责人跟进）；淮安万象城为新店无同比。"
ws["A2"].font = font_sub
ws["A2"].alignment = align_l
ws.row_dimensions[2].height = 30

headers = ["序号", "门店", "重点工作（机会点）", "行动计划（动作+目标+时限）", "截止日期", "状态", "备注"]
for col, name in zip("ABCDEFG", headers):
    c = ws[f"{col}3"]
    c.value = name
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_c
    c.border = thin_border
ws.row_dimensions[3].height = 24

# (门店, 机会点, 行动计划, 截止日期)
D0915, D0916, D0917, D0918, D0920 = (date(2026, 9, d) for d in (15, 16, 17, 18, 20))
actions = [
    ("扬州万象汇", "周销 ¥62,107 同比 -46.1%，超区域降幅（-36%）", "核对去年同期团购/大单基数，梳理可复制动作同步城市经理", D0915),
    ("扬州万象汇", "连带率 1.30 低于区域均值 1.37", "收银台加推 ¥99-199 搭售款（人仔/小车），连带率目标 ≥1.4，周日晚复盘", D0920),
    ("扬州万象汇", "75 折限制促销 9/18 截止", "周四前检查 20 款长尾款堆头与价签到位，9/19 周五恢复原价签并拍照留档", D0918),
    ("扬州京华城", "环比 -21.3%（区域 -12.6%）", "排查上周大单透支与客流变化，结论同步城市经理", D0915),
    ("扬州京华城", "陆文一周销 ¥7,030 仅店人均 41%、客单价 ¥281 全店最低", "丁茜带教高价款推荐+收银搭售，周五检视，周销目标 ≥¥10,000", D0918),
    ("扬州京华城", "促销尾段执行", "9/18 前主推位集中高客单 TECHNIC/超跑系列，活动结束周五恢复价签拍照留档", D0918),
    ("扬州江都金鹰", "完成率 70.6% 低于区域 9.8 个百分点、客单价 ¥345 低 ¥46", "主推 ¥500+ 合款与会员大单邀约，周完成率目标 ≥80%", D0920),
    ("扬州江都金鹰", "邰宇/朱奕然/谭欣蔚周销均低于区域人均 70%", "王伟一对一带教连带推荐，周五检视，各自周销目标 ≥¥7,000", D0918),
    ("扬州江都金鹰", "环比 -25%", "排查客流与缺码断货情况，缺货清单同步补货", D0915),
    ("泰州万象城", "长尾款 ¥3,281 仅区域均值 66%、长尾指标 77.5%（区域唯一未达标）", "收银台+端架调整 20 款长尾款陈列，周长尾目标 ≥¥4,234，周日晚复盘", D0920),
    ("泰州万象城", "植物系列占比 3.5% 低于 4.1% 对标线", "花束/盆栽上移收银动线，本周占比目标 ≥4.5%", D0920),
    ("泰州万象城", "黄小莉（泰州支援）周销 ¥10,394 为店人均 69%", "聚焦 ¥1,000+ 高价款开口，周销目标 ≥¥12,000", D0920),
    ("宿迁宝龙", "周销 ¥12,157 环比 -50%、同比 -57%、完成率 49.5%（区域 80.4%）", "排查客流/排班/活动断档主因，出追赶方案同步城市经理", D0916),
    ("宿迁宝龙", "客单价 ¥270 低于区域均值 ¥120", "主推 ¥299-499 中价位套组+老客复购邀约，客单价目标 ≥¥330", D0920),
    ("宿迁宝龙", "长尾款 ¥207 仅区域门店均值 4%", "20 款长尾款上端架+收银台陈列，周长尾目标 ≥¥983（周指标额），周五检视", D0918),
    ("淮安新亚", "长尾款 ¥3,182 低于区域门店均值×0.7（¥3,461）", "20 款长尾款调至主推位，周长尾目标 ≥¥4,200，周日复盘", D0920),
    ("淮安新亚", "植物系列占比 2.5%（对标 4.1%）", "花束类上收银台第二陈列位，由代班负责人跟进落地并拍照", D0920),
    ("淮安新亚", "爆款集中：法拉利 488 占 7.7%", "周三巡店核对爆款补货深度防周末断码，库存结论同步城市经理", D0916),
    ("淮安万象城", "完成率 57.2% 低于区域均值 23 个百分点", "与商场确认本周及 9 月下旬档期资源，周六日排班加配，周完成率目标 ≥75%", D0916),
    ("淮安万象城", "客单价 ¥330 低区域 ¥60、连带率 1.23 区域最低、植物系列 2.8%", "收银台加购话术落地+花束上第二陈列位，客单价目标 ≥¥370", D0920),
    ("淮安万象城", "黄锦靓周销 ¥4,293 仅店人均 34%", "于思涵带教连带与高价款推荐，周五检视，周销目标 ≥¥8,000", D0918),
]

DATA_START = 4
for i, (store, issue, plan, dl) in enumerate(actions):
    r = DATA_START + i
    ws[f"A{r}"] = i + 1
    ws[f"B{r}"] = store
    ws[f"C{r}"] = issue
    ws[f"D{r}"] = plan
    ws[f"E{r}"] = dl
    ws[f"F{r}"] = "未开始"
    ws[f"G{r}"] = ""
    # 样式
    ws[f"A{r}"].alignment = align_c
    ws[f"B{r}"].alignment = align_c
    ws[f"C{r}"].alignment = align_l
    ws[f"D{r}"].alignment = align_l
    ws[f"E{r}"].alignment = align_c
    ws[f"E{r}"].number_format = "M月D日 ddd"
    ws[f"E{r}"].number_format = "YYYY-MM-DD"
    ws[f"F{r}"].alignment = align_c
    ws[f"G{r}"].alignment = align_l
    for col in "ABCDEFG":
        ws[f"{col}{r}"].font = font_body
        ws[f"{col}{r}"].border = thin_border
    ws.row_dimensions[r].height = 34

DATA_END = DATA_START + len(actions) - 1  # 24

# 列宽
widths = {"A": 6, "B": 14, "C": 38, "D": 48, "E": 12, "F": 10, "G": 16}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

# 状态列数据验证
dv = DataValidation(type="list", formula1='"未开始,进行中,已完成,逾期"', allow_blank=True)
dv.error = "请选择：未开始/进行中/已完成/逾期"
ws.add_data_validation(dv)
dv.add(f"F{DATA_START}:F{DATA_END}")

# 状态列条件格式
ws.conditional_formatting.add(
    f"F{DATA_START}:F{DATA_END}",
    CellIsRule(operator="equal", formula=['"已完成"'], fill=PatternFill("solid", fgColor=XL_GREEN_BG), font=Font(color=XL_GREEN_FT)))
ws.conditional_formatting.add(
    f"F{DATA_START}:F{DATA_END}",
    CellIsRule(operator="equal", formula=['"进行中"'], fill=PatternFill("solid", fgColor=XL_YELLOW_BG), font=Font(color=XL_YELLOW_FT)))
ws.conditional_formatting.add(
    f"F{DATA_START}:F{DATA_END}",
    CellIsRule(operator="equal", formula=['"逾期"'], fill=PatternFill("solid", fgColor=XL_RED_BG), font=Font(color=XL_RED_FT)))
# 截止日期早于今天且未完成 → 红字（静态生成，今天 2026-09-14，本周内不触发，规则留给用户后续使用）

ws.auto_filter.ref = f"A3:G{DATA_END}"
ws.freeze_panes = "A4"

# ============ Sheet2 门店周数据 ============
ws2 = wb.create_sheet("门店周数据（第37周）")
# 锚点：第1行标题、第2行表头、第3-9行数据（7 店）
ws2.merge_cells("A1:K1")
ws2["A1"] = "门店周数据（第 37 周 9/7-9/13，区域：周销 ¥365,109 / 环比 -12.6% / 同比 -36%）"
ws2["A1"].font = font_title
ws2["A1"].fill = fill_title
ws2["A1"].alignment = align_c
ws2.row_dimensions[1].height = 26

h2 = ["门店", "周销（¥）", "周目标（¥）", "完成率", "环比", "同比", "客单价（¥）", "连带率", "长尾款（¥）", "长尾指标完成率", "植物系列占比"]
for col, name in zip("ABCDEFGHIJK", h2):
    c = ws2[f"{col}2"]
    c.value = name
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_c
    c.border = thin_border
ws2.row_dimensions[2].height = 30

# (门店, 周销, 周目标, 环比, 同比, 客单价, 连带率, 长尾款, 长尾完成率, 植物占比)
store_data = [
    ("扬州万象汇", 62105.75, 75600, -0.163, -0.461, 453.3, 1.30, 9998.08, 3.306, 0.041),
    ("扬州京华城", 88650.50, 105840, -0.213, -0.258, 447.7, 1.46, 9026.00, 2.132, 0.064),
    ("扬州江都金鹰", 29359.00, 41580, -0.250, -0.426, 345.4, 1.24, 3506.00, 2.108, 0.095),
    ("泰州万象城", 76509.00, 105840, -0.220, -0.307, 409.1, 1.35, 3281.00, 0.775, 0.035),
    ("宿迁宝龙", 12156.00, 24570, -0.497, -0.570, 270.1, 1.46, 207.00, 0.211, 0.069),
    ("淮安新亚", 44479.00, 30240, 0.912, -0.313, 478.3, 1.57, 3182.00, 2.631, 0.025),
    ("淮安万象城", 51850.00, 90720, 0.125, None, 330.3, 1.23, 5410.00, 1.491, 0.028),
]

S_START = 3
for i, row in enumerate(store_data):
    r = S_START + i
    name, amt, tgt, wow, yoy, atv, jd, lt, lt_rate, bot = row
    ws2[f"A{r}"] = name
    ws2[f"B{r}"] = amt
    ws2[f"C{r}"] = tgt
    # 完成率=周销/周目标（周快照静态值，改数后需手动重算）
    ws2[f"D{r}"] = round(amt / tgt, 4) if (amt and tgt) else ""
    ws2[f"E{r}"] = wow
    ws2[f"F{r}"] = yoy if yoy is not None else "新店无同比"
    ws2[f"G{r}"] = atv
    ws2[f"H{r}"] = jd
    ws2[f"I{r}"] = lt
    ws2[f"J{r}"] = lt_rate
    ws2[f"K{r}"] = bot
    # 样式与格式
    ws2[f"A{r}"].alignment = align_c
    for col in "BCG":
        ws2[f"{col}{r}"].alignment = align_r
        ws2[f"{col}{r}"].number_format = "#,##0"
    for col in "EFJK":
        ws2[f"{col}{r}"].alignment = align_r
        ws2[f"{col}{r}"].number_format = "0.0%"
    ws2[f"I{r}"].alignment = align_r
    ws2[f"I{r}"].number_format = "#,##0"
    ws2[f"H{r}"].alignment = align_r
    ws2[f"H{r}"].number_format = "0.00"
    ws2[f"F{r}"].alignment = align_r
    for col in "ABCDEFGHIJK":
        ws2[f"{col}{r}"].font = font_body
        ws2[f"{col}{r}"].border = thin_border

S_END = S_START + len(store_data) - 1  # 9

for col, w in {"A": 14, "B": 12, "C": 12, "D": 10, "E": 9, "F": 11, "G": 11, "H": 9, "I": 11, "J": 13, "K": 12}.items():
    ws2.column_dimensions[col].width = w

# 完成率条件格式：<60% 红、60-80% 黄、≥80% 绿
ws2.conditional_formatting.add(
    f"D{S_START}:D{S_END}",
    CellIsRule(operator="lessThan", formula=["0.6"], fill=PatternFill("solid", fgColor=XL_RED_BG), font=Font(color=XL_RED_FT)))
ws2.conditional_formatting.add(
    f"D{S_START}:D{S_END}",
    CellIsRule(operator="between", formula=["0.6", "0.799"], fill=PatternFill("solid", fgColor=XL_YELLOW_BG), font=Font(color=XL_YELLOW_FT)))
ws2.conditional_formatting.add(
    f"D{S_START}:D{S_END}",
    CellIsRule(operator="greaterThanOrEqual", formula=["0.8"], fill=PatternFill("solid", fgColor=XL_GREEN_BG), font=Font(color=XL_GREEN_FT)))
# 环比/同比负值红字
for rng in (f"E{S_START}:E{S_END}", f"F{S_START}:F{S_END}"):
    ws2.conditional_formatting.add(
        rng, CellIsRule(operator="lessThan", formula=["0"], font=Font(color=XL_RED_FT)))

ws2.freeze_panes = "B3"

wb.properties.title = "苏中区域门店本周重点工作与行动计划"
wb.save("/Users/a123/WorkBuddy/Claw/outputs/followup_weekly/苏中区域门店本周重点工作与行动计划_20260914.xlsx")
print("saved")
