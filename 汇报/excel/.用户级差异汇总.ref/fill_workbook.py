# -*- coding: utf-8 -*-
"""按 schema.md 图纸填充 用户级差异汇总.xlsx（sheet-agent 不可用时的降级执行）。"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

PATH = r"D:\code\yunyun3\word\汇报\excel\用户级差异汇总.xlsx"

# §6.2 数据：用户, 积分消耗, 积分金额(元), 厂商扣费(元)（差异/差异率为计算列）
rows = [
    ("上班史不得", 207690, 3461.50, 3360.50),
    ("杜崇巍", 120390, 2006.50, 2002.29),
    ("韩昊志", 108930, 1815.50, 1822.68),
    ("小马", 103860, 1731.00, 1710.67),
    ("XiaoH", 2160, 36.00, 862.22),
    ("崔什么", 9900, 165.00, 164.45),
    ("裴", 6960, 116.00, 153.62),
    ("祈悦今朝", 4500, 75.00, 74.75),
    ("森澜兴禾", 1740, 29.00, 28.46),
    ("wx_fpls4pivbmyo", 240, 4.00, 14.95),
    ("wx_fsowbnp6xhc0", 840, 14.00, 13.00),
    ("斯温", 240, 4.00, 4.02),
]

wb = openpyxl.load_workbook(PATH)
ws = wb["用户级差异汇总"]

header_fill = PatternFill("solid", fgColor="4472C4")
total_fill = PatternFill("solid", fgColor="2F5597")
white_bold = Font(bold=True, color="FFFFFF")
red_font = Font(color="9C0006")
thin_bottom = Border(bottom=Side(style="thin"))
thin_top = Border(top=Side(style="thin"))

headers = ["用户", "积分消耗", "积分金额(元)", "厂商扣费(元)", "差异(元)", "差异率"]
for c, h in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=c, value=h)
    cell.fill = header_fill
    cell.font = white_bold
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_bottom

last = 1 + len(rows)  # 数据末行 13
for i, (name, points, amount, vendor) in enumerate(rows, start=2):
    ws.cell(row=i, column=1, value=name).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=i, column=2, value=points).number_format = "#,##0"
    ws.cell(row=i, column=3, value=amount).number_format = "#,##0.00"
    ws.cell(row=i, column=4, value=vendor).number_format = "#,##0.00"
    ws.cell(row=i, column=5, value=f"=C{i}-D{i}").number_format = "#,##0.00"
    ws.cell(row=i, column=6, value=f'=IF(OR(D{i}="",D{i}=0),"",E{i}/D{i})').number_format = "0.00%"
    for c in range(2, 7):
        ws.cell(row=i, column=c).alignment = Alignment(horizontal="right", vertical="center")

tr = last + 1  # 合计行 14
ws.cell(row=tr, column=1, value="合计")
ws.cell(row=tr, column=2, value=f"=SUM(B2:B{last})").number_format = "#,##0"
ws.cell(row=tr, column=3, value=f"=SUM(C2:C{last})").number_format = "#,##0.00"
ws.cell(row=tr, column=4, value=f"=SUM(D2:D{last})").number_format = "#,##0.00"
ws.cell(row=tr, column=5, value=f"=SUM(E2:E{last})").number_format = "#,##0.00"
ws.cell(row=tr, column=6, value=f'=IF(OR(D{tr}="",D{tr}=0),"",E{tr}/D{tr})').number_format = "0.00%"
for c in range(1, 7):
    cell = ws.cell(row=tr, column=c)
    cell.fill = total_fill
    cell.font = white_bold
    cell.border = thin_top
    cell.alignment = Alignment(horizontal="center" if c == 1 else "right", vertical="center")

# 条件格式：E2:F14 数值 < 0 → 红字
for col in ("E", "F"):
    rng = f"{col}2:{col}{tr}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=red_font))

# 列宽（A 列含长用户名 wx_ 前缀）
widths = {"A": 20, "B": 13, "C": 15, "D": 15, "E": 13, "F": 11}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

wb.save(PATH)
print("saved:", PATH)
