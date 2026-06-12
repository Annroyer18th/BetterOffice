import pandas as pd
from openpyxl import load_workbook
import os

# ============= excel以某一列为关键字将包含相同数据的行剪切到新工作表 ======================
EXCEL_FILE = "2025对账单.xlsx"    # Excel文件名
SHEET_NAME = "对账单数据"          # 原数据所在工作表
KEY_COLUMN = "对账时间"             # 关键字列名（比如 部门、姓名、分类、ID）

# 读取原数据
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

# 按关键字分组
groups = df.groupby(KEY_COLUMN)

# 保存每个分组到新工作表，并从原表删除
with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    # 遍历每个分组
    for key, group_data in groups:
        # 写入新工作表
        group_data.to_excel(writer, sheet_name=str(key), index=False)
        
        # 从原表删除这些行
        df = df.drop(group_data.index)

# 覆盖原工作表（剩下的数据）
with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

print(f"✅ 处理完成！")
print(f"📄 原表 {SHEET_NAME} 已删除被移动的行")
print(f"📊 已自动创建 {len(groups)} 个新工作表")