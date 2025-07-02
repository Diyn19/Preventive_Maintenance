import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, numbers
from openpyxl.utils import get_column_letter
from datetime import datetime
import copy

# 自動抓今天的年月
month_tag = datetime.today().strftime("%Y%m")
report_file = f"IM/{month_tag}_HL_Maintain_Report.xlsx"
data_file = "data.xlsx"

# 開啟 data.xlsx 並讀取 IM 分頁
data_wb = load_workbook(data_file)
data_ws = data_wb["IM"]

# 找出台芝案號欄 (直接抓 C 欄案號，從第 2 列開始往下)
case_col_index = 3  # C = 第3欄
last_case_id = None
for row in reversed(list(data_ws.iter_rows(min_row=2, values_only=True))):
    if row[case_col_index - 1]:
        last_case_id = str(row[case_col_index - 1]).strip()
        break

print("最後案號：", last_case_id)

# 開啟報表
report_wb = load_workbook(report_file, data_only=True)
report_ws = report_wb.active

# 從 C2 開始比對案號，收集新資料
start_row = 2
append_rows = []
start_append = False
for row in report_ws.iter_rows(min_row=start_row):
    case_cell = row[2]  # C 欄
    case_id = str(case_cell.value).strip() if case_cell.value else ""
    if last_case_id and case_id == last_case_id:
        start_append = True
        continue
    if start_append:
        values = [cell.value for cell in row]
        if all(v is None for v in values):
            break
        append_rows.append(values)

print(f"共 {len(append_rows)} 列新資料將加入 IM")

# 參考樣式與公式
ref_row = data_ws.max_row
ref_row_height = 21.66
ref_cells = {cell.column: cell for cell in data_ws[ref_row]}
al_formula_cell = data_ws[f"AL{ref_row}"]
am_formula_cell = data_ws[f"AM{ref_row}"]

# 寫入資料
for row_data in append_rows:
    data_ws.append(row_data)
    new_row = data_ws.max_row
    data_ws.row_dimensions[new_row].height = ref_row_height

    for col_idx, value in enumerate(row_data, start=1):
        cell = data_ws.cell(row=new_row, column=col_idx)
        ref_cell = ref_cells.get(col_idx)
        if ref_cell:
            cell.font = copy.copy(ref_cell.font)
            cell.alignment = copy.copy(ref_cell.alignment)
            cell.border = copy.copy(ref_cell.border)
            cell.fill = copy.copy(ref_cell.fill)

        # 強制日期欄位轉為標準格式（這裡假設第 24 欄是日期欄）
        if col_idx == 24 and value:
            try:
                if isinstance(value, str):
                    dt = datetime.strptime(value.strip(), "%Y-%m-%d %H:%M:%S")
                elif isinstance(value, datetime):
                    dt = value
                else:
                    dt = None

                if dt:
                    cell.value = dt
                    cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2  # yyyy/mm/dd
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            except Exception as e:
                print(f"❗ 日期格式轉換錯誤（欄{col_idx}）：{e}")

    # 複製公式
    for col_letter, formula_cell in [("AL", al_formula_cell), ("AM", am_formula_cell)]:
        target_cell = data_ws[f"{col_letter}{new_row}"]
        formula = formula_cell.value.replace(str(ref_row), str(new_row)) if formula_cell.data_type == "f" else formula_cell.value
        target_cell.value = formula
        target_cell.font = copy.copy(formula_cell.font)
        target_cell.alignment = copy.copy(formula_cell.alignment)
        target_cell.border = copy.copy(formula_cell.border)
        target_cell.number_format = formula_cell.number_format

# 儲存檔案
data_wb.save(data_file)
print("✅ 資料更新完成")
