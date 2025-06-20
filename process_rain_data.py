import os
import pandas as pd
import numpy as np
import xlrd
from xlutils.copy import copy
import xlwt
# === 配置路径 ===
main_path = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace_\Source_Data"
template_path = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace_\temp\rain_model.xls"
output_dir = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace_\Rainfall\chck"

# 自动查找降水量表文件
xls_path = None
for file in os.listdir(main_path):
    if "降水量表" in file and file.endswith(".xls"):
        xls_path = os.path.join(main_path, file)
        break
if not xls_path:
    raise FileNotFoundError("❌ 未找到包含 '降水量表' 的 .xls 文件")

xls = pd.ExcelFile(xls_path)

# === 工具函数 ===
def round_to_half(x):
    return round(float(x) * 2) / 2 if pd.notna(x) else np.nan

def get_rain_date(t):
    if pd.isna(t): return np.nan
    return (t - pd.Timedelta(days=1)).date() if t.hour < 9 else t.date()



def write_to_template(template_path, output_path, data_df):
    data_df = data_df.iloc[2:].reset_index(drop=True)
    book = xlrd.open_workbook(template_path, formatting_info=True)
    new_book = copy(book)
    sheet = new_book.get_sheet(0)
    start_row = 2

    # 设置时间格式样式
    datetime_style = xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm')

    for i, (_, row) in enumerate(data_df.iterrows()):
        for j, value in enumerate(row):
            if j == 1 and pd.notna(value):  # 时间列索引 = 1（即“时间”列）
                sheet.write(start_row + i, j, value, datetime_style)
            else:
                sheet.write(start_row + i, j, value)
    new_book.save(output_path)


# === 站点组合（用于补值） ===
station_groups = [
    ["61604000", "61604400", "61623420"],
    ["61604200", "61623420", "61604400"],
    ["61603800", "61623420"],
    ["61629900", "61607000", "61630300"],
    ["61607000", "61629900", "61630310"],
    ["61608000", "61628900", "61606600"],
    ["61608210", "61629650", "61608180"],
    ["61622760", "61622900", "61602880"],
    ["61622400", "61622450"],
    ["61608100", "61608200", "61608225"],
]

# === 读取并清洗每个 sheet（从第二行为列名，第三行起为数据） ===
sheet_data = {}
for sheet in xls.sheet_names:
    df_raw = xls.parse(sheet, header=None)
    if df_raw.shape[0] < 3:
        continue
    df_raw.columns = df_raw.iloc[1]
    df = df_raw.iloc[2:].copy()
    if "时间" in df.columns and "时段降水量" in df.columns:
        df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
        df["时段降水量"] = pd.to_numeric(df["时段降水量"], errors="coerce")
        sheet_data[sheet] = df[["时间", "时段降水量"]]

# === 获取全局时间范围 ===
all_times = pd.concat([df["时间"] for df in sheet_data.values()])
start_time = all_times.min()
end_time = all_times.max()
full_time_index = pd.date_range(start=start_time, end=end_time, freq="1H")

# === 对每个站补齐时间序列 ===
for code in sheet_data:
    df = sheet_data[code].set_index("时间").reindex(full_time_index).reset_index()
    df.rename(columns={"index": "时间"}, inplace=True)
    sheet_data[code] = df

# === 主逻辑：按 group 填补缺值并写入模板 ===
os.makedirs(output_dir, exist_ok=True)

for group in station_groups:
    df_list = [sheet_data[c].rename(columns={"时段降水量": c}) for c in group if c in sheet_data]
    if not df_list:
        continue

    merged = df_list[0]
    for other in df_list[1:]:
        merged = pd.merge(merged, other, on="时间", how="outer")

    for code in group:
        if code not in merged.columns:
            continue
        for i in merged.index:
            val = merged.at[i, code]
            if pd.isna(val):
                others = [merged.at[i, c] for c in group if c != code and c in merged.columns]
                valid_vals = [v for v in others if pd.notna(v)]
                if len(valid_vals) == 2 and all(v == 0 for v in valid_vals):
                    merged.at[i, code] = 0.0
                elif valid_vals:
                    merged.at[i, code] = round_to_half(np.mean(valid_vals))

    # === 每个站导出 ===
    for code in group:
        if code not in merged.columns:
            continue
        df = merged[["时间", code]].copy()
        df.rename(columns={code: "时段降水量"}, inplace=True)
        df["时段降水量"] = df["时段降水量"].apply(round_to_half)
        df["时段长(h)"] = 1
        df["天气状况"] = df["时段降水量"].apply(lambda x: 7 if x > 0 else 9)
        df["降水日"] = df["时间"].apply(get_rain_date)

        daily = df.groupby("降水日")["时段降水量"].sum().apply(round_to_half).reset_index()
        daily.columns = ["降水日", "日降水量"]
        df = pd.merge(df, daily, on="降水日", how="left")

        df_final = pd.DataFrame({
            "站码": [code] * len(df),
            "时间": df["时间"],
            "时段降水量(mm)": df["时段降水量"],
            "时段长(h)": df["时段长(h)"],
            "降水历时": "",
            "日降水量(mm)": df["日降水量"],
            "天气状况": df["天气状况"]
        })

        output_path = os.path.join(output_dir, f"{code}_处理后.xls")
        write_to_template(template_path, output_path, df_final)

print(f"✅ 所有站点数据处理完成，结果已保存至：\n{output_dir}")
from datetime import datetime

# 日志初始化
log_lines = []
log_lines.append(f"处理时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

station_count = 0

for group in station_groups:
    for code in group:
        output_path = os.path.join(output_dir, f"{code}_处理后.xls")
        if os.path.exists(output_path):
            df = pd.read_excel(output_path)
            row_count = len(df)
            log_lines.append(f"✔ 站点 {code} - 处理数据 {row_count} 行")
            station_count += 1

log_lines.append(f"\n总计处理站点数：{station_count}")
log_lines.append("-" * 40)

# 写入日志文件
log_path = os.path.join(output_dir, "rain_processing_log.txt")
with open(log_path, "a", encoding="utf-8") as f:
    f.write("\n".join(log_lines) + "\n\n")

print(f" 已写入处理日志：{log_path}")
