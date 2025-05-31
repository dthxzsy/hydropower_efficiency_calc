import os
import pandas as pd
import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy

# Configuration Constants
MAIN_PATH = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\Source_Data"
REMOVE_SHEETS = [
    "61603780",
    "61603790",
    "61603820",
    "61605870",
    "61606640",
    "61608350",
    "61608360",
    "61627850",
]
FLOW_STATION = [
    "61608000",
    "61604100",
    "61604300",
    "61604400",
    "61608200",
    "61608220",
    "61608310",
    "61608330",
]
FLOW_SOURCE = (
    r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\temp\水位流量关系.xlsx"
)
TEMPLATE_FILES = (
    r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\temp\Template Files.xls"
)
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\Water_Level"
START_TIME = "2025-05-29 09:00:00"
END_TIME = "2025-05-30 09:00:00"

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Load flow relationships once
df_dict = None


def get_input_files(main_path):
    """Discover all relevant .xls files in the main path."""
    return [
        os.path.join(main_path, file)
        for file in os.listdir(main_path)
        if file.endswith(".xls") and "河道水情表" in file
    ]


def filter_sheets(file_path, remove_sheets):
    """Filter out unwanted sheets from the Excel file."""
    xls = pd.ExcelFile(file_path)
    return [sheet for sheet in xls.sheet_names if sheet not in remove_sheets]


def process_sheet_data(file_path, sheet_name):
    """Process data from a specific sheet."""
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = df.iloc[0]  # Set header from first row
    df = df.drop(0).reset_index(drop=True)  # Remove header row
    station_name = "_".join(
        df["测站名称"].unique().astype(str)
    )  # Handle NaN or mixed types

    columns = [
        "站码",
        "时间",
        "水位",
        "流量",
        "断面过水面积",
        "断面平均流速",
        "断面最大流速",
        "河水特征码",
        "水势",
        "测流方法",
        "测积方法",
        "测速方法",
    ]
    available_columns = [col for col in columns if col in df.columns]
    df = df[available_columns]
    df["时间"] = pd.to_datetime(df["时间"], errors="coerce")  # Handle invalid dates

    return df, station_name


def load_flow_relationships(flow_source):
    """Load flow rate relationships from the provided Excel file."""
    df_dict = {}
    xls = pd.ExcelFile(flow_source)
    for sheet in xls.sheet_names:
        df_dict[sheet] = pd.read_excel(flow_source, sheet_name=sheet)
    return df_dict


def merge_with_template(df, start, end):
    """Create a time template and merge with the data to identify missing entries."""
    template_time = pd.DataFrame(
        {"时间": pd.date_range(start=start, end=end, freq="h")}
    )
    merged_df = template_time.merge(df, on="时间", how="outer")
    merged_df["数据缺失"] = merged_df["水位"].isna()
    merged_df_before = merged_df.copy()  # 保存合并前的状态
    return merged_df_before, merged_df


def interpolate_and_clean(merged_df):
    """Interpolate missing values and clean the data."""
    if "水位" in merged_df.columns:
        merged_df["水位"] = pd.to_numeric(merged_df["水位"], errors="coerce")
        merged_df["水位"] = (
            merged_df["水位"].interpolate(method="linear").ffill().bfill().round(2)
        )
    if "站码" in merged_df.columns:
        merged_df["站码"] = merged_df["站码"].ffill().bfill()
    if "测速方法" in merged_df.columns:
        merged_df["测速方法"] = (
            merged_df["测速方法"].astype(str).replace("nan", np.nan).ffill().bfill()
        )
    return merged_df


def calculate_water_level_change(merged_df):
    """Calculate water level change and update water momentum."""
    if "水位" not in merged_df.columns:
        merged_df["水势"] = "6"  # Default if no water level
        return merged_df

    def get_change(row):
        if row.name > 0 and not pd.isna(row["水位"]):
            prev_value = merged_df["水位"].iloc[row.name - 1]
            curr_value = row["水位"]
            if pd.isna(prev_value):
                return "6"
            diff = curr_value - prev_value
            return "4" if diff < 0 else "5" if diff > 0 else "6"
        return "6"

    merged_df["水势"] = merged_df.apply(get_change, axis=1)
    return merged_df


def get_flow_interpolated_np(station_code, water_level):
    """Get interpolated flow value using numpy."""
    if station_code in df_dict and "水位" in df_dict[station_code].columns:
        station_df = df_dict[station_code]
        print(f"Before: station_df['流量'] = {station_df['流量'].tolist()}")

        if station_code in FLOW_STATION:
            if water_level in station_df["水位"].values:
                return station_df[station_df["水位"] == water_level]["流量"].values[0]
            return np.interp(water_level, station_df["水位"], station_df["流量"])
        else:
            return ""  # Return empty string for non-FLOW_STATION stations
    return ""  # Return empty string if conditions not met


def write_to_excel(result, template_path, output_dir, station_name):
    """Append the processed data to an existing Excel template file."""
    # Open the template workbook
    rb = xlrd.open_workbook(template_path, formatting_info=True)
    wb = copy(rb)
    sheet_wr = wb.get_sheet(0)

    # Find the last row with data in the template
    sheet_rd = rb.sheet_by_index(0)
    last_row = sheet_rd.nrows  # Number of rows already in the template

    # Define styles
    date_style = xlwt.XFStyle()
    date_style.num_format_str = "yyyy-mm-dd hh:mm:ss"
    num_style = xlwt.XFStyle()
    num_style.num_format_str = "0.00"

    # Headers (only write if the template is empty)
    headers = [
        "站码",
        "时间",
        "水位",
        "流量",
        "断面过水面积",
        "断面平均流速",
        "断面最大流速",
        "河水特征码",
        "水势",
        "测流方法",
        "测积方法",
        "测速方法",
    ]

    if last_row == 0:  # If template is empty, write headers
        for col_idx, header in enumerate(headers):
            sheet_wr.write(0, col_idx, header)
        start_row = 1
    else:
        start_row = last_row  # Append after existing data

    # Write the data from result
    for row_idx, (_, row) in enumerate(result.iterrows(), start=start_row):
        for col_idx, header in enumerate(headers):
            value = row.get(header, "")  # Use .get() to avoid KeyError
            if pd.isna(value) or value == "":
                sheet_wr.write(row_idx, col_idx, "")  # Write blank for NaN or empty
            elif header == "时间" and isinstance(value, pd.Timestamp):
                sheet_wr.write(row_idx, col_idx, value, date_style)
            elif header in ["水位", "流量"] and isinstance(value, (int, float)):
                sheet_wr.write(row_idx, col_idx, float(value), num_style)
            else:
                sheet_wr.write(row_idx, col_idx, value)

    # Save the output
    output_path = os.path.join(output_dir, f"{station_name}.xls")
    wb.save(output_path)
    return output_path


def main():
    global df_dict
    df_dict = load_flow_relationships(FLOW_SOURCE)
    full_paths = get_input_files(MAIN_PATH)

    final_columns = [
        "站码",
        "时间",
        "水位",
        "流量",
        "断面过水面积",
        "断面平均流速",
        "断面最大流速",
        "河水特征码",
        "水势",
        "测流方法",
        "测积方法",
        "测速方法",
    ]

    for file_path in full_paths:
        new_sheet_names = filter_sheets(file_path, REMOVE_SHEETS)
        for sheet_name in new_sheet_names:
            try:
                df, station_name = process_sheet_data(file_path, sheet_name)
                merged_df_before, merged_df = merge_with_template(
                    df, START_TIME, END_TIME
                )

                # 日志记录缺失数据
                missing_count = merged_df["数据缺失"].sum()
                print(
                    f"Sheet {sheet_name}: {missing_count} missing data points detected."
                )

                # 数据处理：水位填充、水势计算、流量填充
                merged_df = interpolate_and_clean(merged_df)  # 水位填充
                merged_df = calculate_water_level_change(merged_df)  # 水势计算

                if "站码" in merged_df.columns and "水位" in merged_df.columns:
                    merged_df["流量"] = merged_df.apply(
                        lambda row: (
                            round(get_flow_interpolated_np(row["站码"], row["水位"]), 3)
                            if get_flow_interpolated_np(row["站码"], row["水位"]) != ""
                            else ""
                        ),
                        axis=1,
                    )

                # 所有处理完成后，创建 merged_df_after
                merged_df_after = merged_df.copy()

                # 调试：检查 merged_df_before 和 merged_df_after 的索引和数据
                print(
                    f"merged_df_before shape: {merged_df_before.shape}, index: {merged_df_before.index[:5]}"
                )
                print(
                    f"merged_df_after shape: {merged_df_after.shape}, index: {merged_df_after.index[:5]}"
                )

                result = merged_df_before[
                    merged_df_before["水位"].isna() | (merged_df_before["水位"] == "")
                ]
                print(f"Result shape: {result.shape}, index: {result.index[:5]}")

                # 如果 result 为空，跳过后续处理
                if result.empty:
                    print(
                        f"No missing water level data found for {sheet_name}, skipping..."
                    )
                    continue

                # 获取 merged_df_after 中对应索引的行
                filtered_after = merged_df_after.loc[result.index]
                print(f"filtered_after shape after loc: {filtered_after.shape}")

                # 如果 filtered_after 为空，尝试基于时间列匹配
                if filtered_after.empty:
                    print(
                        f"Index matching failed, trying time-based matching for {sheet_name}"
                    )
                    filtered_after = merged_df_after[
                        merged_df_after["时间"].isin(result["时间"])
                    ]
                    print(
                        f"filtered_after shape after time match: {filtered_after.shape}"
                    )

                # 如果仍然为空，记录问题并跳过
                if filtered_after.empty:
                    print(
                        f"Failed to find matching data in merged_df_after for {sheet_name}, skipping..."
                    )
                    continue

                # 确保非 FLOW_STATION 站点的 "流量" 为空白
                if "站码" in filtered_after.columns:
                    station_code = filtered_after["站码"].iloc[0]  # 假设每个表单一站码
                    if station_code not in FLOW_STATION:
                        filtered_after["流量"] = ""

                # 选择指定列
                available_columns = [
                    col for col in final_columns if col in filtered_after.columns
                ]
                if not available_columns:
                    print(
                        f"No columns from final_columns found in filtered_after for {sheet_name}, skipping..."
                    )
                    continue
                filtered_after = filtered_after[available_columns]

                # 使用 write_to_excel 保存筛选后的结果（唯一输出）
                filtered_output_file = write_to_excel(
                    filtered_after,
                    TEMPLATE_FILES,
                    OUTPUT_DIR,
                    f"filtered_after_{station_name}",
                )
                print(f"筛选后的数据已保存到: {filtered_output_file}")

            except Exception as e:
                print(f"Error processing {file_path} - {sheet_name}: {str(e)}")


if __name__ == "__main__":
    main()
