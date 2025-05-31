import os
import pandas as pd
import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy
from flow_data import water_flow_dict  

# 配置常量
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
TEMPLATE_FILES = (
    r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\temp\Template Files.xls"
)
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\Water_Level"

os.makedirs(OUTPUT_DIR, exist_ok=True)


def get_input_files(main_path):
    return [
        os.path.join(main_path, file)
        for file in os.listdir(main_path)
        if file.endswith(".xls") and "河道水情表" in file
    ]


def filter_sheets(file_path, remove_sheets):
    xls = pd.ExcelFile(file_path)
    return [sheet for sheet in xls.sheet_names if sheet not in remove_sheets]


def process_sheet_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = df.iloc[0]
    df = df.drop(0).reset_index(drop=True)
    df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
    station_name = "_".join(df["测站名称"].unique().astype(str))
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
    df = df[[col for col in columns if col in df.columns]]
    return df, station_name


def merge_with_template(df, start, end):
    template_time = pd.DataFrame(
        {"时间": pd.date_range(start=start, end=end, freq="h")}
    )
    merged_df = template_time.merge(df, on="时间", how="outer")
    merged_df["数据缺失"] = merged_df["水位"].isna()
    return merged_df.copy(), merged_df


def interpolate_and_clean(merged_df):
    if "水位" in merged_df.columns:
        merged_df["水位"] = pd.to_numeric(merged_df["水位"], errors="coerce")
        merged_df["水位"] = merged_df["水位"].interpolate().ffill().bfill().round(2)
    if "站码" in merged_df.columns:
        merged_df["站码"] = merged_df["站码"].ffill().bfill()
    if "测速方法" in merged_df.columns:
        merged_df["测速方法"] = (
            merged_df["测速方法"].astype(str).replace("nan", np.nan).ffill().bfill()
        )
    return merged_df


def calculate_water_level_change(merged_df):
    if "水位" not in merged_df.columns:
        merged_df["水势"] = "6"
        return merged_df

    def get_change(row):
        if row.name > 0 and not pd.isna(row["水位"]):
            prev = merged_df["水位"].iloc[row.name - 1]
            curr = row["水位"]
            if pd.isna(prev):
                return "6"
            return "4" if curr < prev else "5" if curr > prev else "6"
        return "6"

    merged_df["水势"] = merged_df.apply(get_change, axis=1)
    return merged_df


def get_flow_interpolated_np(station_code, water_level):
    if station_code in water_flow_dict:
        records = water_flow_dict[station_code]
        levels = [r["水位"] for r in records]
        flows = [r["流量"] for r in records]
        if station_code in FLOW_STATION:
            if water_level in levels:
                return next(r["流量"] for r in records if r["水位"] == water_level)
            return np.interp(water_level, levels, flows)
    return ""


def write_to_excel(result, template_path, output_dir, station_name):
    rb = xlrd.open_workbook(template_path, formatting_info=True)
    wb = copy(rb)
    sheet_wr = wb.get_sheet(0)
    sheet_rd = rb.sheet_by_index(0)
    last_row = sheet_rd.nrows

    date_style = xlwt.XFStyle()
    date_style.num_format_str = "yyyy-mm-dd hh:mm:ss"
    num_style = xlwt.XFStyle()
    num_style.num_format_str = "0.00"

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

    start_row = 0 if last_row == 0 else last_row
    if last_row == 0:
        for col_idx, header in enumerate(headers):
            sheet_wr.write(0, col_idx, header)
        start_row = 1

    for row_idx, (_, row) in enumerate(result.iterrows(), start=start_row):
        for col_idx, header in enumerate(headers):
            value = row.get(header, "")
            if pd.isna(value) or value == "":
                sheet_wr.write(row_idx, col_idx, "")
            elif header == "时间" and isinstance(value, pd.Timestamp):
                sheet_wr.write(row_idx, col_idx, value, date_style)
            elif header in ["水位", "流量"] and isinstance(value, (int, float)):
                sheet_wr.write(row_idx, col_idx, float(value), num_style)
            else:
                sheet_wr.write(row_idx, col_idx, value)

    output_path = os.path.join(output_dir, f"{station_name}.xls")
    wb.save(output_path)
    return output_path


def main():
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

                # 自动获取 START_TIME 和 END_TIME
                start_time = (
                    df["时间"].iloc[2] if len(df) > 2 else df["时间"].dropna().min()
                )
                end_time = df["时间"].dropna().max()

                merged_df_before, merged_df = merge_with_template(
                    df, start_time, end_time
                )

                print(f"Sheet {sheet_name}: 缺失点数 = {merged_df['数据缺失'].sum()}")

                merged_df = interpolate_and_clean(merged_df)
                merged_df = calculate_water_level_change(merged_df)

                if "站码" in merged_df.columns and "水位" in merged_df.columns:
                    merged_df["流量"] = merged_df.apply(
                        lambda row: (
                            round(get_flow_interpolated_np(row["站码"], row["水位"]), 3)
                            if get_flow_interpolated_np(row["站码"], row["水位"]) != ""
                            else ""
                        ),
                        axis=1,
                    )

                merged_df_after = merged_df.copy()
                result = merged_df_before[
                    merged_df_before["水位"].isna() | (merged_df_before["水位"] == "")
                ]

                if result.empty:
                    print(f"✔ 无缺失水位，跳过 {sheet_name}")
                    continue

                filtered_after = merged_df_after.loc[result.index]
                if filtered_after.empty:
                    filtered_after = merged_df_after[
                        merged_df_after["时间"].isin(result["时间"])
                    ]
                if filtered_after.empty:
                    print(f"✘ 匹配失败：{sheet_name}")
                    continue

                station_code = filtered_after["站码"].iloc[0]
                if station_code not in FLOW_STATION:
                    filtered_after["流量"] = ""

                available_columns = [
                    col for col in final_columns if col in filtered_after.columns
                ]
                if not available_columns:
                    print(f"✘ 无匹配列：{sheet_name}")
                    continue

                filtered_after = filtered_after[available_columns]

                filtered_output_file = write_to_excel(
                    filtered_after,
                    TEMPLATE_FILES,
                    OUTPUT_DIR,
                    f"filtered_after_{station_name}",
                )
                print(f" 结果已保存到: {filtered_output_file}")

            except Exception as e:
                print(f"处理错误：{file_path} - {sheet_name}: {str(e)}")


if __name__ == "__main__":
    main()
