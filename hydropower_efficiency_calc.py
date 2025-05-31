import os
import xlrd
from xlutils.copy import copy
import pandas as pd
import numpy as np

# ⚙️ 常量定义
g = 9.8  # 重力加速度


power_efficiency_data = {
    "大水电": {
        "功率": [
            (26.1, 0.72),
            (26.2, 0.721),
            (26.3, 0.723),
            (26.4, 0.724),
            (26.5, 0.725),
            (26.6, 0.726),
            (26.7, 0.727),
            (26.8, 0.728),
            (26.9, 0.729),
            (27.0, 0.731),
            (27.1, 0.732),
            (27.2, 0.733),
            (27.3, 0.734),
            (27.4, 0.735),
            (27.5, 0.736),
            (27.6, 0.737),
            (27.7, 0.739),
            (27.8, 0.740),
            (27.9, 0.741),
            (28.0, 0.742),
            (28.1, 0.743),
            (28.2, 0.744),
            (28.3, 0.745),
            (28.4, 0.746),
            (28.5, 0.747),
            (28.6, 0.748),
            (28.7, 0.749),
            (28.8, 0.750),
            (28.9, 0.751),
            (29.0, 0.752),
            (29.1, 0.753),
            (29.2, 0.754),
            (29.3, 0.755),
            (29.4, 0.756),
            (29.5, 0.757),
            (29.6, 0.758),
            (29.7, 0.759),
            (29.8, 0.76),
            (29.9, 0.761),
            (30.0, 0.762),
            (30.1, 0.763),
            (30.2, 0.764),
            (30.3, 0.764),
            (30.4, 0.765),
            (30.5, 0.766),
            (30.6, 0.767),
            (30.8, 0.769),
            (30.9, 0.769),
            (31.0, 0.770),
            (31.1, 0.770),
            (31.2, 0.770),
            (31.3, 0.771),
            (31.4, 0.771),
            (31.5, 0.772),
            (31.6, 0.772),
            (31.7, 0.773),
            (31.8, 0.773),
            (31.9, 0.774),
            (32.0, 0.774),
            (32.1, 0.774),
            (33.1, 0.774),
            (34.1, 0.774),
            (35.1, 0.774),
            (36.1, 0.774),
            (37.1, 0.774),
            (38.1, 0.774),
            (39.1, 0.774),
            (39.3, 0.774),
        ],
        "说明": "水位(h)、效率(η)",
    },
    "小水电": {
        "功率": [
            (23.3, 0.4454),
            (23.4, 0.4505),
            (23.5, 0.4556),
            (23.6, 0.4606),
            (23.7, 0.4656),
            (23.8, 0.4706),
            (23.9, 0.4755),
            (24.0, 0.4804),
            (24.1, 0.4853),
            (24.2, 0.4901),
            (24.3, 0.4949),
            (24.4, 0.4997),
            (24.5, 0.5044),
            (24.6, 0.5091),
            (24.7, 0.5138),
            (24.8, 0.5184),
            (24.9, 0.523),
            (25.0, 0.5276),
            (25.1, 0.5321),
            (25.2, 0.5366),
            (25.3, 0.541),
            (25.4, 0.5454),
            (25.5, 0.5497),
            (25.6, 0.554),
            (25.7, 0.5582),
            (25.8, 0.5624),
            (25.9, 0.5665),
            (26.0, 0.5706),
            (26.1, 0.5746),
            (26.2, 0.5786),
            (26.3, 0.5826),
            (26.4, 0.5864),
            (26.5, 0.5902),
            (26.6, 0.5939),
            (26.7, 0.5974),
            (26.8, 0.6009),
            (26.9, 0.6042),
            (27.0, 0.6075),
            (27.1, 0.6107),
            (27.2, 0.6136),
            (27.3, 0.6164),
            (27.4, 0.619),
            (27.5, 0.6214),
            (27.6, 0.6236),
            (27.7, 0.6256),
            (27.8, 0.6275),
            (27.9, 0.6293),
            (28.0, 0.631),
            (28.1, 0.6326),
            (28.2, 0.6341),
            (28.3, 0.6355),
            (28.4, 0.6368),
            (28.5, 0.6381),
            (28.6, 0.6393),
            (28.7, 0.6405),
            (28.8, 0.6416),
            (28.9, 0.6427),
            (29.0, 0.6437),
            (29.2, 0.6454),
            (29.4, 0.647),
            (29.6, 0.6485),
            (29.8, 0.6499),
            (30.0, 0.6513),
            (30.2, 0.6526),
            (30.4, 0.6539),
            (30.6, 0.6552),
            (30.8, 0.6565),
            (31.0, 0.6578),
            (31.2, 0.6591),
            (31.4, 0.6604),
        ],
        "说明": "水位(h)、效率(η)",
    },
}


# 参数对应水头基值
base_values = {1: 19.9, 2: 22.63}


# 路径检查函数
def check_file_exists(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"❌ 文件不存在: {path}")


# 读取 Excel 文件为 DataFrame
def read_xls_to_df(path):
    check_file_exists(path)
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    data = [sheet.row_values(r) for r in range(sheet.nrows)]
    df = pd.DataFrame(data[1:], columns=data[0])
    return df


# 主数据读取（含水位、功率、参数等）
def read_main_data(path):
    df = read_xls_to_df(path)
    for col in ["库上水位(m)", "发电功率", "参数"]:
        df[col] = pd.to_numeric(df.get(col, np.nan), errors="coerce")
    return df


#  获取水头基值
def get_base_value(param):
    return base_values.get(param, np.nan)


#  插值获取效率
def interpolate_efficiency(water_level, data_list):
    if pd.isna(water_level):
        return np.nan

    data_list = sorted(data_list, key=lambda x: x[0])
    water_levels = np.array([x[0] for x in data_list])
    efficiencies = np.array([x[1] for x in data_list])

    if (
        len(water_levels) == 0
        or water_level < water_levels.min()
        or water_level > water_levels.max()
    ):
        print(f" 有效水位 {water_level} 超出效率表范围")
        return np.nan

    idx = np.where(np.isclose(water_levels, water_level))[0]
    if len(idx) > 0:
        return efficiencies[idx[0]]

    for i in range(len(water_levels) - 1):
        if water_levels[i] <= water_level <= water_levels[i + 1]:
            x0, x1 = water_levels[i], water_levels[i + 1]
            y0, y1 = efficiencies[i], efficiencies[i + 1]
            return y0 + (water_level - x0) * (y1 - y0) / (x1 - x0)
    return np.nan


#  主计算逻辑（有效水位、效率、出库流量）
def calculate_values(df_main):
    df_main["有效水位"] = df_main.apply(
        lambda row: (
            row["库上水位(m)"] - get_base_value(row["参数"])
            if pd.notna(row["参数"]) and pd.notna(row["库上水位(m)"])
            else np.nan
        ),
        axis=1,
    )

    def get_efficiency(row):
        if pd.isna(row["有效水位"]) or pd.isna(row["参数"]):
            return np.nan
        param = row["参数"]
        if param == 2:
            return interpolate_efficiency(
                row["有效水位"], power_efficiency_data["小水电"]["功率"]
            )
        elif param == 1:
            return interpolate_efficiency(
                row["有效水位"], power_efficiency_data["大水电"]["功率"]
            )
        else:
            print(f" 未知参数类型：{param}")
            return np.nan

    df_main["效率"] = df_main.apply(get_efficiency, axis=1)

    df_main["出库流量(m3/s)"] = df_main.apply(
        lambda row: (
            row["发电功率"] / (g * row["有效水位"] * row["效率"])
            if all(
                [
                    pd.notna(row["发电功率"]),
                    pd.notna(row["有效水位"]),
                    pd.notna(row["效率"]),
                    row["有效水位"] != 0,
                    row["效率"] != 0,
                ]
            )
            else np.nan
        ),
        axis=1,
    ).round(0)

    # 打印前10行测试样本
    print(" 测试样本（前10行）：")
    print(
        df_main[["参数", "库上水位(m)", "有效水位", "效率", "出库流量(m3/s)"]].head(10)
    )

    return df_main


# 📤 写入模板 Excel
def write_to_template(
    template_path, save_path, df, start_row=1, start_col=0, columns=None
):
    check_file_exists(template_path)
    rb = xlrd.open_workbook(template_path, formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(0)

    if columns is None:
        columns = df.columns.tolist()

    # 写入列名
    for c_idx, col_name in enumerate(columns):
        ws.write(start_row - 1, start_col + c_idx, col_name)

    # 写入数据
    for r_idx, row in enumerate(df[columns].values, start=start_row):
        for c_idx, val in enumerate(row):
            ws.write(r_idx, start_col + c_idx, "" if pd.isna(val) else val)

    wb.save(save_path)
    print(f"数据已写入并保存到: {save_path}")


#  主执行逻辑入口
if __name__ == "__main__":
    #  文件路径配置
    main_data_path = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\Source_Data\rsvrSample1.xls"
    template_path = (
        r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\temp\水库_模板文件.xls"
    )
    save_path = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace\Reservoir\rsvrSample1_temp.xls"

    # 步骤①：读取数据
    df_main = read_main_data(main_data_path)

    # 步骤②：计算有效水位、效率、出库流量
    df_main = calculate_values(df_main)

    # 步骤③：字段筛选并写入模板
    columns_to_write = [
        "站码",
        "时间",
        "库上水位(m)",
        "入库流量(m3/s)",
        "蓄水量(m6)",
        "库下水位(m)",
        "出库流量(m3/s)",
        "库水特征码",
        "库水水势",
        "入流时段长",
        "测流方法",
    ]

    write_to_template(
        template_path=template_path,
        save_path=save_path,
        df=df_main,
        start_row=1,
        start_col=0,
        columns=columns_to_write,
    )
