import numpy as np
import pandas as pd
from .constants import g, base_values
from .efficiency_data import power_efficiency_data
from .interpolation import interpolate_efficiency


# 🔽 数据读取函数
def read_main_data(path):
    from .io_utils import read_xls_to_df  # 避免循环依赖，内部导入
    df = read_xls_to_df(path)

    # 强制转换关键列为数值类型
    for col in ["库上水位(m)", "发电功率", "参数"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")  # 转换为数值
        else:
            print(f"⚠️ 警告：列 {col} 不存在，将填充 NaN")
            df[col] = np.nan

    print("📊 数据类型检查：")
    print(df[["库上水位(m)", "发电功率", "参数"]].dtypes)

    return df


# 🔽 获取水头基值
def get_base_value(param):
    return base_values.get(param, np.nan)


# 🔽 主计算逻辑
def calculate_values(df_main):
    # 防止字符串类型引起计算错误
    df_main["库上水位(m)"] = pd.to_numeric(df_main["库上水位(m)"], errors="coerce")

    # 计算有效水位
    df_main["有效水位"] = df_main.apply(
        lambda row: (
            row["库上水位(m)"] - get_base_value(row["参数"])
            if pd.notna(row["参数"]) and pd.notna(row["库上水位(m)"])
            else np.nan
        ),
        axis=1,
    )

    # 插值查效率
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
            print(f"❓ 未知参数类型：{param}")
            return np.nan

    df_main["效率"] = df_main.apply(get_efficiency, axis=1)

    # 计算出库流量 Q = N / (g * η * h)
    df_main["出库流量(m3/s)"] = df_main.apply(
        lambda row: (
            row["发电功率"] / (g * row["有效水位"] * row["效率"])
            if all([
                pd.notna(row["发电功率"]),
                pd.notna(row["有效水位"]),
                pd.notna(row["效率"]),
                row["有效水位"] != 0,
                row["效率"] != 0,
            ])
            else np.nan
        ),
        axis=1,
    ).round(0)

    return df_main
