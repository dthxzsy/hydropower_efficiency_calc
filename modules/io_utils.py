import os
import xlrd
from xlutils.copy import copy
import pandas as pd
import numpy as np

# 路径存在性检查
def check_file_exists(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f" 文件不存在: {path}")


# 读取 Excel (.xls) 为 DataFrame（只读取第一个 sheet）
def read_xls_to_df(path):
    check_file_exists(path)
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    data = [sheet.row_values(r) for r in range(sheet.nrows)]
    df = pd.DataFrame(data[1:], columns=data[0])
    return df


# 写入 DataFrame 到模板 Excel (.xls)，保留格式
def write_to_template(template_path, save_path, df, start_row=1, start_col=0, columns=None):
    check_file_exists(template_path)
    rb = xlrd.open_workbook(template_path, formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(0)

    if columns is None:
        columns = df.columns.tolist()

    # 写列名
    for c_idx, col_name in enumerate(columns):
        ws.write(start_row - 1, start_col + c_idx, col_name)

    # 写数据
    for r_idx, row in enumerate(df[columns].values, start=start_row):
        for c_idx, val in enumerate(row):
            ws.write(r_idx, start_col + c_idx, "" if pd.isna(val) else val)

    # 确保保存目录存在
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    # 保存文件
    wb.save(save_path)
    print(f" 数据已写入并保存到: {save_path}")
