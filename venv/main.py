import os
from modules.io_utils import read_xls_to_df, write_to_template
from modules.calculator import calculate_values
from modules.constants import columns_to_write
# 混合出流模式1，非混合出流模式2
# 设置路径（保持原有数据来源）
main_data_path = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace_\Source_Data\rsvrSample1.xls"
template_path = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace_\temp\水库_模板文件.xls"
save_path = r"C:\Users\Administrator\Desktop\Insert_Data_Workspace_\Reservoir\rsvrSample1_temp.xls"

def main():
    if not os.path.exists(main_data_path):
        print(f" 数据文件不存在: {main_data_path}")
        return
    if not os.path.exists(template_path):
        print(f" 模板文件不存在: {template_path}")
        return

    print(" 读取数据中...")
    df = read_xls_to_df(main_data_path)

    print(" 开始计算有效水位、效率和出库流量...")
    df = calculate_values(df)

    print(" 写入模板文件中...")
    write_to_template(
        template_path=template_path,
        save_path=save_path,
        df=df,
        start_row=1,
        start_col=0,
        columns=columns_to_write,
    )

    print(f" 处理完成，文件已保存至：{save_path}")

if __name__ == "__main__":
    main()
