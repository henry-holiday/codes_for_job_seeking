import os
import pandas as pd

# 全局变量用于存储所有文件的处理结果
global_result_data = {
    "Vendor Name": [],  # 供应商名称
    "RW_WF_ID": [],
    "RW_BINA_FAB_WF_ID": [],
    "RW_WF_GOOD_DIE_QTY": [],
    "生產日期": []  # RW_WF_SORT_DATE 的中文翻译
}

def process_lum_dot(file_path):
    """处理单个 COC 文件"""
    global global_result_data

    # 读取文件
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        df = pd.read_excel(file_path)
    else:
        print(f"Unsupported file format: {file_path}")
        return

    # 初始化当前文件的结果字典
    result_data = {
        "Vendor Name": None,  # 供应商名称
        "RW_WF_ID": [],
        "RW_BINA_FAB_WF_ID": [],
        "RW_WF_GOOD_DIE_QTY": [],
        "生產日期": []  # RW_WF_SORT_DATE 的中文翻译
    }

    # 提取目标字段
    target_fields = ["VENDOR_NAME", "RW_WF_ID", "RW_WF_SORT_DATE", "RW_WF_GOOD_DIE_QTY", "RW_BINA_FAB_WF_ID"]
    for field in target_fields:
        try:
            # 查找字段所在行
            row_index = df[df.iloc[:, 0] == field].index[0]
            # 获取对应的值（向右移动一列或多列）
            values = df.iloc[row_index, 1:].dropna().tolist()
            
            # 特殊处理 RW_WF_GOOD_DIE_QTY，确保其为数值类型
            if field == "RW_WF_GOOD_DIE_QTY":
                numeric_values = []
                for value in values:
                    try:
                        numeric_value = int(value)  # 尝试转换为整数
                    except ValueError:
                        try:
                            numeric_value = float(value)  # 尝试转换为浮点数
                        except ValueError:
                            print(f"Value '{value}' cannot be converted to a number. Skipping...")
                            numeric_value = None  # 转换失败时使用 None 填充
                    numeric_values.append(numeric_value)
                values = numeric_values

            if field == "RW_WF_SORT_DATE":
                result_data["生產日期"] = values
            elif field == "VENDOR_NAME":
                result_data["Vendor Name"] = values[0]  # 记录供应商名称
            else:
                result_data[field] = values
        except IndexError:
            print(f"Field '{field}' not found in the file: {file_path}")
            continue

    # 动态调整列长度
    max_length = max(len(values) for key, values in result_data.items() if isinstance(values, list))
    for key, values in result_data.items():
        if isinstance(values, list):
            result_data[key] += [None] * (max_length - len(values))  # 填充空值以对齐长度

    # 将当前文件的结果追加到全局结果中
    vendor_name = result_data["Vendor Name"]
    for i in range(max_length):
        global_result_data["Vendor Name"].append(vendor_name)
        global_result_data["RW_WF_ID"].append(result_data["RW_WF_ID"][i] if i < len(result_data["RW_WF_ID"]) else None)
        global_result_data["RW_BINA_FAB_WF_ID"].append(result_data["RW_BINA_FAB_WF_ID"][i] if i < len(result_data["RW_BINA_FAB_WF_ID"]) else None)
        global_result_data["RW_WF_GOOD_DIE_QTY"].append(result_data["RW_WF_GOOD_DIE_QTY"][i] if i < len(result_data["RW_WF_GOOD_DIE_QTY"]) else None)
        global_result_data["生產日期"].append(result_data["生產日期"][i] if i < len(result_data["生產日期"]) else None)

# 主程序：遍历文件夹中的所有文件
folder_path = os.getcwd()  # 当前文件夹
for file_name in os.listdir(folder_path):
    if file_name.endswith(('.csv', '.xlsx', '.xls')):
        process_lum_dot(os.path.join(folder_path, file_name))

# 构建最终的 DataFrame
final_df = pd.DataFrame(global_result_data)

# 保存最终结果文件，文件名固定加上 Lumentum
output_file = "Lumentum_All_COC_Results.xlsx"
final_df.to_excel(output_file, index=False)
print(f"All results saved to: {output_file}")
