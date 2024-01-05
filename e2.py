import openpyxl
from collections import Counter

# 提示用户拖放 Excel 文件到命令行窗口
print("请拖放 Excel 文件到此命令行窗口，然后按 Enter 键:")

# 获取用户输入（拖放的文件路径）
excel_file_path = input().strip('\"')  # 去除路径两端可能存在的引号

# 打印路径检查
print(f"Trying to load Excel file from: {excel_file_path}")

# 打开Excel文件
workbook = openpyxl.load_workbook(excel_file_path)

# 选择第一个工作表
sheet = workbook.active

# 指定连续列范围（例如，A 列到 BB 列）
print('请输入开始的列：')
start_column = input()
print('请输入结束的列：')
end_column = input()
column_range = sheet[f"{start_column}:{end_column}"]

# 用于存储每个列的数据
column_data = {}

# 读取数据并存储在字典中
for column in column_range:
    column_letter = column[0].column_letter
    column_values = [cell.value for cell in column]
    column_data[column_letter] = column_values

# 打印每列的数据
for column_letter, values in column_data.items():
    print(f"Column {column_letter}: {values}")

# 统计每列中重复项的个数
for column_letter, values in column_data.items():
    counts = Counter(values)
    print(f"{column_letter}列的统计信息:")
    for value, count in counts.items():
        print(f"{value}: {count} 次")
    print()



# 创建新的工作簿
new_workbook = openpyxl.Workbook()
new_sheet = new_workbook.active


# 将每列的统计信息写入新工作表
for column, values in column_data.items():
    counts = Counter(values)
    new_sheet[f'{column}1'] = f"{column} 列的统计信息"
    row_index = 2
    for value, count in counts.items():
        new_sheet[f'{column}{row_index}'] = f"{value}: {count} 次"
        row_index += 1

# 保存新的Excel表格
import os
new_excel_file_path =os.path.join(os.path.dirname(excel_file_path), 'result.xlsx') 
new_workbook.save(new_excel_file_path)
print(f"结果已保存到: {new_excel_file_path}")

# 关闭工作簿
workbook.close()

# 等待用户按下任意键
input("按任意键退出...") 

