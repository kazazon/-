import os
from openpyxl import Workbook

def find_txt_files(folder_path):
    txt_files = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".txt"):
                txt_file_path = os.path.abspath(os.path.join(root, file))
                txt_files.append((file, txt_file_path))

    return txt_files

# 文件夹路径
folder_path = "."

# 获取所有.txt文件的文件名和绝对路径
txt_files = find_txt_files(folder_path)

# 创建Excel工作簿
wb = Workbook()
ws = wb.active

# 设置表头
ws.append(["文件名", "链接"])

# 将.txt文件的文件名和链接写入Excel文件
for file_name, absolute_path in txt_files:
    link = '=HYPERLINK("' + absolute_path + '","' + file_name + '")'
    ws.append([file_name, link])

# 保存Excel文件
excel_file_path = "txt_files.xlsx"
wb.save(excel_file_path)
print(f"Excel文件已保存到: {excel_file_path}")

input("Press Enter to exit...")