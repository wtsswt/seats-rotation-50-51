import openpyxl

# 打开 Excel 文件
input_file_path = 'seats_51.xlsx'  # 需修改
wb = openpyxl.load_workbook(input_file_path)

#起始工作簿名称(Sheet1就输入a=1)
a = 1 #需修改

# 需要生成的座位表数量
n = 100 #需修改
b = a+n

# 工作表表格名
input_sheet_name = f"Sheet{a}"

# 定义剪切和粘贴操作
operations = [
    ('C1', 'H8'),
    ('G8', 'F11'),
    ('F8', 'G11'),
    ('H8', 'B12'),
    ('B8:E8', 'C12:F12'),
    ('H2:H7', 'I2:I7'),
    ('G1:G7', 'H2:H8'),
    ('F1:F7', 'G2:G8'),
    ('E2:E7', 'F3:F8'),
    ('D2:D7', 'E3:E8'),
    ('C2:C7', 'D3:D8'),
    ('B2:B7', 'C3:C8'),
    ('I2:I7', 'B3:B8'),
    ('B12:F12', 'B2:F2'),
    ('F11:G11', 'F1:G1'),
    ('H8', 'B2'),
]

# 辅助函数，用于剪切和粘贴单元格的内容
def cut_paste(ws, src, dest):
    if ':' in src:
        src_cells = ws[src]
        dest_cells = ws[dest]
        for src_row, dest_row in zip(src_cells, dest_cells):
            for src_cell, dest_cell in zip(src_row, dest_row):
                dest_cell.value = src_cell.value
                src_cell.value = None
    else:
        ws[dest].value = ws[src].value
        ws[src].value = None

# 复制和处理工作表的循环
for i in range(a, b):
    output_sheet_name = f'Sheet{i+1}'

    # 选择要复制的工作表
    input_sheet = wb[input_sheet_name]

    # 复制工作表并命名为'Sheet{i+1}'
    output_sheet = wb.copy_worksheet(input_sheet)
    output_sheet.title = output_sheet_name

    # 激活特定的工作表
    ws = wb[output_sheet_name]

    # 执行所有操作
    for src, dest in operations:
        cut_paste(ws, src, dest)

    # 更新 input_sheet_name 为下一个工作表的名称
    input_sheet_name = output_sheet_name

# 保存修改后的 Excel 文件
wb.save(input_file_path)

# 打印完成信息
print("已完成")