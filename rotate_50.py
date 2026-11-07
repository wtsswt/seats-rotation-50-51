import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

# 打开 Excel 文件
input_file_path = 'seats_50.xlsx'  # 需修改
wb = openpyxl.load_workbook(input_file_path)

#起始工作簿名称(Sheet1就输入a=1)
a = 1 #需修改

# 需要生成的座位表数量
n = 100 #需修改
b = a+n

def custom_rotate(ws):
    special_moves = {
        "F7": "G2",
        "G7": "H1",
        "H7": "G1",
    }
    all_seats = []

    # 收集所有座位，排除特殊格
    for row in range(2, 8):
        for col in ['B','C','D','E','F','G','H','I']:
            cell = f"{col}{row}"
            if cell not in special_moves:
                all_seats.append(cell)
    # 加上G1, H1
    all_seats.extend(['G1','H1'])

    # 读取所有值
    values = {cell: ws[cell].value for cell in all_seats + list(special_moves.keys())}

    new_positions = {}

    # 普通格向右下移动规则
    for cell in all_seats:
        col = column_index_from_string(cell[0])
        row = int(cell[1:])
        # 右移一格，超过I列（9）回B列（2）
        new_col = col + 1 if col < 9 else 2
        # 下移一格，超过7回2
        new_row = row + 1 if row < 7 else 2
        new_cell = f"{get_column_letter(new_col)}{new_row}"
        new_positions[new_cell] = values[cell]

    # 特殊格固定映射
    for src, dest in special_moves.items():
        new_positions[dest] = values[src]

    # 只清空将被写入的新位置的单元格
    for cell in new_positions:
        ws[cell].value = None

    # 写回新值
    for cell, val in new_positions.items():
        ws[cell].value = val

# 复制和处理工作表的循环
for i in range(a, b):
    output_sheet_name = f'Sheet{i+1}'

    # 选择要复制的工作表
    input_sheet = wb[f"Sheet{i}"]

    # 如果目标工作表已存在，则先删除
    if output_sheet_name in wb.sheetnames:
        std = wb[output_sheet_name]
        wb.remove(std)

    # 复制工作表并命名为'Sheet{i+1}'
    output_sheet = wb.copy_worksheet(input_sheet)
    output_sheet.title = output_sheet_name

    # 激活特定的工作表
    ws = wb[output_sheet_name]

    # 调用定制轮换函数
    custom_rotate(ws)

# 保存修改后的 Excel 文件
wb.save(input_file_path)

# 打印完成信息
print("已完成")