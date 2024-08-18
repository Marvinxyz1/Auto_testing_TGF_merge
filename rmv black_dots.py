import os

# 目标文件夹路径
folder_path = '/Users/marvin/Desktop/py/work/Test/tgf_offical_test'

# 要匹配的数字列表
target_numbers = [
    '01-0001', '01-0002', '01-0005', '01-0006', '01-0008', '01-0012', '01-0013', '01-0014',
    '02-0001', '02-0002', '03-0001', '03-0002', '03-0003', '03-0005', '03-0006', '03-0008',
    '03-0012', '03-0014'
]

# 将目标数字列表中的连字符替换为下划线，生成新的匹配列表
target_numbers_underscore = [num.replace('-', '_') for num in target_numbers]

# 合并两个匹配列表
all_target_numbers = target_numbers + target_numbers_underscore

# 获取文件夹中的所有文件
files = os.listdir(folder_path)

# 识别并删除包含目标数字的文件
for file in files:
    for target in all_target_numbers:
        if target in file:
            file_path = os.path.join(folder_path, file)
            os.remove(file_path)
            print(f"已删除文件: {file}")
            break