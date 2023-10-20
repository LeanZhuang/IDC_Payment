import csv
import codecs

# 定义输入文件和输出文件的路径
input_file = 'input.csv'
output_file = 'output.csv'

# 打开输入文件以读取 gbk 编码的 CSV 数据
with codecs.open(input_file, 'r', encoding='gbk') as f_in:
    # 读取CSV数据
    csv_reader = csv.reader(f_in)
    data = [row for row in csv_reader]

# 打开输出文件以写入 utf-8 编码的 CSV 数据
with codecs.open(output_file, 'w', encoding='utf-8') as f_out:
    # 写入CSV数据
    csv_writer = csv.writer(f_out)
    csv_writer.writerows(data)
