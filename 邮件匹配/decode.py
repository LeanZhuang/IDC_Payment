import codecs

# 指定输入文件的编码为 GBK，输出文件的编码为 UTF-8
input_file = './测试.txt'
output_file = 'output.txt'

with codecs.open(input_file, 'r', encoding='gbk') as file_in:
    with codecs.open(output_file, 'w', encoding='utf-8') as file_out:
        for line in file_in:
            file_out.write(line)
