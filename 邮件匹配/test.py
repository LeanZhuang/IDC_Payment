import re

# 读取文本文件
with open('./output.txt', 'r') as file:
    text = file.read()

# 使用正则表达式匹配项目号
idc_list = []
p = re.compile(r'(IDC{1}[a-zA-Z0-9-]+)')
for one in p.findall(text):
    idc_list.append(one)
idc_list = list(set(idc_list))

print(idc_list)

po_list = []
p = re.compile(r'(Z?PO{1}[a-zA-Z0-9-]+)')
for one in p.findall(text):
    po_list.append(one)
po_list = list(set(po_list))

print(po_list)

ac_list = []
p = re.compile(r'(AC{1}[a-zA-Z0-9-]+)')
for one in p.findall(text):
    ac_list.append(one)
ac_list = list(set(ac_list))

print(ac_list)
