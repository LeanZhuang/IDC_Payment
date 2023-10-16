
import os
import glob
import pandas as pd

# 设置下载文件夹路径和目标文件夹路径
download_folder = "/Users/zhuangyuhao/Downloads"
desktop_folder = "/Users/zhuangyuhao/Desktop"

# 获取下载文件夹内最新的xls文件
latest_file = max(glob.glob(os.path.join(download_folder, "*.xls")), key=os.path.getctime)

# 读取xls文件
df = pd.read_excel(latest_file)

# 构建目标文件路径，将文件后缀名改为xlsx
target_file = os.path.join(desktop_folder, os.path.basename(latest_file).replace(".xls", ".xlsx"))

# 将数据保存为xlsx文件
df.to_excel(target_file, index=False)

# 清理内存以释放资源 (可选)
del df