import pandas as pd

def prepare_accured() -> list:
    """读取预提表并生成中间文件，随后生成预提表清单

    Returns:
        list: 带宽与非带宽的预提表清单
    """
    filenames = ['2305', '2306', '2307', '2308', '2309']
    no_bandwidth_list = []
    bandwidth_list = []

    for filename in filenames:
        try:
            expense_bandwidth = pd.read_pickle(f'pkl 文件/{filename} 带宽.pkl')
            expense_no_bandwidth = pd.read_pickle(f'pkl 文件/{filename} 非带宽.pkl')
        except FileNotFoundError:
            expense_bandwidth = pd.read_excel(f'/Users/zhuangyuhao/Documents/Fileport/预算表/{filename} 带宽.xlsx')
            expense_no_bandwidth = pd.read_excel(f'/Users/zhuangyuhao/Documents/Fileport/预算表/{filename} 非带宽.xlsx')
            expense_bandwidth.to_pickle(f'{filename} 带宽.pkl')
            expense_no_bandwidth.to_pickle(f'{filename} 非带宽.pkl')

        no_bandwidth_list.append(expense_no_bandwidth)
        bandwidth_list.append(expense_bandwidth)

    return no_bandwidth_list, bandwidth_list
