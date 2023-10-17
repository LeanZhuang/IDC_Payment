import pandas as pd

def prepare_accured():
    try:
        expense_2305_bandwidth = pd.read_pickle('pkl 文件/2305 带宽.pkl')
        expense_2305_no_bandwidth = pd.read_pickle('pkl 文件/2305 非带宽.pkl')

        expense_2306_bandwidth = pd.read_pickle('pkl 文件/2306 带宽.pkl')
        expense_2306_no_bandwidth = pd.read_pickle('pkl 文件/2306 非带宽.pkl')

        expense_2307_bandwidth = pd.read_pickle('pkl 文件/2307 带宽.pkl')
        expense_2307_no_bandwidth = pd.read_pickle('pkl 文件/2307 非带宽.pkl')

        expense_2308_bandwidth = pd.read_pickle('pkl 文件/2308 带宽.pkl')
        expense_2308_no_bandwidth = pd.read_pickle('pkl 文件/2308 非带宽.pkl')

        expense_2309_bandwidth = pd.read_pickle('pkl 文件/2309 带宽.pkl')
        expense_2309_no_bandwidth = pd.read_pickle('pkl 文件/2309 非带宽.pkl')

    except FileNotFoundError:
        # 2306
        expense_2305_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2305 带宽.xlsx')
        expense_2305_no_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2305 非带宽.xlsx')
        expense_2305_bandwidth.to_pickle('2305 带宽.pkl')
        expense_2305_no_bandwidth.to_pickle('2305 非带宽.pkl')

        # 2306
        expense_2306_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2306 带宽.xlsx')
        expense_2306_no_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2306 非带宽.xlsx')
        expense_2306_bandwidth.to_pickle('2306 带宽.pkl')
        expense_2306_no_bandwidth.to_pickle('2306 非带宽.pkl')

        # 2307
        expense_2307_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2307 带宽.xlsx', sheet_name='202307带宽')
        expense_2307_no_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2307 非带宽.xlsx')
        expense_2307_bandwidth.to_pickle('2307 带宽.pkl')
        expense_2307_no_bandwidth.to_pickle('2307 非带宽.pkl')

        # 2308
        expense_2308_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2308 带宽.xlsx')
        expense_2308_no_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2308 非带宽.xlsx')
        expense_2308_bandwidth.to_pickle('2308 带宽.pkl')
        expense_2308_no_bandwidth.to_pickle('2308 非带宽.pkl')

        # 2309
        expense_2309_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2309 带宽.xlsx')
        expense_2309_no_bandwidth = pd.read_excel('/Users/zhuangyuhao/Documents/Fileport/预算表/2309 非带宽.xlsx')
        expense_2309_bandwidth.to_pickle('2309 带宽.pkl')
        expense_2309_no_bandwidth.to_pickle('2309 非带宽.pkl')

    no_bandwidth_list = [expense_2305_no_bandwidth, expense_2306_no_bandwidth, expense_2308_no_bandwidth, expense_2307_no_bandwidth, expense_2309_no_bandwidth]
    bandwidth_list = [expense_2305_bandwidth, expense_2306_bandwidth, expense_2308_bandwidth, expense_2307_bandwidth, expense_2309_bandwidth]


    return no_bandwidth_list, bandwidth_list