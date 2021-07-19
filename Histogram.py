import numpy as np
import matplotlib.pyplot as plt
import pandas as pd


def cpk(LSL, USL, data):
    data_average = np.average(data)
    data_std = np.std(data)
    if LSL == LSL and USL == USL:
        return data_average, data_std, round(
            min((USL - data_average) / 3 / data_std, ((data_average - LSL) / 3 / data_std)), 2)
    elif USL == USL:
        return data_average, data_std, round((USL - data_average) / 3 / data_std, 2)
    elif LSL == LSL:
        return data_average, data_std, round((data_average - LSL) / 3 / data_std, 2)
    else:
        return data_average, data_std, np.nan


def histogram(Station, pic_dic, config_name, graph_class, item_name, USL, LSL, data_list, Upper_Axis, Lower_Axis,
              ini_pic_order, N):
    data_average, data_std, data_cpk = cpk(LSL, USL, data_list)

    data_N = len(data_list)

    if N == '':
        N_label = data_N
    else:
        N_label = int(N)

    title_1st = 'Mean:%.2f  Std:%.2f \n' % (data_average, data_std)
    title_2nd = 'CPK:' + str(data_cpk) + " N:%d" % N_label + '\n'
    title_3th = 'Limit:[' + str(LSL) + ' , ' + str(USL) + ']'

    save_his = pd.Series([item_name, USL, LSL, data_average, data_std, data_cpk],
                         index=['Item', 'USL', 'LSL', 'Mean', 'STDEV', 'CPK'])

    plt.figure(figsize=(4.5, 6.5))
    plt.hist(x=data_list, bins='auto', color='#00BFFF', alpha=1, density=True, edgecolor='#000000')

    if USL != 'NA':
        plt.axvline(float(USL), color='#FF0000', lw=1, ls='--')

    if LSL != 'NA':
        plt.axvline(float(LSL), color='#FF0000', lw=1, ls='--')

    # 图片X轴范围设置，配置文件有则用配置文件数据
    if Upper_Axis == Upper_Axis:
        plt.xlim(right=Upper_Axis)
    if Lower_Axis == Lower_Axis:
        plt.xlim(left=Lower_Axis)

    plt.title(title_1st + title_2nd + title_3th)
    plt.grid(axis='y')
    plt.xlabel(item_name)
    file_path = pic_dic + graph_class + '-' + config_name + f'-{ini_pic_order}-' + Station + '-' + item_name + '.png'
    plt.savefig(file_path)
    plt.close()
    return ('Finished -> ' + 'Histogram ' + item_name), save_his
