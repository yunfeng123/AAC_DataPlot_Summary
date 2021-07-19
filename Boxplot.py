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


def Boxplot(Station, pic_dic, config_name, graph_class, USL, LSL, data, Upper_Axis, Lower_Axis, boxplot_X_name,
            ini_pic_order, N):
    X = range(1, len(LSL) + 1, 1)

    col_name = data.columns.values.tolist()

    for i in np.arange(len(col_name)):
        col_name[i] = col_name[i].split(' ')[-1].split('_')[0]

    data_N = len(data)
    if N == '':
        N_label = data_N
    else:
        N_label = int(N)

    title_name = data.columns[-1].split(' ')[0]
    title_1st = title_name + " N:%d" % N_label + '\n'
    title_2nd = 'CPK:['

    std_list = []
    std_list_label = []
    cpk_list_label = []
    mean_list = []
    cpk_list = []
    USL = pd.Series(USL, index=data.columns)
    LSL = pd.Series(LSL, index=data.columns)
    for i in data.columns:
        mean_value, std_value, cpk_value = cpk(LSL[i], USL[i], data[i])
        std_list_label.append("'" + str(round(std_value, 2)) + "'")
        cpk_list_label.append("'" + str(round(cpk_value, 2)) + "'")
        std_list.append(std_value)
        mean_list.append(mean_value)
        cpk_list.append(cpk_value)

    save_box = pd.DataFrame([data.columns, USL, LSL, mean_list, std_list, cpk_list],
                            index=['Item', 'USL', 'LSL', 'Mean', 'STDEV', 'CPK'])

    #   title_2nd = title_2nd + ','.join(std_list_label) + ']'
    title_2nd = title_2nd + ','.join(cpk_list_label) + ']'

    title = title_1st + title_2nd

    plt.figure(figsize=(15, 15))
    plt.boxplot(x=data, showmeans=True, meanline=False,
                flierprops=dict(marker='x', markersize=5, linestyle='none', markeredgecolor='r'))

    # Average Line
    plt.plot(range(1, len(col_name) + 1), mean_list, color='#000000', linestyle='--')

    # 图片Y轴范围设置，配置文件有则用配置文件数据
    if Upper_Axis == Upper_Axis:
        plt.ylim(top=Upper_Axis)
    if Lower_Axis == Lower_Axis:
        plt.ylim(bottom=Lower_Axis)

    # USL & LSL
    plt.plot(X, USL, color='r', linestyle='-', label='USL')
    plt.plot(X, LSL, color='r', linestyle='-', label='LSL')

    for a, b in zip(X, mean_list):
        plt.text(a + 0.5, b, '%.2f' % b, ha='center', va='bottom', fontsize=15)

    if len(boxplot_X_name) > 0:
        plt.xticks(range(1, len(col_name) + 1), boxplot_X_name, rotation=45, fontsize=18)
    else:
        plt.xticks(range(1, len(col_name) + 1), col_name, rotation=45, fontsize=18)

    plt.title(title, fontdict=dict(family='Times New Roman', fontsize=25))
    plt.grid(axis='both')

    file_path = pic_dic + graph_class + '-' + config_name + f'-{ini_pic_order}-' + Station + '-' + title_name + '.png'
    plt.savefig(file_path)
    plt.close()
    return ('Finished -> ' + 'Boxplot ' + title_name), save_box
