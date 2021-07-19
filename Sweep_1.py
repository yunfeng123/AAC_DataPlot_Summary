import matplotlib.pyplot as plt
import pandas as pd

# Sweep 矩阵转置后直接画，可以提高效率
def Sweep(Station, pic_dic, config_name, graph_class, USL_List, LSL_List, data_Sweep, Upper_Axis, Lower_Axis, ini_pic_order, N):
    data_column_index = data_Sweep.columns.values.tolist()
    title_1st = (data_column_index[0][data_column_index[0].find(' '):data_column_index[0].find('@')])

    for i in range(len(data_column_index)):
        data_column_index[i] = round(float(data_column_index[i].split('@')[-1]), 2)

    USL = pd.Series(USL_List, index=data_column_index)
    LSL = pd.Series(LSL_List, index=data_column_index)

    # 对数据，LSL，USL根据数据点进行排序
    data_Sweep.columns = data_column_index

    data_Sweep.sort_index(axis=1, ascending=True, inplace=True)
    USL.sort_index(ascending=True, inplace=True)
    LSL.sort_index(ascending=True, inplace=True)

    data_ave = data_Sweep.mean()
    data_ave_3std_P = data_ave + 3 * data_Sweep.std()
    data_ave_3std_N = data_ave - 3 * data_Sweep.std()

    plt.plot(data_Sweep.columns, data_Sweep.transpose(), color='DeepSkyBlue', lw='0.5')

    plt.plot(data_Sweep.columns, USL, color='r', lw='2', ls='--')
    plt.plot(data_Sweep.columns, LSL, color='r', lw='2', ls='--')

    plt.plot(data_Sweep.columns, data_ave, color='Black', lw='1', ls='-')
    plt.plot(data_Sweep.columns, data_ave_3std_P, color='HotPink', lw='1', ls='-')
    plt.plot(data_Sweep.columns, data_ave_3std_N, color='HotPink', lw='1', ls='-')

    plt.xlabel('Frequency')
    plt.ylabel('Magnitude')

    # 图片Y轴范围设置，配置文件有则用配置文件数据
    if Upper_Axis == Upper_Axis:
        plt.ylim(top=Upper_Axis)
    if Lower_Axis == Lower_Axis:
        plt.ylim(bottom=Lower_Axis)

    data_N = len(data_Sweep)
    if N == '':
        N_label = data_N
    else:
        N_label = int(N)

    title = title_1st + '(N=' + str(N_label) + ')'
    plt.title(title)
    plt.grid(axis='both')
    file_path = pic_dic + graph_class + '-' + config_name + f'-{ini_pic_order}-' + Station + '-' + title_1st + '.png'
    plt.savefig(file_path)
    plt.close()
    return 'Finished -> ' + 'Sweep ' + title_1st

    return file_path
