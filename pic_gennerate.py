import pandas as pd
import csv_process
import Histogram
import Boxplot
import Sweep

data_csv, USL, LSL = csv_process.csv_process('D:\LB\Python Interface\X2061_INLINE_TRAP4_V20_BY Config_2_1NF_示例数据.csv')

config_name = 'Main2'
data_ini = pd.read_excel(r'D:\LB\Python Interface\X2061_Configuration File_V20_TRAP4_20210629.xlsx')

graph_flag = ''     # 画什么图
graph_class = ''    # 图的分类标签


boxplot_data_tem = pd.DataFrame()
sweep_data_tem = pd.DataFrame()
USL_tem = []
LSL_tem = []

pring_info = ''

for i in range(len(data_ini['Data Item'])):

    # 画图种类定义+标签
    if data_ini.loc[i, 'Plot Type'] == data_ini.loc[i, 'Plot Type']:
        graph_flag = data_ini.loc[i, 'Plot Type']
        graph_class = data_ini.loc[i, 'Plot Title']

    item_name = data_ini.loc[i, 'Data Item']
    if graph_flag == 'Histogram' and data_ini.loc[i, 'Monitor'] == data_ini.loc[i, 'Monitor']:  # Monitor列判定是否画图

        # USL 配置文件有，则用配置文件，无则用数据USL
        if data_ini.loc[i,'Upper Limit'] == data_ini.loc[i,'Upper Limit']:
            USL_histogram = data_ini.loc[i, 'Upper Limit']
        else:
            USL_histogram = USL[item_name]

        # LSL 配置文件有，则用配置文件，无则用数据LSL
        if data_ini.loc[i,'Lower Limit'] == data_ini.loc[i,'Lower Limit']:
            LSL_histogram = data_ini.loc[i, 'Lower Limit']
        else:
            LSL_histogram = LSL[item_name]

        Histogram.histogram(config_name, graph_class, item_name,USL_histogram,LSL_histogram,data_csv[item_name], data_ini.loc[i, 'Axis Upper Limit'], data_ini.loc[i, 'Axis Lower Limit'])

    elif graph_flag == 'Boxplot' and data_ini.loc[i, 'Monitor'] == data_ini.loc[i, 'Monitor']:# Monitor列判定是否画图
        boxplot_data_tem = pd.concat([boxplot_data_tem, data_csv[item_name]], axis=1)

        # USL 配置文件有，则用配置文件，无则用数据USL
        if data_ini.loc[i, 'Upper Limit'] == data_ini.loc[i, 'Upper Limit']:
            USL_tem.append(data_ini.loc[i, 'Upper Limit'])
        else:
            USL_tem.append(USL[item_name])

        # LSL 配置文件有，则用配置文件，无则用数据LSL
        if data_ini.loc[i,'Lower Limit'] == data_ini.loc[i,'Lower Limit']:
            LSL_tem.append(data_ini.loc[i, 'Lower Limit'])
        else:
            LSL_tem.append(LSL[item_name])

        if data_ini.loc[i+1, 'Plot Type'] == data_ini.loc[i+1, 'Plot Type']:
            Boxplot.Boxplot(config_name, graph_class, USL_tem, LSL_tem, boxplot_data_tem, data_ini.loc[i, 'Axis Upper Limit'], data_ini.loc[i, 'Axis Lower Limit'])
            boxplot_data_tem = pd.DataFrame()
            USL_tem = []
            LSL_tem = []

    elif graph_flag == 'Sweep' and data_ini.loc[i, 'Monitor'] == data_ini.loc[i, 'Monitor']:  # Monitor列判定是否画图
        for j in data_csv.columns:
            if j.find(item_name) >=0:
                sweep_data_tem = pd.concat([sweep_data_tem, data_csv[j]], axis=1)
                USL_tem.append(USL[j])
                LSL_tem.append(LSL[j])
        Sweep.Sweep(config_name, graph_class, USL_tem, LSL_tem, sweep_data_tem, data_ini.loc[i, 'Axis Upper Limit'], data_ini.loc[i, 'Axis Lower Limit'])
        sweep_data_tem = pd.DataFrame()
        USL_tem = []
        LSL_tem = []