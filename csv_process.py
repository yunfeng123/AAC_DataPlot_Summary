import matplotlib.pyplot as plt
import pandas as pd
import os
import xlwings


# 行标签为Config，列标签为Item_Name

def csv_process(csv_path):
    data_raw = pd.read_csv(csv_path, delimiter=',', skiprows=1, low_memory=False)

    USL = data_raw.iloc[2]
    LSL = data_raw.iloc[3]

    # Delete Fail Iitems

    if 'ALERT_CFG' in data_raw.index:
        config_flag = 0
    else:
        config_flag = 1

    for i in data_raw.index:
        if data_raw.loc[i, 'Test Pass/Fail Status'] == 'FAIL':
            data_raw.drop(i, axis=0, inplace=True)

    data = data_raw.iloc[5:]

    Overlay = data['Version'].tolist()[0]
    Station = Overlay.split('_')[2]
    project = Overlay.split('_')[0]
    return data, USL, LSL, Overlay, Station, project
