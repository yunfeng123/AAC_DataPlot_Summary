import pandas as pd
import time

#path = 'D:/桌面文件/0712/X2219_INLINE_TRAP4_V1_BY Config_2_B2 Main1.csv'
path = 'D:/桌面文件/0712/DEVELOPMENT4_ACWJ_EP4-3FT-01_2_DEVELOPMENT4_X2219_INLINE_TRAP4_V1_2021-06-26.csv'

t0 = time.time()
data_head = pd.read_csv(path, header=1, delimiter=',', nrows=5)
USL = data_head.iloc[2]
LSL = data_head.iloc[3]

reader = pd.read_csv(path, delimiter=',', iterator=True, header=None, skiprows=7)

data = pd.DataFrame()
#开多线程？
loop = True
while loop:
    try:
        chunk_0 = reader.get_chunk(5)
        data = pd.concat([data, chunk_0], axis=0)
    except StopIteration:
        loop = False


print(time.time() - t0)

