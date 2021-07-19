import pandas as pd
import time

# Data Frame 横向拼接Series
# out_data = pd.DataFrame()
# data3 = pd.Series([9,9,7,7,4,9],name='K', index=['Item', 'USL', 'LSL', 'Mean', 'STDEV', 'CPK'])
# data2 = out_data.append(data3)
# print(data2)

# Data Frame 纵向拼接Series
# out_data = pd.DataFrame()
# data3 = pd.Series([9,9,7,7,4,9], name='K', index=['Item', 'USL', 'LSL', 'Mean', 'STDEV', 'CPK'])
# data2 = pd.concat([out_data, data3], axis=1)
# print(data2)

# Data Frame 纵向拼接DataFrame
# out_data = pd.DataFrame()
# data3 = pd.DataFrame([[9,9,7,7,4,9], [9,9,7,7,4,9]], index=['K1', 'K2'], columns=['Item', 'USL', 'LSL', 'Mean', 'STDEV', 'CPK'])
# data2 = pd.concat([out_data, data3], axis=1)
# print(data2.transpose())

# Data Frame Mean, STD
out_data = pd.DataFrame()
data3 = pd.DataFrame([[9,9,7,7,4,9], [9,9,7,7,4,9]], index=['K1', 'K2'], columns=['Item', 'USL', 'LSL', 'Mean', 'STDEV', 'CPK'])
data2 = pd.concat([out_data, data3], axis=1)
print(data2)
print(data2.mean() + data2.std())