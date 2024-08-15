import numpy as np
import pandas as pd


np.printoptions(suppress=True)
pd.set_option('display.max_columns', 100)
file_path = './data/CA_EV_registration.csv'
df = pd.read_csv(file_path)
registration_data = df.loc[:, ['County GEOID', 'Registration Valid Date']]
registration_data[['County GEOID', 'Registration Valid Date']] = registration_data[['County GEOID', 'Registration Valid Date']].astype(str)
# print(registration_data.head())
# print(df['County GEOID'].unique())
# print(df['Registration Valid Date'].unique())


data = np.zeros([116, 11]).astype(np.int32)

for row in registration_data.itertuples():
    i = getattr(row, '_1')
    j = getattr(row, '_2')

    try:
        # print(i)
        # print(int(i[-3:]))
        # print(j)
        # print(int(j[0:4])-2010)
        data[int(i[-3:]), int(j[0:4])-2010] += 1
    except ValueError:
        if i == 'Unknown':
            data[0, (int(j[0:4]) - 2010)] += 1
            # print('success')
        # else:
            # print(i)
            # print(j)
        continue

print(data)
np.savetxt('./results/registration_quantity.csv', data, delimiter=',', fmt='%d')
