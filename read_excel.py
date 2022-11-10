# excel_calculations
import pandas as pd

vekting = [1, 1, 1, 1, 1, 1, 1, 1, 1]
poeng_max = [100, 100, 100, 100, 100, 100, 100, 100, 100]
poeng_min = [50, 50, 50, 50, 50, 50, 50, 50, 50]

fileread = pd.read_excel('Skjema_Lyseparken.xlsx')
data = fileread.rename(columns={'Navn på bedrift': 'Navn'})
data = data.set_index('Navn')

for m,mm in enumerate(data):
    #print(m)

    #print(mm)
    #print(data[mm])

    if m == 0:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] <= 50:
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
            elif data.iloc[n,m] > 50:  
                data.iloc[n,m] = vekting[m]*poeng_max[m] 


    if m == 1:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] <= 200:
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
            elif data.iloc[n,m] > 200:  
                data.iloc[n,m] = vekting[m]*poeng_max[m]

    if m == 2:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] < 0:
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
            elif data.iloc[n,m] >= 0:  
                data.iloc[n,m] = vekting[m]*poeng_max[m] 
    if m == 3:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] in ['j', 'J', 'ja', 'Ja', 'JA']:
                data.iloc[n,m] = vekting[m]*poeng_max[m] 
            elif data.iloc[n,m] in ['n', 'N', 'nei', 'Nei', 'NEI']:  
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
    if m == 4:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] in ['j', 'J', 'ja', 'Ja', 'JA']:
                data.iloc[n,m] = vekting[m]*poeng_max[m] 
            elif data.iloc[n,m] in ['n', 'N', 'nei', 'Nei', 'NEI']:  
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
    if m == 5:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] in ['j', 'J', 'ja', 'Ja', 'JA']:
                data.iloc[n,m] = vekting[m]*poeng_max[m] 
            elif data.iloc[n,m] in ['n', 'N', 'nei', 'Nei', 'NEI']:  
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
            else: 
                data.iloc[n,m] = 0
    if m == 6:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] in ['j', 'J', 'ja', 'Ja', 'JA']:
                data.iloc[n,m] = vekting[m]*poeng_max[m] 
            elif data.iloc[n,m] in ['n', 'N', 'nei', 'Nei', 'NEI']:  
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
            else: 
                data.iloc[n,m] = 0
    if m == 7:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] in ['j', 'J', 'ja', 'Ja', 'JA']:
                data.iloc[n,m] = vekting[m]*poeng_max[m] 
            elif data.iloc[n,m] in ['n', 'N', 'nei', 'Nei', 'NEI']:  
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
    if m == 8:
        for n,nn in enumerate(data[mm].index):
            print(data.iloc[n,m])
            if data.iloc[n,m] in ['ROS', 'ros', 'Ros']:
                data.iloc[n,m] = vekting[m]*poeng_min[m] 
            elif data.iloc[n,m] in ['TCFD', 'tcfd']:  
                data.iloc[n,m] = vekting[m]*poeng_max[m]
            else: 
                data.iloc[n,m] = 0 
            

data.to_excel('Lyseparken_poeng.xlsx')
