# -*- coding: utf-8 -*-
"""
Created on Thu Dec 24 15:19:07 2020

@author: Tina
"""

import pandas as pd  

df = pd.read_excel('D:/Users/Tina/Downloads/2303_UMC.xlsx')
print(df.columns)

flag = 0
t = 1
sl = 5 #停損
sw = 5 #停利
n = 1 #持倉日

buy = 0  #進場價
sell = 0 #賣出價
N = 1 #最低持倉日
pf = [] #利潤
case = [] #[[0,110.5],[3,89]]
k=0

while t<=5174:
    #time to enter
    flag=0
    buy = 0
    sell = 0
    if df['Ps'][t-1]<=df['Pm'][t-1] and df['C'][t]>df['Pm'][t] and df['C'][t]>max(df['Pca'][t],df['Pcb'][t]):
        buy = df['O'][t+1]
        case.append([0,buy])
        
        print('買入時點 : %s, 買入價格 :%f'%(df['date'][t+1],buy))
        t+=1
        for i in range(len(case)):
            case[i][0]+=1  
    else:
        for i in range(len(case)):
            case[i][0]+=1  
        t+=1
    
    for m in range(len(case)):
        
        if ((df['C'][t] - case[m][1])/case[m][1])<=(-sl/100) or (df['Ps'][t]<df['Pm'][t] and case[m][0]>=N):
            
            sell = df['O'][t+1]
            k = (sell - case[m][1])/case[m][1]
            pf.append(k)
                
            print('賣出時點 : %s, 賣出價格 :%f'%(df['date'][t+1],sell))
            print('報酬率 = %.3f percent'%(100*k))
            t+=1
            case.pop(m)
        else:
            t+=1
print(case)


