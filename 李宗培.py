# -*- coding: utf-8 -*-
"""
Created on Mon Dec 28 23:59:35 2020

@author: Tina
"""

import pandas as pd  

df = pd.read_excel('D:/Users/Tina/Downloads/2303_UMC.xlsx')
print(df.columns)


t = 1
sl = 5 #停損
N = 1 #最低持倉日
pf = [] #利潤
box = [] #[[0,110.5,0],[持倉天數,買入價,判斷買或賣:0=未賣,1 = 已賣]]
num = [] #暫時
AR=[]



while t<=5174:
    
    
    
    #愉快的一天又開始了，今天有股票要買嗎?
    
    if df['Ps'][t-1]<=df['Pm'][t-1] and df['C'][t]>df['Pm'][t] and df['C'][t]>max(df['Pca'][t],df['Pcb'][t]):
        box.append([0,df['C'][t],0])   #有則放入倉庫
    
        
    
    
    #檢查倉庫有沒有要賣的股票
    for i in range(len(box)):      
        buy = box[i][1] #第i個商品的買入價
        if ((df['C'][t] - buy)/buy)<=(-sl/100) or (df['Ps'][t]<df['Pm'][t] and box[i][0]>=N):
            sell=df['C'][t]  #賣出價
            r = (sell-buy)/buy  #賣出利潤
            pf.append(r)
            box[i][2] =1   #標示此商品已賣出
    
    
    
    
    #整理倉庫，把已賣出商品清掉
    for k in range(len(box)):
        if box[k][2]!=1:
            num.append(box[i])  
    box = num
    num = []
     

       
            
    #沒賣出的股票又要持倉了       
    for j in range(len(box)):
        box[j][0]+=1
    
    
    
    #今日又過了一天
    t+=1
c = (sum(pf)/len(pf))*365
AR.append(c)
print(pf)




            
            