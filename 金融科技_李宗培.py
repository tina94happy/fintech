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
AR = []
while t<=5174:
    #time to enter
    flag=0
    buy = 0
    sell = 0
    if df['Ps'][t-1]<=df['Pm'][t-1] and df['C'][t]>df['Pm'][t] and df['C'][t]>max(df['Pca'][t],df['Pcb'][t]):
        buy = df['O'][t+1]
        n=0
        print('買入時點 : %s, 買入價格 :%f'%(df['date'][t+1],buy))
      
        while flag == 0 :
            #case 1 or case 2 or case 3 happen
            if ((df['C'][t+n] - buy)/buy)<=(-sl/100) or (df['Ps'][t+n]<df['Pm'][t+n] and n>=N):       
                sell = df['O'][t+n+1]
                k = (sell - buy)/buy
                pf.append(k)
                flag = 1
                print('賣出時點 : %s, 賣出價格 :%f'%(df['date'][t+1+n],sell))
                print('報酬率 = %.3f percent'%(100*k))
                t+=1
                
            #none of them happen
            else:
                t+=1
                n+=1
                
                
    #It is not the great time to enter.
    else:
        t+=1
#print(pf)



