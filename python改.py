import pandas as pd  

df = pd.read_excel('D:/Users/Tina/Downloads/2303_UMC.xlsx')
print(df.columns)


t = 1

sl = 19 #停損
N = 60 #最低持倉日
sw = 5
pf = [] #利潤
box = [] #[[0,110.5,0],[持倉天數,買入價,判斷買或賣:0=未賣,1 = 已賣]]
box_2 = []
num = [] #暫時
num_2 = []
AR=[] #年化報酬

po = [] #放10年報酬率最慘的



rk_1 = 0
positive = 0
negetive = 0
mm=0
jj = 0 #賣的次數
jo = 0 #正常買賣AR
ko = 0 #放空AR
cc = 0#總共交易次數

while t<366:
                 
                 
                 
                #愉快的一天又開始了，今天有股票要買嗎?
                
                 if (df['Ps'][t-1]<=df['Pm'][t-1] and df['Ps'][t] > df['Pm'][t]) and (df['C'][t] > df['Pm'][t] and df['C'][t]>max(df['Pca'][t],df['Pcb'][t])):
                     box.append([0,df['O'][t+1],0])   #有則放入倉庫
                          
                #那今天有股票要放空嗎?
                 if (df['Ps'][t-1]>=df['Pm'][t-1] and df['Ps'][t]<df['Pm'][t]) and (df['C'][t]<df['Pm'][t] and df['C'][t]<min(df['Pca'][t],df['Pcb'][t])):
                     box_2.append([0,df['O'][t+1],0])   #有則放入倉庫
                
                    
                
                
                #檢查倉庫有沒有要賣的股票
                 for i in range(len(box)): 
                     
                     buy = box[i][1] #第i個商品的買入價
                     
                     if (((df['C'][t] - buy)/buy)<=(-sl/100)) or(((df['C'][t] - buy)/buy)>=(sw/100)) or((df['Ps'][t]<df['Pm'][t] and box[i][0]>=N)):
                         jj+=1   #賣出次數
                         sell=df['O'][t+1]  #賣出價
                         r = (sell-buy)/buy  #賣出利潤
                         '''
                         print("--------------正常買賣---------------")
                         print("")
                         print(df['date'][t])
                         print("報酬率 = ",r)
                         print("持倉天數 = ",box[i][0])
                         print("報酬率/持蒼天數  = ",r/box[i][0])
                         print("sell =",sell," buy = ",buy)
                         print("ar = ",(r/box[i][0])*365)
                         
                         print("")
                         '''
                         r = (r/box[i][0])*365
                         jo+=r
                         
                         pf.append(r)
                         box[i][2] =1   #標示此商品已賣出
                         cc+=1
                         
       
                           
                #檢查倉庫有沒有要平倉的股票
                 for m in range(len(box_2)): 
                     sell_2 = box_2[m][1] #第m個商品的賣出價
                     if (((df['C'][t] - sell_2)/sell_2)>=(sl/100)) or((df['Ps'][t] > df['Pm'][t] and box_2[m][0]>=N)):
                         mm+=1 #放空次數
                         buy_2=df['O'][t+1]  #買回價
                         rk = (sell_2-buy_2)/buy_2  #偿券利潤
                         
                         rk = (rk/box_2[m][0])*365
                         ko+=rk
                         pf.append(rk)
                         box_2[m][2] =1   #標示此商品已賣出
                         print("--------------放空區---------------")
                         print("")
                         print(df['date'][t])
                         print("報酬率 = ",rk)
                         print("day = ",box_2[m][0])
                         print("rk/day  = ",rk/box_2[m][0])
                         print("放空價 =",sell_2," 買回價 = ",buy_2)
                         print("ar = ",rk)
                         
                         print("")
                         cc+=1
                         
                       
                    
                    
                    
                    
                #整理倉庫，把已賣出商品清掉
                 for k in range(len(box)):
                     if box[k][2]!=1:   
                         num.append(box[k])  
                 box = num
                 num = []
'''
                #整理倉庫，把已平倉商品清掉
                 for p in range(len(box_2)):
                     if box_2[p][2]!=1:
                         num_2.append(box_2[p]) 
                        
                 box_2 = num_2
                 num_2 = []
'''
                            
                     
                
                       
                          
                #沒賣出的股票又要持倉了       
                for j in range(len(box)):
                     box[j][0]+=1
                '''
                #沒賣出的股票又要持倉了       
                 for w in range(len(box_2)):
                     box_2[w][0]+=1
                    '''
                    
                    
                 #今日又過了一天
                 t+=1
            

             #算出這個條件的年化報酬
c = sum(pf)/len(pf)
            
#把各個條件的AR放入AR串列
AR.append([c,N,sl,sw])
print("正常買賣平均AR =",jo/jj)  
print("放空平均AR =",ko/mm )      
print("放空次數 = ",mm)  
print("正常買賣次數 = ",jj)  
print("總交易次數 =",cc)      
      


print(max(AR))


print(mm)
#改了賣出價就好很多 可見當事情不對 隔天立即反映