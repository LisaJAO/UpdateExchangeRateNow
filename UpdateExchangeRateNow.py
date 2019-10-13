#!/usr/bin/env python
# coding: utf-8

# In[9]:


import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as tm
import pandas
import datetime

currency=pandas.read_html('https://rate.bot.com.tw/xrt?Lang=en-US')
currency=currency[1] 
currency=currency.iloc[:,0:5] 
currency.columns=['currency','cash_buying','cash_selling','spot_buying','spot_selliing'] 
currency['currency'] =currency['currency'].str.extract('\((\w+)\)')  
currency= currency.set_index(['currency', 'cash_buying','cash_selling','spot_buying','spot_selliing'])
date=datetime.datetime.now()
currency.to_csv('Currency_CSV.csv', encoding='utf_8_sig')
file1 = open('Currency_CSV.csv', 'r').readlines()
fileout = open('Currency_CSV.csv', 'w')
for line in file1:
    fileout.write(line.replace('-', '0'))
fileout.close()
def currency_rate_Export():    
    currency.to_excel('Currency Rate{:4}{:02}{:02}-Daily.xlsx'.format(date.year,date.month,date.day))
    currencyfile='Currency Rate{:4}{:02}{:02}-Daily.xlsx'.format(date.year,date.month,date.day)
    result_date="Update: Date {:4}.{:02}.{:02}-Time {:02}:{:02}:{:02}".format(date.year,date.month,date.day,date.hour,date.minute,date.second)
    tm.showinfo(title =str(result_date), message ='最新匯率已匯出,檔名為: '+ currencyfile)

    
def rbCountry1_1(): #點選have幣別選項按鈕後處理函式
    n1=0
    for c1 in data["currency"]:  #逐一取出have幣別
        if(c1 == Currency_name.get()):
            Currency_list.append(data.ix[n1, 0])
        n1 += 1    
    n = 0
    for rbCountry1_select in data.ix[:, 0]:  #逐一取得have幣別
        if(rbCountry1_select == Currency_name.get()):  #取得點選的have幣別  
            rate = data.ix[n, "cash_selling"]  #取得匯率
            #break  #找到點選have幣別就離開迴圈
        n += 1 
    r1_1=rate
    rbCountry3() #顯示訊息
    
def rbCountry2_1():  #點選want幣別選項按鈕後處理函式
    n1=0
    for c1 in data["currency"]:  #逐一取出選取want幣別選項
        if(c1 == Currency_name2.get()):
            Currency2_list.append(data.ix[n1, 0])
            n1 += 1    
    n = 0
    for rbCountry2_select in data.ix[:, 0]:  #逐一取得want幣別選項
        if(rbCountry2_select == Currency_name2.get()):  #取得點選的want幣別選項
            #break  #找到點選want幣別就離開迴圈
            n += 1
    rbCountry3() #顯示訊息
        
def rbCountry3(): #點選want幣別 / 選項按鈕後 顯示訊息    
    n1_1=0
    for c1 in data["currency"]:  #逐一取出have幣別
        if(c1 == Currency_name.get()):
            Currency_list.append(data.ix[n1_1, 0])
        n1_1 += 1    
    n0_1 = 0
    for rbCountry1_select in data.ix[:, 0]:  #逐一取得have幣別
        if(rbCountry1_select == Currency_name.get()):  #取得點選的have幣別
            rate = data.ix[n0_1, "cash_selling"]  #取得匯率
            break  #找到點選have幣別就離開迴圈
        n0_1 += 1 
    r1_1=rate
    n2_1=0
    for c1 in data["currency"]:  #逐一取出選取want幣別選項
        if(c1 == Currency_name2.get()):
            Currency2_list.append(data.ix[n2_1, 0])
        n2_1 += 1    
    n0_2 = 0
    for rbCountry2_select in data.ix[:, 0]:  #逐一取得want幣別選項
        if(rbCountry2_select == Currency_name2.get()):  #取得點選的want幣別選項
            rate = data.ix[n0_2, "cash_selling"]  #取得匯率
            break  #找到點選want幣別選項就離開迴圈
        n0_2 += 1
    r2_1=rate

    rate=float(r1_1)/float(r2_1)
    if(float(r1_1)==0 or float(r2_1)==0): #如果沒有資料
        result1.set("Exchange Rate: " + rbCountry1_select + ":"+ rbCountry2_select + " => Sorry, No information !")
        resultNTD.set("Exchange Rate: " + rbCountry1_select + " : NTD => Sorry, No information !")
    else:  
        result1.set("Exchange Rate: " + rbCountry1_select + " : " + rbCountry2_select + " = " + str(round(rate,4)))
        resultNTD.set("Exchange Rate: " + rbCountry1_select + " : NTD = "+ str(round(r1_1,4)))

    
def clickRefresh():  #重新讀取資料
    global data
    data = pandas.read_csv('Currency_CSV.csv')

data = pandas.read_csv('Currency_CSV.csv')

my_window=tk.Tk()
my_window.geometry("600x550")
my_window.title("Welcome Lisa's World - Update Exchange Rate Now")
my_window.iconbitmap('Exchange.ico')

photo=tk.PhotoImage(file='currency conventer.png')
background_window=tk.Label(my_window,
                          text='Welcome\nCurrency Converter',
                          image=photo,
                          compound=tk.CENTER,
                          font=('Calibri',20,'bold italic'),
                          fg='black')

background_window.pack() 

Currency_name = tk.StringVar()  #have幣別文字變數
Currency_name2 = tk.StringVar()  #want幣別文字變數
result1 = tk.StringVar()  #外匯兌外匯文字訊息變數
resultNTD= tk.StringVar()  #台幣兌外匯文字訊息變數
Currency_list = []  #have 貨幣串列
Currency2_list = []  #want 貨幣串列

#建立have貨幣串列
for c1 in data["currency"]:  
    if(c1 not in Currency_list):  #如果串列中無該貨幣就將其加入
        Currency_list.append(c1)
        
#建立第1個貨幣的want貨幣串列
count = 0
for c2 in data["currency"]:  
    if(c2 ==  Currency_list[0]):  #是第1個貨幣的的want
        Currency2_list.append(data.ix[count, 0])
    count += 1

label1 = tk.Label(my_window, text="Which currency you have：", 
                  pady=6, fg="blue", font="Calibri 14")
label1.pack()
lblResult2 = tk.Label(my_window, textvariable=resultNTD, fg="red", font="Calibri 16")
lblResult2.pack(pady=6)


frame1 = tk.Frame(my_window)
frame1.pack()

for i in range(0,2):  #3列選項按鈕
    for j in range(0,10):  #每列8個選項按鈕
        n = i * 10 + j  #第n個選項按鈕
        if(n < len(Currency_list)):
            currency_1 = Currency_list[n]  #取得have幣別名稱
            rbtem = tk.Radiobutton(frame1, text=currency_1, 
                                   variable=Currency_name, 
                                   value=currency_1, command=rbCountry1_1)  #建立選項按鈕
            rbtem.grid(row=i, column=j)  #設定選項按鈕位置
            if(n==0):  #選取第1個have幣別
                rbtem.select()

                
                
                
label2 = tk.Label(my_window, text="Which currency you want ：", 
                  pady=6, fg="blue", font="Calibri 14")
label2.pack()
frame2=tk.Frame(my_window)
frame2.pack()

for ii in range(0,2):  #3列選項按鈕
    for j in range(0,10):  #每列8個選項按鈕
        n = ii * 10 + j  #第n個選項按鈕
        if(n < len(Currency_list)):
            currency_2 = Currency_list[n]  #取得want幣別名稱
            rbtem2 = tk.Radiobutton(frame2, text=currency_2, 
                                   variable=Currency_name2, 
                                   value=currency_2, command=rbCountry2_1)  #建立選項按鈕
            rbtem2.grid(row=ii, column=j)  #設定選項按鈕位置
            if(n==0):  #選取第1個want幣別
                rbtem2.select()

btnDown = tk.Button(my_window, text="Update Now", font="Calibri 14", command=clickRefresh)
btnDown.pack(pady=6)

lblResult1 = tk.Label(my_window, textvariable=result1, fg="red", font="Calibri 16")
lblResult1.pack(pady=6)

rbCountry3()  #顯示匯兌訊息

btn2=tk.Button(my_window,text="Exchange rate list-Export (Excel)",font="Calibri 14",
               width=25,height=1,command=currency_rate_Export)
btn2.pack() 

my_window.mainloop()


# In[ ]:




