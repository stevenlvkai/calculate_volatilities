# calculate_volatilities
calculate the three kinds of volatilities of  42 futures in Chinese futures market
# -*- coding: utf-8 -*-
"""
Created on Thu Dec 21 13:02:50 2017

@author: DELL
"""

import pandas as pd  
from WindPy import *  
import datetime,time  
import os 
import numpy as np
import xlwt

def calculate_volatilities(index,start_date,end_date):
    #由wind导入数据
    w.start()
    stock=w.wsd(index, "open,high,low,close", start_date,end_date)  
    index_data = pd.DataFrame()  
    index_data['open'] =stock.Data[0]  
    index_data['high'] =stock.Data[1]  
    index_data['low']  =stock.Data[2]  
    index_data['close']=stock.Data[3]  
    num=len(index_data['close'])#计算时间长度
    num1=num-1#考虑自由度
    
    #计算历史波动率1
    index_data['log_ret']=np.log(index_data['close']/index_data['close'].shift(1))#对数收益率
    index_data['log_ret']=np.nan_to_num(index_data['log_ret'])#去除第一个交易日的空数据
    mean=np.mean(index_data['log_ret'])
    index_data['log_ret']=index_data['log_ret']-mean
    index_data['historical_volatility']=(index_data['log_ret'])**2
    asum=np.nansum(index_data['historical_volatility'])
    a=np.sqrt(asum/num1)

    #计算Parkinson波动率
    index_data['parkinson_volatility']=(np.log(index_data['high']/index_data['low']))**2
    b=np.sqrt((sum(index_data['parkinson_volatility'])/(4*np.log(2)))/num)
    
    #计算真实平均波动率
    index_data['comp_volatility']=0
    comp_volatility=[0]
    for i in range(num1):
        a1=abs(index_data['close'][i+1]-index_data['close'][i])
        a2=abs(index_data['high'][i+1]-index_data['close'][i])
        a3=abs(index_data['low'][i+1]-index_data['close'][i])
        a4=abs(index_data['high'][i+1]-index_data['low'][i+1])
        if index_data['close'][i+1]>index_data['close'][i]:            
            a0=max(a2,a4)/index_data['close'][i]
        elif index_data['close'][i+1]<index_data['close'][i]:
            a0=-max(a3,a4)/index_data['close'][i]
        else:
            if index_data['close'][i+1]>=index_data['open'][i+1]:
                a0=a4/index_data['close'][i]
            else:
                a0=-a4/index_data['close'][i]              
        comp_volatility.append(a0)
    index_data['comp_volatility']=comp_volatility
    mean=np.mean(index_data['comp_volatility'])
    index_data['comp_volatility']=index_data['comp_volatility']-mean
    index_data['comp_volatility']=(index_data['comp_volatility'])**2
    c=np.sqrt(sum(index_data['comp_volatility'])/num1)
    volatilities=[a,b,c]
    return volatilities
if __name__=='__main__':
    historical_volatilities=['historical volatilities']
    parkinson_volatilities=['parkinson volatilities']
    average_true_volatilities=['average true volatilities']
    index=['RBFI.WI', 'ZNFI.WI', 'HCFI.WI', 'CUFI.WI', 'ALFI.WI', 'PBFI.WI', \
           'NIFI.WI', 'AUFI.WI', 'BUFI.WI', 'AGFI.WI', 'SNFI.WI', 'RUFI.WI', \
           'JFI.WI', 'CSFI.WI', 'CFI.WI', 'JDFI.WI', 'JMFI.WI', 'VFI.WI', \
           'PPFI.WI', 'MFI.WI', 'LFI.WI', 'IFI.WI', 'AFI.WI', 'PFI.WI',  \
           'YFI.WI', 'SFFI.WI', 'ZCFI.WI', 'SMFI.WI', 'FGFI.WI', 'MAFI.WI', \
           'RMFI.WI', 'CFFI.WI', 'TAFI.WI', 'OIFI.WI', 'SRFI.WI', 'WHFI.WI', \
           'IH.CFE', 'IF.CFE', 'IC.CFE', 'TF.CFE', 'T.CFE']
    
    num=len(index)
    for i in range(num):
        volatilities=calculate_volatilities(index[i],'20170101','20171220')
        historical_volatilities.append(volatilities[0])
        parkinson_volatilities.append(volatilities[1])
        average_true_volatilities.append(volatilities[2])
    
    #棉纱期货今年8月正式交易，单独计算
    volatilities_cy=calculate_volatilities('CYFI.WI','20170821','20171220')
    index.append('CYFI.WI')
    historical_volatilities.append(volatilities_cy[0])
    parkinson_volatilities.append(volatilities_cy[1])
    average_true_volatilities.append(volatilities_cy[2])
    
    #输出excel
    index=['index']+index    
    num_index=len(index)
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet('volatilities')
    
    for i in range(num_index):
        worksheet.write(i, 0, index[i])
        worksheet.write(i, 1, historical_volatilities[i])
        worksheet.write(i, 2, parkinson_volatilities[i])
        worksheet.write(i, 3, average_true_volatilities[i])
    workbook.save('D:\\中粮期货-吕凯\\波动收益\\volatilities2.xls')
    
