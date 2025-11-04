#bar chart-monthly(comparison)
from math import ceil
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from bidi.algorithm import get_display
import arabic_reshaper
def bar_cahrt_comparison_monthly(fileAddressOfMainData,firstYear,lastYear,exportDataPath,company_list,firstMonth,lastMonth,name):

    df=pd.read_excel('D:/razaghi/Products_list.xlsx')
    products_list=dict(df)
    barcode=np.array(products_list['barcode'])
    #barcode=[1150010006]
    Coefficient=np.array(products_list['amount per package'])


    df1=pd.read_excel(fileAddressOfMainData)
    df2=pd.read_excel(ad2)
    data=dict(df1)
    totalSalesData=np.array([data['Column 1-Date'],data['Column 2-Product name'],
                      data['Column 3-Product barcode'],data['Column 4-Sale amount'],
                      data['Column 5-Customer code'],data['Column 6-Warehouse name']])


    x_1=[]

    for i in range(int(firstMonth),int(lastMonth)+1):
        summ=0
        for j in range(len(barcode)):
            for k in range(0,len(totalSalesData[2,:])):
                if int(str(totalSalesData[0,k][5:7]))== i and barcode[j]==totalSalesData[2,k] and firstYear[-1] == int(str(totalSalesData[0, k][3]))  and totalSalesData[4,k] in company_list and totalSalesData[5,k]!='انبار مرجوعی' and totalSalesData[5,k]!='انبار ضایعات':
                    summ=summ+ceil((totalSalesData[3,k])/Coefficient[j])
                else:
                    continue
        x_1.append([i,summ])
    xx_1=np.array(x_1)


    x_2=[]

    for i in range(int(firstMonth),int(lastMonth)+1):
        summ=0
        for j in range(len(barcode)):
            
            for k in range(0,len(totalSalesData[2,:])):
                if int(str(totalSalesData[0,k][5:7]))==i and barcode[j]==totalSalesData[2,k] and lastYear[-1] == int(str(totalSalesData[0, k][3])) and totalSalesData[4,k] in company_list and totalSalesData[5,k]!='انبار مرجوعی' and totalSalesData[5,k]!='انبار ضایعات':
                    summ=summ+ceil((totalSalesData[3,k])/Coefficient[j])
                else:
                    continue
        x_2.append([i,summ])
    xx_2=np.array(x_2)
    month=np.array(['فروردین','اردیبهشت','خرداد','تیر','مرداد','شهریور','مهر','آبان','آذر','دی','بهمن','اسفند'])
    x1=[get_display(arabic_reshaper.reshape(d)) for d in month]
    width=0.35
    x11 = np.arange(len(x1))
    y1=xx_1[:,1]
    y2=xx_2[:,1]
    fig, ax = plt.subplots()
    bar1 = ax.bar(x11[int(firstMonth)-1:int(lastMonth)] - width/2, y1, width, label=firstYear)
    bar2 = ax.bar(x11[int(firstMonth)-1:int(lastMonth)] + width/2, y2, width, label=lastYear)
    x = np.arange(len(xx_1[:,0]))
    
   
    plt.rcParams["figure.figsize"] = (25,5)

    for line_x in x11[int(firstMonth)-1:int(lastMonth)]:
        
        plt.axvline(x=line_x + 1.5*width, color='r', linestyle='-.')
    
    ax.set_xlabel(get_display(arabic_reshaper.reshape('ماه')),fontsize=20)
    ax.set_ylabel(get_display(arabic_reshaper.reshape('تعداد کارتن')),fontsize=20)
    ax.set_title(get_display(arabic_reshaper.reshape(f'نمودار میله‌ای فروش کارتنی 12 ماهه اول سال‌های 1403 و 1404-{name}')), fontsize=25)
    ax.spines[['right', 'top','left']].set_visible(False)
    ax.set_yticklabels([])
    ax.set_yticks([])
    
    ax.set_xticks(x11[int(firstMonth)-1:int(lastMonth)+1],x1[int(firstMonth)-1:int(lastMonth)+1],fontsize=20)
    plt.rcParams['font.family'] = 'B nazanin'
    for bar in bar1:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., yval, '%d' % yval, ha='center', va='bottom', fontsize='25')

    for bar in bar2:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., yval, '%d' % yval, ha='center', va='bottom', fontsize='25')
    ax.legend(fontsize=15)
    plt.savefig(f'D:/razaghi/شهروند/report/{name}_monthly_comparison.png', bbox_inches='tight')
    aaa=[]
    for hg in range(0,len(xx_2)):
        changes="{:.2f}".format(((xx_2[hg,1]-xx_1[hg,1])/xx_1[hg,1]))
        aaa.append([hg,xx_1[hg,1],xx_2[hg,1],changes])
    aaa=np.array(aaa)
    comparison_monthly_data={
        'ماه':list(aaa[:,0]),
        '1402':list(aaa[:,1]),
        '1403':list(aaa[:,2]),
        'تغییرات (درصد)': list(aaa[:,3])
           }
    comparison_monthly_data=pd.DataFrame(comparison_monthly_data)
    file_path = exportDataPath  # Define the file path and name
    comparison_monthly_data.to_excel(file_path, index=False)
    print(name)
#########################################################################################################################################################
#bar_cahrt_comparison_monthly(adress1,name1,adress2,name2,file path,company)
 company_list=pd.read_excel('D:/razaghi/company_list.xlsx')
 company_list_data=dict(company_list)
 names=list(company_list_data.keys())

for name in names:
    
    company_code_list=[company_list_data[name]]
    bar_cahrt_comparison_monthly('D:/example/example_total sales.xlsx','2024','2025',
                                 f'D:/example/{name}-monthly.xlsx', company_code_list,1,3,name)