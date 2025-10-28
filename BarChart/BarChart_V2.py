#bar chart-monthly(comparison)
from math import ceil
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from bidi.algorithm import get_display
import arabic_reshaper
def bar_cahrt_comparison_monthly(FileAddressMainDatabase,first_considered_year,
                                 second_considered_year,saving_file_path,company_list,
                                 first_number_of_month,last_number_of_month,name):

    df=pd.read_excel('D:/razaghi/Products_list.xlsx')
    products_list=dict(df)
    barcode=np.array(products_list['barcode'])
    Coefficient=np.array(products_list['amount per package'])


    df1=pd.read_excel(FileAddressMainDatabase)
    data1_1=dict(df1)
    data2_1=np.array([data1_1['Column 1-Date'],data1_1['Column 2-Product name'],
                      data1_1['Column 3-Product barcode'],data1_1['Column 4-Sale amount'],
                      data1_1['Column 5-Customer code'],data1_1['Column 6-Warehouse name']])


    x_1=[]

    for i in range(int(first_number_of_month),int(last_number_of_month)+1):
        summ=0
        for j in range(len(barcode)):
            for k in range(0,len(data2_1[2,:])):
                if int(str(data2_1[0,k][5:7]))== i and barcode[j]==data2_1[2,k] and first_considered_year[-1] == int(str(data2_1[0, k][3]))  and data2_1[4,k] in company_list and data2_1[5,k]!='انبار مرجوعی' and data2_1[5,k]!='انبار ضایعات':
                    summ=summ+ceil((data2_1[3,k])/Coefficient[j])
                else:
                    continue
        x_1.append([i,summ])
    xx_1=np.array(x_1)


    x_2=[]

    for i in range(int(first_number_of_month),int(last_number_of_month)+1):
        summ=0
        for j in range(len(barcode)):
            
            for k in range(0,len(data2_1[2,:])):
                if int(str(data2_1[0,k][5:7]))==i and barcode[j]==data2_1[2,k] and first_considered_year[-1] == int(str(data2_1[0, k][3])) and data2_1[4,k] in company_list and data2_1[5,k]!='انبار مرجوعی' and data2_1[5,k]!='انبار ضایعات':
                    summ=summ+ceil((data2_1[3,k])/Coefficient[j])
                else:
                    continue
        x_2.append([i,summ])
    xx_2=np.array(x_2)
    
    #Names of the Persian months
    month=np.array(['فروردین','اردیبهشت','خرداد','تیر','مرداد','شهریور','مهر','آبان','آذر','دی','بهمن','اسفند'])
    
    #Names of the Gregorian months
    #month=np.array(['January','February','March','April','May','June','July','August','September','October','November','December'])
    
    x1=[get_display(arabic_reshaper.reshape(d)) for d in month]
    width=0.35
    x11 = np.arange(len(x1))
    y1=xx_1[:,1]
    y2=xx_2[:,1]
    fig, ax = plt.subplots()
    bar1 = ax.bar(x11[int(first_number_of_month)-1:int(last_number_of_month)] - width/2, y1, width, label=str(first_considered_year))
    bar2 = ax.bar(x11[int(first_number_of_month)-1:int(last_number_of_month)] + width/2, y2, width, label=str(second_considered_year))
    x = np.arange(len(xx_1[:,0]))
    
   
    plt.rcParams["figure.figsize"] = (25,5)

    for line_x in x11[int(first_number_of_month)-1:int(last_number_of_month)]:
        
        plt.axvline(x=line_x + 1.5*width, color='r', linestyle='-.')
    
    ax.set_xlabel(get_display(arabic_reshaper.reshape('Month')),fontsize=20)
    ax.set_ylabel(get_display(arabic_reshaper.reshape('Quantity (cartons)')),fontsize=20)
    ax.set_title(get_display(arabic_reshaper.reshape(f'title-{name}')), fontsize=25)
    ax.spines[['right', 'top','left']].set_visible(False)
    ax.set_yticklabels([])
    ax.set_yticks([])
    
    ax.set_xticks(x11[int(first_number_of_month)-1:int(last_number_of_month)+1],x1[int(first_number_of_month)-1:int(last_number_of_month)+1],fontsize=20)
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
        'month':list(aaa[:,0]),
        f'{first_considered_year}':list(aaa[:,1]),
        f'{second_considered_year}':list(aaa[:,2]),
        'Changes(Percentage)': list(aaa[:,3])
           }
    comparison_monthly_data=pd.DataFrame(comparison_monthly_data)
    file_path = saving_file_path  # Define the file path and name
    comparison_monthly_data.to_excel(file_path, index=False)
    print(name)
#########################################################################################################################################################
company_list=pd.read_excel('D:/example/company_code_list.xlsx')
names=list(company_list.keys())

for name in names:
    
    company_code_list=[company_list[name]]
    bar_cahrt_comparison_monthly('D:/example/total-1401 to 1404.xlsx','1403','1404',
                                 f'D:/example/{name}-monthly.xlsx',company_code_list,1,3,name) 