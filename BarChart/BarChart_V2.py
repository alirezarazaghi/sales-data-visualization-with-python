import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from bidi.algorithm import get_display
import arabic_reshaper
from math import ceil



def BarChart_V2(fileAddress_1,first_considered_year,fileAddress_2,
                                 second_considered_year,saving_file_path,company_code_list,
                                 first_number_of_month,last_number_of_month):
    company_list=pd.read_excel('D:/razaghi/company_list.xlsx')
    company_list_data=dict(company_list)
    names=list(company_list_data.keys())
    print("names: ",names)

    for name in names:
        df=pd.read_excel('D:/razaghi/Products_list.xlsx')
        products_list=dict(df)
        df1=pd.read_excel('D:/razaghi/total-1401 to 1404.xlsx')
        data1=dict(df1)
        data2=np.array([data1['Column 1-Date'],data1['Column 2-Product name'],
                          data1['Column 3-Product barcode'],data1['Column 4-Sale amount'],
                          data1['Column 5-Customer code'],data1['Column 6-Warehouse name']])
        barcode=np.array(products_list['barcode'])
        Coefficient=np.array(products_list['amount per package'])

   

        company_code_list=company_list_data[name].values
        print("name: ",name)
        print('company_code_list: ',company_code_list)

    a1=[]
    for i in range(0,len(barcode)):
        summ1=0
        for j in range(0,len(data2[3,:])):
            if str(data2[0, j][5:7]) in ['01', '02', '03','04','05','06','07','08','09','10','11','12'] and 3 == int(str(data2[0, j][3])) and barcode[i]==data2[3,j] and data2[8,j]!='انبار مرجوعی' and data2[8,j]!='انبار ضایعات'and data2[7,j] in company_code_list:
                summ1=summ1+ceil(data2[4,j]/Coefficient[i])
            else:
                continue
        a1.append([barcode[i],summ1])
    a1=np.array(a1)
    # print('a1: ',a1)
    a2=[]
    for i in range(0,len(barcode)):
        summ2=0
        for j in range(0,len(data2[2,:])):
            if str(data2[0, j][5:7]) in ['01', '02', '03','04','05','06','07','08','09','10','11','12'] and 4 == int(str(data2[0, j][3])) and barcode[i]==data2[3,j] and data2[8,j]!='انبار مرجوعی' and data2[8,j]!='انبار ضایعات'and data2[7,j] in company_code_list:
                summ2=summ2+ceil(data2[4,j]/Coefficient[i])
            else:
                continue
        a2.append([barcode[i],summ2])
    a2=np.array(a2)
    # print('a2: ',a2)
    
    y1_final=[]
    y2_final=[]
    for cc in range(0,len(a1[:,1])):
        if a1[cc,1]!=0 or a2[cc,1]!=0:
            y1_final.append(a1[cc])
            y2_final.append(a2[cc])
        else:
            continue
        
    y1_final=np.array(y1_final)
    y2_final=np.array(y2_final)
    z1=[]
    for e in y1_final[:,0]:
        for c in range(0,len(products_list['barcode'])):
            if products_list['barcode'][c]==e:
                z1.append(products_list['name'][c])
            else:
                continue
    x1=[get_display(arabic_reshaper.reshape(d)) for d in z1]
    x = np.arange(len(x1))
    plt.rcParams["figure.figsize"] = (25,6)
    fig, ax = plt.subplots()
    width=0.5
    bar1=ax.bar(x - width/2, y1_final[:,1], width=width, label='1402')
    bar2=ax.bar(x + width/2, y2_final[:,1], width=width, label='1403',color='orange')
    ax.set_ylabel(get_display(arabic_reshaper.reshape('تعداد کارتن')), fontsize='30')
    ax.set_title(get_display(arabic_reshaper.reshape(f'مقایسه فروش کارتنی سال‌های 1403 و 1404 به تفکیک کالا-{name}-12 ماهه')), fontsize=35)
    ax.set_xticks(x)
    ax.set_yticks([])
    ax.set_yticklabels([])
    ax.spines[['right', 'top','left']].set_visible(False)
    zz=[get_display(arabic_reshaper.reshape(d)) for d in z1]
    ax.set_xticklabels(zz, rotation=90, fontsize='30')

    plt.rcParams['font.family'] = 'B nazanin'
    ax.legend(fontsize='25')
    plt.savefig(f'D:/razaghi/total/{name}_bar_chart_comparison.png', bbox_inches='tight')
    aa=[]
    for hg in range(0,len(z1)):
        aa.append([z1[hg],y1_final[hg,1],y2_final[hg,1]])
    aa=np.array(aa)
    final_data={
        'product name':list(aa[:,0]),
        '1402':list(aa[:,1]),
        '1403':list(aa[:,2])
    }
    final_data=pd.DataFrame(final_data)
    final_data.to_excel(f'D:/razaghi/total/{name}_products.xlsx', index=False)

    
