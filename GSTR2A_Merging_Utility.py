import pandas as pd
from tkinter import *
import os
import glob
from tkinter import messagebox ,filedialog
import datetime
import warnings
import numpy as np
FotaGui =Tk ()
LogGui =Tk ()
FotaGui .geometry ("500x500")
LogGui .geometry ("250x250")
FotaGui .title ("Утилита для объединения GSTR2A")
LogGui .title ("Журнал всех действий")
Label_1 =Label (FotaGui ,text ="Это утилита для объединения GSTR 2A",font ="раз 16")
Label_1 .pack ()
warnings .filterwarnings ('ignore')
def file_path ():
    global filepath
    global label_head7
    filepath =StringVar ()
    global now
    now =datetime .datetime .now ()
    if (filepath ==""):
        filepath =filedialog .askopenfilename (initialdir =os .getcwd (),
        title ="Выберите файл",filetypes =[("Все файлы","*.*"),("PDF-файлы","*.pdf"),("Текстовые файлы","*.текст"),("Файлы Excel","*.xlsx")])
    else :
        filepath =filedialog .askopenfilename (initialdir =filepath ,
        title ="Выберите файл",filetypes =[("Все файлы","*.*"),("PDF-файлы","*.pdf"),("Текстовые файлы","*.текст"),("Файлы Excel","*.xlsx")])
    extension =os .path .splitext (filepath )[1 ]
    filename =os .path .splitext (filepath )[0 ]
    pth =os .path .dirname (filepath )
    files =glob .glob (os .path .join (pth ,'*{ext}'.format (ext =extension )))
    for f in files :
        label_head7 =Label (LogGui ,text ='{n}Файл {fil} выбран.'.format (fil =f ,n =now .strftime ('%y-%m-%d %H:%M:%S')),bd =1 ,relief ='solid',
        font ='Раз 10',anchor =N )
        label_head7 .pack ()
def Combine_GSTR2A_File ():
    import pandas as pd
    import glob
    import os
    global filepath
    pth =os .path .dirname (filepath )
    filenames =glob .glob (pth +"/*.xlsx")
    i =0
    for file in filenames :
        i =i +1
    if i <1 :
        print ("Загрузите минимум 2 файла")
    elif i >60 :
        print ("Максимальная емкость — 60 файлов одновременно.")
    else :
        pass
    cum_size =0
    for file in filenames :
        size =os .path .getsize (file )
        cum_size =cum_size +size
        if size >31457280 :
            print ("Пожалуйста, загрузите файл меньшего размера. Максимальный лимит — 30 МБ.")
        elif cum_size >314572800 :
            print ("Общий размер файла для всех файлов составляет более 300 МБ. Пожалуйста, используйте файлы меньшего размера")
            break
        else :
            pass
    df2 =pd .DataFrame ()
    for files in filenames :
        df =pd .read_excel (files ,sheet_name =1 )
        df1 =df .drop ([0 ,1 ,2 ,3 ,4 ],axis =0 )
        df1 =df1 .dropna (how ='all')
        df1 ['File_name']=files
        df2 =df2 .append (df1 )
    df3 =df2
    df3 .rename (columns ={'Налог на товары и услуги – GSTR 2A':'GSTIN_of_Supplier'},inplace =True )
    df3 .rename (columns ={'Безымянный: 1':'Юридическое_название_поставщика'},inplace =True )
    df3 .rename (columns ={'Безымянный: 2':'Inv_CN_DN_Number_Original'},inplace =True )
    df3 .rename (columns ={'Безымянный: 3':'Inv_CN_DN_Type_Original'},inplace =True )
    df3 .rename (columns ={'Безымянный: 4':'Inv_CN_DN_Date_Original'},inplace =True )
    df3 .rename (columns ={'Безымянный: 5':'Inv_CN_DN_Value_Original'},inplace =True )
    df3 .rename (columns ={'Безымянный: 6':'Place_Of_Supply'},inplace =True )
    df3 .rename (columns ={'Безымянный: 7':'Supply_Attract_Reverse_Charge'},inplace =True )
    df3 .rename (columns ={'Безымянный: 8':'GST_Rate'},inplace =True )
    df3 .rename (columns ={'Безымянный: 9':'Taxable_Value_Rs'},inplace =True )
    df3 .rename (columns ={'Безымянный: 10':'IGST_Rs'},inplace =True )
    df3 .rename (columns ={'Безымянный: 11':'CGST_Rs'},inplace =True )
    df3 .rename (columns ={'Безымянный: 12':'SGST_Rs'},inplace =True )
    df3 .rename (columns ={'Безымянный: 13':'Cess'},inplace =True )
    df3 .rename (columns ={'Безымянный: 14':'GSTR_1_5_Filing_Status'},inplace =True )
    df3 .rename (columns ={'Безымянный: 15':'GSTR_1_5_Filing_Date'},inplace =True )
    df3 .rename (columns ={'Безымянный: 16':'GSTR_1_5_Filing_Period'},inplace =True )
    df3 .rename (columns ={'Безымянный: 17':'GSTR_3B_Filing_Status'},inplace =True )
    df3 .rename (columns ={'Безымянный: 18':'Amendment_made_if_any'},inplace =True )
    df3 .rename (columns ={'Безымянный: 19':'Tax_Period_in_which_Amended'},inplace =True )
    df3 .rename (columns ={'Безымянный: 20':'Effective_date_of_cancellation'},inplace =True )
    df3 .rename (columns ={'Безымянный: 21':'Source'},inplace =True )
    df3 .rename (columns ={'Безымянный: 22':'IRN'},inplace =True )
    df3 .rename (columns ={'Безымянный: 23':'IRN_Date'},inplace =True )
    filt =df3 ['Inv_CN_DN_Number_Original'].str .contains ('Total',na =False )
    df3 =df3 [~filt ]
    df3 ['Inv_CN_DN_Date_Unique']=df3 ['Inv_CN_DN_Date_Original'].str .replace ("-",".")
    df3 ['Total_tax']=df3 ['IGST_Rs']+df3 ['CGST_Rs']+df3 ['SGST_Rs']
    df3 ['Unique_ID']=df3 ['GSTIN_of_Supplier']+"/"+df3 ['Inv_CN_DN_Number_Original']+"/"+df3 [
    'Inv_CN_DN_Date_Unique']
    df3 ['Sheet_Name']=("B2B")
    df3 ['PAN_Number']=df3 ["GSTIN_of_Supplier"].apply (lambda x :x [2 :12 :1 ])
    df3 =df3 .replace (np .nan ,"",regex =True )
    label_head21 =Label (FotaGui ,text ='Таблица B2B объединяется... Пожалуйста, подождите')
    df2 =pd .DataFrame ()
    for files in filenames :
        df =pd .read_excel (files ,sheet_name =2 )
        df1 =df .drop ([0 ,1 ,2 ,3 ,4 ,5 ],axis =0 )
        df1 =df1 .dropna (how ='all')
        df1 ['File_name']=files
        df2 =df2 .append (df1 )
    df4 =df2 .dropna (how ='all')
    df4 .rename (
    columns ={'                                      Налог на товары и услуги – GSTR-2A':'Inv_CN_DN_Number_Original'},
    inplace =True )
    df4 .rename (columns ={'Безымянный: 1':'Inv_CN_DN_Date_Original'},inplace =True )
    df4 .rename (columns ={'Безымянный: 2':'GSTIN_of_Supplier'},inplace =True )
    df4 .rename (columns ={'Безымянный: 3':'Юридическое_название_поставщика'},inplace =True )
    df4 .rename (columns ={'Безымянный: 4':'Inv_CN_DN_Type_Revised'},inplace =True )
    df4 .rename (columns ={'Безымянный: 5':'Inv_CN_DN_Number_Revised'},inplace =True )
    df4 .rename (columns ={'Безымянный: 6':'Inv_CN_DN_Date_Revised'},inplace =True )
    df4 .rename (columns ={'Безымянный: 7':'Inv_CN_DN_Value_Revised'},inplace =True )
    df4 .rename (columns ={'Безымянный: 8':'Place_Of_Supply'},inplace =True )
    df4 .rename (columns ={'Безымянный: 9':'Supply_Attract_Reverse_Charge'},inplace =True )
    df4 .rename (columns ={'Безымянный: 10':'GST_Rate'},inplace =True )
    df4 .rename (columns ={'Безымянный: 11':'Taxable_Value_Rs'},inplace =True )
    df4 .rename (columns ={'Безымянный: 12':'IGST_Rs'},inplace =True )
    df4 .rename (columns ={'Безымянный: 13':'CGST_Rs'},inplace =True )
    df4 .rename (columns ={'Безымянный: 14':'SGST_Rs'},inplace =True )
    df4 .rename (columns ={'Безымянный: 15':'Cess'},inplace =True )
    df4 .rename (columns ={'Безымянный: 16':'GSTR_1_5_Filing_Status'},inplace =True )
    df4 .rename (columns ={'Безымянный: 17':'GSTR_1_5_Filing_Date'},inplace =True )
    df4 .rename (columns ={'Безымянный: 18':'GSTR_1_5_Filing_Period'},inplace =True )
    df4 .rename (columns ={'Безымянный: 19':'GSTR_3B_Filing_Status'},inplace =True )
    df4 .rename (columns ={'Безымянный: 20':'Effective_date_of_cancellation'},inplace =True )
    df4 .rename (columns ={'Безымянный: 21':'Amendment_made_if_any'},inplace =True )
    df4 .rename (columns ={'Безымянный: 22':'Original_tax_period_in_which_reported'},inplace =True )
    filt =df4 ['Inv_CN_DN_Number_Revised'].str .contains ('Total',na =False )
    df4 =df4 [~filt ]
    df4 ['Inv_CN_DN_Date_Unique']=df4 ['Inv_CN_DN_Date_Original'].str .replace ("-",".")
    df4 ['Total_tax']=df4 ['IGST_Rs']+df4 ['CGST_Rs']+df4 ['SGST_Rs']
    df4 ['Unique_ID']=df4 ['GSTIN_of_Supplier']+"/"+df4 ['Inv_CN_DN_Number_Original']+"/"+df4 [
    'Inv_CN_DN_Date_Unique']
    df4 ["Inv_CN_DN_Date_Revised_Unique"]=df4 ['Inv_CN_DN_Date_Revised'].str .replace ("-",".")
    df4 ['Sheet_Name']=("B2BA")
    df4 ['PAN_Number']=df4 ["GSTIN_of_Supplier"].apply (lambda x :x [2 :12 :1 ])
    df4 =df4 .replace (np .nan ,"",regex =True )
    df2 =pd .DataFrame ()
    for files in filenames :
        df =pd .read_excel (files ,sheet_name =3 )
        df1 =df .drop ([0 ,1 ,2 ,3 ,4 ],axis =0 )
        df1 =df1 .dropna (how ='all')
        df1 ['File_name']=files
        df2 =df2 .append (df1 )
    df5 =df2 .dropna (how ='all')
    df5 .rename (
    columns ={'                                             Налог на товары и услуги – GSTR-2A':'GSTIN_of_Supplier'},
    inplace =True )
    df5 .rename (columns ={'Безымянный: 1':'Юридическое_название_поставщика'},inplace =True )
    df5 .rename (columns ={'Безымянный: 2':'Credit_Debit_Note_Original'},inplace =True )
    df5 .rename (columns ={'Безымянный: 3':'Inv_CN_DN_Number_Original'},inplace =True )
    df5 .rename (columns ={'Безымянный: 4':'Inv_CN_DN_Type_Original'},inplace =True )
    df5 .rename (columns ={'Безымянный: 5':'Inv_CN_DN_Date_Original'},inplace =True )
    df5 .rename (columns ={'Безымянный: 6':'Inv_CN_DN_Value_Original'},inplace =True )
    df5 .rename (columns ={'Безымянный: 7':'Place_Of_Supply'},inplace =True )
    df5 .rename (columns ={'Безымянный: 8':'Supply_Attract_Reverse_Charge'},inplace =True )
    df5 .rename (columns ={'Безымянный: 9':'GST_Rate'},inplace =True )
    df5 .rename (columns ={'Безымянный: 10':'Taxable_Value_Rs'},inplace =True )
    df5 .rename (columns ={'Безымянный: 11':'IGST_Rs'},inplace =True )
    df5 .rename (columns ={'Безымянный: 12':'CGST_Rs'},inplace =True )
    df5 .rename (columns ={'Безымянный: 13':'SGST_Rs'},inplace =True )
    df5 .rename (columns ={'Безымянный: 14':'Cess'},inplace =True )
    df5 .rename (columns ={'Безымянный: 15':'GSTR_1_5_Filing_Status'},inplace =True )
    df5 .rename (columns ={'Безымянный: 16':'GSTR_1_5_Filing_Date'},inplace =True )
    df5 .rename (columns ={'Безымянный: 17':'GSTR_1_5_Filing_Period'},inplace =True )
    df5 .rename (columns ={'Безымянный: 18':'GSTR_3B_Filing_Status'},inplace =True )
    df5 .rename (columns ={'Безымянный: 19':'Amendment_made_if_any'},inplace =True )
    df5 .rename (columns ={'Безымянный: 20':'Tax_Period_in_which_Amended'},inplace =True )
    df5 .rename (columns ={'Безымянный: 21':'Effective_date_of_cancellation'},inplace =True )
    df5 .rename (columns ={'Безымянный: 22':'Source'},inplace =True )
    df5 .rename (columns ={'Безымянный: 23':'IRN'},inplace =True )
    df5 .rename (columns ={'Безымянный: 24':'IRN_Date'},inplace =True )
    filt =df5 ['Inv_CN_DN_Number_Original'].str .contains ('Total',na =False )
    df5 =df5 [~filt ]
    df5 ['Inv_CN_DN_Date_Unique']=df5 ['Inv_CN_DN_Date_Original'].str .replace ("-",".")
    df5 ['Total_tax']=df5 ['IGST_Rs']+df5 ['CGST_Rs']+df5 ['SGST_Rs']
    df5 ['Unique_ID']=df5 ['GSTIN_of_Supplier']+"/"+df5 ['Inv_CN_DN_Number_Original']+"/"+df5 [
    'Inv_CN_DN_Date_Unique']
    df5 ['Sheet_Name']=("CDNR")
    df5 ['PAN_Number']=df5 ["GSTIN_of_Supplier"].apply (lambda x :x [2 :12 :1 ])
    df5 =df5 .replace (np .nan ,"",regex =True )
    df2 =pd .DataFrame ()
    for files in filenames :
        df =pd .read_excel (files ,sheet_name =4 )
        df1 =df .drop ([0 ,1 ,2 ,3 ,4 ,5 ],axis =0 )
        df1 =df1 .dropna (how ='all')
        df1 ['File_name']=files
        df2 =df2 .append (df1 )
    df6 =df2 .dropna (how ='all')
    df6 .rename (columns ={'                             Налог на товары и услуги – GSTR2A':'Credit_Debit_Note_Original'},
    inplace =True )
    df6 .rename (columns ={'Безымянный: 1':'Inv_CN_DN_Number_Original'},inplace =True )
    df6 .rename (columns ={'Безымянный: 2':'Inv_CN_DN_Date_Original'},inplace =True )
    df6 .rename (columns ={'Безымянный: 3':'GSTIN_of_Supplier'},inplace =True )
    df6 .rename (columns ={'Безымянный: 4':'Юридическое_название_поставщика'},inplace =True )
    df6 .rename (columns ={'Безымянный: 5':'Credit_Debit_Note_Revised'},inplace =True )
    df6 .rename (columns ={'Безымянный: 6':'Inv_CN_DN_Number_Revised'},inplace =True )
    df6 .rename (columns ={'Безымянный: 7':'Inv_CN_DN_Type_Revised'},inplace =True )
    df6 .rename (columns ={'Безымянный: 8':'Inv_CN_DN_Date_Revised'},inplace =True )
    df6 .rename (columns ={'Безымянный: 9':'Inv_CN_DN_Value_Revised'},inplace =True )
    df6 .rename (columns ={'Безымянный: 10':'Place_Of_Supply'},inplace =True )
    df6 .rename (columns ={'Безымянный: 11':'Supply_Attract_Reverse_Charge'},inplace =True )
    df6 .rename (columns ={'Безымянный: 12':'GST_Rate'},inplace =True )
    df6 .rename (columns ={'Безымянный: 13':'Taxable_Value_Rs'},inplace =True )
    df6 .rename (columns ={'Безымянный: 14':'IGST_Rs'},inplace =True )
    df6 .rename (columns ={'Безымянный: 15':'CGST_Rs'},inplace =True )
    df6 .rename (columns ={'Безымянный: 16':'SGST_Rs'},inplace =True )
    df6 .rename (columns ={'Безымянный: 17':'Cess'},inplace =True )
    df6 .rename (columns ={'Безымянный: 18':'GSTR_1_5_Filing_Status'},inplace =True )
    df6 .rename (columns ={'Безымянный: 19':'GSTR_1_5_Filing_Date'},inplace =True )
    df6 .rename (columns ={'Безымянный: 20':'GSTR_1_5_Filing_Period'},inplace =True )
    df6 .rename (columns ={'Безымянный: 21':'GSTR_3B_Filing_Status'},inplace =True )
    df6 .rename (columns ={'Безымянный: 22':'Amendment_made_if_any'},inplace =True )
    df6 .rename (columns ={'Безымянный: 23':'Original_tax_period_in_which_reported'},inplace =True )
    df6 .rename (columns ={'Безымянный: 24':'Effective_date_of_cancellation'},inplace =True )
    filt =df6 ['Inv_CN_DN_Number_Revised'].str .contains ('Total',na =False )
    df6 =df6 [~filt ]
    df6 ['Inv_CN_DN_Date_Unique']=df6 ['Inv_CN_DN_Date_Original'].str .replace ("-",".")
    df6 ['Total_tax']=df6 ['IGST_Rs']+df6 ['CGST_Rs']+df6 ['SGST_Rs']
    df6 ['Unique_ID']=df6 ['GSTIN_of_Supplier']+"/"+df6 ['Inv_CN_DN_Number_Original']+"/"+df6 [
    'Inv_CN_DN_Date_Unique']
    df6 ["Inv_CN_DN_Date_Revised_Unique"]=df6 ['Inv_CN_DN_Date_Revised'].str .replace ("-",".")
    df6 ['Sheet_Name']=("CDNRA")
    df6 ['PAN_Number']=df6 ["GSTIN_of_Supplier"].apply (lambda x :x [2 :12 :1 ])
    df6 =df6 .replace (np .nan ,"",regex =True )
    df8 =df3 .append (df4 )
    df9 =df8 .append (df5 )
    df10 =df9 .append (df6 )
    df10 ['PAN_Number']=df10 ["GSTIN_of_Supplier"].apply (lambda x :x [2 :12 :1 ])
    df10 =df10 .replace (np .nan ,"",regex =True )
    df10 ["Ultimate_Unique"]=df10 ["Sheet_Name"]+"/"+df10 ["Supply_Attract_Reverse_Charge"]+df10 ["GSTR_1_5_Filing_Status"]+"/"+df10 ["Unique_ID"]
    df10 ["PAN_3_Way_Key"]=np .where (df10 ["Sheet_Name"]=="B2BA",
    df10 ["PAN_Number"]+"/"+df10 ["Inv_CN_DN_Number_Revised"]+"/"
    +df10 ["Inv_CN_DN_Date_Revised_Unique"],
    df10 ["PAN_Number"]+"/"+df10 ["Inv_CN_DN_Number_Original"]
    +"/"+df10 ["Inv_CN_DN_Date_Unique"])
    df10 ["PAN_2_Way_Key_PAN_InvNo"]=np .where (df10 ["Sheet_Name"]=="B2BA",
    df10 ["PAN_Number"]+"/"+df10 ["Inv_CN_DN_Number_Revised"]
    ,df10 ["PAN_Number"]+"/"+df10 ["Inv_CN_DN_Number_Original"])
    df10 ["PAN_2_Way_Key_PAN_InvDt"]=np .where (df10 ["Sheet_Name"]=="B2BA",
    df10 ["PAN_Number"]+"/"+df10 ["Inv_CN_DN_Date_Revised_Unique"]
    ,df10 ["PAN_Number"]+"/"+df10 ["Inv_CN_DN_Date_Unique"])
    df11 =df10 [df10 ['GSTR_1_5_Filing_Status']=="N"]
    df12 =df10 [(df10 ['Supply_Attract_Reverse_Charge']=="Y")&(df10 ['GSTR_1_5_Filing_Status']=="Y")]
    df13 =df10 [(df10 ['Supply_Attract_Reverse_Charge']=="N")&(df10 ['GSTR_1_5_Filing_Status']=="Y")&(
    df10 ['Total_tax']<1 )]
    df14 =df10 [(df10 ['Supply_Attract_Reverse_Charge']=="N")&(df10 ['GSTR_1_5_Filing_Status']=="Y")&(
    df10 ['Total_tax']>=1 )]
    extension =os .path .splitext (filepath )[1 ]
    filename =os .path .splitext (filepath )[0 ]
    pth =os .path .dirname (filepath )
    newfile =os .path .join (pth ,filename +'GSTR2A_all_combined'+extension )
    writer =pd .ExcelWriter (newfile ,engine ='openpyxl')
    df3 .to_excel (writer ,sheet_name ="B2B")
    df4 .to_excel (writer ,sheet_name ="B2BA")
    df5 .to_excel (writer ,sheet_name ="CDNR")
    df6 .to_excel (writer ,sheet_name ="CDNRA")
    titles =list (df10 .columns )
    titles [0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [10 ],titles [11 ],titles [12 ],titles [13 ],titles [14 ],titles [15 ],titles [16 ],titles [17 ],titles [18 ],titles [
    19 ],titles [20 ],titles [21 ],titles [22 ],titles [23 ],titles [24 ],titles [25 ],titles [26 ],titles [27 ],titles [28 ],titles [29 ],titles [30 ],titles [31 ],titles [32 ],titles [33 ],titles [34 ],titles [35 ]=titles [24 ],titles [28 ],titles [
    0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [
    10 ],titles [11 ],titles [
    12 ],titles [13 ],titles [
    26 ],titles [25 ],titles [
    21 ],titles [27 ],titles [
    14 ],titles [15 ],titles [
    16 ],titles [17 ],titles [
    18 ],titles [19 ],titles [
    20 ],titles [22 ],titles [
    23 ],titles [29 ],titles [
    30 ],titles [31 ],titles [
    32 ],titles [33 ],titles [
    34 ],titles [35 ]
    df10 [titles ].to_excel (writer ,sheet_name ="All_Combined")
    titles =list (df11 .columns )
    titles [0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [10 ],titles [11 ],titles [12 ],titles [13 ],titles [14 ],titles [15 ],titles [16 ],titles [17 ],titles [18 ],titles [
    19 ],titles [20 ],titles [21 ],titles [22 ],titles [23 ],titles [24 ],titles [25 ],titles [26 ],titles [27 ],titles [28 ],titles [29 ],titles [30 ],titles [31 ],titles [32 ],titles [33 ],titles [34 ],titles [35 ]=titles [24 ],titles [28 ],titles [
    0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [
    10 ],titles [11 ],titles [
    12 ],titles [13 ],titles [
    26 ],titles [25 ],titles [
    21 ],titles [27 ],titles [
    14 ],titles [15 ],titles [
    16 ],titles [17 ],titles [
    18 ],titles [19 ],titles [
    20 ],titles [22 ],titles [
    23 ],titles [29 ],titles [
    30 ],titles [31 ],titles [
    32 ],titles [33 ],titles [
    34 ],titles [35 ]
    df11 [titles ].to_excel (writer ,sheet_name ="GSTR_1_Не сохранено")
    titles =list (df12 .columns )
    titles [0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [10 ],titles [11 ],titles [12 ],titles [13 ],titles [14 ],titles [15 ],titles [16 ],titles [17 ],titles [18 ],titles [
    19 ],titles [20 ],titles [21 ],titles [22 ],titles [23 ],titles [24 ],titles [25 ],titles [26 ],titles [27 ],titles [28 ],titles [29 ],titles [30 ],titles [31 ],titles [32 ],titles [33 ],titles [34 ],titles [35 ]=titles [24 ],titles [28 ],titles [
    0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [
    10 ],titles [11 ],titles [
    12 ],titles [13 ],titles [
    26 ],titles [25 ],titles [
    21 ],titles [27 ],titles [
    14 ],titles [15 ],titles [
    16 ],titles [17 ],titles [
    18 ],titles [19 ],titles [
    20 ],titles [22 ],titles [
    23 ],titles [29 ],titles [
    30 ],titles [31 ],titles [
    32 ],titles [33 ],titles [
    34 ],titles [35 ]
    df12 [titles ].to_excel (writer ,sheet_name ="GSTR_Filed_RCM_Yes")
    titles =list (df13 .columns )
    titles [0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [10 ],titles [11 ],titles [12 ],titles [13 ],titles [14 ],titles [15 ],titles [16 ],titles [17 ],titles [18 ],titles [
    19 ],titles [20 ],titles [21 ],titles [22 ],titles [23 ],titles [24 ],titles [25 ],titles [26 ],titles [27 ],titles [28 ],titles [29 ],titles [30 ],titles [31 ],titles [32 ],titles [33 ],titles [34 ],titles [35 ]=titles [24 ],titles [28 ],titles [
    0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [
    10 ],titles [11 ],titles [
    12 ],titles [13 ],titles [
    26 ],titles [25 ],titles [
    21 ],titles [27 ],titles [
    14 ],titles [15 ],titles [
    16 ],titles [17 ],titles [
    18 ],titles [19 ],titles [
    20 ],titles [22 ],titles [
    23 ],titles [29 ],titles [
    30 ],titles [31 ],titles [
    32 ],titles [33 ],titles [
    34 ],titles [35 ]
    df13 [titles ].to_excel (writer ,sheet_name ="Tax_Zero_Cases")
    titles =list (df14 .columns )
    titles [0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [10 ],titles [11 ],titles [12 ],titles [13 ],titles [14 ],titles [15 ],titles [16 ],titles [17 ],titles [18 ],titles [
    19 ],titles [20 ],titles [21 ],titles [22 ],titles [23 ],titles [24 ],titles [25 ],titles [26 ],titles [27 ],titles [28 ],titles [29 ],titles [30 ],titles [31 ],titles [32 ],titles [33 ],titles [34 ],titles [35 ]=titles [24 ],titles [28 ],titles [
    0 ],titles [1 ],titles [2 ],titles [3 ],titles [4 ],titles [5 ],titles [6 ],titles [7 ],titles [8 ],titles [9 ],titles [
    10 ],titles [11 ],titles [
    12 ],titles [13 ],titles [
    26 ],titles [25 ],titles [
    21 ],titles [27 ],titles [
    14 ],titles [15 ],titles [
    16 ],titles [17 ],titles [
    18 ],titles [19 ],titles [
    20 ],titles [22 ],titles [
    23 ],titles [29 ],titles [
    30 ],titles [31 ],titles [
    32 ],titles [33 ],titles [
    34 ],titles [35 ]
    df14 [titles ].to_excel (writer ,sheet_name ="Working_Cases")
    writer .save ()
    messagebox .showinfo ('Output','Все файлы GSTR2A объединены!. \n Нажмите ОК')
label_0 =Label (FotaGui ,text ='\n')
label_0 .pack ()
label_0 =Label (FotaGui ,text ='Шаг: 1. Выберите файл, нажав кнопку «Обзор» !!!',font ='раз 11',anchor =N ,bd =1 ,relief ='solid')
label_0 .pack ()
Browsebutton =Button (FotaGui ,width =15 ,text ="BROWSE",command =file_path )
Browsebutton .pack ()
label_head3 =Label (FotaGui ,text ='\n'
'\n'
)
label_head3 .pack ()
label_head4 =Label (FotaGui ,text ='Шаг 2: Нажмите кнопку действия, расположенную ниже:',bd =1 ,relief ='solid',
font ='раз 12',anchor =N )
label_head4 .pack ()
Button_1 =Button (FotaGui ,text ="Объединить файлы GSTR2A",command =Combine_GSTR2A_File )
Button_1 .pack ()
label_head12 =Label (FotaGui ,text ="   \n"
"\n"
"\n"
"\n Требуется обратная связь для улучшения Программы."
"\n На основе отзывов программа может быть улучшена с учетом потребностей конкретных пользователей."
"\n Отправьте свой отзыв по адресу pranav.tulshyan@gmail.com. ",font ="Раз 10 ")
label_head12 .pack ()
LogGui .mainloop ()
FotaGui .mainloop ()