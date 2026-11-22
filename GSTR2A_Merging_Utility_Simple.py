from openpyxl import Workbook
from datetime import date
from tkinter import *
import os
import glob
from tkinter import messagebox ,filedialog
import datetime
import warnings
import numpy as np
from UliPlot .XLSX import auto_adjust_xlsx_column_width
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
    df2 =pd .DataFrame ()
    for file in filenames :
        df =pd .read_excel (file ,header =4 ,sheet_name ="B2B")
        df .rename (columns ={'Детали счета':'InvoiceNumber'},inplace =True )
        df .rename (columns ={'Безымянный: 3':'Тип счета'},inplace =True )
        df .rename (columns ={'Безымянный: 4':'Дата счета'},inplace =True )
        df .rename (columns ={'Безымянный: 5':'Стоимость счета'},inplace =True )
        df .rename (columns ={'Сумма налога':'Интегрированный налог (₹)'},inplace =True )
        df .rename (columns ={'Безымянный: 11':'Центральный налог (₹)'},inplace =True )
        df .rename (columns ={'Безымянный: 12':'Налог штата/UT (₹)'},inplace =True )
        df .rename (columns ={'Безымянный: 13':'Цесс (₹)'},inplace =True )
        df =df .drop (0 )
        df =df .dropna (how ='all')
        df =df [~df .InvoiceNumber .str .contains ('Total')]
        df2 =df2 .append (df )
    df3 =pd .DataFrame ()
    for file in filenames :
        df4 =pd .read_excel (file ,header =5 ,sheet_name ="B2BA")
        df4 .rename (columns ={'Детали счета':'Тип счета'},inplace =True )
        df4 .rename (columns ={'Безымянный: 5':'InvoiceNumber2'},inplace =True )
        df4 .rename (columns ={'Безымянный: 6':'Дата счета'},inplace =True )
        df4 .rename (columns ={'Безымянный: 7':'Стоимость счета'},inplace =True )
        df4 .rename (columns ={'Сумма налога':'Интегрированный налог (₹)'},inplace =True )
        df4 .rename (columns ={'Безымянный: 13':'Центральный налог (₹)'},inplace =True )
        df4 .rename (columns ={'Безымянный: 14':'Налог штата/UT (₹)'},inplace =True )
        df4 .rename (columns ={'Безымянный: 15':'Цесс (₹)'},inplace =True )
        df4 =df4 .drop (0 )
        df4 =df4 .dropna (how ='all')
        df4 =df4 [~df4 .InvoiceNumber2 .str .contains ('Total')]
        df3 =df3 .append (df4 )
    df6 =pd .DataFrame ()
    for file in filenames :
        df5 =pd .read_excel (file ,header =4 ,sheet_name ="CDNR")
        df5 .rename (columns ={'Безымянный: 3':'Notenumber'},inplace =True )
        df5 .rename (columns ={'Безымянный: 4':'Примечание Тип поставки'},inplace =True )
        df5 .rename (columns ={'Безымянный: 5':'Дата заметки'},inplace =True )
        df5 .rename (columns ={'Безымянный: 6':'Стоимость примечания (₹)'},inplace =True )
        df5 .rename (columns ={'Сумма налога':'Интегрированный налог (₹)'},inplace =True )
        df5 .rename (columns ={'Безымянный: 12':'Центральный налог (₹)'},inplace =True )
        df5 .rename (columns ={'Безымянный: 13':'Налог штата/UT (₹)'},inplace =True )
        df5 .rename (columns ={'Безымянный: 14':'Цесс (₹)'},inplace =True )
        df5 =df5 .drop (0 )
        df5 =df5 .dropna (how ='all')
        df5 =df5 [~df5 .Notenumber .str .contains ('Total')]
        df6 =df6 .append (df5 )
    df7 =pd .DataFrame ()
    for file in filenames :
        df8 =pd .read_excel (file ,header =5 ,sheet_name ="CDNRA")
        df8 .rename (columns ={'Безымянный: 6':'NoteNumber'},inplace =True )
        df8 .rename (columns ={'Безымянный: 7':'Примечание Тип поставки'},inplace =True )
        df8 .rename (columns ={'Безымянный: 8':'Дата заметки'},inplace =True )
        df8 .rename (columns ={'Безымянный: 9':'Стоимость примечания (₹)'},inplace =True )
        df8 .rename (columns ={'Сумма налога':'Интегрированный налог (₹)'},inplace =True )
        df8 .rename (columns ={'Безымянный: 15':'Центральный налог (₹)'},inplace =True )
        df8 .rename (columns ={'Безымянный: 16':'Налог штата/UT (₹)'},inplace =True )
        df8 .rename (columns ={'Безымянный: 17':'Цесс (₹)'},inplace =True )
        df8 =df8 .drop (0 )
        df8 =df8 .dropna (how ='all')
        df8 =df8 [~df8 .NoteNumber .str .contains ('Total')]
        df7 =df7 .append (df8 )
    q =date .today ()
    filename =f'GSTR2A as on {q }'
    extension =os .path .splitext (filepath )[1 ]
    path =os .path .join (pth ,filename +extension )
    writer =pd .ExcelWriter (path ,engine ='xlsxwriter',options ={'strings_to_formulas':False })
    df2 .to_excel (writer ,sheet_name ="B2B",index =False )
    auto_adjust_xlsx_column_width (df2 ,writer ,sheet_name ="B2B",margin =10 )
    df3 .to_excel (writer ,sheet_name ="B2BA",index =False )
    auto_adjust_xlsx_column_width (df2 ,writer ,sheet_name ="B2BA",margin =10 )
    df6 .to_excel (writer ,sheet_name ="CDNR",index =False )
    auto_adjust_xlsx_column_width (df2 ,writer ,sheet_name ="CDNR",margin =10 )
    df7 .to_excel (writer ,sheet_name ="CDNRA",index =False )
    auto_adjust_xlsx_column_width (df2 ,writer ,sheet_name ="CDNRA",margin =10 )
    writer .save ()
    writer .close ()
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
LogGui .mainloop ()
FotaGui .mainloop ()