import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog ,messagebox
import os
import datetime
import datacompy
import ast
FotaGui =Tk ()
LogGui =Tk ()
FotaGui .geometry ("500x500")
LogGui .geometry ("400x400")
FotaGui .title ("Приложение для сравнения двух файлов Excel")
LogGui .title ("Журнал всех действий")
oldfile =StringVar ()
newfile =StringVar ()
def old_file ():
    global oldfile
    global label_head7
    global now
    now =datetime .datetime .now ()
    if (oldfile ==""):
        oldfile =filedialog .askopenfilename (initialdir =os .getcwd (),
        title ="Выберите файл",filetypes =[("Все файлы","*.*"),("PDF-файлы","*.pdf"),("Текстовые файлы","*.текст"),("Файлы Excel","*.xlsx")])
    else :
        oldfile =filedialog .askopenfilename (initialdir =oldfile ,
        title ="Выберите файл",filetypes =[("Все файлы","*.*"),("PDF-файлы","*.pdf"),("Текстовые файлы","*.текст"),("Файлы Excel","*.xlsx")])
    label_head7 =Label (LogGui ,text ='{n}Файл {fil} выбран.'.format (fil =oldfile ,n =now .strftime ('%y-%m-%d %H:%M:%S')),bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
def new_file ():
    global now
    global newfile
    now =datetime .datetime .now ()
    if (newfile ==""):
        newfile =filedialog .askopenfilename (initialdir =os .getcwd (),
        title ="Выберите файл",filetypes =[("Все файлы","*.*"),("PDF-файлы","*.pdf"),("Текстовые файлы","*.текст"),("Файлы Excel","*.xlsx")])
    else :
        newfile =filedialog .askopenfilename (initialdir =newfile ,
        title ="Выберите файл",filetypes =[("Все файлы","*.*"),("PDF-файлы","*.pdf"),("Текстовые файлы","*.текст"),("Файлы Excel","*.xlsx")])
    label_head7 =Label (LogGui ,text ='{n}Файл {fil} выбран.'.format (fil =newfile ,n =now .strftime ('%y-%m-%d %H:%M:%S')),bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
def read_files ():
    global oldfile
    global newfile
    global df1
    global df2
    global e_1
    global e_2
    global e_3
    global commoncol
    df1 =pd .read_excel (oldfile ,sheet_name =0 )
    df2 =pd .read_excel (newfile ,sheet_name =0 )
    commoncol =np .intersect1d (df2 .columns ,df1 .columns )
    messagebox .showinfo ("Output","Мы прочитали оба файла. Общие столбцы: {}. Количество общих столбцов {}.".format (",".join (commoncol ),len (commoncol )))
    label_2 =Label (FotaGui ,text ="Введите название минимум 1 и максимум трех названий общих столбцов.")
    label_2 .pack ()
    e_1 =Entry (FotaGui ,width =50 ,bg ='blue',fg ='white',borderwidth =4 )
    e_1 .pack ()
    e_2 =Entry (FotaGui ,width =50 ,bg ='blue',fg ='white',borderwidth =4 )
    e_2 .pack ()
    e_3 =Entry (FotaGui ,width =50 ,bg ='blue',fg ='white',borderwidth =4 )
    e_3 .pack ()
    Button_1 =Button (FotaGui ,text ="Compare_Excel",command =compare_files )
    Button_1 .pack ()
def compare_files ():
    global oldfile
    global newfile
    global df1
    global df2
    global commoncol
    global e_1
    global e_2
    global e_3
    str1 =e_1 .get ()
    str2 =e_2 .get ()
    str3 =e_3 .get ()
    listcomcol =[]
    listcomcol .append (str3 )
    listcomcol .append (str2 )
    listcomcol .append (str1 )
    a_set =set (commoncol )
    b_set =set (listcomcol )
    common =list (set (a_set .intersection (b_set )))
    comparevalues =datacompy .Compare (df1 ,df2 ,join_columns =common )
    print (comparevalues .report ())
    df =pd .DataFrame ()
    df .to_csv (os .path .join (os .path .dirname (newfile ),"Comparison_File.csv"))
    outputfile =os .path .join (os .path .dirname (newfile ),"Comparison_File.csv")
    with open (outputfile ,mode ="r+",encoding ='utf-8')as report_file :
        report_file .write (comparevalues .report ())
    messagebox .showinfo ("Output","Отчет о сравнении экспортирован в файл csv по адресу {}.".format (outputfile ))
    label_head12 =Label (FotaGui ,text ="   \n"
    "\n"
    "\n"
    "\n Требуется обратная связь для улучшения Программы."
    "\n На основе отзывов программа может быть улучшена с учетом потребностей конкретных пользователей."
    "\n Отправьте свой отзыв по адресу pranav.tulshyan@gmail.com. ",font ="Раз 10 ")
    label_head12 .pack ()
label_0 =Label (FotaGui ,text ='\n')
label_0 .pack ()
label_0 =Label (FotaGui ,text ='Шаг: 1. Выберите файл, нажав кнопку «Обзор» !!!',font ='раз 11',anchor =N ,bd =1 ,relief ='solid')
label_0 .pack ()
Browsebutton =Button (FotaGui ,width =15 ,text ="BROWSE_Old_File-df1",command =old_file )
Browsebutton .pack ()
Browsebutton =Button (FotaGui ,width =15 ,text ="BROWSE_New_file-df2",command =new_file )
Browsebutton .pack ()
label_head3 =Label (FotaGui ,text ='\n'
'\n'
)
label_head3 .pack ()
label_head4 =Label (FotaGui ,text ='Шаг 2: Нажмите кнопку действия, расположенную ниже:',bd =1 ,relief ='solid',
font ='раз 12',anchor =N )
label_head4 .pack ()
Button_1 =Button (FotaGui ,text ="Read_Excel",command =read_files )
Button_1 .pack ()
LogGui .mainloop ()
FotaGui .mainloop ()