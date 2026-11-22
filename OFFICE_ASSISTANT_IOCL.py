from tkinter import *
from tkinter import filedialog ,messagebox ,ttk
import pandas as pd
import os
import io
import glob
from openpyxl import load_workbook
from shutil import copyfile
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileReader
from pdfminer .pdfinterp import PDFResourceManager ,PDFPageInterpreter
from pdfminer .converter import TextConverter
from pdfminer .layout import LAParams
from pdfminer .pdfpage import PDFPage
import sys ,getopt
import datetime
import tabula as tb
import tabulate
FotaGui =Tk ()
LogGui =Tk ()
FotaGui .geometry ('500x600')
FotaGui .title ('Офис-ассистент для корпоративных клиентов  ')
LogGui .geometry ('800x400')
LogGui .title ('Журнал всех действий:')
filepath =''
label_head0 =Label (FotaGui ,text ="   Офисный помощник для корпоративных пользователей:"
"\n Программа для автоматизации различных офисных функций"
"\n",font ="Раз 10 ")
label_head0 .pack ()
label_head0 =Label (FotaGui ,text ="\n "
"\n Нажмите HELP_INFO для получения дополнительной информации.",font ="Раз 10 ",anchor =W )
label_head0 .pack ()
def HELP_INFO ():
    root =Tk ()
    label_head0 =Label (root ,text ="   Программа для офисных функций",font ="Раз 10 ")
    label_head0 .pack ()
    label_head1 =Label (root ,text ='Это программа для выполнения различных задач в файлах PDF, текстовых файлах или Excel за 2 шага.'
    ' \n   .',bd =1 ,relief ='solid',font ='раз 12',anchor =N )
    label_head1 .pack ()
    label_head2 =Label (root ,text ='Шаг 1: Нажмите кнопку «Обзор», чтобы выбрать файл, для которого вы хотите выполнить действие.'
    '\n Кнопку «Очистить память» можно использовать для очистки выбранных файлов и повторного выбора файла.',
    bd =1 ,relief ='solid',font ='раз 12',anchor =N )
    label_head2 .pack ()
    label_head3 =Label (root ,text ='Шаг 2: Нажмите кнопку действия, расположенную ниже:',bd =1 ,relief ='solid',
    font ='раз 12',anchor =N )
    label_head3 .pack ()
    label_head4 =Label (root ,
    text ='\n «РАЗДЕЛИТЬ ФАЙЛЫ» Если вы хотите разделить файлы, вам необходимо указать:'
    '\n Точное имя столбца, на основе которого вы хотите разделить файл.'
    '\n Файлы можно разделить на несколько листов в одном файле или в разных файлах.'
    ,
    bd =1 ,relief ='solid',font ='раз 12',anchor =N
    )
    label_head4 .pack ()
    label_head5 =Label (root ,text ='ПРИМЕЧАНИЕ 2. Для объединения файлов Excel/pdf/текста обратите внимание ниже:'
    '\n Все объединяемые файлы Excel, PDF или текстовые файлы следует хранить в одной папке.'
    '\n При нажатии кнопки «Обзор» невозможно выбрать несколько файлов.'
    '\n Итак, в этой папке нужно выбрать только один файл.'
    '\n Программа автоматически прочитает все остальные файлы PDF/Excel/текстовые файлы и объединит их.'
    '\n «ОБЪЕДИНИТЬ НОРМАЛЬНЫЕ ФАЙЛЫ EXCEL»: в случае объединения Excel есть два варианта. '
    '\n 1. Объединить файлы: при этом будет объединен ПЕРВЫЙ лист всех файлов Excel.'
    '\n 2. Объединить листы: в случае, если пользователю необходимо выбрать 1 файл Excel, затем все листы этого файла Excel будут объединены. '
    '\n'
    '\n "Combine GSTR2A file": It is to be used specificaly for the GSTR 2A downloaded from GST Site (For Tax People Specifically)'
    '\n'
    '\n «Объединить PDF-файлы»: PDF-файлы будут объединены в той последовательности, в которой они хранятся в папке.'
    '\n'
    '\n «Объединить файлы SAP TXT»: это позволит объединить различные текстовые файлы. Убедитесь, что заголовки строк одинаковы для всех текстовых файлов.    '
    ,bd =1 ,relief ='solid',font ='раз 12',anchor =N
    )
    label_head5 .pack ()
    label_head6 =Label (root ,text ='По любым другим вопросам/проблемам/отзывам: обращайтесь'
    '\nПранав П. Тулшян'
    '\n Электронная почта= pranav.tulshyan@gmail.com'
    '\n Mob: 920*******'
    '\n'
    '\n Нам нужны отзывы об этой программе, чтобы ее можно было улучшить и расширить для удовлетворения потребностей большего числа пользователей.',bd =1 ,relief ='solid',
    font ='раз 11',anchor =N )
    label_head6 .pack ()
    root .title ('Помощь/Информация о программе')
    root .mainloop ()
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
def sendtofile (colslist ,filepath ):
    df =pd .read_excel (filepath )
    cols =e_1 .get ()
    pth =os .path .dirname (filepath )
    colslist =list (set (df [cols ].values ))
    global now
    for i in colslist :
        df [df [cols ]==i ].to_excel ("{}/{}.xlsx".format (pth ,i ),sheet_name =i ,index =False )
    messagebox .showinfo ('Output','You data has been split into {} and {} files has been created.Click OK. \n All Files stored in same folder'.format (
    ', '.join (colslist ),len (colslist )))
    label_head7 =Label (LogGui ,
    text ='{n}Файлы были разделены на разные файлы.'.format (n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
    print ('\nCompleted')
    print ('Спасибо за использование этой программы.')
    return
def sendtosheet (colslist ):
    cols =e_1 .get ()
    extension =os .path .splitext (filepath )[1 ]
    filename =os .path .splitext (filepath )[0 ]
    pth =os .path .dirname (filepath )
    newfile =os .path .join (pth ,filename +'_Sheet_Split_Auto'+extension )
    df =pd .read_excel (filepath )
    colslist =list (set (df [cols ].values ))
    copyfile (filepath ,newfile )
    for j in colslist :
        writer =pd .ExcelWriter (newfile ,engine ='openpyxl')
        for myname in colslist :
            mydf =df .loc [df [cols ]==myname ]
            mydf .to_excel (writer ,sheet_name =myname ,index =False )
        writer .save ()
    messagebox .showinfo ('Output',
    'You data has been split into {} and {} sheets has been created under single file named {new}.\n Click on OK .'.format (
    ', '.join (colslist ),len (colslist ),new =newfile ))
    label_head7 =Label (LogGui ,
    text ='{n}Файлы были разделены на разные листы.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
    print ('\nCompleted')
    print ('Спасибо за использование этой программы.')
    return
def SPLIT_FILE ():
    global filepath
    global e_1
    global e_2
    splitwin =Tk ()
    label_1 =Label (splitwin ,text ='Введите точное имя столбца, значение которого вы хотите разделить.')
    label_1 .pack ()
    e_1 =Entry (splitwin ,width =50 ,bg ='blue',fg ='white',borderwidth =4 )
    e_1 .pack ()
    Browsebutton =Button (splitwin ,width =20 ,text ="Разделить файлы",command =SPLIT_FILE2 )
    Browsebutton .pack ()
    splitwin .mainloop ()
def SPLIT_FILE2 ():
    df =pd .read_excel (filepath )
    cols =e_1 .get ()
    colslist =list (set (df [cols ].values ))
    messagebox .showinfo ('Проверьте вывод',
    'You data will split based on these values {} and create {} files or sheets based on next selection. If you are ready to proceed Click OK or close the dialog box to re-start.'.format (
    ', '.join (colslist ),len (colslist )))
    response =messagebox .askyesno ('Разделить файлы','Хотите ли вы разделить несколько листов в одном файле ИЛИ в разных файлах? '
    '\nНажмите «Да» для разных листов в одном файле.!'
    '\n НАЖМИТЕ «Нет» для разных файлов')
    df =pd .read_excel (filepath )
    cols =e_1 .get ()
    colslist =list (set (df [cols ].values ))
    if response ==0 :
        sendtofile (colslist ,filepath )
    elif response ==1 :
        sendtosheet (colslist )
    else :
        messagebox .showerror ('Output',"Что-то пошло не так")
def Combine_File ():
    global filepath
    global sheet_pos
    combwin =Tk ()
    combwin .title ('Объединение файлов Excel')
    label_comb1 =Label (combwin ,text ='Примечание. Существует два варианта объединения файла Excel.'
    '\n '
    '\n Вариант 1: Объединить файлы --> Это объединит первый лист всех файлов Excel.'
    '\n Вариант 2: объединить листы --> Это объединит все листы этого конкретного файла Excel.'
    '\n'
    '\n While choosing option 2, please ensure that u have only that 1 excel file (whose sheets are to be combined) in that folder.')
    label_comb1 .pack ()
    label_1 =Label (combwin ,text ='\n'
    '\n')
    label_1 .pack ()
    Browsebutton =Button (combwin ,width =15 ,text ="Вариант 1: объединить файлы",command =Combine_File2 )
    Browsebutton .pack ()
    Browsebutton =Button (combwin ,width =15 ,text ="Вариант 2: Объединение листов",command =Combine_File3 )
    Browsebutton .pack ()
    label_1 =Label (combwin ,text ='\n'
    '\n')
    label_1 .pack ()
    combwin .mainloop ()
def Combine_File2 ():
    global filepath
    global sheet_pos
    pth =os .path .dirname (filepath )
    extension =os .path .splitext (filepath )[1 ]
    files =glob .glob (os .path .join (pth ,'*.xls*'))
    newfile =os .path .join (pth ,'All_Files_Combined_Auto.xlsx')
    df =pd .DataFrame ()
    response =messagebox .askyesno ('Важная проверка','Необходимо убедиться, что заголовок строки для всех объединяемых файлов Excel '
    'абсолютно одинаковы.\n Любая разница в заголовках строк может привести к неправильному объединению Excel.!'
    '\n Хотите продолжить?')
    if response ==1 :
        for f in files :
            data =pd .read_excel (f ,sheet_name =0 )
            data ["File_Name"]=f
            df =df .append (data )
        df .to_excel (newfile ,sheet_name ='combined',index =False )
        messagebox .showinfo ('Output','Все файлы Excel в выбранной папке объединены.\n Нажмите «ОК».')
        label_head7 =Label (LogGui ,
        text ='{n}Файлы Excel были объединены.'.format (
        n =now .strftime ('%y-%m-%d %H:%M:%S')),
        bd =1 ,relief ='solid',
        font ='Раз 10',anchor =N )
        label_head7 .pack ()
    else :
        messagebox .showinfo ('Output','Операция прекращена. Пожалуйста, вернитесь после выравнивания названий заголовков строк. \n Нажмите ОК')
        label_head7 =Label (LogGui ,
        text ='{n}Операция была прекращена.'.format (
        n =now .strftime ('%y-%m-%d %H:%M:%S')),
        bd =1 ,relief ='solid',
        font ='Раз 10',anchor =N )
        label_head7 .pack ()
def Combine_File3 ():
    global filepath
    global sheet_pos
    pth =os .path .dirname (filepath )
    df =pd .DataFrame ()
    df2 =pd .DataFrame ()
    xl =pd .ExcelFile (filepath )
    newfile =os .path .join (pth ,'All_Sheets_Combined_Auto.xlsx')
    response =messagebox .askyesno ('Важная проверка','Необходимо убедиться, что заголовок строки для всех объединяемых листов Excel '
    'абсолютно одинаковы.\n Любая разница в заголовках строк может привести к неправильному объединению Excel.!'
    '\n Хотите продолжить?')
    if response ==1 :
        res =len (xl .sheet_names )
        while res >0 :
            res -=1
            df =pd .read_excel (filepath ,sheet_name =res )
            df2 =df2 .append (df )
        df2 .to_excel (newfile ,sheet_name ='combined',index =False )
        messagebox .showinfo ('Output','Все листы Excel в выбранном файле Excel объединены. \n Нажмите ОК')
        label_head7 =Label (LogGui ,
        text ='{n}Листы Excel были объединены.'.format (
        n =now .strftime ('%y-%m-%d %H:%M:%S')),
        bd =1 ,relief ='solid',
        font ='Раз 10',anchor =N )
        label_head7 .pack ()
    else :
        messagebox .showinfo ('Output','Операция прекращена. Пожалуйста, вернитесь после изменения названий заголовков строк.\n Нажмите «ОК».')
        label_head7 =Label (LogGui ,
        text ='{n}Операция была прекращена.'.format (
        n =now .strftime ('%y-%m-%d %H:%M:%S')),
        bd =1 ,relief ='solid',
        font ='Раз 10',anchor =N )
        label_head7 .pack ()
def Combine_PDF ():
    import glob
    import os
    global filepath
    global label_head7
    global now
    pth =os .path .dirname (filepath )
    filenames =glob .glob (pth +"/*.pdf")
    merged =PdfFileMerger ()
    for files in filenames :
        merged .append (files )
    extension =os .path .splitext (filepath )[1 ]
    filename =os .path .splitext (filepath )[0 ]
    pth =os .path .dirname (filepath )
    newfile =os .path .join (pth ,'Combined_Pdf_File_Auto'+extension )
    merged .write (newfile )
    merged .close ()
    messagebox .showinfo ('Output','Все PDF-файлы в выбранной папке были объединены.\n Нажмите «ОК».')
    label_head7 =Label (LogGui ,
    text ='{n}Файлы PDF объединены.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
def Combine_SAP_Txt_files ():
    import pandas as pd
    import glob
    global filepath
    import os
    pth =os .path .dirname (filepath )
    extension =os .path .splitext (filepath )[1 ]
    newfile =os .path .join (pth ,'Combined_Text_File_Auto.txt')
    filenames =glob .glob (pth +"/*.текст")
    df2 =pd .DataFrame ()
    for files in filenames :
        df =pd .read_csv (files ,sep ="\t",low_memory =False ,encoding ='cp1252')
        df2 =df2 .append (df )
    df2 .to_csv (newfile ,sep ="\t")
    messagebox .showinfo ('Output','Все текстовые файлы в выбранной папке были объединены.\n Нажмите «ОК».')
    label_head7 =Label (LogGui ,
    text ='{n}Текстовые файлы SAP объединены.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
    print ('Все текстовые файлы SAP в выбранной папке были объединены.')
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
    label_head21 =Label (FotaGui ,text ='Таблица B2B объединяется... Пожалуйста, подождите')
    label_head7 =Label (LogGui ,
    text ='{n}Выполняется операция..... Пожалуйста, подождите. Лист B2B Читается.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
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
    df4 ['Sheet_Name']=("B2BA")
    label_head7 =Label (LogGui ,
    text ='{n}Выполняется операция..... Пожалуйста, подождите. Лист B2BA Читается.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
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
    label_head7 =Label (LogGui ,
    text ='{n}Выполняется операция..... Пожалуйста, подождите. Лист CDNR читается.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
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
    df6 ['Sheet_Name']=("CDNRA")
    label_head7 =Label (LogGui ,
    text ='{n}Выполняется операция..... Пожалуйста, подождите. Лист CDNRA читается.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
    df8 =df3 .append (df4 )
    df9 =df8 .append (df5 )
    df10 =df9 .append (df6 )
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
    label_head7 =Label (LogGui ,
    text ='{n}Файлы GSTR 2A объединены.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
def Extract_Text ():
    global filepath
    global label_head7
    global now
    pdfDir =os .path .dirname (filepath )
    txtDir =os .path .dirname (filepath )
    if pdfDir =="":pdfDir =os .getcwd ()+"\\"
    for pdf in os .listdir (pdfDir ):
        fileExtension =pdf .split (".")[-1 ]
        if fileExtension =="pdf":
            pdfFilename =pdfDir +"\\"+pdf
            text =convert (pdfFilename )
            textFilename =txtDir +"\\"+pdf +".txt"
            textFile =open (textFilename ,"w")
            textFile .write (text )
        textFile .close
    messagebox .showinfo ('Output','Все PDF-файлы были преобразованы в текстовый файл.\n Нажмите «ОК».')
    label_head7 =Label (LogGui ,
    text ='{n}Файлы PDF преобразованы в текст.'.format (
    n =now .strftime ('%y-%m-%d %H:%M:%S')),
    bd =1 ,relief ='solid',
    font ='Раз 10',anchor =N )
    label_head7 .pack ()
def convert (fname ,pages =None ):
    if not pages :
        pagenums =set ()
    else :
        pagenums =set (pages )
    output =io .StringIO ()
    manager =PDFResourceManager ()
    converter =TextConverter (manager ,output ,laparams =LAParams ())
    interpreter =PDFPageInterpreter (manager ,converter )
    infile =open (fname ,'rb')
    for page in PDFPage .get_pages (infile ,pagenums ):
        interpreter .process_page (page )
    infile .close ()
    converter .close ()
    text =output .getvalue ()
    output .close
    return text
def Extract_Table ():
    global filepath
    global label_head7
    global now
    pth =os .path .dirname (filepath )
    df =pd .DataFrame ()
    tables =tb .read_pdf (filepath ,pages ="all",multiple_tables =True )
    for table in tables :
        i =0
        excel_file =table .to_excel ('pdf+convert.xlsx',sheet_name ="Table_"+i )
        i +=1
def Clear_Memory ():
    messagebox .showinfo ('Очистка памяти','Выбранный файл удален из памяти.')
    now =datetime .datetime .now ()
    label_head12 =Label (LogGui ,text ='{n}: выбранный файл удален из памяти. Вы можете просмотреть файл еще раз '.format (n =now .strftime ("%y-%m-%d %H:%M:%S")))
    label_head12 .pack ()
Browsebutton =Button (FotaGui ,width =15 ,text ="HELP_INFO",command =HELP_INFO )
Browsebutton .pack ()
label_0 =Label (FotaGui ,text ='\n'
'\n Шаг: 1. Выберите файл, нажав кнопку «Обзор» !!!',font ='раз 11',anchor =N )
label_0 .pack ()
Browsebutton =Button (FotaGui ,width =15 ,text ="BROWSE",command =file_path )
Browsebutton .pack ()
Browsebutton =Button (FotaGui ,width =20 ,text ="Очистить память",command =Clear_Memory )
Browsebutton .pack ()
label_20 =Label (FotaGui ,text ='\n'
'\n Шаг: 2 Нажмите на действие, которое вы хотите выполнить !!!'
'\n',font ='раз 11',anchor =N )
label_20 .pack ()
Browsebutton =Button (FotaGui ,width =20 ,text ="Разделить файлы Excel",command =SPLIT_FILE )
Browsebutton .pack ()
Browsebutton =Button (FotaGui ,width =20 ,text ="Объедините обычные файлы Excel",command =Combine_File )
Browsebutton .pack ()
Browsebutton =Button (FotaGui ,width =20 ,text ="Объединение PDF-файлов",command =Combine_PDF )
Browsebutton .pack ()
Browsebutton =Button (FotaGui ,width =20 ,text ="Объединить текстовые файлы",command =Combine_SAP_Txt_files )
Browsebutton .pack ()
Browsebutton =Button (FotaGui ,width =20 ,text ="Объединение файлов GSTR2A",command =Combine_GSTR2A_File )
Browsebutton .pack ()
Browsebutton =Button (FotaGui ,width =20 ,text ="Конвертировать PDF в текст",command =Extract_Text )
Browsebutton .pack ()
Browsebutton =Button (FotaGui ,width =20 ,text ="Извлечь таблицу в Excel",command =Extract_Table )
Browsebutton .pack ()
label_head11 =Label (LogGui ,text ='Журнал всех действий:',anchor =W )
label_head11 .pack ()
label_head12 =Label (FotaGui ,text ="   \n"
"\n"
"\n"
"\n Требуется обратная связь для улучшения Программы."
"\n На основе отзывов программа может быть улучшена с учетом потребностей конкретных пользователей."
"\n Отправьте свой отзыв по адресу pranav.tulshyan@gmail.com. ",font ="Раз 10 ")
label_head12 .pack ()
LogGui .mainloop ()
FotaGui .mainloop ()