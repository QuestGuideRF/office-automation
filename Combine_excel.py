import pandas as pd
filepath ="G:\OFFICE AUTOMATION FILES\Объединить файлы Excel\Несколько листов\Несколько листов.xlsx"
xl =pd .ExcelFile (filepath )
listt =xl .sheet_names
df2 =pd .DataFrame ()
for l in listt :
    df1 =pd .read_excel (filepath ,sheet_name =l )
    df2 =df2 .append (df1 )
df2 .to_excel ("G:\OFFICE AUTOMATION FILES\Объединить файлы Excel\Несколько листов\Объединенные листы.xlsx",sheet_name ="Combined",index =False )