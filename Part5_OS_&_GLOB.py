import os
import glob
ppt =os .getcwd ()
filepath ="G:\OFFICE AUTOMATION FILES\Split Excel Files\Financial Sample.xlsx"
dir =os .path .dirname (filepath )
extension =os .path .splitext (filepath )[1 ]
filename =os .path .splitext (filepath )[0 ]
joincheck =os .path .join (dir ,filename +"_signed"+
extension )
listcheck =os .listdir (dir )
joincheck_2 =os .path .join (dir ,listcheck [0 ])
filesize =os .path .getsize (filepath )
multiplefile =glob .glob (dir +"/*.xlsx")
print ('Команда os.getcwd дает такой вывод '+ppt )
print ("Это полный путь к файлу: "+filepath )
print ("Это имя каталога: "+dir )
print ("Это расширение: "+extension )
print ("Это только путь до имени файла без расширения: "+filename )
print ("3 входа объединяются, образуя полный путь к файлу: "+joincheck )
print (listcheck )
print ('Это присоединение файла через индексацию списка.'+joincheck_2 )
print (filesize )
print (multiplefile )
for f in multiplefile :
    print ("Мы можем повторять вот так..!!")