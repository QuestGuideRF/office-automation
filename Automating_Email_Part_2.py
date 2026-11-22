import os
import pandas as pd
import smtplib
from email .message import EmailMessage
from getpass import getpass
sender_email =input ("Введите свой адрес электронной почты  ")
sender_pass =getpass ("Введите свой пароль ")
df =pd .read_excel ("E:\Для электронной почты.xlsx")
receivers_email =df ["EMAIL_ID"].values
sub =("Тестовая почта ")
attach_files =df ["Файлы, которые нужно прикрепить"]
name =df ["NAME"].values
zipped =zip (receivers_email ,attach_files ,name )
for (a ,b ,c )in zipped :
    msg =EmailMessage ()
    files =[(r"C:\Users\SHUBHAM\Desktop\Attachments\{}.pdf".format (b ))]
    for file in files :
        with open (file ,'rb')as f :
            file_data =f .read ()
            file_name =f .name
        msg ['From']=sender_email
        msg ['To']=a
        msg ['Subject']=sub
        msg .set_content (f"hello {c }! I have something for you.")
        msg .add_attachment (file_data ,maintype ='application',subtype ='octet-stream',filename ="{}.pdf".format (b ))
        with smtplib .SMTP_SSL ('smtp.gmail.com',465 )as smtp :
            smtp .login (sender_email ,sender_pass )
            smtp .send_message (msg )
print ("Все письма отправлены!")