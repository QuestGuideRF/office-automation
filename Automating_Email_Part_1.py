import smtplib
from getpass import getpass
from email .message import EmailMessage
senders_email =input ("Введите свой адрес электронной почты: ")
senders_password =getpass ("Введите свой пароль ")
receivers_mail =['shubham26032000@gmail.com','shubham2000.03.26@gmail.com','shubham2000.09.26@gmail.com']
sub =['TEST','TEST2','ЗАКЛЮЧИТЕЛЬНОЕ ИСПЫТАНИЕ']
attach_files =['1.jpg','2.jpg','3.jpg']
zipped =zip (receivers_mail ,sub ,attach_files )
for (a ,b ,c )in zipped :
    msg =EmailMessage ()
    with open (c ,'rb')as f :
        file_data =f .read ()
        file_name =f .name
    msg ['From']=senders_email
    msg ['To']=a
    msg ['Subject']=b
    msg .set_content (f'Hello ! How are you??')
    msg .add_attachment (file_data ,maintype ='img',subtype ='imghdr',filename =f .name )
    with smtplib .SMTP_SSL ('smtp.gmail.com','465')as smtp :
        smtp .login (senders_email ,senders_password )
        smtp .send_message (msg )
print ("Все письма отправлены!")