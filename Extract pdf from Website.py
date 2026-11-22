import os
def extract_url_pdf (input_url ,folder_path =os .getcwd ()):
    import os
    import requests
    from urllib .parse import urljoin
    from bs4 import BeautifulSoup
    import pandas as pd
    import datetime
    url =input_url
    folder_location =folder_path
    if not os .path .exists (folder_location ):os .mkdir (folder_location )
    response =requests .get (url )
    soup =BeautifulSoup (response .text ,"html.parser")
    link_text =list ()
    link_href =list ()
    link_file =list ()
    counter =0
    for link in soup .select ("а[href$='.pdf']"):
        filename =os .path .join (folder_location ,link ['href'].split ('/')[-1 ])
        with open (filename ,'wb')as f :
            f .write (requests .get (urljoin (url ,link ['href'])).content )
        link_text .append (str (link .text ))
        link_href .append (link ['href'])
        link_file .append (link ['href'].split ('/')[-1 ])
        counter +=1
        print (counter ,"-Файлы, извлеченные из URL-адреса с именем ",link ['href'].split ('/')[-1 ])
    table_dict ={"Text":link_text ,"Url_Link":link_href ,"Имя файла":link_file }
    df =pd .DataFrame (table_dict )
    time_stamp =datetime .datetime .now ().strftime ('%Y-%m-%d %H-%M-%S')
    print ("Создание файла Excel с именем файла, URL-ссылкой и текстом ссылки...")
    new_excel_file =os .path .join (folder_location ,"Excel_Output_"+time_stamp +".xlsx")
    writer =pd .ExcelWriter (new_excel_file ,engine ='openpyxl')
    df .to_excel (writer ,sheet_name ="Output")
    writer .save ()
    print ("Все файлы PDF загружены и файл Excel создан.")
extract_url_pdf (input_url ="https://www.icai.org/post.html?post_id=17843")
extract_url_pdf (input_url ="https://cbic-gst.gov.in/central-tax-notifications.html")