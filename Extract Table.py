import os
def extract_url_table (input_url ,folder_path =os .getcwd ()):
    import pandas as pd
    import datetime
    url =input_url
    table =pd .read_html (url )[0 ]
    time_stamp =datetime .datetime .now ().strftime ('%Y-%m-%d %H-%M-%S')
    new_excel_file =os .path .join (folder_path ,"Excel_Table_Output_"+time_stamp +".xlsx")
    writer =pd .ExcelWriter (new_excel_file ,engine ='openpyxl')
    table .to_excel (writer ,sheet_name ="Output")
    writer .save ()
    print ("Таблица в URL-адресе преобразована в файл Excel и сохранена в..",new_excel_file )
extract_url_table ("https://cbic-gst.gov.in/central-tax-notifications.html")
extract_url_table ("https://trends.builtwith.com/websitelist/Responsive-Tables")