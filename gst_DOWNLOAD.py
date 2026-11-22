from selenium import webdriver
from selenium .webdriver .common .keys import Keys
import time
from selenium .webdriver .common .by import By
import os
path ="C:\Program Files (x86)\chromedriver.exe"
driver =webdriver .Chrome (path )
driver .get ("https://www.gst.gov.in/")
print (driver .title )
time .sleep (5 )
driver .find_element_by_xpath ("/html/body/div[1]/header/div[2]/div/div/ul/li/a").click ()
time .sleep (5 )
ele_user =driver .find_element_by_id ("username")
ele_userpass =driver .find_element_by_id ("user_pass")
print (ele_userpass .is_displayed ())
print (ele_userpass .is_enabled ())
user =os .environ .get ("gst_bihar_normal_user")
passwd =os .environ .get ("gst_bihar_normal_pass")
ele_user .send_keys (user )
ele_userpass .send_keys (passwd )
time .sleep (5 )
driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div/div/div/div/div/form/div[6]/div/button").click ()
time .sleep (5 )
print ("Нажата кнопка входа")
year =["2020"]
quarter =["q","qq","qqq","qqqq"]
months =["Apr","May","Jun"]
driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[3]/div/div[1]/button/span").click ()
time .sleep (7 )
for each in quarter :
    for mon in months :
        time .sleep (5 )
        driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[1]/select").click ()
        fin_year =driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[1]/select")
        fin_year .send_keys (year )
        time .sleep (2 )
        driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[2]/select").click ()
        qtr =driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[2]/select")
        qtr .send_keys (each )
        driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[3]/select").click ()
        month =driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[3]/select")
        month .send_keys (mon )
        search_key =driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[4]/кнопка")
        search_key .send_keys (Keys .RETURN )
        print ("Вошёл в панель возврата")
        time .sleep (6 )
        driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[4]/div[3]/div[2]/div[1]/div/div/div/div[1]/button").click ()
        print ("Нажал кнопку просмотра GSTR-3B.")
        time .sleep (6 )
        try :
            download_but =WebDriverWait (driver ,5 ).until (
            EC .element_to_be_clickable ((By .XPATH ,'/html/body/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[9]/div/button[2]'))).click ()
        except Exception as e :
            print (e )
        print ("Нажал кнопку загрузки GSTR-3B.")
        driver .find_element_by_xpath ("/html/body/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[9]/div/a").click ()
        time .sleep (5 )
        print ("Нажал кнопку «Баак»")
time .sleep (5 )
driver .close ()