import ClointFusion as cf
import os
import time

from numpy import string_


cf.OFF_semi_automatic_mode()
cf.window_show_desktop()


cf.launch_website_h("https://avinashtechlvr.github.io/ClointFusion-Training-Task-6/")
cf.scrape_save_contents_to_notepad('C:/Users/Karthick/Desktop/Task-5')
cf.browser_quit_h()


cf.launch_any_exe_bat_application('C:/Users/Karthick/Desktop/Task-5/notepad-contents.txt')
cf.key_press('ctrl+a')
cf.key_press('ctrl+c')
cf.window_close_windows('notepad')
cf.excel_create_file("C:/Users/Karthick/Desktop/Task-5","scrap_from_notepad")

cf.launch_any_exe_bat_application('C:/Users/Karthick/Desktop/Task-5/scrap_from_notepad.xlsx')
cf.key_press('ctrl+v')
cf.key_press('ctrl+s')
cf.key_press('alt+f3')
cf.key_write_enter("A1")
for i in range(2):
    cf.key_press("ctrl+-")
    cf.key_press('Down')
    cf.key_press('Down')
    cf.key_hit_enter()
for i in range(5):
    cf.key_press('Down')
    cf.key_press("ctrl+-")
    cf.key_press('Down')
    cf.key_press('Down')
    cf.key_hit_enter()

   
cf.key_press("ctrl+s")
cf.window_close_windows('Excel')
cf.launch_website_h("https://www.xe.com/currencyconverter/convert/?Amount=2000&From=GBP&To=CAD")
for j in range(5):

    From_currency = cf.excel_get_single_cell('C:/Users/Karthick/Desktop/Task-5/scrap_from_notepad.xlsx',columnName="From",cellNumber=j)
    To_currency=cf.excel_get_single_cell('C:/Users/Karthick/Desktop/Task-5/scrap_from_notepad.xlsx',columnName="To",cellNumber=j)
    Amount=str(cf.excel_get_single_cell('C:/Users/Karthick/Desktop/Task-5/scrap_from_notepad.xlsx',columnName="Amount",cellNumber=j))
    # print(From_currency)
    # print(To_currency)
    # print(Amount)
    time.sleep(3)
    cf.mouse_click(350,827,single_double_triple='single')
    cf.browser_mouse_click_h('From')
    cf.key_write_enter(From_currency)
    cf.key_press('Tab')
    cf.key_press('Tab')
    cf.key_write_enter(To_currency)
    cf.browser_mouse_click_h("Amount")
    cf.key_write_enter(Amount)
    cf.key_press('ctrl+a')
    cf.key_write_enter(Amount)
    Amount=cf.browser_locate_element_h('//*[@id="__next"]/div[2]/div[2]/section/div[2]/div/main/form/div[2]/div[1]/p[2]',get_text=True).split()

    Amount[0]=Amount[0].replace(',','')
    print(Amount[0])
    cf.excel_set_single_cell('C:/Users/Karthick/Desktop/Task-5/scrap_from_notepad.xlsx',columnName="Converted",cellNumber=j,setText=Amount[0])
cf.browser_quit_h()
cf.browser_navigate_h("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1621775318&rver=7.0.6737.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26RpsCsrfState%3da28024fb-5807-740b-f967-1e49a65bda97&id=292841&aadredir=1&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=90015")

#SIgnin outlook
cf.key_write_enter("karthickps.clointfusion@outlook.com")
cf.key_write_enter("Pskar135#")
time.sleep(5)
cf.browser_mouse_click_h('New Message')
time.sleep(3)
cf.key_write_enter('shrinidhi.clointfusion@gmail.com')
cf.key_write_enter('fharookshaik.clointfusion@gmail.com')
cf.key_write_enter('avinash.clointfusion@gmail.com')
for k in range(3):
    cf.key_press('Tab')
cf.key_write_enter("Task-5 ")
cf.key_press('Tab')
cf.key_write_enter("Currency Convertor and storing in Excel")
cf.launch_any_exe_bat_application('C:/Users/Karthick/Desktop/Task-5/scrap_from_notepad.xlsx')
cf.key_press('ctrl+a')
cf.key_press('ctrl+c')
cf.window_close_windows('Excel')
cf.key_press('ctrl+v')
cf.browser_mouse_click_h("Attach")
cf.browser_mouse_click_h("Browse this computer")
time.sleep(2)
cf.key_write_enter(r'C:\Users\Karthick\Desktop\Task-5\scrap_from_notepad.xlsx')
