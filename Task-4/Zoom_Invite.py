import ClointFusion as cf
from ClointFusion.ClointFusion import key_press
import keyboard
cf.OFF_semi_automatic_mode()
import time

#Updating Hackathon Version and Date in Excel
cf.launch_any_exe_bat_application(r"C:\Users\Karthick\Desktop\Task-4\Email.xlsx")
cf.key_press('alt+f3')
cf.key_write_enter('c2')
for i in range(5):
    cf.key_write_enter('9.0')
cf.key_press('alt+f3')
cf.key_write_enter('d2')
for i in range(5):
    cf.key_write_enter('12th June 2021')
cf.key_press('ctrl+s')
cf.window_close_windows("excel")

#Date and Hackathon versions which are stored in excel file
version = cf.excel_get_single_cell(excel_path=r"C:\Users\Karthick\Desktop\Task-4\Email.xlsx", columnName="Hackathon version", cellNumber=0)
date = cf.excel_get_single_cell(excel_path=r"C:\Users\Karthick\Desktop\Task-4\Email.xlsx", columnName="Date", cellNumber=0)


#Changing the hackathon version and date in the given text file
cf.launch_any_exe_bat_application(r"C:\Users\Karthick\Desktop\Task-4\Email.txt") 
cf.key_press('ctrl+h')
cf.key_write_enter('6.0')
cf.key_press('tab')
cf.key_write_enter('9.0')
cf.key_hit_enter()
for i in range(4):
    cf.key_press('tab')
cf.key_hit_enter()
cf.key_hit_enter()

cf.key_press('ctrl+h')
cf.key_write_enter('3.0')
cf.key_press('tab')
cf.key_write_enter('9.0')
cf.key_hit_enter()
for i in range(4):
    cf.key_press('tab')
cf.key_hit_enter()
cf.key_hit_enter()

cf.key_press('ctrl+h')
cf.key_write_enter('March 13, 2021')
cf.key_press('tab')
cf.key_write_enter(date)
cf.key_hit_enter()
for i in range(4):
    cf.key_press('tab')
cf.key_hit_enter()
cf.key_hit_enter()
cf.key_press('ctrl+s')
cf.window_close_windows("notepad")


#Browser and outlook
cf.browser_navigate_h("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1621775318&rver=7.0.6737.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26RpsCsrfState%3da28024fb-5807-740b-f967-1e49a65bda97&id=292841&aadredir=1&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=90015")

#SIgnin outlook
cf.key_write_enter("karthickps.clointfusion@outlook.com")
cf.key_write_enter("Pskar135#")
time.sleep(5)


row_col=cf.excel_get_row_column_count(r"C:\Users\Karthick\Desktop\Task-4\Email.xlsx")
rows=row_col[0]-1
name='Folks'
for i in range(rows):
    cf.launch_any_exe_bat_application(r"C:\Users\Karthick\Desktop\Task-4\Email.txt") 
    cf.key_press('ctrl+h')
    cf.key_write_enter(name)
    cf.key_press('tab')
    name=cf.excel_get_single_cell(excel_path=r"C:\Users\Karthick\Desktop\Task-4\Email.xlsx", columnName="Name", cellNumber=i)
    cf.key_write_enter(name)
    cf.key_hit_enter()
    for j in range(4):
        cf.key_press('tab')
    cf.key_hit_enter()
    cf.key_hit_enter()
    cf.key_press('esc')
    cf.key_press('ctrl+s')
    cf.key_press('ctrl+a')
    cf.key_press('ctrl+c')
    cf.window_close_windows("notepad")
    cf.browser_mouse_click_h('New Message')
    time.sleep(2)
    emailid=cf.excel_get_single_cell(excel_path=r"C:\Users\Karthick\Desktop\Task-4\Email.xlsx", columnName="Email Id", cellNumber=i)
    cf.key_write_enter(emailid)
    for k in range(3):
        cf.key_press('tab')
    cf.key_write_enter("Task-4 Zoom invite via Outlook")
    cf.key_press('tab')
    cf.key_press('ctrl+v')

