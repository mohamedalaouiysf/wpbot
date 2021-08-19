import pywhatkit as wp
import os
import xlrd
from datetime import datetime
import time
loc = input("destination : ")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
msg=input("wrtie ur message :")
for i in range(sheet.nrows):
    phone = "+212"+str(int(sheet.cell_value(i, 0)))
    now = datetime.now()
    current_time_h = int(now.strftime("%H"))
    current_time_m = int(now.strftime("%M"))
    waitTime=10
    try:
        if(current_time_m != 59):
            current_time_m=current_time_m+1
            wp.sendwhatmsg(phone,msg,current_time_h,current_time_m,waitTime,True,True)
        else:
            if(current_time_h != 23):
                current_time_h=current_time_h+1
                wp.sendwhatmsg(phone,msg,current_time_h,current_time_m,waitTime,True,True)
            else:
                current_time_h=0
                current_time_m=0
                wp.sendwhatmsg(phone,msg,current_time_h,current_time_m,waitTime,True,True)
    except CallTimeException:
        wp.sendwhatmsg(phone,msg,current_time_h,current_time_m,30,True,True)
    time.sleep(10)
os.system("pause")


