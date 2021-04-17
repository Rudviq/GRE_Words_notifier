#Code written by:Rudviq Bhavsar
#Final testing done on:17/04/2021

#Notifier to display Advanced GRE words .
import random
from plyer import notification
import openpyxl as op
import time


def notifyMe(title,message):
    notification.notify(
        title=title,
        message= message,
        app_icon="C:\\Users\\Rudviq\\PycharmProjects\\GRE_Notifier\\Advance.ico",
        timeout=4,
    )

if __name__=='__main__':
    # while True:
    #     notifyMe("Hey Rud, This is your first notifier","Reminding you to upgrade this App for Vocab reminder!.")
    #     time.sleep(60*60)
    while True:
        wb=op.load_workbook("GREWords.xlsx")
        # print(wb.sheetnames)
        # print(wb.active.title)
        sh1=wb['Advance Words']
        # print(sh1['B2'].value)
        n=random.randrange(2,307)
        b='B'+ str(n)
        c='C'+str(n)
        notifyMe(sh1[b].value,sh1[c].value)
        time.sleep(10*60)

