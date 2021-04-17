#Code written by:Rudviq Bhavsar
#Final testing done on:15/12/2020

#App for implementing a notifier to display the Basic level GRE words.

import random
from plyer import notification
import openpyxl as op
import time


def notifyMe(title,message):
    notification.notify(
        title=title,
        message= message,
        app_icon="C:\\Users\\Rudviq\\PycharmProjects\\GRE_Notifier\\Basic.ico",
        timeout=20,
    )

if __name__=='__main__':

    while True:
        wb=op.load_workbook("GREWords.xlsx")
        sh1=wb['Basic Words']
        n=random.randrange(2,307)
        b='B'+ str(n)
        c='C'+str(n)
        notifyMe(sh1[b].value,sh1[c].value)
        time.sleep(10*60)