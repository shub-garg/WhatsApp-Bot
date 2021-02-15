import pywhatkit as kt
import time
import os
import datetime
import pandas as pd 

# Reading data from excel file
file =(r'C:\Users\Shubham garg\Desktop\Hi.xls') 
#Enter the Phone Numbers column
cols = [0]  
newData = pd.read_excel(file, usecols=cols)
print(newData)

#display welcome msg
print("Let's Automate Whatsapp!")
i = 0
msg = input("Enter the message :")

#To extract system time
c_t = datetime.datetime.now() 
ct = c_t.strftime("%H:%M:%S")
t1 = ct.rsplit(":")
h = int(t1[0])
m = int(t1[1])
s = int(t1[2])

if (s < 57) :
    m = m + 1
else :
    m = m + 2
p_num = []
p_num = newData.values.tolist()
n = len(p_num)

# Loop for sending the message
for i in range(n):
    p_num[i] = "+91" + str(p_num[i])
    kt.sendwhatmsg(p_num[i],msg,h,m,10)
    time.sleep(15)
    os.system("TASKKILL /IM Chrome.exe")
    m = m + 1