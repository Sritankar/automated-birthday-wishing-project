from os import name
import pandas as pd
import datetime
import smtplib
import time
import requests
from win10toast import ToastNotifier

#your gmail credentials here
GMAIL_ID = 'your_email_here'
GMAIL_PWD = 'your_password_here'

#for desktop notification
toast = ToastNotifier()

#Function for sending email
def sendEmail(to,sub,msg):
    s = smtplib.SMTP('smtp.gmail.com',587)                          #conncection to gmail
    s.starttls()                                                    #starting the session
    s.login(GMAIL_ID,GMAIL_PWD)                                     #login using credentials
    s.sendmail(GMAIL_ID,to,f"Subject : {sub}\n\n{msg}")             #sending email
    s.quit()                                                        #quit the session
    print(f"Email sent to {to} with subject {sub} and message : {msg}")
    toast.show_toast("Email Sent!" , f"{name} was sent e-mail", threaded=True, icon_path=None, duration=6)

    while toast.notification_active():
        time.sleep(0.1)
        
def sendsms(to,msg,name,sub):
    url = "https://www.fast2sms.com/dev/bulk"
    payload = f"sender_id=FSTSMS&message={msg}&language=english&route=p&numbers={to}"
    headers = {
        'authorization': "API_KEY_HERE",
        'Content-Type': "application/x-www-form-urlencoded",
        'Cache-Control': "no-cache",
        }

    response = requests.request("POST", url, data=payload, headers=headers)
    print(response.text)
    print(f"SMS sent to {to} with subject : {sub} and message : {msg}")
    toast.show_toast("SMS Sent!" , f"{name} was sent message", threaded=True, icon_path=None, duration=6)

    while toast.notification_active():
        time.sleep(0.1)

if name=="main":
    df = pd.read_excel("excelsheet.xlsx")                           #read the excel sheet having all the details
    today = datetime.datetime.now().strftime("%d-%m")               #today's date in format : DD-MM
    yearNow = datetime.datetime.now().strftime("%Y")                #current year in format : YY
    writeInd = []                                                   #writeindex list

    for index,item in df.iterrows():
        msg = f"Many Many Happy Returns of the day dear {item['NAME']} !!!!!!\n\n\nThis is an automated email from **** sent using Python.\n"
        bday = item['Birthday'].strftime("%d-%m")                   #stripping the birthday in excel sheet as : DD-MM
        if (today==bday) and yearNow not in str(item['Year']):      #condition checking
            sendEmail(item['Email'], "Happy Birthday", msg)         #calling the sendEmail function
            sendsms(item['Contact'], msg, item['NAME'], "Happy Birthday")       #calling the sendsms function
            writeInd.append(index)                                  

    for i in writeInd:
        yr = df.loc[i,'Year']
        df.loc[i,'Year'] = str(yr) + ',' + str(yearNow)             #this will record the years in which email has been sent

    df.to_excel('excelsheet.xlsx', index=False)