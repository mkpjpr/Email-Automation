#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os 
import pandas as pd
import smtplib
from email.message import EmailMessage
from getpass import getpass

sender_email=input("Enter Your Email ID  ")
sender_pass=getpass("Enter Your Password ")


attach_folder=input("Enter Folder with attachments:")

body_path=input("Enter Text file with the Message:")

df=pd.read_excel(r"C:\Users\yoges\Desktop\CA Mukesh Pareek\New folder\Automate Template and Bulk Mail\Receiver Details.xlsx")
receivers_email=df["EMAIL_ID"].values
sub=("Transforming Financial Processes with Tech-Driven Chartered Accounting Solutions ")
attach_files=df["Files to be attached"].values
name=df["NAME"].values

print(attach_files)

zipped=zip(receivers_email,attach_files,name)
i=1
for(a,b,c) in zipped:
    
    msg=EmailMessage()
#     print(b)
    
    files=[(attach_folder+r"\{}.pdf".format(b))]
    print(files)
    
    file=files[0] #this selects the first item in the atachement list

    with open(file,'rb') as f:

        file_data=f.read()
        file_name=f.name

        print("File name is {}".format(file_name))

    msg['From']=sender_email
    msg['To']=a
    msg['Subject']=sub

    with open(body_path,encoding='utf-8') as g:


        body = g.read()
        body=body.format(c)

    msg.set_content(body+"<h3 style='color: red'> </h3>",subtype='html') 
    
    msg.add_attachment(file_data,maintype='application',subtype='octet-stream',filename="{}.pdf".format(b))

    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:

        smtp.login(sender_email,sender_pass)

        smtp.send_message(msg)
        
    print(str(i)+"  Sent")
    
    i=i+1
        

print("All mail sent!")

