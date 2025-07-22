print("Jsglp")
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jun  6 15:15:35 2021

@author: jani4785
"""

print("Jsglp")
print("Jsglp")
print("Jsglp")
print("Jsglp")
print("Jsglp")
print("Jsglp")
print("Jgslp")
print("Jsglp")
print("Jsglp")
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jun  1 13:55:41 2021

@author: jani4785
"""

print("Jsglp")
print("Jsglp")
print("Jsglp")


credentials=r"/Users/jani4785/Desktop/JsglpPython/gmailapi_c2ventures/c2ventures1.json"


#STEP 1) CONNECTION TO GMAIL API

#from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import time
import base64
import email
import mailparser
import json
import csv
import pandas as pd

# If modifying these scopes, delete the file stoken.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('c2ventures1.json'):
        creds = Credentials.from_authorized_user_file('c2ventures1.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                #modified below to point to my .json file
                credentials, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('c2ventures1.json', 'w') as token:
            token.write(creds.to_json())

    conn = build('gmail', 'v1', credentials=creds)
    
    return conn

service=main()

#STEP 2) IMPORTING ALL MESSAGE ID'S 
def importMsgids(service):
    
    '''emailSum=service.users().getProfile(userId="me").execute()
    print("summary of email-->",emailSum)
    print("\n")
    
    labelnm_lableIds=service.users().labels().list(userId="me").execute()
    print("all lablenames & lableid's")
    print(labelnm_lableIds)'''
    
    
    z=service.users().messages().list(userId="me",labelIds="INBOX",maxResults=20).execute()
    msg_ids=[]
    #print(z)
    
    page_token=None
    
    if 'nextPageToken' in z:
        page_token=z['nextPageToken']
        
    while page_token:

        z=service.users().messages().list(userId="me",labelIds="INBOX",pageToken=page_token).execute()
        zx=z.get('messages',[])
        #print((zx))
        for msg in zx:
            #print(msg['id'])
            msg_ids.append(msg['id'])

         
        
        if 'nextPageToken' in z:
            page_token=z['nextPageToken']
        else:
            break
    
    
    print(len(msg_ids))
    return msg_ids


##STEP 3), Importing each message and exporting different columns into CSV file
def msgMainLoads(service):
  
    #createing list to be able to append into csv file
    final_list=[]
    
    #adding counter to see whcih message that we are currently at
    count=0
    #importing message id's from previous function above
    mymsgid=importMsgids(service)
    #mymsgid=['179c8b0427658618', , ]
    
    try:   
        #give number of messages that you want to download
        for msgids in mymsgid[0:100000]:
            #Creating dictionary to add each column and its message value
            mydict={}
            count=count+1
            print(count,' ',msgids)
            
            #Importing raw message in bytes
            message=service.users().messages().get(userId='me',id=msgids,format='raw').execute()
            #encoding into ASCII
            msg_raw=base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
            #print(msg_raw)
            
            #Converting into readable string data type
            msg_str=email.message_from_bytes(msg_raw)
            #print(msg_str.get_content_maintype())
            #print(msg_str.get_payload())
            
            #verifying message type (text or MultPart, most of them are multiparts)
            contenttype=msg_str.get_content_maintype()
            print(contenttype)
            
            mail = mailparser.parse_from_bytes(msg_raw)
            print(mail.date)
            print(mail.from_)
            print(mail.subject)
            #print(mail.MESSAGE_ID)
            #print(mail.to)
            print(mail.mailparser)
            
            #adding MESSAGE_ID,DATE,SUBJECT etc for csv purpose
            mydict['MESSAGE_ID']=message['id']
            mydict['DATE']=mail.date
            mydict['FROM']=mail.from_
            mydict['SUBJECT']=mail.subject
            mydict['MSG_SNIPPET']=message['snippet']
            
            
            #Observation
            #content_1 is for the main body
            #content_2 is for the Diagnostic code, final recipient email, status codes etc
            dict_content_1=''
            if contenttype=='multipart':                
                partType1gc=msg_str.get_content_maintype()
                print(partType1gc)
                bodyParts=msg_str.get_payload()
                firstBodyPart=bodyParts[0]
                #print(firstBodyPart)
                x=firstBodyPart.values()
                
                x=str(x).replace('\r\n','')
                x=str(x).replace('\t',' ')
                #print(x)
                x=str(x).lower()
                print("printing x")
                print(x)
                
                if x=='''['text/plain; charset="utf-8"']''':
                    print('''printed from -->text/plain; charset="utf-8"''')
                    print(str(firstBodyPart).replace('''Content-Type: text/plain; charset="UTF-8"''',''))
                    dict_content_1=dict_content_1+str(firstBodyPart).replace('''Content-Type: text/plain; charset="UTF-8"''','')
                
                if x=='''['quoted-printable', 'text/plain; charset=utf-8']''':
                    print(''''printed from -->quoted-printable', 'text/plain; charset=utf-8''')
                    print(str(firstBodyPart)[85:])
                    dict_content_1=dict_content_1+str(firstBodyPart)[85:]
                    
                
                print(firstBodyPart.get_content_maintype())
                if firstBodyPart.get_content_maintype()=="multipart":
                    bpartsfull=firstBodyPart.get_payload()
                    bpartsfull2=bpartsfull[0]
                    
                    x=bpartsfull2.values()
                    x=str(x).replace('\r\n','')
                    x=str(x).replace('\t',' ')
                    #print(x)
                    x=str(x).lower()
                    print("printing x")
                    print(x)
                    print("gurrr-->")
                    #print(bpartsfull[0])
                    if x=='''['text/plain; charset="us-ascii"', 'quoted-printable']''':
                        print('''printed from -->text/plain; charset="us-ascii"', 'quoted-printable''')
                        print(str(bpartsfull2)[89:])
                        dict_content_1=dict_content_1+str(bpartsfull2)[89:]
                    
                    if x=='''['text/plain; charset="utf-8"', 'base64']''':
                        print(''''printed from if condition -->text/plain; charset="utf-8"', 'base64''')
                        message_bytes = base64.b64decode(str(bpartsfull2)[77:1000000])
                        finalMsg=message_bytes.decode('utf-8')
                        print(finalMsg)
                        dict_content_1=dict_content_1+finalMsg
                    
                    if x=='''['base64', 'text/plain; charset="utf-8"']''':
                        message_bytes = base64.b64decode(str(bpartsfull2)[77:1000000])
                        finalMsg=message_bytes.decode('utf-8')
                        print(finalMsg)
                        dict_content_1=dict_content_1+finalMsg


                    if x=='''['text/plain; charset="utf-8"', '1.0', 'base64', 'address verification request']''':
                        message_bytes = base64.b64decode(str(bpartsfull2)[143:])
                        finalMsg=message_bytes.decode('utf-8')
                        print(finalMsg)
                        dict_content_1=dict_content_1+finalMsg
                        
                    #check if it has any negative impact , if it does , get rid of it    
                    if bpartsfull2.get_content_maintype()=="multipart":
                        partsfull3=bpartsfull2.get_payload()
                        print("---->")
                        partsfull33=partsfull3[0]
                        
                        message_bytes = base64.b64decode(str(partsfull33)[77:1000000])
                        finalMsg=message_bytes.decode('utf-8')
                        print(finalMsg)
                        dict_content_1=dict_content_1+finalMsg
                        
                    else:
                        print("printing from elseme")
                        print(bpartsfull2)
                        dict_content_1=dict_content_1+str(bpartsfull2)
                        
                        
                        
                elif firstBodyPart.get_content_maintype()=="text":
                    if x=='''['text/plain; charset="utf-8"', 'base64']''':
                                             
                        efparts=firstBodyPart.get_payload()
                        message_bytes = base64.b64decode(str(efparts))
                        finalMsg=message_bytes.decode('utf-8')
                        print('''printed from -->'text/plain; charset="utf-8"', 'base64''')   
                        print(finalMsg)
                        dict_content_1=dict_content_1+finalMsg
                        
                    if x=='''['text/plain; charset="us-ascii"', 'quoted-printable']''':
                        print('''printed from -->'ttext/plain; charset="us-ascii"', quoted-printable''')   
                        print(str(firstBodyPart)[88:])
                        dict_content_1=dict_content_1+str(firstBodyPart)[88:]
                    
                    if x=='''['text/plain; charset=utf-8; format=flowed', '8bit']''':
                        print('''printed from -->'text/plain; charset=utf-8; format=flowed', '8bit''') 
                           
                        #print(str(firstBodyPart)[92:])
                        message_bytes = base64.b64decode(str(firstBodyPart)[92:1000000])
                        finalMsg=message_bytes.decode('utf-8')
                        print(finalMsg)
                        dict_content_1=dict_content_1+finalMsg
                    
                    #below else created for this purpose only  ['text/plain;\r\n\tcharset="us-ascii"', '7bit']
                    #msg id: 178dbc567f03778c to know more about below else
                    
                    else:
                        print('''printed from ---> else from text''')
                        print(str(firstBodyPart)[77:])
                        dict_content_1=dict_content_1+str(firstBodyPart)[77:]
                        
                        #print(efparts)
                        
                    
                
                #print(firstBodyPart)
                #outf.write(str(firstBodyPart))
                

                

                        
                    

            else:
                print("\n")
                print("printing non Multipart message")   
                #print(str(msg_str))
  
                parts=msg_str.get_payload()
                print(parts)
                #dont uncomment below , because it stopping loop sometimes
                #message_bytes = base64.b64decode(str(parts))
                #finalMsg=message_bytes.decode('utf-8')
                #print(finalMsg)
                dict_content_1=dict_content_1+str(parts)
                
                #if you add below linke file breaking again, so may be dont add
                #also when i noticed for 100 message it only missed two bodies
                #notmultiPart=notmultiPart+str(parts)
                #print(str(parts))
                #dict_content_1=dict_content_1+str(parts)
                print("above messaged printed final else nonmultipart")
                #outf.write(str(parts))
        
           
            mydict['CONTENT_FULLBODY']=dict_content_1
            #mydict['CONTENT2']=content_2
            #mydict['NO_DIAGNOSTIC_CODE']=notmultiPart
            final_list.append(mydict)
            
    except:
        pass
    
        #print(final_list)
    
    #exporting final_list into csv file
    csvF=r"/Users/jani4785/Desktop/JsglpPython/gmailapi_c2ventures/InboxFinalc2v/InboxFinalc2v1.csv"
    with open(csvF, 'w', encoding='utf-8', newline = '') as csvfile: 
        fieldnames = ['MESSAGE_ID','DATE','FROM','SUBJECT','MSG_SNIPPET','CONTENT_FULLBODY']#,'CONTENT2','NO_DIAGNOSTIC_CODE']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter = ',')
        writer.writeheader()
        for val in final_list:
            writer.writerow(val)
        
    



if __name__ == '__main__':
    #print("sldkf")
    print(importMsgids(service))
    msgMainLoads(service)