#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jul 27 11:07:14 2021

@author: jani4785
"""

import csv
from exchangelib import Credentials, Account
from exchangelib.ewsdatetime import EWSDateTime

from datetime import timedelta



USERNAME= 'mhrynuik@deerislecap.com'
PASSWORD= 'NoahButtons55'
credentials= Credentials(USERNAME, PASSWORD)
account= Account(USERNAME, credentials=credentials, autodiscover=True)

count=0

final_list=[]

for item in account.inbox.all().order_by('-datetime_received')[:1000]:
    mydict={}
    count=count+1
    print('subject->',item.subject)
    #print(dir(item))
    print('message_id->',item.message_id)
    #message_id
    print('datetime_received->', item.datetime_received)
    #print('headers->', item.headers)
    #print('body->', item.body)
    print('author->', item.author)
    print('text_body->', item.text_body)
    print('cc_recipient->', item.cc_recipients)
    #print('unique_body->', item.unique_body)
    
    mydict['MESSAGE_ID']=item.message_id
    mydict['DATE']=item.datetime_received
    mydict['FROM']=item.author
    mydict['SUBJECT']=item.subject
    mydict['CONTENT']=item.text_body
    mydict['CC_RECIPIENT']=item.cc_recipients
    
    
    final_list.append(mydict)
    
csvF=r"/Users/jani4785/Desktop/JsglpPython/outlookapi/outlookMichaelDeerIsle/Junk_Michael.csv"
with open(csvF, 'w', encoding='utf-8', newline = '') as csvfile: 
    fieldnames = ['MESSAGE_ID','DATE','FROM','CC_RECIPIENT','SUBJECT','CONTENT']#,'CONTENT2','NO_DIAGNOSTIC_CODE']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter = ',')
    writer.writeheader()
    for val in final_list:
        writer.writerow(val)
    
'''['ELEMENT_NAME', 
 'FIELDS', 
 'ID_ELEMENT_CLS', 
 'INSERT_AFTER_FIELD', 
 'NAMESPACE', '___id', 
 '__attachments', 
 '__author', 
 '__bcc_recipients', 
 '__body', 
 '__categories', 
 '__cc_recipients', 
 '__class__', 
 '__conversation_id', 
 '__conversation_index', 
 '__conversation_topic', 
 '__culture', 
 '__datetime_created', 
 '__datetime_received', 
 '__datetime_sent', 
 '__delattr__', 
 '__dict__', 
 '__dir__', 
 '__display_cc', 
 '__display_to', 
 '__doc__', 
 '__effective_rights', 
 '__eq__', 
 '__format__', 
 '__ge__', 
 '__getattribute__', 
 '__gt__', 
 '__has_attachments', 
 '__hash__', 
 '__headers', 
 '__importance', 
 '__in_reply_to', 
 '__init__', 
 '__init_subclass__', 
 '__is_associated', 
 '__is_delivery_receipt_requested', 
 '__is_draft', 
 '__is_from_me', 
 '__is_read', 
 '__is_read_receipt_requested', 
 '__is_resend', 
 '__is_response_requested', 
 '__is_submitted', 
 '__is_unmodified', 
 '__item_class', 
 '__last_modified_name', 
 '__last_modified_time', 
 '__le__', 
 '__lt__', 
 '__message_id', 
 '__mime_content', 
 '__module__', 
 '__ne__', 
 '__new__', 
 '__parent_folder_id', 
 '__received_by', 
 '__received_representing', 
 '__reduce__', 
 '__reduce_ex__', 
 '__references', 
 '__reminder_due_by', 
 '__reminder_is_set', 
 '__reminder_message_data', 
 '__reminder_minutes_before_start', 
 '__reply_to', 
 '__repr__', 
 '__response_objects', 
 '__sender', 
 '__sensitivity', 
 '__setattr__', 
 '__size', 
 '__sizeof__', 
 '__slots__', 
 '__str__', 
 '__subclasshook__', 
 '__subject', 
 '__text_body', 
 '__to_recipients', 
 '__unique_body', 
 '__web_client_edit_form_query_string', 
 '__web_client_read_form_query_string', 
 '_clear', 
 '_create', 
 '_delete', 
 '_field_vals', 
 '_fields_lock', 
 '_id', 
 '_slots_keys', 
 '_update', 
 '_update_fieldnames', 
 'account', 
 'add_field', 
 'archive', 
 'attach', 
 'attachments', 
 'attribute_fields', 
 'author', 
 'bcc_recipients', 
 'body', 
 'categories', 
 'cc_recipients', 
 'changekey', 
 'clean', 
 'conversation_id', 
 'conversation_index', 
 'conversation_topic', 
 'copy', 
 'create_forward', 
 'create_reply', 
 'create_reply_all', 
 'culture', 
 'datetime_created', 
 'datetime_received', 
 'datetime_sent', 
 'delete', 
 'deregister', 
 'detach', 
 'display_cc', 
 'display_to', 
 'effective_rights', 
 'folder', 
 'forward', 
 'from_xml', 
 'get_field_by_fieldname', 
 'has_attachments', 
 'headers', 
 'id', 
 'id_from_xml', 
 'importance', 
 'in_reply_to', 
 'is_associated', 
 'is_delivery_receipt_requested', 
 'is_draft', 
 'is_from_me', 
 'is_read', 
 'is_read_receipt_requested', 
 'is_resend', 
 'is_response_requested', 
 'is_submitted', 
 'is_unmodified', 
 'item_class', 
 'last_modified_name', 
 'last_modified_time', 
 'mark_as_junk', 
 'message_id', 
 'mime_content', 
 'move', 
 'move_to_trash', 
 'parent_folder_id', 
 'received_by', 
 'received_representing', 
 'references', 
 'refresh', 
 'register', 
 'reminder_due_by', 
 'reminder_is_set', 
 'reminder_message_data', 
 'reminder_minutes_before_start', 
 'remove_field', 
 'reply', 
 'reply_all', 
 'reply_to', 
 'request_tag', 
 'response_objects', 
 'response_tag', 
 'save', 
 'send', 
 'send_and_save', 
 'sender', 
 'sensitivity', 
 'size', 
 'soft_delete', 
 'subject', 
 'supported_fields', 
 'text_body', 
 'to_id_xml', 
 'to_recipients', 
 'to_xml', 
 'unique_body', 
 'validate_field', 
 'web_client_edit_form_query_string', 
 'web_client_read_form_query_string']'''










