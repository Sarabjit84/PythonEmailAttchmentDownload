import imaplib
import email
from email.parser import HeaderParser
import os
import base64

import environ

env = environ.Env()

environ.Env.read_env()

server = 'outlook.office365.com'
user = env('EMAIL_USER')    #EMAIL ID
password = env('EMAIL_PASSWORD') #EMAIL PASSWORD
outputdir = env('OUTPUT_DIR')
keywords_list = ["testai"]

def connect(server, user, password):
    imap_conn = imaplib.IMAP4_SSL(server)
    imap_conn.login(user=user, password=password)
    return imap_conn

def delete_email(instance, email_id):
    typ, delete_response = instance.fetch(email_id, '(FLAGS)')
    typ, response = instance.store(email_id, '+FLAGS', r'(\Deleted)')
    print(delete_response)
    print(response)

#def download_attachmet (instance, email_id)
    #code to save attachment in local folder

def downloaAttachmentsInEmail(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    mail = email.message_from_bytes(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(outputdir + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))

""" Try for download"""
for  keyword in keywords_list:
    conn = connect(server, user, password)
    conn.select('Test1')
    typ, msg = conn.search(None, '(SUBJECT "' + keyword + '")')
    print(msg)
    msg = msg[0].split()
    for email_id in msg:
        print(email_id)
    downloaAttachmentsInEmail (conn, email_id, outputdir)
print('Downloaded')