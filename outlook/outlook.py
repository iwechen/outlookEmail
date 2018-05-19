#coding:utf-8
'''
Created on 2018年5月7日
@author: chenwei
Email:iwechen123@gmail.com
'''
import email
import imaplib
import smtplib
import datetime
import email.mime.multipart
from .config import imap_server,imap_port
import base64


class Outlook():
    def __init__(self):
        mydate = datetime.datetime.now()-datetime.timedelta(1)
        self.today = mydate.strftime("%d-%b-%Y")
        print(self.today)
        # self.imap = imaplib.IMAP4_SSL('imap-mail.outlook.com')
        # self.smtp = smtplib.SMTP('smtp-mail.outlook.com')

    def login(self, username, password):
        self.username = username
        self.password = password
        while True:
            try:
                self.imap = imaplib.IMAP4_SSL(imap_server,imap_port)
                r, d = self.imap.login(username, password)
                assert r == 'OK', 'login failed'
                print("登录成功！", d)
            except:
                print(" > Sign In ...")
                break
            # self.imap.logout()
            break

    def list(self):
        # self.login()
        return self.imap.list()

    def select(self, str):
        return self.imap.select(str)

    def inbox(self):
        return self.imap.select("Inbox")
        # return self.imap.select("payment")

    def junk(self):
        return self.imap.select("Junk")

    def logout(self):
        return self.imap.logout()

    def today(self):
        mydate = datetime.datetime.now()
        return mydate.strftime("%d-%b-%Y")

    def unreadIdsToday(self):
        r, d = self.imap.search(None, '(SINCE "'+self.today+'")', 'UNSEEN')
        list = d[0].split(' ')
        return list

    def getIdswithWord(self, ids, word):
        stack = []
        for id in ids:
            self.getEmail(id)
            if word in self.mailbody().lower():
                stack.append(id)
        return stack

    def hasUnread(self):
        list = self.unreadIds()
        return list != ['']

    def unreadIds(self):
        r, d = self.imap.search(None, "UNSEEN")
        list = d[0].split()
        return list

    def readIdsToday(self):
        r, d = self.imap.search(None, '(SINCE "'+self.today+'")', 'SEEN')
        list = d[0].split()
        return list

    def allIds(self):
        r, d = self.imap.search(None, "ALL")
        list = d[0].split()
        return list

    def readIds(self):
        r, d = self.imap.search(None, "SEEN")
        list = d[0].split()
        return list

    def getEmail(self, ids):
        r, d = self.imap.fetch(ids, "(RFC822)")
        # print(d[0][1])
        try:
            self.raw_email = d[0][1].decode('utf-8')
        except:
            self.raw_email = d[0][1].decode('gb2312')
        # print(self.raw_email)
        self.email_message = email.message_from_string(self.raw_email)
        # print(self.email_message)
        return self.email_message

    def unread(self):
        list = self.unreadIds()
        latest_id = list[-1]
        # print(latest_id)
        return self.getEmail(latest_id)

    def read(self):
        list = self.readIds()
        latest_id = list[-1]
        return self.getEmail(latest_id)

    def readToday(self):
        list = self.readIdsToday()
        latest_id = list[-1]
        return self.getEmail(latest_id)

    def unreadToday(self):
        list = self.unreadIdsToday()
        latest_id = list[-1]
        return self.getEmail(latest_id)

    def readOnly(self, folder):
        return self.imap.select(folder, readonly=True)

    def writeEnable(self, folder):
        return self.imap.select(folder, readonly=False)

    def rawRead(self):
        list = self.readIds()
        latest_id = list[-1]
        r, d = self.imap.fetch(latest_id, "(RFC822)")
        self.raw_email = d[0][1]
        return self.raw_email

    def mailbody(self):
        if self.email_message.is_multipart():
            for payload in self.email_message.get_payload():
                # if payload.is_multipart(): ...
                body = (
                    payload.get_payload()
                    .split(self.email_message['from'])[0]
                    .split('\r\n\r\n2015')[0]
                )
                return body
        else:
            body = (
                self.email_message.get_payload()
                .split(self.email_message['from'])[0]
                .split('\r\n\r\n2015')[0]
            )
            return body

    def mailsubject(self):
        return self.email_message['Subject']

    def mailfrom(self):
        return self.email_message['from']

    def mailto(self):
        return self.email_message['to']

    def mailreturnpath(self):
        return self.email_message['Return-Path']

    def mailreplyto(self):
        return self.email_message['Reply-To']

    def mailall(self):
        return self.email_message

    def mailbodydecoded(self):
        return base64.urlsafe_b64decode(self.mailbody())
