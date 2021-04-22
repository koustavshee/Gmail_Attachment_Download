# Something in lines of http://stackoverflow.com/questions/348630/how-can-i-download-all-emails-with-attachments-from-gmail
# Make sure you have IMAP enabled in your gmail settings.
# Make sure less secure apps are enabled in gmail settings
# Right now it won't download same file name twice even if their contents are different.
#### ****Originally made by baali [https://gist.github.com/baali]**** ####
# Modified to run on python 3.9

import email
import getpass, imaplib
import os
import sys
import time
from tqdm import tqdm

#start time
start = time.time()

#initialize counter
count = 0

#seting file list
file_list = []

FolderName = input('Enter File Name: ')
detach_dir = input('Enter path: ')

if FolderName not in os.listdir(detach_dir):
    os.chdir(detach_dir)
    os.mkdir(FolderName)

userName = input('Enter your GMail username: ')
passwd = getpass.getpass('Enter your password: ')

try:
    imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
    typ, accountDetails = imapSession.login(userName, passwd)
    if typ != 'OK':
        print('Not able to sign in!')
    
    imapSession.select('inbox')
    typ, data = imapSession.search(None, 'ALL')
    if typ != 'OK':
        print('Error searching Inbox.')
        
    
    #Iterating over all emails
    for msgId in data[0].split() and tqdm(data[0].split()):
        typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
        if typ != 'OK':
            print('Error fetching mail.')
            

        emailBody = messageParts[0][1]
        mail = email.message_from_bytes(emailBody)
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                # print part.as_string()
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue

            fileName = part.get_filename()
            total_size = len(str(part))

            if bool(fileName):
                filePath = os.path.join(detach_dir, FolderName, fileName)
                if not os.path.isfile(filePath) :
                    file_list.append(fileName)
                    count=count+1
                    with open(filePath, 'wb') as fp:
                        fp.write(part.get_payload(decode=True))
                    fp.close()
    imapSession.close()
    imapSession.logout()
    textpath = os.path.join(detach_dir, FolderName, "file_list.txt")
    if not os.path.isfile(textpath):
        with open(textpath, 'w+') as tp:
            for fname in file_list:
                tp.write("%s\n" % fname)
            tp.write("\nTotal no of files %s" % count)
except :
    print('Not able to download all attachments.')

# end time
end = time.time()
ExecutionTime = "{:.2f}".format(end - start)

# total time taken
print(f"Runtime of the program is {ExecutionTime} s")
print(f"Total attachment found: {count}")