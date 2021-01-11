""" 
    Ana Lucía Hernández 
    08/05/2020
    *Then edited by Sebastian Gonzalez to add zip files*

    Email file extractor to store .csv data in attachments
    Taken from: https://gist.github.com/baali/2633554

    To run:
        python3 gmail_extract.py [--mode daily --folder directory_number]
    
    -> If no directory_number is specified, mails will be taken from INBOX.
    -> 'full' is the default mode option, that allows to download all mails in the directory. If 'daily' is specified, only mails from the day before 
    (from 0:00 hrs to 24:00 hrs) will be taken from the directory. 
    -> directory_number can be an integer in [1,7] range, according values are described below in the gmail_folders variable
    -> both arguments are optional.

"""
import email
import getpass, imaplib
import os
import sys
import base64
import argparse
from datetime import datetime
import csv


#for the zips files
import shutil
import io
import zipfile

#fill
detach_dir = '' # directory in which the attachments folder will be stored
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('') #if not, create one in.. FILL

ap = argparse.ArgumentParser()
ap.add_argument("--mode", nargs='?', required=False, help="Time frame mode of mail extraction. Can be 'daily' or 'full'")
ap.add_argument("--folder", required=False, help="Folder number to get all mail from, available numbers are [1,7]")

#fill this with your information
email_name = ''
pwd = ''
gmail_folders = [
    '', #empty string to get All Mail
    'INBOX',
    'Drafts',
    'Important',
    'Junk', #spam
    'Flagged', #starred
    'Trash'
]

args = vars(ap.parse_args())
dir_number = args['folder']
mode = 'full' if args['mode'] == None else args['mode']
folder = 1 if dir_number == None else int(dir_number) -1
if folder > 6 or folder < 0:
    raise Exception('Invalid folder number. Choose from [1,7].')

imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
typ, accountDetails = imapSession.login(email_name, pwd)
if typ != 'OK':
    raise Exception('Not able to sign in.')

folder_s=''
if folder == 1:
    folder_s = gmail_folders[1]
elif folder > 1:
    for k in imapSession.list()[1]:
        if ('\\'+gmail_folders[folder]) in k.decode("utf-8"):
            f_split = k.decode("utf-8").split(' ')
            folder_s = f_split[len(f_split)-1].split('"')[1]

imapSession.select() if folder == 0 else imapSession.select(folder_s)

months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
search_mode = 'ALL'
if mode == 'daily':
    date = str(datetime.now().date()).split('-')
    day_before, day_now, month, year = str(int(date[2])-1), date[2], months[int(date[1])-1], date[0]
    search_mode = '(SINCE "'+ day_before+'-'+month+'-'+ year+'"' + ' BEFORE "'+ day_now+'-'+month+'-'+ year+'")'

typ, data = imapSession.search(None, search_mode)
if typ != 'OK':
    raise Exception('Error searching Inbox.')

other_files = {} # lleva cuenta de cuantos files mandaron que no sean csvs y de que tipo 

messages = data[0].split()
# Iterating over emails
for msgId in range(len(messages)-1, -1, -1): #get from newer to older
    typ, messageParts = imapSession.fetch(messages[msgId], '(RFC822)')
    if typ != 'OK':
        raise Exception('Error fetching mail.')

    emailBody = messageParts[0][1]
    mail = email.message_from_string(emailBody.decode('utf-8'))

    #print('sender: ', mail['From'])
    #print('date: ', mail['Date'])
    # print(mail.get_content_maintype())
    if mail.get_content_maintype() == 'multipart' or mail.is_multipart(): # get only mails with attachments or other components
        dir_zipped='' #fill where to put zippeds, put it in a blank directory please
        dir_unzip = '' #fill where tu unzip


        for part in mail.walk():
            was_zip = False
            
            print(part.get_filename())
            # print('\n', part.get_payload(decode=True), '\n')
            if part.get_content_maintype() == 'multipart': continue
            # if part.get_content_maintype() == 'text': continue
            # if part.get('Content-Disposition') == 'inline': continue
            if part.get('Content-Disposition') is None: continue
            
            attachment_name = part.get_filename()
            try:
                att_ext = attachment_name.split('.')[1]
            except:
                att_ext = attachment_name
            if type(att_ext) == str:        
              if att_ext == 'zip' or att_ext.startswith('=?UTF-8?b?'):
                was_zip = True
                #if zip file we will create an empty zip file first and then move
                #it to the directory we want
                print('found a zip')
                print('sender: ', mail['From'])
                print('date: ', mail['Date'])
                print('subject: ', mail['Subject'], '\n')


                name_temp_z = 'test22.zip'
                zf = zipfile.ZipFile(name_temp_z, 'w')
                path = dir_zipped + name_temp_z

                # This works:
                zfi = zipfile.ZipInfo(part.get_filename())
                zf.writestr(zfi, part.get_payload(decode=True))
                shutil.move(name_temp_z, dir_zipped) 
                print(dir_zipped + name_temp_z)
                print(part.get_filename())
                #this still doesnt work but looks like it create the file
                with zipfile.ZipFile(path, 'r') as zip_ref:
                  zip_ref.extractall(dir_unzip)

                zip_ref.close()
                os.remove(path)


          #add if rar then 
            if att_ext == 'rar':
              print('rar was found')

            if att_ext != 'csv' or 'zip' or 'rar':
                if att_ext not in other_files.keys():
                    other_files[att_ext] = 1
                else:
                    other_files[att_ext] +=1
              # To evaluate if the filename matches the one we want, we get the name then make a slice and see if it matches the string 
            if was_zip:
              attachment_names = []
              #list of files to record in directory
              print('the files found in the zip where:')
              for file in os.listdir(dir_unzip):
                print(file + ' attached')
                attachment_names.append(file)
              print('subject: ', mail['Subject'], '\n')
              print(attachment_names)
              for attachment_name in attachment_names:
                try:
                  if attachment_name[:19] == 'DETALLE_ESTACIONES_':
                      print('\t', attachment_name)
                      day, month, year = attachment_name[19:27][-2:], attachment_name[19:27][-4:-2], attachment_name[19:27][:-4] 
                      date = day + '_' + month + '_' + year
                      new_dir = os.path.join(detach_dir, 'attachments', str(date))
                      if not os.path.isdir(new_dir): # if directory is nonexistent, create it
                          os.makedirs(new_dir)

                      past_dir = dir_unzip + attachment_name
                      if (os.path.exists(new_dir + '/' + attachment_name) == False):
                        print(new_dir + '/' + attachment_name)
                        shutil.move(past_dir, new_dir)

                      # despues de guardarlo, se carga otra vez y se intenta parsear 
                      # si no se puede se lee el archivo sin modo bytes y se guarda
                      filePath = os.path.join(new_dir, attachment_name)
                      try:
                          with open(filePath, mode='r') as csv_file_o:
                              csv_reader = csv.reader(csv_file_o, delimiter=',')
                              for row in csv_reader:
                                  break
                      except UnicodeDecodeError:
                        fp = open(filePath, 'w')
                        fp.write(part.get_payload(decode=False))
                        fp.close()
                      except TypeError:
                        print('error')
                        print(attachment_name)
                except TypeError:
                  print(attachment_name)


              #vamos a limpiar la data que quedo en el archivo de unzipping que no
              #contenia archivos que nos interesaran
              folder = dir_unzip
              for filename in os.listdir(folder):
                  file_path = os.path.join(folder, filename)
                  try:
                      if os.path.isfile(file_path) or os.path.islink(file_path):
                          os.unlink(file_path)
                      elif os.path.isdir(file_path):
                          shutil.rmtree(file_path)
                  except Exception as e:
                      print('Failed to delete %s. Reason: %s' % (file_path, e))

            else:      
                try:
                    if attachment_name[:19] == 'DETALLE_ESTACIONES_':
                        #print('\t', attachment_name)
                        day, month, year = attachment_name[19:27][-2:], attachment_name[19:27][-4:-2], attachment_name[19:27][:-4] 
                        date = day + '_' + month + '_' + year
                        new_dir = os.path.join(detach_dir, 'attachments', str(date))
                        if not os.path.isdir(new_dir): # if directory is nonexistent, create it
                            os.makedirs(new_dir)
                        filePath = os.path.join(new_dir, attachment_name)
                        # download file
                        if not os.path.isfile(filePath):
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()

                        # despues de guardarlo, se carga otra vez y se intenta parsear 
                        # si no se puede se lee el archivo sin modo bytes y se guarda
                        try:
                            with open(filePath, mode='r') as csv_file_o:
                                csv_reader = csv.reader(csv_file_o, delimiter=',')
                                for row in csv_reader:
                                    break
                        except UnicodeDecodeError:
                            fp = open(filePath, 'w')
                            fp.write(part.get_payload(decode=False))
                            fp.close()
                except TypeError:
                    print('typerror')
                    print(attachment_name)
                


            
    # print('\n')
    #imapSession.store(messages[msgId], '+FLAGS', '(SEEN)')

print("\nArchivos diferentes de csv:")
print(other_files)
print("\n")
imapSession.close()
imapSession.logout()