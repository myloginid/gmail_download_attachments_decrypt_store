# Something in line of http://stackoverflow.com/questions/348630/how-can-i-download-all-emails-with-attachments-from-gmail
# Make sure you have IMAP enabled in your gmail settings.
# Downloads zip and pdf files.
# Decrypts pdf files where possible.
# Renames file according to mail date and time
# tries to decrypt old files also 
import PyPDF2                                                                
import email
import getpass, imaplib, os, sys
from email.utils import parsedate_tz, mktime_tz, formatdate
import time , datetime
from PyPDF2 import PdfFileReader, PdfFileWriter

passwords =  ['PASSWORD1','PASSWORD2','PASSWORD3']
detach_dir = '/home/YOURFOLDER/attachments'
userName = 'YOURNAME@gmail.com'
passwd = 'YOURPASSWORD'
imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
typ, accountDetails = imapSession.login(userName, passwd)
if typ != 'OK':
    print 'Not able to sign in!'

search = [
    [  '(TO "YOURNAME@gmail.com" FROM "estatement@icicibank.com" SINCE "20-DEC-2014")' , "ICICI"],
    [  '(TO "YOURNAME@gmail.com" FROM "statements@axisbank.com" SINCE "20-DEC-2014")' , "AXIS"],
    [  '(FROM "YOURNAME@sharekhan.com" SINCE "20-DEC-2014")' , "SHARES"],
    [  '(FROM "YOURNAME@camsonline.com" SINCE "20-DEC-2014")' , "FUNDS"]
]

for s in range(0, len(search)):
    print " loop starting for  = " + search[s][0]  + " and folder " + search[s][1]
    search_str = search[s][0]
    working_dir = search[s][1]
    if working_dir not in os.listdir(detach_dir):
        os.mkdir(working_dir)
        
    imapSession.select('[Gmail]/All Mail')
    typ, data = imapSession.search(None, search_str )
    print "ok till search "
    if typ != 'OK':
        print 'Error searching Inbox.'
        raise
    # Iterating over all emails
    print "for loop starting"
    for msgId in data[0].split():
        typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
        if typ != 'OK':
            print 'Error fetching mail.'
            raise
        emailBody = messageParts[0][1]
        mail = email.message_from_string(emailBody)
        date_tuple = email.utils.parsedate_tz(mail['Date'])
        if date_tuple:
          local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
          mail_dt = local_date.strftime("%Y%m%d-%H%M%S-")
        # print mail['From']
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                # print part.as_string()
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue
            if bool(part.get_filename()):
                fileName = (mail_dt +  ' '.join(part.get_filename().split()).replace (" ", "_")).lower()
                filePath = os.path.join(detach_dir, working_dir, fileName)
                if not os.path.isfile( filePath )  and not os.path.isfile(os.path.join(detach_dir, working_dir, "Decrypted_" +  fileName))  :
                    print fileName
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                if os.path.isfile( filePath ) and  '.pdf' in fileName.lower():
                    print "reading file" + filePath
                    try:
                        pdfOne = PdfFileReader(file( filePath, "rb"))
                    except:
                        print "Read Error " + filePath
                    else:
                        pdfOne = PdfFileReader(file( filePath, "rb"))                        
                        if pdfOne.isEncrypted:
                            for p in passwords:
                                print "trying with " + p
                                try:
                                    pdfOne.decrypt(p)
                                except NotImplementedError:
                                    print "Decrypt Error" + filePath
                                else:
                                    dec_success = pdfOne.decrypt(p)
                                    if dec_success == 0 :
                                        continue
                                    else:
                                        print "dec_success = " + str(dec_success)
                                        print "Password that worked = " + p
                                        try:
                                            x = pdfOne.getNumPages()
                                        except PyPDF2.utils.PdfReadError:
                                            print "Get Pages Error" + filePath
                                        else:
                                            output = PdfFileWriter()
                                            for pg in pdfOne.pages:
                                                output.addPage(pg)
                                            # for i in range(0,pdfOne.getNumPages()):
                                                # output.addPage(pdfOne.getPage(i))
                                            outputStream = file(   os.path.join(detach_dir, working_dir, "Decrypted_" +  fileName) , "wb")
                                            output.write(outputStream)
                                            outputStream.close()
                                            if os.path.exists(os.path.join(detach_dir, working_dir, "Decrypted_" +  fileName)):
                                                # Delete original file if decrypted file exists 
                                                print "deleting " + filePath
                                                os.remove( filePath )
                                            break
                            print "no password worked for  " + filePath
imapSession.close()
imapSession.logout()
