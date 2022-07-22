
# region Import
from email.message import Message
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
from email.utils import decode_params
from re import sub
import smtplib
import imaplib
import email
import os
# region Only for Read_Email_ByCredentialConsole
# import the required libraries
# https://www.geeksforgeeks.org/how-to-read-emails-from-gmail-using-gmail-api-in-python/
# pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
# pip install beautifulsoup4
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path
import base64
import email
from bs4 import BeautifulSoup
# endregion
# endregion

# region Class Service_Email_Class
class Service_Email_Class:
    @staticmethod
    def Delete_Email(IMAPDetail_Dict,EmailDelete_Dict_List):        
        Service_Email_Helper_Class_Var = Service_Email_Helper_Class()
        IMAP = Service_Email_Helper_Class_Var.Crate_IMAP(IMAPDetail_Dict)

        # --
        for EachItem in EmailDelete_Dict_List:      
            Folder = EachItem["Folder"]
            Filter = EachItem["Filter"]

            IMAP.select(Folder)

            status, messages = None,None
            if Filter.find("FROM")!=-1:
                # search for specific mails by sender
                status, messages = IMAP.search(None,Filter)
            if Filter.find("ALL")!=-1:
                # to get all mails
                status, messages = IMAP.search(None, "ALL")
            if Filter.find("SUBJECT")!=-1:        
                # to get mails by subject
                status, messages = IMAP.search(None, 'SUBJECT "Thanks for Subscribing to our Newsletter !"')
            if Filter.find("SINCE")!=-1:                    
                # to get mails after a specific date
                status, messages = IMAP.search(None, 'SINCE "01-JAN-2020"')
            if Filter.find("BEFORE")!=-1:                                
                # to get mails before a specific date
                status, messages = IMAP.search(None, 'BEFORE "01-JAN-2020"')
                # convert messages to a list of email IDs

            # ---
            if messages[0].decode() != "":
                messages = messages[0].split(b' ')
                for mail in messages:
                    _, msg = IMAP.fetch(mail, "(RFC822)")
                    # you can delete the for loop for performance if you have a long list of emails
                    # because it is only for printing the SUBJECT of target email to delete
                    for response in msg:
                        if isinstance(response, tuple):
                            msg = email.message_from_bytes(response[1])
                            # decode the email subject
                            subject = decode_header(msg["Subject"])[0][0]
                            if isinstance(subject, bytes):
                                # if it's a bytes type, decode to str
                                subject = subject.decode()
                            print("Deleting", subject)
                    # mark the mail as deleted
                    IMAP.store(mail, "+FLAGS", "\\Deleted")
                # permanently remove mails that are marked as deleted
                # from the selected mailbox (in this case, INBOX)
            else:
                pass
        # permanently remove mails that are marked as deleted
        # from the selected mailbox (in this case, INBOX)
        IMAP.expunge()
        # close the mailbox
        IMAP.close()
        # logout from the account
        IMAP.logout()

    @staticmethod
    def Create_Mail(SMTPDetail_Dict,EmailFromToDetail_Dict_List):
        Service_Email_Helper_Class_Var = Service_Email_Helper_Class()
        smtp = Service_Email_Helper_Class_Var.Crate_SMTP(SMTPDetail_Dict)

        # Provide some data to the sendmail function!
        for EachItem in EmailFromToDetail_Dict_List:        
            for EachToItem in EachItem["EmailTo_Id_List"]:   
                try:
                    smtp.sendmail(from_addr=EachItem["EmailFrom_Id"],to_addrs=EachToItem, msg=EachItem["Email_Message"].as_string())
                    print("Sucess:Create_Mail:"+str(EachItem["EmailFrom_Id"])+":"+str(EachToItem))
                except:
                    print("Error:Create_Mail:"+str(EachItem["EmailFrom_Id"])+":"+str(EachToItem))

        # Finally, don't forget to close the connection
        smtp.quit()

    @staticmethod
    def Read_Email(IMAPDetail_Dict,EmailReadFrom_List):        
        Service_Email_Helper_Class_Var = Service_Email_Helper_Class()
        IMAP = Service_Email_Helper_Class_Var.Crate_IMAP(IMAPDetail_Dict)

        Message_ListOfList = []
        for EachItem in EmailReadFrom_List:
            Message_Encrypt_List = Service_Email_Helper_Class_Var.Read_Email_FromSearchResult(Service_Email_Helper_Class_Var.ReadSearch_Message('FROM',EachItem,IMAP),IMAP)
            # printing them by the order they are displayed in your gmail
            Message_ListOfList.append(Service_Email_Helper_Class_Var.ReadDecrypt_Email(Message_Encrypt_List))
        return Message_ListOfList


        # Provide some data to the sendmail function!
        for EachItem in EmailFromToDetail_Dict_List:        
                try:
                    smtp.sendmail(from_addr=EachItem["EmailFrom_Id"],to_addrs=EachToItem, msg=EachItem["Email_Message"].as_string())
                    print("Sucess:Create_Mail:"+str(EachItem["EmailFrom_Id"])+":"+str(EachToItem))
                except:
                    print("Error:Create_Mail:"+str(EachItem["EmailFrom_Id"])+":"+str(EachToItem))

        # Finally, don't forget to close the connection
        smtp.quit()

    @staticmethod
    def Read_Email_FromCredentialGoogleCloudConsole(File_Path_Credential):
    	# Define the SCOPES. If modifying it, delete the token.pickle file.
        SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

    	# Variable creds will store the user access token.
    	# If no valid token found, we will create one.
        creds = None

    	# The file token.pickle contains the user access token.
    	# Check if it exists
        if os.path.exists('token.pickle'):

    		# Read the token from the file and store it in the variable creds
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

    	# If credentials are not available or are invalid, ask the user to log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(File_Path_Credential, SCOPES)
                creds = flow.run_local_server(port=0)

    		# Save the access token in token.pickle file for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

    	# Connect to the Gmail API
        service = build('gmail', 'v1', credentials=creds)

    	# request a list of all the messages
        result = service.users().messages().list(userId='me').execute()

    	# We can also pass maxResults to get any number of emails. Like this:
    	# result = service.users().messages().list(maxResults=200, userId='me').execute()
        messages = result.get('messages')

    	# messages is a list of dictionaries where each dictionary contains a message id.

    	# iterate through all the messages
        for msg in messages:
    		# Get the message from its id
            txt = service.users().messages().get(userId='me', id=msg['id']).execute()

    		# Use try-except to avoid any Errors
            try:
    			# Get value of 'payload' from dictionary 'txt'
                payload = txt['payload']
                headers = payload['headers']

    			# Look for Subject and Sender Email in the headers
                for d in headers:
                    if d['name'] == 'Subject':
                        subject = d['value']
                    if d['name'] == 'From':
                        sender = d['value']

    			# The Body of the message is in Encrypted format. So, we have to decode it.
    			# Get the data and decode it with base 64 decoder.
                parts = payload.get('parts')[0]
                data = parts['body']['data']
                data = data.replace("-","+").replace("_","/")
                decoded_data = base64.b64decode(data)

    			# Now, the data obtained is in lxml. So, we will parse
    			# it with BeautifulSoup library
                soup = BeautifulSoup(decoded_data , "lxml")
                body = soup.body()

    			# Printing the subject, sender's email and message
                print("Subject: ", subject)
                print("From: ", sender)
                print("Message: ", body)
                print('\n')
            except:
                pass
# endregion

# region Class Service_Email_Class
class Service_Email_Helper_Class:
    @staticmethod
    def ReadDecrypt_Email(Message_Encrypt_List):
        Message_Decrypt_List = []
        for msg in Message_Encrypt_List[::-1]:
          	for sent in msg:
                  if type(sent) is tuple:
                      content = str(sent[1], 'utf-8')
                      data = str(content)
          
          			# Handling errors related to unicodenecode
                      try:
                          indexstart = data.find("ltr")
                          data2 = data[indexstart + 5: len(data)]
                          indexend = data2.find("</div>")
          
          				# printtng the required content which we need
          				# to extract from our email i.e our body
                        #   print(data2[0: indexend])
                          Message_Decrypt_List.append(data2[0: indexend])           
                      except UnicodeDecodeError as e:
                          pass
        return Message_Decrypt_List

    @staticmethod
    def Crate_IMAP(IMAPDetail_Dict):
        # initialize connection to our
        # email server, we will use gmail here
        # this is done to make SSL connection with GMAIL
        IMAP = imaplib.IMAP4_SSL(IMAPDetail_Dict["IMAP_Adr"])

        # logging the user in
        IMAP.login(IMAPDetail_Dict["IMAP_Mail_Id"],IMAPDetail_Dict["IMAP_Mail_Psw"])

        # calling function to check for email under this label
        IMAP.select('Inbox')

        return IMAP

    @staticmethod
    def Crate_SMTP(SMTPDetail_Dict):
        # initialize connection to our
        # email server, we will use gmail here
        smtp = smtplib.SMTP(SMTPDetail_Dict["SMTP_Adr"],SMTPDetail_Dict["SMTP_Port"])
        smtp.ehlo()
        smtp.starttls()

        # Login with your email and password
        smtp.login(SMTPDetail_Dict["SMTP_Mail_Id"],SMTPDetail_Dict["SMTP_Mail_Psw"])

        return smtp

    @staticmethod
    def Create_Message(subject="Notification", ContentDetail_Dict=None, img=None, attachment=None):
    
        # build message contents
        msg = MIMEMultipart()
    
        # Add Subject
        msg["Subject"] = subject
    
        # Add text contents
        if ContentDetail_Dict["Content_Type"]=="Html":
            msg.attach(MIMEText(ContentDetail_Dict["Content"],'html'))

        if ContentDetail_Dict["Content_Type"]=="Text":
            msg.attach(MIMEText(ContentDetail_Dict["Content"]))
    
        # Check if we have anything
        # given in the img parameter
        if img is not None:
        
            # Check whether we have the lists of images or not!
            if type(img) is not list:
            
                # if it isn't a list, make it one
                img = [img]
    
            # Now iterate through our list
            for one_img in img:
            
                # read the image binary data
                img_data = open(one_img, "rb").read()
                # Attach the image data to MIMEMultipart
                # using MIMEImage, we add the given filename use os.basename
                msg.attach(MIMEImage(img_data, name=os.path.basename(one_img)))
    
        # We do the same for
        # attachments as we did for images
        if attachment is not None:
        
            # Check whether we have the
            # lists of attachments or not!
            if type(attachment) is not list:
            
                # if it isn't a list, make it one
                attachment = [attachment]
    
            for one_attachment in attachment:
            
                with open(one_attachment, "rb") as f:
                
                    # Read in the attachment
                    # using MIMEApplication
                    file = MIMEApplication(f.read(), name=os.path.basename(one_attachment))
                file[
                    "Content-Disposition"
                ] = f'attachment;\
    			filename="{os.path.basename(one_attachment)}"'
    
                # At last, Add the attachment to our message object
                msg.attach(file)
        return msg

    @staticmethod
    def ReadSearch_Message(key, value,IMAP):
        result,data = IMAP.search(None, key, '"{}"'.format(value))
        return data

    @staticmethod
    def Read_Body_FromMessage(Message):
        Service_Email_Helper_Class_Var = Service_Email_Helper_Class()
        if Message.is_multipart():
            return Service_Email_Helper_Class_Var.Read_Body_FromMessage(Message.get_payload(0))
        else:
            return Message.get_payload(None, True)

    @staticmethod
    def Read_Email_FromSearchResult(Search_Result,IMAP):
        msgs = []
        for num in Search_Result[0].split():
            typ, data = IMAP.fetch(num, '(RFC822)')
            msgs.append(data)
        return msgs
# endregion

# region Declare
# region Create
SMTPDetail_Dict ={
    "SMTP_Adr":"smtp.gmail.com",
    "SMTP_Port":587,
    "SMTP_Mail_Id":"abhisheksakti2014@gmail.com",
    "SMTP_Mail_Psw":"rrgzirgsqkgzheuw",
}

Service_Email_Helper_Class_Var = Service_Email_Helper_Class()

html = """<html>
    <head></head>
    <body>
        <p style='color:orange;'>This is my first HTML automated email!</p>
        <p>Alessio</p>
    </body>
</html>"""

EmailFromToDetail_Dict_List = [
    {
        "EmailFrom_Id":"abhisheksakti2014@gmail.com",
        "EmailTo_Id_List":["abhisheksakti2014@gmail.com"],
        "Email_Message":Service_Email_Helper_Class_Var.Create_Message("Good!",{"Content":"Hi there!","Content_Type":"Text"} ,None,None)
    },
    {
        "EmailFrom_Id":"abhisheksakti2014@gmail.com",
        "EmailTo_Id_List":["abhisheksakti2014@gmail.com"],
        "Email_Message":Service_Email_Helper_Class_Var.Create_Message("Good!",{"Content":html,"Content_Type":"Html"} ,None,None)
    }
]
# endregion
# region Read
IMAPDetail_Dict ={
    "IMAP_Adr":"imap.gmail.com",
    "IMAP_Mail_Id":"abhisheksakti2014@gmail.com",
    "IMAP_Mail_Psw":"rrgzirgsqkgzheuw",
}

EmailReadFrom_List =["abhisheksakti2014@gmail.com"]

EmailDelete_Dict_List = [
    {
        "Folder":"INBOX",
        "Filter": "FROM "+'"'+"googlealerts-noreply@google.com"+'"'        
    },
    {
        "Folder":"INBOX",
        "Filter": "FROM "+'"'+"birthdays@facebookmail.com"+'"' 
    }   
]
# endregion
# endregion

# region Test
Service_Email_Class_Var = Service_Email_Class()
# Service_Email_Class_Var.Create_Mail(SMTPDetail_Dict,EmailFromToDetail_Dict_List)
# Result = Service_Email_Class_Var.Read_Email(IMAPDetail_Dict,EmailReadFrom_List)
# print(Result)
# Service_Email_Class_Var.Delete_Email(IMAPDetail_Dict,EmailDelete_Dict_List)
Service_Email_Class_Var.Read_Email_FromCredentialGoogleCloudConsole("S:/2_PCRC/C22_PyDj_Sb_An/43_Prjct_Atmn_Email/A1_In/Credentials.json")
# endregion