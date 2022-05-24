#########################################################################
#
# Program : GmailBot  
#  
# Created : 05/09/2022
#   
# Programmer : Oscar Gastelum 
#  
# 
# ########################################################################  
#  Program description/ purpose : GmailBot aims to receive, read and also
# reply to emails. Can be configured to run specific events based on email 
# information that act as trigger events. GmailBot actively listens for 
# changes in the configured gmail accounts inbox. Every time a change
# occurs, GmailBot gets a callback. Using the push notification features
# from the GmailApi we get real time updates as changes occur. We then 
# extract the email received and display it. 
# GmailBot can be configured to act based on keywords from the emails 
# sender, subject or message. For example, if x sender, do y action... 
# if x keyword(s) in the message, extract values and insert them in a 
# database. If message subject "SPAM" and from x user, use the list of 
# emails from the emails from the emails message and spam all of them with 
# something. It's applications are endless and ultimately depend on the 
# developers imagination. 
#
#
#--------------------------- Change History ------------------------------
# Programmer : 
# Date :
# Description of change : 
#  
#########################################################################/ 

#imports for gmail api services 
from __future__ import print_function
import shutil
from tkinter import filedialog
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os.path
import pickle
import os.path
import base64
from bs4 import BeautifulSoup
from concurrent.futures import TimeoutError
from google.cloud import pubsub_v1

import pyinputplus as pyip #input validation
import ezgmail # send emails
import re # regex




#CONSTANTS
#TODO: add your own email for sendEmailReplyIfSender()
SEND_EMAIL = 'notARealEmail@mail.com'



#--------------------------- Functions  ------------------------------




"""Will get the x amount of emails from the inbox.
Args:
    numOfResults (int): number of emails to request. Set to 1 for this program
Returns:
    dictionary : dictionary with 'subject', 'from' and 'message' keys 
"""
def getEmails(numOfResults):

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
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
  
        # Save the access token in token.pickle file for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
  
    # Connect to the Gmail API
    service = build('gmail', 'v1', credentials=creds)
  
    # request a list of all the messages
    result = service.users().messages().list(userId='me').execute()
  
    # We can also pass maxResults to get any number of emails. Like this:
    result = service.users().messages().list(maxResults=numOfResults, userId='me').execute()
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
  

            #create and return a dictionary with the email elements
            return {'subject' : subject, 'from' : sender, 'message' : body}

        except:
            pass



"""send an email reply if the sender matches the email it received.
Args:
    email  (dictionary): Dictionary object retrun from getEmail()
    sender (string): if sender matches the email it received, send email reply.
    sub    (string): subject
    msg    (string): message for the email reply
"""
def sendEmailReplyIfSender(email, sender, sub, msg):
    try:
        if(re.search(sender, email['from'] )):
            ezgmail.send(sender, sub, msg)
            print('\nEmail reply sent to: %s.' % (SEND_EMAIL))
    except Exception as e:
        print(e)




"""Display the email received from the getEmail() function in the  
form of a dictionary.
Args:
    emailDict (dictionary): The email dictionary to display.
"""    
def displayEmail(emailDict):
    print('\n%s' % (''.center(50, '-')))
    print('Email Received : \n')
    print('From: %s'        % (emailDict['from']))
    print('Subject: %s'     % (emailDict['subject']))
    print('Message: \n%s\n%s%s' % ('~'.center(50, '~'), str(emailDict['message'])[4:-5], '~'.center(50, '~') ))




"""Actively listens for emails and performs triggered events when email contains
x elements. 
Args:
    project_id (string): the id for the project you created in google developer 
    page.
    subscription_id (string): subcrittion id found in developer page
    timeout (double): time before exiting listening method. Set to None to always
    listen.
    topicName (string): the name of the project found in topics tab of the 
    google developer projects page.
"""
def listenForEmails(project_id, subscription_id, timeout, topicName):
    subscriber = pubsub_v1.SubscriberClient()
    # The `subscription_path` method creates a fully qualified identifier
    # in the form `projects/{project_id}/subscriptions/{subscription_id}`
    subscription_path = subscriber.subscription_path(project_id, subscription_id)

    # callback is just for notifying you that msf was received 
    def callback(message: pubsub_v1.subscriber.message.Message) -> None:
        #acknowledges the message.
        message.ack()

        print('inside\n')
        #check email if callback from 
        e = getEmails(1)
        if(e != None):
            #display the email contents.
            displayEmail(e)

            #PERFORM THE SPECIFIED TASK GIVEN A TRIGGER KEY FROM EMAIL

            #if email matches keyword sender, reply to email
            subject = "OUT OF THE OFFICE"
            msg = "If you get this email I am on vacation and will not return until later this month."
            sendEmailReplyIfSender(e, SEND_EMAIL, subject, msg)
                    

    streaming_pull_future = subscriber.subscribe(subscription_path, 
    callback=callback)
    print(f"Listening for messages on {subscription_path}..\n")

    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://mail.google.com/']
    cred = Credentials.from_authorized_user_file('token.json', SCOPES)

    gmail = build('gmail', 'v1', credentials=cred)
    request = {'labelIds': ['INBOX'],'topicName': topicName}
    gmail.users().watch(userId='me', body=request).execute()

    # Wrap subscriber in a 'with' block to automatically call close() when done.
    with subscriber:
        try:
            # When `timeout` is not set, result() will block indefinitely,
            # unless an exception is encountered first.
            streaming_pull_future.result(timeout=timeout)
        except TimeoutError:
            streaming_pull_future.cancel()  # Trigger the shutdown.
            streaming_pull_future.result()  # Block until the shutdown is complete.



#TODO: rename file if is different name 
##initialize ezgmail, retrieve toke.json
def initializeEZGmail():
    credFile = 'credentials.json'
    
    if(not os.path.exists(credFile)):
        print(f'\nNo "{credFile}" file found in cwd directory.')
        a = pyip.inputYesNo(f'Would you like to specify its location?')
        
        if(a == 'yes'):
            print(f'\nPlease select your "{credFile}" file.')
            tempLoc = selectPath('f')
            shutil.move(tempLoc, credFile)

        else:
            print(f'\nCannot proceed without having {credFile} file available.')
            print('Terminating program. Exit code 1.')
            exit(1)

    #catch ezgmail exceptions
    try:
        ezgmail.init()
    except ezgmail.EZGmailException:
        print('ERROR: An error has occurred while initializing ezgmail.')

    

"""Prompt the user to select a directory or file path given the parameters "d"
or "f" respectively.
Args:
    type (char): select a directory or file path given the parameters "d"
or "f" respectively.
Returns:
    path: path to file or directory.
"""    
def selectPath(type):
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    if((type == 'f') or (type == 'F') ):
        p = filedialog.askopenfilename()
    elif((type == 'd')  or (type == 'D') ):
        p = filedialog.askdirectory()  # show an "Open" dialog box and return the path to the selected file
    else:
        print('\nInvlaid Parameter. Valid Parameters are "d" for directory or "f" for file path. Terminating program.')
        exit(1)
    return p



#all code commented out will be user for future version of GmailBot. Feel free to explore. 

#
#"""Display the email received from the getEmail() function in the  
#form of a dictionary.
#Args:
#    emailDict (dictionary): The email dictionary to display.
#"""    
#def displayEmail(emailDict):
#    print('\n%s' % (''.center(50, '-')))
#    print('Email Received : \n')
#    print('From: %s'        % (emailDict['sender']))
#    print('Subject: %s'     % (textToParagraph(emailDict['subject'])))
#    print('Message: \n%s\n%s\n%s' % ('~'.center(50, '~'), textToParagraph(emailDict['message']), '~'.center(50, '~') ))
#
#
#def textToParagraph (text):
#    return re.sub("(.{55})", "\\1\n", text, 0, re.DOTALL)
#	
#"""get the last email received to email address. Return dictionary
#containing:
#- sender 
#- recipient 
#- subject
#- message body
#- timestamp
#"""    
#def getLastEmail():
#    #get Gmailthread of recent emails 
#    recentEmails = ezgmail.recent(maxResults = 1)
#
#    #get the rest of the emails information
#    theEmail = recentEmails[0].messages[0]
#
#    #remove html brackets from message to get string text only 
#    msgString = BeautifulSoup(theEmail.body, features="lxml") #avoid errors by passing "features='lxml'"
#
#    #return a dictionary with the emails data
#    return{ 'sender'   : theEmail.sender   , 
#            'recipient': theEmail.recipient, 
#            'subject'  : theEmail.subject  , 
#            'message'  : ' '.join(msgString.get_text().split()), #remove long empty spaces           
#            'timestamp': theEmail.timestamp }
#



#"""write data to a dictionary object.
#Args:
#    dict (dictionary): previously saved dictionary object. 
#    key (dictionary key): key value to write to .
#    data (dictionary value): value to write to dictionary.
#Returns:
#    dictionary : new dictionary object with updated values.
#"""    
#def writeToDict(dict, key, data):
#    d = {}
#    d = dict
#    d[key] = data
#    return dict




#"""write dictionary object to a .txt file. 
#Args:
#    dict (dictionary): dictionary object to write.
#    file (.txt): .txt file to write to.
#"""    
#def writeDictToFile(dict, file):
#    # opening file in write mode (binary)
#    file = open(file, "wb")
#    
#    # serializing dictionary 
#    pickle.dump(dict, file)
#    
#    # closing the file
#    file.close()




#"""read dictionary data from file and return the saved dictrionary.
#Args:
#    file (.txt file with dictionary object): dictionary file saved object.
#Returns:
#    dictionary : saved dictionary object from file.
#"""
#def readDictFromFile(file):
#    try:
#
#        # reading the data from the file
#        with open(file, 'rb') as handle:
#            data = handle.read()
#            
#        # reconstructing the data as dictionary
#        d = pickle.loads(data)
#        return d
#
#    except Exception as e:
#        print(f'ERROR: {e}. Terminating program. System exit 1.')
#        exit(1)




#def readDataFile():
#
#    gmailBotData = {}
#    #if data file already exists, load data
#    print(f'\nSearching for {DATA_FILE} data file...')
#    if(os.path.exists(DATA_FILE)):
#        print(f'"{DATA_FILE} directory found.')
#        gmailBotData = readDictFromFile(DATA_FILE)
#    else:
#        gmailBotData = writeToDict(gmailBotData,
#                                  'gmailBotDir', 
#                                   os.getcwd())
#
#        writeDictToFile(gmailBotData, DATA_FILE)
#        print(f'"{DATA_FILE} directory created.')
#




#acts as the point of execution for any program.
def main():
    """ For listeningForEmails() """
    
    #TODO: DEVELOPER, add your own attributes for projectID, subscriptionId and topic name
    #have credential files in the same directory already set up and downloaded  
    
    #Setting the environment variable
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"]= 'Google_Application_Credentials.json'
    

    ##the project id name 
    #projectID = "your project id name "
    ##the subscription id name 
    #subscriptionID = "your subscriptionID name "
    ##Number of seconds the subscriber should listen for messages
    #timeout = None
    ##topic name 
    #topicName = 'your topic name '

    #initialize ezgmail to send emails
    initializeEZGmail()

    #start listening for emails sent to inbox 
    listenForEmails(projectID, subscriptionID, timeout, topicName)

    print('\n%s' % ('Program terminated successfully'.center(50, '*')))
#--------------------------- Code  -----------------------------------


#acts as the point of execution for any program.
main()


