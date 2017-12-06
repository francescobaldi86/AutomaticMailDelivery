import httplib2
import os
import io
import csv
import datetime
import calendar
import logging

from apiclient.discovery import build
from apiclient.http import MediaIoBaseDownload
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import win32com.client as win32

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Drive API Python Quickstart'
dir_path = os.path.dirname(os.path.realpath(__file__))


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'drive-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def download_file(service):
    results = service.files().list(
        pageSize=10, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])
    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            if item['name'] == 'IPESE Afterlunch Seminars Planning':
                fileId = item['id']
    return (items, fileId)


def read_file(service, fileId):
    request = service.files().export_media(fileId=fileId, mimeType='text/csv')
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download " + str(int(status.progress() * 100)) + '%')
    csvContent = fh.getvalue().decode("utf-8")
    f = io.StringIO(csvContent)
    reader = csv.reader(f, delimiter=',')
    output = {}
    output["Dates"] = []
    output["Presenter Names"] = []
    output["Presentation Titles"] = []
    for row in reader:
        if row[1][:2] == '20':
            output["Presenter Names"].append(row[0])
            output["Dates"].append(row[1])
            output["Presentation Titles"].append(row[2])
    return output


def find_date(orderedList):
    tomorrowDate =  str(datetime.date.today() + datetime.timedelta(days=1))
    sendMailDecision = False
    counter = 0
    idEvent = ''
    for presentationDate in orderedList["Dates"]:
        if presentationDate ==tomorrowDate:
            sendMailDecision = True
            idEvent = counter
        counter = counter + 1
    if sendMailDecision:
        with open('mail_sent.csv', 'r') as csvfile:
            mailSentCheck = csv.reader(csvfile)
            for row in mailSentCheck:
                if len(row) == 2:
                    if (row[0] == tomorrowDate) & (row[1] == 'sent'):
                        sendMailDecision = False
                    else:
                        continue
    if sendMailDecision:
        with open('mail_sent.csv', 'w') as csvfile:
            csvwriter = csv.writer(csvfile, delimiter=',')
            csvwriter.writerow([tomorrowDate, 'sent'])
    return (sendMailDecision, idEvent)

def ord(n):
    return str(n)+("th" if 4<=n%100<=20 else {1:"st",2:"nd",3:"rd"}.get(n%10, "th"))


def write_mail(orderedList, idEvent):
    # This
    presenterName = orderedList["Presenter Names"][idEvent]
    date = orderedList["Dates"][idEvent]
    day = ord(int(date[-2:]))
    monthN = int(date[-5:-3])
    month = calendar.month_name[monthN]
    presentationTitle = orderedList["Presentation Titles"][idEvent]
    mailBody = "Hi everyone, \n it’s time again for an afterlunch IPESE seminar. It will take place tomorrow, the "
    mailBody = mailBody + day + " of " + month
    mailBody = mailBody + " at 13 in Emosson, as usual. Presenting this time will be "
    mailBody = mailBody + presenterName
    mailBody = mailBody + " who will talk about "
    mailBody = mailBody + presentationTitle + ".\n\n"
    mailBody = mailBody + "Have a nice week to everyone \n\n"
    mailBody = mailBody + "PhD Francesco Baldi \n"
    mailBody = mailBody + "Postdoctoral Researcher \n"
    mailBody = mailBody + "Industrial Processes and Energy Systems engineering (IPESE) \n"
    mailBody = mailBody + "École Polytechnique Fédérale de Lausanne (EPFL) \n"
    mailBody = mailBody + "Tel: +41 21 69 58277 \n"
    mailBody = mailBody + "Mail: Francesco.baldi@epfl.ch \n"
    return mailBody


def send_mail(destinationAddress, mailSubject, mailBody):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = destinationAddress
    mail.Subject = mailSubject
    mail.Body = mailBody
    # mail.HTMLBody = '<h2>HTML Message body</h2>'  # this field is optional
    mail.Send()

def main():
    logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', filename='log.log', level=logging.DEBUG)
    logging.info('The program successfully started')
    # First some info about the email to write
    # destinationAddress = "ipese@groupes.epfl.ch"
    # destinationAddress = "stefano.moret@epfl.ch"
    destinationAddress = "francesco.baldi@epfl.ch"
    mailSubject = "Next IPESE After-lunch seminar - AUTOMATIC MAIL TEST"
    # Getting the required credentials
    credentials = get_credentials()
    logging.info('get_credentials executed successfully')
    http = credentials.authorize(httplib2.Http())
    service = build('drive', 'v3', http=http)
    # Downloading the file
    (items, fileId) = download_file(service)
    logging.info('Destination file downloaded successfully')
    orderedList = read_file(service, fileId)
    (sendMailDecision, idEvent) = find_date(orderedList)
    logging.info('Date find completed, with decision: %s', str(sendMailDecision))
    if sendMailDecision:
        logging.info('There will be a seminar tomorrow')
        mailBody = write_mail(orderedList, idEvent)
        send_mail(destinationAddress, mailSubject, mailBody)
        logging.info('Email sent')
    else:
        logging.info("No seminar tomorrow. No email has been sent")





if __name__ == '__main__':
    main()