#myLoadsheddingApp
#import libraries
from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import requests
import json
from bs4 import BeautifulSoup
import time
import schedule
import time
import os
import sys
import openpyxl
import datetime
import random

SCOPES = ['https://www.googleapis.com/auth/calendar']


# global stage, to test for repeating stages
prev_status = 0
prev_date = str(datetime.date.today())
current_date = ''

# global var to check whether times on stages have passed
old_time = True

# loading in excel file with loadshedding times
wrkbk = openpyxl.load_workbook("Loadshedding.xlsx")

# global var to hold stage num
stage_num = 0

# path to root, Not permanent fix, needs better solution - check dir for .txt file
url = 'C:\\Users\Ghost\Documents'

#populating user_headers with headers from .txt file
with open('user-agents.txt') as f:
     user_headers = f.readlines()

# Closes app
def End():
    print("Exiting Program")
    sys.exit()

# Return time in HH:MM:SS - 17:45:17
def checkTime():

    now = datetime.datetime.now()

    current_time = now.strftime("%H:%M:%S")
    return current_time

# Returns date 2022-09-06 #YYYY, MM, DD
def currDate():
    currentDate = str(datetime.date.today())
    return currentDate

# Returns day - 09 - for 9th
def checkDay():

    currentDate = currDate()

    res = currentDate[8:]
    if res[0] == '0':
        res = res[1 : : ]
        day = int(res)
        return day
    else:
        day = int(res)
        return day

# checks if scheduled time have passed or not
def checkOldTime(start_time):

    global old_time

    start_time = start_time
    current_time = checkTime()
    if start_time[0:5] >= current_time[0:5]:
        old_time = False
        return old_time
    else:
        return old_time

# Returns stage via API - exp, 3
def getLoadStatus():

    global user_headers

    user_agent = random.choice(user_headers)
    headers = {'User-Agent': user_agent.replace("\n", "")}

    stage = requests.get('https://loadshedding.eskom.co.za/LoadShedding/GetStatus', headers=headers).text

    return stage

# Calls this function every 5 minutes
def callStatus():

    global stage_num
    global current_date
    global prev_status
    global prev_date

    current_date = currDate() #2022-09-06

    # check if day changed, then reset prev_status otherwize events won't be added, due to if check
    if current_date > prev_date:
        prev_date = current_date
        prev_status = 0

    status = int(getLoadStatus()) #3

    if status != 99 and status != -1:

        stage_num = status - 1 #0 -> 3 ->  2

        print('Stage currently is: Stage ' + str(stage_num))

        day = checkDay() #7
            
        mainSys(status, day) # ( 3, 7)

        print('End Of Call')
    else:
        print('API is broken, App will be stopped')
        End()

# Function to simplify code and to do all logic
def mainSysCall(day, service):

    global wrkbk
    global stage_num
    global old_time

    #prev_status = prev_status
    #status = "stage 1"
    #compare status

    ws = wrkbk['Stage_' + str(stage_num)]

    #loop through excel
    for row in ws.iter_rows(): #loops through sheet with specific stage name
        val = row[(day + 1)].value # sets val  = row , but row number is = day + 1
        if val == stage_num:
            start_time = str(row[0].value)
            end_time = str(row[1].value)
            start_date = str(currDate())
            end_date = str(currDate())
            print(f'{start_time} {end_time}')
            #2022-09-09 #YYYY, MM, DD
            # if end_time = 00:30:00
            if end_time[0:5] == '00:30':
                temp_end_date = end_date[8:]# getting date
                temp_end_date = str(int(temp_end_date)+1) # converting to int and adding 1 then back to str
                end_date = con_date[0:8] + temp_end_date
                #call calander function with time information and add to calander
                createCalEvents(start_time, end_time, start_date, end_date, service)
            else:
                checkOldTime(start_time)
                if old_time == False:
                    createCalEvents(start_time, end_time, start_date, end_date, service)
                else:
                    print('Old Time')

# Main sys
def mainSys(status, day):

    global prev_status

    if status != 1: # 1 = not loadshedding
        if status == 2 and prev_status != status: # stage 1
            prev_status = status
            service, now = eventSetup()
            deleteEvents(service, now)
            mainSysCall(day, service)
            checkEvents(service, now)

        elif status == 3 and prev_status != status: # stage 2
            prev_status = status
            service, now = eventSetup()
            deleteEvents(service, now)
            mainSysCall(day, service)
            checkEvents(service, now)

        elif status == 4 and prev_status != status: # stage 3
            prev_status = status
            service, now = eventSetup()
            deleteEvents(service, now)
            mainSysCall(day, service)
            checkEvents(service, now)

        elif status == 5 and prev_status != status: # stage 4
            prev_status = status
            service, now = eventSetup()
            deleteEvents(service, now)
            mainSysCall(day, service)
            checkEvents(service, now)

        elif status == 6 and prev_status != status: # stage 5
            prev_status = status
            service, now = eventSetup()
            deleteEvents(service, now)
            mainSysCall(day, service)
            checkEvents(service, now)

        elif status == 7 and prev_status != status: # stage 6
            prev_status = status
            service, now = eventSetup()
            deleteEvents(service, now)
            mainSysCall(day, service)
            checkEvents(service, now)

        else:
            if prev_status == status:
                print('Stage Already Implimented')
            else:
                print('No stage found')
    else:
        if status == 1:
            print('No Loadshedding!')
        else:
            print('Something Went Wrong!')

# function to create events
def eventSetup():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    return service, now

def createCalEvents(start_time, end_time, start_date, end_date, service):

    event = {
      'summary': 'LOADSHEDDING',
      'location': 'Home',
      'description': 'Loadshedding APP',
      'start': {
        'dateTime': start_date + 'T' + start_time,
        'timeZone': 'Africa/Windhoek',
      },
      'end': {
        'dateTime': end_date + 'T' + end_time,
        'timeZone': 'Africa/Windhoek',
      },
      'recurrence': [
        'RRULE:FREQ=DAILY;COUNT=1'
      ],
      'reminders': {
        'useDefault': False,
        'overrides': [
          {'method': 'email', 'minutes': 0.5 * 60},
          {'method': 'popup', 'minutes': 5},
        ],
      },
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    #print ('Event created: %s' % (event.get('htmlLink')))
    print ('Event created')

def checkEvents(service, now):
    
    print('Getting the upcoming 5 events')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=5, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])

def deleteEvents(service, now):
    
    print('Deleting Events!')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=20, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found to be deleted')
    for event in events:
        if event['description'] == 'Loadshedding APP':
            print('Found Event!')
            check_event_id = event['id']
            service.events().delete(calendarId='primary', eventId=check_event_id).execute()
            print('Done deleting!')

# Main function, being initialized by __name__ == '__main__'
def main():
    schedule.every(0.2666666).minutes.do(callStatus)

    while True:
            schedule.run_pending()
            for fname in os.listdir(url):
                    if fname.endswith('.txt'):
                            time.sleep(0.3)
                            End()
            time.sleep(1)

if __name__ == '__main__':
    main()