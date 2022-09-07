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
import keyboard
import schedule
import time
import os
import sys
import ctypes  # An included library with Python install.
import re
import openpyxl

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']


#global stage, to test for repeating events
prev_status = ""
# need to add day/time varible to check for same day

#loading in excel file with loadshedding times
wrkbk = openpyxl.load_workbook("Loadshedding.xlsx")

stage_num = 0

currentDate = str(datetime.date.today())

if_pass = False

#path to root #Not permanent fix, needs better solution
url = 'C:\\Users\Ghost\Documents'

def End():
    print("Exiting Program")
    sys.exit()


def currDate():
    currentDate = str(datetime.date.today())
    #returns in 2022-09-06 #YYYY, MM, DD
    return currentDate

def check_day():

    currentDate = currDate()

    res = currentDate[8:]
    if res[0] == '0':
        res = res[1 : : ]
        day = int(res)
        return day
    else:
        day = int(res)
        return day



def get_load_status():

    res = requests.get('https://loadshedding.eskom.co.za/LoadShedding/GetStatus')

    stage = res.text[:1]


    return stage

def call_status():

    global stage_num

    status = get_load_status() #3

    tmp_int = int(status) # 3

    stage_num = tmp_int - 1 #0 -> 3 ->  2

    day = check_day() #7

    main(status, day) # ( 3, 7)

    print('End')


def main(status, day):

    status = status
    day = day
    global prev_status
    global wrkbk
    global stage_num

    #prev_status = prev_status
    #status = "stage 1"
    #compare status

    ws = wrkbk['Stage_' + str(stage_num)]

    if status != '1': # 1 = not loadshedding
        if status == '2' and prev_status != status: # stage 1
            prev_status = status
            #call time function
            #call calander function with time information and add to calander
            print(f"End Of Add {status}")
        elif status == '3' and prev_status != status: # stage 2
            prev_status = status
            #loop through excel
            for row in ws.iter_rows():
                val = row[(day + 1)].value
                if val == stage_num:
                    start_time = row[0].value
                    end_time = row[1].value  # Need to have global variables 
                    print(f'{start_time} {end_time}')
            #call calander function with time information and add to calander
            print(f"End Of Add {status}")
        elif status == '4' and prev_status != status: # stage 3
            prev_status = status
            #call stage decoder
            #call time function
            #call calander function with time information and add to calander
            print(f"End Of Add {status}")
        elif status == '5' and prev_status != status: # stage 4
            prev_status = status
            #call stage decoder
            #call time function
            #call calander function with time information and add to calander
            print(f"End Of Add {status}")
        elif status == '6' and prev_status != status: # stage 5
            prev_status = status
            #call stage decoder
            #call time function
            #call calander function with time information and add to calander
            print(f"End Of Add {status}")
        elif status == '7' and prev_status != status: # stage 6
            prev_status = status
            #call stage decoder
            #call time function
            #call calander function with time information and add to calander
            print(f"End Of Add {status}")
        else:
            print("No stage found! or Stage already implimented!")
    else:
        print("Something went wrong or No Loadshedding!")
    # check status, if loadshed, then get times for stage, and add to calander, if not, then wait 5 min, and check again
    
##        """Shows basic usage of the Google Calendar API.
##        Prints the start and name of the next 10 events on the user's calendar.
##        """
##        creds = None
##        # The file token.pickle stores the user's access and refresh tokens, and is
##        # created automatically when the authorization flow completes for the first
##        # time.
##        if os.path.exists('token.pickle'):
##            with open('token.pickle', 'rb') as token:
##                creds = pickle.load(token)
##        # If there are no (valid) credentials available, let the user log in.
##        if not creds or not creds.valid:
##            if creds and creds.expired and creds.refresh_token:
##                creds.refresh(Request())
##            else:
##                flow = InstalledAppFlow.from_client_secrets_file(
##                    'credentials.json', SCOPES)
##                creds = flow.run_local_server(port=0)
##            # Save the credentials for the next run
##            with open('token.pickle', 'wb') as token:
##                pickle.dump(creds, token)
##
##        service = build('calendar', 'v3', credentials=creds)
##
##        # Call the Calendar API
##        now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
##        ###

        service = ""
        now = ""
        service, now =  get_service_cal(service, now)
        print(service)
        currentDate = str(datetime.date.today())
        
##        #Callable event to create multiple entries
##        def createEvent(begin_1, end_1, currentDate, endDate, stage):
##            
##            begin_1 = begin_1
##            end_1 = end_1
##            endDate = endDate
##            currentDate = currentDate
##            stage = stage
##            #newID = newID
##            
##            event = {
##              'summary': stage,
##              #'id': 'testy1',
##              'location': 'Home',
##              'description': 'Loadshedding App.',
##              'start': {
##                'dateTime': currentDate + 'T' + begin_1,
##                'timeZone': 'Africa/Windhoek',
##              },
##              'end': {
##                'dateTime': endDate + 'T' + end_1,
##                'timeZone': 'Africa/Windhoek',
##                #'id' : 'testy1',
##              },
##              'recurrence': [
##                'RRULE:FREQ=DAILY;COUNT=1'
##              ],
##              'reminders': {
##                'useDefault': False,
##                'overrides': [
##                  {'method': 'email', 'minutes': 0.5 * 60},
##                  {'method': 'popup', 'minutes': 10},
##                ],
##              },
##            }
##
##            event = service.events().insert(calendarId='primary', body=event).execute()
##            print ('Event created: %s' % (event.get('htmlLink')))
            
        



        ######
        # Using readlines()
        file1 = open('times.txt', 'r')
        Lines = file1.readlines()
 
        count = 0
        time1 = ""
        time2 = ""
        time3 = ""
        # Strips the newline character
        for line in Lines:
            if (count == 0):
                time1 = line.strip()
            if (count == 1):
                time2 = line.strip()
            if (count == 2):
                time3 = line.strip()
            count += 1

        #######

        print('Checking availability')
        events_result = service.events().list(calendarId='primary', timeMin=now,
                                            maxResults=10, singleEvents=True,
                                            orderBy='updated').execute() #startTime
        events = events_result.get('items', [])
        #print(events)
        space = 0
        if not events:
            space = 1
        
        ########
        if (time1 != "" and space == 1):
            n = 6
            new = [time1[i:i+n] for i in range(0, len(time1), n)]
            final_time1 = [i.strip() for i in new]
            with open('time1.txt', 'w') as f:
                for item in final_time1:
                    f.write("%s\n" % item)
        
            file1 = open('time1.txt', 'r')
            Lines = file1.readlines()
            count = 0
            begin_1 = ""
            end_1 = ""
            for line in Lines:
                if (count == 0):
                    line = line.replace('-', '')
                    begin_1 = line.strip()
                    #print(begin_1)
                if(count == 1):
                    end_1 = line.strip()
                    #print(end_1)
                count += 1
            begin_1 = begin_1 + ":00"
            end_1 = end_1 + ":00"
            print(begin_1)
            print(end_1)
            testingEnd = end_1[0:2]
            if(testingEnd == "00"):
                date = str(datetime.date.today() + datetime.timedelta(days=1))
                endDate = date
            time3id = "time1"
            if((testingEnd == "00") == False):
                endDate = currentDate
            stage = "Time 1 LOADSHEDDING!"
            createEvent(begin_1, end_1, currentDate, endDate, stage, service)
        if (time2 != "" and space == 1):
            n = 6
            new = [time2[i:i+n] for i in range(0, len(time2), n)]
            final_time2 = [i.strip() for i in new]
            with open('time2.txt', 'w') as f:
                for item in final_time2:
                    f.write("%s\n" % item)
        
            file1 = open('time2.txt', 'r')
            Lines = file1.readlines()
            count = 0
            begin_1 = ""
            end_1 = ""
            for line in Lines:
                if (count == 0):
                    line = line.replace('-', '')
                    begin_1 = line.strip()
                    #print(begin_1)
                if(count == 1):
                    end_1 = line.strip()
                    #print(end_1)
                count += 1
            begin_1 = begin_1 + ":00"
            end_1 = end_1 + ":00"
            print(begin_1)
            print(end_1)
            testingEnd = end_1[0:2]
            if(testingEnd == "00"):
                date = str(datetime.date.today() + datetime.timedelta(days=1))
                endDate = date
            time3id = "time2"
            if((testingEnd == "00") == False):
                endDate = currentDate
            stage = "Time 2 LOADSHEDDING!"
            createEvent(begin_1, end_1, currentDate, endDate, stage, service)
        if (time3 != "" and space == 1):
            n = 6
            new = [time3[i:i+n] for i in range(0, len(time3), n)]
            final_time3 = [i.strip() for i in new]
            with open('time3.txt', 'w') as f:
                for item in final_time3:
                    f.write("%s\n" % item)
        
            file1 = open('time3.txt', 'r')
            Lines = file1.readlines()
            count = 0
            begin_1 = ""
            end_1 = ""
            for line in Lines:
                if (count == 0):
                    line = line.replace('-', '')
                    begin_1 = line.strip()
                    #print(begin_1)
                if(count == 1):
                    end_1 = line.strip()
                    #print(end_1)
                count += 1
            begin_1 = begin_1 + ":00"
            end_1 = end_1 + ":00"
            print(begin_1)
            print(end_1)
            testingEnd = end_1[0:2]
            if(testingEnd == "00"):
                date = str(datetime.date.today() + datetime.timedelta(days=1))
                endDate = date
            time3id = "time3"
            if((testingEnd == "00") == False):
                endDate = currentDate
            stage = "Time 3 LOADSHEDDING!"
            createEvent(begin_1, end_1, currentDate, endDate, stage, service)


            
        #createEvent(time1, time2, currentDate)

        ####
        
        print('Getting the upcoming 10 events')
        events_result = service.events().list(calendarId='primary', timeMin=now,
                                            maxResults=10, singleEvents=True,
                                            orderBy='updated').execute() #startTime
        events = events_result.get('items', [])
        #print(events)
        if not events:
            print('No upcoming events found.')

        #if not empty
        
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            check_id = event['id']
            print(check_id)
        func3()

        #this part is unneeded i think, already did something at start of code
        bool1 = (title != "Loadshedding suspended until further notice")
        if (space == 0 and bool1 == False):
            for event in events:
                service.events().delete(calendarId='primary', eventId=event['id']).execute()

                #for events in events, if name = summary, delete
        



def get_stage_num(word):
    print("Getting stage num")

    #get stage name: Stage 2
    title = word[:7]
    phase = title[-1:]
    return phase

def get_service_cal(service, now):
    
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

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    ###

    return service, now

#Callable event to create multiple entries
def createEvent(begin_1, end_1, currentDate, endDate, stage, service):
    
    begin_1 = begin_1
    end_1 = end_1
    endDate = endDate
    currentDate = currentDate
    stage = stage
    service = service
    #newID = newID
    
    event = {
      'summary': stage,
      #'id': 'testy1',
      'location': 'Home',
      'description': 'Loadshedding App.',
      'start': {
        'dateTime': currentDate + 'T' + begin_1,
        'timeZone': 'Africa/Windhoek',
      },
      'end': {
        'dateTime': endDate + 'T' + end_1,
        'timeZone': 'Africa/Windhoek',
        #'id' : 'testy1',
      },
      'recurrence': [
        'RRULE:FREQ=DAILY;COUNT=1'
      ],
      'reminders': {
        'useDefault': False,
        'overrides': [
          {'method': 'email', 'minutes': 0.5 * 60},
          {'method': 'popup', 'minutes': 10},
        ],
      },
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print ('Event created: %s' % (event.get('htmlLink')))

def delete_all_events():
    print("Deleted")



if __name__ == '__main__':
    main()
