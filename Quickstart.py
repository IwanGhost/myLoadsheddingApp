#import libraries
from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
#from Final_L import Grab_Data

import requests
import sys
from bs4 import BeautifulSoup
import ctypes  # An included library with Python install.
import re

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def main():
    def func3():
        print("Exiting")
        sys.exit()
    print("Running main frame")
    #ctypes to display gui
    def Mbox(title, text, style):
        return ctypes.windll.user32.MessageBoxW(0, text, title, style)

    #get url
    #url_1 = 'https://loadshed.org/zones/5576941750976512/loadshedding-eskom-co-za/stellenbosch'
    url = 'https://mydorpie.com/m/?page=loadshedding&suburb=Stellenbosch&region=Stellenbosch&province=Western-Cape'

    #set headers
    headers = {"User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36'}

    #identify headers and url
    page = requests.get(url, headers=headers)
    #page2 = requests.get(url_2, headers=headers)

    #parse websites
    soup = BeautifulSoup(page.content, 'html.parser')
    #soup2 = BeautifulSoup(page2.content, 'html.parser')

    #grabs data
    title = soup.find("div", {"id": "bh2yellow"}).get_text()

    #display grabbed data using gui
    #Mbox('Loadshedding Status', title1, 1)

    #check if there is loadshedding
    #if loadshedding then extract stage
    if title == "Loadshedding suspended until further notice":
        delete_all_events()
    if title != "Loadshedding suspended until further notice":
        phase = int(get_stage_num(title))
        print(phase)
        #grabs all tables
        table = soup.find_all('table')[0]
        
        #set variables
        inc=2
        stage=0
        num=1
        times = ""
    
        #some addition to get stage to the correct stage
        stage = stage + inc + phase

        #loop through table to extract times
        for sibling in soup.find_all('table')[1].tr.next_siblings:
            for td in sibling:
                #print(num)
                #check if number of stage on table match stage calculated
                if(num == stage):
                    times = td
                num = num+1

        #remove unneeded elements in string
        times = str(times)
        times = times.replace("<","").replace(">","").replace("t","").replace("d","").replace("c","").replace("l","").replace("a","").replace("s","").replace("=","").replace('"','').replace("l","").replace("m","").replace("n","").replace("f","").replace("b","").replace("r","").replace("/","").replace(" ","")
        times = str(times)
        #print(times)

        #set cut to 12 since every 12th caharcter indicated end of time
        n=12

        #cut up string
        new = [times[i:i+n] for i in range(0, len(times), n)]

        #remove unneeded elements from previous loop
        final_times = [i.strip() for i in new]

        #print out times underneath each other
        for t in final_times:
            print(t)

        #write times to text file
        with open('times.txt', 'w') as f:
            for item in final_times:
                f.write("%s\n" % item)
        f.close()
    
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
