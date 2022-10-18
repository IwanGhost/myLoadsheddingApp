# myLoadsheddingApp
Just a small project to track loadshedding and set reminders on your calendar.

## Note:
This app is made for windows only, and only works for Stellenbosch zone 2, currently.

## (WIP)How to use

- Create a folder on your PC where you would like to run everything from
- Go to: https://developers.google.com/calendar/api/quickstart/python
- Follow the steps given on the website, namely:
  - Enable the Google Calendar API
  - Going to credentials and adding a OAuth client ID
  - Make sure it is set to desktop app
  - Make sure to download the credentials.json file to the folder you created and rename it to credentials.json if it is not already so
- Make sure to install the google calendar API libraries: 
  ```bash
  pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
  ```
- Download and run [Quickstart_Setup.py](https://github.com/IwanSmit/myLoadsheddingApp/blob/main/Quickstart_Setup.py) from within the folder you created
  - Make sure the credentials.json file is also within the folder you created
  - This will make sure you get a token.pickle file, also just make sure you have one, after running it, and make sure the file name is token.pickle
- Download the Quickstart.py, Loadsheding.xlsx and user-agents.txt files into the folder you created
- Now just make sure you have a internet connection and run Quickstart.py

## Disclaimer
myLoadsheddingApp is provided by Iwan Smit "as is" and "with all faults". The provider makes no representations or warranties of any kind concerning the safety, suitability, lack of viruses, inaccuracies, typographical errors, or other harmful components of this software. There are inherent dangers in the use of any software, and you are solely responsible for determining whether myLoadsheddingApp is compatible with your equipment and other software installed on your equipment. You are also solely responsible for the protection of your equipment and backup of your data, and the provider will not be liable for any damages you may suffer in connection with using, modifying, or distributing this software.
