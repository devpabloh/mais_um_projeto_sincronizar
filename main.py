# importando os arquivos que contém a classe GoogleCalendarSync
from google_calendar_sync import GoogleCalendarSync 
from outlook_calendar_sync import OutlookCalendarSync 

def main():
    # configurando o google calendar
    google_credentials_file = 'path/to/credentials.json'