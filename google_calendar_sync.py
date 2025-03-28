import os
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

class GoogleCalendarSync:
    # Define SCOPES as a class attribute
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    
    def __init__(self, credentials_file, token_file, calendar_id='primary'):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.calendar_id = calendar_id
        self.service = self.authenticate()

    def authenticate(self):
        creds = None
        # O arquivo token.json armazena os tokens de acesso e atualização do usuário
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_info(json.loads(open(self.token_file).read()), self.SCOPES)
        
        # Se não houver credenciais (válidas) disponíveis, permita que o usuário faça login.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, self.SCOPES)
                creds = flow.run_local_server(port=8090)
            
            # Salve as credenciais para a próxima execução
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())
        
        # Construa o serviço
        return build('calendar', 'v3', credentials=creds)

    def list_events(self):
        # Exemplo de listagem de eventos
        events_result = self.service.events().list(
            calendarId=self.calendar_id,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])
        return events

    def create_event(self, event):
        # Cria um novo evento
        created_event = self.service.events().insert(calendarId=self.calendar_id, body=event).execute()
        return created_event
