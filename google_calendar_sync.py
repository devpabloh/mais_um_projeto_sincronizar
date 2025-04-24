import os
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import os.path
from googleapiclient.errors import HttpError


class GoogleCalendarSync:
    # Define SCOPES as a class attribute
    SCOPES = ["https://www.googleapis.com/auth/calendar"]

    def __init__(self, credentials_file, token_file, calendar_id="primary"):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.calendar_id = calendar_id
        self.service = self.authenticate()

    def authenticate(self):
        creds = None
        # O arquivo token.json armazena os tokens de acesso e atualização do usuário
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_info(
                json.loads(open(self.token_file).read()), self.SCOPES
            )

        # Se não houver credenciais (válidas) disponíveis, permita que o usuário faça login.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, self.SCOPES
                )
                creds = flow.run_local_server(port=8090)

            # Salve as credenciais para a próxima execução
            with open(self.token_file, "w") as token:
                token.write(creds.to_json())

        # Construa o serviço
        return build("calendar", "v3", credentials=creds)

    def list_events(self, from_date=None):
        """Lista eventos a partir de uma data específica"""
        # Se from_date não for fornecido, usar a data atual
        if from_date is None:
            from_date = datetime.now().replace(
                hour=0, minute=0, second=0, microsecond=0
            )

        # Converter para formato ISO
        time_min = from_date.isoformat() + "Z"  # 'Z' indica UTC

        print(
            f"Buscando eventos do Google a partir de: {from_date.strftime('%d/%m/%Y')}"
        )

        # Listar eventos
        events_result = (
            self.service.events()
            .list(
                calendarId=self.calendar_id,
                timeMin=time_min,
                maxResults=100,  # Aumentar para pegar mais eventos
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )

        events = events_result.get("items", [])
        print(f"Encontrados {len(events)} eventos no Google Calendar")

        # Debug: mostrar os eventos encontrados
        for event in events:
            print(
                f"  - Google Event: {event.get('summary', 'Sem título')} (ID: {event.get('id', 'N/A')})"
            )

        return events

    # Certifique-se de que o método create_event retorna o evento criado com seu ID
    def create_event(self, event_data):
        """Cria um evento no Google Calendar"""
        try:
            event = (
                self.service.events()
                .insert(calendarId=self.calendar_id, body=event_data)
                .execute()
            )
            print(
                f"Evento criado no Google Calendar: {event.get('summary', 'Sem título')}"
            )
            return event  # Retorna o evento criado com seu ID
        except Exception as e:
            print(f"Erro ao criar evento no Google Calendar: {e}")
            raise e

    def update_event(self, event_id, event):
        """Atualiza um evento existente"""
        updated_event = (
            self.service.events()
            .update(calendarId=self.calendar_id, eventId=event_id, body=event)
            .execute()
        )
        return updated_event

    def delete_event(self, event_id):
        """Exclui um evento"""
        try:
            print(f"Excluindo do Google: {event_id}")
            self.service.events().delete(
                calendarId="primary", eventId=event_id
            ).execute()
            print(f"Evento {event_id} excluído com sucesso do Google Calendar")
            return True
        except HttpError as error:
            # Verificar se o erro é 410 (Gone) - evento já foi excluído
            if error.resp.status == 410:
                print(f"Evento {event_id} já foi excluído do Google Calendar")
                return True  # Considerar como sucesso, pois o evento não existe mais
            else:
                print(f"Erro ao excluir evento do Google: {error}")
                return False
        except Exception as e:
            print(f"Erro ao excluir evento do Google: {e}")
            return False
