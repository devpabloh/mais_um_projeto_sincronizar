# outlook_calendar_sync.py
import requests
import json
import msal
import webbrowser
from datetime import datetime, timedelta


class OutlookCalendarSync:
    def __init__(self, client_id, client_secret, tenant_id, redirect_uri=None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.redirect_uri = redirect_uri
        self.token = None
        self.calendar_id = None
        self.user_id = None
        self.authenticate()

    def authenticate(self):
        # Usando o fluxo de autenticação de dispositivo
        app = msal.PublicClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )

        # Tentar obter token do cache
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(
                ["https://graph.microsoft.com/.default"], account=accounts[0]
            )
            if result and "access_token" in result:
                self.token = result["access_token"]
                try:
                    self.user_id = self._get_user_id()
                    return  # Autenticação bem-sucedida com token em cache
                except Exception:
                    # Se falhar ao obter ID do usuário, o token pode estar inválido
                    pass

        # Se não houver token em cache ou o token estiver inválido, usar fluxo de dispositivo
        print("Iniciando autenticação com fluxo de dispositivo...")
        flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])

        if "user_code" not in flow:
            raise Exception(
                f"Falha ao iniciar fluxo de dispositivo: {flow.get('error')}"
            )

        print(f"Para autenticar, use o código: {flow['user_code']}")
        print(f"Acesse {flow['verification_uri']} e insira o código acima.")

        # Abrir o navegador automaticamente para facilitar
        webbrowser.open(flow["verification_uri"])

        # Aguardar a autenticação do usuário
        result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            self.token = result["access_token"]
            # Obter ID do usuário
            self.user_id = self._get_user_id()
        else:
            print(f"Erro de autenticação: {result.get('error')}")
            print(f"Descrição: {result.get('error_description')}")
            raise Exception("Falha ao obter token de acesso")

    def _get_user_id(self):
        """Obter ID do usuário autenticado"""
        url = "https://graph.microsoft.com/v1.0/me"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json().get("id")
        else:
            raise Exception(
                f"Erro ao obter ID do usuário: {response.status_code}, {response.text}"
            )

    def list_calendars(self):
        """Listar calendários disponíveis"""
        url = "https://graph.microsoft.com/v1.0/me/calendars"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            calendars = response.json().get("value", [])
            return [{"id": cal["id"], "name": cal["name"]} for cal in calendars]
        else:
            error_message = (
                f"Erro ao listar calendários: {response.status_code}, {response.text}"
            )
            print(error_message)
            raise Exception(error_message)

    def set_calendar_id(self, calendar_id):
        """Definir o ID do calendário a ser usado"""
        self.calendar_id = calendar_id

    def list_events(self, from_date=None):
        """Lista eventos a partir de uma data específica"""
        if not self.calendar_id:
            raise Exception(
                "ID do calendário não definido. Use set_calendar_id() primeiro."
            )

        # Se from_date não for fornecido, usar a data atual
        if from_date is None:
            from_date = datetime.now().replace(
                hour=0, minute=0, second=0, microsecond=0
            )

        url = f"https://graph.microsoft.com/v1.0/me/calendars/{self.calendar_id}/events"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

        # Converter a data para formato ISO 8601
        start_datetime = from_date.isoformat()

        # Filtrar eventos que começam após a data especificada
        params = {
            "$filter": f"start/dateTime ge '{start_datetime}'",
            "$orderby": "start/dateTime asc",
            "$top": 100,  # Limitar número de resultados para melhor performance
        }

        print(
            f"Buscando eventos do Outlook a partir de: {from_date.strftime('%d/%m/%Y')}"
        )

        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            events = response.json().get("value", [])
            print(f"Encontrados {len(events)} eventos no Outlook Calendar")

            # Debug: mostrar os eventos encontrados
            for event in events:
                print(
                    f"  - Outlook Event: {event.get('subject', 'Sem título')} (ID: {event.get('id', 'N/A')})"
                )

            return events
        else:
            # Adicionar mais detalhes sobre o erro
            error_message = f"Erro ao listar eventos na API do Outlook. Status: {response.status_code}, Resposta: {response.text}"
            print(error_message)  # Imprimir o erro para debug
            raise Exception(error_message)

    # Certifique-se de que o método create_event retorna o evento criado com seu ID
    def create_event(self, event_data):
        """Cria um evento no Outlook Calendar"""
        if not self.calendar_id:
            raise Exception(
                "ID do calendário não definido. Use set_calendar_id() primeiro."
            )

        url = f"https://graph.microsoft.com/v1.0/me/calendars/{self.calendar_id}/events"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

        response = requests.post(url, headers=headers, json=event_data)
        if response.status_code == 201:  # 201 Created
            created_event = response.json()
            print(
                f"Evento criado no Outlook Calendar: {created_event.get('subject', 'Sem título')}"
            )
            return created_event  # Retorna o evento criado com seu ID
        else:
            error_message = f"Erro ao criar evento na API do Outlook. Status: {response.status_code}, Resposta: {response.text}"
            print(error_message)
            raise Exception(error_message)

    def update_event(self, event_id, event):
        """Atualiza um evento existente"""
        url = f"https://graph.microsoft.com/v1.0/me/calendars/{self.calendar_id}/events/{event_id}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

        # Garantir a formatação correta para eventos de dia inteiro antes de enviar
        if event.get("isAllDay"):
            # Start time deve ser meia-noite UTC
            if "start" in event and "dateTime" in event["start"]:
                try:
                    start_date_str = event["start"]["dateTime"].split("T")[0]
                    event["start"]["dateTime"] = f"{start_date_str}T00:00:00"
                    event["start"]["timeZone"] = "UTC"
                except Exception as e:
                    print(f"Erro ao processar start dateTime para evento all-day: {e}")
                    # Pode ser necessário um tratamento de erro mais robusto aqui

            # End time deve ser meia-noite UTC do dia seguinte
            if "end" in event and "dateTime" in event["end"]:
                try:
                    # Tenta obter a data de início para calcular o fim corretamente
                    current_start_date_str = ""
                    if "start" in event and "dateTime" in event["start"]:
                         current_start_date_str = event["start"]["dateTime"].split("T")[0]

                    if current_start_date_str:
                         start_dt_obj = datetime.fromisoformat(current_start_date_str)
                         # O fim de um evento de dia inteiro no Outlook é a meia-noite do dia em que termina,
                         # que é efetivamente o início do dia seguinte ao start para um evento de um dia.
                         # Se o evento do Google/Expresso já tem um 'end.date', usamos ele,
                         # caso contrário, calculamos como start + 1 dia.
                         
                         # Verifica se o 'end.dateTime' original já representa o dia correto
                         original_end_date_str = event["end"]["dateTime"].split("T")[0]
                         original_end_dt_obj = datetime.fromisoformat(original_end_date_str)

                         # Se o 'end date' original não for pelo menos um dia depois do 'start date', ajusta.
                         # Para um evento de um único dia inteiro, o 'end.date' do Google é o dia seguinte.
                         # Outlook espera que o 'end.dateTime' para um evento de dia inteiro seja 00:00 do dia seguinte ao término visual.
                         # Ex: Evento de dia inteiro em 2023-10-10.
                         # Google: start.date=2023-10-10, end.date=2023-10-11
                         # Outlook: start.dateTime=2023-10-10T00:00:00Z, end.dateTime=2023-10-11T00:00:00Z

                        # Usa a data de término já formatada se ela for posterior à data de início
                         if original_end_dt_obj > start_dt_obj:
                            event["end"]["dateTime"] = f"{original_end_date_str}T00:00:00"
                         else: # Garante que o fim seja no mínimo o dia seguinte ao início
                            end_dt_obj = start_dt_obj + timedelta(days=1)
                            event["end"]["dateTime"] = f"{end_dt_obj.strftime('%Y-%m-%d')}T00:00:00"
                         event["end"]["timeZone"] = "UTC"
                    else:
                        # Fallback se a data de início não puder ser determinada
                         end_date_str = event["end"]["dateTime"].split("T")[0]
                         event["end"]["dateTime"] = f"{end_date_str}T00:00:00" # Poderia ser impreciso
                         event["end"]["timeZone"] = "UTC"

                except Exception as e:
                    print(f"Erro ao processar end dateTime para evento all-day: {e}")

        print(f"Outlook Update Payload: {json.dumps(event, indent=2)}") # Log para depuração
        response = requests.patch(url, headers=headers, data=json.dumps(event))
        if response.status_code in (200, 201, 204):
            return response.json() if response.text else {}
        else:
            error_message = f"Erro ao atualizar evento na API do Outlook. Status: {response.status_code}, Resposta: {response.text}"
            print(error_message)
            raise Exception(error_message)

    def delete_event(self, event_id):
        """Exclui um evento"""
        url = f"https://graph.microsoft.com/v1.0/me/calendars/{self.calendar_id}/events/{event_id}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }
        response = requests.delete(url, headers=headers)
        if response.status_code in (200, 201, 204):
            return True
        else:
            error_message = f"Erro ao excluir evento na API do Outlook. Status: {response.status_code}, Resposta: {response.text}"
            print(error_message)
            raise Exception(error_message)
