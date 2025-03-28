# outlook_calendar_sync.py
import requests
import json
import msal
import webbrowser

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
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
        # Tentar obter token do cache
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=accounts[0])
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
            raise Exception(f"Falha ao iniciar fluxo de dispositivo: {flow.get('error')}")
        
        print(f"Para autenticar, use o código: {flow['user_code']}")
        print(f"Acesse {flow['verification_uri']} e insira o código acima.")
        
        # Abrir o navegador automaticamente para facilitar
        webbrowser.open(flow['verification_uri'])
        
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
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json().get('id')
        else:
            raise Exception(f"Erro ao obter ID do usuário: {response.status_code}, {response.text}")

    def list_calendars(self):
        """Listar calendários disponíveis"""
        url = "https://graph.microsoft.com/v1.0/me/calendars"
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            calendars = response.json().get('value', [])
            return [{'id': cal['id'], 'name': cal['name']} for cal in calendars]
        else:
            error_message = f"Erro ao listar calendários: {response.status_code}, {response.text}"
            print(error_message)
            raise Exception(error_message)

    def set_calendar_id(self, calendar_id):
        """Definir o ID do calendário a ser usado"""
        self.calendar_id = calendar_id

    def list_events(self):
        if not self.calendar_id:
            raise Exception("ID do calendário não definido. Use set_calendar_id() primeiro.")
            
        url = f"https://graph.microsoft.com/v1.0/me/calendars/{self.calendar_id}/events"
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            events = response.json().get('value', [])
            return events
        else:
            # Adicionar mais detalhes sobre o erro
            error_message = f"Erro ao listar eventos na API do Outlook. Status: {response.status_code}, Resposta: {response.text}"
            print(error_message)  # Imprimir o erro para debug
            raise Exception(error_message)

    def create_event(self, event):
        if not self.calendar_id:
            raise Exception("ID do calendário não definido. Use set_calendar_id() primeiro.")
            
        url = f"https://graph.microsoft.com/v1.0/me/calendars/{self.calendar_id}/events"
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        response = requests.post(url, headers=headers, data=json.dumps(event))
        if response.status_code in (200, 201):
            return response.json()
        else:
            error_message = f"Erro ao criar evento na API do Outlook. Status: {response.status_code}, Resposta: {response.text}"
            print(error_message)
            raise Exception(error_message)
