# importando os arquivos que contém a classe GoogleCalendarSync
from google_calendar_sync import GoogleCalendarSync 
from outlook_calendar_sync import OutlookCalendarSync 

def main():
    # CONFIGURANDO O GOOGLE CALENDAR
    google_credentials_file = 'credentials.json' # caminho para o arquivo de credenciais do google console
    google_token_file = 'token.json' # caminho para o arquivo de token do google console
    google_calendar_id = "primary" # Identificador do calendário do google, pelo que eu li, geralmente usamos primary

    # CONFIGURANDO O OUTLOOK CALENDAR
    outlook_client_id = "55368e06-d089-466d-ac29-7e3ed48e8c3b"
    outlook_client_secret = "mbK8Q~llHzwaQ4YuVeUEv3XI4kMCiTO-QalUka~q"
    outlook_tenant_id = "3ef67eec-8b26-4d7d-9c40-145014dee515"
    outlook_calendar_id = "AAMkADM4YWE1NWRhLWMyNDEtNDUyYy1hNzFiLTZjMTYwMTMxYTI3ZQBGAAAAAABPzfZHPSIEQY-bwshVFPIjBwD9zZOMHrMzS6ymJIc30mPTAAAAAAEGAAD9zZOMHrMzS6ymJIc30mPTAAACH1QGAAA="
    outlook_redirect_uri = "http://localhost:50141"

    # Inicializa os clientes de cada API
    print("\n=== Autenticando no Google Calendar ===")
    google_sync = GoogleCalendarSync(google_credentials_file, google_token_file, google_calendar_id)
    
    # Inicializar o sincronizador do Outlook
    print("\n=== Autenticando no Outlook Calendar ===")
    outlook_sync = OutlookCalendarSync(
        outlook_client_id, 
        outlook_client_secret, 
        outlook_tenant_id, 
        outlook_redirect_uri
    )
    
    # Definir o ID do calendário do Outlook
    outlook_sync.set_calendar_id(outlook_calendar_id)
    
    # Listar eventos do Google
    try:
        print("\n=== Eventos do Google Calendar ===")
        google_events = google_sync.list_events()
        if google_events:
            for event in google_events:
                print(f" - {event.get('summary', 'Sem título')} ({event.get('start', {}).get('dateTime', 'Sem data')})")
        else:
            print("Nenhum evento encontrado no calendário do Google.")
    except Exception as e:
        print(f"Erro ao listar eventos do Google: {e}")
    
    # Listar eventos do Outlook
    try:
        print("\n=== Eventos do Outlook Calendar ===")
        outlook_events = outlook_sync.list_events()
        if outlook_events:
            for event in outlook_events:
                print(f" - {event.get('subject', 'Sem título')} ({event.get('start', {}).get('dateTime', 'Sem data')})")
        else:
            print("Nenhum evento encontrado no calendário do Outlook.")
    except Exception as e:
        print(f"Erro ao listar eventos do Outlook: {e}")

if __name__ == "__main__":
    main()





