# importando os arquivos que contém as classes necessárias
from google_calendar_sync import GoogleCalendarSync 
from outlook_calendar_sync import OutlookCalendarSync
from calendar_synchronizer import CalendarSynchronizer
# Importar o Expresso se estiver usando
# from vcard_sync2 import sincronizarExpresso

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
    
    # OPCIONAL: Configurar Expresso
    # expresso_sync = None
    # try:
    #     print("\n=== Autenticando no Expresso ===")
    #     expresso_sync = sincronizarExpresso("seu_usuario", "sua_senha")
    #     expresso_sync.login()
    #     expresso_sync.selecionarCalendario()
    # except Exception as e:
    #     print(f"Erro ao configurar Expresso: {e}")
    #     expresso_sync = None
    
    # Criar o sincronizador
    # Se não estiver usando o Expresso:
    synchronizer = CalendarSynchronizer(google_sync, outlook_sync)
    # Se estiver usando o Expresso:
    # synchronizer = CalendarSynchronizer(google_sync, outlook_sync, expresso_sync)
    
    print("\n=== Iniciando sincronização em tempo real ===")
    print("Este modo irá:")
    print("- Monitorar apenas as mudanças em ambos os calendários")
    print("- Sincronizar apenas eventos novos, atualizados ou excluídos")
    print("- Não duplicar eventos existentes")
    print("- Verificar mudanças a cada 20 segundos")
    
    # Iniciar sincronização em tempo real
    # Parâmetros:
    # - interval: tempo entre sincronizações (20 segundos)
    # - cleanup_interval: tempo entre limpezas (86400 segundos = 1 dia)
    # - days_to_keep: dias no passado para manter (0 = apenas hoje e futuro)
    synchronizer.start_realtime_sync(interval=20, cleanup_interval=86400, days_to_keep=0)

if __name__ == "__main__":
    main()





