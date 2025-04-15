# importando os arquivos que contém as classes necessárias
from google_calendar_sync import GoogleCalendarSync
from outlook_calendar_sync import OutlookCalendarSync
from calendar_synchronizer import CalendarSynchronizer

# Importar o Expresso - descomentar esta linha
from vcard_sync2 import sincronizarExpresso
import time
from datetime import datetime, timedelta


def main():
    # CONFIGURANDO O GOOGLE CALENDAR
    google_credentials_file = (
        "credentials.json"  # caminho para o arquivo de credenciais do google console
    )
    google_token_file = (
        "token.json"  # caminho para o arquivo de token do google console
    )
    google_calendar_id = "primary"  # Identificador do calendário do google, pelo que eu li, geralmente usamos primary

    # CONFIGURANDO O OUTLOOK CALENDAR
    outlook_client_id = "55368e06-d089-466d-ac29-7e3ed48e8c3b"
    outlook_client_secret = "mbK8Q~llHzwaQ4YuVeUEv3XI4kMCiTO-QalUka~q"
    outlook_tenant_id = "3ef67eec-8b26-4d7d-9c40-145014dee515"
    outlook_calendar_id = "AAMkADM4YWE1NWRhLWMyNDEtNDUyYy1hNzFiLTZjMTYwMTMxYTI3ZQBGAAAAAABPzfZHPSIEQY-bwshVFPIjBwD9zZOMHrMzS6ymJIc30mPTAAAAAAEGAAD9zZOMHrMzS6ymJIc30mPTAAACH1QGAAA="
    outlook_redirect_uri = "http://localhost:50141"

    # Inicializa os clientes de cada API
    print("\n=== Autenticando no Google Calendar ===")
    google_sync = GoogleCalendarSync(
        google_credentials_file, google_token_file, google_calendar_id
    )

    # Inicializar o sincronizador do Outlook
    print("\n=== Autenticando no Outlook Calendar ===")
    outlook_sync = OutlookCalendarSync(
        outlook_client_id,
        outlook_client_secret,
        outlook_tenant_id,
        outlook_redirect_uri,
    )

    # Definir o ID do calendário do Outlook
    outlook_sync.set_calendar_id(outlook_calendar_id)

    # CONFIGURAR EXPRESSO - descomentar e configurar com suas credenciais
    expresso_sync = None
    try:
        print("\n=== Autenticando no Expresso ===")
        expresso_sync = sincronizarExpresso("pablo.henrique1", "@Taisatt84671514")
        expresso_sync.login()
        expresso_sync.selecionarCalendario()
    except Exception as e:
        print(f"Erro ao configurar Expresso: {e}")
        expresso_sync = None

    # Criar o sincronizador com os três calendários
    synchronizer = CalendarSynchronizer(google_sync, outlook_sync, expresso_sync)

    print("\n=== Iniciando sincronização em tempo real ===")
    print("Este modo irá:")
    print("- Monitorar as mudanças nos três calendários")
    print("- Sincronizar eventos novos, atualizados ou excluídos")
    print("- Não duplicar eventos existentes")
    print("- Verificar mudanças a cada 20 segundos")

    # Testar sincronização criando eventos
    teste_sincronizacao = False  # Defina como True para ativar o teste

    if teste_sincronizacao:
        print("\n=== MODO DE TESTE DE SINCRONIZAÇÃO ATIVADO ===")
        print("Criando eventos de teste para verificar a sincronização...")

        # Criar evento no Google
        print("\n=== Criando evento de teste no Google Calendar ===")
        evento_google = {
            "summary": f'Evento de teste do Google - {datetime.now().strftime("%H:%M:%S")}',
            "description": "Este é um evento de teste para verificar a sincronização.",
            "start": {
                "dateTime": (
                    datetime.now().replace(hour=10, minute=0, second=0)
                    + timedelta(days=1)
                ).isoformat(),
                "timeZone": "America/Recife",
            },
            "end": {
                "dateTime": (
                    datetime.now().replace(hour=11, minute=0, second=0)
                    + timedelta(days=1)
                ).isoformat(),
                "timeZone": "America/Recife",
            },
        }

        try:
            result_google = google_sync.create_event(evento_google)
            print(f"Evento criado no Google com ID: {result_google.get('id')}")

            # Esperar para sincronizar
            print("Aguardando 30 segundos para iniciar sincronização...")
            time.sleep(30)

            # Realizar uma sincronização manual
            print("Executando sincronização manual...")
            stats = synchronizer.sync_changes_only()
            print(f"Resultado da sincronização: {stats}")

            # Criar evento no Outlook
            print("\n=== Criando evento de teste no Outlook Calendar ===")
            evento_outlook = {
                "subject": f'Evento de teste do Outlook - {datetime.now().strftime("%H:%M:%S")}',
                "body": {
                    "contentType": "HTML",
                    "content": "Este é um evento de teste para verificar a sincronização do Outlook.",
                },
                "start": {
                    "dateTime": (
                        datetime.now().replace(hour=14, minute=0, second=0)
                        + timedelta(days=1)
                    ).isoformat(),
                    "timeZone": "America/Recife",
                },
                "end": {
                    "dateTime": (
                        datetime.now().replace(hour=15, minute=0, second=0)
                        + timedelta(days=1)
                    ).isoformat(),
                    "timeZone": "America/Recife",
                },
            }

            result_outlook = outlook_sync.create_event(evento_outlook)
            print(f"Evento criado no Outlook com ID: {result_outlook.get('id')}")

            # Esperar para sincronizar
            print("Aguardando 30 segundos para segunda sincronização...")
            time.sleep(30)

            # Realizar segunda sincronização manual
            print("Executando segunda sincronização manual...")
            stats = synchronizer.sync_changes_only()
            print(f"Resultado da segunda sincronização: {stats}")

        except Exception as e:
            print(f"Erro durante o teste de sincronização: {e}")

    # Iniciar sincronização
    synchronizer.start_realtime_sync(
        interval=60, cleanup_interval=86400, days_to_keep=0
    )


if __name__ == "__main__":
    main()
