# calendar_synchronizer.py
from datetime import datetime, timedelta
import time

# Adicione/modifique estas partes no arquivo calendar_synchronizer.py

class CalendarSynchronizer:
    def __init__(self, google_sync, outlook_sync):
        self.google_sync = google_sync
        self.outlook_sync = outlook_sync
        # Armazenar o estado atual dos calendários para detectar mudanças
        self.google_events_cache = {}
        self.outlook_events_cache = {}
        self.last_sync_time = datetime.now()
    
    def _update_caches(self):
        """Atualiza os caches com o estado atual dos calendários"""
        # Obter eventos atuais - usar data atual para pegar eventos recentes
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        google_events = self.google_sync.list_events(today)
        outlook_events = self.outlook_sync.list_events(today)
        
        print(f"Eventos encontrados - Google: {len(google_events)}, Outlook: {len(outlook_events)}")
        
        # Atualizar cache do Google
        new_google_cache = {}
        for event in google_events:
            if 'id' in event:
                new_google_cache[event['id']] = event
        
        # Atualizar cache do Outlook
        new_outlook_cache = {}
        for event in outlook_events:
            if 'id' in event:
                new_outlook_cache[event['id']] = event
        
        # Detectar mudanças desde a última sincronização
        google_added = {id: event for id, event in new_google_cache.items() 
                       if id not in self.google_events_cache}
        
        outlook_added = {id: event for id, event in new_outlook_cache.items() 
                        if id not in self.outlook_events_cache}
        
        # Imprimir informações de debug
        if google_added:
            print(f"Novos eventos detectados no Google: {len(google_added)}")
            for id, event in google_added.items():
                print(f"  - Novo no Google: {event.get('summary', 'Sem título')} ({id})")
                
                # Verificar se já existe correspondente no Outlook
                found = False
                for outlook_id, outlook_event in new_outlook_cache.items():
                    if self._events_match(event, outlook_event):
                        found = True
                        print(f"    - Já existe no Outlook: {outlook_event.get('subject', 'Sem título')} ({outlook_id})")
                        break
                
                if not found:
                    print(f"    - Precisa ser criado no Outlook")
        
        if outlook_added:
            print(f"Novos eventos detectados no Outlook: {len(outlook_added)}")
            for id, event in outlook_added.items():
                print(f"  - Novo no Outlook: {event.get('subject', 'Sem título')} ({id})")
        
        # Atualizar caches
        old_google_cache = self.google_events_cache
        old_outlook_cache = self.outlook_events_cache
        
        self.google_events_cache = new_google_cache
        self.outlook_events_cache = new_outlook_cache
        self.last_sync_time = datetime.now()
        
        return {
            'google': {
                'added': google_added,
                'updated': {},  # Simplificando para focar em eventos novos
                'deleted': {}
            },
            'outlook': {
                'added': outlook_added,
                'updated': {},
                'deleted': {}
            }
        }
    
    def sync_changes_only(self):
        """Sincroniza apenas as mudanças detectadas desde a última sincronização"""
        print(f"\n=== Verificando mudanças desde {self.last_sync_time.strftime('%H:%M:%S')} ===")
        
        # Detectar mudanças
        changes = self._update_caches()
        
        # Contadores
        stats = {
            'google_to_outlook': {'created': 0, 'updated': 0, 'deleted': 0},
            'outlook_to_google': {'created': 0, 'updated': 0, 'deleted': 0}
        }
        
        # Processar eventos adicionados no Google
        if changes['google']['added']:
            print(f"\nProcessando {len(changes['google']['added'])} novos eventos do Google...")
            
        for event_id, google_event in changes['google']['added'].items():
            print(f"Processando evento do Google: {google_event.get('summary', 'Sem título')}")
            
            # Verificar se já existe correspondente no Outlook
            found = False
            for outlook_id, outlook_event in self.outlook_events_cache.items():
                if self._events_match(google_event, outlook_event):
                    found = True
                    print(f"  - Evento já existe no Outlook: {outlook_event.get('subject', 'Sem título')}")
                    break
            
            if not found:
                try:
                    outlook_event = self._format_google_to_outlook(google_event)
                    if outlook_event:
                        print(f"  - Convertendo para formato Outlook: {outlook_event.get('subject', 'Sem título')}")
                        result = self.outlook_sync.create_event(outlook_event)
                        print(f"  - Criado no Outlook: {outlook_event.get('subject', 'Sem título')}")
                        stats['google_to_outlook']['created'] += 1
                except Exception as e:
                    print(f"  - Erro ao criar evento no Outlook: {e}")
        
        # Processar eventos adicionados no Outlook
        if changes['outlook']['added']:
            print(f"\nProcessando {len(changes['outlook']['added'])} novos eventos do Outlook...")
            
        for event_id, outlook_event in changes['outlook']['added'].items():
            print(f"Processando evento do Outlook: {outlook_event.get('subject', 'Sem título')}")
            
            # Verificar se já existe correspondente no Google
            found = False
            for google_id, google_event in self.google_events_cache.items():
                if self._events_match(google_event, outlook_event):
                    found = True
                    print(f"  - Evento já existe no Google: {google_event.get('summary', 'Sem título')}")
                    break
            
            if not found:
                try:
                    google_event = self._format_outlook_to_google(outlook_event)
                    if google_event:
                        print(f"  - Convertendo para formato Google: {google_event.get('summary', 'Sem título')}")
                        result = self.google_sync.create_event(google_event)
                        print(f"  - Criado no Google: {google_event.get('summary', 'Sem título')}")
                        stats['outlook_to_google']['created'] += 1
                except Exception as e:
                    print(f"  - Erro ao criar evento no Google: {e}")
        
        return stats
    
    def _events_match(self, google_event, outlook_event):
        """Verifica se um evento do Google corresponde a um evento do Outlook"""
        # Comparar título
        google_title = google_event.get('summary', '').strip()
        outlook_title = outlook_event.get('subject', '').strip()
        
        if not google_title or not outlook_title:
            return False
            
        title_match = google_title.lower() == outlook_title.lower()
        
        # Comparar data/hora de início
        google_start = google_event.get('start', {}).get('dateTime', '')
        outlook_start = outlook_event.get('start', {}).get('dateTime', '')
        
        if not google_start or not outlook_start:
            return False
            
        time_match = False
        try:
            google_dt = datetime.fromisoformat(google_start.replace('Z', '+00:00'))
            outlook_dt = datetime.fromisoformat(outlook_start.replace('Z', '+00:00'))
            time_diff = abs((google_dt - outlook_dt).total_seconds())
            time_match = time_diff < 300  # 5 minutos de tolerância
        except (ValueError, TypeError) as e:
            print(f"Erro ao comparar datas: {e}")
            return False
            
        if title_match and time_match:
            print(f"Evento correspondente encontrado: '{google_title}' = '{outlook_title}'")
            return True
            
        return False
    
    def start_realtime_sync(self, interval=20):
        """Inicia sincronização em tempo real a cada 'interval' segundos"""
        print(f"\n=== Iniciando sincronização em tempo real a cada {interval} segundos ===")
        print("Monitorando apenas mudanças (criações, atualizações, exclusões)")
        print("Pressione Ctrl+C para interromper")
        
        # Inicializar caches
        self._update_caches()
        print("Estado inicial dos calendários armazenado. Monitorando mudanças...")
        
        try:
            while True:
                start_time = time.time()
                
                # Sincronizar apenas mudanças
                stats = self.sync_changes_only()
                
                # Resumo das operações
                total_ops = sum(sum(category.values()) for category in stats.values())
                
                if total_ops > 0:
                    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Mudanças sincronizadas:")
                    print(f"- Google → Outlook: {stats['google_to_outlook']['created']} criados, "
                          f"{stats['google_to_outlook']['updated']} atualizados, "
                          f"{stats['google_to_outlook']['deleted']} excluídos")
                    print(f"- Outlook → Google: {stats['outlook_to_google']['created']} criados, "
                          f"{stats['outlook_to_google']['updated']} atualizados, "
                          f"{stats['outlook_to_google']['deleted']} excluídos")
                else:
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] Nenhuma mudança detectada")
                
                # Calcular tempo de espera
                elapsed = time.time() - start_time
                wait_time = max(0, interval - elapsed)
                
                if wait_time > 0:
                    print(f"Aguardando {wait_time:.1f} segundos...")
                    time.sleep(wait_time)
                
        except KeyboardInterrupt:
            print("\n=== Sincronização em tempo real interrompida pelo usuário ===")

    # Add these methods to your CalendarSynchronizer class
    
    def _format_google_to_outlook(self, google_event):
        """Converte um evento do Google para o formato do Outlook"""
        # Verificar se o evento tem os campos necessários
        if not google_event.get('summary') or not google_event.get('start'):
            return None
            
        # Criar evento no formato do Outlook
        outlook_event = {
            'subject': google_event.get('summary'),
            'body': {
                'contentType': 'HTML',
                'content': google_event.get('description', 'Evento sincronizado do Google Calendar')
            },
            'start': {
                'dateTime': google_event.get('start', {}).get('dateTime', ''),
                'timeZone': google_event.get('start', {}).get('timeZone', 'UTC')
            },
            'end': {
                'dateTime': google_event.get('end', {}).get('dateTime', ''),
                'timeZone': google_event.get('end', {}).get('timeZone', 'UTC')
            }
        }
        
        # Adicionar localização se disponível
        if google_event.get('location'):
            outlook_event['location'] = {
                'displayName': google_event.get('location')
            }
            
        return outlook_event
    
    def _format_outlook_to_google(self, outlook_event):
        """Converte um evento do Outlook para o formato do Google"""
        # Verificar se o evento tem os campos necessários
        if not outlook_event.get('subject') or not outlook_event.get('start'):
            return None
            
        # Criar evento no formato do Google
        google_event = {
            'summary': outlook_event.get('subject'),
            'description': outlook_event.get('body', {}).get('content', 'Evento sincronizado do Outlook Calendar'),
            'start': {
                'dateTime': outlook_event.get('start', {}).get('dateTime', ''),
                'timeZone': outlook_event.get('start', {}).get('timeZone', 'UTC')
            },
            'end': {
                'dateTime': outlook_event.get('end', {}).get('dateTime', ''),
                'timeZone': outlook_event.get('end', {}).get('timeZone', 'UTC')
            }
        }
        
        # Adicionar localização se disponível
        if outlook_event.get('location'):
            google_event['location'] = outlook_event.get('location', {}).get('displayName', '')
            
        return google_event