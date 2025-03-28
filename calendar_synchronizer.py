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
        # Adicionar mapeamento entre IDs de eventos para rastrear correspondências
        self.google_to_outlook_map = {}  # Google ID -> Outlook ID
        self.outlook_to_google_map = {}  # Outlook ID -> Google ID
    
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
        google_updated = {id: event for id, event in new_google_cache.items() 
                         if id in self.google_events_cache and self._is_event_updated(event, self.google_events_cache[id])}
        google_deleted = {id: self.google_events_cache[id] for id in self.google_events_cache 
                         if id not in new_google_cache}
        
        outlook_added = {id: event for id, event in new_outlook_cache.items() 
                        if id not in self.outlook_events_cache}
        outlook_updated = {id: event for id, event in new_outlook_cache.items() 
                          if id in self.outlook_events_cache and self._is_event_updated(event, self.outlook_events_cache[id])}
        outlook_deleted = {id: self.outlook_events_cache[id] for id in self.outlook_events_cache 
                          if id not in new_outlook_cache}
        
        # Debug info
        if google_added:
            print(f"Novos eventos detectados no Google: {len(google_added)}")
            for id, event in google_added.items():
                print(f"  - {event.get('summary', 'Sem título')} ({id})")
        
        if outlook_added:
            print(f"Novos eventos detectados no Outlook: {len(outlook_added)}")
            for id, event in outlook_added.items():
                print(f"  - {event.get('subject', 'Sem título')} ({id})")
        
        # Atualizar caches
        self.google_events_cache = new_google_cache
        self.outlook_events_cache = new_outlook_cache
        self.last_sync_time = datetime.now()
        
        return {
            'google': {
                'added': google_added,
                'updated': google_updated,
                'deleted': google_deleted
            },
            'outlook': {
                'added': outlook_added,
                'updated': outlook_updated,
                'deleted': outlook_deleted
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
        for event_id, google_event in changes['google']['added'].items():
            print(f"Processando novo evento do Google: {google_event.get('summary', 'Sem título')}")
            
            # Verificar se este evento já tem um mapeamento (para evitar duplicatas)
            if event_id in self.google_to_outlook_map:
                print(f"  - Já sincronizado com o Outlook com ID: {self.google_to_outlook_map[event_id]}")
                continue
                
            # Verificar se este é um evento recentemente sincronizado do Outlook
            is_synced_from_outlook = False
            for outlook_id, outlook_event in self.outlook_events_cache.items():
                if self._events_match(google_event, outlook_event):
                    print(f"  - Corresponde a um evento existente no Outlook: {outlook_event.get('subject', 'Sem título')}")
                    self._store_event_mapping(event_id, outlook_id)
                    is_synced_from_outlook = True
                    break
            
            if not is_synced_from_outlook:
                try:
                    outlook_event = self._format_google_to_outlook(google_event)
                    if outlook_event:
                        print(f"  - Criando no Outlook: {outlook_event.get('subject', 'Sem título')}")
                        result = self.outlook_sync.create_event(outlook_event)
                        outlook_id = result.get('id')
                        if outlook_id:
                            self._store_event_mapping(event_id, outlook_id)
                            print(f"  - Criado no Outlook com ID: {outlook_id}")
                            stats['google_to_outlook']['created'] += 1
                except Exception as e:
                    print(f"  - Erro ao criar evento no Outlook: {e}")
        
        # Processar eventos atualizados no Google
        for event_id, google_event in changes['google']['updated'].items():
            if event_id in self.google_to_outlook_map:
                outlook_id = self.google_to_outlook_map[event_id]
                try:
                    outlook_event = self._format_google_to_outlook(google_event)
                    if outlook_event:
                        print(f"Atualizando no Outlook: {outlook_event.get('subject', 'Sem título')}")
                        self.outlook_sync.update_event(outlook_id, outlook_event)
                        stats['google_to_outlook']['updated'] += 1
                except Exception as e:
                    print(f"Erro ao atualizar evento no Outlook: {e}")
        
        # Processar eventos excluídos no Google
        for event_id, google_event in changes['google']['deleted'].items():
            if event_id in self.google_to_outlook_map:
                outlook_id = self.google_to_outlook_map[event_id]
                try:
                    print(f"Excluindo do Outlook: {google_event.get('summary', 'Sem título')}")
                    self.outlook_sync.delete_event(outlook_id)
                    # Remover dos mapeamentos
                    del self.outlook_to_google_map[outlook_id]
                    del self.google_to_outlook_map[event_id]
                    stats['google_to_outlook']['deleted'] += 1
                except Exception as e:
                    print(f"Erro ao excluir evento do Outlook: {e}")
        
        # Processar eventos adicionados no Outlook
        for event_id, outlook_event in changes['outlook']['added'].items():
            print(f"Processando novo evento do Outlook: {outlook_event.get('subject', 'Sem título')}")
            
            # Verificar se este evento já tem um mapeamento
            if event_id in self.outlook_to_google_map:
                print(f"  - Já sincronizado com o Google com ID: {self.outlook_to_google_map[event_id]}")
                continue
                
            # Verificar se este é um evento recentemente sincronizado do Google
            is_synced_from_google = False
            for google_id, google_event in self.google_events_cache.items():
                if self._events_match(google_event, outlook_event):
                    print(f"  - Corresponde a um evento existente no Google: {google_event.get('summary', 'Sem título')}")
                    self._store_event_mapping(google_id, event_id)
                    is_synced_from_google = True
                    break
            
            if not is_synced_from_google:
                try:
                    google_event = self._format_outlook_to_google(outlook_event)
                    if google_event:
                        print(f"  - Criando no Google: {google_event.get('summary', 'Sem título')}")
                        result = self.google_sync.create_event(google_event)
                        google_id = result.get('id')
                        if google_id:
                            self._store_event_mapping(google_id, event_id)
                            print(f"  - Criado no Google com ID: {google_id}")
                            stats['outlook_to_google']['created'] += 1
                except Exception as e:
                    print(f"  - Erro ao criar evento no Google: {e}")
        
        # Processar eventos atualizados no Outlook
        for event_id, outlook_event in changes['outlook']['updated'].items():
            if event_id in self.outlook_to_google_map:
                google_id = self.outlook_to_google_map[event_id]
                try:
                    google_event = self._format_outlook_to_google(outlook_event)
                    if google_event:
                        print(f"Atualizando no Google: {google_event.get('summary', 'Sem título')}")
                        self.google_sync.update_event(google_id, google_event)
                        stats['outlook_to_google']['updated'] += 1
                except Exception as e:
                    print(f"Erro ao atualizar evento no Google: {e}")
        
        # Processar eventos excluídos no Outlook
        for event_id, outlook_event in changes['outlook']['deleted'].items():
            if event_id in self.outlook_to_google_map:
                google_id = self.outlook_to_google_map[event_id]
                try:
                    print(f"Excluindo do Google: {outlook_event.get('subject', 'Sem título')}")
                    self.google_sync.delete_event(google_id)
                    # Remover dos mapeamentos
                    del self.google_to_outlook_map[google_id]
                    del self.outlook_to_google_map[event_id]
                    stats['outlook_to_google']['deleted'] += 1
                except Exception as e:
                    print(f"Erro ao excluir evento do Google: {e}")
        
        return stats
    
    def _is_event_updated(self, current_event, cached_event):
        """Verifica se um evento foi atualizado comparando campos relevantes"""
        # Para Google
        if 'updated' in current_event and 'updated' in cached_event:
            return current_event['updated'] != cached_event['updated']
        
        # Para Outlook
        if 'lastModifiedDateTime' in current_event and 'lastModifiedDateTime' in cached_event:
            return current_event['lastModifiedDateTime'] != cached_event['lastModifiedDateTime']
        
        # Comparação manual de campos importantes
        important_fields = ['summary', 'subject', 'start', 'end', 'location', 'description']
        
        for field in important_fields:
            if field in current_event and field in cached_event:
                # Para campos aninhados como start e end
                if isinstance(current_event[field], dict) and isinstance(cached_event[field], dict):
                    # Comparar dateTime para campos de data/hora
                    if 'dateTime' in current_event[field] and 'dateTime' in cached_event[field]:
                        if current_event[field]['dateTime'] != cached_event[field]['dateTime']:
                            return True
                # Para campos simples
                elif current_event[field] != cached_event[field]:
                    return True
        
        return False
    
    def _events_match(self, google_event, outlook_event):
        """Verifica se um evento do Google corresponde a um evento do Outlook"""
        # Comparar título
        google_title = google_event.get('summary', '').strip()
        outlook_title = outlook_event.get('subject', '').strip()
        
        if not google_title or not outlook_title:
            return False
            
        title_match = google_title.lower() == outlook_title.lower()
        if not title_match:
            return False
            
        # Comparar data/hora de início
        google_start = google_event.get('start', {}).get('dateTime', '')
        outlook_start = outlook_event.get('start', {}).get('dateTime', '')
        
        if not google_start or not outlook_start:
            return False
            
        try:
            google_dt = datetime.fromisoformat(google_start.replace('Z', '+00:00'))
            outlook_dt = datetime.fromisoformat(outlook_start.replace('Z', '+00:00'))
            time_diff = abs((google_dt - outlook_dt).total_seconds())
            if time_diff > 300:  # 5 minutos de tolerância
                return False
        except (ValueError, TypeError):
            return False
            
        return True
    
    def start_realtime_sync(self, interval=20):
        """Inicia sincronização em tempo real a cada 'interval' segundos"""
        print(f"\n=== Iniciando sincronização em tempo real a cada {interval} segundos ===")
        print("Monitorando apenas mudanças (criações, atualizações, exclusões)")
        print("Pressione Ctrl+C para interromper")
        
        # Inicializar caches e mapeamentos
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
        if not google_event.get('summary') or not google_event.get('start'):
            return None
            
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
        
        if google_event.get('location'):
            outlook_event['location'] = {
                'displayName': google_event.get('location')
            }
            
        return outlook_event
    
    def _format_outlook_to_google(self, outlook_event):
        """Converte um evento do Outlook para o formato do Google"""
        if not outlook_event.get('subject') or not outlook_event.get('start'):
            return None
            
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
        
        if outlook_event.get('location'):
            google_event['location'] = outlook_event.get('location', {}).get('displayName', '')
            
        return google_event

    def _is_event_updated(self, current_event, cached_event):
        """Checks if an event has been updated by comparing relevant fields"""
        # For Google
        if 'updated' in current_event and 'updated' in cached_event:
            return current_event['updated'] != cached_event['updated']
        
        # For Outlook
        if 'lastModifiedDateTime' in current_event and 'lastModifiedDateTime' in cached_event:
            return current_event['lastModifiedDateTime'] != cached_event['lastModifiedDateTime']
        
        # Manual comparison of important fields
        important_fields = ['summary', 'subject', 'start', 'end', 'location', 'description']
        
        for field in important_fields:
            if field in current_event and field in cached_event:
                # Para campos aninhados como start e end
                if isinstance(current_event[field], dict) and isinstance(cached_event[field], dict):
                    # Comparar dateTime para campos de data/hora
                    if 'dateTime' in current_event[field] and 'dateTime' in cached_event[field]:
                        if current_event[field]['dateTime'] != cached_event[field]['dateTime']:
                            return True
                # Para campos simples
                elif current_event[field] != cached_event[field]:
                    return True
        
        return False
        
    def _store_event_mapping(self, google_id, outlook_id):
        """Armazena o mapeamento entre IDs de eventos do Google e Outlook"""
        print(f"Mapeando evento: Google ID {google_id} <-> Outlook ID {outlook_id}")
        self.google_to_outlook_map[google_id] = outlook_id
        self.outlook_to_google_map[outlook_id] = google_id