# calendar_synchronizer.py
from database import DatabaseManager
from datetime import datetime, timedelta, timezone
import time
import re
import json

# Adicione/modifique estas partes no arquivo calendar_synchronizer.py


class CalendarSynchronizer:
    def __init__(self, google_sync, outlook_sync, expresso_sync=None):
        self.google_sync = google_sync
        self.outlook_sync = outlook_sync
        self.expresso_sync = expresso_sync  # Pode ser None se não estiver usando

        # Inicializar o gerenciador de banco de dados
        self.db = DatabaseManager()

        # Manter estas propriedades para compatibilidade com código existente
        self.google_events_cache = {}
        self.outlook_events_cache = {}
        self.last_sync_time = datetime.now()
        self.google_to_outlook_map = {}
        self.outlook_to_google_map = {}

    def _update_caches(self):
        """Atualiza os caches com o estado atual dos calendários"""
        # Obter eventos atuais - usar data atual para pegar eventos recentes
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        google_events = self.google_sync.list_events(today)
        outlook_events = self.outlook_sync.list_events(today)

        # Adicionado suporte para Expresso (opcional)
        expresso_events = []
        if hasattr(self, "expresso_sync") and self.expresso_sync:
            expresso_events = self.expresso_sync.obterEventos()

        print(
            f"Eventos encontrados - Google: {len(google_events)}, Outlook: {len(outlook_events)}"
        )
        if hasattr(self, "expresso_sync") and self.expresso_sync:
            print(f"Eventos encontrados - Expresso: {len(expresso_events)}")

        # Armazenar todos os eventos no banco de dados
        for event in google_events:
            if "id" in event:
                self.db.store_google_event(event)

        for event in outlook_events:
            if "id" in event:
                self.db.store_outlook_event(event)

        if hasattr(self, "expresso_sync") and self.expresso_sync:
            for event in expresso_events:
                if "id" in event:
                    self.db.store_expresso_event(event)

        # Para compatibilidade com o código existente, ainda atualizar os caches
        new_google_cache = {}
        for event in google_events:
            if "id" in event:
                new_google_cache[event["id"]] = event

        new_outlook_cache = {}
        for event in outlook_events:
            if "id" in event:
                new_outlook_cache[event["id"]] = event

        # Continuação da lógica existente para detectar mudanças...
        google_added = {
            id: event
            for id, event in new_google_cache.items()
            if id not in self.google_events_cache
        }
        google_updated = {
            id: event
            for id, event in new_google_cache.items()
            if id in self.google_events_cache
            and self._is_event_updated(event, self.google_events_cache[id])
        }
        google_deleted = {
            id: self.google_events_cache[id]
            for id in self.google_events_cache
            if id not in new_google_cache
        }

        outlook_added = {
            id: event
            for id, event in new_outlook_cache.items()
            if id not in self.outlook_events_cache
        }
        outlook_updated = {
            id: event
            for id, event in new_outlook_cache.items()
            if id in self.outlook_events_cache
            and self._is_event_updated(event, self.outlook_events_cache[id])
        }
        outlook_deleted = {
            id: self.outlook_events_cache[id]
            for id in self.outlook_events_cache
            if id not in new_outlook_cache
        }

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

        # Reconstruir os mapas baseado no banco de dados
        self._rebuild_maps_from_db()

        return {
            "google": {
                "added": google_added,
                "updated": google_updated,
                "deleted": google_deleted,
            },
            "outlook": {
                "added": outlook_added,
                "updated": outlook_updated,
                "deleted": outlook_deleted,
            },
        }

    def _rebuild_maps_from_db(self):
        """Reconstrói os mapas de ID a partir do banco de dados"""
        mappings = self.db.get_all_mappings()

        # Limpar mapas existentes
        self.google_to_outlook_map = {}
        self.outlook_to_google_map = {}

        # Adicionar mapas para Expresso
        self.google_to_expresso_map = {}
        self.outlook_to_expresso_map = {}
        self.expresso_to_google_map = {}
        self.expresso_to_outlook_map = {}

        # Reconstruir a partir do banco de dados
        for mapping in mappings:
            _, outlook_id, google_id, expresso_id, _, _ = mapping

            if outlook_id and google_id:
                self.google_to_outlook_map[google_id] = outlook_id
                self.outlook_to_google_map[outlook_id] = google_id

            # Adicionar mapeamentos para Expresso
            if google_id and expresso_id:
                self.google_to_expresso_map[google_id] = expresso_id
                self.expresso_to_google_map[expresso_id] = google_id

            if outlook_id and expresso_id:
                self.outlook_to_expresso_map[outlook_id] = expresso_id
                self.expresso_to_outlook_map[expresso_id] = outlook_id

    def sync_changes_only(self):
        """Sincroniza apenas as mudanças detectadas desde a última sincronização"""
        print(
            f"\n=== Verificando mudanças desde {self.last_sync_time.strftime('%H:%M:%S')} ==="
        )

        # Recarregar a página do Expresso antes de começar a sincronização, se existir
        if hasattr(self, "expresso_sync") and self.expresso_sync and self.expresso_sync.driver:
            try:
                print("Atualizando página do Expresso...")
                self.expresso_sync.selecionarCalendario()  # Isso vai recarregar a página do calendário
            except Exception as e:
                print(f"Erro ao atualizar página do Expresso: {e}")

        # Manter um registro de eventos que estão sendo sincronizados
        events_being_synced = set()

        # Detectar mudanças usando o método existente (que agora também atualiza o banco de dados)
        changes = self._update_caches()

        # Contadores
        stats = {
            "google_to_outlook": {"created": 0, "updated": 0, "deleted": 0},
            "outlook_to_google": {"created": 0, "updated": 0, "deleted": 0},
        }

        # Adicionar contador para Expresso se necessário
        if hasattr(self, "expresso_sync") and self.expresso_sync:
            stats["google_to_expresso"] = {"created": 0, "updated": 0, "deleted": 0}
            stats["outlook_to_expresso"] = {"created": 0, "updated": 0, "deleted": 0}
            stats["expresso_to_google"] = {"created": 0, "updated": 0, "deleted": 0}
            stats["expresso_to_outlook"] = {"created": 0, "updated": 0, "deleted": 0}

        # Processar eventos adicionados no Google
        for event_id, google_event in changes["google"]["added"].items():
            print(f"Processando novo evento do Google: {google_event.get('summary', 'Sem título')}")
            
            # Verificar se o evento já foi processado nesta sessão
            if event_id in events_being_synced:
                print(f"  - Evento já está sendo processado nesta sessão, ignorando")
                continue
            
            # Adicionar à lista de eventos sendo processados
            events_being_synced.add(event_id)
            
            # Verificar se já existe no banco de dados antes de tudo
            mapped_ids = self.db.get_mapped_ids(event_id, "google")
            if mapped_ids and mapped_ids[0]:  # Existe mapeamento com Outlook
                outlook_id = mapped_ids[0]
                print(f"  - Já mapeado com evento do Outlook ID: {outlook_id}")
                continue
            
            # Verificar primeiro se há ID correspondente nas propriedades estendidas
            outlook_id = self._find_matching_event_by_id(event_id, "google", self.outlook_events_cache)
            if outlook_id:
                print(
                    f"  - Corresponde a um evento existente no Outlook pelo ID: {outlook_id}"
                )
                self.db.map_events(
                    google_id=event_id, outlook_id=outlook_id, origem="id_match"
                )
                self._store_event_mapping(google_id=event_id, outlook_id=outlook_id)
                outlook_match_found = True
            else:
                # Verificar correspondência baseada em propriedades
                outlook_match_found = False
                expresso_match_found = False

                # Verificar se o evento já existe no Outlook (usando o novo método)
                exists, outlook_id = self._check_event_already_exists(google_event, 'google', self.outlook_events_cache, 'outlook')
                if exists:
                    print(f"  - Evento já existe no Outlook com ID: {outlook_id}")
                    self.db.map_events(
                        google_id=event_id, outlook_id=outlook_id, origem="match"
                    )
                    self._store_event_mapping(google_id=event_id, outlook_id=outlook_id)
                    outlook_match_found = True
                    
                    # IMPORTANTE: Se o evento for encontrado, verificar se precisa ser atualizado
                    # para garantir que as versões estejam sincronizadas
                    try:
                        outlook_event_atualizado = self._format_google_to_outlook(
                            google_event
                        )
                        if outlook_event_atualizado:
                            print(
                                f"  - Atualizando no Outlook: {outlook_event_atualizado.get('subject', 'Sem título')}"
                            )
                            self.outlook_sync.update_event(
                                outlook_id, outlook_event_atualizado
                            )
                            stats["google_to_outlook"]["updated"] += 1
                    except Exception as e:
                        print(f"  - Erro ao atualizar evento no Outlook: {e}")

                # Se não tiver match no Outlook, criar evento
                if not outlook_match_found:
                    try:
                        outlook_event = self._format_google_to_outlook(google_event)
                        if outlook_event:
                            print(
                                f"  - Criando no Outlook: {outlook_event.get('subject', 'Sem título')}"
                            )
                            result = self.outlook_sync.create_event(outlook_event)
                            outlook_id = result.get("id")
                            if outlook_id:
                                # Primeiro armazenar o evento no banco
                                self.db.store_outlook_event(result)
                                # Depois mapear os eventos
                                self.db.map_events(
                                    google_id=event_id,
                                    outlook_id=outlook_id,
                                    origem="google",
                                )
                                self._store_event_mapping(
                                    google_id=event_id, outlook_id=outlook_id
                                )
                                print(f"  - Criado no Outlook com ID: {outlook_id}")
                                stats["google_to_outlook"]["created"] += 1
                    except Exception as e:
                        print(f"  - Erro ao criar evento no Outlook: {e}")

                # Verificar e processar Expresso independentemente do Outlook
                if hasattr(self, "expresso_sync") and self.expresso_sync:
                    # Verificar se o evento já existe no Expresso (usando o novo método)
                    expresso_events_cache = {}
                    try:
                        expresso_events = self.expresso_sync.obterEventos()
                        for event in expresso_events:
                            if "id" in event:
                                expresso_events_cache[event["id"]] = event
                    except Exception as e:
                        print(f"  - Erro ao obter eventos do Expresso: {e}")
                        # Continuar mesmo com erro
                        
                    # Verificar duplicação usando o novo método
                    exists, expresso_id = self._check_event_already_exists(google_event, 'google', expresso_events_cache, 'expresso')
                    if exists:
                        print(f"  - Evento já existe no Expresso com ID: {expresso_id}")
                        self.db.map_events(
                            google_id=event_id,
                            expresso_id=expresso_id,
                            origem="match",
                        )
                        self._store_event_mapping(
                            google_id=event_id, expresso_id=expresso_id
                        )
                        expresso_match_found = True

                    # Se não encontrou match, criar novo evento no Expresso
                    if not expresso_match_found:
                        try:
                            expresso_event = (
                                self.expresso_sync._format_google_to_expresso(
                                    google_event
                                )
                            )
                            if expresso_event:
                                print(
                                    f"  - Criando no Expresso: {expresso_event.get('titulo', 'Sem título')}"
                                )
                                result = self.expresso_sync.create_event(expresso_event)
                                expresso_id = result.get("id")
                                if expresso_id:
                                    # Aqui também precisamos armazenar o evento no banco antes de mapear
                                    self.db.store_expresso_event(result)
                                    # Depois mapear os eventos
                                    self.db.map_events(
                                        google_id=event_id,
                                        expresso_id=expresso_id,
                                        origem="google",
                                    )
                                    self._store_event_mapping(
                                        google_id=event_id, expresso_id=expresso_id
                                    )
                                    print(
                                        f"  - Criado no Expresso com ID: {expresso_id}"
                                    )
                                    stats["google_to_expresso"]["created"] += 1
                        except Exception as e:
                            print(f"  - Erro ao criar evento no Expresso: {e}")
        else:
            print(f"  - Evento já sincronizado anteriormente")

        # Processar eventos atualizados no Google
        for event_id, google_event in changes["google"]["updated"].items():
            if event_id in self.google_to_outlook_map:
                outlook_id = self.google_to_outlook_map[event_id]
                try:
                    outlook_event = self._format_google_to_outlook(google_event)
                    if outlook_event:
                        print(
                            f"Atualizando no Outlook: {outlook_event.get('subject', 'Sem título')}"
                        )
                        self.outlook_sync.update_event(outlook_id, outlook_event)
                        stats["google_to_outlook"]["updated"] += 1
                except Exception as e:
                    print(f"Erro ao atualizar evento no Outlook: {e}")

        # Processar eventos excluídos no Google
        for event_id, google_event in changes["google"]["deleted"].items():
            print(
                f"Detectada exclusão no Google: {google_event.get('summary', 'Sem título')} (ID: {event_id})"
            )

            # Verificar mapeamentos para Outlook
            if event_id in self.google_to_outlook_map:
                outlook_id = self.google_to_outlook_map[event_id]
                try:
                    print(f"  - Excluindo do Outlook: {outlook_id}")
                    self.outlook_sync.delete_event(outlook_id)
                    stats["google_to_outlook"]["deleted"] += 1
                except Exception as e:
                    print(f"  - Erro ao excluir evento do Outlook: {e}")

            # Verificar mapeamentos para Expresso
            if (
                hasattr(self, "expresso_sync")
                and self.expresso_sync
                and event_id in self.google_to_expresso_map
            ):
                expresso_id = self.google_to_expresso_map[event_id]
                try:
                    print(f"  - Excluindo do Expresso: {expresso_id}")
                    self.expresso_sync.delete_event(expresso_id)
                    stats["google_to_expresso"]["deleted"] += 1
                except Exception as e:
                    print(f"  - Erro ao excluir evento do Expresso: {e}")

            # Remover todos os mapeamentos relacionados
            self._remove_all_mappings(google_id=event_id)

        # Processar eventos adicionados no Outlook
        for event_id, outlook_event in changes["outlook"]["added"].items():
            print(
                f"Processando novo evento do Outlook: {outlook_event.get('subject', 'Sem título')} (ID: {event_id})"
            )

            # Verificar se o evento já foi processado nesta sessão
            if event_id in events_being_synced:
                print(f"  - Evento já está sendo processado nesta sessão, ignorando")
                continue
            
            # Adicionar à lista de eventos sendo processados
            events_being_synced.add(event_id)
            
            # Verificar se este evento já tem um mapeamento
            if event_id in self.outlook_to_google_map:
                print(
                    f"  - Já sincronizado com o Google com ID: {self.outlook_to_google_map[event_id]}"
                )
                continue

            # Verificar se este é um evento recentemente sincronizado do Google usando o novo método
            exists, google_id = self._check_event_already_exists(outlook_event, 'outlook', self.google_events_cache, 'google')
            if exists:
                print(f"  - Evento já existe no Google com ID: {google_id}")
                self._store_event_mapping(google_id=google_id, outlook_id=event_id)
                is_synced_from_google = True
            else:
                is_synced_from_google = False
                try:
                    google_event = self._format_outlook_to_google(outlook_event)
                    if google_event:
                        print(
                            f"  - Criando no Google: {google_event.get('summary', 'Sem título')}"
                        )
                        # Log detalhado do evento formatado para depuração
                        print(f"  - Dados do evento para o Google: {json.dumps(google_event, default=str)}")
                        
                        result = self.google_sync.create_event(google_event)
                        google_id = result.get("id")
                        if google_id:
                            # Primeiro armazenar o evento no banco
                            self.db.store_google_event(result)
                            # Depois mapear os eventos
                            self._store_event_mapping(
                                google_id=google_id, outlook_id=event_id
                            )
                            print(f"  - Criado no Google com ID: {google_id}")
                            stats["outlook_to_google"]["created"] += 1
                        else:
                            print("  - ERRO: Não foi possível obter ID do evento criado no Google")
                except Exception as e:
                    print(f"  - Erro ao criar evento no Google: {str(e)}")
                    if hasattr(e, 'response') and hasattr(e.response, 'text'):
                        print(f"  - Detalhes do erro: {e.response.text}")

            # Adicionar sincronização com Expresso
            if hasattr(self, "expresso_sync") and self.expresso_sync:
                # Verificar se este evento já tem um mapeamento com Expresso
                expresso_match_found = False

                # Verificar no banco de dados
                mapped_ids = self.db.get_mapped_ids(event_id, "outlook")
                if (
                    mapped_ids and mapped_ids[1]
                ):  # [1] corresponde ao expresso_id no retorno de get_mapped_ids
                    print(f"  - Já sincronizado com o Expresso com ID: {mapped_ids[1]}")
                    expresso_match_found = True

                if not expresso_match_found:
                    # Verificar se o evento já existe no Expresso usando o novo método
                    expresso_events_cache = {}
                    try:
                        expresso_events = self.expresso_sync.obterEventos()
                        for event in expresso_events:
                            if "id" in event:
                                expresso_events_cache[event["id"]] = event
                        
                        print(f"  - Encontrados {len(expresso_events_cache)} eventos no Expresso para comparação")
                        
                        # Verificar duplicação usando o novo método
                        exists, expresso_id = self._check_event_already_exists(outlook_event, 'outlook', expresso_events_cache, 'expresso')
                        if exists:
                            print(f"  - Evento já existe no Expresso com ID: {expresso_id}")
                            self.db.map_events(
                                outlook_id=event_id,
                                expresso_id=expresso_id,
                                origem="match",
                            )
                            self._store_event_mapping(
                                outlook_id=event_id, expresso_id=expresso_id
                            )
                            expresso_match_found = True
                    except Exception as e:
                        print(f"  - Erro ao obter eventos do Expresso: {str(e)}")
                        # Continuar mesmo com erro, tentando criar o evento no Expresso

                    # Se não encontrou match, criar novo evento no Expresso
                    if not expresso_match_found:
                        try:
                            expresso_event = self._format_outlook_to_expresso(
                                outlook_event
                            )
                            if expresso_event:
                                print(
                                    f"  - Criando no Expresso: {expresso_event.get('titulo', 'Sem título')}"
                                )
                                # Log detalhado para depuração
                                print(f"  - Dados do evento para o Expresso: {json.dumps(expresso_event, default=str)}")
                                
                                result = self.expresso_sync.create_event(expresso_event)
                                expresso_id = result.get("id") if isinstance(result, dict) else result
                                if expresso_id:
                                    # Armazenar o evento no banco
                                    self.db.store_expresso_event(result if isinstance(result, dict) else {"id": expresso_id})
                                    # Mapear os eventos
                                    self.db.map_events(
                                        outlook_id=event_id,
                                        expresso_id=expresso_id,
                                        origem="outlook",
                                    )
                                    self._store_event_mapping(
                                        outlook_id=event_id, expresso_id=expresso_id
                                    )
                                    print(f"  - Criado no Expresso com ID: {expresso_id}")
                                    stats["outlook_to_expresso"]["created"] += 1
                                else:
                                    print("  - ERRO: Não foi possível obter ID do evento criado no Expresso")
                            else:
                                print("  - ERRO: Falha ao formatar evento do Outlook para o Expresso")
                        except Exception as e:
                            print(f"  - Erro ao criar evento no Expresso: {str(e)}")
                            # Continuar a execução mesmo após erro

        # Processar eventos atualizados no Outlook
        for event_id, outlook_event in changes["outlook"]["updated"].items():
            if event_id in self.outlook_to_google_map:
                google_id = self.outlook_to_google_map[event_id]
                try:
                    google_event = self._format_outlook_to_google(outlook_event)
                    if google_event:
                        print(
                            f"Atualizando no Google: {google_event.get('summary', 'Sem título')}"
                        )
                        self.google_sync.update_event(google_id, google_event)
                        stats["outlook_to_google"]["updated"] += 1
                except Exception as e:
                    print(f"Erro ao atualizar evento no Google: {e}")

        # Processar eventos excluídos no Outlook
        for event_id, outlook_event in changes["outlook"]["deleted"].items():
            if event_id in self.outlook_to_google_map:
                google_id = self.outlook_to_google_map[event_id]
                try:
                    print(
                        f"Excluindo do Google: {outlook_event.get('subject', 'Sem título')}"
                    )
                    self.google_sync.delete_event(google_id)
                    # Remover dos mapeamentos
                    del self.google_to_outlook_map[google_id]
                    del self.outlook_to_google_map[event_id]
                    stats["outlook_to_google"]["deleted"] += 1
                except Exception as e:
                    print(f"Erro ao excluir evento do Google: {e}")

        # Processar eventos no Expresso, se existir
        if hasattr(self, "expresso_sync") and self.expresso_sync:
            # Obter cache de eventos do Expresso
            expresso_events_cache = {}
            expresso_events = self.expresso_sync.obterEventos()
            for event in expresso_events:
                if "id" in event:
                    expresso_events_cache[event["id"]] = event

            # Detectar eventos adicionados no Expresso desde a última sincronização
            expresso_added = {}
            expresso_mapped_ids = {}

            # Mapear eventos existentes do Expresso para evitar duplicações
            for event_id, event in expresso_events_cache.items():
                # Verificar se já está mapeado
                mapped_ids = self.db.get_mapped_ids(event_id, "expresso")
                if mapped_ids:
                    expresso_mapped_ids[event_id] = mapped_ids
                    continue

                # Se não estiver mapeado, é um novo evento
                expresso_added[event_id] = event

            # Processar novos eventos do Expresso
            for event_id, expresso_event in expresso_added.items():
                print(
                    f"Processando novo evento do Expresso: {expresso_event.get('titulo', 'Sem título')}"
                )

                # Tentar criar no Google
                try:
                    google_event = self.expresso_sync._format_expresso_to_google(
                        expresso_event
                    )
                    if google_event:
                        print(
                            f"  - Criando no Google: {google_event.get('summary', 'Sem título')}"
                        )
                        result = self.google_sync.create_event(google_event)
                        google_id = result.get("id")
                        if google_id:
                            # Primeiro armazenar evento na tabela google_events
                            self.db.store_google_event(result)

                            # Depois mapear os eventos
                            self.db.map_events(
                                expresso_id=event_id,
                                google_id=google_id,
                                origem="expresso",
                            )
                            print(f"  - Criado no Google com ID: {google_id}")
                            stats["expresso_to_google"]["created"] += 1
                except Exception as e:
                    print(f"  - Erro ao criar evento do Expresso no Google: {e}")

                # Tentar criar no Outlook
                try:
                    outlook_event = self.expresso_sync._format_expresso_to_outlook(
                        expresso_event
                    )
                    if outlook_event:
                        print(
                            f"  - Criando no Outlook: {outlook_event.get('subject', 'Sem título')}"
                        )
                        result = self.outlook_sync.create_event(outlook_event)
                        outlook_id = result.get("id")
                        if outlook_id:
                            # Primeiro armazenar evento na tabela outlook_events
                            self.db.store_outlook_event(result)

                            # Depois mapear os eventos
                            self.db.map_events(
                                expresso_id=event_id,
                                outlook_id=outlook_id,
                                origem="expresso",
                            )
                            print(f"  - Criado no Outlook com ID: {outlook_id}")
                            stats["expresso_to_outlook"]["created"] += 1
                except Exception as e:
                    print(f"  - Erro ao criar evento do Expresso no Outlook: {e}")

            # Detectar eventos excluídos no Expresso
            if hasattr(self, "expresso_sync") and self.expresso_sync:
                # Obter IDs de eventos atuais do Expresso
                expresso_current_ids = set()
                expresso_events = self.expresso_sync.obterEventos()
                for event in expresso_events:
                    if "id" in event:
                        expresso_current_ids.add(event["id"])
                
                # Obter todos os mapeamentos relacionados ao Expresso
                expresso_mappings = {}
                for mapping in self.db.get_all_mappings():
                    _, outlook_id, google_id, expresso_id, _, _ = mapping
                    if expresso_id:
                        expresso_mappings[expresso_id] = (google_id, outlook_id)
                
                # Detectar IDs que estavam mapeados mas não estão mais nos eventos atuais
                for expresso_id, (google_id, outlook_id) in expresso_mappings.items():
                    if expresso_id not in expresso_current_ids:
                        print(f"Detectada exclusão no Expresso ID: {expresso_id}")
                        
                        # Excluir do Google se tiver mapeamento
                        if google_id:
                            try:
                                print(f"  - Excluindo do Google: {google_id}")
                                self.google_sync.delete_event(google_id)
                                stats["expresso_to_google"]["deleted"] += 1
                            except Exception as e:
                                print(f"  - Erro ao excluir evento do Google: {e}")
                        
                        # Excluir do Outlook se tiver mapeamento
                        if outlook_id:
                            try:
                                print(f"  - Excluindo do Outlook: {outlook_id}")
                                self.outlook_sync.delete_event(outlook_id)
                                stats["expresso_to_outlook"]["deleted"] += 1
                            except Exception as e:
                                print(f"  - Erro ao excluir evento do Outlook: {e}")
                        
                        # Remover todos os mapeamentos relacionados
                        self._remove_all_mappings(expresso_id=expresso_id)
                        
                        # Marcar como excluído no banco de dados
                        self.db.mark_event_deleted(expresso_id, "expresso")

        # Detectar e processar eventos atualizados no Expresso
        if hasattr(self, "expresso_sync") and self.expresso_sync:
            expresso_events_cache = {}
            expresso_events = self.expresso_sync.obterEventos()
            
            # Mapear todos os eventos do Expresso para verificar atualizações
            for event in expresso_events:
                if "id" in event:
                    expresso_id = event["id"]
                    expresso_events_cache[expresso_id] = event
                    
                    # Verificar se este evento tem mapeamento no banco
                    mapped_ids = self.db.get_mapped_ids(expresso_id, "expresso")
                    if mapped_ids and (mapped_ids[0] or mapped_ids[1]): # Se tiver Google ID ou Outlook ID mapeado
                        google_id = mapped_ids[1]  # No retorno do get_mapped_ids para expresso, [1] é o Google ID
                        outlook_id = mapped_ids[0]  # [0] é o Outlook ID
                        
                        # Verificar mudanças no evento do Expresso
                        # Como não temos um campo de "lastModified" no Expresso, vamos usar um método mais manual
                        # Podemos obter o evento do banco de dados e comparar campos relevantes
                        
                        # Atualizar no Google
                        if google_id:
                            try:
                                # Converter o evento do Expresso para o formato do Google
                                google_event = self.expresso_sync._format_expresso_to_google(event)
                                if google_event:
                                    # Adicionar o campo ID para que a API do Google saiba qual evento atualizar
                                    google_event["id"] = google_id
                                    print(f"Atualizando no Google: {google_event.get('summary', 'Sem título')}")
                                    self.google_sync.update_event(google_id, google_event)
                                    stats["expresso_to_google"]["updated"] += 1
                            except Exception as e:
                                print(f"Erro ao atualizar evento no Google: {e}")
                        
                        # Atualizar no Outlook
                        if outlook_id:
                            try:
                                # Converter o evento do Expresso para o formato do Outlook
                                outlook_event = self.expresso_sync._format_expresso_to_outlook(event)
                                if outlook_event:
                                    print(f"Atualizando no Outlook: {outlook_event.get('subject', 'Sem título')}")
                                    self.outlook_sync.update_event(outlook_id, outlook_event)
                                    stats["expresso_to_outlook"]["updated"] += 1
                            except Exception as e:
                                print(f"Erro ao atualizar evento no Outlook: {e}")

        # Após sincronização completa:
        events_being_synced.clear()

        return stats

    def _is_event_updated(self, current_event, cached_event):
        """Verifica se um evento foi atualizado comparando campos relevantes"""
        # Para Google
        if "updated" in current_event and "updated" in cached_event:
            return current_event["updated"] != cached_event["updated"]

        # Para Outlook
        if (
            "lastModifiedDateTime" in current_event
            and "lastModifiedDateTime" in cached_event
        ):
            return (
                current_event["lastModifiedDateTime"]
                != cached_event["lastModifiedDateTime"]
            )

        # Comparação manual de campos importantes
        important_fields = [
            "summary",
            "subject",
            "start",
            "end",
            "location",
            "description",
        ]

        for field in important_fields:
            if field in current_event and field in cached_event:
                # Para campos aninhados como start e end
                if isinstance(current_event[field], dict) and isinstance(
                    cached_event[field], dict
                ):
                    # Comparar dateTime para campos de data/hora
                    if (
                        "dateTime" in current_event[field]
                        and "dateTime" in cached_event[field]
                    ):
                        if (
                            current_event[field]["dateTime"]
                            != cached_event[field]["dateTime"]
                        ):
                            return True
                # Para campos simples
                elif current_event[field] != cached_event[field]:
                    return True

        return False

    def _events_match(self, google_event, outlook_event):
        """Verifica se um evento do Google corresponde a um evento do Outlook"""
        # Primeiro verificar se temos IDs mapeados no banco de dados
        if "id" in google_event and "id" in outlook_event:
            google_id = google_event["id"]
            outlook_id = outlook_event["id"]
            
            # Verificar se já existe um mapeamento no banco de dados
            mapped_ids = self.db.get_mapped_ids(google_id, "google")
            if mapped_ids and mapped_ids[0] == outlook_id:
                print(f"Match por ID no banco: Google {google_id} -> Outlook {outlook_id}")
                return True
            
            mapped_ids = self.db.get_mapped_ids(outlook_id, "outlook")
            if mapped_ids and mapped_ids[0] == google_id:
                print(f"Match por ID no banco: Outlook {outlook_id} -> Google {google_id}")
                return True
            
            # Verificar se há IDs externos armazenados nos eventos
            # Outlook em Google
            if "extendedProperties" in google_event and "private" in google_event["extendedProperties"]:
                private_props = google_event["extendedProperties"]["private"]
                if "outlook_id" in private_props and private_props["outlook_id"] == outlook_id:
                    print(f"Match por ID externo: Google contém Outlook ID {outlook_id}")
                    return True
            
            # Google em Outlook (depende de como os IDs são armazenados no Outlook)
            # Este é apenas um exemplo, pode não ser aplicável dependendo da API
            if "singleValueExtendedProperties" in outlook_event:
                for prop in outlook_event["singleValueExtendedProperties"]:
                    if prop.get("name") == "google_id" and prop.get("value") == google_id:
                        print(f"Match por ID externo: Outlook contém Google ID {google_id}")
                        return True
        
        # Se não encontrar mapeamento por ID, continuar com a verificação por conteúdo
        
        # Comparar título
        google_title = google_event.get("summary", "").strip()
        outlook_title = outlook_event.get("subject", "").strip()

        if not google_title or not outlook_title:
            return False

        title_match = google_title.lower() == outlook_title.lower()
        if not title_match:
            # Se os títulos não correspondem exatamente, podemos verificar se um contém o outro
            # para eventos que podem ter sido editados ligeiramente
            if (google_title.lower() in outlook_title.lower() or 
                outlook_title.lower() in google_title.lower()) and (
                len(google_title) > 5 and len(outlook_title) > 5):  # Evitar falsos positivos com títulos muito curtos
                title_match = True
                # Neste caso, os títulos são similares mas não idênticos
                print(f"Títulos similares: '{google_title}' e '{outlook_title}'")
                # Neste caso, precisamos verificar a data/hora com maior rigor
                # para ter certeza de que é o mesmo evento
            else:
                return False
        
        # Comparar data/hora com maior precisão
        google_start = google_event.get("start", {}).get("dateTime", "")
        outlook_start = outlook_event.get("start", {}).get("dateTime", "")
        
        # Verificar se ambos são eventos de dia inteiro
        google_all_day = "date" in google_event.get("start", {}) and "dateTime" not in google_event.get("start", {})
        outlook_all_day = "date" in outlook_event.get("start", {}) and "dateTime" not in outlook_event.get("start", {})
        
        # Se ambos forem de dia inteiro, comparar as datas
        if google_all_day and outlook_all_day:
            google_date = google_event.get("start", {}).get("date", "")
            outlook_date = outlook_event.get("start", {}).get("date", "")
            if google_date and outlook_date:
                date_match = google_date == outlook_date
                if date_match:
                    print(f"Match de eventos de dia inteiro: {google_date}")
                    return True
                else:
                    return False
        
        # Se chegou aqui, pelo menos um dos eventos não é de dia inteiro
        if google_start and outlook_start:
            try:
                google_dt = datetime.fromisoformat(google_start.replace("Z", "+00:00"))
                outlook_dt = datetime.fromisoformat(outlook_start.replace("Z", "+00:00"))
                
                # Converter para UTC para comparação adequada
                if hasattr(google_dt, "tzinfo") and google_dt.tzinfo:
                    google_dt = google_dt.astimezone(timezone.utc)
                if hasattr(outlook_dt, "tzinfo") and outlook_dt.tzinfo:
                    outlook_dt = outlook_dt.astimezone(timezone.utc)
                
                # Tolerância de 5 minutos (300 segundos)
                time_diff = abs((google_dt - outlook_dt).total_seconds())
                if time_diff > 300:
                    print(f"Horários diferem por {time_diff} segundos")
                    return False
                
                # Se chegou aqui, título e horário correspondem
                print(f"Match de evento com horário específico: '{google_title}' em {google_dt}")
                return True
            except (ValueError, TypeError) as e:
                print(f"Erro ao comparar datas: {e}")
                # Se houver erro na conversão de data/hora, verificar só o título
                return title_match
        
        # Se chegou aqui e nenhuma das verificações definiu um match ou não-match,
        # retornar o resultado da verificação de título
        return title_match

    def start_realtime_sync(self, interval=20, cleanup_interval=86400, days_to_keep=0):
        """
        Inicia sincronização em tempo real a cada 'interval' segundos.

        Args:
            interval (int): Intervalo em segundos entre cada sincronização.
            cleanup_interval (int): Intervalo em segundos para limpar o banco de dados.
                                   Padrão é 86400 (1 dia).
            days_to_keep (int): Dias no passado para manter no banco de dados.
                               0 = mantém apenas eventos de hoje em diante.
        """
        print(
            f"\n=== Iniciando sincronização em tempo real a cada {interval} segundos ==="
        )
        print("Monitorando apenas mudanças (criações, atualizações, exclusões)")
        print(
            f"Limpeza automática do banco de dados a cada {cleanup_interval/3600} horas"
        )
        print("Pressione Ctrl+C para interromper")

        # Executar limpeza inicial
        self.cleanup_database(days_to_keep)

        # Inicializar caches e mapeamentos
        self._update_caches()
        print("Estado inicial dos calendários armazenado. Monitorando mudanças...")

        # Controle de tempo para limpeza periódica
        last_cleanup_time = time.time()

        try:
            while True:
                start_time = time.time()

                # Atualizar a página do Expresso antes de cada sincronização
                if hasattr(self, "expresso_sync") and self.expresso_sync and self.expresso_sync.driver:
                    try:
                        print("Atualizando página do Expresso...")
                        self.expresso_sync.selecionarCalendario()
                    except Exception as e:
                        print(f"Erro ao atualizar página do Expresso: {e}")

                # Verificar se é hora de limpar o banco de dados
                current_time = time.time()
                if current_time - last_cleanup_time >= cleanup_interval:
                    print("\n=== Executando limpeza periódica do banco de dados ===")
                    self.cleanup_database(days_to_keep)
                    last_cleanup_time = current_time

                # Sincronizar apenas mudanças
                stats = self.sync_changes_only()

                # Resumo das operações
                total_ops = sum(sum(category.values()) for category in stats.values())

                if total_ops > 0:
                    print(
                        f"\n[{datetime.now().strftime('%H:%M:%S')}] Mudanças sincronizadas:"
                    )
                    print(
                        f"- Google → Outlook: {stats['google_to_outlook']['created']} criados, "
                        f"{stats['google_to_outlook']['updated']} atualizados, "
                        f"{stats['google_to_outlook']['deleted']} excluídos"
                    )
                    print(
                        f"- Outlook → Google: {stats['outlook_to_google']['created']} criados, "
                        f"{stats['outlook_to_google']['updated']} atualizados, "
                        f"{stats['outlook_to_google']['deleted']} excluídos"
                    )
                else:
                    print(
                        f"[{datetime.now().strftime('%H:%M:%S')}] Nenhuma mudança detectada"
                    )

                # Calcular tempo de espera
                elapsed = time.time() - start_time
                wait_time = max(0, interval - elapsed)

                if wait_time > 0:
                    print(f"Aguardando {wait_time:.1f} segundos...")
                    time.sleep(wait_time)

        except KeyboardInterrupt:
            print("\n=== Sincronização em tempo real interrompida pelo usuário ===")
        finally:
            # Garantir que o banco de dados seja fechado corretamente
            if hasattr(self, "db"):
                self.db.close()
                print("Conexão com o banco de dados fechada.")

    # Add these methods to your CalendarSynchronizer class

    def _format_google_to_outlook(self, google_event):
        """Converte um evento do Google para o formato do Outlook"""
        if not google_event.get("summary") or not google_event.get("start"):
            return None

        outlook_event = {
            "subject": google_event.get("summary"),
            "body": {
                "contentType": "HTML",
                "content": google_event.get(
                    "description", "Evento sincronizado do Google Calendar"
                ),
            },
            "start": {
                "dateTime": google_event.get("start", {}).get("dateTime", ""),
                "timeZone": google_event.get("start", {}).get("timeZone", "UTC"),
            },
            "end": {
                "dateTime": google_event.get("end", {}).get("dateTime", ""),
                "timeZone": google_event.get("end", {}).get("timeZone", "UTC"),
            },
        }

        if google_event.get("location"):
            outlook_event["location"] = {"displayName": google_event.get("location")}

        return outlook_event

    def _format_outlook_to_google(self, outlook_event):
        """Converte um evento do Outlook para o formato do Google"""
        if not outlook_event.get("subject") or not outlook_event.get("start"):
            return None

        google_event = {
            "summary": outlook_event.get("subject"),
            "description": outlook_event.get("body", {}).get(
                "content", "Evento sincronizado do Outlook Calendar"
            ),
            "start": {
                "dateTime": outlook_event.get("start", {}).get("dateTime", ""),
                "timeZone": outlook_event.get("start", {}).get("timeZone", "UTC"),
            },
            "end": {
                "dateTime": outlook_event.get("end", {}).get("dateTime", ""),
                "timeZone": outlook_event.get("end", {}).get("timeZone", "UTC"),
            },
        }

        if outlook_event.get("location"):
            google_event["location"] = outlook_event.get("location", {}).get(
                "displayName", ""
            )

        return google_event

    def _store_event_mapping(self, google_id=None, outlook_id=None, expresso_id=None):
        """Armazena o mapeamento entre IDs de eventos"""
        # Imprimir informações de mapeamento
        mapping_info = []
        if google_id:
            mapping_info.append(f"Google ID {google_id}")
        if outlook_id:
            mapping_info.append(f"Outlook ID {outlook_id}")
        if expresso_id:
            mapping_info.append(f"Expresso ID {expresso_id}")

        print(f"Mapeando evento: {' <-> '.join(mapping_info)}")

        # Verificar se já existe um mapeamento parcial que possa ser completado
        existing_mapping = None
        origem = "sync"

        if google_id:
            google_mapping = self.db.get_mapped_ids(google_id, "google")
            if google_mapping and (google_mapping[0] or google_mapping[1]):
                existing_mapping = google_mapping
                outlook_id = outlook_id or google_mapping[0]
                expresso_id = expresso_id or google_mapping[1]

        if outlook_id and not existing_mapping:
            outlook_mapping = self.db.get_mapped_ids(outlook_id, "outlook")
            if outlook_mapping and (outlook_mapping[0] or outlook_mapping[1]):
                existing_mapping = outlook_mapping
                google_id = google_id or outlook_mapping[0]
                expresso_id = expresso_id or outlook_mapping[1]

        if expresso_id and not existing_mapping:
            expresso_mapping = self.db.get_mapped_ids(expresso_id, "expresso")
            if expresso_mapping and (expresso_mapping[0] or expresso_mapping[1]):
                existing_mapping = expresso_mapping
                google_id = google_id or expresso_mapping[1]
                outlook_id = outlook_id or expresso_mapping[0]

        # Armazenar no banco de dados com todos os IDs disponíveis
        self.db.map_events(
            google_id=google_id,
            outlook_id=outlook_id,
            expresso_id=expresso_id,
            origem=origem,
        )

        # Também manter nos mapas em memória para compatibilidade
        if google_id and outlook_id:
            self.google_to_outlook_map[google_id] = outlook_id
            self.outlook_to_google_map[outlook_id] = google_id

        # Adicionar mapeamentos para Expresso
        if google_id and expresso_id:
            self.google_to_expresso_map[google_id] = expresso_id
            self.expresso_to_google_map[expresso_id] = google_id

        if outlook_id and expresso_id:
            self.outlook_to_expresso_map[outlook_id] = expresso_id
            self.expresso_to_outlook_map[expresso_id] = outlook_id

    def cleanup_database(self, days_to_keep=0):
        """
        Limpa eventos antigos do banco de dados.

        Args:
            days_to_keep (int): Número de dias passados a manter.
                               0 = mantém apenas eventos de hoje em diante.
        """
        print(f"\n=== Iniciando limpeza do banco de dados ===")
        print(f"Mantendo eventos de hoje e futuros (+ {days_to_keep} dias no passado)")

        result = self.db.cleanup_old_events(days_to_keep)

        # Após limpar o banco de dados, atualizar os caches em memória
        self._update_caches()

        return result

    def _check_event_already_exists(self, event, source_type, target_cache, target_type=None):
        """
        Verifica se um evento já existe no cache de destino
        
        Args:
            event: Evento a verificar
            source_type: Tipo do evento origem ('google', 'outlook', 'expresso')
            target_cache: Cache de eventos de destino para verificação
            target_type: Tipo do evento de destino ('google', 'outlook', 'expresso')
            
        Returns:
            (bool, str): Tupla com (existe?, id do evento correspondente)
        """
        print(f"Verificando duplicação de evento tipo '{source_type}' em '{target_type or 'destino'}'...")
        
        if not target_type:
            if source_type == 'google':
                target_type = 'outlook'
            elif source_type == 'outlook':
                target_type = 'google'
            else:
                target_type = 'google'  # Default para Expresso
        
        # Se não tiver título, não podemos comparar adequadamente
        event_title = None
        if source_type == 'google':
            event_title = event.get('summary', '').strip()
        elif source_type == 'outlook':
            event_title = event.get('subject', '').strip()
        elif source_type == 'expresso':
            event_title = event.get('titulo', '').strip()
            
        if not event_title:
            print(f"  - Evento sem título, não é possível verificar duplicação")
            return False, None
        else:
            print(f"  - Verificando evento com título: '{event_title}'")
            
        # Extração de data/hora depende do tipo de evento
        event_datetime = None
        try:
            if source_type == 'google':
                if 'dateTime' in event.get('start', {}):
                    dt_str = event['start']['dateTime']
                    event_datetime = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
                    print(f"  - Data/hora do evento: {event_datetime}")
                elif 'date' in event.get('start', {}):
                    # Evento de dia inteiro
                    dt_str = event['start']['date']
                    event_datetime = datetime.fromisoformat(dt_str)
                    print(f"  - Data do evento (dia inteiro): {event_datetime}")
            elif source_type == 'outlook':
                if 'dateTime' in event.get('start', {}):
                    dt_str = event['start']['dateTime']
                    event_datetime = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
                    print(f"  - Data/hora do evento: {event_datetime}")
                elif 'date' in event.get('start', {}):
                    # Evento de dia inteiro
                    dt_str = event['start']['date']
                    event_datetime = datetime.fromisoformat(dt_str)
                    print(f"  - Data do evento (dia inteiro): {event_datetime}")
            elif source_type == 'expresso':
                if 'data' in event and 'inicio' in event:
                    data = event['data']
                    hora = event['inicio']
                    if data and hora and ':' in hora:
                        dia, mes, ano = data.split('/')
                        hora, minuto = hora.split(':')
                        # Criar um datetime aware com timezone UTC
                        event_datetime = datetime(int(ano), int(mes), int(dia), int(hora), int(minuto), tzinfo=timezone.utc)
                        print(f"  - Data/hora do evento: {event_datetime}")
        except (ValueError, TypeError, IndexError) as e:
            print(f"  - Erro ao extrair data/hora: {e}")
            # Continuar apenas com verificação de título
            
        print(f"  - Cache destino ({target_type}) tem {len(target_cache)} eventos para comparar")
        
        # Verificar cada evento no cache de destino
        for target_id, target_event in target_cache.items():
            # Extrair título do evento de destino
            target_title = None
            if target_type == 'google':
                target_title = target_event.get('summary', '').strip()
            elif target_type == 'outlook':
                target_title = target_event.get('subject', '').strip()
            elif target_type == 'expresso':
                target_title = target_event.get('titulo', '').strip()
                
            if not target_title:
                continue
                
            # Verificar correspondência de título
            title_match = event_title.lower() == target_title.lower()
            if not title_match:
                # Verificar títulos similares
                if (event_title.lower() in target_title.lower() or 
                    target_title.lower() in event_title.lower()) and (
                    len(event_title) > 5 and len(target_title) > 5):
                    title_match = True
                    print(f"  - Títulos similares: '{event_title}' e '{target_title}'")
                else:
                    continue
            else:
                print(f"  - Títulos idênticos: '{event_title}' e '{target_title}'")
                    
            # Se não temos data/hora para comparar, retornar com base apenas no título
            if not event_datetime:
                print(f"  - Match por título: '{event_title}' = '{target_title}'")
                return True, target_id
                
            # Extrair data/hora do evento de destino
            target_datetime = None
            try:
                if target_type == 'google':
                    if 'dateTime' in target_event.get('start', {}):
                        dt_str = target_event['start']['dateTime']
                        target_datetime = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
                        print(f"  - Data/hora do evento destino: {target_datetime}")
                    elif 'date' in target_event.get('start', {}):
                        dt_str = target_event['start']['date']
                        target_datetime = datetime.fromisoformat(dt_str)
                        print(f"  - Data do evento destino (dia inteiro): {target_datetime}")
                elif target_type == 'outlook':
                    if 'dateTime' in target_event.get('start', {}):
                        dt_str = target_event['start']['dateTime']
                        target_datetime = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
                        print(f"  - Data/hora do evento destino: {target_datetime}")
                    elif 'date' in target_event.get('start', {}):
                        dt_str = target_event['start']['date']
                        target_datetime = datetime.fromisoformat(dt_str)
                        print(f"  - Data do evento destino (dia inteiro): {target_datetime}")
                elif target_type == 'expresso':
                    if 'data' in target_event and 'inicio' in target_event:
                        data = target_event['data']
                        hora = target_event['inicio']
                        if data and hora and ':' in hora:
                            dia, mes, ano = data.split('/')
                            hora, minuto = hora.split(':')
                            # Criar um datetime aware com timezone UTC
                            target_datetime = datetime(int(ano), int(mes), int(dia), int(hora), int(minuto), tzinfo=timezone.utc)
                            print(f"  - Data/hora do evento destino: {target_datetime}")
            except (ValueError, TypeError, IndexError) as e:
                print(f"  - Erro ao extrair data/hora do evento de destino: {e}")
                # Se não conseguimos extrair a data do destino, consideramos o match pelo título
                print(f"  - Match apenas por título (erro na data): '{event_title}'")
                return True, target_id
                
            # Se não conseguimos extrair a data do evento de destino
            if not target_datetime:
                print(f"  - Match apenas por título (destino sem data): '{event_title}'")
                return True, target_id
                
            # Verificar se ambos os datetimes são do mesmo tipo (aware ou naive)
            # Se um for aware e outro naive, converter para que sejam comparáveis
            if (event_datetime.tzinfo is None and target_datetime.tzinfo is not None):
                # event_datetime é naive e target_datetime é aware
                # Converter event_datetime para aware (UTC)
                event_datetime = event_datetime.replace(tzinfo=timezone.utc)
            elif (event_datetime.tzinfo is not None and target_datetime.tzinfo is None):
                # event_datetime é aware e target_datetime é naive
                # Converter target_datetime para aware (UTC)
                target_datetime = target_datetime.replace(tzinfo=timezone.utc)
                
            # Comparar as datas/horas com tolerância de 24 horas (86400 segundos)
            try:
                time_diff = abs((event_datetime - target_datetime).total_seconds())
                if time_diff <= 86400:  # 24 horas
                    print(f"  - Match completo (título e data) para evento: '{event_title}' em {event_datetime}")
                    print(f"  - Diferença de tempo: {time_diff} segundos ({time_diff/3600:.2f} horas)")
                    return True, target_id
                else:
                    print(f"  - Mesmo título mas data/hora muito diferente: {time_diff/3600:.2f} horas de diferença")
            except TypeError as e:
                print(f"  - Erro ao comparar datas: {e}. Assumindo match pelo título.")
                return True, target_id
            
        # Se chegou aqui, nenhum evento correspondente foi encontrado
        print("  - Nenhum evento correspondente encontrado")
        return False, None

    def _find_matching_event_by_id(self, event_id, source_type, target_cache):
        """Tenta encontrar um evento correspondente usando IDs externos armazenados em campos personalizados"""
        # Primeiro verificar no banco de dados
        if source_type == "google":
            mapped_ids = self.db.get_mapped_ids(event_id, "google")
            if mapped_ids and mapped_ids[0]:
                return mapped_ids[0]  # Retorna o ID do Outlook
        elif source_type == "outlook":
            mapped_ids = self.db.get_mapped_ids(event_id, "outlook")
            if mapped_ids and mapped_ids[0]:
                return mapped_ids[0]  # Retorna o ID do Google
        
        # Se não encontrou no banco, procurar por correspondência de conteúdo
        if source_type == "google":
            google_event = self.google_events_cache.get(event_id)
            if google_event:
                for outlook_id, outlook_event in target_cache.items():
                    if self._events_match(google_event, outlook_event):
                        return outlook_id
        elif source_type == "outlook":
            outlook_event = self.outlook_events_cache.get(event_id)
            if outlook_event:
                for google_id, google_event in target_cache.items():
                    if self._events_match(google_event, outlook_event):
                        return google_id
        
        return None
