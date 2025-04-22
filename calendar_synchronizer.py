# calendar_synchronizer.py
from database import DatabaseManager
from datetime import datetime, timedelta
import time

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
            print(
                f"Processando novo evento do Google: {google_event.get('summary', 'Sem título')}"
            )

            # Verificar no banco de dados se este evento já existe em algum mapeamento
            mapped_ids = self.db.get_mapped_ids(event_id, "google")

            if mapped_ids is None or (mapped_ids[0] is None and mapped_ids[1] is None):
                # Verificar correspondência baseada em propriedades
                outlook_match_found = False
                expresso_match_found = False

                # Verificar correspondência com eventos do Outlook
                for outlook_id, outlook_event in self.outlook_events_cache.items():
                    if self._events_match(google_event, outlook_event):
                        print(
                            f"  - Corresponde a um evento existente no Outlook: {outlook_event.get('subject', 'Sem título')}"
                        )
                        self.db.map_events(
                            google_id=event_id, outlook_id=outlook_id, origem="match"
                        )
                        self._store_event_mapping(
                            google_id=event_id, outlook_id=outlook_id
                        )
                        outlook_match_found = True
                        break

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
                    # Primeiramente, verificar se há match com algum evento do Expresso
                    expresso_events_cache = {}
                    expresso_events = self.expresso_sync.obterEventos()
                    for event in expresso_events:
                        if "id" in event:
                            expresso_events_cache[event["id"]] = event

                    # Verificar match com eventos Expresso
                    for expresso_id, expresso_event in expresso_events_cache.items():
                        # Comparar usando critérios similares aos usados em _events_match
                        google_title = google_event.get("summary", "").strip()
                        expresso_title = expresso_event.get("titulo", "").strip()

                        if google_title.lower() == expresso_title.lower():
                            print(
                                f"  - Corresponde a um evento existente no Expresso: {expresso_title}"
                            )
                            self.db.map_events(
                                google_id=event_id,
                                expresso_id=expresso_id,
                                origem="match",
                            )
                            self._store_event_mapping(
                                google_id=event_id, expresso_id=expresso_id
                            )
                            expresso_match_found = True
                            break

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
            print(f"Detectada exclusão no Google: {google_event.get('summary', 'Sem título')} (ID: {event_id})")
            
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
            if hasattr(self, "expresso_sync") and self.expresso_sync and event_id in self.google_to_expresso_map:
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
                f"Processando novo evento do Outlook: {outlook_event.get('subject', 'Sem título')}"
            )

            # Verificar se este evento já tem um mapeamento
            if event_id in self.outlook_to_google_map:
                print(
                    f"  - Já sincronizado com o Google com ID: {self.outlook_to_google_map[event_id]}"
                )
                continue

            # Verificar se este é um evento recentemente sincronizado do Google
            is_synced_from_google = False
            for google_id, google_event in self.google_events_cache.items():
                if self._events_match(google_event, outlook_event):
                    print(
                        f"  - Corresponde a um evento existente no Google: {google_event.get('summary', 'Sem título')}"
                    )
                    self._store_event_mapping(google_id=google_id, outlook_id=event_id)
                    is_synced_from_google = True
                    break

            if not is_synced_from_google:
                try:
                    google_event = self._format_outlook_to_google(outlook_event)
                    if google_event:
                        print(
                            f"  - Criando no Google: {google_event.get('summary', 'Sem título')}"
                        )
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
                except Exception as e:
                    print(f"  - Erro ao criar evento no Google: {e}")

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
                    # Verificar correspondência com eventos do Expresso
                    expresso_events_cache = {}
                    expresso_events = self.expresso_sync.obterEventos()
                    for event in expresso_events:
                        if "id" in event:
                            expresso_events_cache[event["id"]] = event

                    for expresso_id, expresso_event in expresso_events_cache.items():
                        # Comparar título
                        outlook_title = outlook_event.get("subject", "").strip()
                        expresso_title = expresso_event.get("titulo", "").strip()

                        if outlook_title.lower() == expresso_title.lower():
                            print(
                                f"  - Corresponde a um evento existente no Expresso: {expresso_title}"
                            )
                            self.db.map_events(
                                outlook_id=event_id,
                                expresso_id=expresso_id,
                                origem="match",
                            )
                            self._store_event_mapping(
                                outlook_id=event_id, expresso_id=expresso_id
                            )
                            expresso_match_found = True
                            break

                # Se não encontrou match, criar novo evento no Expresso
                if not expresso_match_found:
                    try:
                        expresso_event = self.expresso_sync._format_outlook_to_expresso(
                            outlook_event
                        )
                        if expresso_event:
                            print(
                                f"  - Criando no Expresso: {expresso_event.get('titulo', 'Sem título')}"
                            )
                            result = self.expresso_sync.create_event(expresso_event)
                            expresso_id = result.get("id")
                            if expresso_id:
                                # Armazenar o evento no banco
                                self.db.store_expresso_event(result)
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
                    except Exception as e:
                        print(f"  - Erro ao criar evento no Expresso: {e}")

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

            # Também precisamos adicionar código para atualizar eventos do Expresso no Google e Outlook
            # E código para deletar eventos excluídos do Expresso nos outros sistemas
            # (Código omitido por brevidade, mas seguiria a mesma estrutura dos processos de Google/Outlook)

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
        # Comparar título
        google_title = google_event.get("summary", "").strip()
        outlook_title = outlook_event.get("subject", "").strip()

        if not google_title or not outlook_title:
            return False

        title_match = google_title.lower() == outlook_title.lower()
        if not title_match:
            return False

        # Comparar data/hora de início
        google_start = google_event.get("start", {}).get("dateTime", "")
        outlook_start = outlook_event.get("start", {}).get("dateTime", "")

        if not google_start or not outlook_start:
            return False

        try:
            google_dt = datetime.fromisoformat(google_start.replace("Z", "+00:00"))
            outlook_dt = datetime.fromisoformat(outlook_start.replace("Z", "+00:00"))
            time_diff = abs((google_dt - outlook_dt).total_seconds())
            if time_diff > 300:  # 5 minutos de tolerância
                return False
        except (ValueError, TypeError):
            return False

        return True

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
            google_id=google_id, outlook_id=outlook_id, expresso_id=expresso_id, origem=origem
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

    def _events_match_expresso(self, evento1, evento2, source1, source2):
        """
        Verifica se dois eventos correspondem entre si baseado nos tipos de fonte

        Args:
            evento1: Primeiro evento
            evento2: Segundo evento
            source1: Tipo do primeiro evento ('google', 'outlook', 'expresso')
            source2: Tipo do segundo evento ('google', 'outlook', 'expresso')
        """
        # Extrair títulos dos eventos
        if source1 == "google":
            title1 = evento1.get("summary", "").strip()
        elif source1 == "outlook":
            title1 = evento1.get("subject", "").strip()
        elif source1 == "expresso":
            title1 = evento1.get("titulo", "").strip()
        else:
            return False

        if source2 == "google":
            title2 = evento2.get("summary", "").strip()
        elif source2 == "outlook":
            title2 = evento2.get("subject", "").strip()
        elif source2 == "expresso":
            title2 = evento2.get("titulo", "").strip()
        else:
            return False

        # Comparar títulos
        if not title1 or not title2:
            return False

        title_match = title1.lower() == title2.lower()
        if not title_match:
            return False

        # Comparar datas
        # Extrair data/hora de acordo com o tipo de evento
        try:
            if source1 == "google":
                start1_str = evento1.get("start", {}).get("dateTime", "")
                start1 = (
                    datetime.fromisoformat(start1_str.replace("Z", "+00:00"))
                    if start1_str
                    else None
                )
            elif source1 == "outlook":
                start1_str = evento1.get("start", {}).get("dateTime", "")
                start1 = (
                    datetime.fromisoformat(start1_str.replace("Z", "+00:00"))
                    if start1_str
                    else None
                )
            elif source1 == "expresso":
                data1 = evento1.get("data", "")
                inicio1 = evento1.get("inicio", "")
                if data1 and inicio1 and ":" in inicio1:
                    dia, mes, ano = data1.split("/")
                    hora, minuto = inicio1.split(":")
                    start1 = datetime(
                        int(ano), int(mes), int(dia), int(hora), int(minuto)
                    )
                else:
                    start1 = None

            if source2 == "google":
                start2_str = evento2.get("start", {}).get("dateTime", "")
                start2 = (
                    datetime.fromisoformat(start2_str.replace("Z", "+00:00"))
                    if start2_str
                    else None
                )
            elif source2 == "outlook":
                start2_str = evento2.get("start", {}).get("dateTime", "")
                start2 = (
                    datetime.fromisoformat(start2_str.replace("Z", "+00:00"))
                    if start2_str
                    else None
                )
            elif source2 == "expresso":
                data2 = evento2.get("data", "")
                inicio2 = evento2.get("inicio", "")
                if data2 and inicio2 and ":" in inicio2:
                    dia, mes, ano = data2.split("/")
                    hora, minuto = inicio2.split(":")
                    start2 = datetime(
                        int(ano), int(mes), int(dia), int(hora), int(minuto)
                    )
                else:
                    start2 = None

            # Se não foi possível extrair datas, não considerar match
            if not start1 or not start2:
                return title_match  # Pelo menos os títulos correspondem

            # Verificar diferença de tempo (tolerância de 5 minutos)
            time_diff = abs((start1 - start2).total_seconds())
            if time_diff > 300:  # 5 minutos
                return False

        except (ValueError, TypeError, IndexError, AttributeError) as e:
            # Se houver erro na comparação de datas, considerar apenas o título
            print(f"Erro ao comparar datas: {e}")
            return title_match

        # Ambos título e data/hora correspondem
        return True

    def _format_google_to_expresso(self, google_event):
        """Converte um evento do Google para o formato do Expresso"""
        if not google_event.get("summary") or not google_event.get("start"):
            return None
        
        # Extrair data e hora do evento
        start_datetime = None
        end_datetime = None
        dia_inteiro = False
        
        # Verificar se é um evento de dia inteiro
        if "date" in google_event.get("start", {}):
            # Evento de dia inteiro
            start_date = google_event["start"]["date"]
            end_date = google_event["end"]["date"] if "date" in google_event.get("end", {}) else start_date
            
            # Converter para formato DD/MM/YYYY
            start_dt = datetime.fromisoformat(start_date)
            data_formatada = start_dt.strftime("%d/%m/%Y")
            
            dia_inteiro = True
            inicio = "00:00"
            fim = "23:59"
        else:
            # Evento com horário específico
            start_datetime_str = google_event["start"]["dateTime"]
            end_datetime_str = google_event["end"]["dateTime"]
            
            # Converter para objetos datetime
            start_dt = datetime.fromisoformat(start_datetime_str.replace("Z", "+00:00"))
            end_dt = datetime.fromisoformat(end_datetime_str.replace("Z", "+00:00"))
            
            # Formatar data e hora
            data_formatada = start_dt.strftime("%d/%m/%Y")
            inicio = start_dt.strftime("%H:%M")
            fim = end_dt.strftime("%H:%M")
            dia_inteiro = False
        
        # Criar o evento no formato do Expresso
        expresso_event = {
            "titulo": google_event.get("summary", "Sem título"),
            "descricao": google_event.get("description", "Evento sincronizado do Google Calendar"),
            "data": data_formatada,
            "inicio": inicio,
            "fim": fim,
            "dia_inteiro": dia_inteiro,
            "localizacao": google_event.get("location", "")
        }
        
        return expresso_event
    
    def _format_outlook_to_expresso(self, outlook_event):
        """Converte um evento do Outlook para o formato do Expresso"""
        if not outlook_event.get("subject") or not outlook_event.get("start"):
            return None
        
        # Extrair data e hora do evento
        start_datetime = None
        end_datetime = None
        dia_inteiro = False
        
        # Verificar se é um evento de dia inteiro
        if "dateTime" not in outlook_event.get("start", {}):
            # Evento de dia inteiro
            dia_inteiro = True
            inicio = "00:00"
            fim = "23:59"
            
            # Usar a data como está
            data_formatada = outlook_event.get("start", {}).get("date", "")
            
            # Converter para formato DD/MM/YYYY se necessário
            if data_formatada and len(data_formatada) == 10:  # formato YYYY-MM-DD
                dt = datetime.fromisoformat(data_formatada)
                data_formatada = dt.strftime("%d/%m/%Y")
        else:
            # Evento com horário específico
            start_datetime_str = outlook_event["start"]["dateTime"]
            end_datetime_str = outlook_event["end"]["dateTime"]
            
            # Converter para objetos datetime
            start_dt = datetime.fromisoformat(start_datetime_str.replace("Z", "+00:00"))
            end_dt = datetime.fromisoformat(end_datetime_str.replace("Z", "+00:00"))
            
            # Formatar data e hora
            data_formatada = start_dt.strftime("%d/%m/%Y")
            inicio = start_dt.strftime("%H:%M")
            fim = end_dt.strftime("%H:%M")
            dia_inteiro = False
        
        # Criar o evento no formato do Expresso
        expresso_event = {
            "titulo": outlook_event.get("subject", "Sem título"),
            "descricao": outlook_event.get("body", {}).get("content", "Evento sincronizado do Outlook Calendar"),
            "data": data_formatada,
            "inicio": inicio,
            "fim": fim,
            "dia_inteiro": dia_inteiro,
            "localizacao": outlook_event.get("location", {}).get("displayName", "")
        }
        
        return expresso_event

    def _normalize_event_for_comparison(self, event, source_type):
        """Normaliza um evento para facilitar a comparação entre diferentes fontes"""
        normalized = {}
        
        # Extrair título normalizado
        if source_type == 'google':
            normalized['title'] = event.get('summary', '').strip().lower()
        elif source_type == 'outlook':
            normalized['title'] = event.get('subject', '').strip().lower()
        elif source_type == 'expresso':
            normalized['title'] = event.get('titulo', '').strip().lower()
        
        # Extrair e normalizar data/hora
        if source_type == 'google':
            if 'start' in event:
                if 'dateTime' in event['start']:
                    dt_str = event['start']['dateTime']
                    try:
                        dt = datetime.fromisoformat(dt_str.replace('Z', '+00:00'))
                        normalized['date'] = dt.date().isoformat()
                        normalized['time'] = dt.time().strftime('%H:%M')
                    except:
                        pass
        elif source_type == 'outlook':
            if 'start' in event:
                if 'dateTime' in event['start']:
                    dt_str = event['start']['dateTime']
                    try:
                        dt = datetime.fromisoformat(dt_str.replace('Z', '+00:00'))
                        normalized['date'] = dt.date().isoformat()
                        normalized['time'] = dt.time().strftime('%H:%M')
                    except:
                        pass
        elif source_type == 'expresso':
            data = event.get('data', '')
            hora = event.get('inicio', '')
            
            if data and '/' in data:
                try:
                    dia, mes, ano = data.split('/')
                    normalized['date'] = f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"
                except:
                    pass
                
            if hora and ':' in hora:
                normalized['time'] = hora
        
        return normalized

    def _validate_event_data(self, event_data, source_type):
        required_fields = {
            'google': ['summary', 'start'],
            'outlook': ['subject', 'start'],
            'expresso': ['titulo', 'data', 'inicio']
        }
        
        for field in required_fields[source_type]:
            if source_type == 'google' or source_type == 'outlook':
                if field not in event_data:
                    return False, f"Campo obrigatório ausente: {field}"
                if field == 'start' and 'dateTime' not in event_data[field]:
                    return False, f"Campo obrigatório ausente: {field}.dateTime"
            else:
                if field not in event_data:
                    return False, f"Campo obrigatório ausente: {field}"
        
        return True, ""

    def _validate_and_normalize_attendees(self, attendees, target_type):
        if not attendees:
            return []
        
        validated = []
        for attendee in attendees:
            email = None
            if isinstance(attendee, dict) and 'email' in attendee:
                email = attendee['email']
            elif isinstance(attendee, str):
                email = attendee.strip()
            
            # Verificar se o e-mail é válido (formato simples)
            if email and '@' in email and '.' in email:
                if target_type == 'google':
                    validated.append({'email': email})
                elif target_type == 'outlook':
                    validated.append({
                        'emailAddress': {'address': email},
                        'type': 'required'
                    })
        
        return validated

    def _format_expresso_to_google(self, expresso_event):
        """Converte um evento do Expresso para o formato do Google Calendar"""
        if not expresso_event.get("titulo") or not expresso_event.get("data"):
            return None
        
        # Extrair data e criar objeto datetime
        data_str = expresso_event.get("data", "")
        inicio_str = expresso_event.get("inicio", "")
        fim_str = expresso_event.get("fim", "")
        
        # Extrair componentes da data (formato DD/MM/YYYY)
        try:
            if "/" in data_str:
                dia, mes, ano = data_str.split("/")
                data_iso = f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"
            else:
                data_iso = data_str
            
            # Criar objeto do evento para o Google
            google_event = {
                "summary": expresso_event.get("titulo", "Sem título"),
                "description": expresso_event.get("descricao", "Evento sincronizado do Expresso"),
                "location": expresso_event.get("localizacao", ""),
            }
            
            # Verificar se é evento de dia inteiro
            if expresso_event.get("dia_inteiro", False):
                google_event["start"] = {"date": data_iso}
                # Para eventos de dia inteiro, a data de término é o dia seguinte
                from datetime import datetime, timedelta
                end_date = (datetime.fromisoformat(data_iso) + timedelta(days=1)).date().isoformat()
                google_event["end"] = {"date": end_date}
            else:
                # Evento com horário específico
                if inicio_str and ":" in inicio_str:
                    hora, minuto = inicio_str.split(":")
                    start_datetime = f"{data_iso}T{hora.zfill(2)}:{minuto.zfill(2)}:00"
                    google_event["start"] = {
                        "dateTime": start_datetime,
                        "timeZone": "America/Recife"
                    }
                
                if fim_str and ":" in fim_str:
                    hora, minuto = fim_str.split(":")
                    end_datetime = f"{data_iso}T{hora.zfill(2)}:{minuto.zfill(2)}:00"
                    google_event["end"] = {
                        "dateTime": end_datetime,
                        "timeZone": "America/Recife"
                    }
            
            # Tratar participantes - este é o ponto crítico do erro
            participantes = expresso_event.get("participantes", "")
            if participantes:
                # Validar e normalizar os e-mails dos participantes
                attendees = self._validate_and_normalize_attendees(participantes.split(','), 'google')
                if attendees:  # Só incluir se houver participantes válidos
                    google_event["attendees"] = attendees
            
            return google_event
        except Exception as e:
            print(f"Erro ao converter evento do Expresso para Google: {e}")
            return None

    def _remove_all_mappings(self, google_id=None, outlook_id=None, expresso_id=None):
        """Remove todos os mapeamentos relacionados ao evento"""
        # Obter todos os IDs mapeados
        if google_id:
            mapped_ids = self.db.get_mapped_ids(google_id, "google")
            if mapped_ids:
                outlook_id = mapped_ids[0] or outlook_id
                expresso_id = mapped_ids[1] or expresso_id
        elif outlook_id:
            mapped_ids = self.db.get_mapped_ids(outlook_id, "outlook")
            if mapped_ids:
                google_id = mapped_ids[0] or google_id
                expresso_id = mapped_ids[1] or expresso_id
        elif expresso_id:
            mapped_ids = self.db.get_mapped_ids(expresso_id, "expresso")
            if mapped_ids:
                outlook_id = mapped_ids[0] or outlook_id
                google_id = mapped_ids[1] or google_id
        
        # Remover do banco de dados
        # Criar um método no DatabaseManager para remover mapeamentos
        self.db.remove_mapping(google_id=google_id, outlook_id=outlook_id, expresso_id=expresso_id)
        
        # Remover dos mapas em memória
        if google_id and outlook_id:
            if google_id in self.google_to_outlook_map:
                del self.google_to_outlook_map[google_id]
            if outlook_id in self.outlook_to_google_map:
                del self.outlook_to_google_map[outlook_id]
        
        if google_id and expresso_id:
            if google_id in self.google_to_expresso_map:
                del self.google_to_expresso_map[google_id]
            if expresso_id in self.expresso_to_google_map:
                del self.expresso_to_google_map[expresso_id]
            
        if outlook_id and expresso_id:
            if outlook_id in self.outlook_to_expresso_map:
                del self.outlook_to_expresso_map[outlook_id]
            if expresso_id in self.expresso_to_outlook_map:
                del self.expresso_to_outlook_map[expresso_id]
