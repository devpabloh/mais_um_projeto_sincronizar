import sqlite3
from datetime import datetime, timedelta
import os


class DatabaseManager:
    def __init__(self, db_file="calendar_sync.db"):
        self.db_file = db_file
        self.conn = None
        self.setup_database()

    def setup_database(self):
        """Configura o banco de dados e cria tabelas se não existirem"""
        # Verificar se o banco de dados já existe
        db_exists = os.path.exists(self.db_file)

        # Conectar ao banco de dados (será criado se não existir)
        self.conn = sqlite3.connect(self.db_file)

        # Habilitar chaves estrangeiras
        self.conn.execute("PRAGMA foreign_keys = ON")

        cursor = self.conn.cursor()

        # Criar tabelas
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS outlook_events (
            id VARCHAR(255) PRIMARY KEY,
            subject VARCHAR(255),
            start_datetime DATETIME,
            end_datetime DATETIME,
            location TEXT,
            description TEXT,
            is_all_day BOOLEAN,
            last_modified DATETIME,
            status VARCHAR(50),
            created_at DATETIME,
            updated_at DATETIME
        )
        """
        )

        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS google_events (
            id VARCHAR(255) PRIMARY KEY,
            summary VARCHAR(255),
            start_datetime DATETIME,
            end_datetime DATETIME,
            location TEXT,
            description TEXT,
            is_all_day BOOLEAN,
            last_modified DATETIME,
            status VARCHAR(50),
            created_at DATETIME,
            updated_at DATETIME
        )
        """
        )

        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS expresso_events (
            id VARCHAR(255) PRIMARY KEY,
            titulo VARCHAR(255),
            data_inicio DATETIME,
            data_fim DATETIME,
            local TEXT,
            descricao TEXT,
            is_all_day BOOLEAN,
            participantes TEXT,
            last_modified DATETIME,
            status VARCHAR(50),
            created_at DATETIME,
            updated_at DATETIME
        )
        """
        )

        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS eventos_sincronizados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            outlook_event_id VARCHAR(255),
            google_event_id VARCHAR(255),
            expresso_event_id VARCHAR(255),
            ultima_sincronizacao DATETIME,
            origem_criacao VARCHAR(20),
            FOREIGN KEY (outlook_event_id) REFERENCES outlook_events(id),
            FOREIGN KEY (google_event_id) REFERENCES google_events(id),
            FOREIGN KEY (expresso_event_id) REFERENCES expresso_events(id)
        )
        """
        )

        # Criar índices para otimização
        cursor.execute(
            "CREATE INDEX IF NOT EXISTS idx_outlook_id ON eventos_sincronizados(outlook_event_id)"
        )
        cursor.execute(
            "CREATE INDEX IF NOT EXISTS idx_google_id ON eventos_sincronizados(google_event_id)"
        )
        cursor.execute(
            "CREATE INDEX IF NOT EXISTS idx_expresso_id ON eventos_sincronizados(expresso_event_id)"
        )

        # Salvar as alterações
        self.conn.commit()

        if not db_exists:
            print(f"Banco de dados '{self.db_file}' criado com sucesso.")

    # Métodos para manipular eventos do Outlook
    def store_outlook_event(self, event):
        """Armazena ou atualiza um evento do Outlook"""
        cursor = self.conn.cursor()
        now = datetime.now().isoformat()

        # Verificar se o evento já existe
        cursor.execute("SELECT id FROM outlook_events WHERE id = ?", (event["id"],))
        exists = cursor.fetchone()

        if exists:
            # Atualizar evento existente
            cursor.execute(
                """
                UPDATE outlook_events SET
                subject = ?, start_datetime = ?, end_datetime = ?,
                location = ?, description = ?, is_all_day = ?,
                last_modified = ?, status = ?, updated_at = ?
                WHERE id = ?
            """,
                (
                    event.get("subject", ""),
                    event.get("start", {}).get("dateTime", ""),
                    event.get("end", {}).get("dateTime", ""),
                    event.get("location", {}).get("displayName", ""),
                    event.get("body", {}).get("content", ""),
                    event.get("isAllDay", False),
                    event.get("lastModifiedDateTime", now),
                    "ativo",
                    now,
                    event["id"],
                ),
            )
        else:
            # Inserir novo evento
            cursor.execute(
                """
                INSERT INTO outlook_events (
                    id, subject, start_datetime, end_datetime,
                    location, description, is_all_day,
                    last_modified, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    event["id"],
                    event.get("subject", ""),
                    event.get("start", {}).get("dateTime", ""),
                    event.get("end", {}).get("dateTime", ""),
                    event.get("location", {}).get("displayName", ""),
                    event.get("body", {}).get("content", ""),
                    event.get("isAllDay", False),
                    event.get("lastModifiedDateTime", now),
                    "ativo",
                    now,
                    now,
                ),
            )

        self.conn.commit()
        return event["id"]

    # Métodos para manipular eventos do Google
    def store_google_event(self, event):
        """Armazena ou atualiza um evento do Google"""
        cursor = self.conn.cursor()
        now = datetime.now().isoformat()

        # Verificar se o evento já existe
        cursor.execute("SELECT id FROM google_events WHERE id = ?", (event["id"],))
        exists = cursor.fetchone()

        if exists:
            # Atualizar evento existente
            cursor.execute(
                """
                UPDATE google_events SET
                summary = ?, start_datetime = ?, end_datetime = ?,
                location = ?, description = ?, is_all_day = ?,
                last_modified = ?, status = ?, updated_at = ?
                WHERE id = ?
            """,
                (
                    event.get("summary", ""),
                    event.get("start", {}).get("dateTime", ""),
                    event.get("end", {}).get("dateTime", ""),
                    event.get("location", ""),
                    event.get("description", ""),
                    "dateTime"
                    not in event.get("start", {}),  # se não tem dateTime, é all day
                    event.get("updated", now),
                    "ativo",
                    now,
                    event["id"],
                ),
            )
        else:
            # Inserir novo evento
            cursor.execute(
                """
                INSERT INTO google_events (
                    id, summary, start_datetime, end_datetime,
                    location, description, is_all_day,
                    last_modified, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    event["id"],
                    event.get("summary", ""),
                    event.get("start", {}).get("dateTime", ""),
                    event.get("end", {}).get("dateTime", ""),
                    event.get("location", ""),
                    event.get("description", ""),
                    "dateTime" not in event.get("start", {}),
                    event.get("updated", now),
                    "ativo",
                    now,
                    now,
                ),
            )

        self.conn.commit()
        return event["id"]

    # Métodos para manipular eventos do Expresso
    def store_expresso_event(self, event):
        """Armazena ou atualiza um evento do Expresso"""
        cursor = self.conn.cursor()
        now = datetime.now().isoformat()

        # Verificar se o evento já existe
        cursor.execute("SELECT id FROM expresso_events WHERE id = ?", (event["id"],))
        exists = cursor.fetchone()

        if exists:
            # Atualizar evento existente
            cursor.execute(
                """
                UPDATE expresso_events SET
                titulo = ?, data_inicio = ?, data_fim = ?,
                local = ?, descricao = ?, is_all_day = ?,
                participantes = ?, last_modified = ?, status = ?, updated_at = ?
                WHERE id = ?
            """,
                (
                    event.get("titulo", ""),
                    f"{event.get('data', '')} {event.get('inicio', '')}",
                    f"{event.get('data', '')} {event.get('fim', '')}",
                    "",  # local não está disponível no objeto de evento Expresso
                    event.get("descricao", ""),
                    False,  # assumindo que eventos no Expresso não são de dia inteiro
                    event.get("participantes", ""),
                    now,  # não temos data de modificação no Expresso
                    "ativo",
                    now,
                    event["id"],
                ),
            )
        else:
            # Inserir novo evento
            cursor.execute(
                """
                INSERT INTO expresso_events (
                    id, titulo, data_inicio, data_fim,
                    local, descricao, is_all_day, participantes,
                    last_modified, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    event["id"],
                    event.get("titulo", ""),
                    f"{event.get('data', '')} {event.get('inicio', '')}",
                    f"{event.get('data', '')} {event.get('fim', '')}",
                    "",  # local não está disponível no objeto de evento Expresso
                    event.get("descricao", ""),
                    False,  # assumindo que eventos no Expresso não são de dia inteiro
                    event.get("participantes", ""),
                    now,  # não temos data de modificação no Expresso
                    "ativo",
                    now,
                    now,
                ),
            )

        self.conn.commit()
        return event["id"]

    # Métodos para gerenciar mapeamentos entre eventos
    def map_events(
        self, outlook_id=None, google_id=None, expresso_id=None, origem=None
    ):
        """Criar ou atualizar mapeamento entre eventos"""
        cursor = self.conn.cursor()
        now = datetime.now().isoformat()
        mapping_id = None

        # Procurar mapeamento existente
        if outlook_id:
            cursor.execute(
                "SELECT id FROM eventos_sincronizados WHERE outlook_event_id = ?",
                (outlook_id,),
            )
            mapping = cursor.fetchone()
            if mapping:
                mapping_id = mapping[0]

        if google_id and not mapping_id:
            cursor.execute(
                "SELECT id FROM eventos_sincronizados WHERE google_event_id = ?",
                (google_id,),
            )
            mapping = cursor.fetchone()
            if mapping:
                mapping_id = mapping[0]

        if expresso_id and not mapping_id:
            cursor.execute(
                "SELECT id FROM eventos_sincronizados WHERE expresso_event_id = ?",
                (expresso_id,),
            )
            mapping = cursor.fetchone()
            if mapping:
                mapping_id = mapping[0]

        if mapping_id:
            # Atualizar mapeamento existente
            cursor.execute(
                """
                UPDATE eventos_sincronizados SET
                outlook_event_id = COALESCE(?, outlook_event_id),
                google_event_id = COALESCE(?, google_event_id),
                expresso_event_id = COALESCE(?, expresso_event_id),
                ultima_sincronizacao = ?
                WHERE id = ?
            """,
                (outlook_id, google_id, expresso_id, now, mapping_id),
            )
        else:
            # Criar novo mapeamento
            cursor.execute(
                """
                INSERT INTO eventos_sincronizados (
                    outlook_event_id, google_event_id, expresso_event_id,
                    ultima_sincronizacao, origem_criacao
                ) VALUES (?, ?, ?, ?, ?)
            """,
                (outlook_id, google_id, expresso_id, now, origem),
            )

        self.conn.commit()

    def get_mapped_ids(self, event_id, source):
        """Retorna IDs mapeados de um evento"""
        cursor = self.conn.cursor()

        if source == "outlook":
            cursor.execute(
                """
                SELECT google_event_id, expresso_event_id FROM eventos_sincronizados
                WHERE outlook_event_id = ?
            """,
                (event_id,),
            )
        elif source == "google":
            cursor.execute(
                """
                SELECT outlook_event_id, expresso_event_id FROM eventos_sincronizados
                WHERE google_event_id = ?
            """,
                (event_id,),
            )
        elif source == "expresso":
            cursor.execute(
                """
                SELECT outlook_event_id, google_event_id FROM eventos_sincronizados
                WHERE expresso_event_id = ?
            """,
                (event_id,),
            )
        else:
            return None

        return cursor.fetchone()

    def mark_event_deleted(self, event_id, source):
        """Marca um evento como excluído"""
        cursor = self.conn.cursor()
        now = datetime.now().isoformat()

        if source == "outlook":
            cursor.execute(
                """
                UPDATE outlook_events SET status = 'excluído', updated_at = ?
                WHERE id = ?
            """,
                (now, event_id),
            )
        elif source == "google":
            cursor.execute(
                """
                UPDATE google_events SET status = 'excluído', updated_at = ?
                WHERE id = ?
            """,
                (now, event_id),
            )
        elif source == "expresso":
            cursor.execute(
                """
                UPDATE expresso_events SET status = 'excluído', updated_at = ?
                WHERE id = ?
            """,
                (now, event_id),
            )

        self.conn.commit()

    def get_all_mappings(self):
        """Retorna todos os mapeamentos de eventos"""
        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT id, outlook_event_id, google_event_id, expresso_event_id, 
                   ultima_sincronizacao, origem_criacao
            FROM eventos_sincronizados
        """
        )
        return cursor.fetchall()

    def close(self):
        """Fecha a conexão com o banco de dados"""
        if self.conn:
            self.conn.close()

    def cleanup_old_events(self, days_to_keep=0):
        """
        Remove eventos antigos do banco de dados.

        Args:
            days_to_keep (int): Mantém eventos a partir de hoje até X dias no passado.
                                0 = mantém apenas eventos de hoje em diante.
        """
        cursor = self.conn.cursor()
        now = datetime.now()
        cutoff_date = (
            (now - timedelta(days=days_to_keep))
            .replace(hour=0, minute=0, second=0, microsecond=0)
            .isoformat()
        )

        print(f"Removendo eventos anteriores a {cutoff_date.split('T')[0]}")

        # Remover mapeamentos para eventos do Google que serão excluídos
        cursor.execute(
            """
            DELETE FROM eventos_sincronizados 
            WHERE google_event_id IN (
                SELECT id FROM google_events 
                WHERE end_datetime < ? AND end_datetime != ''
            )
        """,
            (cutoff_date,),
        )

        # Remover mapeamentos para eventos do Outlook que serão excluídos
        cursor.execute(
            """
            DELETE FROM eventos_sincronizados 
            WHERE outlook_event_id IN (
                SELECT id FROM outlook_events 
                WHERE end_datetime < ? AND end_datetime != ''
            )
        """,
            (cutoff_date,),
        )

        # Remover mapeamentos para eventos do Expresso que serão excluídos
        cursor.execute(
            """
            DELETE FROM eventos_sincronizados 
            WHERE expresso_event_id IN (
                SELECT id FROM expresso_events 
                WHERE data_fim < ? AND data_fim != ''
            )
        """,
            (cutoff_date,),
        )

        # Excluir eventos antigos de cada tabela
        cursor.execute(
            "DELETE FROM google_events WHERE end_datetime < ? AND end_datetime != ''",
            (cutoff_date,),
        )
        deleted_google = cursor.rowcount

        cursor.execute(
            "DELETE FROM outlook_events WHERE end_datetime < ? AND end_datetime != ''",
            (cutoff_date,),
        )
        deleted_outlook = cursor.rowcount

        cursor.execute(
            "DELETE FROM expresso_events WHERE data_fim < ? AND data_fim != ''",
            (cutoff_date,),
        )
        deleted_expresso = cursor.rowcount

        # Remover mapeamentos órfãos (que não têm mais eventos associados)
        cursor.execute(
            """
            DELETE FROM eventos_sincronizados
            WHERE (outlook_event_id IS NULL OR 
                   outlook_event_id NOT IN (SELECT id FROM outlook_events))
              AND (google_event_id IS NULL OR 
                   google_event_id NOT IN (SELECT id FROM google_events))
              AND (expresso_event_id IS NULL OR 
                   expresso_event_id NOT IN (SELECT id FROM expresso_events))
        """
        )
        deleted_mappings = cursor.rowcount

        self.conn.commit()

        print(f"Limpeza concluída. Removidos:")
        print(f"- Eventos do Google: {deleted_google}")
        print(f"- Eventos do Outlook: {deleted_outlook}")
        print(f"- Eventos do Expresso: {deleted_expresso}")
        print(f"- Mapeamentos órfãos: {deleted_mappings}")

        return {
            "google": deleted_google,
            "outlook": deleted_outlook,
            "expresso": deleted_expresso,
            "mappings": deleted_mappings,
        }
