"""Local, workbook-scoped conversation memory for the finance assistant."""

from __future__ import annotations

import json
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class AssistantMemory:
    """SQLite chat archive. Each conversation belongs to one selected workbook."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self):
        connection = sqlite3.connect(self.path, timeout=10)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA busy_timeout = 10000")
        return connection

    @contextmanager
    def _connection(self):
        connection = self._connect()
        try:
            with connection:
                yield connection
        finally:
            connection.close()

    def _initialize(self):
        with self._connection() as connection:
            connection.execute("PRAGMA journal_mode = WAL")
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS conversations (
                    id TEXT PRIMARY KEY,
                    workbook_path TEXT NOT NULL,
                    title TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_conversations_workbook_updated
                    ON conversations(workbook_path, updated_at DESC);
                CREATE TABLE IF NOT EXISTS messages (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    conversation_id TEXT NOT NULL REFERENCES conversations(id) ON DELETE CASCADE,
                    role TEXT NOT NULL CHECK(role IN ('user', 'assistant')),
                    content TEXT NOT NULL,
                    metadata_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_messages_conversation_id
                    ON messages(conversation_id, id);
                CREATE TABLE IF NOT EXISTS memory_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    workbook_path TEXT NOT NULL,
                    memory_key TEXT NOT NULL,
                    memory_value TEXT NOT NULL,
                    source_conversation_id TEXT REFERENCES conversations(id) ON DELETE SET NULL,
                    source_message_id INTEGER REFERENCES messages(id) ON DELETE SET NULL,
                    confidence REAL NOT NULL DEFAULT 1.0,
                    approved INTEGER NOT NULL DEFAULT 0 CHECK(approved IN (0, 1)),
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(workbook_path, memory_key)
                );
                """
            )

    @staticmethod
    def _workbook_key(workbook_path):
        return str(Path(workbook_path).resolve()).casefold()

    def create_conversation(self, workbook_path, title="Nueva conversación"):
        conversation_id = uuid.uuid4().hex
        now = _now()
        with self._connection() as connection:
            connection.execute(
                "INSERT INTO conversations(id, workbook_path, title, created_at, updated_at) VALUES(?, ?, ?, ?, ?)",
                (conversation_id, self._workbook_key(workbook_path), title[:120], now, now),
            )
        return conversation_id

    def list_conversations(self, workbook_path, limit=50):
        with self._connection() as connection:
            rows = connection.execute(
                """SELECT c.id, c.title, c.created_at, c.updated_at, COUNT(m.id) AS message_count
                   FROM conversations c LEFT JOIN messages m ON m.conversation_id = c.id
                   WHERE c.workbook_path = ?
                   GROUP BY c.id ORDER BY c.updated_at DESC LIMIT ?""",
                (self._workbook_key(workbook_path), max(1, min(int(limit), 100))),
            ).fetchall()
        return [dict(row) for row in rows]

    def get_conversation(self, conversation_id, workbook_path):
        with self._connection() as connection:
            row = connection.execute(
                "SELECT id, title, created_at, updated_at FROM conversations WHERE id = ? AND workbook_path = ?",
                (str(conversation_id), self._workbook_key(workbook_path)),
            ).fetchone()
        return dict(row) if row else None

    def get_messages(self, conversation_id, workbook_path, limit=100):
        if not self.get_conversation(conversation_id, workbook_path):
            raise ValueError("No encontré esa conversación para el libro seleccionado.")
        with self._connection() as connection:
            rows = connection.execute(
                """SELECT id, role, content, metadata_json, created_at FROM (
                       SELECT id, role, content, metadata_json, created_at FROM messages
                       WHERE conversation_id = ? ORDER BY id DESC LIMIT ?
                   ) ORDER BY id ASC""",
                (str(conversation_id), max(1, min(int(limit), 200))),
            ).fetchall()
        messages = []
        for row in rows:
            item = dict(row)
            try:
                item["metadata"] = json.loads(item.pop("metadata_json") or "{}")
            except json.JSONDecodeError:
                item["metadata"] = {}
                item.pop("metadata_json", None)
            messages.append(item)
        return messages

    def append_message(self, conversation_id, workbook_path, role, content, metadata=None):
        if role not in {"user", "assistant"}:
            raise ValueError("El tipo de mensaje no es válido.")
        content = str(content or "")[:12000]
        now = _now()
        workbook_key = self._workbook_key(workbook_path)
        with self._connection() as connection:
            conversation = connection.execute(
                "SELECT title FROM conversations WHERE id = ? AND workbook_path = ?",
                (str(conversation_id), workbook_key),
            ).fetchone()
            if not conversation:
                raise ValueError("No encontré esa conversación para el libro seleccionado.")
            cursor = connection.execute(
                "INSERT INTO messages(conversation_id, role, content, metadata_json, created_at) VALUES(?, ?, ?, ?, ?)",
                (str(conversation_id), role, content, json.dumps(metadata or {}, ensure_ascii=False, default=str), now),
            )
            if role == "user" and conversation["title"] == "Nueva conversación":
                title = " ".join(content.split())[:72]
                connection.execute(
                    "UPDATE conversations SET title = ?, updated_at = ? WHERE id = ?",
                    (title or "Nueva conversación", now, str(conversation_id)),
                )
            else:
                connection.execute(
                    "UPDATE conversations SET updated_at = ? WHERE id = ?",
                    (now, str(conversation_id)),
                )
            return cursor.lastrowid

    def update_proposal_status(self, conversation_id, workbook_path, plan_id, status):
        if status not in {"applied", "cancelled", "expired"}:
            raise ValueError("El estado de propuesta no es válido.")
        conversation = self.get_conversation(conversation_id, workbook_path)
        if not conversation:
            return False
        with self._connection() as connection:
            rows = connection.execute(
                "SELECT id, metadata_json FROM messages WHERE conversation_id = ? AND role = 'assistant' ORDER BY id DESC",
                (str(conversation_id),),
            ).fetchall()
            for row in rows:
                try:
                    metadata = json.loads(row["metadata_json"] or "{}")
                except json.JSONDecodeError:
                    continue
                if metadata.get("plan_id") == plan_id:
                    metadata["status"] = status
                    connection.execute(
                        "UPDATE messages SET metadata_json = ? WHERE id = ?",
                        (json.dumps(metadata, ensure_ascii=False), row["id"]),
                    )
                    return True
        return False

    def delete_conversation(self, conversation_id, workbook_path):
        with self._connection() as connection:
            cursor = connection.execute(
                "DELETE FROM conversations WHERE id = ? AND workbook_path = ?",
                (str(conversation_id), self._workbook_key(workbook_path)),
            )
            return cursor.rowcount > 0

    def list_memories(self, workbook_path):
        with self._connection() as connection:
            rows = connection.execute(
                """SELECT id, memory_key, memory_value, confidence, created_at, updated_at
                   FROM memory_items WHERE workbook_path = ? AND approved = 1
                   ORDER BY updated_at DESC, memory_key COLLATE NOCASE""",
                (self._workbook_key(workbook_path),),
            ).fetchall()
        return [dict(row) for row in rows]

    def save_memory(self, workbook_path, key, value, conversation_id="", source_message_id=None):
        key = str(key or "").strip().lower()
        value = " ".join(str(value or "").split())
        if not key or len(key) > 60 or not all(char.isalnum() or char in "_-" for char in key):
            raise ValueError("El nombre de esta memoria no es válido.")
        if not value or len(value) > 600:
            raise ValueError("El contenido de la memoria debe tener entre 1 y 600 caracteres.")
        workbook_key = self._workbook_key(workbook_path)
        now = _now()
        with self._connection() as connection:
            if conversation_id and not connection.execute(
                "SELECT 1 FROM conversations WHERE id = ? AND workbook_path = ?",
                (str(conversation_id), workbook_key),
            ).fetchone():
                raise ValueError("No encontré el chat de origen de esta memoria.")
            count = connection.execute(
                "SELECT COUNT(*) FROM memory_items WHERE workbook_path = ? AND approved = 1",
                (workbook_key,),
            ).fetchone()[0]
            existing = connection.execute(
                "SELECT id FROM memory_items WHERE workbook_path = ? AND memory_key = ?",
                (workbook_key, key),
            ).fetchone()
            if count >= 100 and not existing:
                raise ValueError("Este libro ya tiene 100 memorias guardadas. Elimina alguna antes de agregar otra.")
            connection.execute(
                """INSERT INTO memory_items(workbook_path, memory_key, memory_value, source_conversation_id,
                       source_message_id, confidence, approved, created_at, updated_at)
                   VALUES(?, ?, ?, ?, ?, 1.0, 1, ?, ?)
                   ON CONFLICT(workbook_path, memory_key) DO UPDATE SET
                       memory_value = excluded.memory_value,
                       source_conversation_id = excluded.source_conversation_id,
                       source_message_id = excluded.source_message_id,
                       confidence = excluded.confidence, approved = 1, updated_at = excluded.updated_at""",
                (workbook_key, key, value, conversation_id or None, source_message_id, now, now),
            )

    def save_memories(self, workbook_path, updates, conversation_id="", source_message_id=None):
        prepared = []
        for item in updates:
            key = str(item.get("key", "")).strip().lower()
            value = " ".join(str(item.get("value", "") or "").split())
            if not key or len(key) > 60 or not all(char.isalnum() or char in "_-" for char in key):
                raise ValueError("El nombre de esta memoria no es válido.")
            if not value or len(value) > 600:
                raise ValueError("El contenido de la memoria debe tener entre 1 y 600 caracteres.")
            prepared.append((key, value))
        prepared = list(dict(prepared).items())
        if not prepared:
            return
        workbook_key = self._workbook_key(workbook_path)
        now = _now()
        with self._connection() as connection:
            if conversation_id and not connection.execute(
                "SELECT 1 FROM conversations WHERE id = ? AND workbook_path = ?",
                (str(conversation_id), workbook_key),
            ).fetchone():
                raise ValueError("No encontré el chat de origen de esta memoria.")
            existing_keys = {row[0] for row in connection.execute(
                "SELECT memory_key FROM memory_items WHERE workbook_path = ? AND approved = 1",
                (workbook_key,),
            ).fetchall()}
            if len(existing_keys | {key for key, _ in prepared}) > 100:
                raise ValueError("Este libro ya tiene 100 memorias guardadas. Elimina alguna antes de agregar otra.")
            for key, value in prepared:
                connection.execute(
                    """INSERT INTO memory_items(workbook_path, memory_key, memory_value, source_conversation_id,
                           source_message_id, confidence, approved, created_at, updated_at)
                       VALUES(?, ?, ?, ?, ?, 1.0, 1, ?, ?)
                       ON CONFLICT(workbook_path, memory_key) DO UPDATE SET
                           memory_value = excluded.memory_value,
                           source_conversation_id = excluded.source_conversation_id,
                           source_message_id = excluded.source_message_id,
                           confidence = excluded.confidence, approved = 1, updated_at = excluded.updated_at""",
                    (workbook_key, key, value, conversation_id or None, source_message_id, now, now),
                )

    def delete_memory(self, workbook_path, memory_id):
        with self._connection() as connection:
            cursor = connection.execute(
                "DELETE FROM memory_items WHERE id = ? AND workbook_path = ?",
                (int(memory_id), self._workbook_key(workbook_path)),
            )
            return cursor.rowcount > 0

    def ai_history(self, conversation_id, workbook_path, limit=12):
        messages = self.get_messages(conversation_id, workbook_path, limit=limit)
        history = []
        for item in messages:
            content = item["content"]
            metadata = item.get("metadata") or {}
            if metadata.get("kind") == "proposal":
                status = metadata.get("status", "pending")
                state = {"pending": "propuesta mostrada; aún no guardada", "applied": "cambio confirmado y guardado",
                         "cancelled": "propuesta cancelada; no se guardó", "expired": "propuesta caducada; no se guardó"}.get(status, "propuesta no confirmada")
                actions = metadata.get("actions", [])
                content += "\n[Estado de la propuesta: " + state + ". Resumen: " + json.dumps(actions, ensure_ascii=False)[:3500] + "]"
            history.append({"role": item["role"], "content": content[:1800]})
        history = history[-max(1, min(int(limit), 16)):]
        compact = []
        total = 0
        for item in reversed(history):
            remaining = 10000 - total
            if remaining <= 0:
                break
            content = item["content"][:remaining]
            compact.append({"role": item["role"], "content": content})
            total += len(content)
        return list(reversed(compact))
