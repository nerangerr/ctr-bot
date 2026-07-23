import os
import logging
from typing import List, Dict, Any, Optional
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime

logger = logging.getLogger(__name__)

class DatabaseManager:
    def __init__(self, db_url: str):
        self.db_url = db_url
        self.init_tables()

    def get_connection(self):
        return psycopg2.connect(self.db_url)

    def init_tables(self):
        try:
            conn = self.get_connection()
            cur = conn.cursor()

            cur.execute("""
                CREATE TABLE IF NOT EXISTS employees (
                    id SERIAL PRIMARY KEY,
                    full_name TEXT UNIQUE NOT NULL,
                    telegram_id BIGINT UNIQUE NOT NULL,
                    registered_at TIMESTAMP DEFAULT NOW()
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS assignments (
                    id SERIAL PRIMARY KEY,
                    user_id BIGINT NOT NULL,
                    short_task TEXT NOT NULL,
                    description TEXT,
                    deadline DATE,
                    responsible TEXT[],
                    protocol_number TEXT,
                    status TEXT DEFAULT 'active',
                    created_at TIMESTAMP DEFAULT NOW(),
                    completed_at TIMESTAMP,
                    owner_id BIGINT,
                    completion_report TEXT
                );
            """)

            conn.commit()
            cur.close()
            conn.close()
            logger.info("PostgreSQL tables initialized")
        except Exception as e:
            logger.error(f"DB init error: {e}")
            raise

    def register_employee(self, full_name: str, telegram_id: int) -> bool:
        try:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO employees (full_name, telegram_id)
                VALUES (%s, %s)
                ON CONFLICT (telegram_id) DO UPDATE SET full_name = EXCLUDED.full_name
            """, (full_name, telegram_id))
            conn.commit()
            cur.close()
            conn.close()
            return True
        except Exception as e:
            logger.error(f"Register employee error: {e}")
            return False

    def get_employee_by_name(self, full_name: str) -> Optional[Dict[str, Any]]:
        try:
            conn = self.get_connection()
            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute("SELECT * FROM employees WHERE full_name = %s", (full_name,))
            result = cur.fetchone()
            cur.close()
            conn.close()
            return dict(result) if result else None
        except Exception as e:
            logger.error(f"Get employee error: {e}")
            return None

    def get_employee_by_id(self, telegram_id: int) -> Optional[Dict[str, Any]]:
        try:
            conn = self.get_connection()
            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute("SELECT * FROM employees WHERE telegram_id = %s", (telegram_id,))
            result = cur.fetchone()
            cur.close()
            conn.close()
            return dict(result) if result else None
        except Exception as e:
            logger.error(f"Get employee by id error: {e}")
            return None

    def get_all_employees(self) -> List[Dict[str, Any]]:
        try:
            conn = self.get_connection()
            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute("SELECT * FROM employees ORDER BY full_name")
            results = cur.fetchall()
            cur.close()
            conn.close()
            return [dict(r) for r in results]
        except Exception as e:
            logger.error(f"Get all employees error: {e}")
            return []

    def save_assignment(self, user_id: int, assignment: Dict[str, Any]) -> int:
        try:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO assignments
                (user_id, short_task, description, deadline, responsible, protocol_number, status, owner_id, completion_report, completed_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (
                user_id,
                assignment.get('short_task', ''),
                assignment.get('description', ''),
                assignment.get('deadline'),
                assignment.get('responsible', []),
                assignment.get('protocol_number'),
                assignment.get('status', 'active'),
                assignment.get('owner_id'),
                assignment.get('completion_report'),
                assignment.get('completed_at')
            ))
            assignment_id = cur.fetchone()[0]
            conn.commit()
            cur.close()
            conn.close()
            return assignment_id
        except Exception as e:
            logger.error(f"DB save error: {e}")
            return 0

    # ========== ИСПРАВЛЕННЫЙ МЕТОД (сравнение по ФИО) ==========

    def get_user_assignments(self, user_id: int, user_name: str) -> List[Dict[str, Any]]:
        """
        Возвращает поручения, где пользователь является владельцем (user_id)
        ИЛИ указан в поле responsible (как ответственный) по ФИО.
        Поле deadline преобразуется в строку YYYY-MM-DD.
        """
        try:
            conn = self.get_connection()
            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute("""
                SELECT * FROM assignments
                WHERE user_id = %s OR %s = ANY(responsible)
                ORDER BY id
            """, (user_id, user_name))
            rows = cur.fetchall()
            cur.close()
            conn.close()
            result = []
            for row in rows:
                d = dict(row)
                if d.get("deadline") is not None:
                    d["deadline"] = d["deadline"].isoformat()
                result.append(d)
            return result
        except Exception as e:
            logger.error(f"DB fetch error: {e}")
            return []

    # =========================================================

    def delete_assignment(self, user_id: int, assignment_id: int) -> bool:
        try:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("DELETE FROM assignments WHERE user_id = %s AND id = %s", (user_id, assignment_id))
            affected = cur.rowcount
            conn.commit()
            cur.close()
            conn.close()
            return affected > 0
        except Exception as e:
            logger.error(f"DB delete error: {e}")
            return False

    def update_assignment(self, user_id: int, assignment_id: int, field: str, value: str) -> bool:
        try:
            conn = self.get_connection()
            cur = conn.cursor()
            if field == "deadline":
                cur.execute("UPDATE assignments SET deadline = %s WHERE user_id = %s AND id = %s", (value, user_id, assignment_id))
            elif field == "responsible":
                resp_list = [v.strip() for v in value.split(",") if v.strip()]
                cur.execute("UPDATE assignments SET responsible = %s WHERE user_id = %s AND id = %s", (resp_list, user_id, assignment_id))
            elif field == "short_task":
                cur.execute("UPDATE assignments SET short_task = %s WHERE user_id = %s AND id = %s", (value, user_id, assignment_id))
            elif field == "description":
                cur.execute("UPDATE assignments SET description = %s WHERE user_id = %s AND id = %s", (value, user_id, assignment_id))
            else:
                return False
            affected = cur.rowcount
            conn.commit()
            cur.close()
            conn.close()
            return affected > 0
        except Exception as e:
            logger.error(f"DB update error: {e}")
            return False

    def mark_completed(self, user_id: int, assignment_id: int) -> bool:
        try:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("""
                UPDATE assignments
                SET status = 'completed', completed_at = NOW()
                WHERE user_id = %s AND id = %s
            """, (user_id, assignment_id))
            affected = cur.rowcount
            conn.commit()
            cur.close()
            conn.close()
            return affected > 0
        except Exception as e:
            logger.error(f"DB complete error: {e}")
            return False

    def clear_user_assignments(self, user_id: int) -> bool:
        try:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("DELETE FROM assignments WHERE user_id = %s", (user_id,))
            conn.commit()
            cur.close()
            conn.close()
            return True
        except Exception as e:
            logger.error(f"DB clear error: {e}")
            return False

    def get_all_assignments(self) -> List[Dict[str, Any]]:
        """Возвращает все активные поручения с deadline в виде строки YYYY-MM-DD."""
        try:
            conn = self.get_connection()
            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute("SELECT * FROM assignments WHERE status != 'completed'")
            rows = cur.fetchall()
            cur.close()
            conn.close()
            result = []
            for row in rows:
                d = dict(row)
                if d.get("deadline") is not None:
                    d["deadline"] = d["deadline"].isoformat()
                result.append(d)
            return result
        except Exception as e:
            logger.error(f"DB all assignments error: {e}")
            return []