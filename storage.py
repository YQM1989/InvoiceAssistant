from __future__ import annotations

import re
import shutil
import sqlite3
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional


DB_FILE = "invoice_assistant.db"
SCHEMA_VERSION = "3"
PARTY_TYPES = ("seller", "buyer")
INVOICE_POOLS = ("approved", "pending")
TABLE_BY_POOL = {
    "approved": "approved_invoices",
    "pending": "pending_invoices",
}
INVALID_COMPANY_TOKENS = {
    "",
    "名称",
    "名称:",
    "名称：",
    "名 称",
    "销售方",
    "购买方",
    "销售方名称",
    "购买方名称",
    "购方名称",
    "销方名称",
    "统一社会信用代码",
    "纳税人识别号",
}


@dataclass
class InvoiceRecord:
    id: int
    seller_name: str
    seller_name_norm: str
    buyer_name: str
    buyer_name_norm: str
    invoice_number: str
    invoice_number_norm: str
    amount: float
    invoice_date: str
    source_path: str
    original_filename: str
    stored_filename: str
    imported_at: str
    needs_review: int
    review_reason: str
    approval_status: str
    approval_reason: str
    batch_id: str
    file_finalized: int
    raw_text: str


@dataclass
class WhitelistEntry:
    id: int
    party_type: str
    standard_name: str
    standard_norm: str
    created_at: str


@dataclass
class WhitelistAlias:
    id: int
    whitelist_id: int
    alias_name: str
    alias_norm: str
    created_at: str


def clean_company_name(name: str) -> str:
    value = (name or "").strip()
    if not value:
        return ""
    value = unicodedata.normalize("NFKC", value)
    value = value.replace("\u3000", " ")
    value = value.replace("（", "(").replace("）", ")")
    value = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", value)
    value = re.sub(r"[\t\r\n]+", " ", value)
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"^(?:购买方信息|销售方信息|购买方|销售方|购方|销方)?\s*名称[:：\s]*", "", value, flags=re.IGNORECASE)
    value = value.strip(" :：;；")
    return value


def normalize_company(name: str) -> str:
    value = clean_company_name(name)
    if not value:
        return ""
    return value.upper()


def is_valid_company_name(name: str) -> bool:
    value = clean_company_name(name)
    if not value:
        return False
    compact = value.replace(" ", "")
    if compact in INVALID_COMPANY_TOKENS:
        return False
    if re.fullmatch(r"[:：]+", compact):
        return False
    if re.fullmatch(r"[A-Z0-9]+", compact):
        return False
    return len(compact) >= 4


def normalize_invoice_number(invoice_number: str) -> str:
    if not invoice_number:
        return ""
    normalized = "".join(invoice_number.split()).upper()
    return "".join(ch for ch in normalized if ch.isalnum())


class InvoiceStore:
    SORTABLE_FIELDS = {
        "id": "id",
        "seller": "seller_name_norm",
        "buyer": "buyer_name_norm",
        "invoice_no": "invoice_number_norm",
        "amount": "amount",
        "date": "invoice_date",
        "approval": "approval_status",
        "review": "needs_review",
        "file": "stored_filename",
        "imported_at": "imported_at",
    }

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._ensure_db()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _ensure_db(self) -> None:
        existed = self.db_path.exists() and self.db_path.stat().st_size > 0
        with self._connect() as conn:
            self._ensure_meta_table(conn)
            current_version = self._get_meta(conn, "schema_version")
            if existed and current_version and current_version != SCHEMA_VERSION:
                self._backup_db()
            self._create_tables(conn)
            self._create_indexes(conn)
            self._set_meta(conn, "schema_version", SCHEMA_VERSION)

    def _ensure_meta_table(self, conn: sqlite3.Connection) -> None:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS app_meta (
                meta_key TEXT PRIMARY KEY,
                meta_value TEXT NOT NULL
            )
            """
        )

    def _get_meta(self, conn: sqlite3.Connection, key: str) -> str:
        row = conn.execute("SELECT meta_value FROM app_meta WHERE meta_key = ?", (key,)).fetchone()
        return str(row[0]) if row else ""

    def _set_meta(self, conn: sqlite3.Connection, key: str, value: str) -> None:
        conn.execute(
            """
            INSERT INTO app_meta(meta_key, meta_value) VALUES (?, ?)
            ON CONFLICT(meta_key) DO UPDATE SET meta_value = excluded.meta_value
            """,
            (key, value),
        )

    def _backup_db(self) -> None:
        backup_name = f"{self.db_path.stem}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{self.db_path.suffix}"
        backup_path = self.db_path.with_name(backup_name)
        if self.db_path.exists() and not backup_path.exists():
            shutil.copy2(self.db_path, backup_path)

    def _create_tables(self, conn: sqlite3.Connection) -> None:
        self._create_invoice_table(conn, TABLE_BY_POOL["approved"], "normal")
        self._create_invoice_table(conn, TABLE_BY_POOL["pending"], "pending")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS whitelist_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                party_type TEXT NOT NULL,
                standard_name TEXT NOT NULL,
                standard_norm TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(party_type, standard_norm)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS whitelist_aliases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                whitelist_id INTEGER NOT NULL,
                alias_name TEXT NOT NULL,
                alias_norm TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(whitelist_id, alias_norm),
                FOREIGN KEY(whitelist_id) REFERENCES whitelist_entries(id) ON DELETE CASCADE
            )
            """
        )

    def _create_invoice_table(self, conn: sqlite3.Connection, table_name: str, default_status: str) -> None:
        conn.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {table_name} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                seller_name TEXT NOT NULL DEFAULT '',
                seller_name_norm TEXT NOT NULL DEFAULT '',
                buyer_name TEXT NOT NULL DEFAULT '',
                buyer_name_norm TEXT NOT NULL DEFAULT '',
                invoice_number TEXT NOT NULL DEFAULT '',
                invoice_number_norm TEXT NOT NULL DEFAULT '',
                amount REAL NOT NULL DEFAULT 0,
                invoice_date TEXT NOT NULL DEFAULT '',
                source_path TEXT NOT NULL DEFAULT '',
                original_filename TEXT NOT NULL DEFAULT '',
                stored_filename TEXT NOT NULL DEFAULT '',
                imported_at TEXT NOT NULL DEFAULT '',
                needs_review INTEGER NOT NULL DEFAULT 0,
                review_reason TEXT NOT NULL DEFAULT '',
                approval_status TEXT NOT NULL DEFAULT '{default_status}',
                approval_reason TEXT NOT NULL DEFAULT '',
                batch_id TEXT NOT NULL DEFAULT '',
                file_finalized INTEGER NOT NULL DEFAULT 0,
                raw_text TEXT NOT NULL DEFAULT ''
            )
            """
        )

    def _create_indexes(self, conn: sqlite3.Connection) -> None:
        for table_name in TABLE_BY_POOL.values():
            conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table_name}_date ON {table_name}(invoice_date)")
            conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table_name}_seller ON {table_name}(seller_name_norm)")
            conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table_name}_buyer ON {table_name}(buyer_name_norm)")
            conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table_name}_invoice_no ON {table_name}(invoice_number_norm)")
            conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table_name}_approval ON {table_name}(approval_status)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_whitelist_type ON whitelist_entries(party_type, standard_name)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_whitelist_alias_norm ON whitelist_aliases(alias_norm)")

    def add_invoice(
        self,
        seller_name: str,
        buyer_name: str,
        invoice_number: str,
        amount: float,
        invoice_date: str,
        source_path: str,
        original_filename: str,
        stored_filename: str,
        needs_review: int,
        review_reason: str,
        approval_status: str,
        approval_reason: str,
        batch_id: str,
        file_finalized: int,
        raw_text: str,
        pool: str = "approved",
        imported_at: str | None = None,
    ) -> int:
        table_name = self._table_for_pool(pool)
        imported_at = imported_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            cur = conn.execute(
                f"""
                INSERT INTO {table_name} (
                    seller_name, seller_name_norm, buyer_name, buyer_name_norm,
                    invoice_number, invoice_number_norm, amount, invoice_date, source_path,
                    original_filename, stored_filename, imported_at, needs_review, review_reason,
                    approval_status, approval_reason, batch_id, file_finalized, raw_text
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    clean_company_name(seller_name),
                    normalize_company(seller_name),
                    clean_company_name(buyer_name),
                    normalize_company(buyer_name),
                    invoice_number,
                    normalize_invoice_number(invoice_number),
                    amount,
                    invoice_date or "",
                    source_path,
                    original_filename,
                    stored_filename,
                    imported_at,
                    int(needs_review),
                    review_reason,
                    approval_status,
                    approval_reason,
                    batch_id,
                    int(file_finalized),
                    raw_text,
                ),
            )
            return int(cur.lastrowid)

    def invoice_exists(
        self,
        seller_name: str,
        buyer_name: str,
        invoice_number: str,
        pool: str = "all",
        exclude_pool: Optional[str] = None,
        exclude_id: Optional[int] = None,
    ) -> bool:
        seller_norm = normalize_company(seller_name)
        buyer_norm = normalize_company(buyer_name)
        invoice_norm = normalize_invoice_number(invoice_number)
        if not seller_norm or not buyer_norm or not invoice_norm:
            return False

        table_names = self._tables_for_scope(pool)
        with self._connect() as conn:
            for current_table in table_names:
                sql = (
                    f"SELECT 1 FROM {current_table} "
                    "WHERE seller_name_norm = ? AND buyer_name_norm = ? AND invoice_number_norm = ?"
                )
                params: list[object] = [seller_norm, buyer_norm, invoice_norm]
                if exclude_id is not None and exclude_pool and current_table == self._table_for_pool(exclude_pool):
                    sql += " AND id <> ?"
                    params.append(exclude_id)
                row = conn.execute(sql, params).fetchone()
                if row is not None:
                    return True
        return False

    def get_invoice(self, invoice_id: int, pool: str = "approved") -> Optional[InvoiceRecord]:
        table_name = self._table_for_pool(pool)
        with self._connect() as conn:
            row = conn.execute(f"SELECT * FROM {table_name} WHERE id = ?", (invoice_id,)).fetchone()
        return self._row_to_invoice(row) if row else None

    def get_invoices(self, invoice_ids: Iterable[int], pool: str = "approved") -> list[InvoiceRecord]:
        ids = [int(x) for x in invoice_ids]
        if not ids:
            return []
        table_name = self._table_for_pool(pool)
        placeholders = ",".join("?" for _ in ids)
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT * FROM {table_name} WHERE id IN ({placeholders}) ORDER BY id",
                ids,
            ).fetchall()
        return [self._row_to_invoice(row) for row in rows]

    def update_invoice_core(
        self,
        invoice_id: int,
        seller_name: str,
        buyer_name: str,
        invoice_number: str,
        amount: float,
        invoice_date: str,
        needs_review: int,
        review_reason: str,
        approval_status: str,
        approval_reason: str,
        pool: str = "approved",
    ) -> None:
        table_name = self._table_for_pool(pool)
        with self._connect() as conn:
            conn.execute(
                f"""
                UPDATE {table_name}
                SET seller_name = ?,
                    seller_name_norm = ?,
                    buyer_name = ?,
                    buyer_name_norm = ?,
                    invoice_number = ?,
                    invoice_number_norm = ?,
                    amount = ?,
                    invoice_date = ?,
                    needs_review = ?,
                    review_reason = ?,
                    approval_status = ?,
                    approval_reason = ?
                WHERE id = ?
                """,
                (
                    clean_company_name(seller_name),
                    normalize_company(seller_name),
                    clean_company_name(buyer_name),
                    normalize_company(buyer_name),
                    invoice_number,
                    normalize_invoice_number(invoice_number),
                    amount,
                    invoice_date or "",
                    int(needs_review),
                    review_reason,
                    approval_status,
                    approval_reason,
                    invoice_id,
                ),
            )

    def update_invoice_parties(
        self,
        invoice_ids: Iterable[int],
        seller_name: Optional[str] = None,
        buyer_name: Optional[str] = None,
        approval_status: Optional[str] = None,
        approval_reason: Optional[str] = None,
        pool: str = "approved",
    ) -> None:
        ids = [int(x) for x in invoice_ids]
        if not ids:
            return
        table_name = self._table_for_pool(pool)
        sets: list[str] = []
        params: list[object] = []
        if seller_name is not None:
            cleaned = clean_company_name(seller_name)
            sets.extend(["seller_name = ?", "seller_name_norm = ?"])
            params.extend([cleaned, normalize_company(cleaned)])
        if buyer_name is not None:
            cleaned = clean_company_name(buyer_name)
            sets.extend(["buyer_name = ?", "buyer_name_norm = ?"])
            params.extend([cleaned, normalize_company(cleaned)])
        if approval_status is not None:
            sets.append("approval_status = ?")
            params.append(approval_status)
        if approval_reason is not None:
            sets.append("approval_reason = ?")
            params.append(approval_reason)
        if not sets:
            return
        placeholders = ",".join("?" for _ in ids)
        with self._connect() as conn:
            conn.execute(
                f"UPDATE {table_name} SET {', '.join(sets)} WHERE id IN ({placeholders})",
                [*params, *ids],
            )

    def update_invoice_file(
        self,
        invoice_id: int,
        source_path: str,
        stored_filename: str,
        file_finalized: int,
        pool: str = "approved",
    ) -> None:
        table_name = self._table_for_pool(pool)
        with self._connect() as conn:
            conn.execute(
                f"""
                UPDATE {table_name}
                SET source_path = ?, stored_filename = ?, file_finalized = ?
                WHERE id = ?
                """,
                (source_path, stored_filename, int(file_finalized), invoice_id),
            )

    def delete_invoices(self, invoice_ids: Iterable[int], pool: str = "approved") -> None:
        ids = [int(x) for x in invoice_ids]
        if not ids:
            return
        table_name = self._table_for_pool(pool)
        placeholders = ",".join("?" for _ in ids)
        with self._connect() as conn:
            conn.execute(f"DELETE FROM {table_name} WHERE id IN ({placeholders})", ids)

    def query_invoices(
        self,
        seller_keyword: str = "",
        buyer_keyword: str = "",
        start_month: str = "",
        end_month: str = "",
        approval_status: str = "",
        keyword: str = "",
        page: int = 1,
        page_size: int = 20,
        order_by: str = "date",
        ascending: bool = False,
        pool: str = "approved",
    ) -> tuple[list[InvoiceRecord], int]:
        table_name = self._table_for_pool(pool)
        where_sql, params = self._build_invoice_filter(
            seller_keyword,
            buyer_keyword,
            start_month,
            end_month,
            approval_status,
            keyword,
        )
        order_sql = self._build_order_sql(order_by, ascending)
        with self._connect() as conn:
            total = int(conn.execute(f"SELECT COUNT(1) FROM {table_name} WHERE {where_sql}", params).fetchone()[0])
            offset = max(0, (page - 1) * page_size)
            rows = conn.execute(
                f"SELECT * FROM {table_name} WHERE {where_sql} ORDER BY {order_sql} LIMIT ? OFFSET ?",
                [*params, page_size, offset],
            ).fetchall()
        return [self._row_to_invoice(row) for row in rows], total

    def query_all_filtered(
        self,
        seller_keyword: str = "",
        buyer_keyword: str = "",
        start_month: str = "",
        end_month: str = "",
        approval_status: str = "",
        keyword: str = "",
        order_by: str = "date",
        ascending: bool = False,
        pool: str = "approved",
    ) -> list[InvoiceRecord]:
        table_name = self._table_for_pool(pool)
        where_sql, params = self._build_invoice_filter(
            seller_keyword,
            buyer_keyword,
            start_month,
            end_month,
            approval_status,
            keyword,
        )
        order_sql = self._build_order_sql(order_by, ascending)
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT * FROM {table_name} WHERE {where_sql} ORDER BY {order_sql}",
                params,
            ).fetchall()
        return [self._row_to_invoice(row) for row in rows]

    def get_pending_invoices(self) -> list[InvoiceRecord]:
        return self.query_all_filtered(pool="pending", order_by="imported_at", ascending=False)

    def approve_pending_invoice(
        self,
        pending_id: int,
        seller_name: str,
        buyer_name: str,
        source_path: str,
        stored_filename: str,
    ) -> int:
        with self._connect() as conn:
            row = conn.execute(
                f"SELECT * FROM {TABLE_BY_POOL['pending']} WHERE id = ?",
                (pending_id,),
            ).fetchone()
            if not row:
                raise ValueError(f"Pending invoice {pending_id} not found")
            record = self._row_to_invoice(row)
            cur = conn.execute(
                f"""
                INSERT INTO {TABLE_BY_POOL['approved']} (
                    seller_name, seller_name_norm, buyer_name, buyer_name_norm,
                    invoice_number, invoice_number_norm, amount, invoice_date, source_path,
                    original_filename, stored_filename, imported_at, needs_review, review_reason,
                    approval_status, approval_reason, batch_id, file_finalized, raw_text
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    clean_company_name(seller_name),
                    normalize_company(seller_name),
                    clean_company_name(buyer_name),
                    normalize_company(buyer_name),
                    record.invoice_number,
                    normalize_invoice_number(record.invoice_number),
                    record.amount,
                    record.invoice_date,
                    source_path,
                    record.original_filename,
                    stored_filename,
                    record.imported_at,
                    record.needs_review,
                    record.review_reason,
                    "normal",
                    "",
                    record.batch_id,
                    1,
                    record.raw_text,
                ),
            )
            conn.execute(f"DELETE FROM {TABLE_BY_POOL['pending']} WHERE id = ?", (pending_id,))
            return int(cur.lastrowid)

    def get_all_party_names(self, party_type: str) -> list[str]:
        self._validate_party_type(party_type)
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT standard_name FROM whitelist_entries WHERE party_type = ? ORDER BY standard_name",
                (party_type,),
            ).fetchall()
        return [str(row[0]) for row in rows]

    def get_all_invoice_party_names(self, party_type: str, pool: str = "approved") -> list[str]:
        self._validate_party_type(party_type)
        if pool == "all":
            values = set(self.get_all_invoice_party_names(party_type, "approved"))
            values.update(self.get_all_invoice_party_names(party_type, "pending"))
            return sorted(values)
        table_name = self._table_for_pool(pool)
        column = "seller_name" if party_type == "seller" else "buyer_name"
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT DISTINCT {column} FROM {table_name} WHERE {column} <> '' ORDER BY {column}"
            ).fetchall()
        return [str(row[0]) for row in rows]

    def get_all_months(self, pool: str = "approved") -> list[str]:
        table_name = self._table_for_pool(pool)
        with self._connect() as conn:
            rows = conn.execute(
                f"""
                SELECT DISTINCT substr(invoice_date, 1, 7) AS month_value
                FROM {table_name}
                WHERE length(invoice_date) >= 7
                ORDER BY month_value DESC
                """
            ).fetchall()
        return [str(row[0]) for row in rows if row[0]]

    def get_filtered_total_amount(
        self,
        seller_keyword: str = "",
        buyer_keyword: str = "",
        start_month: str = "",
        end_month: str = "",
        approval_status: str = "",
        keyword: str = "",
        pool: str = "approved",
    ) -> float:
        table_name = self._table_for_pool(pool)
        where_sql, params = self._build_invoice_filter(
            seller_keyword,
            buyer_keyword,
            start_month,
            end_month,
            approval_status,
            keyword,
        )
        with self._connect() as conn:
            value = conn.execute(f"SELECT COALESCE(SUM(amount), 0) FROM {table_name} WHERE {where_sql}", params).fetchone()[0]
        return float(value or 0)

    def get_whitelist_entries(self, party_type: str) -> list[WhitelistEntry]:
        self._validate_party_type(party_type)
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM whitelist_entries WHERE party_type = ? ORDER BY standard_name",
                (party_type,),
            ).fetchall()
        return [self._row_to_whitelist_entry(row) for row in rows]

    def get_whitelist_entry_by_norm(self, party_type: str, normalized_name: str) -> Optional[WhitelistEntry]:
        self._validate_party_type(party_type)
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM whitelist_entries WHERE party_type = ? AND standard_norm = ?",
                (party_type, normalized_name),
            ).fetchone()
        return self._row_to_whitelist_entry(row) if row else None

    def get_whitelist_aliases(self, whitelist_id: int) -> list[WhitelistAlias]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM whitelist_aliases WHERE whitelist_id = ? ORDER BY alias_name",
                (whitelist_id,),
            ).fetchall()
        return [self._row_to_whitelist_alias(row) for row in rows]

    def get_whitelist_map(self, party_type: str) -> dict[str, str]:
        self._validate_party_type(party_type)
        mapping: dict[str, str] = {}
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT e.standard_name, e.standard_norm, a.alias_name, a.alias_norm
                FROM whitelist_entries e
                LEFT JOIN whitelist_aliases a ON a.whitelist_id = e.id
                WHERE e.party_type = ?
                ORDER BY e.standard_name
                """,
                (party_type,),
            ).fetchall()
        for row in rows:
            mapping[str(row["standard_norm"])] = str(row["standard_name"])
            if row["alias_norm"]:
                mapping[str(row["alias_norm"])] = str(row["standard_name"])
        return mapping

    def resolve_standard_name(self, party_type: str, name: str) -> str:
        norm = normalize_company(name)
        if not norm:
            return ""
        mapping = self.get_whitelist_map(party_type)
        return mapping.get(norm, clean_company_name(name))

    def add_whitelist_item(
        self,
        party_type: str,
        name: str,
        conn: sqlite3.Connection | None = None,
    ) -> tuple[bool, str]:
        self._validate_party_type(party_type)
        cleaned = clean_company_name(name)
        norm = normalize_company(cleaned)
        if not is_valid_company_name(cleaned):
            return False, ""
        close_after = False
        if conn is None:
            conn = self._connect()
            close_after = True
        try:
            existing = conn.execute(
                "SELECT standard_name FROM whitelist_entries WHERE party_type = ? AND standard_norm = ?",
                (party_type, norm),
            ).fetchone()
            if existing:
                return False, str(existing[0])
            conn.execute(
                """
                INSERT INTO whitelist_entries(party_type, standard_name, standard_norm, created_at)
                VALUES (?, ?, ?, ?)
                """,
                (party_type, cleaned, norm, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            )
            if close_after:
                conn.commit()
            return True, cleaned
        finally:
            if close_after:
                conn.close()

    def add_whitelist_items(
        self,
        party_type: str,
        names: Iterable[str],
        conn: sqlite3.Connection | None = None,
    ) -> tuple[int, list[str]]:
        inserted = 0
        duplicates: list[str] = []
        close_after = False
        if conn is None:
            conn = self._connect()
            close_after = True
        try:
            for name in names:
                ok, existing = self.add_whitelist_item(party_type, name, conn=conn)
                cleaned = clean_company_name(name)
                if ok:
                    inserted += 1
                elif cleaned:
                    duplicates.append(existing or cleaned)
            if close_after:
                conn.commit()
            return inserted, duplicates
        finally:
            if close_after:
                conn.close()

    def update_whitelist_item(self, entry_id: int, new_name: str) -> tuple[bool, str]:
        cleaned = clean_company_name(new_name)
        norm = normalize_company(cleaned)
        if not is_valid_company_name(cleaned):
            return False, "名称无效"
        with self._connect() as conn:
            row = conn.execute("SELECT party_type FROM whitelist_entries WHERE id = ?", (entry_id,)).fetchone()
            if not row:
                return False, "记录不存在"
            conflict = conn.execute(
                "SELECT 1 FROM whitelist_entries WHERE party_type = ? AND standard_norm = ? AND id <> ?",
                (str(row[0]), norm, entry_id),
            ).fetchone()
            if conflict:
                return False, "规范化后与现有白名单重复"
            conn.execute(
                "UPDATE whitelist_entries SET standard_name = ?, standard_norm = ? WHERE id = ?",
                (cleaned, norm, entry_id),
            )
        return True, ""

    def delete_whitelist_item(self, entry_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM whitelist_entries WHERE id = ?", (entry_id,))

    def add_whitelist_alias(self, whitelist_id: int, alias_name: str) -> tuple[bool, str]:
        cleaned = clean_company_name(alias_name)
        norm = normalize_company(cleaned)
        if not is_valid_company_name(cleaned):
            return False, "别名无效"
        with self._connect() as conn:
            row = conn.execute("SELECT party_type FROM whitelist_entries WHERE id = ?", (whitelist_id,)).fetchone()
            if not row:
                return False, "白名单项不存在"
            conflict = conn.execute(
                """
                SELECT 1
                FROM whitelist_aliases a
                INNER JOIN whitelist_entries e ON e.id = a.whitelist_id
                WHERE e.party_type = ? AND a.alias_norm = ?
                """,
                (str(row[0]), norm),
            ).fetchone()
            if conflict:
                return False, "别名重复"
            standard_conflict = conn.execute(
                "SELECT 1 FROM whitelist_entries WHERE party_type = ? AND standard_norm = ?",
                (str(row[0]), norm),
            ).fetchone()
            if standard_conflict:
                return False, "别名与已有标准名重复"
            conn.execute(
                """
                INSERT INTO whitelist_aliases(whitelist_id, alias_name, alias_norm, created_at)
                VALUES (?, ?, ?, ?)
                """,
                (whitelist_id, cleaned, norm, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            )
        return True, ""

    def update_whitelist_alias(self, alias_id: int, new_name: str) -> tuple[bool, str]:
        cleaned = clean_company_name(new_name)
        norm = normalize_company(cleaned)
        if not is_valid_company_name(cleaned):
            return False, "别名无效"
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT a.whitelist_id, e.party_type
                FROM whitelist_aliases a
                INNER JOIN whitelist_entries e ON e.id = a.whitelist_id
                WHERE a.id = ?
                """,
                (alias_id,),
            ).fetchone()
            if not row:
                return False, "别名不存在"
            conflict = conn.execute(
                """
                SELECT 1
                FROM whitelist_aliases a
                INNER JOIN whitelist_entries e ON e.id = a.whitelist_id
                WHERE e.party_type = ? AND a.alias_norm = ? AND a.id <> ?
                """,
                (str(row["party_type"]), norm, alias_id),
            ).fetchone()
            if conflict:
                return False, "规范化后与现有别名重复"
            conn.execute(
                "UPDATE whitelist_aliases SET alias_name = ?, alias_norm = ? WHERE id = ?",
                (cleaned, norm, alias_id),
            )
        return True, ""

    def delete_whitelist_alias(self, alias_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM whitelist_aliases WHERE id = ?", (alias_id,))

    def _build_invoice_filter(
        self,
        seller_keyword: str = "",
        buyer_keyword: str = "",
        start_month: str = "",
        end_month: str = "",
        approval_status: str = "",
        keyword: str = "",
    ) -> tuple[str, list[object]]:
        where = ["1=1"]
        params: list[object] = []

        seller_norm = normalize_company(seller_keyword)
        if seller_norm:
            where.append("seller_name_norm LIKE ?")
            params.append(f"%{seller_norm}%")

        buyer_norm = normalize_company(buyer_keyword)
        if buyer_norm:
            where.append("buyer_name_norm LIKE ?")
            params.append(f"%{buyer_norm}%")

        start_month = start_month.strip()
        if start_month:
            where.append("substr(invoice_date, 1, 7) >= ?")
            params.append(start_month)

        end_month = end_month.strip()
        if end_month:
            where.append("substr(invoice_date, 1, 7) <= ?")
            params.append(end_month)

        approval_status = (approval_status or "").strip()
        if approval_status:
            where.append("approval_status = ?")
            params.append(approval_status)

        keyword = (keyword or "").strip()
        if keyword:
            keyword_norm = normalize_company(keyword)
            keyword_invoice = normalize_invoice_number(keyword)
            clauses = [
                "seller_name_norm LIKE ?",
                "buyer_name_norm LIKE ?",
                "invoice_number_norm LIKE ?",
                "stored_filename LIKE ?",
            ]
            where.append("(" + " OR ".join(clauses) + ")")
            params.extend([
                f"%{keyword_norm}%",
                f"%{keyword_norm}%",
                f"%{keyword_invoice or keyword_norm}%",
                f"%{keyword}%",
            ])

        return " AND ".join(where), params

    @classmethod
    def _build_order_sql(cls, order_by: str, ascending: bool) -> str:
        column = cls.SORTABLE_FIELDS.get(order_by, "invoice_date")
        direction = "ASC" if ascending else "DESC"
        if column in {"invoice_date", "imported_at"}:
            return f"{column} {direction}, id {direction}"
        if column == "amount":
            return f"{column} {direction}, invoice_date DESC, id DESC"
        return f"{column} {direction}, invoice_date DESC, id DESC"

    @staticmethod
    def _row_to_invoice(row: sqlite3.Row) -> InvoiceRecord:
        return InvoiceRecord(
            id=int(row["id"]),
            seller_name=str(row["seller_name"] or ""),
            seller_name_norm=str(row["seller_name_norm"] or ""),
            buyer_name=str(row["buyer_name"] or ""),
            buyer_name_norm=str(row["buyer_name_norm"] or ""),
            invoice_number=str(row["invoice_number"] or ""),
            invoice_number_norm=str(row["invoice_number_norm"] or ""),
            amount=float(row["amount"] or 0),
            invoice_date=str(row["invoice_date"] or ""),
            source_path=str(row["source_path"] or ""),
            original_filename=str(row["original_filename"] or ""),
            stored_filename=str(row["stored_filename"] or ""),
            imported_at=str(row["imported_at"] or ""),
            needs_review=int(row["needs_review"] or 0),
            review_reason=str(row["review_reason"] or ""),
            approval_status=str(row["approval_status"] or "pending"),
            approval_reason=str(row["approval_reason"] or ""),
            batch_id=str(row["batch_id"] or ""),
            file_finalized=int(row["file_finalized"] or 0),
            raw_text=str(row["raw_text"] or ""),
        )

    @staticmethod
    def _row_to_whitelist_entry(row: sqlite3.Row) -> WhitelistEntry:
        return WhitelistEntry(
            id=int(row["id"]),
            party_type=str(row["party_type"]),
            standard_name=str(row["standard_name"]),
            standard_norm=str(row["standard_norm"]),
            created_at=str(row["created_at"]),
        )

    @staticmethod
    def _row_to_whitelist_alias(row: sqlite3.Row) -> WhitelistAlias:
        return WhitelistAlias(
            id=int(row["id"]),
            whitelist_id=int(row["whitelist_id"]),
            alias_name=str(row["alias_name"]),
            alias_norm=str(row["alias_norm"]),
            created_at=str(row["created_at"]),
        )

    @staticmethod
    def _validate_party_type(party_type: str) -> None:
        if party_type not in PARTY_TYPES:
            raise ValueError(f"Unsupported party type: {party_type}")

    @staticmethod
    def _table_for_pool(pool: str) -> str:
        if pool not in TABLE_BY_POOL:
            raise ValueError(f"Unsupported invoice pool: {pool}")
        return TABLE_BY_POOL[pool]

    @staticmethod
    def _tables_for_scope(scope: str) -> list[str]:
        if scope == "all":
            return [TABLE_BY_POOL["approved"], TABLE_BY_POOL["pending"]]
        return [InvoiceStore._table_for_pool(scope)]
