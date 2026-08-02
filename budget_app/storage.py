from __future__ import annotations

import hashlib
import hmac
import os
import secrets
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable


DB_PATH = Path(os.environ.get("DB_PATH", "data/orcamentos.db"))


@dataclass
class Item:
    codigo: str
    produto: str
    pmc: str
    desconto: str
    valor_unitario: Decimal
    quantidade: Decimal

    @property
    def desconto_percentual(self) -> Decimal:
        value = (self.desconto or "").strip().replace("%", "").replace(".", "").replace(",", ".")
        if not value:
            return Decimal("0")
        try:
            desconto = Decimal(value)
        except InvalidOperation as exc:
            raise ValueError(f"Desconto inválido: {self.desconto}") from exc
        if desconto < 0 or desconto > 100:
            raise ValueError("Desconto deve estar entre 0 e 100%.")
        return desconto

    @property
    def valor_unitario_com_desconto(self) -> Decimal:
        fator = Decimal("1") - (self.desconto_percentual / Decimal("100"))
        return self.valor_unitario * fator

    @property
    def total(self) -> Decimal:
        return (self.valor_unitario_com_desconto * self.quantidade).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )


@dataclass
class Orcamento:
    cliente_nome: str
    farmaceutico_responsavel: str
    cpf: str
    telefone: str
    email: str | None
    data_orcamento: str
    localidade: str
    itens: list[Item]

    @property
    def total(self) -> Decimal:
        return sum((item.total for item in self.itens), Decimal("0"))


def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS orcamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                criado_em TEXT NOT NULL,
                cliente_nome TEXT NOT NULL,
                farmaceutico_responsavel TEXT NOT NULL DEFAULT '',
                cpf TEXT NOT NULL,
                telefone TEXT NOT NULL,
                email TEXT,
                data_orcamento TEXT NOT NULL,
                localidade TEXT NOT NULL,
                total TEXT NOT NULL,
                docx_path TEXT NOT NULL,
                pdf_path TEXT,
                pdf_status TEXT NOT NULL
            )
            """
        )
        migrate_nullable_email(conn)
        migrate_farmaceutico_responsavel(conn)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS orcamento_itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                orcamento_id INTEGER NOT NULL,
                codigo TEXT NOT NULL,
                produto TEXT NOT NULL,
                pmc TEXT NOT NULL,
                desconto TEXT NOT NULL,
                valor_unitario TEXT NOT NULL,
                quantidade TEXT NOT NULL,
                total TEXT NOT NULL,
                FOREIGN KEY (orcamento_id) REFERENCES orcamentos(id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS farmaceuticos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE,
                ativo INTEGER NOT NULL DEFAULT 1,
                criado_em TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS app_users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                perfil TEXT NOT NULL DEFAULT 'admin',
                criado_em TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS app_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                token_hash TEXT NOT NULL UNIQUE,
                criado_em TEXT NOT NULL,
                expira_em TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES app_users(id)
            )
            """
        )
        bootstrap_admin_from_env(conn)
        seed_farmaceuticos_from_history(conn)


def migrate_nullable_email(conn: sqlite3.Connection) -> None:
    columns = conn.execute("PRAGMA table_info(orcamentos)").fetchall()
    email_column = next((column for column in columns if column["name"] == "email"), None)
    if email_column is None or not email_column["notnull"]:
        return

    conn.execute("ALTER TABLE orcamentos RENAME TO orcamentos_old")
    conn.execute(
        """
        CREATE TABLE orcamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            criado_em TEXT NOT NULL,
            cliente_nome TEXT NOT NULL,
            cpf TEXT NOT NULL,
            telefone TEXT NOT NULL,
            email TEXT,
            data_orcamento TEXT NOT NULL,
            localidade TEXT NOT NULL,
            total TEXT NOT NULL,
            docx_path TEXT NOT NULL,
            pdf_path TEXT,
            pdf_status TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        INSERT INTO orcamentos (
            id, criado_em, cliente_nome, cpf, telefone, email, data_orcamento,
            localidade, total, docx_path, pdf_path, pdf_status
        )
        SELECT
            id, criado_em, cliente_nome, cpf, telefone,
            NULLIF(NULLIF(email, ''), 'não tem'),
            data_orcamento, localidade, total, docx_path, pdf_path, pdf_status
        FROM orcamentos_old
        """
    )
    conn.execute("DROP TABLE orcamentos_old")


def migrate_farmaceutico_responsavel(conn: sqlite3.Connection) -> None:
    columns = conn.execute("PRAGMA table_info(orcamentos)").fetchall()
    if any(column["name"] == "farmaceutico_responsavel" for column in columns):
        return
    conn.execute(
        "ALTER TABLE orcamentos ADD COLUMN farmaceutico_responsavel TEXT NOT NULL DEFAULT ''"
    )


def seed_farmaceuticos_from_history(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        """
        SELECT DISTINCT farmaceutico_responsavel
        FROM orcamentos
        WHERE TRIM(farmaceutico_responsavel) <> ''
        """
    ).fetchall()
    now = datetime.now().isoformat(timespec="seconds")
    conn.executemany(
        """
        INSERT OR IGNORE INTO farmaceuticos (nome, ativo, criado_em)
        VALUES (?, 1, ?)
        """,
        [(row["farmaceutico_responsavel"], now) for row in rows],
    )


def list_farmaceuticos(include_inactive: bool = False) -> list[sqlite3.Row]:
    with connect() as conn:
        if include_inactive:
            query = "SELECT * FROM farmaceuticos ORDER BY ativo DESC, nome COLLATE NOCASE"
            return list(conn.execute(query))
        return list(
            conn.execute(
                """
                SELECT *
                FROM farmaceuticos
                WHERE ativo = 1
                ORDER BY nome COLLATE NOCASE
                """
            )
        )


def get_farmaceutico(farmaceutico_id: int) -> sqlite3.Row | None:
    with connect() as conn:
        return conn.execute(
            "SELECT * FROM farmaceuticos WHERE id = ?",
            (farmaceutico_id,),
        ).fetchone()


def add_farmaceutico(nome: str) -> None:
    nome = nome.strip()
    if not nome:
        raise ValueError("Informe o nome do farmacêutico(a).")
    with connect() as conn:
        conn.execute(
            """
            INSERT INTO farmaceuticos (nome, ativo, criado_em)
            VALUES (?, 1, ?)
            ON CONFLICT(nome) DO UPDATE SET ativo = 1
            """,
            (nome, datetime.now().isoformat(timespec="seconds")),
        )


def update_farmaceutico(farmaceutico_id: int, nome: str, ativo: bool = True) -> None:
    nome = nome.strip()
    if not nome:
        raise ValueError("Informe o nome do farmacêutico(a).")
    with connect() as conn:
        try:
            conn.execute(
                """
                UPDATE farmaceuticos
                SET nome = ?, ativo = ?
                WHERE id = ?
                """,
                (nome, 1 if ativo else 0, farmaceutico_id),
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError("Já existe um responsável com este nome.") from exc


def delete_farmaceutico(farmaceutico_id: int) -> None:
    with connect() as conn:
        conn.execute("DELETE FROM farmaceuticos WHERE id = ?", (farmaceutico_id,))


def password_hash(password: str, salt: str | None = None, iterations: int = 260_000) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        bytes.fromhex(salt),
        iterations,
    ).hex()
    return f"pbkdf2_sha256${iterations}${salt}${digest}"


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        algorithm, iterations_raw, salt, digest = stored_hash.split("$", 3)
        iterations = int(iterations_raw)
    except ValueError:
        return False
    if algorithm != "pbkdf2_sha256":
        return False
    candidate = password_hash(password, salt=salt, iterations=iterations).split("$", 3)[-1]
    return hmac.compare_digest(candidate, digest)


def token_hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def has_users() -> bool:
    with connect() as conn:
        row = conn.execute("SELECT 1 FROM app_users LIMIT 1").fetchone()
        return row is not None


def bootstrap_admin_from_env(conn: sqlite3.Connection) -> None:
    username = os.environ.get("APP_ADMIN_USER", "").strip()
    password = os.environ.get("APP_ADMIN_PASSWORD", "")
    if not username and not password:
        return
    if not username or not password:
        raise RuntimeError("Configure APP_ADMIN_USER e APP_ADMIN_PASSWORD juntos.")
    if len(password) < 8:
        raise RuntimeError("APP_ADMIN_PASSWORD precisa ter pelo menos 8 caracteres.")

    row = conn.execute("SELECT 1 FROM app_users LIMIT 1").fetchone()
    if row is not None:
        return

    conn.execute(
        """
        INSERT INTO app_users (username, password_hash, perfil, criado_em)
        VALUES (?, ?, 'admin', ?)
        """,
        (
            username,
            password_hash(password),
            datetime.now().isoformat(timespec="seconds"),
        ),
    )


def create_user(username: str, password: str, perfil: str = "admin") -> int:
    username = username.strip()
    if not username:
        raise ValueError("Informe o usuário.")
    if len(password) < 8:
        raise ValueError("A senha deve ter pelo menos 8 caracteres.")
    with connect() as conn:
        try:
            cur = conn.execute(
                """
                INSERT INTO app_users (username, password_hash, perfil, criado_em)
                VALUES (?, ?, ?, ?)
                """,
                (
                    username,
                    password_hash(password),
                    perfil,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError("Este usuário já existe.") from exc
        return int(cur.lastrowid)


def verify_login(username: str, password: str) -> sqlite3.Row | None:
    username = username.strip()
    with connect() as conn:
        row = conn.execute(
            "SELECT * FROM app_users WHERE username = ?",
            (username,),
        ).fetchone()
    if row is None or not verify_password(password, row["password_hash"]):
        return None
    return row


def delete_expired_sessions(conn: sqlite3.Connection) -> None:
    conn.execute(
        "DELETE FROM app_sessions WHERE expira_em <= ?",
        (datetime.now().isoformat(timespec="seconds"),),
    )


def create_session(user_id: int, hours: int = 12) -> str:
    token = secrets.token_urlsafe(32)
    now = datetime.now()
    with connect() as conn:
        delete_expired_sessions(conn)
        conn.execute(
            """
            INSERT INTO app_sessions (user_id, token_hash, criado_em, expira_em)
            VALUES (?, ?, ?, ?)
            """,
            (
                user_id,
                token_hash(token),
                now.isoformat(timespec="seconds"),
                (now + timedelta(hours=hours)).isoformat(timespec="seconds"),
            ),
        )
    return token


def get_user_by_session(token: str | None) -> sqlite3.Row | None:
    if not token:
        return None
    with connect() as conn:
        delete_expired_sessions(conn)
        return conn.execute(
            """
            SELECT u.id, u.username, u.perfil
            FROM app_sessions s
            JOIN app_users u ON u.id = s.user_id
            WHERE s.token_hash = ? AND s.expira_em > ?
            """,
            (token_hash(token), datetime.now().isoformat(timespec="seconds")),
        ).fetchone()


def delete_session(token: str | None) -> None:
    if not token:
        return
    with connect() as conn:
        conn.execute("DELETE FROM app_sessions WHERE token_hash = ?", (token_hash(token),))


def save_orcamento(
    orcamento: Orcamento,
    docx_path: Path,
    pdf_path: Path | None,
    pdf_status: str,
) -> int:
    with connect() as conn:
        cur = conn.execute(
            """
            INSERT INTO orcamentos (
                criado_em, cliente_nome, farmaceutico_responsavel, cpf, telefone,
                email, data_orcamento, localidade, total, docx_path, pdf_path,
                pdf_status
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().isoformat(timespec="seconds"),
                orcamento.cliente_nome,
                orcamento.farmaceutico_responsavel,
                orcamento.cpf,
                orcamento.telefone,
                orcamento.email,
                orcamento.data_orcamento,
                orcamento.localidade,
                str(orcamento.total),
                str(docx_path),
                str(pdf_path) if pdf_path else None,
                pdf_status,
            ),
        )
        orcamento_id = int(cur.lastrowid)
        conn.executemany(
            """
            INSERT INTO orcamento_itens (
                orcamento_id, codigo, produto, pmc, desconto, valor_unitario,
                quantidade, total
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    orcamento_id,
                    item.codigo,
                    item.produto,
                    item.pmc,
                    item.desconto,
                    str(item.valor_unitario),
                    str(item.quantidade),
                    str(item.total),
                )
                for item in orcamento.itens
            ],
        )
        return orcamento_id


def list_orcamentos(limit: int = 30) -> list[sqlite3.Row]:
    with connect() as conn:
        return list(
            conn.execute(
                """
                SELECT *
                FROM orcamentos
                ORDER BY id DESC
                LIMIT ?
                """,
                (limit,),
            )
        )


def get_orcamento(orcamento_id: int) -> sqlite3.Row | None:
    with connect() as conn:
        return conn.execute(
            "SELECT * FROM orcamentos WHERE id = ?",
            (orcamento_id,),
        ).fetchone()


def get_itens(orcamento_id: int) -> list[sqlite3.Row]:
    with connect() as conn:
        return list(
            conn.execute(
                """
                SELECT *
                FROM orcamento_itens
                WHERE orcamento_id = ?
                ORDER BY id
                """,
                (orcamento_id,),
            )
        )


def delete_orcamento(orcamento_id: int) -> sqlite3.Row | None:
    with connect() as conn:
        row = conn.execute(
            "SELECT * FROM orcamentos WHERE id = ?",
            (orcamento_id,),
        ).fetchone()
        if row is None:
            return None
        conn.execute("DELETE FROM orcamento_itens WHERE orcamento_id = ?", (orcamento_id,))
        conn.execute("DELETE FROM orcamentos WHERE id = ?", (orcamento_id,))
        return row


def iter_export_rows() -> Iterable[sqlite3.Row]:
    with connect() as conn:
        yield from conn.execute(
            """
            SELECT
                o.id,
                o.criado_em,
                o.cliente_nome,
                o.farmaceutico_responsavel,
                o.cpf,
                o.telefone,
                o.email,
                o.data_orcamento,
                o.localidade,
                o.total AS total_orcamento,
                i.codigo,
                i.produto,
                i.pmc,
                i.desconto,
                i.valor_unitario,
                i.quantidade,
                i.total AS total_item
            FROM orcamentos o
            JOIN orcamento_itens i ON i.orcamento_id = o.id
            ORDER BY o.id DESC, i.id
            """
        )
