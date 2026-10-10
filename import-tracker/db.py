"""
Хранилище данных на чистом SQLite (stdlib, без ORM) — чтобы было легко
читать и дорабатывать под себя.

Структура:
  users      — пользователи (ты и коллеги)
  shipments  — ПОСТАВКИ (у каждой свой номер)
  invoices   — ИНВОЙСЫ (принадлежат ровно одной поставке)
  steps      — шаги процесса по каждому инвойсу (это и есть прогресс)
  files      — загруженные документы (физически лежат в папках на диске)
"""

import os
import sqlite3
import hashlib
import secrets
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
FILES_DIR = DATA_DIR / "files"           # сюда складываются сами документы
DB_PATH = DATA_DIR / "tracker.db"

DATA_DIR.mkdir(exist_ok=True)
FILES_DIR.mkdir(exist_ok=True)


# --- Шаблон процесса по инвойсу -------------------------------------------
# Это НАСТРАИВАЕМЫЙ список. При создании инвойса эти шаги копируются к нему.
# Поменяешь здесь — новые инвойсы будут создаваться с новым списком.
# kind: входящий | создать | отправка | согласование
DEFAULT_STEPS = [
    ("Инвойс", "входящий"),
    ("Пакинг лист", "входящий"),
    ("Перевод", "входящий"),
    ("Камьен", "создать"),
    ("Спецификация", "создать"),
    ("Русификация", "создать"),
    ("Русификация отправлена на согл. (бренд-менеджер РУС)", "отправка"),
    ("РУС согласовано", "согласование"),
    ("Файл товаров и паллет", "входящий"),
    ("Склад", "создать"),
    ("Документы отправлены брокеру", "отправка"),
    ("Национальный каталог", "создать"),
    ("Транзит", "создать"),
]

STATUSES = ["не начато", "в работе", "готово"]

# Короткие подписи этапов для шапки матрицы (параллельно DEFAULT_STEPS)
STAGE_SHORT = [
    "Инвойс", "Пакинг", "Перевод", "Камьен", "Специф.", "Русиф.",
    "Рус→согл", "Рус ОК", "Товары", "Склад", "Брокер", "Нац.кат", "Транзит",
]


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    conn = get_conn()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            username      TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role          TEXT NOT NULL DEFAULT 'user',   -- admin | user
            created_at    TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS shipments (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            number     TEXT UNIQUE NOT NULL,              -- номер поставки
            title      TEXT,
            note       TEXT,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS invoices (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            shipment_id INTEGER NOT NULL REFERENCES shipments(id) ON DELETE CASCADE,
            number      TEXT NOT NULL,                    -- номер инвойса
            supplier    TEXT,
            amount      TEXT,
            note        TEXT,
            created_at  TEXT NOT NULL,
            -- поля, заполняемые автоматически из Excel
            inv_date         TEXT,
            incoterms        TEXT,
            currency         TEXT,
            total_qty        TEXT,
            pallets          TEXT,
            boxes            TEXT,
            net_weight       TEXT,
            gross_weight     TEXT,
            contract         TEXT,
            consignee        TEXT,
            inn              TEXT,
            packing_list_no  TEXT,
            source_file      TEXT,
            UNIQUE (shipment_id, number)
        );

        CREATE TABLE IF NOT EXISTS items (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id  INTEGER NOT NULL REFERENCES invoices(id) ON DELETE CASCADE,
            position    INTEGER,
            hs_code     TEXT,     -- Код ТН ВЭД
            item_code   TEXT,     -- Артикул
            name        TEXT,     -- Наименование (рус.)
            name_en     TEXT,     -- Наименование (англ.)
            barcode     TEXT,
            origin      TEXT,
            qty         TEXT,
            unit_price  TEXT,
            total_price TEXT,
            net_weight  TEXT,
            gross_weight TEXT,
            box_count   TEXT
        );

        CREATE TABLE IF NOT EXISTS steps (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id  INTEGER NOT NULL REFERENCES invoices(id) ON DELETE CASCADE,
            position    INTEGER NOT NULL,
            name        TEXT NOT NULL,
            kind        TEXT NOT NULL,
            status      TEXT NOT NULL DEFAULT 'не начато',
            note        TEXT,
            updated_at  TEXT
        );

        CREATE TABLE IF NOT EXISTS article_ref (
            article    TEXT PRIMARY KEY,   -- нормализованный артикул
            hs_code    TEXT,               -- Код ТН ВЭД
            ean        TEXT,
            name_rus   TEXT,               -- рус. наименование
            origin     TEXT,               -- страна (Origine)
            producer   TEXT,               -- Fournisseur / производитель
            source     TEXT,               -- 'transit' | 'novinki'
            source_name TEXT,              -- 'TRANSIT' или бренд/файл новинок
            supply_key INTEGER DEFAULT 0,
            updated_at TEXT
        );

        CREATE TABLE IF NOT EXISTS files (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id   INTEGER NOT NULL REFERENCES invoices(id) ON DELETE CASCADE,
            step_id      INTEGER REFERENCES steps(id) ON DELETE SET NULL,
            original_name TEXT NOT NULL,
            stored_path  TEXT NOT NULL,
            size         INTEGER,
            uploaded_by  TEXT,
            uploaded_at  TEXT NOT NULL
        );
        """
    )
    # Мягкая миграция: если база создавалась старой версией — добираем новые колонки
    existing = {r["name"] for r in conn.execute("PRAGMA table_info(invoices)")}
    for col in ("inv_date", "incoterms", "currency", "total_qty", "pallets",
                "boxes", "net_weight", "gross_weight", "contract", "consignee",
                "inn", "packing_list_no", "source_file"):
        if col not in existing:
            conn.execute(f"ALTER TABLE invoices ADD COLUMN {col} TEXT")
    existing_items = {r["name"] for r in conn.execute("PRAGMA table_info(items)")}
    if "name_en" not in existing_items:
        conn.execute("ALTER TABLE items ADD COLUMN name_en TEXT")
    existing_ref = {r["name"] for r in conn.execute("PRAGMA table_info(article_ref)")}
    if "source_name" not in existing_ref:
        conn.execute("ALTER TABLE article_ref ADD COLUMN source_name TEXT")
    conn.commit()
    _load_ref_seed_if_empty(conn)
    conn.close()


def _load_ref_seed_if_empty(conn) -> None:
    """Первый запуск: если справочник пуст, заполняем его из встроенного
    файла ref_seed.json.gz (TRANSIT + НОВИНКИ, собранные заранее)."""
    try:
        cnt = conn.execute("SELECT COUNT(*) c FROM article_ref").fetchone()["c"]
    except Exception:
        return
    if cnt:
        return
    seed = BASE_DIR / "ref_seed.json.gz"
    if not seed.exists():
        return
    import json as _json
    import gzip as _gzip
    try:
        with _gzip.open(seed, "rt", encoding="utf-8") as f:
            rows = _json.load(f)
    except Exception:
        return
    conn.executemany(
        """INSERT OR IGNORE INTO article_ref
           (article, hs_code, ean, name_rus, origin, producer, source, source_name, supply_key, updated_at)
           VALUES (?,?,?,?,?,?,?,?,?,?)""",
        [(r.get("article", ""), r.get("hs_code", ""), r.get("ean", ""),
          r.get("name_rus", ""), r.get("origin", ""), r.get("producer", ""),
          r.get("source", "seed"),
          r.get("source_name", "TRANSIT" if r.get("source") == "transit" else "Новинки"),
          r.get("supply_key", 0), now()) for r in rows],
    )
    conn.commit()


# Поля инвойса, которые заполняет парсер Excel
INVOICE_IMPORT_FIELDS = [
    "number", "supplier", "amount", "inv_date", "incoterms", "currency",
    "total_qty", "pallets", "boxes", "net_weight", "gross_weight",
    "contract", "consignee", "inn", "packing_list_no",
]


# --- Пользователи / пароли -------------------------------------------------
def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000)
    return f"{salt}${dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        salt, _ = stored.split("$", 1)
    except ValueError:
        return False
    return secrets.compare_digest(hash_password(password, salt), stored)


def create_user(username: str, password: str, role: str = "user") -> None:
    conn = get_conn()
    conn.execute(
        "INSERT INTO users (username, password_hash, role, created_at) VALUES (?,?,?,?)",
        (username, hash_password(password), role, now()),
    )
    conn.commit()
    conn.close()


def get_user(username: str):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    return row


def list_users():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM users ORDER BY username").fetchall()
    conn.close()
    return rows


def ensure_admin() -> str | None:
    """Создаёт первого админа, если пользователей ещё нет.
    Логин/пароль берутся из ENV или по умолчанию admin/admin."""
    conn = get_conn()
    count = conn.execute("SELECT COUNT(*) AS c FROM users").fetchone()["c"]
    conn.close()
    if count:
        return None
    username = os.getenv("ADMIN_USER", "admin")
    password = os.getenv("ADMIN_PASSWORD", "admin")
    create_user(username, password, role="admin")
    return username


# --- Поставки --------------------------------------------------------------
def list_shipments():
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT s.*,
               (SELECT COUNT(*) FROM invoices i WHERE i.shipment_id = s.id) AS invoice_count
        FROM shipments s ORDER BY s.created_at DESC
        """
    ).fetchall()
    conn.close()
    return rows


def create_shipment(number: str, title: str = "", note: str = "") -> int:
    conn = get_conn()
    cur = conn.execute(
        "INSERT INTO shipments (number, title, note, created_at) VALUES (?,?,?,?)",
        (number.strip(), title.strip(), note.strip(), now()),
    )
    conn.commit()
    sid = cur.lastrowid
    conn.close()
    # сразу создаём папку поставки на диске
    shipment_folder(number)
    return sid


def get_shipment(shipment_id: int):
    conn = get_conn()
    row = conn.execute("SELECT * FROM shipments WHERE id = ?", (shipment_id,)).fetchone()
    conn.close()
    return row


# --- Инвойсы ---------------------------------------------------------------
def list_all_invoices():
    """Все инвойсы по всем поставкам (для общего экрана)."""
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT i.*, s.number AS shipment_number, s.id AS shipment_id,
               (SELECT COUNT(*) FROM files f WHERE f.invoice_id = i.id) AS file_count
        FROM invoices i
        JOIN shipments s ON s.id = i.shipment_id
        ORDER BY s.created_at DESC, i.created_at
        """
    ).fetchall()
    conn.close()
    return rows


def list_invoices(shipment_id: int):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM invoices WHERE shipment_id = ? ORDER BY created_at",
        (shipment_id,),
    ).fetchall()
    conn.close()
    return rows


def create_invoice(shipment_id: int, number: str, supplier: str = "",
                   amount: str = "", note: str = "") -> int:
    conn = get_conn()
    cur = conn.execute(
        """INSERT INTO invoices (shipment_id, number, supplier, amount, note, created_at)
           VALUES (?,?,?,?,?,?)""",
        (shipment_id, number.strip(), supplier.strip(), amount.strip(), note.strip(), now()),
    )
    invoice_id = cur.lastrowid
    # сразу создаём подпапку инвойса внутри папки поставки
    sh = conn.execute("SELECT number FROM shipments WHERE id = ?", (shipment_id,)).fetchone()
    if sh:
        invoice_folder(sh["number"], number)
    # Копируем шаблон шагов именно этому инвойсу
    for pos, (name, kind) in enumerate(DEFAULT_STEPS):
        conn.execute(
            """INSERT INTO steps (invoice_id, position, name, kind, status, updated_at)
               VALUES (?,?,?,?, 'не начато', ?)""",
            (invoice_id, pos, name, kind, now()),
        )
    conn.commit()
    conn.close()
    return invoice_id


def get_invoice(invoice_id: int):
    conn = get_conn()
    row = conn.execute("SELECT * FROM invoices WHERE id = ?", (invoice_id,)).fetchone()
    conn.close()
    return row


# Разрешённые к записи колонки инвойса (чтобы не было SQL-инъекции через имена колонок)
_INVOICE_COLS = {
    "supplier", "amount", "note", "inv_date", "incoterms", "currency",
    "total_qty", "pallets", "boxes", "net_weight", "gross_weight",
    "contract", "consignee", "inn", "packing_list_no", "source_file",
}


def set_invoice_fields(invoice_id: int, fields: dict) -> None:
    """Обновляет только переданные (непустые) поля инвойса."""
    cols = {k: v for k, v in fields.items()
            if k in _INVOICE_COLS and v not in (None, "")}
    if not cols:
        return
    sets = ", ".join(f"{k} = ?" for k in cols)
    conn = get_conn()
    conn.execute(f"UPDATE invoices SET {sets} WHERE id = ?",
                 (*cols.values(), invoice_id))
    conn.commit()
    conn.close()


# --- Товары (строки инвойса) ----------------------------------------------
def get_items(invoice_id: int):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM items WHERE invoice_id = ? ORDER BY position, id",
        (invoice_id,),
    ).fetchall()
    conn.close()
    return rows


def replace_items(invoice_id: int, rows: list[dict]) -> None:
    """Полностью заменяет товары инвойса (при повторном импорте)."""
    conn = get_conn()
    conn.execute("DELETE FROM items WHERE invoice_id = ?", (invoice_id,))
    for pos, it in enumerate(rows):
        conn.execute(
            """INSERT INTO items (invoice_id, position, hs_code, item_code, name, name_en,
                                  barcode, origin, qty, unit_price, total_price,
                                  net_weight, gross_weight, box_count)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (invoice_id, pos, it.get("hs_code", ""), it.get("item_code", ""),
             it.get("name", ""), it.get("name_en", ""), it.get("barcode", ""),
             it.get("origin", ""), it.get("qty", ""), it.get("unit_price", ""),
             it.get("total_price", ""), it.get("net_weight", ""),
             it.get("gross_weight", ""), it.get("box_count", "")),
        )
    conn.commit()
    conn.close()


# --- Шаги / прогресс -------------------------------------------------------
def get_step(step_id: int):
    conn = get_conn()
    row = conn.execute("SELECT * FROM steps WHERE id = ?", (step_id,)).fetchone()
    conn.close()
    return row


def get_steps(invoice_id: int):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM steps WHERE invoice_id = ? ORDER BY position", (invoice_id,)
    ).fetchall()
    conn.close()
    return rows


def update_step(step_id: int, status: str, note: str = "") -> None:
    conn = get_conn()
    row = conn.execute(
        "SELECT invoice_id, position FROM steps WHERE id = ?", (step_id,)
    ).fetchone()
    conn.execute(
        "UPDATE steps SET status = ?, note = ?, updated_at = ? WHERE id = ?",
        (status, note.strip(), now(), step_id),
    )
    # Автопереход: когда этап отмечен "готово", следующий этап по порядку
    # сам становится "в работе" (только если он ещё "не начато").
    if row and status == "готово":
        conn.execute(
            """UPDATE steps SET status = 'в работе', updated_at = ?
               WHERE invoice_id = ? AND position = ? AND status = 'не начато'""",
            (now(), row["invoice_id"], row["position"] + 1),
        )
    conn.commit()
    conn.close()


def progress(invoice_id: int) -> dict:
    """Процент готовности инвойса + текущий шаг (первый невыполненный)."""
    steps = get_steps(invoice_id)
    total = len(steps) or 1
    done = sum(1 for s in steps if s["status"] == "готово")
    current = next((s["name"] for s in steps if s["status"] != "готово"), "Закрыто")
    return {
        "done": done,
        "total": len(steps),
        "percent": round(done * 100 / total),
        "current": current,
    }


# --- Файлы -----------------------------------------------------------------
def list_files(invoice_id: int):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM files WHERE invoice_id = ? ORDER BY uploaded_at DESC",
        (invoice_id,),
    ).fetchall()
    conn.close()
    return rows


def add_file(invoice_id: int, step_id, original_name: str, stored_path: str,
             size: int, uploaded_by: str) -> None:
    conn = get_conn()
    conn.execute(
        """INSERT INTO files (invoice_id, step_id, original_name, stored_path,
                              size, uploaded_by, uploaded_at)
           VALUES (?,?,?,?,?,?,?)""",
        (invoice_id, step_id, original_name, stored_path, size, uploaded_by, now()),
    )
    conn.commit()
    conn.close()


def get_file(file_id: int):
    conn = get_conn()
    row = conn.execute("SELECT * FROM files WHERE id = ?", (file_id,)).fetchone()
    conn.close()
    return row


def set_step_status(step_id: int, status: str) -> None:
    """Меняет только статус шага (заметку не трогает)."""
    conn = get_conn()
    conn.execute("UPDATE steps SET status = ?, updated_at = ? WHERE id = ?",
                 (status, now(), step_id))
    conn.commit()
    conn.close()


def delete_step_files(step_id: int) -> None:
    """Удаляет все файлы, привязанные к шагу (и с диска)."""
    conn = get_conn()
    rows = conn.execute("SELECT stored_path FROM files WHERE step_id = ?", (step_id,)).fetchall()
    for r in rows:
        try:
            os.remove(r["stored_path"])
        except OSError:
            pass
    conn.execute("DELETE FROM files WHERE step_id = ?", (step_id,))
    conn.commit()
    conn.close()


def delete_file(file_id: int) -> None:
    conn = get_conn()
    row = conn.execute("SELECT stored_path FROM files WHERE id = ?", (file_id,)).fetchone()
    if row:
        try:
            os.remove(row["stored_path"])
        except OSError:
            pass
    conn.execute("DELETE FROM files WHERE id = ?", (file_id,))
    conn.commit()
    conn.close()


def _safe(name: str) -> str:
    """Безопасное имя папки/файла (без запрещённых символов)."""
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in str(name)).strip("_") or "x"


def shipment_folder(shipment_number: str) -> Path:
    """Папка поставки на диске (создаётся при необходимости)."""
    folder = FILES_DIR / _safe(shipment_number)
    folder.mkdir(parents=True, exist_ok=True)
    return folder


# --- Справочник артикулов (TRANSIT / НОВИНКИ) ------------------------------
def ref_upsert_transit(rows: list[dict]) -> int:
    """TRANSIT имеет приоритет — перезаписывает запись по артикулу."""
    conn = get_conn()
    conn.executemany(
        """INSERT OR REPLACE INTO article_ref
           (article, hs_code, ean, name_rus, origin, producer, source, source_name, supply_key, updated_at)
           VALUES (?,?,?,?,?,?, 'transit', 'TRANSIT', ?, ?)""",
        [(r["article"], r.get("hs_code", ""), r.get("ean", ""), r.get("name_rus", ""),
          r.get("origin", ""), r.get("producer", ""), r.get("supply_key", 0), now())
         for r in rows],
    )
    conn.commit()
    conn.close()
    return len(rows)


def ref_upsert_novinki(rows: list[dict], source_name: str = "Новинки") -> int:
    """НОВИНКИ добавляют только артикулы, которых ещё нет (TRANSIT не трогаем)."""
    conn = get_conn()
    conn.executemany(
        """INSERT OR IGNORE INTO article_ref
           (article, hs_code, ean, name_rus, origin, producer, source, source_name, supply_key, updated_at)
           VALUES (?,?,?,?,?,?, 'novinki', ?, 0, ?)""",
        [(r["article"], r.get("hs_code", ""), r.get("ean", ""), r.get("name_rus", ""),
          r.get("origin", ""), r.get("producer", ""), r.get("source_name", source_name), now())
         for r in rows],
    )
    conn.commit()
    conn.close()
    return len(rows)


def ref_brands() -> dict:
    """Содержимое справочника: TRANSIT (кол-во) + список брендов Новинок (кол-во)."""
    conn = get_conn()
    tr = conn.execute("SELECT COUNT(*) c FROM article_ref WHERE source='transit'").fetchone()["c"]
    nv = conn.execute(
        """SELECT COALESCE(source_name,'Новинки') AS name, COUNT(*) c
           FROM article_ref WHERE source='novinki'
           GROUP BY COALESCE(source_name,'Новинки') ORDER BY name"""
    ).fetchall()
    conn.close()
    return {"transit": tr, "novinki": [dict(r) for r in nv]}


def ref_get(article: str):
    conn = get_conn()
    row = conn.execute("SELECT * FROM article_ref WHERE article = ?", (article,)).fetchone()
    conn.close()
    return row


def ref_get_many(articles: list[str]) -> dict:
    """Пакетный поиск: {article: row} для переданных артикулов."""
    if not articles:
        return {}
    conn = get_conn()
    out = {}
    uniq = list({a for a in articles if a})
    for i in range(0, len(uniq), 400):
        chunk = uniq[i:i + 400]
        q = "SELECT * FROM article_ref WHERE article IN (%s)" % ",".join("?" * len(chunk))
        for r in conn.execute(q, chunk):
            out[r["article"]] = r
    conn.close()
    return out


def ref_search(q: str = "", source: str = "", brand: str = "", limit: int = 50, offset: int = 0):
    """Поиск/просмотр справочника с фильтром по источнику/бренду и пагинацией."""
    where, params = [], []
    if source in ("transit", "novinki"):
        where.append("source = ?")
        params.append(source)
    if brand:
        where.append("COALESCE(source_name,'Новинки') = ?")
        params.append(brand)
    if q:
        where.append("(article LIKE ? OR name_rus LIKE ? OR hs_code LIKE ? OR ean LIKE ?)")
        qa = f"%{q.upper()}%"
        qp = f"%{q}%"
        params += [qa, qp, qp, qp]
    wsql = (" WHERE " + " AND ".join(where)) if where else ""
    conn = get_conn()
    total = conn.execute("SELECT COUNT(*) c FROM article_ref" + wsql, params).fetchone()["c"]
    rows = conn.execute(
        "SELECT * FROM article_ref" + wsql + " ORDER BY article LIMIT ? OFFSET ?",
        params + [limit, offset],
    ).fetchall()
    conn.close()
    return rows, total


def ref_stats() -> dict:
    conn = get_conn()
    total = conn.execute("SELECT COUNT(*) c FROM article_ref").fetchone()["c"]
    tr = conn.execute("SELECT COUNT(*) c FROM article_ref WHERE source='transit'").fetchone()["c"]
    nv = conn.execute("SELECT COUNT(*) c FROM article_ref WHERE source='novinki'").fetchone()["c"]
    last = conn.execute("SELECT MAX(updated_at) m FROM article_ref").fetchone()["m"]
    conn.close()
    return {"total": total, "transit": tr, "novinki": nv, "updated": last}


def invoice_folder(shipment_number: str, invoice_number: str) -> Path:
    """Папка инвойса (подпапка внутри папки поставки)."""
    folder = FILES_DIR / _safe(shipment_number) / _safe(invoice_number)
    folder.mkdir(parents=True, exist_ok=True)
    return folder
