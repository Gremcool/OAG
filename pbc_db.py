"""
pbc_db.py
---------
Storage for the OAG Preliminary Information Request tracker (standalone SQLite database).

Tables: pbc_items, pbc_documents (attachments stored as BLOBs), pbc_comments,
pbc_logs and pbc_meta. Everything lives in one file, so backing up that single
file backs up the items, documents and comments together.
"""
import hashlib
import json
import os
from contextlib import contextmanager
from datetime import datetime

import pandas as pd

import sqlite3

# Standalone database. Set DATABASE_PATH (e.g. a Railway volume path) to keep data across deploys.
DB_FILE = os.getenv("DATABASE_PATH", "pbc_tracker.db")
_db_dir = os.path.dirname(DB_FILE)
if _db_dir and not os.path.exists(_db_dir):
    os.makedirs(_db_dir, exist_ok=True)


def get_conn():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

STATUSES = ["Pending", "In Progress", "Partially Submitted", "Submitted", "Not Applicable"]
DONE_STATUSES = ("Submitted", "Not Applicable")
MAX_FILE_MB = 50


@contextmanager
def _conn():
    conn = get_conn()
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------------------
# Schema + seeding
# ---------------------------------------------------------------------------
def init():
    with _conn() as conn:
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS pbc_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sort_order REAL,
                ref TEXT,
                section TEXT,
                subsection TEXT,
                description TEXT,
                info_type TEXT,
                directorate TEXT,
                responsible TEXT,
                deadline TEXT,
                status TEXT DEFAULT 'Pending',
                notes TEXT,
                created_at TEXT,
                updated_at TEXT
            )""")
        c.execute("""
            CREATE TABLE IF NOT EXISTS pbc_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                filename TEXT,
                mime TEXT,
                size INTEGER,
                sha256 TEXT,
                note TEXT,
                data BLOB,
                uploaded_at TEXT
            )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_pbc_docs_item ON pbc_documents(item_id)")
        c.execute("""
            CREATE TABLE IF NOT EXISTS pbc_comments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                author TEXT,
                body TEXT,
                created_at TEXT
            )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_pbc_comments_item ON pbc_comments(item_id)")
        c.execute("""
            CREATE TABLE IF NOT EXISTS pbc_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                message TEXT
            )""")
        c.execute("CREATE TABLE IF NOT EXISTS pbc_meta (key TEXT PRIMARY KEY, value TEXT)")


def get_meta(key, default=None):
    with _conn() as conn:
        r = conn.execute("SELECT value FROM pbc_meta WHERE key=?", (key,)).fetchone()
    return r[0] if r else default


def set_meta(key, value):
    with _conn() as conn:
        conn.execute("INSERT OR REPLACE INTO pbc_meta(key, value) VALUES (?, ?)", (key, value))


def get_request_meta():
    try:
        return json.loads(get_meta("request_meta", "{}") or "{}")
    except Exception:
        return {}


def seed_if_needed(seed_path):
    """First run only: load the bundled request list (parsed from the OAG Word file)."""
    if get_meta("seeded") == "1":
        return 0
    with _conn() as conn:
        n = conn.execute("SELECT COUNT(*) FROM pbc_items").fetchone()[0]
    if n > 0 or not os.path.exists(seed_path):
        set_meta("seeded", "1")
        return 0
    with open(seed_path, encoding="utf-8") as f:
        seed = json.load(f)
    added = import_items(seed["items"])
    set_meta("request_meta", json.dumps(seed.get("meta", {})))
    set_meta("seeded", "1")
    log(f"📥 Loaded {added} request items from the OAG Preliminary Information Request")
    return added


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
def log(message):
    with _conn() as conn:
        conn.execute("INSERT INTO pbc_logs(timestamp, message) VALUES (?, ?)", (_now(), message))


def get_logs(limit=200):
    with _conn() as conn:
        return pd.read_sql(
            "SELECT timestamp, message FROM pbc_logs ORDER BY id DESC LIMIT ?", conn, params=(limit,))


# ---------------------------------------------------------------------------
# Items
# ---------------------------------------------------------------------------
ITEM_COLS = ["ref", "section", "subsection", "description", "info_type",
             "directorate", "responsible", "deadline", "status", "notes"]


def load_items():
    with _conn() as conn:
        df = pd.read_sql("""
            SELECT i.*,
                   COALESCE(d.n, 0) AS doc_count,
                   COALESCE(d.bytes, 0) AS doc_bytes
            FROM pbc_items i
            LEFT JOIN (SELECT item_id, COUNT(*) n, SUM(size) bytes
                       FROM pbc_documents GROUP BY item_id) d ON d.item_id = i.id
            ORDER BY i.sort_order, i.id""", conn)
    for col in ITEM_COLS:
        if col in df.columns:
            df[col] = df[col].fillna("")
    return df


def get_item(item_id):
    with _conn() as conn:
        conn.row_factory = None
        cur = conn.execute("SELECT * FROM pbc_items WHERE id=?", (item_id,))
        row = cur.fetchone()
        if not row:
            return None
        cols = [d[0] for d in cur.description]
    item = dict(zip(cols, row))
    return {k: ("" if v is None else v) for k, v in item.items()}


def _next_sort_order(conn, section):
    """Place a new item at the end of its section (or at the very end for a new section)."""
    row = conn.execute("SELECT MAX(sort_order) FROM pbc_items WHERE section=?", (section,)).fetchone()
    if row and row[0] is not None:
        a = row[0]
        nxt = conn.execute("SELECT MIN(sort_order) FROM pbc_items WHERE sort_order > ?", (a,)).fetchone()
        return (a + nxt[0]) / 2 if nxt and nxt[0] is not None else a + 10
    row = conn.execute("SELECT MAX(sort_order) FROM pbc_items").fetchone()
    return (row[0] or 0) + 10


def save_item(item, item_id=None):
    """Insert (item_id None) or update. Returns the item id."""
    now = _now()
    with _conn() as conn:
        c = conn.cursor()
        if item_id is None:
            order = item.get("sort_order") or _next_sort_order(conn, item.get("section", ""))
            cols = ["sort_order"] + ITEM_COLS + ["created_at", "updated_at"]
            vals = [order] + [item.get(k, "") for k in ITEM_COLS] + [now, now]
            c.execute(f"INSERT INTO pbc_items ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})", vals)
            return c.lastrowid
        sets = ", ".join(f"{k}=?" for k in ITEM_COLS) + ", updated_at=?"
        vals = [item.get(k, "") for k in ITEM_COLS] + [now, item_id]
        c.execute(f"UPDATE pbc_items SET {sets} WHERE id=?", vals)
        return item_id


def set_status(item_id, status):
    with _conn() as conn:
        conn.execute("UPDATE pbc_items SET status=?, updated_at=? WHERE id=?", (status, _now(), item_id))


def delete_item(item_id):
    with _conn() as conn:
        conn.execute("DELETE FROM pbc_documents WHERE item_id=?", (item_id,))
        conn.execute("DELETE FROM pbc_comments WHERE item_id=?", (item_id,))
        conn.execute("DELETE FROM pbc_items WHERE id=?", (item_id,))


def _dedupe_key(ref, description_text):
    return ((ref or "").strip().lower(), " ".join((description_text or "").split())[:80].lower())


def _respace(conn):
    """Keep one section together: new items join the END of their existing section, new sections go last.
    Relative order of everything already in the tracker is preserved."""
    rows = conn.execute("SELECT id, section, sort_order FROM pbc_items ORDER BY sort_order, id").fetchall()
    rank = {}
    for _, sec, _ in rows:
        rank.setdefault(sec, len(rank))
    for i, (iid, _, _) in enumerate(sorted(rows, key=lambda r: (rank[r[1]], r[2], r[0])), 1):
        conn.execute("UPDATE pbc_items SET sort_order=? WHERE id=?", (i * 10, iid))


def import_items(items, dry_run=False, keep_duplicates=False):
    """ADD-ONLY import. Existing items, statuses, notes, documents and comments are never changed or deleted.

    * Items that look identical to one already present (same ref + start of wording) are skipped,
      unless keep_duplicates=True, in which case they are added as separate copies.
    * New items are placed at the end of their section (a new section is added at the end).
    * With dry_run=True nothing is written and the number of items that WOULD be added is returned.
    """
    from pbc_parser import html_to_text  # local import keeps module import light

    with _conn() as conn:
        existing = {
            _dedupe_key(r[0], html_to_text(r[1]))
            for r in conn.execute("SELECT ref, description FROM pbc_items").fetchall()
        }
        base = conn.execute("SELECT MAX(sort_order) FROM pbc_items").fetchone()[0] or 0
        added = 0
        now = _now()
        for it in items:
            key = _dedupe_key(it.get("ref"), html_to_text(it.get("description", "")))
            if not keep_duplicates and key in existing:
                continue
            existing.add(key)
            added += 1
            if dry_run:
                continue
            cols = ["sort_order"] + ITEM_COLS + ["created_at", "updated_at"]
            vals = [base + added] + [it.get(k, "") for k in ITEM_COLS] + [now, now]
            conn.execute(f"INSERT INTO pbc_items ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})", vals)
        if added and not dry_run:
            _respace(conn)
    return added


def totals():
    """Headline counts, shown on the Import tab so you can see nothing is lost."""
    with _conn() as conn:
        return {
            "items": conn.execute("SELECT COUNT(*) FROM pbc_items").fetchone()[0],
            "documents": conn.execute("SELECT COUNT(*) FROM pbc_documents").fetchone()[0],
            "comments": conn.execute("SELECT COUNT(*) FROM pbc_comments").fetchone()[0],
        }


# ---------------------------------------------------------------------------
# Documents (attachments)
# ---------------------------------------------------------------------------
def add_document(item_id, filename, data: bytes, mime="", note=""):
    """Returns (ok, message)."""
    if len(data) > MAX_FILE_MB * 1024 * 1024:
        return False, f"'{filename}' is larger than {MAX_FILE_MB} MB."
    digest = hashlib.sha256(data).hexdigest()
    with _conn() as conn:
        dup = conn.execute(
            "SELECT filename FROM pbc_documents WHERE item_id=? AND sha256=?", (item_id, digest)).fetchone()
        if dup:
            return False, f"'{filename}' is already attached to this item (same content as '{dup[0]}')."
        conn.execute(
            "INSERT INTO pbc_documents(item_id, filename, mime, size, sha256, note, data, uploaded_at) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (item_id, filename, mime or "application/octet-stream", len(data), digest, note, data, _now()))
    return True, f"'{filename}' attached."


def list_documents(item_id):
    """Metadata only (no blobs)."""
    with _conn() as conn:
        rows = conn.execute(
            "SELECT id, filename, mime, size, note, uploaded_at FROM pbc_documents "
            "WHERE item_id=? ORDER BY id", (item_id,)).fetchall()
    return [dict(id=r[0], filename=r[1], mime=r[2], size=r[3], note=r[4] or "", uploaded_at=r[5]) for r in rows]


def get_document_data(doc_id):
    with _conn() as conn:
        r = conn.execute("SELECT data FROM pbc_documents WHERE id=?", (doc_id,)).fetchone()
    return bytes(r[0]) if r and r[0] is not None else b""


def delete_document(doc_id):
    with _conn() as conn:
        conn.execute("DELETE FROM pbc_documents WHERE id=?", (doc_id,))


def update_document_note(doc_id, note):
    with _conn() as conn:
        conn.execute("UPDATE pbc_documents SET note=? WHERE id=?", (note, doc_id))


def docs_summary():
    """{item_id: [(filename, size), ...]} for showing file chips on the cards."""
    out = {}
    with _conn() as conn:
        for item_id, name, size in conn.execute(
                "SELECT item_id, filename, size FROM pbc_documents ORDER BY id").fetchall():
            out.setdefault(item_id, []).append((name, size))
    return out


def iter_documents_for_zip():
    """Yields (ref, item_id, filename, bytes) one file at a time to keep memory low."""
    with _conn() as conn:
        ids = conn.execute(
            "SELECT d.id, i.ref, i.id, d.filename FROM pbc_documents d "
            "JOIN pbc_items i ON i.id = d.item_id ORDER BY i.sort_order, d.id").fetchall()
    for doc_id, ref, item_id, filename in ids:
        yield ref or "no-ref", item_id, filename, get_document_data(doc_id)


# ---------------------------------------------------------------------------
# Comments (a discussion thread on every request item)
# ---------------------------------------------------------------------------
def add_comment(item_id, author, body):
    body = (body or "").strip()
    if not body:
        return False
    with _conn() as conn:
        conn.execute("INSERT INTO pbc_comments(item_id, author, body, created_at) VALUES (?,?,?,?)",
                     (item_id, (author or "").strip() or "Anonymous", body, _now()))
    return True


def list_comments(item_id):
    with _conn() as conn:
        rows = conn.execute("SELECT id, author, body, created_at FROM pbc_comments WHERE item_id=? ORDER BY id",
                            (item_id,)).fetchall()
    return [dict(id=r[0], author=r[1], body=r[2], created_at=r[3]) for r in rows]


def delete_comment(comment_id):
    with _conn() as conn:
        conn.execute("DELETE FROM pbc_comments WHERE id=?", (comment_id,))


def comments_all():
    """{item_id: [(author, body, created_at), ...]} oldest first - for card previews and exports."""
    out = {}
    with _conn() as conn:
        for item_id, author, body, ts in conn.execute(
                "SELECT item_id, author, body, created_at FROM pbc_comments ORDER BY id").fetchall():
            out.setdefault(item_id, []).append((author, body, ts))
    return out


# ---------------------------------------------------------------------------
# Editable lists (rename / merge a dropdown value across every request)
# ---------------------------------------------------------------------------
LIST_FIELDS = {
    "directorate": "Directorate",
    "responsible": "Responsible person",
    "section": "Section",
    "subsection": "Sub-section",
}


def list_values(field):
    """[(value, number_of_items), ...] for a list field, A-Z. Blank values are skipped."""
    if field not in LIST_FIELDS:
        raise ValueError("unknown list")
    with _conn() as conn:
        rows = conn.execute(
            f"SELECT {field}, COUNT(*) FROM pbc_items WHERE TRIM(COALESCE({field}, '')) <> '' "
            f"GROUP BY {field}").fetchall()
    return sorted(rows, key=lambda r: r[0].lower())


def rename_value(field, old, new):
    """Rename `old` to `new` on every item. If `new` already exists the two are merged.
    Returns the number of items changed."""
    if field not in LIST_FIELDS:
        raise ValueError("unknown list")
    new = " ".join((new or "").split())
    if not new or new == old:
        return 0
    with _conn() as conn:
        cur = conn.execute(f"UPDATE pbc_items SET {field}=?, updated_at=? WHERE {field}=?", (new, _now(), old))
        n = cur.rowcount
    log(f"✏️ Renamed {LIST_FIELDS[field].lower()} '{old}' → '{new}' ({n} item{'s' if n != 1 else ''})")
    return n
