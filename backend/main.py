"""Exam portal — FastAPI backend (manually graded, capture-only).

Semua jawaban DINILAI MANUAL oleh pengawas. Portal hanya menangkap jawaban:
- Q tipe "mcq"  -> radio, jawaban = index pilihan
- Q tipe "text" -> textarea, jawaban = teks
- Q tipe "group"-> kartu berisi beberapa sub (a/b/c/d), tiap sub = slot sendiri
TIDAK ada auto-grade, TIDAK ada skor, TIDAK ada leaderboard. Admin melihat
seluruh jawaban per attempt dan meng-export Excel untuk dinilai di luar.

Storage: SQLite at data/quiz.db (override with EXAM_DB). Attempt lifecycle: /api/start creates a
row (status='started', started_at), answers autosave via /api/save while the
student works, /api/submit flips it to status='submitted' + submitted_at. A
refresh resumes the same session (server-anchored timer, no fresh 90 min).
Multi-attempt = separate rows per NIM (the invigilator checks manually).
"""

import hmac
import os
import io
import json
import re
import sqlite3
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path

from fastapi import FastAPI, Form, Header, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from questions import QUESTIONS, TOTAL_POINTS

BASE_DIR = Path(__file__).resolve().parent
FRONTEND = BASE_DIR.parent / "frontend"
DATA = BASE_DIR.parent / "data"
DATA.mkdir(parents=True, exist_ok=True)
# Override for tests/demos: export EXAM_DB=/tmp/demo.db
DB_PATH = Path(os.environ.get("EXAM_DB", str(DATA / "quiz.db")))

# Admin key comes from the environment — never commit a real one.
#   export EXAM_ADMIN_KEY=$(python3 -c "import secrets;print(secrets.token_hex(16))")
ADMIN_KEY = os.environ.get("EXAM_ADMIN_KEY", "")

LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

EXAM_TITLE = os.environ.get("EXAM_TITLE", "Exam Portal")
app = FastAPI(title=EXAM_TITLE)
app.mount("/static", StaticFiles(directory=str(FRONTEND)), name="static")


# ------------------------------------------------------------------ slot model
def slots():
    """Flat list of answer slots across all questions.

    A group sub WITHOUT inputs -> one slot ("10a"). A sub WITH inputs -> one slot
    PER input, id "{qid}{subkey}_{inkey}" ("13a_z1"), points carried at the sub
    level (each input slot points=0; question total comes from QUESTIONS).
    """
    out = []
    for q in QUESTIONS:
        if q["type"] == "group":
            for s in q["subs"]:
                subs_id = f"{q['id']}{s['key']}"
                if s.get("inputs"):
                    for inp in s["inputs"]:
                        out.append({
                            "id": f"{subs_id}_{inp['key']}",
                            "qid": q["id"],
                            "sub_key": s["key"],
                            "section": q["section"],
                            "topic": q["topic"],
                            "points": 0,
                            "sub_points": s["points"],
                            "kind": inp.get("kind", "text"),
                            "question": q["question"],
                            "sub_question": s["question"],
                            "inkey": inp["key"],
                            "inlabel": inp.get("label", inp["key"]),
                            "placeholder": inp.get("placeholder"),
                            "options": None,
                        })
                else:
                    out.append({
                        "id": subs_id,
                        "qid": q["id"],
                        "sub_key": s["key"],
                        "section": q["section"],
                        "topic": q["topic"],
                        "points": s["points"],
                        "sub_points": s["points"],
                        "kind": s.get("kind", "text"),
                        "question": q["question"],
                        "sub_question": s["question"],
                        "inkey": None,
                        "inlabel": None,
                        "placeholder": s.get("placeholder"),
                        "options": None,
                    })
        else:
            out.append({
                "id": str(q["id"]),
                "qid": q["id"],
                "sub_key": None,
                "section": q["section"],
                "topic": q["topic"],
                "points": q["points"],
                "sub_points": q["points"],
                "kind": q["type"],  # mcq | text
                "question": q["question"],
                "sub_question": None,
                "inkey": None,
                "inlabel": None,
                "placeholder": q.get("placeholder"),
                "options": q.get("options"),
            })
    return out


def question_points(qid):
    """Total points for a question (from QUESTIONS, not slot sums)."""
    for q in QUESTIONS:
        if q["id"] == qid:
            return q["points"]
    return 0


SLOTS = slots()
SLOT_IDS = [s["id"] for s in SLOTS]


# ----------------------------------------------------------------------- db
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS attempts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nim TEXT NOT NULL,
                name TEXT NOT NULL,
                answers TEXT NOT NULL,
                submitted_at REAL NOT NULL,
                attempt_no INTEGER,
                status TEXT NOT NULL DEFAULT 'submitted',
                started_at REAL
            )
            """
        )
        conn.execute("DROP INDEX IF EXISTS idx_attempts_nim")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_attempts_nim ON attempts(nim)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_attempts_name ON attempts(name)")
        # migrate older tables (pre-session build): add status + started_at if absent
        cols = [c[1] for c in conn.execute("PRAGMA table_info(attempts)").fetchall()]
        if "status" not in cols:
            conn.execute("ALTER TABLE attempts ADD COLUMN status TEXT NOT NULL DEFAULT 'submitted'")
        if "started_at" not in cols:
            conn.execute("ALTER TABLE attempts ADD COLUMN started_at REAL")
        # legacy rows (pre-session build): status default submitted, started_at backfilled
        conn.execute(
            "UPDATE attempts SET status = 'submitted', started_at = COALESCE(started_at, submitted_at) "
            "WHERE status IS NULL OR started_at IS NULL"
        )


def validate_student(nim: str, name: str):
    nim = (nim or "").strip()
    name = (name or "").strip()
    if not nim:
        raise HTTPException(status_code=400, detail="Masukkan NIM Anda.")
    if len(nim) != 10 or not nim.isdigit() or not (20 <= int(nim[:2]) <= 30):
        raise HTTPException(status_code=400, detail="NIM harus 10 digit diawali tahun masuk 20-30.")
    if not name:
        raise HTTPException(status_code=400, detail="Masukkan nama lengkap Anda.")
    if len(name) > 100:
        raise HTTPException(status_code=400, detail="Nama terlalu panjang (maks 100 karakter).")
    return nim, name


DURATION_SECONDS = 90 * 60


init_db()


def check_admin(key: str | None):
    if not ADMIN_KEY:
        # No key configured: admin surface is closed, not open.
        raise HTTPException(status_code=503, detail="EXAM_ADMIN_KEY is not set on the server")
    if not key or not hmac.compare_digest(key, ADMIN_KEY):
        raise HTTPException(status_code=401, detail="Invalid admin key")


def _norm(v):
    if isinstance(v, str):
        return v.strip()
    return v


def _fmt_ts(ts):
    t = datetime.fromtimestamp(ts, tz=timezone(timedelta(hours=7)))
    return t.strftime("%Y-%m-%d %H:%M")


def _given_display(slot, answers):
    """Human-readable student answer for a slot (for admin + exports)."""
    given = answers.get(slot["id"])
    if given is None:
        return ""
    if slot["kind"] == "mcq":
        if isinstance(given, int) and 0 <= given < len(slot["options"]):
            return f"{LETTERS[given]}. {slot['options'][given]}"
        if isinstance(given, str) and given.isdigit() and 0 <= int(given) < len(slot["options"]):
            i = int(given)
            return f"{LETTERS[i]}. {slot['options'][i]}"
        return str(given)
    return str(given)


# ------------------------------------------------------------------- routes
@app.get("/")
def index():
    return FileResponse(FRONTEND / "index.html")


@app.get("/admin")
def admin_page():
    # Serve login page WITHOUT header check; security lives on the API endpoints.
    return FileResponse(FRONTEND / "admin.html")


@app.get("/api/questions")
def get_questions():
    """Full exam structure for the student UI (cards + slots, no answers)."""
    return {
        "title": EXAM_TITLE,
        "duration_minutes": 90,
        "total_points": TOTAL_POINTS,
        "questions": QUESTIONS,
    }


@app.post("/api/start")
async def start_session(nim: str = Form(...), name: str = Form(...)):
    """Log an exam start. Reuses an open (started, not submitted) session for the
    same NIM so a refresh never creates a duplicate start row."""
    nim, name = validate_student(nim, name)
    with db() as conn:
        row = conn.execute(
            "SELECT id, nim, name, attempt_no, started_at FROM attempts "
            "WHERE lower(nim)=lower(?) AND status='started' ORDER BY id DESC LIMIT 1",
            (nim,),
        ).fetchone()
        if row:
            return {
                "attempt_id": row["id"],
                "nim": row["nim"],
                "name": row["name"],
                "attempt_no": row["attempt_no"],
                "started_at": row["started_at"],
                "duration_seconds": DURATION_SECONDS,
                "server_now": time.time(),
                "resumed": True,
            }
        conn.execute("BEGIN IMMEDIATE")
        cnt = conn.execute(
            "SELECT COUNT(*) FROM attempts WHERE lower(nim)=lower(?)", (nim,)
        ).fetchone()[0]
        attempt_no = cnt + 1
        now = time.time()
        cur = conn.execute(
            "INSERT INTO attempts (nim, name, answers, submitted_at, attempt_no, status, started_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (nim, name, "{}", now, attempt_no, "started", now),
        )
        return {
            "attempt_id": cur.lastrowid,
            "nim": nim,
            "name": name,
            "attempt_no": attempt_no,
            "started_at": now,
            "duration_seconds": DURATION_SECONDS,
            "server_now": now,
            "resumed": False,
        }


@app.get("/api/session")
def get_session(attempt_id: int):
    """Resume payload for one attempt: current answers + time left (server-anchored)."""
    with db() as conn:
        row = conn.execute(
            "SELECT id, nim, name, answers, attempt_no, status, started_at, submitted_at "
            "FROM attempts WHERE id=?", (attempt_id,),
        ).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Sesi tidak ditemukan.")
    try:
        answers = json.loads(row["answers"] or "{}")
    except Exception:
        answers = {}
    now = time.time()
    if row["status"] == "submitted":
        remaining = 0
    else:
        remaining = max(0, int((row["started_at"] + DURATION_SECONDS) - now))
    return {
        "attempt_id": row["id"],
        "nim": row["nim"],
        "name": row["name"],
        "attempt_no": row["attempt_no"],
        "status": row["status"],
        "started_at": row["started_at"],
        "submitted_at": row["submitted_at"],
        "duration_seconds": DURATION_SECONDS,
        "server_now": now,
        "remaining_seconds": remaining,
        "answers": answers,
    }


@app.post("/api/save")
async def save_progress(attempt_id: int = Form(...), answers: str = Form("{}")):
    """Autosave partial answers onto an open session (no-op after submit)."""
    try:
        answers = json.loads(answers)
    except Exception:
        raise HTTPException(status_code=400, detail="Format jawaban tidak valid.")
    if not isinstance(answers, dict):
        raise HTTPException(status_code=400, detail="Format jawaban tidak valid.")
    clean = {}
    for sid in SLOT_IDS:
        if sid in answers:
            v = _norm(answers[sid])
            if v is not None and v != "":
                clean[sid] = v
    with db() as conn:
        conn.execute(
            "UPDATE attempts SET answers=? WHERE id=? AND status='started'",
            (json.dumps(clean, ensure_ascii=False), attempt_id),
        )
    return {"ok": True, "saved": len(clean)}


@app.post("/api/submit")
async def submit(
    nim: str = Form(...),
    name: str = Form(...),
    answers: str = Form("{}"),
    attempt_id: int | None = Form(default=None),
    coretan: UploadFile | None = File(default=None),
):
    """Finalize the started session (or a legacy direct submit). No grading."""
    nim, name = validate_student(nim, name)
    try:
        answers = json.loads(answers)
    except Exception:
        raise HTTPException(status_code=400, detail="Format jawaban tidak valid.")
    if not isinstance(answers, dict):
        raise HTTPException(status_code=400, detail="Format jawaban tidak valid.")

    clean = {}
    for sid in SLOT_IDS:
        if sid in answers:
            v = _norm(answers[sid])
            if v is not None and v != "":
                clean[sid] = v

    with db() as conn:
        # Prefer flipping the started row we created at /api/start.
        if attempt_id:
            row = conn.execute(
                "SELECT id, attempt_no FROM attempts WHERE id=? AND lower(nim)=lower(?) "
                "AND lower(name)=lower(?) AND status='started'",
                (attempt_id, nim, name),
            ).fetchone()
            if row:
                conn.execute(
                    "UPDATE attempts SET answers=?, status='submitted', submitted_at=? "
                    "WHERE id=?",
                    (json.dumps(clean, ensure_ascii=False), time.time(), row["id"]),
                )
                return {"ok": True, "attempt_no": row["attempt_no"], "submitted": True,
                        "received_slots": len(clean)}
        # fallback: plain insert (legacy path)
        conn.execute("BEGIN IMMEDIATE")
        cnt = conn.execute(
            "SELECT COUNT(*) FROM attempts WHERE lower(nim)=lower(?)", (nim,)
        ).fetchone()[0]
        attempt_no = cnt + 1
        now = time.time()
        conn.execute(
            "INSERT INTO attempts (nim, name, answers, submitted_at, attempt_no, status, started_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (nim, name, json.dumps(clean, ensure_ascii=False), now, attempt_no, "submitted", now),
        )
    return {"ok": True, "attempt_no": attempt_no, "received_slots": len(clean)}


@app.get("/api/scoreboard")
def scoreboard(x_admin_key: str | None = Header(default=None)):
    """Admin review: all attempts with every answer mapped to slots. No score."""
    check_admin(x_admin_key)
    with db() as conn:
        rows = conn.execute(
            "SELECT id, nim, name, answers, attempt_no, status, started_at, submitted_at FROM attempts "
            "ORDER BY lower(nim), COALESCE(attempt_no, 1), COALESCE(started_at, submitted_at)"
        ).fetchall()

    groups = {}
    for r in rows:
        groups.setdefault(r["nim"].lower(), []).append(r)

    out = []
    for nim_lower, attempts in groups.items():
        for r in attempts:
            try:
                answers = json.loads(r["answers"] or "{}")
            except Exception:
                answers = {}
            details = []
            for slot in SLOTS:
                details.append({
                    "id": slot["id"],
                    "section": slot["section"],
                    "topic": slot["topic"],
                    "points": slot["points"],
                    "sub_points": slot["sub_points"],
                    "q_points": question_points(slot["qid"]),
                    "qid": slot["qid"],
                    "sub_key": slot["sub_key"],
                    "inkey": slot["inkey"],
                    "inlabel": slot["inlabel"],
                    "given": _given_display(slot, answers),
                    "raw": answers.get(slot["id"]),
                })
            out.append({
                "id": r["id"],
                "nim": r["nim"],
                "name": r["name"],
                "attempt_no": r["attempt_no"] or 1,
                "total_attempts": len(attempts),
                "status": r["status"],
                "started_at": r["started_at"] or r["submitted_at"],
                "submitted_at": r["submitted_at"] if r["status"] == "submitted" else None,
                "details": details,
            })
    return out


@app.get("/api/export_xlsx")
def export_xlsx(x_admin_key: str | None = Header(default=None)):
    """One row per attempt; one column per answer slot + poin header; blank grade col."""
    check_admin(x_admin_key)
    wb = Workbook()
    ws = wb.active
    ws.title = "Attempts"
    header = ["NIM", "Name", "Attempt#", "Status", "Started (UTC+7)", "Submitted (UTC+7)"]
    for s in SLOTS:
        if s["inkey"]:
            header.append(f"Q{s['qid']}{s['sub_key']} · {s['inlabel']}")
        else:
            header.append(f"Q{s['id']} ({s['sub_points']} poin)")
    header += ["Catatan pengawas"]
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    with db() as conn:
        rows = conn.execute(
            "SELECT nim, name, answers, attempt_no, status, started_at, submitted_at FROM attempts "
            "ORDER BY lower(nim), COALESCE(attempt_no, 1), COALESCE(started_at, submitted_at)"
        ).fetchall()
    for r in rows:
        try:
            answers = json.loads(r["answers"] or "{}")
        except Exception:
            answers = {}
        st = r["status"] or "submitted"
        row = [
            r["nim"], r["name"], r["attempt_no"] or 1,
            "submitted" if st == "submitted" else "started (belum kirim)",
            _fmt_ts(r["started_at"] or r["submitted_at"]),
            _fmt_ts(r["submitted_at"]) if st == "submitted" else "",
        ]
        row += [_given_display(s, answers) for s in SLOTS]
        row += [""]
        ws.append(row)

    for col in ws.columns:
        letter = col[0].column_letter
        best = 8
        for cell in col:
            if cell.value is not None and len(str(cell.value)) > best:
                best = min(len(str(cell.value)), 70)
        ws.column_dimensions[letter].width = best + 2
    ws.freeze_panes = "E2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="midterm_ai_attempts.xlsx"'},
    )


@app.delete("/api/attempts/{attempt_id}")
def delete_attempt(attempt_id: int, x_admin_key: str | None = Header(default=None)):
    """Admin-only: remove one attempt so a student can redo (also clears a stuck session)."""
    check_admin(x_admin_key)
    with db() as conn:
        cur = conn.execute("DELETE FROM attempts WHERE id=?", (attempt_id,))
    if cur.rowcount == 0:
        raise HTTPException(status_code=404, detail="Attempt tidak ditemukan.")
    return {"ok": True, "deleted": attempt_id}


@app.delete("/api/attempts")
def delete_all_for_nim(nim: str = "", x_admin_key: str | None = Header(default=None)):
    """Admin-only: remove ALL attempts for one NIM (clean slate for a full retake)."""
    check_admin(x_admin_key)
    nim = (nim or "").strip()
    if not nim:
        raise HTTPException(status_code=400, detail="NIM wajib diisi.")
    with db() as conn:
        cur = conn.execute("DELETE FROM attempts WHERE lower(nim)=lower(?)", (nim,))
    return {"ok": True, "deleted_rows": cur.rowcount}


@app.get("/health")
def health():
    return {"status": "ok"}
