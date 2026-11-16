# Exam Portal — capture-only, manually graded

A small FastAPI + SQLite exam portal for **open-ended papers that a human marks afterwards**:
multiple choice, free-text essays, and multi-part calculation questions (data tables,
parameter tables, several numeric inputs per sub-question). It captures answers, autosaves
them, and exports everything to Excel/DOCX for offline grading.

It deliberately has **no answer key and no auto-grading** — that is the point. If an exam has
worked calculations and essays, a spreadsheet of what each student actually wrote beats a
fragile grader every time.

Built for a 90-minute AI/ML midterm: 15 questions, 100 points, 60+ students per sitting,
a few hundred attempts.

## What it does

- **Question types** — `mcq` (radio), `text` (textarea, optional numeric line), and `group`
  (multi-part card with a/b/c/d, each with its own points, free text, a number, or **several
  numbered inputs**).
- **Rich cards** — data tables (`table`), a Parameter/Value table (`params`), per-question
  hints and neutral format placeholders.
- **Resume on refresh** — the attempt row is created on start, answers autosave every change,
  and a reload puts the student back where they were.
- **Server-anchored timer** — 90 minutes is computed from `started_at` on the server, so
  reloading the page cannot buy more time.
- **Proctored review screen** — the last step is a review of every answer; the real submit
  button lives at the bottom behind a confirm. Blanks are allowed until the student submits
  voluntarily; the timer never pauses and auto-submits at zero.
- **Admin console** (`/admin`, `X-Admin-Key`) — scoreboard of every attempt, full answer
  detail per student, and **Excel / DOCX export** for marking offline.
- **Zero LLM, zero external services** — SQLite on disk, no CDN, no telemetry.

## Quickstart

```bash
git clone <this repo> && cd exam-portal
python3 -m venv venv && . venv/bin/activate
pip install -r requirements.txt

export EXAM_ADMIN_KEY=$(python3 -c "import secrets;print(secrets.token_hex(16))")
echo "admin key: $EXAM_ADMIN_KEY"

cd backend
uvicorn main:app --host 127.0.0.1 --port 8022
```

Open <http://127.0.0.1:8022/> for the paper and <http://127.0.0.1:8022/admin> for the
console. The shipped `backend/questions.py` is **demo content** — replace it with your own
questions; the format is documented at the top of that file.

## Configuration

| Variable | Default | Meaning |
|---|---|---|
| `EXAM_ADMIN_KEY` | *(empty)* | Admin key for `/api/scoreboard` and exports. **When unset, the admin surface is closed (503)** rather than open. |
| `EXAM_DB` | `data/quiz.db` | SQLite location. Point it at a tmp file for experiments. |

Set these in the systemd unit (`systemd-unit.service.example`) rather than editing code.
Never commit a real key; `frontend/admin.html` ships with `const KEY = ""` on purpose.

## Endpoints

| Method | Path | Purpose |
|---|---|---|
| GET | `/` `/static/*` | Student paper |
| GET | `/api/questions` | Paper **without** any answer data |
| POST | `/api/start` | NIM + name → attempt row, resumes if one is already open |
| POST | `/api/save` | Autosave answers for the current attempt |
| POST | `/api/submit` | Flip the attempt to `submitted`, stamp the time |
| GET | `/health` | Liveness |
| GET | `/admin` | Admin HTML (unlocks with the key in the page) |
| GET | `/api/scoreboard` | Every attempt (needs `X-Admin-Key`) |
| GET | `/api/export_xlsx` | Excel export for offline grading (needs key) |
| GET | `/api/export_docx` | Per-student DOCX of one attempt (needs key) |

## Deploy

Behind a reverse proxy on localhost, one systemd unit per paper:

```bash
sudo cp systemd-unit.service.example /etc/systemd/system/exam-portal.service
sudo systemctl daemon-reload && sudo systemctl enable --now exam-portal
curl -s localhost:8022/health
```

Then publish it however you already publish localhost services (Cloudflare quick tunnel,
nginx, WireGuard). Keep it bound to `127.0.0.1`.

## What is NOT in this repo

The deployment this was extracted from holds real student attempts and a real exam paper.
This copy is **sanitized**:

- `data/` and every `*.db` are excluded (student NIM, names, answers);
- the admin key moved to `EXAM_ADMIN_KEY` (the real one was hardcoded in the original and is
  not in git history here);
- `backend/questions.py` is demo content, not the live paper;
- the `venv/` and per-deployment backups were dropped.

## License

MIT — see [LICENSE](LICENSE).
