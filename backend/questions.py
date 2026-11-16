"""Question bank — DEMO CONTENT. Replace this file with your own paper.

The engine never grades anything: there are no answer keys here, by design.
Every answer is captured and exported (Excel/DOCX) for a human to mark.

Question shape
--------------
Every question needs:
    id, section ("A"|"B"|"C"), topic, type, points, question

type="mcq"    -> `options` (list of strings). No `correct` key exists.
type="text"   -> one free-text box. Optional: kind="num" (numeric line),
                 placeholder (FORMAT hints only — never hint at the answer).
type="group"  -> `subs`: [{key, points, question, ...}] rendered as a/b/c/d
                 inside one card. A sub with `inputs` gets one answer box per
                 input: [{"key","kind","label"}] where kind="num" or "text".

Optional extras on any question or sub:
    table  -> {"caption", "headers", "rows"}  a data table on the card
    params -> [["Parameter", "Value"], ...]   a 2-column parameter table
    hint   -> short coaching line under the question

Run the server and open / to see this demo paper render exactly as a real one
would; the point values below sum to TOTAL_POINTS.
"""

# ---------------------------------------------------------------- SECTION A
A_MCQ = [
    {
        "id": 1,
        "section": "A",
        "topic": "Demo — Multiple Choice",
        "type": "mcq",
        "points": 3,
        "question": (
            "[DEMO] A model is trained on past exam data to predict a new student's "
            "score. Which kind of learning is this?"
        ),
        "options": [
            "Unsupervised Learning",
            "Supervised Learning",
            "Reinforcement Learning",
            "Clustering",
        ],
    },
]

# ---------------------------------------------------------------- SECTION B
B_TEXT = [
    {
        "id": 2,
        "section": "B",
        "topic": "Demo — Short Essay",
        "type": "text",
        "points": 5,
        "question": (
            "[DEMO] In two or three sentences, explain the difference between a "
            "parameter and a hyperparameter. Give one example of each."
        ),
        "hint": "Maksimal 3 kalimat.",
    },
]

# ---------------------------------------------------------------- SECTION C
C_GROUP = [
    {
        "id": 3,
        "section": "C",
        "topic": "Demo — Calculation with a data table",
        "type": "group",
        "points": 7,
        "question": (
            "[DEMO] A company models service time against machine age. Use the table."
        ),
        "table": {
            "caption": "Machine age and service time",
            "headers": ["Machine age (years)", "Service time (minutes)"],
            "rows": [
                ["1.0", "28"], ["2.0", "32"], ["3.0", "37"],
                ["4.0", "43"], ["5.0", "48"],
            ],
        },
        "subs": [
            {
                "key": "a",
                "points": 3,
                "question": "[DEMO] Write the linear regression equation from this data.",
                "kind": "text",
                "placeholder": "contoh format: y = 1 + 1x",
            },
            {
                "key": "b",
                "points": 4,
                "question": (
                    "[DEMO] Use the model to predict the service time for a 7-year-old "
                    "machine. Round to the nearest minute."
                ),
                "kind": "num",
            },
        ],
    },
    {
        "id": 4,
        "section": "C",
        "topic": "Demo — Multi-input calculation",
        "type": "group",
        "points": 5,
        "question": "[DEMO] Given one neuron, compute the pre-activation and the output.",
        "params": [
            ["Weights (w1, w2)", "0.5, 0.5"],
            ["Bias (b)", "-1"],
            ["Inputs (x1, x2)", "1, 2"],
            ["Activation", "sigmoid"],
        ],
        "subs": [
            {
                "key": "a",
                "points": 3,
                "question": "[DEMO] Compute z and the activation output.",
                "inputs": [
                    {"key": "z", "kind": "num", "label": "z"},
                    {"key": "out", "kind": "num", "label": "output (sigmoid)"},
                ],
            },
            {
                "key": "b",
                "points": 2,
                "question": "[DEMO] Does the output rise or fall if x1 becomes 2? Explain briefly.",
                "kind": "text",
                "placeholder": "contoh: output naik karena z membesar",
            },
        ],
    },
]

QUESTIONS = A_MCQ + B_TEXT + C_GROUP

# Total poin per section (for the header/admin info)
SECTION_POINTS = {"A": sum(q["points"] for q in A_MCQ),
                  "B": sum(q["points"] for q in B_TEXT),
                  "C": sum(q["points"] for q in C_GROUP)}
TOTAL_POINTS = sum(SECTION_POINTS.values())
