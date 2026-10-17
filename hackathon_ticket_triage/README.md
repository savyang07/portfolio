# Triage Copilot

Agentic AI demo for ServiceNow ticket triage — built for the **Keck Medicine of USC Agentic AI in Healthcare Hackathon**.

Runs entirely offline on synthetic data. No API key or ServiceNow instance needed.

---

## What it does

The demo shows a Data Engineering on-call queue before and after an AI triage agent:

| Before | After |
|---|---|
| 38 raw, unsorted tickets | 22 unique issues (16 duplicates collapsed) |
| No priority scores | AI priority score 0–100 with per-factor breakdown |
| PHI-exposure ticket buried at row 19 | Critical ticket surfaced to top automatically |
| 8 mistyped incidents (wrong SLA clock) | Type/team flagged for review |
| ~3.8 hrs estimated manual triage | ~33 min estimated AI-assisted review |

Four agent modules make the story visible as real code:

```
Summarizer → Grouper → Priority Scorer → Router
```

---

## Quick start

```bash
cd hackathon_ticket_triage
pip install -r requirements.txt
streamlit run app.py
```

Opens at `http://localhost:8501`.

---

## Folder map

```
hackathon_ticket_triage/
├── app.py                  # Thin router: page config, load data, two tabs
├── config/
│   └── scoring.json        # Priority weights, thresholds, colors — edit here
├── agents/
│   ├── summarizer.py       # Ticket → plain-English summary
│   ├── grouper.py          # Tickets → duplicate groups
│   ├── scorer.py           # Ticket + config → 0-100 score with breakdown
│   └── router.py           # Ticket → queue + owner + confidence
├── tools/
│   └── servicenow_mock.py  # Mock MCP layer (mirrors real ServiceNow Table API)
├── ui/
│   ├── components.py       # Shared CSS, helpers, color constants
│   ├── before_tab.py       # "Before" tab: raw inbox, pain-point callouts
│   └── after_tab.py        # "After" tab: Run Triage animation, results
├── triage/
│   └── offline_engine.py   # Keyword rules + TF-IDF for live new-ticket input
├── data/
│   └── tickets.json        # 38 precomputed synthetic tickets
├── tests/
│   ├── test_scorer.py      # Unit tests for priority scorer
│   └── test_grouper.py     # Unit tests for duplicate grouper
└── docs/
    └── DEMO_GUIDE.md       # Full demo script, Q&A, Mermaid diagram
```

---

## What's mocked vs real

| Component | Demo | Production path |
|---|---|---|
| Ticket data | 38 synthetic tickets in `data/tickets.json` | ServiceNow Table API via MCP |
| Summaries | Precomputed `ai_summary` field | LLM call (e.g. Claude claude-haiku-5-5) |
| Duplicate grouping | Precomputed `group_id` field | Embedding similarity + LLM cluster labels |
| Priority scores | Precomputed `ai_priority_score` | Weighted rule engine or fine-tuned model |
| Routing | Precomputed `ai_queue` / `ai_owner` | Team load + calendar lookup + LLM |
| ServiceNow writes | No-op (logs `[MOCK MCP]` to console) | PATCH `/api/now/table/incident/<sys_id>` |

See `tools/servicenow_mock.py` for the exact endpoint each mock function would call.

---

## Scoring config

Edit `config/scoring.json` to tune the priority model:

```json
{
  "factors": [
    {"name": "patient_care_criticality", "weight": 30},
    {"name": "users_affected",           "weight": 25},
    ...
  ],
  "thresholds": {"p1_min": 85, "p2_min": 65, "p3_min": 40}
}
```

All weights are **placeholders** — replace with values validated against your historical ticket data before going live.

---

## Running tests

```bash
pytest tests/ -v
```

All 15 tests should pass (10 scorer + 5 grouper).

---

## Swapping in a real model

To replace precomputed outputs with live LLM calls:

1. Set `DEMO_MODE = False` in `agents/summarizer.py` and `agents/router.py`.
2. Fill in `_call_llm()` in `summarizer.py` with your Anthropic SDK call.
3. Point `tools/servicenow_mock.py` functions at your live ServiceNow MCP server.
4. Replace `ai_*` fields in `tickets.json` with live reads from ServiceNow.

---

## Data sources

All ticket data is **100% synthetic** — generated for this demo. No real patient, employee, or Keck Medicine data is used anywhere. No PHI. The scoring weights are placeholder values and have not been validated against production data.

See `docs/DEMO_GUIDE.md` for the full responsible-AI disclosure and next-steps roadmap.
