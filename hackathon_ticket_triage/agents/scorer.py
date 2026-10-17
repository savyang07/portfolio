"""
Priority Scorer Agent — ERWG formula

Reads per-ticket ERWG dimension scores (already precomputed by the AI agent
in the demo) and computes the final priority using the Enterprise Reporting
Working Group formula:

    raw   = patient_safety×8 + compliance×7 + financial×6 + utilization×4
    score = round(raw × org_multiplier)    (org_multiplier ∈ {1.0,0.8,0.6,0.4,0.0})

Thresholds: Critical ≥70, High 50–69, Medium 30–49, Low ≤29

In production the AI agent would derive each 0–4 dimension score from
ticket text, form fields, SLA data, and training on past scored requests.
The output shape is identical to the demo.

Input:  ticket dict (from tickets.json), scoring config dict (from
        config/scoring.json)
Output: {
    "score":  int,             # 0–100 ERWG score
    "label":  str,             # "Critical" / "High" / "Medium" / "Low"
    "needs_more_info": bool,   # True if AI flagged missing info
    "info_requests": list[str],# questions for the BA to ask the requester
    "breakdown": [             # 5 rows — one per ERWG dimension
        {
          "key":          str,   # factor key
          "label":        str,   # human-readable name
          "score":        int,   # 0–4 dimension score
          "weight_label": str,   # "×8", "×7", "×6", "×4", or "multiplier"
          "points":       int,   # contribution to raw or effective score
          "level_label":  str,   # text description of this score level
        }
    ]
}
"""

ORG_PCT = {4: 1.0, 3: 0.8, 2: 0.6, 1: 0.4, 0: 0.0}


def score_ticket(ticket: dict, config: dict) -> dict:
    """Compute ERWG score and return label + 5-row breakdown."""
    org        = int(ticket.get("ai_org_impact",     0))
    patient    = int(ticket.get("ai_patient_safety", 0))
    compliance = int(ticket.get("ai_compliance",     0))
    financial  = int(ticket.get("ai_financial",      0))
    utilization= int(ticket.get("ai_utilization",    0))

    org_pct  = ORG_PCT.get(org, 0.0)
    raw      = patient * 8 + compliance * 7 + financial * 6 + utilization * 4
    score    = round(raw * org_pct)

    label    = _label_for_score(score, config["thresholds"])
    breakdown= _build_breakdown(
        org, org_pct, patient, compliance, financial, utilization, score, config
    )

    return {
        "score":           score,
        "label":           label,
        "needs_more_info": bool(ticket.get("ai_needs_more_info", False)),
        "info_requests":   list(ticket.get("ai_info_requests", [])),
        "breakdown":       breakdown,
    }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _label_for_score(score: int, thresholds: dict) -> str:
    if score >= thresholds["critical_min"]:
        return "Critical"
    if score >= thresholds["high_min"]:
        return "High"
    if score >= thresholds["medium_min"]:
        return "Medium"
    return "Low"


def _build_breakdown(
    org: int, org_pct: float,
    patient: int, compliance: int, financial: int, utilization: int,
    score: int, config: dict,
) -> list:
    factors_cfg = {f["key"]: f for f in config["factors"]}

    def level_label(factor_key: str, val: int) -> str:
        lvls = factors_cfg.get(factor_key, {}).get("levels", {})
        return lvls.get(str(val), f"Score {val}")

    raw_patient    = patient    * 8
    raw_compliance = compliance * 7
    raw_financial  = financial  * 6
    raw_utilization= utilization* 4

    rows = [
        {
            "key":          "org_impact",
            "label":        "Organizational Impact",
            "score":        org,
            "weight_label": "multiplier",
            "points":       score,   # the multiplier shrinks all other contributions
            "level_label":  level_label("org_impact", org),
            "multiplier":   f"{int(org_pct*100)}%",
        },
        {
            "key":          "patient_safety",
            "label":        "Patient Impact & Safety Risk",
            "score":        patient,
            "weight_label": "×8",
            "points":       round(raw_patient    * org_pct),
            "level_label":  level_label("patient_safety", patient),
        },
        {
            "key":          "compliance",
            "label":        "Compliance & Regulatory",
            "score":        compliance,
            "weight_label": "×7",
            "points":       round(raw_compliance * org_pct),
            "level_label":  level_label("compliance", compliance),
        },
        {
            "key":          "financial",
            "label":        "Financial Impact",
            "score":        financial,
            "weight_label": "×6",
            "points":       round(raw_financial  * org_pct),
            "level_label":  level_label("financial", financial),
        },
        {
            "key":          "utilization",
            "label":        "Report / Information Utilization",
            "score":        utilization,
            "weight_label": "×4",
            "points":       round(raw_utilization* org_pct),
            "level_label":  level_label("utilization", utilization),
        },
    ]
    return rows
