"""
Offline triage engine for the 'Triage a new ticket' tab.

The demo is precomputed-only (no API key), so new tickets typed in live are scored
by this lightweight stand-in: keyword rules for routing/priority plus TF-IDF cosine
similarity against the existing queue to find duplicates. In production this module
is where the LLM call would go — the output shape is the same.
"""
import math
import re
from collections import Counter

STOP = set("""a an the and or of to in on for is are was were be been it this that with
at by from as i my we our you your me can could please hi hello thanks thank not no
any some has have had do does did there what which who how when why get got its im""".split())

# (keywords, queue, category, base score, rationale)
RULES = [
    (["phi", "mrn", "patient names", "hipaa", "exposed", "visible to everyone", "all staff"],
     "DE – Data Access & Security", "Security / PHI", 95,
     "Mentions patient identifiers in a broadly shared place — potential PHI exposure."),
    (["talend", "job", "failed", "etl", "load", "feed", "lambda", "pipeline"],
     "DE – Pipelines (Talend/ETL)", "Pipeline failure", 70,
     "Describes a failed or delayed data load."),
    (["stale", "yesterday", "not updated", "didn't refresh", "old data", "blank", "no rows"],
     "DE – Pipelines (Talend/ETL)", "Stale data", 72,
     "Report showing old or missing data usually means an upstream load failed."),
    (["redshift", "timeout", "timing out", "slow", "query", "wlm"],
     "DE – Redshift / Warehouse", "Performance", 62,
     "Warehouse performance symptoms."),
    (["omop", "cerner", "snomed", "concept", "mapping", "duplicate rows", "doesn't match"],
     "DE – Clinical Data (Cerner/OMOP)", "Data quality", 55,
     "Clinical data correctness issue."),
    (["access", "permission", "permission denied", "license", "need access"],
     "DE – Data Access & Security", "Access request", 32,
     "Access provisioning — should be a Request, not an Incident."),
    (["dashboard", "tableau", "subscription", "workbook", "report"],
     "BI – Reporting (Tableau)", "Reporting", 45,
     "Reporting / dashboard issue."),
    (["password", "locked out", "login", "log in", "printer", "laptop", "vpn"],
     "IT Service Desk (not DE)", "Misrouted", 15,
     "End-user IT issue — belongs to the IT Service Desk."),
]

URGENCY = [("huddle", 12), ("transfer center", 10), ("ed ", 8), ("emergency", 8),
           ("cms", 8), ("deadline", 8), ("due", 4), ("urgent", 6), ("asap", 5),
           ("all units", 8), ("down", 6)]

REQUEST_HINTS = ["can i get", "please give", "need access", "request", "would like",
                 "set up", "question", "what counts", "how do i"]


GREETING = re.compile(r"^\s*(hi|hello|hey|good (morning|afternoon)|dear)[^,.!\n]*[,.!]?\s*", re.I)


def heuristic_summary(text, max_words=24):
    """Offline fallback: drop greetings/sign-offs, keep the first meaningful sentence."""
    body = GREETING.sub("", text.strip())
    body = re.split(r"\b(thanks|thank you|regards)\b", body, flags=re.I)[0].strip()
    first = re.split(r"(?<=[.!?])\s+", body)[0] if body else text.strip()
    words = first.split()
    s = " ".join(words[:max_words]) + ("…" if len(words) > max_words else "")
    return s[:1].upper() + s[1:]


def tokens(text):
    return [w for w in re.findall(r"[a-z0-9_]+", text.lower()) if w not in STOP and len(w) > 2]


class SimilarityIndex:
    def __init__(self, tickets):
        self.tickets = tickets
        docs = [tokens(t["short_description"] + " " + t["description"]) for t in tickets]
        df = Counter(w for d in docs for w in set(d))
        n = len(docs)
        self.idf = {w: math.log((1 + n) / (1 + c)) + 1 for w, c in df.items()}
        self.vecs = [self._vec(d) for d in docs]

    def _vec(self, toks):
        tf = Counter(toks)
        v = {w: c * self.idf.get(w, math.log(len(self.tickets) + 1) + 1) for w, c in tf.items()}
        norm = math.sqrt(sum(x * x for x in v.values())) or 1.0
        return {w: x / norm for w, x in v.items()}

    def top(self, text, k=3):
        q = self._vec(tokens(text))
        scored = [(sum(q[w] * v.get(w, 0) for w in q), t) for v, t in zip(self.vecs, self.tickets)]
        return sorted(scored, key=lambda s: -s[0])[:k]


def triage(text, index, queues, group_threshold=0.30):
    low = text.lower()
    best = None
    for kws, queue, cat, base, why in RULES:
        hits = [k for k in kws if k in low]
        if hits and (best is None or len(hits) * 10 + base > best[0]):
            best = (len(hits) * 10 + base, queue, cat, base, why, hits)
    if best is None:
        queue, cat, score, why, hits = "BI – Reporting (Tableau)", "General", 35, \
            "No strong signal — defaulting to BI intake for a human look.", []
        confidence = 0.45
    else:
        _, queue, cat, score, why, hits = best
        confidence = min(0.95, 0.6 + 0.1 * len(hits))

    bumps = [(w, b) for w, b in URGENCY if re.search(r"\b" + re.escape(w.strip()) + r"\b", low)]
    score = min(99, score + sum(b for _, b in bumps))

    similar = index.top(text)
    match = None
    top_sim = similar[0][0] if similar else 0
    # grouped tickets match at a lower bar (the group is already confirmed);
    # a lone open ticket needs a closer match to call it a duplicate
    if similar and (top_sim >= group_threshold and similar[0][1]["group_id"] or top_sim >= 0.5):
        match = similar[0][1]
        # duplicates inherit the existing issue's triage — same queue, owner, category
        queue, cat = match["ai_route_queue"], match["ai_category"]
        score = match["ai_priority_score"]
        confidence = max(confidence, round(min(0.95, 0.5 + top_sim), 2))

    should_be_request = (cat == "Access request" or any(h in low for h in REQUEST_HINTS)) and cat in (
        "Access request", "Reporting", "General", "Data change request")

    summary = heuristic_summary(text)

    reasons = [match["ai_rationale"] if match else why]
    if hits and not match:
        reasons.append("Signals: " + ", ".join(f"“{h}”" for h in hits[:4]) + ".")
    if bumps and not match:
        reasons.append("Urgency bump for: " + ", ".join(w.strip() for w, _ in bumps) + ".")
    if match:
        where = f" ({match['group_id']})" if match["group_id"] else " (open, unassigned)"
        reasons.append(f"Matches {match['number']}{where} at similarity {top_sim:.2f}, "
                       f"so it inherits that ticket's triage.")

    return {
        "summary": summary,
        "category": cat,
        "score": int(score),
        "queue": queue,
        "owner": (match["ai_route_owner"] if match else queues[queue]),
        "confidence": round(confidence, 2),
        "rationale": " ".join(reasons),
        "should_be_request": should_be_request,
        "group_match": match,
        "similar": similar,
    }
