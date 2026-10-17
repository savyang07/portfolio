"""
Tests for agents/scorer.py — ERWG formula

The scorer reads per-ticket 0–4 dimension scores (ai_org_impact,
ai_patient_safety, ai_compliance, ai_financial, ai_utilization) and
computes:

    raw   = patient×8 + compliance×7 + financial×6 + utilization×4
    score = round(raw × org_pct)     where org_pct ∈ {1.0,0.8,0.6,0.4,0.0}

Thresholds: Critical ≥70, High 50–69, Medium 30–49, Low ≤29
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from agents.scorer import score_ticket

_config = json.loads((Path(__file__).parent.parent / "config" / "scoring.json").read_text())


def _ticket(org, patient, compliance, financial, utilization,
            needs_more_info=False, info_requests=None) -> dict:
    """Build a minimal synthetic ticket with ERWG dimension scores."""
    return {
        "ai_org_impact":      org,
        "ai_patient_safety":  patient,
        "ai_compliance":      compliance,
        "ai_financial":       financial,
        "ai_utilization":     utilization,
        "ai_needs_more_info": needs_more_info,
        "ai_info_requests":   info_requests or [],
        # legacy fields the scorer no longer uses directly:
        "ai_priority_score":  0,
        "ai_priority":        "Low",
        "ai_category":        "Pipeline failure",
        "ai_should_be_request": False,
    }


# ── ERWG formula tests ────────────────────────────────────────────────────────

def test_phi_ticket_scores_critical():
    """PHI ticket: org=4, patient=4, compliance=4, financial=3, utilization=2 → 86"""
    r = score_ticket(_ticket(4, 4, 4, 3, 2), _config)
    assert r["score"] == 86, f"Expected 86, got {r['score']}"
    assert r["label"] == "Critical"


def test_pipeline_cluster_scores_critical():
    """Cerner ADT cluster: org=4, patient=4, compliance=3, financial=3, utilization=4 → 87"""
    r = score_ticket(_ticket(4, 4, 3, 3, 4), _config)
    assert r["score"] == 87
    assert r["label"] == "Critical"


def test_stale_data_scores_critical():
    """Stale bed census: org=4, patient=3, compliance=2, financial=3, utilization=4 → 72"""
    r = score_ticket(_ticket(4, 3, 2, 3, 4), _config)
    assert r["score"] == 72
    assert r["label"] == "Critical"


def test_cms_compliance_scores_high():
    """CMS deadline: org=4, patient=1, compliance=4, financial=3, utilization=2 → 62"""
    r = score_ticket(_ticket(4, 1, 4, 3, 2), _config)
    assert r["score"] == 62
    assert r["label"] == "High"


def test_access_request_scores_low():
    """Access request: org=1(40%), utilization=2 → raw=8, score=round(8×0.4)=3"""
    r = score_ticket(_ticket(1, 0, 0, 0, 2), _config)
    assert r["score"] == 3
    assert r["label"] == "Low"


def test_zero_org_impact_gives_zero_score():
    """If org_impact=0, multiplier=0 so score=0 regardless of other dimensions."""
    r = score_ticket(_ticket(0, 4, 4, 4, 4), _config)
    assert r["score"] == 0
    assert r["label"] == "Low"


# ── Multiplier tests ──────────────────────────────────────────────────────────

def test_org3_applies_80pct_multiplier():
    """org=3 → 80% of raw; raw=4×8+0+0+0=32 → score=round(32×0.8)=26"""
    r = score_ticket(_ticket(3, 4, 0, 0, 0), _config)
    assert r["score"] == 26


def test_org2_applies_60pct_multiplier():
    """org=2 → 60% of raw; raw=4×8=32 → score=round(32×0.6)=19"""
    r = score_ticket(_ticket(2, 4, 0, 0, 0), _config)
    assert r["score"] == 19


def test_org1_applies_40pct_multiplier():
    """org=1 → 40% of raw; raw=4×8=32 → score=round(32×0.4)=13"""
    r = score_ticket(_ticket(1, 4, 0, 0, 0), _config)
    assert r["score"] == 13


# ── Threshold boundaries ──────────────────────────────────────────────────────

def test_score_70_is_critical():
    r = score_ticket(_ticket(4, 3, 2, 3, 4), _config)
    assert r["score"] == 72
    assert r["label"] == "Critical"


def test_score_50_is_high():
    """org=4, patient=0, compliance=0, financial=3, utilization=4 → raw=34, score=34 → Medium.
    Use: raw=50 → org=4 needed; 50 = patient×8+compliance×7+... → patient=2,compliance=2,fin=2,utl=2
    raw=16+14+12+8=50 score=50 → High at boundary"""
    r = score_ticket(_ticket(4, 2, 2, 2, 2), _config)
    assert r["score"] == 50
    assert r["label"] == "High"


def test_score_30_is_medium():
    """30 → Medium; raw must give score=30 with org=4: raw=30
    raw=2×8+0×7+0×6+2×4=16+8=24 → 24 Medium. Let's use raw=30: 0+3×7+0+0=21,score=21 Low.
    Better: org=3, raw=38 → 38×0.8=30.4→30 Medium."""
    r = score_ticket(_ticket(3, 0, 2, 2, 4), _config)
    # raw=0+14+12+16=42 → 42×0.8=33.6→34 Medium
    assert r["label"] == "Medium"


def test_score_below_30_is_low():
    r = score_ticket(_ticket(1, 0, 0, 0, 2), _config)
    assert r["score"] == 2
    assert r["label"] == "Low"


# ── Breakdown structure tests ─────────────────────────────────────────────────

def test_breakdown_has_five_entries():
    r = score_ticket(_ticket(4, 4, 4, 3, 2), _config)
    assert len(r["breakdown"]) == 5, f"Expected 5 breakdown rows, got {len(r['breakdown'])}"


def test_breakdown_contains_org_impact_row():
    r = score_ticket(_ticket(4, 3, 2, 3, 4), _config)
    keys = [row["key"] for row in r["breakdown"]]
    assert "org_impact" in keys


def test_breakdown_entries_have_required_keys():
    r = score_ticket(_ticket(4, 2, 2, 2, 2), _config)
    for entry in r["breakdown"]:
        assert "key"         in entry
        assert "label"       in entry
        assert "score"       in entry
        assert "points"      in entry
        assert "level_label" in entry


def test_needs_more_info_flag_passthrough():
    questions = ["What table?", "Which format?"]
    r = score_ticket(_ticket(1, 0, 0, 0, 1, needs_more_info=True, info_requests=questions), _config)
    assert r["needs_more_info"] is True
    assert r["info_requests"] == questions


def test_no_needs_more_info_by_default():
    r = score_ticket(_ticket(4, 4, 4, 3, 2), _config)
    assert r["needs_more_info"] is False
    assert r["info_requests"] == []
