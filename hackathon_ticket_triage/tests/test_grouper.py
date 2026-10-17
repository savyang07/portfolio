"""
Tests for agents/grouper.py

The grouper maps tickets to their duplicate groups. These tests verify that
group assignments are consistent with the source data and that singletons
are handled correctly.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from agents.grouper import find_groups


def _make_tickets():
    return [
        {"number": "INC001", "group_id": "GRP-A", "group_role": "parent",
         "short_description": "foo", "description": "foo"},
        {"number": "INC002", "group_id": "GRP-A", "group_role": "child",
         "short_description": "foo dupe", "description": "foo dupe"},
        {"number": "INC003", "group_id": "", "group_role": "",
         "short_description": "bar", "description": "bar"},
    ]


def _make_groups():
    return {
        "GRP-A": {
            "title": "Foo outage",
            "root_cause": "Something broke",
            "recommended_action": "Fix it",
            "notify": ["Team A"],
            "comms_draft": "Hi all",
        }
    }


def test_grouped_tickets_assigned_to_correct_group():
    result = find_groups(_make_tickets(), _make_groups())
    assert result["assignments"]["INC001"] == "GRP-A"
    assert result["assignments"]["INC002"] == "GRP-A"


def test_singleton_ticket_assigned_none():
    result = find_groups(_make_tickets(), _make_groups())
    assert result["assignments"]["INC003"] is None


def test_groups_output_contains_group_data():
    result = find_groups(_make_tickets(), _make_groups())
    assert "GRP-A" in result["groups"]
    assert result["groups"]["GRP-A"]["title"] == "Foo outage"


def test_group_member_count_correct():
    result = find_groups(_make_tickets(), _make_groups())
    assert result["groups"]["GRP-A"]["member_count"] == 2


def test_all_tickets_accounted_for():
    tickets = _make_tickets()
    result = find_groups(tickets, _make_groups())
    assert len(result["assignments"]) == len(tickets)
