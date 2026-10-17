"""
Mock ServiceNow MCP Tool Layer

This module simulates the MCP (Model Context Protocol) server that would
connect the AI agents to Keck's live ServiceNow environment.

In production, each function below maps to a ServiceNow Table API endpoint
exposed through an approved MCP server. For the hackathon demo, every
function reads from the local tickets.json file and logs a "mock MCP call"
so the demo code is honest about what is real vs. simulated.

NOTE FOR REVIEWERS: This is a demonstration only. No real ServiceNow data
is accessed. All tickets are synthetic (no PHI, no real patient or employee
data).
"""
import json
from pathlib import Path

_DATA_PATH = Path(__file__).parent.parent / "data" / "tickets.json"


def _load_data() -> dict:
    """Load the synthetic ticket dataset from disk."""
    return json.loads(_DATA_PATH.read_text())


def _mock_log(endpoint: str, params: dict = None) -> None:
    """Log a simulated MCP call so the demo is transparent about what's real."""
    param_str = f" params={params}" if params else ""
    print(f"[MOCK MCP] GET {endpoint}{param_str}")


# ---------------------------------------------------------------------------
# Mock tool functions (each mirrors a real ServiceNow Table API endpoint)
# ---------------------------------------------------------------------------

def get_tickets(assignment_group: str = "Data Engineering", state: str = "New") -> list:
    """
    Return all open tickets assigned to a group.

    Production endpoint:
        GET /api/now/table/incident
            ?assignment_group=<sys_id>&state=1&sysparm_fields=...
        (also: /api/now/table/sc_request for Requests)

    MCP tool name: servicenow.get_tickets
    """
    _mock_log("/api/now/table/incident", {"assignment_group": assignment_group, "state": state})
    data = _load_data()
    return [
        t for t in data["tickets"]
        if t.get("assignment_group") == assignment_group
        and t.get("state") == state
    ]


def get_ticket_details(ticket_number: str) -> dict:
    """
    Return full details for a single ticket.

    Production endpoint:
        GET /api/now/table/incident
            ?number=<ticket_number>&sysparm_limit=1

    MCP tool name: servicenow.get_ticket_details
    """
    _mock_log(f"/api/now/table/incident", {"number": ticket_number})
    data = _load_data()
    matches = [t for t in data["tickets"] if t["number"] == ticket_number]
    if not matches:
        raise KeyError(f"Ticket {ticket_number} not found in mock data.")
    return matches[0]


def update_ticket(ticket_number: str, fields: dict) -> dict:
    """
    Update fields on a ticket (e.g., priority, assignment_group, work_notes).

    Production endpoint:
        PATCH /api/now/table/incident/<sys_id>
              body: {fields}

    MCP tool name: servicenow.update_ticket

    NOTE: In the demo this is a no-op. No data is written.
    """
    _mock_log(f"/api/now/table/incident/{ticket_number}", fields)
    print(f"[MOCK MCP] (demo) Would update {ticket_number} with {fields} — no write in demo mode.")
    return {"status": "mock_success", "ticket_number": ticket_number, "updated_fields": fields}


def assign_ticket(ticket_number: str, queue: str, owner: str) -> dict:
    """
    Assign a ticket to a queue and an individual owner.

    Production endpoint:
        PATCH /api/now/table/incident/<sys_id>
              body: {"assignment_group": <group_sys_id>, "assigned_to": <user_sys_id>}

    MCP tool name: servicenow.assign_ticket

    NOTE: In the demo this is a no-op.
    """
    _mock_log(f"/api/now/table/incident/{ticket_number}",
              {"assignment_group": queue, "assigned_to": owner})
    print(f"[MOCK MCP] (demo) Would assign {ticket_number} → {queue} / {owner}")
    return {"status": "mock_success", "ticket_number": ticket_number, "queue": queue, "owner": owner}


def add_work_note(ticket_number: str, note: str) -> dict:
    """
    Post a work note to a ticket (internal note, not visible to caller).

    Production endpoint:
        PATCH /api/now/table/incident/<sys_id>
              body: {"work_notes": "<note text>"}

    MCP tool name: servicenow.add_work_note

    NOTE: In the demo this is a no-op.
    """
    _mock_log(f"/api/now/table/incident/{ticket_number}", {"work_notes": note[:40] + "…"})
    print(f"[MOCK MCP] (demo) Would add work note to {ticket_number}")
    return {"status": "mock_success", "ticket_number": ticket_number}
