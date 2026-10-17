"""
Grouping Agent

Finds duplicate tickets and groups them so a single engineer can resolve
them together.

In the demo, group assignments are precomputed in tickets.json (group_id,
group_role fields). This module wraps that data in a clean interface.

In production, this agent would use embedding similarity (e.g., compare
ticket descriptions via an LLM embedding model) to find near-duplicates
automatically — no precomputed labels needed.

Input:  tickets (list of ticket dicts), groups_data (dict from tickets.json)
Output: {
    "groups": {
        "GRP-XX": {
            "title": str,
            "root_cause": str,
            "recommended_action": str,
            "notify": list,
            "comms_draft": str,
            "member_count": int,
        }
    },
    "assignments": {
        "INC0104511": "GRP-01",   # grouped ticket → group id
        "INC0104522": None,       # singleton → None
    }
}
"""


def find_groups(tickets: list, groups_data: dict) -> dict:
    """
    Return group data and per-ticket assignments from precomputed fields.

    In production, replace the precomputed lookup with embedding similarity
    across ticket descriptions to find duplicates automatically.
    """
    assignments = {}
    member_counts: dict = {}

    for ticket in tickets:
        gid = ticket.get("group_id") or None
        assignments[ticket["number"]] = gid
        if gid:
            member_counts[gid] = member_counts.get(gid, 0) + 1

    # Enrich group data with member counts
    enriched_groups = {}
    for gid, gdata in groups_data.items():
        enriched_groups[gid] = {**gdata, "member_count": member_counts.get(gid, 0)}

    return {"groups": enriched_groups, "assignments": assignments}
