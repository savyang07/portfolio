"""
Routing Agent

Decides which queue and owner each ticket should go to, based on its
category, content, and who has resolved similar tickets in the past.

In the demo, routing decisions are precomputed in tickets.json
(ai_route_queue, ai_route_owner, ai_confidence).

In production, this agent would:
  1. Pull the team calendar and current workload from ServiceNow via MCP.
  2. Query a vector store of resolved tickets to find the expert for this
     category.
  3. Use an LLM to reason across those signals and return a routing decision.

Input:  ticket dict, queues dict (queue name → default owner)
Output: {
    "queue": str,
    "owner": str,
    "confidence": float,
    "rationale": str,
}
"""

DEMO_MODE = True


def route_ticket(ticket: dict, queues: dict) -> dict:
    """
    Return the routing decision for a single ticket.

    In demo mode, returns the precomputed values from tickets.json.
    In production, calls an LLM with team availability and history signals.
    """
    if DEMO_MODE:
        queue = ticket.get("ai_route_queue", next(iter(queues)))
        return {
            "queue": queue,
            "owner": ticket.get("ai_route_owner", queues.get(queue, "Unassigned")),
            "confidence": float(ticket.get("ai_confidence", 0.75)),
            "rationale": ticket.get("ai_rationale", ""),
        }
    return _call_llm(ticket, queues)  # pragma: no cover


def _call_llm(ticket: dict, queues: dict) -> dict:  # pragma: no cover
    """
    Production implementation: route via LLM reasoning.

    Pull team workload from ServiceNow MCP (tools/servicenow_mock.py),
    then send ticket + team context to Claude and parse the JSON response.
    """
    raise NotImplementedError("Set DEMO_MODE=False only after adding your LLM call here.")
