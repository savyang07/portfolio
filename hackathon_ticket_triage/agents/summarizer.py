"""
Summarizer Agent

Reads each ticket and produces a short, plain-English summary of what the
user is reporting. In the demo, summaries are precomputed in tickets.json.

In production, this agent sends the ticket description to an LLM
(e.g., Claude claude-sonnet-4-5) with a structured prompt and parses the
returned summary. The output shape is identical — swap `demo_mode=False`
and add your API key to use a live model.

Input:  ticket dict (from tickets.json)
Output: {"summary": str}
"""

# Set to True to use precomputed demo data.
# Set to False and implement _call_llm() below to use a live model.
DEMO_MODE = True


def summarize(ticket: dict) -> dict:
    """
    Return a short summary for a single ticket.

    In demo mode, returns the precomputed ai_summary from tickets.json.
    In production, calls an LLM with the raw ticket description.
    """
    if DEMO_MODE:
        return {"summary": ticket.get("ai_summary", ticket.get("short_description", ""))}
    # Production path (not used in demo):
    return _call_llm(ticket)  # pragma: no cover


def _call_llm(ticket: dict) -> dict:  # pragma: no cover
    """
    Production implementation: send ticket to an LLM and parse summary.

    Expected prompt pattern:
        "Summarize this ServiceNow ticket in one plain sentence (max 20 words),
         focusing on what is broken or needed. Ticket: {description}"

    Replace with your Anthropic SDK call:
        import anthropic
        client = anthropic.Anthropic()
        msg = client.messages.create(model="claude-sonnet-4-5", max_tokens=60, ...)
        return {"summary": msg.content[0].text.strip()}
    """
    raise NotImplementedError("Set DEMO_MODE=False only after adding your LLM call here.")
