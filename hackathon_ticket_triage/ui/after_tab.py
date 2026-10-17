"""
After Tab — "AI Triage Agent"

Shows the same tickets after they have been processed by the four agents.
The "Run Triage" button plays the agent steps step-by-step so the audience
sees agents working, not just a finished result.

State in st.session_state:
  triage_done   — bool, whether the animation has been played this session
  decisions     — dict of routing decisions made by the user
  linked        — set of group IDs whose children the user linked
  notified      — set of group IDs that received the status update
"""
import html
import json
import time
from pathlib import Path

import pandas as pd
import streamlit as st

from ui.components import age, pill, priority_pill, plabel_for, pkey, score_box, PCOLOR, PBG
from agents.scorer import score_ticket

_CFG_PATH = Path(__file__).parent.parent / "config" / "scoring.json"
_SCORING_CFG = json.loads(_CFG_PATH.read_text())

PRIORITY_ORDER = ["Critical", "High", "Medium", "Low"]


# ---------------------------------------------------------------------------
# Session state defaults
# ---------------------------------------------------------------------------
def _init_state():
    defaults = {
        "triage_done": False,
        "decisions": {},
        "linked": set(),
        "notified": set(),
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ---------------------------------------------------------------------------
# Build the issues view (collapse duplicate groups into one row)
# ---------------------------------------------------------------------------
def build_issues(df: pd.DataFrame, groups: dict) -> pd.DataFrame:
    rows = []
    for gid, g in df[df.group_id != ""].groupby("group_id"):
        parent = g[g.group_role == "parent"].iloc[0].to_dict()
        top = g.ai_priority_score.max()
        rows.append({
            **parent,
            "key": gid,
            "is_group": True,
            "n": len(g),
            "members": g.number.tolist(),
            "ai_priority_score": int(top),
            "ai_priority": plabel_for(top),
            "ai_summary": groups[gid]["title"],
            "oldest_hrs": g.age_hrs.max(),
            "n_request": int(g.ai_should_be_request.sum()),
        })
    for _, r in df[df.group_id == ""].iterrows():
        rows.append({
            **r.to_dict(),
            "key": r.number,
            "is_group": False,
            "n": 1,
            "members": [r.number],
            "oldest_hrs": r.age_hrs,
            "n_request": int(r.ai_should_be_request),
        })
    return pd.DataFrame(rows).sort_values("ai_priority_score", ascending=False)


# ---------------------------------------------------------------------------
# Run Triage animation
# ---------------------------------------------------------------------------
def _run_triage_animation(n_tickets: int) -> None:
    steps = [
        ("🔍", "Summarizer", f"Reading {n_tickets} tickets and generating plain-English summaries…"),
        ("🔗", "Grouping agent", "Comparing tickets for duplicates and clustering related issues…"),
        ("⚖️",  "Priority scorer", "Assigning 0–4 ERWG dimension scores and computing final priority…"),
        ("📬", "Routing agent", "Matching each issue to the right queue and suggested owner…"),
    ]
    placeholder = st.empty()
    for icon, name, desc in steps:
        with placeholder.container():
            st.markdown(
                f'<div class="agent-step">'
                f'<span class="running">⟳ {icon} {name}</span>&nbsp;&nbsp;{html.escape(desc)}'
                f'</div>',
                unsafe_allow_html=True,
            )
        time.sleep(1.5)
    with placeholder.container():
        st.success(
            "✅ All four agents finished. "
            f"{n_tickets} tickets → issues scored, grouped, and routed."
        )
    time.sleep(0.8)
    placeholder.empty()
    st.session_state.triage_done = True
    st.rerun()


# ---------------------------------------------------------------------------
# KPI strip
# ---------------------------------------------------------------------------
def _render_kpis(df: pd.DataFrame, issues: pd.DataFrame, manual_min: float, review_min: float) -> None:
    n_tickets = len(df)
    n_issues  = len(issues)
    n_dupes   = n_tickets - n_issues
    n_critical= int((issues.ai_priority == "Critical").sum())
    n_pending = int(df.get("ai_needs_more_info", pd.Series(dtype=bool)).sum())
    before_min= n_tickets * manual_min
    after_min = n_issues  * review_min

    st.markdown(f"""
<div class="kpis" style="grid-template-columns:repeat(5,1fr)">
 <div class="kpi"><div class="lbl">Tickets in queue</div><div class="val">{n_tickets}</div>
   <div class="delta">all unassigned at shift start</div></div>
 <div class="kpi"><div class="lbl">Unique issues</div><div class="val">{n_issues}</div>
   <div class="delta">↓ {n_dupes} duplicates collapsed ({n_dupes / n_tickets:.0%})</div></div>
 <div class="kpi"><div class="lbl">Critical issues</div><div class="val">{n_critical}</div>
   <div class="delta">surfaced to the top automatically</div></div>
 <div class="kpi"><div class="lbl">Pending clarification</div><div class="val">{n_pending}</div>
   <div class="delta">⚠️ AI flagged — BA needs more info</div></div>
 <div class="kpi"><div class="lbl">Triage time (est.)</div>
   <div class="val">{after_min:.0f}<small> min</small></div>
   <div class="delta">vs {before_min / 60:.1f} hrs manual (−{1 - after_min / before_min:.0%})</div></div>
</div>
<p style="font-size:.75rem; color:#6B7684; margin-top:-.3rem;">
  ⚑ KPI estimates are illustrative — based on team time-study averages, not live measurements.
</p>""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Triage queue sub-tab (inner tab ①)
# ---------------------------------------------------------------------------
def _render_triage_queue(df: pd.DataFrame, issues: pd.DataFrame, queues: dict, groups: dict = None) -> None:
    c1, c2, c3 = st.columns([1.1, 2, 2])
    collapse = c1.toggle("Collapse duplicates", value=True)
    pf = c2.multiselect(
        "Priority",
        PRIORITY_ORDER,
        default=PRIORITY_ORDER,
    )
    qf = c3.multiselect("Suggested queue", list(queues), placeholder="All queues")

    view = issues.copy() if collapse else _flat_view(df)
    view = view[view.ai_priority.isin(pf)]
    if qf:
        view = view[view.ai_route_queue.isin(qf)]

    n_pending = int(view.get("ai_needs_more_info", pd.Series(dtype=bool)).sum()) if "ai_needs_more_info" in view.columns else 0
    decided   = len(st.session_state.decisions)
    st.caption(
        f"{len(view)} {'issues' if collapse else 'tickets'} · "
        f"sorted by AI priority score · {decided} routed · "
        f"{n_pending} pending clarification"
    )

    for _, r in view.iterrows():
        _render_issue_card(r, queues, df, groups or {})


def _flat_view(df: pd.DataFrame) -> pd.DataFrame:
    v = df.copy()
    v["key"]       = v["number"]
    v["is_group"]  = False
    v["n"]         = 1
    v["oldest_hrs"]= v["age_hrs"]
    v["n_request"] = v["ai_should_be_request"].astype(int)
    return v.sort_values("ai_priority_score", ascending=False)


def _render_issue_card(r, queues: dict, df: pd.DataFrame, groups: dict = None) -> None:
    k   = r["key"]
    dec = st.session_state.decisions.get(k)

    needs_info  = bool(r.get("ai_needs_more_info", False))
    info_reqs   = list(r.get("ai_info_requests",   []))

    with st.container(border=True):
        # ── Card header row ──────────────────────────────────────────────────
        a, b, c = st.columns([0.55, 5.2, 3.2], vertical_alignment="center")

        if needs_info:
            a.markdown(
                '<div style="width:58px;height:58px;border-radius:12px;'
                'background:#FFF8E6;border:1.5px solid #F5D88A;'
                'display:flex;flex-direction:column;align-items:center;'
                'justify-content:center;font-size:1.5rem;">⚠️</div>',
                unsafe_allow_html=True,
            )
        else:
            a.markdown(score_box(r.ai_priority_score, r.ai_priority), unsafe_allow_html=True)

        badges = pill(r.ai_category, "#3A4552", "#EEF1F4")
        if needs_info:
            badges += pill("⚠️ Needs More Info", "#7A5E00", "#FFF8E6")
        if r.is_group:
            badges += pill(f"🔗 {r.n} tickets grouped", "#0F6E6E", "#E2F2F1")
        if r.get("n_request", 0) or r.ai_should_be_request:
            badges += pill("↻ Should be a Request", "#7A4B00", "#FFF1D6")
        if r.ai_route_queue == "IT Service Desk (not DE)":
            badges += pill("✕ Not Data Engineering", "#7A1F1D", "#FCE9E8")
        if not r.is_group and str(r.snow_priority).startswith("4") and r.ai_priority_score >= 70:
            badges += pill("⚑ Filed as Low", "#7A1F1D", "#FCE9E8")

        b.markdown(
            f'<div class="summ">{html.escape(str(r.ai_summary))}</div>'
            f'<div class="orig">{html.escape(str(r.number))} · '
            f'"{html.escape(str(r.short_description))}" · '
            f'{html.escape(str(r.caller_dept))} · open {age(r.oldest_hrs)}</div>'
            f'{badges}',
            unsafe_allow_html=True,
        )

        queue = dec["queue"] if dec else r.ai_route_queue
        owner = queues.get(queue, "—") if (dec and dec["queue"] != r.ai_route_queue) else r.ai_route_owner
        status_label = {"accepted": "✅ Routed", "rerouted": "✏️ Rerouted"}.get(
            dec["status"] if dec else "", "Suggested"
        )
        c.markdown(
            f'<div class="route"><span class="muted">{status_label} →</span><br>'
            f'<b>{html.escape(str(queue))}</b><br>👤 {html.escape(str(owner))}</div>'
            f'<div class="conf"><div style="width:{r.ai_confidence * 100:.0f}%"></div></div>'
            f'<div class="muted">confidence {r.ai_confidence:.0%}</div>',
            unsafe_allow_html=True,
        )

        # ── Needs More Info panel (shown outside the expander for visibility) ─
        if needs_info and info_reqs:
            st.markdown(
                '<div class="needs-info">⚠️ <b>AI flagged — more information needed before scoring.</b><br>'
                'The Business Analyst should ask the requester:<ul>'
                + "".join(f"<li>{html.escape(q)}</li>" for q in info_reqs)
                + "</ul>Once the BA provides the answers, the agent will re-triage with a full ERWG score.</div>",
                unsafe_allow_html=True,
            )

        # ── Score breakdown expander ─────────────────────────────────────────
        expander_label = (
            "⚠️ Needs More Info — no score yet · details"
            if needs_info
            else f"Why this score ({r.ai_priority_score}) & route · ERWG breakdown · details"
        )
        with st.expander(expander_label):
            d1, d2 = st.columns([3, 2])
            with d1:
                if needs_info:
                    st.info(
                        "This ticket was flagged as incomplete. The ERWG score will be calculated "
                        "once the BA obtains the missing information from the requester.",
                        icon="⚠️",
                    )
                else:
                    # ERWG 5-row breakdown table
                    st.markdown("**ERWG Priority Score Breakdown**")
                    result    = score_ticket(r.to_dict(), _SCORING_CFG)
                    breakdown = result["breakdown"]

                    # Build display table
                    bd_rows = []
                    for row in breakdown:
                        if row["key"] == "org_impact":
                            bd_rows.append({
                                "Dimension":      row["label"],
                                "Score (0–4)":    row["score"],
                                "Weight":         row["weight_label"],
                                "Multiplier":     row.get("multiplier", ""),
                                "Level":          row["level_label"],
                            })
                        else:
                            bd_rows.append({
                                "Dimension":      row["label"],
                                "Score (0–4)":    row["score"],
                                "Weight":         row["weight_label"],
                                "Points contrib": row["points"],
                                "Level":          row["level_label"],
                            })

                    # Org impact row first, then the four additive rows
                    org_row   = [bd_rows[0]]
                    addl_rows = bd_rows[1:]
                    bd_df = pd.DataFrame(addl_rows)[["Dimension", "Score (0–4)", "Weight", "Points contrib", "Level"]]
                    st.dataframe(bd_df, hide_index=True, use_container_width=True)

                    org_info = bd_rows[0]
                    st.markdown(
                        f"**Organizational Impact (multiplier):** {org_info['Dimension']} — "
                        f"score {org_info['Score (0–4)']} → **{org_info.get('Multiplier','')} of raw score applied**  \n"
                        f"_{org_info['Level']}_  \n\n"
                        f"**Final ERWG score: {r.ai_priority_score}** "
                        f"→ {priority_pill(r.ai_priority)}",
                        unsafe_allow_html=True,
                    )

                    st.markdown(f"**Rationale:** {r.ai_rationale}")
                    if r.ai_similar_history:
                        st.markdown(f"**Who's handled this before:** {r.ai_similar_history}")
                    if r.is_group and groups:
                        g = groups.get(r["key"], {})
                        if g:
                            st.markdown(f"**Likely root cause:** {g.get('root_cause', '')}")
                            st.markdown(f"**Recommended action:** {g.get('recommended_action', '')}")
                        members = df[df.number.isin(r.members)][
                            ["number", "type", "channel", "short_description", "caller", "ai_priority_score"]
                        ].sort_values("ai_priority_score", ascending=False)
                        st.dataframe(members, hide_index=True, use_container_width=True)
                    else:
                        st.markdown("**Original description**")
                        st.code(r.description, language=None, wrap_lines=True)

                    # Form field context
                    jtext = str(r.get("justification_text", "") or "")
                    if jtext:
                        with st.expander("ServiceNow form fields (submitted by requester)"):
                            st.markdown(f"**Justification:** {jtext}")
                            st.markdown(
                                f"**Needed by:** {r.get('when_needed','—')} · "
                                f"**Recurring:** {'Yes' if r.get('is_recurring') else 'No'} · "
                                f"**Frequency:** {r.get('frequency','—')} · "
                                f"**Analytics type:** {r.get('analytics_type','—')}"
                            )

            with d2:
                st.markdown(f"**Scored by:** {r.scored_by}")
                st.markdown(f"**Filed as:** {r.type} · {r.snow_priority} via {r.channel}")
                st.markdown(f"**Caller:** {r.caller} ({r.caller_dept})")
                x1, x2 = st.columns(2)
                if not needs_info:
                    if x1.button("✅ Accept route", key=f"acc_{k}", use_container_width=True):
                        st.session_state.decisions[k] = {"status": "accepted", "queue": r.ai_route_queue}
                        st.rerun()
                new_q = x2.selectbox(
                    "Reroute to", list(queues), key=f"rq_{k}",
                    index=list(queues).index(queue) if queue in queues else 0,
                    label_visibility="collapsed",
                )
                if new_q != queue and not needs_info:
                    st.session_state.decisions[k] = {"status": "rerouted", "queue": new_q}
                    st.rerun()


# ---------------------------------------------------------------------------
# Duplicate groups sub-tab (inner tab ②)
# ---------------------------------------------------------------------------
def _render_groups(df: pd.DataFrame, issues: pd.DataFrame, groups: dict) -> None:
    n_tickets = len(df)
    n_issues  = len(issues)
    n_dupes   = n_tickets - n_issues
    st.markdown(
        f"**{n_tickets} tickets → {n_issues} issues.** "
        f"{len(groups)} groups absorb {sum(len(df[df.group_id == g]) for g in groups)} tickets. "
        "Fix the root once, link the rest, send one update to everyone affected."
    )
    gsum = issues[issues.is_group].sort_values("ai_priority_score", ascending=False)
    for _, r in gsum.iterrows():
        gid     = r["key"]
        g       = groups[gid]
        members = df[df.group_id == gid].sort_values("opened_dt")
        lbl     = r.ai_priority
        with st.container(border=True):
            h1, h2 = st.columns([6, 2], vertical_alignment="center")
            h1.markdown(
                f'<div class="gcard-h">{priority_pill(lbl)} {html.escape(g["title"])}</div>'
                f'<div class="muted">{gid} · {len(members)} tickets · '
                f'{(members.channel == "Monitoring").sum()} automated alerts · '
                f'{members[members.caller_dept != "Automated"].caller_dept.nunique()} affected teams · '
                f'first seen {age(members.age_hrs.max())} ago</div>',
                unsafe_allow_html=True,
            )
            h2.markdown(
                f'<div class="route" style="text-align:right">→ '
                f'<b>{html.escape(r.ai_route_queue)}</b>'
                f'<br>👤 {html.escape(r.ai_route_owner)}</div>',
                unsafe_allow_html=True,
            )
            l, rcol = st.columns([3, 2])
            with l:
                st.markdown(f"**Root cause:** {g['root_cause']}")
                st.markdown(f"**Do this:** {g['recommended_action']}")
                tbl = members[["number", "group_role", "channel", "short_description", "caller", "age_hrs"]].copy()
                tbl["age_hrs"]    = tbl.age_hrs.map(age)
                tbl["group_role"] = tbl.group_role.map({"parent": "⭐ parent", "child": "child"})
                tbl.columns = ["Number", "Role", "Channel", "Original short description", "Caller", "Age"]
                st.dataframe(tbl, hide_index=True, use_container_width=True)
            with rcol:
                st.markdown("**Update for:** " + ", ".join(g["notify"]))
                st.text_area("Draft", g["comms_draft"], height=190, key=f"comms_{gid}", label_visibility="collapsed")
                b1, b2 = st.columns(2)
                linked = gid in st.session_state.linked
                if b1.button(
                    "✓ Children linked" if linked else f"🔗 Link {len(members) - 1} to parent",
                    key=f"link_{gid}", disabled=linked, use_container_width=True,
                ):
                    st.session_state.linked.add(gid)
                    st.toast(f"Linked {len(members) - 1} tickets to {r.number} (demo)")
                    st.rerun()
                sent = gid in st.session_state.notified
                if b2.button(
                    "✓ Update sent" if sent else "✉️ Send update",
                    key=f"send_{gid}", disabled=sent, use_container_width=True,
                ):
                    st.session_state.notified.add(gid)
                    st.toast("Update posted to linked tickets (demo — nothing actually sent)")
                    st.rerun()


# ---------------------------------------------------------------------------
# Routing sub-tab (inner tab ③)
# ---------------------------------------------------------------------------
def _render_routing(df: pd.DataFrame, issues: pd.DataFrame, queues: dict) -> None:
    st.markdown(
        "**Before:** every ticket lands in one shared *Data Engineering* queue. "
        "**After:** each issue goes straight to the team that owns it, with a suggested owner "
        "based on who has fixed similar tickets."
    )
    eff_queue = issues.apply(
        lambda r: st.session_state.decisions.get(r["key"], {}).get("queue", r.ai_route_queue),
        axis=1,
    )
    counts = pd.DataFrame({
        "Tickets": df.ai_route_queue.value_counts(),
        "Unique issues": eff_queue.value_counts(),
    }).fillna(0).astype(int).reindex(list(queues))
    counts.index = [q.split("–", 1)[-1].strip().replace(" (not DE)", "") for q in counts.index]
    m1, m2 = st.columns([3, 2])
    with m1:
        st.bar_chart(counts, horizontal=True, stack=False, height=320, color=["#9FB4C2", "#0F6E6E"])
    with m2:
        acc = sum(1 for d in st.session_state.decisions.values() if d["status"] == "accepted")
        rer = sum(1 for d in st.session_state.decisions.values() if d["status"] == "rerouted")
        st.metric("Routing decisions made", f"{acc + rer} / {len(issues)}")
        st.metric("AI route accepted as-is", f"{acc / (acc + rer):.0%}" if acc + rer else "—",
                  help="Live proxy for the routing-accuracy KPI.")
        st.metric("Sent back to IT Service Desk",
                  int((df.ai_route_queue == "IT Service Desk (not DE)").sum()))

    st.divider()
    cols = st.columns(3)
    for i, q in enumerate(queues):
        sub = issues[eff_queue == q]
        with cols[i % 3].container(border=True):
            st.markdown(
                f"**{q}**  \n<span class='muted'>👤 {queues[q]} · {len(sub)} issues</span>",
                unsafe_allow_html=True,
            )
            for _, r in sub.iterrows():
                lbl = r.ai_priority
                tag = f" · {r.n} tickets" if r.is_group else ""
                st.markdown(
                    f"{priority_pill(lbl)} {html.escape(r.ai_summary[:70])}"
                    f"{'…' if len(r.ai_summary) > 70 else ''}"
                    f"<span class='muted'>{tag}</span>",
                    unsafe_allow_html=True,
                )


# ---------------------------------------------------------------------------
# Live triage sub-tab (inner tab ④)
# ---------------------------------------------------------------------------
def _render_live_triage(df: pd.DataFrame, queues: dict) -> None:
    from triage.offline_engine import SimilarityIndex, triage as _triage

    SAMPLES = {
        "Duplicate of an open issue (bed census)":
            "Hi team, the bed census report on 6 North still shows yesterday's numbers. We need it for the bed huddle. Thanks!",
        "Access request filed as incident":
            "Can I get access to the Bed Capacity dashboard in Tableau? I get permission denied.",
        "Not Data Engineering":
            "My laptop VPN keeps disconnecting and I'm locked out of email.",
        "Possible PHI exposure":
            "A workbook in the All Staff project shows patient names and MRN numbers — I think everyone can see it.",
        "New pipeline failure":
            "Talend job RADIOLOGY_ORDERS_DAILY failed overnight with a null pointer in tMap_2. Radiology volume report is blank.",
    }

    st.markdown(
        "Paste or pick a ticket the way it would arrive from email or the portal. "
        "The demo engine summarizes, scores, routes, and checks it against the open queue for duplicates."
    )
    s    = st.selectbox("Sample ticket", list(SAMPLES))
    text = st.text_area("Ticket text", SAMPLES[s], height=110, key=f"txt_{s}")

    with st.expander("ServiceNow form fields (requester provides these at submission)"):
        col_a, col_b, col_c = st.columns(3)
        justification = col_a.text_area(
            "Justify the priority level",
            "Nightly pipeline failure causing stale ED boarding data.",
            height=80,
        )
        when_needed   = col_b.text_input("When is this needed by?", "2026-10-15")
        frequency     = col_c.selectbox("Frequency", ["Daily", "Weekly", "Monthly", "Ad-hoc"], index=0)
        is_recurring  = col_a.checkbox("Recurring report or pipeline?", value=True)
        analytics_type= col_b.selectbox("Analytics type", ["EA", "Surgery", "OPT", "Unknown/Other"], index=0)

    if st.button("⚡ Score & route this ticket", type="primary"):
        idx = SimilarityIndex(df.to_dict("records"))
        res = _triage(text, idx, queues)
        pl  = plabel_for(res["score"])

        with st.container(border=True):
            a, b, c = st.columns([0.55, 5.2, 3.2], vertical_alignment="center")
            a.markdown(score_box(res["score"], pl), unsafe_allow_html=True)
            badges = pill(res["category"], "#3A4552", "#EEF1F4")
            gm = res["group_match"]
            if gm is not None:
                badges += pill(f"🔗 Duplicate → {gm['group_id'] or gm['number']}", "#0F6E6E", "#E2F2F1")
            if res["should_be_request"]:
                badges += pill("↻ Should be a Request", "#7A4B00", "#FFF1D6")
            if res["queue"] == "IT Service Desk (not DE)":
                badges += pill("✕ Not Data Engineering", "#7A1F1D", "#FCE9E8")
            b.markdown(f'<div class="summ">{html.escape(res["summary"])}</div>{badges}', unsafe_allow_html=True)
            c.markdown(
                f'<div class="route"><span class="muted">Suggested →</span><br>'
                f'<b>{html.escape(res["queue"])}</b><br>👤 {html.escape(res["owner"])}</div>'
                f'<div class="conf"><div style="width:{res["confidence"] * 100:.0f}%"></div></div>'
                f'<div class="muted">confidence {res["confidence"]:.0%}</div>',
                unsafe_allow_html=True,
            )
            st.markdown(f"**Why:** {res['rationale']}")
            if gm is not None and gm.get("group_id"):
                st.success(
                    "Would be linked to an existing group — no new investigation needed. "
                    "Caller gets the existing status update automatically."
                )
            elif gm is not None:
                st.warning(f"Same issue as open ticket **{gm['number']}** — would start a new duplicate group.")

        st.markdown("**Most similar open tickets**")
        st.dataframe(pd.DataFrame([{
            "Similarity": round(sc, 2), "Number": t["number"],
            "Short description": t["short_description"], "Group": t["group_id"] or "—",
        } for sc, t in res["similar"]]), hide_index=True, use_container_width=True)

    st.caption(
        "Offline demo mode: keyword routing + TF-IDF similarity. "
        "In production this step is an LLM call trained on past scored requests."
    )


# ---------------------------------------------------------------------------
# Data table sub-tab (inner tab ⑤)
# ---------------------------------------------------------------------------
def _render_data_table(df: pd.DataFrame) -> None:
    st.markdown(
        "The single queryable table developers asked for — every ticket with its ERWG score, "
        "summary, owner and status. In production this lands in Redshift so the team can query it directly."
    )
    cols = [
        "number", "type", "ai_priority_score", "ai_priority",
        "ai_org_impact", "ai_patient_safety", "ai_compliance", "ai_financial", "ai_utilization",
        "ai_needs_more_info", "ai_summary", "ai_category",
        "ai_route_queue", "ai_route_owner", "ai_confidence", "group_id",
        "ai_should_be_request", "opened_at", "state", "scored_by",
    ]
    available = [c for c in cols if c in df.columns]
    out = df[available].sort_values("ai_priority_score", ascending=False)
    st.dataframe(out, hide_index=True, use_container_width=True, height=440, column_config={
        "ai_priority_score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%d"),
        "ai_confidence":     st.column_config.NumberColumn("Confidence", format="percent"),
    })
    st.download_button("⬇️ Download CSV", out.to_csv(index=False), "ticket_triage.csv", "text/csv")
    with st.expander("Proposed Redshift table"):
        st.code("""CREATE TABLE de_triage.ticket_scores (
    ticket_number        VARCHAR(20)  NOT NULL,
    ticket_type          VARCHAR(20),
    priority_score       SMALLINT,
    priority_label       VARCHAR(20),   -- Critical / High / Medium / Low
    erwg_org_impact      SMALLINT,      -- 0-4
    erwg_patient_safety  SMALLINT,      -- 0-4
    erwg_compliance      SMALLINT,      -- 0-4
    erwg_financial       SMALLINT,      -- 0-4
    erwg_utilization     SMALLINT,      -- 0-4
    needs_more_info      BOOLEAN,
    ai_summary           VARCHAR(500),
    category             VARCHAR(50),
    suggested_queue      VARCHAR(100),
    suggested_owner      VARCHAR(100),
    confidence           DECIMAL(3,2),
    duplicate_group_id   VARCHAR(20),
    should_be_request    BOOLEAN,
    opened_at            TIMESTAMP,
    status               VARCHAR(20),
    scored_by            VARCHAR(50),
    scored_at            TIMESTAMP DEFAULT GETDATE()
)
DISTKEY(ticket_number) SORTKEY(priority_score, opened_at);""", language="sql")


# ---------------------------------------------------------------------------
# Main render function
# ---------------------------------------------------------------------------
def render_after(data: dict, df: pd.DataFrame, now, manual_min: float = 6, review_min: float = 1.5) -> None:
    """
    Render the After (AI Triage Agent) tab.

    Args:
        data:        Full data dict from tickets.json (includes queues and groups).
        df:          Ticket DataFrame.
        now:         Demo reference datetime.
        manual_min:  Slider value for manual triage time per ticket.
        review_min:  Slider value for AI review time per issue.
    """
    _init_state()
    queues = data["queues"]
    groups = data["groups"]
    issues = build_issues(df, groups)

    # ── Top bar: Run Triage / Reset ──────────────────────────────────────────
    col_run, col_reset, col_info = st.columns([1, 1, 6])

    if col_run.button("▶ Run Triage", type="primary", use_container_width=True,
                      disabled=st.session_state.triage_done):
        _run_triage_animation(len(df))

    if col_reset.button("↺ Reset", use_container_width=True):
        st.session_state.triage_done  = False
        st.session_state.decisions    = {}
        st.session_state.linked       = set()
        st.session_state.notified     = set()
        st.rerun()

    if not st.session_state.triage_done:
        col_info.markdown(
            '<span style="color:#6B7684; font-size:.9rem;">'
            '← Click <b>Run Triage</b> to watch the four agents work.</span>',
            unsafe_allow_html=True,
        )
        st.info(
            "The AI triage agents will run here.  \n"
            "You'll see each agent step play out before the results appear.",
            icon="🤖",
        )
        st.caption("Synthetic data only — no PHI.")
        return

    # ── KPI panel ────────────────────────────────────────────────────────────
    _render_kpis(df, issues, manual_min, review_min)

    # ── Clock line ───────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="hdr">'
        f'<div class="sub">AI-assisted triage — summarize, group, score (ERWG), route.</div>'
        f'<div class="clock">🕘 {now:%a %b %d}, {now.hour % 12 or 12}:{now:%M %p} · on-call shift start · queue: <b>Data Engineering</b></div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── Inner tabs ────────────────────────────────────────────────────────────
    inner = st.tabs([
        "① Triage queue",
        "② Duplicate groups",
        "③ Routing",
        "④ Triage a new ticket",
        "⑤ Data table",
    ])

    with inner[0]:
        _render_triage_queue(df, issues, queues, groups)

    with inner[1]:
        _render_groups(df, issues, groups)

    with inner[2]:
        _render_routing(df, issues, queues)

    with inner[3]:
        _render_live_triage(df, queues)

    with inner[4]:
        _render_data_table(df)

    st.caption("Synthetic data only — no PHI, no real Keck tickets.")
