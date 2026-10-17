"""
Generates the dummy ServiceNow tickets used by the demo.

Every ticket has two halves:
  * the raw ServiceNow fields, as the ticket lands in the shared Data Engineering queue
  * the ai_* fields, which are what the triage model WOULD return (precomputed so the
    demo runs with no API key and no network)

All people, ticket numbers and details are fictional.

Run:  python data/generate_tickets.py   -> writes data/tickets.json and data/tickets.csv
"""
import csv
import json
from datetime import datetime, timedelta
from pathlib import Path

DEMO_NOW = datetime(2026, 10, 15, 9, 0)  # Thursday, 9:00 AM — on-call shift start
SCORED_BY = "Triage AI v0.3 (precomputed)"

# ---------------------------------------------------------------------------
# Routing targets: queue -> default owner
# ---------------------------------------------------------------------------
QUEUES = {
    "DE – Pipelines (Talend/ETL)": "Dev Patel",
    "DE – Redshift / Warehouse": "Luis Ortega",
    "DE – Clinical Data (Cerner/OMOP)": "Sam Whitfield",
    "BI – Reporting (Tableau)": "Priya Raman",
    "DE – Data Access & Security": "Jordan Lee",
    "IT Service Desk (not DE)": "Service Desk Tier 1",
}

# ---------------------------------------------------------------------------
# Duplicate / similar-issue groups
# ---------------------------------------------------------------------------
GROUPS = {
    "GRP-01": {
        "title": "Nightly Cerner ADT load failed → bed census & ED reports stale",
        "root_cause": "Talend job CERNER_ADT_DAILY failed at 02:14 on a Redshift connection timeout and both retries failed. Every downstream bed/ED report is showing yesterday's data.",
        "recommended_action": "Fix and re-run CERNER_ADT_DAILY once, close the 3 monitoring alerts as children, and send one status update to the 3 affected report users.",
        "notify": ["Nursing Operations", "Capacity Management", "Emergency Department"],
        "comms_draft": "Subject: Bed census & ED boarding reports delayed this morning\n\nHi all — the overnight load that feeds the Bed Census, Bed Availability and ED Boarding reports failed at 2:14 AM. Data Engineering is re-running it now; we expect reports to show today's numbers by 10:30 AM. You don't need to open a new ticket — we've linked yours to INC0104511 and will update you here when it's done.\n\n— Data Engineering On-Call",
    },
    "GRP-02": {
        "title": "Redshift queries timing out (WLM queue saturated)",
        "root_cause": "CloudWatch shows the default WLM queue waiting >300s since ~3 PM yesterday. A long-running ad-hoc research query is holding most of the slots.",
        "recommended_action": "Kill / move the long-running query to a research WLM queue, then confirm the Tableau extract refresh succeeds. Link the 3 user tickets to the alarm.",
        "notify": ["Quality & Analytics", "BI Reporting"],
        "comms_draft": "Subject: Redshift slowness — cause found\n\nHi — the slowdowns and timeouts you reported are caused by one long-running query that's tying up shared capacity. We're moving it to a separate queue now; queries and Tableau refreshes should return to normal within the hour. Your ticket is linked to INC0104490.\n\n— Data Engineering",
    },
    "GRP-03": {
        "title": "Access requests for the Bed Capacity Tableau dashboard",
        "root_cause": "Five managers emailed the help desk for dashboard access, which auto-opened Incidents. These are access requests, not outages.",
        "recommended_action": "Convert all 5 to a single access Request (RITM), approve against the Bed Capacity AD group, and close the Incidents before they breach SLA.",
        "notify": ["Requesters (5 managers)"],
        "comms_draft": "Subject: Your Bed Capacity dashboard access\n\nHi — thanks for reaching out. Your note came in as an incident, so we've converted it to an access request to get it approved faster. Once your manager approves (usually same day), you'll be able to open the dashboard. For next time, the 'Request Tableau Access' form in the portal goes straight to the right team.\n\n— Data Engineering",
    },
    "GRP-04": {
        "title": "Lab results ingestion Lambda timing out",
        "root_cause": "lab-results-ingest Lambda hit its 900s limit twice; S3 landing prefix shows ~3x normal object count, likely a vendor backfill dump.",
        "recommended_action": "Confirm the vendor backfill, raise Lambda concurrency / batch size temporarily, and treat the 3 alerts as one incident.",
        "notify": ["Laboratory Informatics"],
        "comms_draft": "Subject: Lab results feed running behind\n\nHi — the lab results feed into the warehouse is behind because of a large vendor file drop overnight. We're processing it now; lab-based reports may be up to 4 hours behind until about noon.\n\n— Data Engineering",
    },
    "GRP-05": {
        "title": "OMOP vocabulary is out of date",
        "root_cause": "The OMOP concept tables are on the March vocabulary release; newer SNOMED codes are missing, which is what the researcher is seeing.",
        "recommended_action": "Fulfil RITM0031207 (load latest Athena vocabulary) and close INC0104505 as resolved-by-request.",
        "notify": ["Research Informatics"],
        "comms_draft": "Subject: Missing SNOMED codes in OMOP\n\nHi — the missing codes are because our OMOP vocabulary is due for a refresh. That's already requested under RITM0031207 and scheduled for this sprint; we've linked your ticket to it.\n\n— Data Engineering",
    },
    "GRP-06": {
        "title": "Login / password issues sent to Data Engineering",
        "root_cause": "Account lockout and password-expiry tickets that belong to the IT Service Desk, not Data Engineering.",
        "recommended_action": "Reassign both to the IT Service Desk queue immediately.",
        "notify": ["IT Service Desk"],
        "comms_draft": "Subject: Your login issue\n\nHi — login and password problems are handled by the IT Service Desk, so we've moved your ticket there. They'll reach out shortly. If it's urgent, call the Service Desk line.\n\n— Data Engineering",
    },
}

# ---------------------------------------------------------------------------
# Tickets
#   hrs_ago: how long before DEMO_NOW the ticket was opened
#   filed_pri: priority as it landed in ServiceNow (email defaults to 4 - Low)
#   score: AI priority score 0-100
# ---------------------------------------------------------------------------
T = []


def t(number, ttype, channel, hrs_ago, caller, dept, filed_pri, short, desc,
      summary, category, score, queue, confidence, rationale,
      group=None, role=None, should_be_request=False, owner=None, history=None):
    T.append(dict(
        number=number, type=ttype, channel=channel,
        opened_at=(DEMO_NOW - timedelta(hours=hrs_ago)).strftime("%Y-%m-%d %H:%M"),
        caller=caller, caller_dept=dept,
        snow_priority=filed_pri, state="New",
        assignment_group="Data Engineering", assigned_to="",
        short_description=short, description=desc,
        ai_summary=summary, ai_category=category, ai_priority_score=score,
        ai_route_queue=queue, ai_route_owner=owner or QUEUES[queue],
        ai_confidence=confidence, ai_rationale=rationale,
        ai_should_be_request=should_be_request,
        ai_similar_history=history or "",
        group_id=group or "", group_role=role or "",
        scored_by=SCORED_BY,
    ))


# ---------------- GRP-01: Cerner ADT load failure ----------------
t("INC0104511", "Incident", "Monitoring", 6.8, "talend-monitor", "Automated", "3 - Moderate",
  "Talend job CERNER_ADT_DAILY failed - exit code 1",
  "Job: CERNER_ADT_DAILY\nServer: talend-prd-02\nStatus: FAILED (exit 1)\nError: java.sql.SQLException: [Amazon](500150) Error setting/closing connection: Connection timed out.\nStep: tRedshiftOutput_3 (stg_adt_events)",
  "Nightly Cerner ADT load failed on a Redshift connection timeout; all ADT-based reports are now stale.",
  "Pipeline failure", 94, "DE – Pipelines (Talend/ETL)", 0.95,
  "Root job for 5 downstream tickets. ADT feeds bed census and ED boarding — clinical operations impact during morning bed huddle.",
  "GRP-01", "parent",
  history="INC0098812 (Aug 21) — same job, same timeout; fixed by Dev Patel (connection pool size).")
t("INC0104512", "Incident", "Monitoring", 6.3, "talend-monitor", "Automated", "3 - Moderate",
  "Talend job CERNER_ADT_DAILY retry 1 failed",
  "Job: CERNER_ADT_DAILY (auto-retry 1/2)\nStatus: FAILED (exit 1)\nError: Connection timed out.",
  "Automatic retry #1 of the ADT load failed with the same error.",
  "Pipeline failure", 88, "DE – Pipelines (Talend/ETL)", 0.97,
  "Same job and error as INC0104511 — duplicate alert.", "GRP-01", "child")
t("INC0104513", "Incident", "Monitoring", 5.8, "talend-monitor", "Automated", "3 - Moderate",
  "Talend job CERNER_ADT_DAILY retry 2 failed",
  "Job: CERNER_ADT_DAILY (auto-retry 2/2)\nStatus: FAILED (exit 1)\nError: Connection timed out. No further retries scheduled.",
  "Final automatic retry of the ADT load failed; no more retries are scheduled.",
  "Pipeline failure", 88, "DE – Pipelines (Talend/ETL)", 0.97,
  "Same job and error as INC0104511 — duplicate alert.", "GRP-01", "child")
t("INC0104520", "Incident", "Email", 2.2, "Angela Brooks", "Nursing Operations", "4 - Low",
  "Bed census dashboard shows yesterday's numbers",
  "Good morning, the bed census dashboard is still showing yesterday's census. We use this for the 8am bed huddle. Can someone take a look? Thanks, Angela",
  "Bed census dashboard is showing yesterday's data; needed for the 8 AM bed huddle.",
  "Stale data", 86, "DE – Pipelines (Talend/ETL)", 0.89,
  "Stale bed census matches the failed ADT load (GRP-01). Filed as Low via email but affects daily bed huddle.",
  "GRP-01", "child")
t("INC0104524", "Incident", "Portal", 1.6, "Marcus Hill", "Capacity Management", "3 - Moderate",
  "Bed availability report not updated this morning",
  "The Bed Availability report last refresh time says 10/14 11:58 PM. Transfer center is asking for current numbers.",
  "Bed availability report hasn't refreshed since last night; transfer center needs current numbers.",
  "Stale data", 87, "DE – Pipelines (Talend/ETL)", 0.9,
  "Last refresh time lines up with the failed ADT load (GRP-01).", "GRP-01", "child")
t("INC0104529", "Incident", "Email", 0.7, "Teresa Nguyen", "Emergency Department", "4 - Low",
  "ED boarding report blank for today",
  "Hi, ED boarding report has no rows for today. Is something down?",
  "ED boarding report has no rows for today.",
  "Stale data", 87, "DE – Pipelines (Talend/ETL)", 0.86,
  "ED boarding report is built from ADT events — same root cause as GRP-01.", "GRP-01", "child")

# ---------------- GRP-02: Redshift WLM saturation ----------------
t("INC0104490", "Incident", "Monitoring", 17.5, "cloudwatch", "Automated", "3 - Moderate",
  "CloudWatch alarm: Redshift WLM queue wait > 300s",
  "ALARM: redshift-prd-wlm-queue-wait\nCluster: kmed-dw-prd\nMetric: WLMQueueWaitTime (default queue) > 300000 ms for 3 datapoints",
  "Redshift default query queue has been backed up for 3+ intervals; queries are waiting over 5 minutes.",
  "Performance", 78, "DE – Redshift / Warehouse", 0.93,
  "Platform-level alarm explaining 3 user-reported slowness tickets.", "GRP-02", "parent",
  history="INC0101307 (Sep 30) — same alarm; Luis Ortega moved research workload to separate WLM queue.")
t("INC0104470", "Incident", "Portal", 18.4, "Rachel Kim", "Quality & Analytics", "3 - Moderate",
  "Queries timing out in Redshift since yesterday afternoon",
  "Simple SELECTs on the encounters table are taking 5+ minutes or timing out in DBeaver. Started around 3pm.",
  "Analyst queries on Redshift have been timing out since ~3 PM yesterday.",
  "Performance", 72, "DE – Redshift / Warehouse", 0.88,
  "Timing matches the WLM alarm (GRP-02).", "GRP-02", "child")
t("INC0104478", "Incident", "Email", 15.1, "Kevin Osei", "BI Reporting", "4 - Low",
  "Tableau extract refresh failing - timeout",
  "Our Readmissions extract refresh failed twice overnight with 'query timed out'. Workbook: Readmissions_30day.",
  "Readmissions Tableau extract failed twice overnight due to query timeouts.",
  "Performance", 74, "DE – Redshift / Warehouse", 0.84,
  "Extract failure is a symptom of the Redshift queue backlog (GRP-02), not a Tableau problem.", "GRP-02", "child")
t("INC0104497", "Incident", "Email", 12.0, "Dr. Alan Fisher", "Research", "4 - Low",
  "SQL Workbench very slow, can't run my readmissions query",
  "My query has been running for 2 hours. Is the server down?",
  "Researcher's query has been running for 2 hours — likely the one saturating the queue.",
  "Performance", 70, "DE – Redshift / Warehouse", 0.8,
  "Same symptom as GRP-02; this caller's long-running query may be the cause.", "GRP-02", "child")

# ---------------- GRP-03: Tableau access (should be requests) ----------------
for num, hrs, caller, dept, short, desc in [
    ("INC0104455", 26.0, "Linda Park", "Nursing - 5 North", "Can't open Bed Capacity dashboard",
     "I get 'You do not have permission to access this view' when I click the Bed Capacity link."),
    ("INC0104461", 22.5, "Omar Haddad", "Perioperative Services", "Need access to bed dashboard",
     "Please give me access to the bed capacity dashboard, my director asked me to review it daily."),
    ("INC0104483", 14.0, "Grace Liu", "Case Management", "Permission denied on Tableau",
     "Permission denied when opening Tableau > Operations > Bed Capacity."),
    ("INC0104502", 9.5, "Brian Torres", "Nursing - 7 West", "Bed capacity access",
     "Hi can I get access to the bed capacity report? Thanks"),
    ("INC0104527", 1.2, "Nadia Rahman", "Patient Flow", "Access to Bed Capacity Tableau",
     "New to the Patient Flow team. Need Bed Capacity dashboard access."),
]:
    t(num, "Incident", "Email", hrs, caller, dept, "4 - Low", short, desc,
      f"{caller.split()[0]} needs view access to the Bed Capacity Tableau dashboard.",
      "Access request", 32, "DE – Data Access & Security", 0.94,
      "Access provisioning, not break-fix. Should be a Request (RITM); convert before SLA breach.",
      "GRP-03", "parent" if num == "INC0104455" else "child", should_be_request=True)

# ---------------- GRP-04: Lambda lab ingestion ----------------
t("INC0104440", "Incident", "Monitoring", 30.0, "cloudwatch", "Automated", "3 - Moderate",
  "Lambda lab-results-ingest timeout (900s)",
  "Function: lab-results-ingest\nError: Task timed out after 900.00 seconds\nInvocations failed: 14",
  "Lab results ingestion Lambda is timing out; 14 invocations failed.",
  "Pipeline failure", 73, "DE – Pipelines (Talend/ETL)", 0.9,
  "Feeds lab reporting. Root alert for 2 related alarms.", "GRP-04", "parent")
t("INC0104448", "Incident", "Monitoring", 25.0, "cloudwatch", "Automated", "3 - Moderate",
  "Lambda lab-results-ingest timeout (900s)",
  "Function: lab-results-ingest\nError: Task timed out after 900.00 seconds\nInvocations failed: 9",
  "Repeat of the lab ingestion Lambda timeout; 9 more failures.",
  "Pipeline failure", 70, "DE – Pipelines (Talend/ETL)", 0.96,
  "Identical alert text to INC0104440.", "GRP-04", "child")
t("INC0104466", "Incident", "Monitoring", 20.0, "cloudwatch", "Automated", "4 - Low",
  "S3 prefix landing/labs object count anomaly",
  "Anomaly detector: s3://kmed-landing/labs/ object count 3.1x 7-day average.",
  "Lab landing folder received ~3x the normal number of files — likely why the Lambda is timing out.",
  "Pipeline failure", 66, "DE – Pipelines (Talend/ETL)", 0.78,
  "Probable cause of GRP-04 timeouts (vendor backfill).", "GRP-04", "child")

# ---------------- GRP-05: OMOP vocab ----------------
t("RITM0031207", "Request", "Portal", 70.0, "Dr. Hannah Cole", "Research Informatics", "4 - Low",
  "Load latest Athena vocabulary into OMOP CDM",
  "Please refresh the OMOP vocabulary tables to the latest Athena release (Aug 2026). We need new SNOMED and LOINC concepts for an IRB study.",
  "Research wants the OMOP vocabulary refreshed to the Aug 2026 Athena release for an IRB study.",
  "Data change request", 46, "DE – Clinical Data (Cerner/OMOP)", 0.92,
  "Planned work; resolves INC0104505 as well.", "GRP-05", "parent")
t("INC0104505", "Incident", "Email", 8.0, "Victor Alvarez", "Research Informatics", "4 - Low",
  "SNOMED codes missing in OMOP concept table",
  "Several SNOMED codes from 2026 aren't in concept. Is the load broken?",
  "Researcher can't find 2026 SNOMED codes in the OMOP concept table.",
  "Data quality", 44, "DE – Clinical Data (Cerner/OMOP)", 0.85,
  "Not a broken load — vocabulary is out of date. Covered by RITM0031207.", "GRP-05", "child",
  should_be_request=True)

# ---------------- GRP-06: misrouted logins ----------------
t("INC0104458", "Incident", "Email", 24.0, "Paul Mendes", "Radiology", "4 - Low",
  "Locked out of my Cerner account",
  "I tried my password too many times and now I'm locked out of PowerChart.",
  "User is locked out of PowerChart after too many password attempts.",
  "Misrouted", 40, "IT Service Desk (not DE)", 0.97,
  "Account lockout — IT Service Desk handles this, not Data Engineering.", "GRP-06", "parent")
t("INC0104509", "Incident", "Email", 7.0, "Sofia Reyes", "Pharmacy", "4 - Low",
  "Can't log in to PowerChart, password expired",
  "Password expired message and the reset link isn't working.",
  "User's PowerChart password expired and the reset link fails.",
  "Misrouted", 40, "IT Service Desk (not DE)", 0.96,
  "Password reset — belongs to the IT Service Desk.", "GRP-06", "child")

# ---------------- Singletons ----------------
t("INC0104493", "Incident", "Email", 11.0, "Erin Walsh", "Compliance", "4 - Low",
  "Patient names visible in workbook shared to all staff",
  "Hi, not sure who to send this to. The 'Discharge Follow-up' workbook in the All Staff project shows patient names and MRNs. I don't think that's supposed to be visible to everyone.",
  "Possible PHI exposure: patient names and MRNs visible in a Tableau workbook shared with all staff.",
  "Security / PHI", 97, "DE – Data Access & Security", 0.93,
  "Potential HIPAA exposure. Filed as Low via email and sitting unassigned for 11 hours — restrict the workbook now and loop in Privacy.",
  history="No similar tickets in 12 months.")
t("INC0104515", "Incident", "Monitoring", 3.0, "talend-monitor", "Automated", "3 - Moderate",
  "Pharmacy dispensing feed delayed 6 hours",
  "SLA monitor: PHARM_DISPENSE_HOURLY last success 02:55. Expected every 60 min.",
  "Pharmacy dispensing feed hasn't loaded in 6 hours.",
  "Pipeline failure", 76, "DE – Pipelines (Talend/ETL)", 0.82,
  "Hourly clinical feed down for 6 cycles; separate job from GRP-01 but same server — check talend-prd-02.",
  owner="Maya Chen",
  history="INC0099940 (Sep 4) — pharmacy feed stall; Maya Chen restarted the JobServer.")
t("INC0104474", "Incident", "Portal", 16.0, "Rachel Kim", "Quality & Analytics", "3 - Moderate",
  "Duplicate patient rows in readmissions report",
  "Some encounters appear twice in the 30-day readmissions output, inflating the rate. Example FINs attached.",
  "Readmissions report double-counts some encounters, inflating the readmission rate.",
  "Data quality", 71, "DE – Clinical Data (Cerner/OMOP)", 0.84,
  "Affects a reported quality metric. Likely an encounter-alias join issue.")
t("INC0104518", "Incident", "Portal", 4.5, "Angela Brooks", "Nursing Operations", "3 - Moderate",
  "Wrong unit mapping for 7 West in bed census",
  "7 West beds are showing under 7 East since the unit renovation. Counts are wrong for both units.",
  "7 West beds are mapped to 7 East in the bed census after the renovation.",
  "Data quality", 64, "DE – Clinical Data (Cerner/OMOP)", 0.8,
  "Related to bed census but a different issue from GRP-01 (mapping table, not the failed load).")
t("INC0104430", "Incident", "Portal", 40.0, "Monica Shah", "Quality & Analytics", "3 - Moderate",
  "CMS quality measures extract failing on null discharge dates",
  "Monthly CMS extract fails validation: 212 records with null discharge_dt. Submission due 10/20.",
  "Monthly CMS quality extract fails validation on 212 null discharge dates; due Oct 20.",
  "Pipeline failure", 68, "DE – Pipelines (Talend/ETL)", 0.83,
  "Regulatory submission with a hard deadline in 5 days.")
t("INC0104486", "Incident", "Portal", 13.0, "Kevin Osei", "BI Reporting", "4 - Low",
  "Tableau subscription emails not sending",
  "Scheduled subscriptions for the Daily Ops workbook didn't go out this morning.",
  "Daily Ops Tableau subscription emails didn't send this morning.",
  "Reporting", 48, "BI – Reporting (Tableau)", 0.81,
  "Reporting platform issue; no data impact.")
t("INC0104526", "Incident", "Email", 1.0, "Dr. Ravi Menon", "Surgery", "4 - Low",
  "Surgical case volume report doesn't match Cerner",
  "September case volume on the dashboard is 1,412 but SurgiNet shows 1,438. Which is right?",
  "September surgical case volume on the dashboard is 26 cases lower than SurgiNet.",
  "Data quality", 52, "BI – Reporting (Tableau)", 0.74,
  "Likely add-on cases excluded by report filter; needs analyst review.")
t("INC0104531", "Incident", "Monitoring", 0.4, "talend-monitor", "Automated", "4 - Low",
  "Talend job OR_CASES_HOURLY runtime 3x baseline",
  "Runtime 38 min vs 12 min baseline. Job still completed.",
  "Hourly OR cases job is running 3x slower than normal but still completing.",
  "Performance", 45, "DE – Pipelines (Talend/ETL)", 0.76,
  "Warning only. Watch — may be the same Redshift contention as GRP-02.")
t("INC0104507", "Incident", "Email", 8.5, "Marcus Hill", "Capacity Management", "4 - Low",
  "What counts as observation status in the census table?",
  "Quick question — does census include observation patients or only inpatients?",
  "Question about whether observation patients are included in the census table.",
  "Question", 18, "BI – Reporting (Tableau)", 0.88,
  "Definition question, not an incident. Answer with the data dictionary link.", should_be_request=True)
t("INC0104522", "Incident", "Email", 2.0, "Diane Foster", "Finance", "4 - Low",
  "Printer on 4th floor not printing",
  "The printer by the finance area isn't printing anything.",
  "4th-floor finance printer isn't printing.",
  "Misrouted", 10, "IT Service Desk (not DE)", 0.99,
  "Hardware issue — not Data Engineering.")
t("INC0104533", "Incident", "Email", 0.3, "Tom Becker", "Revenue Cycle", "4 - Low",
  "Vendor needs SFTP drop to S3",
  "Our new coding vendor needs to drop files to us daily. Can you set that up?",
  "Revenue Cycle needs a new daily SFTP-to-S3 drop for a coding vendor.",
  "Access request", 38, "DE – Data Access & Security", 0.86,
  "New integration request — should be a Request with security review.", should_be_request=True)
t("RITM0031190", "Request", "Portal", 52.0, "Rachel Kim", "Quality & Analytics", "4 - Low",
  "Add attending physician NPI to encounters table",
  "Please add attending_npi to dw.encounters so we can attribute measures by provider.",
  "Add attending physician NPI to the encounters table for provider-level measures.",
  "Data change request", 42, "DE – Redshift / Warehouse", 0.9,
  "Standard enhancement; schedule into sprint.")
t("RITM0031195", "Request", "Portal", 46.0, "Dr. Hannah Cole", "Research Informatics", "4 - Low",
  "Create Redshift sandbox schema for research team",
  "Need a sandbox schema (research_sbx) with create/drop rights for 4 analysts.",
  "Research needs a Redshift sandbox schema with write access for 4 analysts.",
  "Data change request", 36, "DE – Redshift / Warehouse", 0.89,
  "Standard request. Pairs well with GRP-02 fix (separate research WLM queue).")
t("RITM0031201", "Request", "Portal", 34.0, "Maya Chen", "Data Engineering", "4 - Low",
  "Talend Studio license for new contractor",
  "New contractor starting Monday needs a Talend Studio license.",
  "Talend Studio license needed for a contractor starting Monday.",
  "Access request", 26, "DE – Data Access & Security", 0.91,
  "Routine license request.")
t("RITM0031210", "Request", "Portal", 28.0, "Dev Patel", "Data Engineering", "4 - Low",
  "Decommission ETL_LEGACY linked server",
  "All jobs migrated off ETL_LEGACY. Request to remove the linked server and credentials.",
  "Remove the ETL_LEGACY linked server now that all jobs are migrated.",
  "Data change request", 22, "DE – Pipelines (Talend/ETL)", 0.87,
  "Cleanup task; low urgency.")
t("RITM0031214", "Request", "Portal", 5.0, "Teresa Nguyen", "Emergency Department", "4 - Low",
  "New dashboard: ED length of stay by hour",
  "Would like a dashboard showing ED LOS by arrival hour for staffing decisions.",
  "ED wants a new dashboard of length of stay by arrival hour for staffing.",
  "Reporting", 30, "BI – Reporting (Tableau)", 0.86,
  "New build request; goes to BI intake.")


def priority_label(score):
    if score >= 85:
        return "P1 – Critical"
    if score >= 65:
        return "P2 – High"
    if score >= 40:
        return "P3 – Medium"
    return "P4 – Low"


def main():
    out = Path(__file__).parent
    for row in T:
        row["ai_priority"] = priority_label(row["ai_priority_score"])
    payload = {
        "demo_now": DEMO_NOW.strftime("%Y-%m-%d %H:%M"),
        "queues": QUEUES,
        "groups": GROUPS,
        "tickets": sorted(T, key=lambda r: r["opened_at"]),
    }
    (out / "tickets.json").write_text(json.dumps(payload, indent=2))
    with open(out / "tickets.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(T[0].keys()))
        w.writeheader()
        w.writerows(payload["tickets"])
    print(f"Wrote {len(T)} tickets in {len(GROUPS)} groups to {out}")


if __name__ == "__main__":
    main()
