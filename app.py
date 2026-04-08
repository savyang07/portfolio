pip install streamlit openai pillow pillow-heif google-api-python-client google-auth google-auth-httplib2

import streamlit as st
import json
import uuid
import io
import os
import base64
from datetime import datetime, date
from typing import Optional

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BSD Receipt Processor",
    page_icon="🧾",
    layout="centered",
)

# ── Inject custom CSS ──────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* ── global resets ──────────────────────────────────────────────── */
  html, body, [class*="css"] {
    font-family: 'Inter', 'Segoe UI', sans-serif;
  }

  /* ── hero header ────────────────────────────────────────────────── */
  .hero {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    border-radius: 16px;
    padding: 2rem 2.4rem 1.8rem;
    margin-bottom: 1.6rem;
    position: relative;
    overflow: hidden;
  }
  .hero::before {
    content: "";
    position: absolute;
    top: -60px; right: -60px;
    width: 200px; height: 200px;
    background: radial-gradient(circle, rgba(229,57,53,.18) 0%, transparent 70%);
    border-radius: 50%;
  }
  .hero h1 {
    color: #fff;
    font-size: 1.75rem;
    font-weight: 700;
    margin: 0 0 .35rem;
    letter-spacing: -.5px;
  }
  .hero p {
    color: #94a3b8;
    font-size: .95rem;
    margin: 0;
  }

  /* ── section card ───────────────────────────────────────────────── */
  .card {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1.2rem;
  }
  .card-title {
    color: #e2e8f0;
    font-size: 1rem;
    font-weight: 600;
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: .5rem;
  }

  /* ── status pill ────────────────────────────────────────────────── */
  .pill {
    display: inline-flex;
    align-items: center;
    gap: .35rem;
    padding: .28rem .75rem;
    border-radius: 999px;
    font-size: .8rem;
    font-weight: 600;
  }
  .pill-success { background:#dcfce7; color:#166534; }
  .pill-error   { background:#fee2e2; color:#991b1b; }
  .pill-info    { background:#dbeafe; color:#1e40af; }

  /* ── receipt result row ─────────────────────────────────────────── */
  .receipt-card {
    background:#0f172a;
    border:1px solid #1e3a5f;
    border-radius:10px;
    padding:1rem 1.2rem;
    margin-bottom:.8rem;
  }
  .receipt-header {
    display:flex;
    justify-content:space-between;
    align-items:center;
    margin-bottom:.6rem;
  }
  .receipt-filename { color:#60a5fa; font-weight:600; font-size:.95rem; }
  .item-table { width:100%; border-collapse:collapse; margin-top:.5rem; }
  .item-table th {
    text-align:left; color:#64748b;
    font-size:.75rem; font-weight:600;
    text-transform:uppercase; letter-spacing:.05em;
    padding:.3rem .5rem; border-bottom:1px solid #1e293b;
  }
  .item-table td {
    color:#cbd5e1; font-size:.85rem;
    padding:.35rem .5rem; border-bottom:1px solid #0f172a;
  }
  .item-table tr:last-child td { border-bottom:none; }

  /* ── success banner ─────────────────────────────────────────────── */
  .success-banner {
    background: linear-gradient(135deg,#064e3b,#065f46);
    border:1px solid #10b981;
    border-radius:12px;
    padding:1.4rem 1.6rem;
    text-align:center;
    margin-top:1rem;
  }
  .success-banner h3 { color:#d1fae5; font-size:1.3rem; margin:0 0 .4rem; }
  .success-banner p  { color:#6ee7b7; font-size:.9rem; margin:0; }

  /* ── override Streamlit defaults for dark feel ──────────────────── */
  [data-testid="stForm"] {
    background: transparent !important;
    border: none !important;
    padding: 0 !important;
  }
  .stSelectbox label, .stFileUploader label { color:#94a3b8 !important; }
</style>
""", unsafe_allow_html=True)

# ── Lazy imports (only load when needed) ──────────────────────────────────────
@st.cache_resource
def get_openai_client():
    import openai
    api_key = st.secrets.get("OPENAI_API_KEY") or os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        return None
    return openai.OpenAI(api_key=api_key)

@st.cache_resource
def get_gdrive_service():
    from googleapiclient.discovery import build
    from google.oauth2.service_account import Credentials
    try:
        creds_dict = dict(st.secrets.get("GCP_SERVICE_ACCOUNT", {}))
        if not creds_dict:
            raw = os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "{}")
            creds_dict = json.loads(raw)
        if not creds_dict:
            return None, None
        scopes = [
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets",
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        drive   = build("drive",   "v3", credentials=creds)
        sheets  = build("sheets",  "v4", credentials=creds)
        return drive, sheets
    except Exception:
        return None, None

# ── Config / mappings ─────────────────────────────────────────────────────────
PROPERTIES = [
    "-- Select --",
    "12345 Main St",
    "456 Oak Ave",
    "789 Pine Rd",
    "321 Cedar Blvd",
    "Other (enter below)",
]

PAYABLE_PARTIES = [
    "-- Select --",
    "Home Depot",
    "Lowe's",
    "Ace Hardware",
    "Amazon",
    "Fastenal",
    "Other (enter below)",
]

PAYMENT_METHODS = [
    "-- Select --",
    "Company Credit Card",
    "Company Check",
    "Personal Card (Reimbursable)",
    "Cash",
    "ACH / Wire",
    "Other",
]

COST_CODE_MAP = {
    # Labour / GCs
    "general conditions": "01000 - General Conditions",
    "supervision":        "01000 - General Conditions",
    "temporary":          "01500 - Temporary Facilities",
    "porta potty":        "01500 - Temporary Facilities",
    "dumpster":           "01500 - Temporary Facilities",
    "safety":             "01550 - Safety",
    "permit":             "01450 - Permits & Fees",
    "inspection":         "01450 - Permits & Fees",
    # Site / demo
    "demolition":         "02000 - Demolition",
    "excavat":            "02300 - Earthwork",
    "grading":            "02300 - Earthwork",
    "concrete":           "03000 - Concrete",
    "cement":             "03000 - Concrete",
    "rebar":              "03200 - Reinforcing Steel",
    # Structural / framing
    "lumber":             "06100 - Rough Carpentry",
    "framing":            "06100 - Rough Carpentry",
    "plywood":            "06100 - Rough Carpentry",
    "osb":                "06100 - Rough Carpentry",
    "engineered wood":    "06170 - Engineered Wood",
    "finish carpentry":   "06200 - Finish Carpentry",
    "millwork":           "06200 - Finish Carpentry",
    # Moisture / envelope
    "roofing":            "07500 - Roofing",
    "shingle":            "07500 - Roofing",
    "waterproof":         "07100 - Waterproofing",
    "insulation":         "07200 - Insulation",
    "weatherstrip":       "07900 - Sealants & Caulking",
    "caulk":              "07900 - Sealants & Caulking",
    "sealant":            "07900 - Sealants & Caulking",
    # Openings
    "door":               "08100 - Doors",
    "window":             "08500 - Windows",
    "glass":              "08800 - Glazing",
    # Finishes
    "drywall":            "09250 - Drywall",
    "gypsum":             "09250 - Drywall",
    "tile":               "09300 - Tile",
    "flooring":           "09600 - Flooring",
    "hardwood":           "09600 - Flooring",
    "laminate":           "09600 - Flooring",
    "vinyl":              "09600 - Flooring",
    "carpet":             "09680 - Carpet",
    "paint":              "09900 - Paint & Coatings",
    "primer":             "09900 - Paint & Coatings",
    "stain":              "09900 - Paint & Coatings",
    "coating":            "09900 - Paint & Coatings",
    # MEP
    "plumbing":           "15000 - Plumbing",
    "pipe":               "15000 - Plumbing",
    "fixture":            "15000 - Plumbing",
    "faucet":             "15000 - Plumbing",
    "toilet":             "15000 - Plumbing",
    "hvac":               "15600 - HVAC",
    "ductwork":           "15600 - HVAC",
    "furnace":            "15600 - HVAC",
    "ac unit":            "15600 - HVAC",
    "electrical":         "16000 - Electrical",
    "wiring":             "16000 - Electrical",
    "conduit":            "16000 - Electrical",
    "outlet":             "16000 - Electrical",
    "breaker":            "16000 - Electrical",
    "panel":              "16000 - Electrical",
    "light":              "16500 - Lighting",
    "fixture":            "16500 - Lighting",
    # Equipment / tools
    "tool":               "01600 - Equipment & Tools",
    "equipment":          "01600 - Equipment & Tools",
    "drill":              "01600 - Equipment & Tools",
    "blade":              "01600 - Equipment & Tools",
    "fastener":           "06050 - Fasteners & Adhesives",
    "screw":              "06050 - Fasteners & Adhesives",
    "nail":               "06050 - Fasteners & Adhesives",
    "bolt":               "06050 - Fasteners & Adhesives",
    "adhesive":           "06050 - Fasteners & Adhesives",
    "glue":               "06050 - Fasteners & Adhesives",
    # Landscaping / exterior
    "landscape":          "02900 - Landscaping",
    "sod":                "02900 - Landscaping",
    "mulch":              "02900 - Landscaping",
    # Cabinetry / millwork
    "cabinet":            "06410 - Cabinetry",
    "vanity":             "06410 - Cabinetry",
    "countertop":         "06410 - Cabinetry",
    # Appliances
    "appliance":          "11400 - Appliances",
    "refrigerator":       "11400 - Appliances",
    "dishwasher":         "11400 - Appliances",
    "microwave":          "11400 - Appliances",
    "range":              "11400 - Appliances",
    "washer":             "11400 - Appliances",
    "dryer":              "11400 - Appliances",
    # Misc
    "cleaning":           "01710 - Cleaning",
    "clean":              "01710 - Cleaning",
    "hauling":            "02080 - Hauling",
    "delivery":           "01300 - Project Management",
    "shipping":           "01300 - Project Management",
    "coffee":             "01350 - Miscellaneous Expenses",
    "food":               "01350 - Miscellaneous Expenses",
}

DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "")  # override via secrets
SPREADSHEET_ID  = os.environ.get("SPREADSHEET_ID", "")
SHEET_NAME      = "Master Data"

SHEET_HEADERS = [
    "Date Paid", "Date Invoiced", "Unique ID", "Claim Number", "Worker Name",
    "Hours", "Item Name", "Property", "QB Property", "Amount", "Payable Party",
    "Project Description", "Invoice Number", "Cost Code", "Payment Method",
    "Status", "Form", "Drive Link", "Equation Description",
]

# ── Helper functions ──────────────────────────────────────────────────────────

def assign_cost_code(item_name: str) -> str:
    lower = item_name.lower()
    for keyword, code in COST_CODE_MAP.items():
        if keyword in lower:
            return code
    return "01350 - Miscellaneous Expenses"


def convert_heic(file_bytes: bytes) -> bytes:
    """Convert HEIC/HEIF bytes to JPEG bytes."""
    from pillow_heif import register_heif_opener
    from PIL import Image
    register_heif_opener()
    img = Image.open(io.BytesIO(file_bytes))
    out = io.BytesIO()
    img.convert("RGB").save(out, format="JPEG")
    return out.getvalue()


def extract_receipt_data(client, image_bytes: bytes, mime: str) -> dict:
    """Call OpenAI vision model, return structured receipt data."""
    b64 = base64.standard_b64encode(image_bytes).decode()
    data_url = f"data:{mime};base64,{b64}"

    prompt = """You are a receipt data extraction specialist.
Extract ALL information from this receipt and return ONLY valid JSON (no markdown, no extra text).

Required JSON schema:
{
  "date": "YYYY-MM-DD",       // receipt/invoice date; use today if unclear
  "invoice_number": "string", // blank if not found
  "items": [
    { "name": "string", "price": 0.00 }
  ],
  "tax": 0.00,                // total tax + all extra fees/surcharges not tied to a specific item
  "grand_total": 0.00
}

Rules:
- Assume year is 20xx for 2-digit years.
- Separate combined items into individual line items when possible.
- Do NOT include tax or fees as items; put them in the "tax" field.
- Prices are positive numbers in USD.
- If an item has no separate price, estimate 0.
"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": data_url, "detail": "high"}},
                    {"type": "text", "text": prompt},
                ],
            }
        ],
        max_tokens=1500,
        temperature=0,
    )
    raw = response.choices[0].message.content.strip()
    # Strip possible markdown fences
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw)


def upload_to_drive(drive_service, file_bytes: bytes, filename: str, folder_id: str) -> str:
    """Upload bytes to Google Drive, return shareable link."""
    from googleapiclient.http import MediaIoBaseUpload
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype="image/jpeg", resumable=False)
    meta = {"name": filename, "parents": [folder_id] if folder_id else []}
    created = drive_service.files().create(body=meta, media_body=media, fields="id, webViewLink").execute()
    # Make it readable by anyone with the link
    drive_service.permissions().create(
        fileId=created["id"],
        body={"type": "anyone", "role": "reader"},
    ).execute()
    return created.get("webViewLink", "")


def ensure_sheet_headers(sheets_service, spreadsheet_id: str):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{SHEET_NAME}!A1:Z1",
    ).execute()
    values = result.get("values", [])
    if not values:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="RAW",
            body={"values": [SHEET_HEADERS]},
        ).execute()


def append_rows_to_sheet(sheets_service, spreadsheet_id: str, rows: list[list]):
    sheets_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"{SHEET_NAME}!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()


def build_sheet_rows(
    extracted: dict,
    items_with_tax: list[dict],
    property_val: str,
    payable_party: str,
    payment_method: str,
    drive_link: str,
    date_paid: str,
) -> list[list]:
    rows = []
    for item in items_with_tax:
        unique_id = str(uuid.uuid4())[:8].upper()
        cost_code = assign_cost_code(item["name"])
        row = [
            date_paid,                          # Date Paid
            extracted.get("date", ""),           # Date Invoiced
            unique_id,                           # Unique ID
            "",                                  # Claim Number
            "",                                  # Worker Name
            "",                                  # Hours
            item["name"],                        # Item Name
            property_val,                        # Property
            "",                                  # QB Property
            round(item["final_amount"], 2),      # Amount
            payable_party,                       # Payable Party
            "",                                  # Project Description
            extracted.get("invoice_number", ""), # Invoice Number
            cost_code,                           # Cost Code
            payment_method,                      # Payment Method
            "Pending",                           # Status
            "MATERIALS",                         # Form
            drive_link,                          # Drive Link
            "",                                  # Equation Description
        ]
        rows.append(row)
    return rows


def allocate_tax(items: list[dict], tax: float) -> list[dict]:
    subtotal = sum(i["price"] for i in items) or 1
    result = []
    for item in items:
        share = (item["price"] / subtotal) * tax
        result.append({
            "name": item["name"],
            "price": item["price"],
            "tax_share": round(share, 2),
            "final_amount": round(item["price"] + share, 2),
        })
    return result


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
  <h1>🧾 BSD Receipt Processor</h1>
  <p>Upload receipts → AI extraction → structured rows in Google Sheets + Drive</p>
</div>
""", unsafe_allow_html=True)

# ── Configuration sidebar ──────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    st.caption("Override defaults for this session")

    sidebar_openai = st.text_input("OpenAI API Key", type="password",
                                   placeholder="sk-…", key="openai_key")
    sidebar_drive_folder = st.text_input("Drive Folder ID",
                                         value=DRIVE_FOLDER_ID, key="drive_folder")
    sidebar_sheet_id = st.text_input("Spreadsheet ID",
                                      value=SPREADSHEET_ID, key="sheet_id")
    st.divider()
    st.caption("💡 For production, set keys in `.streamlit/secrets.toml`")

    with st.expander("📋 Cost Code Reference"):
        st.dataframe(
            {"Keyword": list(COST_CODE_MAP.keys()), "Code": list(COST_CODE_MAP.values())},
            height=300, use_container_width=True,
        )

# ── Main form ──────────────────────────────────────────────────────────────────
with st.form("receipt_form", clear_on_submit=False):

    # — Metadata card —
    st.markdown('<div class="card"><div class="card-title">📋 Submission Details</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        property_select = st.selectbox("Property *", PROPERTIES)
        property_override = st.text_input("Custom Property (if 'Other')", placeholder="e.g. 999 Elm St")
    with col2:
        party_select = st.selectbox("Payable Party *", PAYABLE_PARTIES)
        party_override = st.text_input("Custom Payable Party (if 'Other')", placeholder="e.g. Local Lumber Co")

    col3, col4 = st.columns(2)
    with col3:
        payment_method = st.selectbox("Payment Method *", PAYMENT_METHODS)
    with col4:
        date_paid = st.date_input("Date Paid *", value=date.today())

    st.markdown('</div>', unsafe_allow_html=True)

    # — Upload card —
    st.markdown('<div class="card"><div class="card-title">📎 Upload Receipts</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Upload receipt images (JPG, PNG, HEIC/HEIF supported)",
        type=["jpg", "jpeg", "png", "heic", "heif"],
        accept_multiple_files=True,
        help="Drag & drop or click to browse. Multiple receipts supported.",
    )
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) ready", icon="📁")
    st.markdown('</div>', unsafe_allow_html=True)

    # — Submit —
    submitted = st.form_submit_button("🚀 Process & Submit Receipts", use_container_width=True, type="primary")

# ── Processing ─────────────────────────────────────────────────────────────────
if submitted:
    # Resolve property / party
    property_val  = property_override.strip()  if "Other" in property_select  else property_select
    payable_party = party_override.strip()     if "Other" in party_select     else party_select

    # Validation
    errors = []
    if not property_val or property_val == "-- Select --":
        errors.append("Property is required.")
    if not payable_party or payable_party == "-- Select --":
        errors.append("Payable Party is required.")
    if payment_method == "-- Select --":
        errors.append("Payment Method is required.")
    if not uploaded_files:
        errors.append("At least one receipt image is required.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    # Resolve config
    openai_key   = sidebar_openai or (st.secrets.get("OPENAI_API_KEY") if hasattr(st, "secrets") else None) or os.environ.get("OPENAI_API_KEY", "")
    drive_folder = sidebar_drive_folder or DRIVE_FOLDER_ID
    sheet_id     = sidebar_sheet_id or SPREADSHEET_ID
    date_paid_str = date_paid.strftime("%Y-%m-%d")

    if not openai_key:
        st.error("⚠️ OpenAI API key not configured. Add it in the sidebar or `secrets.toml`.")
        st.stop()

    import openai as _openai
    client = _openai.OpenAI(api_key=openai_key)
    drive_service, sheets_service = get_gdrive_service()

    # Warn if integrations unavailable
    if not drive_service:
        st.warning("⚠️ Google Drive not configured — Drive upload skipped. Rows will still be shown.", icon="☁️")
    if not sheets_service or not sheet_id:
        st.warning("⚠️ Google Sheets not configured — Sheet write skipped. Rows will still be shown.", icon="📊")

    # Ensure headers
    if sheets_service and sheet_id:
        try:
            ensure_sheet_headers(sheets_service, sheet_id)
        except Exception as ex:
            st.warning(f"Could not verify sheet headers: {ex}")

    st.markdown("---")
    st.markdown("### Processing Results")

    total_rows = 0
    all_results = []

    for idx, uf in enumerate(uploaded_files):
        fname = uf.name
        fext  = fname.rsplit(".", 1)[-1].lower()

        with st.status(f"🔄 Processing **{fname}** ({idx+1}/{len(uploaded_files)})…", expanded=True) as status:

            # 1. Read bytes
            raw_bytes = uf.read()

            # 2. HEIC conversion
            if fext in ("heic", "heif"):
                st.write("Converting HEIC → JPEG…")
                try:
                    raw_bytes = convert_heic(raw_bytes)
                    fname = fname.rsplit(".", 1)[0] + ".jpg"
                    fext = "jpg"
                except Exception as ex:
                    status.update(label=f"❌ HEIC conversion failed: {ex}", state="error")
                    all_results.append({"file": uf.name, "error": str(ex)})
                    continue

            mime = "image/jpeg" if fext in ("jpg", "jpeg") else "image/png"

            # 3. Extract via OpenAI
            st.write("Extracting receipt data with GPT-4o…")
            try:
                extracted = extract_receipt_data(client, raw_bytes, mime)
            except Exception as ex:
                status.update(label=f"❌ Extraction failed: {ex}", state="error")
                all_results.append({"file": uf.name, "error": str(ex)})
                continue

            items        = extracted.get("items", [])
            tax          = float(extracted.get("tax", 0))
            grand_total  = float(extracted.get("grand_total", 0))

            if not items:
                status.update(label=f"⚠️ No line items found in {uf.name}", state="error")
                all_results.append({"file": uf.name, "error": "No items extracted"})
                continue

            items_with_tax = allocate_tax(items, tax)

            # 4. Drive upload
            drive_link = ""
            if drive_service and drive_folder:
                st.write("Uploading to Google Drive…")
                try:
                    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
                    dname = f"receipt_{ts}_{fname}"
                    drive_link = upload_to_drive(drive_service, raw_bytes, dname, drive_folder)
                except Exception as ex:
                    st.warning(f"Drive upload failed: {ex}")

            # 5. Build & write sheet rows
            rows = build_sheet_rows(
                extracted, items_with_tax,
                property_val, payable_party, payment_method,
                drive_link, date_paid_str,
            )

            if sheets_service and sheet_id:
                st.write("Writing to Google Sheets…")
                try:
                    append_rows_to_sheet(sheets_service, sheet_id, rows)
                except Exception as ex:
                    st.warning(f"Sheet write failed: {ex}")

            total_rows += len(rows)
            status.update(label=f"✅ {fname} — {len(rows)} item(s) processed", state="complete")
            all_results.append({
                "file": fname,
                "extracted": extracted,
                "items_with_tax": items_with_tax,
                "drive_link": drive_link,
                "rows": rows,
                "grand_total": grand_total,
            })

    # ── Show results ──────────────────────────────────────────────────────────
    st.markdown("---")
    for res in all_results:
        if "error" in res:
            st.error(f"**{res['file']}** → {res['error']}")
            continue

        items_with_tax = res["items_with_tax"]

        with st.expander(f"📄 {res['file']}  —  {len(items_with_tax)} item(s)  |  Total: ${res['grand_total']:.2f}", expanded=True):
            # Table of items
            table_rows = ""
            for it in items_with_tax:
                cc = assign_cost_code(it["name"])
                table_rows += f"""
                <tr>
                  <td>{it['name']}</td>
                  <td>${it['price']:.2f}</td>
                  <td>${it['tax_share']:.2f}</td>
                  <td><strong>${it['final_amount']:.2f}</strong></td>
                  <td><span style="color:#60a5fa;font-size:.8rem">{cc}</span></td>
                </tr>"""

            st.markdown(f"""
            <div class="receipt-card">
              <div class="receipt-header">
                <span class="receipt-filename">📅 Invoice Date: {res['extracted'].get('date','—')}</span>
                <span class="pill pill-success">✓ Processed</span>
              </div>
              <table class="item-table">
                <thead><tr>
                  <th>Item</th><th>Price</th><th>Tax Share</th><th>Final</th><th>Cost Code</th>
                </tr></thead>
                <tbody>{table_rows}</tbody>
              </table>
            </div>
            """, unsafe_allow_html=True)

            if res["drive_link"]:
                st.markdown(f"☁️ **Drive link:** [{res['file']}]({res['drive_link']})")
            else:
                st.caption("Drive link: not available (Drive not configured)")

    # ── Final success banner ───────────────────────────────────────────────────
    if total_rows > 0:
        st.markdown(f"""
        <div class="success-banner">
          <h3>🎉 Form Fully Submitted!</h3>
          <p>{len([r for r in all_results if 'error' not in r])} receipt(s) processed &nbsp;·&nbsp;
             {total_rows} row(s) written to Google Sheets</p>
        </div>
        """, unsafe_allow_html=True)

    elif all(("error" in r) for r in all_results):
        st.error("All receipts failed to process. Please check the errors above.")


