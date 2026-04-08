

import streamlit as st
import json
from openai import OpenAI
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import tempfile
from PIL import Image
import io
import pillow_heif

# -----------------------------
# Streamlit UI cleanup
# -----------------------------
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {visibility: hidden;}
    .viewerBadge_link__1S137 {display: none !important;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none;}
    .viewerBadge_link__qRIco {display: none;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# -----------------------------
# Clients / Secrets
# -----------------------------
client = OpenAI(api_key=st.secrets["openai_api_key"])
creds_dict = st.secrets["gcp_service_account"]

# -----------------------------
# Helpers
# -----------------------------
def convert_heic_to_jpeg(uploaded_file):
    uploaded_file.seek(0)
    heif_file = pillow_heif.read_heif(uploaded_file.read())
    image = Image.frombytes(heif_file.mode, heif_file.size, heif_file.data)

    jpeg_bytes = io.BytesIO()
    image.save(jpeg_bytes, format="JPEG")
    jpeg_bytes.seek(0)

    # Add required file attributes for OpenAI
    jpeg_bytes.name = uploaded_file.name.replace(".heic", ".jpeg")
    return jpeg_bytes

def upload_file_to_drive(uploaded_file, filename, folder_id=None):
    gauth = GoogleAuth()
    creds_dict = st.secrets["gcp_service_account"]
    scope = ["https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gauth.credentials = creds
    drive = GoogleDrive(gauth)

    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        tmp_file.write(uploaded_file.getbuffer())
        tmp_path = tmp_file.name

    file_metadata = {"title": filename}
    if folder_id:
        file_metadata["parents"] = [{"id": folder_id}]

    gfile = drive.CreateFile(file_metadata)
    gfile.SetContentFile(tmp_path)
    gfile.Upload(param={"supportsAllDrives": True})

    return gfile["id"], gfile["alternateLink"]

def upload_to_google_sheet(df: pd.DataFrame):
    from gspread.utils import rowcol_to_a1

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gs_client = gspread.authorize(creds)

    sheet = gs_client.open("BSD Master Data Submittals")
    worksheet = sheet.worksheet("Master Data")

    existing = worksheet.get_all_values()

    # If empty, write headers first
    if not existing:
        worksheet.append_row(df.columns.tolist(), value_input_option="USER_ENTERED")
        start_row = 2
    else:
        start_row = len(existing) + 1

    data = df.values.tolist()
    end_row = start_row + len(data) - 1
    end_col = len(df.columns)
    cell_range = f"A{start_row}:{rowcol_to_a1(end_row, end_col)}"

    worksheet.update(cell_range, data, value_input_option="USER_ENTERED")

# -----------------------------
# Cost code mapping
# -----------------------------
cost_code_mapping_text = """00030 - Financing Fees
01020 - First Aid/safety/Inspect/Carp./Lab
01025 - Safety Supplies
01100 - Surveying
01200 - Hydro/Gas/Telus Services
01210 - Temp Hydro
01220 - Temporary Heat
01230 - Temporary Lighting & Security Lighting
01240 - Temporary Water
01250 - Temporary Fencing
01350 - Miscellaneous Expenses
01400 - Tree Protection
01510 - Site Office
01520 - Sanitary Facilities
01560 - Project Construction Signs
01721 - Pressure Washing
01750 - Disposal Bins/Fees
01760 - Protect Finishes
01810 - Hoist/ crane/Scaffold rental
01820 - Winter Protection
02270 - Erosion & Sediment Control
02300 - Site Services (Fence)
02310 - Finish Grading
02315 - Excavation & Backfill
02600 - Drainaige & Stormwater
02621 - Foundation Drain Tile
02700 - Exterior Hardscape
02705 - Exterior Decking
02773 - Curbs & Gutters & Sidewalk
02820 - Fencing & Gates (Fnds, Stone & Alumn)
02900 - Landscaping
02910 - Irrigation Systems
03050 - Concrete Material
03100 - Formwork Material
03210 - Reinforcing Steel Material and Hardware
04200 - Masonry
04400 - Stone Veneer
05090 - Exterior Railing and Guardrail
05095 - Driveway Gates & Fencing
05100 - Steel Beams
05700 - Metal Chimney Cap
05710 - Deck Flashing
06060 - Framing Lumber
06175 - Wood Trusses
06200 - Interior Finishing Material
06220 - Finishing Labor
06410 - Custom Cabinets
06415 - Bath Vanity
06420 - Stone/Countertop - Material
06425 - Stone/Countertop - Fabrication
06430 - Interior Railings
06450 - Fireplace Mantels
07200 - Interior Waterproofing/Shower pan
07210 - Building Insulation
07220 - Building Exterior Waterproofing/Vapour Barrier
07311 - Roofing System
07450 - Siding/Trims - Material
07465 - Stucco
07500 - Torch & Decking
07600 - Metal Roofing - Prepainted Aluminum
07714 - Gutter & Downspouts
07920 - Sealants & Caulking
08210 - Interior Doors
08215 - Exterior Doors
08216 - Front/Entrance Door
08220 - Closet Doors - Bifolds
08360 - Garage Door
08560 - Window Material
08580 - Window Waterproofing
08600 - Skylights
08700 - Cabinetry and finish hardware
08800 - Door hardware
09200 - Drywall Systems
09300 - Exterior Tile Work- Material
09640 - Wood Flooring - Material
09650 - Interior Tile Work- Material
09680 - Carpeting - Material
09900 - Painting Exterior
09905 - Painting Interior
09910 - Wallpaper Material
10810 - Residential Washroom Accessories
10820 - Shower Enclosures
10830 - Bathroom Mirrors
10840 - Mirror and Glazing
10850 - Wine Rack
10900 - Closet Specialties
11450 - Appliances
11455 - Built-in Vacuum
11460 - Outdoor Kitchen BBQ & Sink
12490 - Window Treatment
12500 - Furniture
13150 - Swimming Pools
13160 - Generator
13170 - Dry Sauna
13180 - Hot Tubs
15015 - Plumbing Rough in
15300 - Fire Protection (Sprinklers)
15410 - Plumbing Fixtures
15500 - Radiant Heating
15610 - Wine Cellar Cooling Unit
15700 - Air Conditioning/HRV
15750 - Fire Place Inserts
16050 - General Electrical
16100 - Solar System
16500 - Fixtures
16800 - Low Voltage (Security, Internet)
16900 - Sound and Audio"""

# -----------------------------
# App
# -----------------------------
st.title("BSD Receipt Submittals")

with st.form("receipt_form"):
    # Dropdown 1
    st.markdown("#### Property")
    property_dropdown = st.selectbox(
        "Select Property",
        [
            "",
            "Coto",
            "Milford",
            "647 Navy",
            "645 Navy",
            "Sagebrush",
            "Paramount",
            "126 Scenic",
            "San Marino",
            "King Arthur",
            "Via Sonoma",
            "Highland",
            "Channel View",
            "Paseo De las Estrellas",
            "Marguerite",
            "BSD SHOP",
            "5 Montepellier",
            "Sycamore",
        ],
    )
    property_manual = st.text_input("Or enter manually:", key="property_manual_input")

    # Dropdown 2
    st.markdown("#### Payable Party")
    payable_party_dropdown = st.selectbox(
        "Select from list",
        ["", "Christian Granados (Vendor)", "Jessica Ajtun", "Andres De Jesus", "Nick Yuh (Vendor)"],
        key="dropdown",
    )
    payable_party_manual = st.text_input("Or enter manually:", key="manual_input")

    st.markdown("#### Payment Method")
    payment_method_dropdown = st.selectbox(
        "Select from list",
        ["", "AMEX", "Zelle (Construction)", "Zelle (Materials)"],
        key="pay_drop",
    )
    payment_method_manual = st.text_input("Or enter manually:", key="pay_manual_input")

    uploaded_files = st.file_uploader(
        "Upload Receipt Image",
        type=["jpg", "jpeg", "png", "heif", "heic"],
        accept_multiple_files=True,
    )

    submitted = st.form_submit_button("Submit Form")

    if submitted:
        # Validate all fields
        property_val = property_manual.strip() if property_manual.strip() else property_dropdown
        payable_party_val = payable_party_manual.strip() if payable_party_manual.strip() else payable_party_dropdown
        payment_method_val = (
            payment_method_manual.strip()
            if payment_method_manual.strip()
            else (payment_method_dropdown if payment_method_dropdown else None)
        )

        if not property_val or not payable_party_val or not uploaded_files:
            st.error("Please complete all fields and upload a receipt.")
        else:
            for uploaded_file in uploaded_files:
                with st.spinner("Uploading and processing..."):
                    if uploaded_file is not None and uploaded_file.type in ["image/heic", "image/heif"]:
                        try:
                            uploaded_file = convert_heic_to_jpeg(uploaded_file)
                        except Exception as e:
                            st.error(f"Failed to convert HEIC file: {e}")
                            st.stop()  # Stop execution if conversion fails

                    # Upload file to OpenAI
                    file_id = client.files.create(file=uploaded_file, purpose="vision").id

                    # Build prompt
                    response = client.responses.create(
                        model="gpt-4.1-mini",
                        input=[
                            {
                                "role": "user",
                                "content": [
                                    {
                                        "type": "input_text",
                                        "text": (
                                            "From this receipt image, extract:\n"
                                            "- The date of purchase. This is a recent receipt. If the receipt shows a 2-digit year (e.g. '4-2-25'), assume it's 20xx, not 19xx.\n"
                                            "- A list of items with each item's 'name' and 'price'. If there is a small fee for an item, combine the two prices. The fee should not have it's own row. \n"
                                            "- The total tax amount (if present), plus any extra fees, surcharges, or additional subtotal line items not associated with specific items (e.g., environmental fees, cartage, service fees, recycling fees, etc). These should all be treated as part of 'tax'.\n"
                                            "- Assign a cost code to each item based on its name using the mapping below. Do not give any justification. If no matching cost code assign under Miscellaneous Expenses cost code.\n\n"
                                            "Each assigned cost code should be in the format: 'CODE - Description'.\n\n"
                                            "Return a JSON object in this format:\n"
                                            "{\n"
                                            '  "date": "YYYY-MM-DD",\n'
                                            '  "items": [\n'
                                            '    {"name": "Item name", "price": 0.00, "cost_code": "00030 - Financing Fees"},\n'
                                            '    ...\n'
                                            "  ],\n"
                                            '  "tax": 0.00\n'
                                            "}\n\n"
                                            "Here is the cost code mapping:\n"
                                            + cost_code_mapping_text
                                        ),
                                    },
                                    {"type": "input_image", "file_id": file_id},
                                ],
                            }
                        ],
                    )

                    drive_file_id, drive_link = upload_file_to_drive(
                        uploaded_file,
                        uploaded_file.name,
                        folder_id="1f1tN4BaGPn5oruX7ngNsZPBKl23FZkWu",
                    )

                    raw_text = response.output[0].content[0].text
                    cleaned_text = raw_text.strip("```json").strip("```").strip()
                    parsed = json.loads(cleaned_text)

                    date_val = parsed["date"]
                    items = parsed["items"]
                    tax_val = parsed.get("tax", 0.0)

                    df = pd.DataFrame(items)
                    df["price"] = df["price"].astype(float)

                    subtotal = df["price"].sum()
                    df["tax_share"] = df["price"] / subtotal * tax_val
                    df["amount"] = (df["price"] + df["tax_share"]).round(2)

                    df["Date Invoiced"] = date_val
                    df["Property"] = property_val
                    df["Payable Party"] = payable_party_val
                    df.rename(columns={"name": "Item Name", "cost_code": "Cost Code"}, inplace=True)

                    df["Date Paid"] = None
                    df["Unique ID"] = None
                    df["Worker Name"] = None
                    df["Hours"] = None
                    df["Claim Number"] = None
                    df["QB Property"] = None
                    df["Invoice Number"] = None
                    df["Project Description"] = df["Item Name"]
                    df["Status"] = None
                    df["Form"] = "MATERIALS"
                    df["Drive Link"] = drive_link
                    df["Equation Description"] = df["Item Name"]
                    df["Payment Method"] = payment_method_val if payment_method_val is not None else ""

                    final_df = df[
                        [
                            "Date Paid",
                            "Date Invoiced",
                            "Unique ID",
                            "Claim Number",
                            "Worker Name",
                            "Hours",
                            "Item Name",
                            "Property",
                            "QB Property",
                            "amount",
                            "Payable Party",
                            "Project Description",
                            "Invoice Number",
                            "Cost Code",
                            "Payment Method",
                            "Status",
                            "Form",
                            "Drive Link",
                            "Equation Description",
                        ]
                    ].copy()
                    final_df.rename(columns={"amount": "Amount"}, inplace=True)

                    upload_to_google_sheet(final_df)

            st.success("Form Fully Submitted!")