"""
Smart Vehicle Management System — Supabase (PostgreSQL) Edition
==================================================================
Enterprise-grade backend: Supabase Postgres via the official `supabase-py`
SDK, replacing SQLite. Adds self-service user registration with admin
approval, in addition to requisition approval. No email — everything
happens live inside the app. Exports: Excel (.xlsx) and PDF (.pdf).

Author: Senior Python Developer (generated for Kazi Nur Shahin)
"""

import io
import hashlib
import random
from datetime import datetime, date

import pandas as pd
import plotly.express as px
import streamlit as st
from supabase import create_client, Client
from fpdf import FPDF
from fpdf.enums import XPos, YPos

# =========================================================
# 1. COMPANY INFO & PAGE CONFIG
# =========================================================
COMPANY_NAME = "Renaissaince Barind Ltd."
COMPANY_ADDRESS = "Ishwardi EPZ, Pakshi, Pabna"

st.set_page_config(
    page_title=f"{COMPANY_NAME} | Vehicle Management System",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

USERS_TABLE = "users"
REQUISITIONS_TABLE = "requisitions"

VEHICLE_TYPES = ["Private Car", "HIACE", "Pick-up Van", "Covered Van", "Truck", "Shipment Vehicle", "Other"]
DEPARTMENTS = [
    "Accounts", "Warehouse", "Factory Merchandising", "Commercial", "Floor Operation",
    "TSD", "QMS", "Production Planning & Control", "Cutting", "R & D", "Admin",
    "Technical", "Finishing", "Quality", "IE", "HR & Compliance", "Other",
]
STATUS_OPTIONS = ["Pending", "Approved", "Rejected"]
STATUS_BADGE = {"Pending": "🟡 Pending", "Approved": "🟢 Approved", "Rejected": "🔴 Rejected"}

# =========================================================
# 2. STYLING (desktop / Windows browser + mobile / Android friendly)
# =========================================================
st.markdown("""
<style>
    /* ---- Base (desktop / Windows browser) ---- */
    .main > div { padding-top: 1.2rem; }
    div.stButton > button { border-radius: 8px; font-weight: 600; }
    .req-card {
        background: white; border: 1px solid #eee; border-radius: 12px;
        padding: 14px 18px; margin-bottom: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .badge-pending  { background:#FFF3CD; color:#856404; padding:4px 12px; border-radius:14px; font-weight:700; white-space:nowrap; }
    .badge-approved { background:#D4EDDA; color:#155724; padding:4px 12px; border-radius:14px; font-weight:700; white-space:nowrap; }
    .badge-rejected { background:#F8D7DA; color:#721C24; padding:4px 12px; border-radius:14px; font-weight:700; white-space:nowrap; }
    .driver-box {
        background:#EAF7EE; border:1px solid #B7E4C7; border-radius:10px; padding:10px 14px; margin-top:8px;
    }
    .company-banner {
        text-align:center; padding: 6px 8px 14px 8px;
    }
    .company-banner h1 { margin-bottom: 2px; font-size: 1.9rem; }
    .company-banner .addr { color:#555; font-weight:600; margin:0 0 4px 0; }
    .company-banner .tag  { color:#888; margin:0; font-size:0.95rem; }

    /* Let wide tables/dataframes scroll horizontally instead of squeezing on small screens */
    .stDataFrame, .stDataEditor { overflow-x: auto; }

    /* ---- Mobile / Android phone (narrow viewport) ---- */
    @media (max-width: 640px) {
        .main > div { padding-top: 0.6rem; padding-left: 0.6rem; padding-right: 0.6rem; }
        .company-banner h1 { font-size: 1.35rem; }
        .company-banner .addr { font-size: 0.85rem; }
        .company-banner .tag  { font-size: 0.8rem; }
        .req-card { padding: 10px 12px; font-size: 0.9rem; }
        div.stButton > button, div.stDownloadButton > button { width: 100%; font-size: 0.95rem; }
        [data-testid="stMetricValue"] { font-size: 1.3rem; }
        .badge-pending, .badge-approved, .badge-rejected { padding: 3px 9px; font-size: 0.8rem; }
    }
</style>
""", unsafe_allow_html=True)


def company_header(subtitle: str = ""):
    """Reusable company name + address banner, shown at the top of every dashboard."""
    st.markdown(
        f"""
        <div class="company-banner">
            <h1>🚗 {COMPANY_NAME}</h1>
            <p class="addr">📍 {COMPANY_ADDRESS}</p>
            {f'<p class="tag">{subtitle}</p>' if subtitle else ''}
        </div>
        """,
        unsafe_allow_html=True,
    )


def badge_class(status: str) -> str:
    return {"Pending": "badge-pending", "Approved": "badge-approved", "Rejected": "badge-rejected"}.get(status, "badge-pending")


def hash_password(raw: str) -> str:
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


# =========================================================
# 3. SUPABASE CONNECTION
# =========================================================
@st.cache_resource(show_spinner="Connecting to Supabase...")
def get_supabase_client() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


def check_tables_ready() -> tuple[bool, str]:
    """Verify the users/requisitions tables exist and are reachable."""
    sb = get_supabase_client()
    try:
        sb.table(USERS_TABLE).select("id").limit(1).execute()
        sb.table(REQUISITIONS_TABLE).select("id").limit(1).execute()
        return True, ""
    except Exception as e:
        return False, str(e)


# ---------------------- USERS TABLE HELPERS ----------------------
def get_user_by_username(username: str):
    sb = get_supabase_client()
    res = sb.table(USERS_TABLE).select("*").eq("username", username).limit(1).execute()
    return res.data[0] if res.data else None


def register_user(data: dict):
    sb = get_supabase_client()
    sb.table(USERS_TABLE).insert(data).execute()


def update_user(username: str, updates: dict):
    sb = get_supabase_client()
    sb.table(USERS_TABLE).update(updates).eq("username", username).execute()


def delete_user(username: str):
    """Permanently remove a user account (e.g. after they leave the company)."""
    sb = get_supabase_client()
    sb.table(USERS_TABLE).delete().eq("username", username).execute()


def fetch_all_users() -> pd.DataFrame:
    sb = get_supabase_client()
    res = sb.table(USERS_TABLE).select("*").order("created_at", desc=True).execute()
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(
        columns=["id", "username", "full_name", "designation", "employee_id", "department",
                 "mobile", "role", "status", "created_at"]
    )


# ------------------- REQUISITIONS TABLE HELPERS -------------------
def generate_request_id() -> str:
    return f"REQ-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(100, 999)}"


def insert_requisition(data: dict):
    sb = get_supabase_client()
    sb.table(REQUISITIONS_TABLE).insert(data).execute()


def update_requisition(request_id: str, updates: dict):
    sb = get_supabase_client()
    sb.table(REQUISITIONS_TABLE).update(updates).eq("request_id", request_id).execute()


def fetch_all_requisitions() -> pd.DataFrame:
    sb = get_supabase_client()
    res = sb.table(REQUISITIONS_TABLE).select("*").order("created_at", desc=True).execute()
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=[
        "id", "request_id", "created_at", "username", "applicant_name", "department", "mobile_number",
        "date_of_travel", "time_of_travel", "destination", "passenger_count", "vehicle_type", "purpose",
        "special_request", "status", "driver_name", "driver_contact", "vehicle_number", "approved_by",
        "action_timestamp",
    ])


def fetch_requisitions_by_user(username: str) -> pd.DataFrame:
    """Server-side filtered fetch — strict data isolation: only this user's rows are ever requested."""
    sb = get_supabase_client()
    res = (
        sb.table(REQUISITIONS_TABLE)
        .select("*")
        .eq("username", username)
        .order("created_at", desc=True)
        .execute()
    )
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=[
        "id", "request_id", "created_at", "username", "applicant_name", "department", "mobile_number",
        "date_of_travel", "time_of_travel", "destination", "passenger_count", "vehicle_type", "purpose",
        "special_request", "status", "driver_name", "driver_contact", "vehicle_number", "approved_by",
        "action_timestamp",
    ])


# =========================================================
# 4. AUTH: SIGN IN + SELF REGISTRATION
# =========================================================
def get_bootstrap_admin():
    """Fallback super-admin defined in secrets, used only to bootstrap the very
    first login before any 'admin' role exists in the users table."""
    try:
        return st.secrets["admin"]["username"], st.secrets["admin"]["password"]
    except Exception:
        return None, None


def attempt_login(username: str, password: str):
    boot_user, boot_pass = get_bootstrap_admin()
    if boot_user and username == boot_user and password == boot_pass:
        return {"username": username, "full_name": "Super Admin (Bootstrap)", "role": "admin",
                "department": "Management", "designation": "System Administrator"}

    record = get_user_by_username(username)
    if not record:
        return None
    if record.get("status") != "Approved":
        st.error(f"⏳ Your account status is **{record.get('status', 'Pending')}**. Please wait for admin approval.")
        return "PENDING"
    if record.get("password") != hash_password(password):
        return None
    return {
        "username": record["username"],
        "full_name": record.get("full_name", username),
        "role": record.get("role", "user"),
        "department": record.get("department", ""),
        "designation": record.get("designation", ""),
    }


def login_view():
    company_header("Smart Vehicle Management System — Sign in or request a new account")
    ready, err = check_tables_ready()
    if not ready:
        st.error(
            "⚠️ Could not reach the `users` / `requisitions` tables in Supabase. "
            "Make sure you've run the setup SQL from `SETUP_GUIDE.md` in the Supabase SQL Editor, "
            "and that `SUPABASE_URL` / `SUPABASE_KEY` in secrets are correct."
        )
        with st.expander("Technical details"):
            st.code(err)
        st.stop()

    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        tab_login, tab_register = st.tabs(["🔐 Sign In", "📝 Request New User ID"])

        with tab_login:
            with st.form("login_form"):
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                submitted = st.form_submit_button("Login", type="primary", use_container_width=True)

            if submitted:
                result = attempt_login(username.strip(), password)
                if result == "PENDING":
                    pass  # error already shown
                elif result is None:
                    st.error("❌ Invalid username or password.")
                else:
                    st.session_state.auth_user = result
                    st.rerun()

        with tab_register:
            st.caption("New employees can request an account here. Your default password will be your **Employee ID**. "
                       "An admin must approve your account before you can log in.")
            with st.form("register_form", clear_on_submit=True):
                full_name = st.text_input("Full Name *")
                username_r = st.text_input("Choose a Username *")
                designation = st.text_input("Designation *")
                employee_id = st.text_input("Employee ID *", help="This will also be your default password.")
                department = st.selectbox("Department *", DEPARTMENTS)
                mobile = st.text_input("Mobile Number *")
                reg_submitted = st.form_submit_button("Submit Request", type="primary", use_container_width=True)

            if reg_submitted:
                errors = []
                if not full_name.strip():
                    errors.append("Full Name is required.")
                if not username_r.strip():
                    errors.append("Username is required.")
                if not employee_id.strip():
                    errors.append("Employee ID is required.")
                if not mobile.strip():
                    errors.append("Mobile Number is required.")
                if not errors and get_user_by_username(username_r.strip()):
                    errors.append("This username is already taken. Please choose another.")

                if errors:
                    for e in errors:
                        st.error(e)
                else:
                    register_user({
                        "username": username_r.strip(),
                        "password": hash_password(employee_id.strip()),
                        "full_name": full_name.strip(),
                        "designation": designation.strip(),
                        "employee_id": employee_id.strip(),
                        "department": department,
                        "mobile": mobile.strip(),
                        "role": "user",
                        "status": "Pending",
                    })
                    st.success(
                        "✅ Your account request has been submitted! Your default password is your "
                        f"**Employee ID ({employee_id.strip()})**. Please wait for admin approval before signing in."
                    )


def logout_button():
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        del st.session_state["auth_user"]
        st.rerun()


# =========================================================
# 5. EXCEL + PDF REPORT GENERATORS
# =========================================================
def build_excel_report(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Requisitions")
    return buf.getvalue()


class ReportPDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(15, 98, 254)
        self.cell(0, 9, COMPANY_NAME, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.set_font("Helvetica", "", 10)
        self.set_text_color(90, 90, 90)
        self.cell(0, 6, COMPANY_ADDRESS, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.set_font("Helvetica", "B", 12)
        self.set_text_color(0, 0, 0)
        self.cell(0, 8, "Vehicle Requisition Report", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")


def build_pdf_report(df: pd.DataFrame, filters_summary: str) -> bytes:
    pdf = ReportPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
              new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.multi_cell(0, 6, filters_summary)
    pdf.ln(2)

    total = len(df)
    approved = int((df["status"] == "Approved").sum()) if not df.empty else 0
    rejected = int((df["status"] == "Rejected").sum()) if not df.empty else 0
    pending = int((df["status"] == "Pending").sum()) if not df.empty else 0

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Summary Statistics", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", size=9)
    pdf.set_fill_color(230, 230, 230)
    for label, value in [("Total", total), ("Approved", approved), ("Rejected", rejected), ("Pending", pending)]:
        pdf.cell(45, 7, f"{label}: {value}", border=1, align="C", fill=True)
    pdf.ln(12)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Requisition Records", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    headers = ["Req ID", "Applicant", "Department", "Date", "Time", "Destination",
               "Vehicle", "Status", "Driver", "Contact", "Vehicle No."]
    cols = ["request_id", "applicant_name", "department", "date_of_travel", "time_of_travel",
            "destination", "vehicle_type", "status", "driver_name", "driver_contact", "vehicle_number"]
    widths = [30, 30, 28, 20, 15, 30, 22, 20, 28, 30, 24]

    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(15, 98, 254)
    pdf.set_text_color(255, 255, 255)
    for h, w in zip(headers, widths):
        pdf.cell(w, 7, h, border=1, align="C", fill=True)
    pdf.ln()

    pdf.set_font("Helvetica", size=8)
    pdf.set_text_color(0, 0, 0)
    fill = False
    for _, row in df.iterrows():
        pdf.set_fill_color(245, 245, 245)
        for col, w in zip(cols, widths):
            val = str(row.get(col, "") or "")
            if len(val) > 22:
                val = val[:20] + "…"
            pdf.cell(w, 7, val, border=1, fill=fill)
        pdf.ln()
        fill = not fill

    return bytes(pdf.output())


# =========================================================
# 6. SESSION STATE INIT
# =========================================================
if "auth_user" not in st.session_state:
    login_view()
    st.stop()

user = st.session_state.auth_user

# =========================================================
# 7. SIDEBAR
# =========================================================
st.sidebar.title(f"🚗 {COMPANY_NAME}")
st.sidebar.caption(f"📍 {COMPANY_ADDRESS}")
st.sidebar.markdown("---")
st.sidebar.markdown(f"**{user['full_name']}**")
st.sidebar.caption(f"Role: {user['role'].capitalize()}")
if st.sidebar.button("🔄 Refresh Data", use_container_width=True):
    st.rerun()
st.sidebar.markdown("---")
logout_button()

# =========================================================
# 8. GENERAL USER DASHBOARD
# =========================================================
if user["role"] != "admin":
    company_header("👤 Employee Dashboard")
    tab1, tab2 = st.tabs(["📋 New Requisition", "📍 My Requests / Live Status"])

    with tab1:
        st.subheader("Submit a New Vehicle Requisition")
        with st.form("req_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                applicant_name = st.text_input("Applicant Name *", value=user["full_name"])
                department = st.selectbox("Department *", DEPARTMENTS,
                                           index=DEPARTMENTS.index(user["department"]) if user.get("department") in DEPARTMENTS else 0)
                mobile_number = st.text_input("Mobile Number *", placeholder="01XXXXXXXXX")
                passenger_count = st.number_input("Passenger Count *", min_value=1, max_value=50, value=1)
            with c2:
                date_of_travel = st.date_input("Date of Travel *", min_value=date.today())
                time_of_travel = st.time_input("Time of Travel *")
                destination = st.text_input("Destination *")
                vehicle_type = st.selectbox("Vehicle Type Required *", VEHICLE_TYPES)

            purpose = st.text_area("Purpose of Travel *", height=90)
            special_request = st.text_area("Special Request (optional)", height=70)
            submitted = st.form_submit_button("🚀 Submit Requisition", type="primary", use_container_width=True)

        if submitted:
            errors = []
            if not applicant_name.strip():
                errors.append("Applicant Name is required.")
            if not mobile_number.strip():
                errors.append("Mobile Number is required.")
            if not destination.strip():
                errors.append("Destination is required.")
            if not purpose.strip():
                errors.append("Purpose of Travel is required.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                request_id = generate_request_id()
                data = {
                    "request_id": request_id,
                    "username": user["username"],
                    "applicant_name": applicant_name.strip(),
                    "department": department,
                    "mobile_number": mobile_number.strip(),
                    "date_of_travel": str(date_of_travel),
                    "time_of_travel": time_of_travel.strftime("%H:%M"),
                    "destination": destination.strip(),
                    "passenger_count": int(passenger_count),
                    "vehicle_type": vehicle_type,
                    "purpose": purpose.strip(),
                    "special_request": special_request.strip(),
                    "status": "Pending",
                    "driver_name": "",
                    "driver_contact": "",
                    "vehicle_number": "",
                    "approved_by": "",
                }
                with st.spinner("Saving to Supabase..."):
                    try:
                        insert_requisition(data)
                        st.success(f"✅ Requisition submitted! Your Request ID is **{request_id}**")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Failed to save requisition: {e}")

    with tab2:
        st.subheader("My Requests — Live Status")
        with st.spinner("Loading your requests..."):
            my_df = fetch_requisitions_by_user(user["username"])
        if my_df.empty:
            st.info("You haven't submitted any requisitions yet.")
        else:
            for _, r in my_df.iterrows():
                with st.container():
                    st.markdown(f"""
                    <div class="req-card">
                        <b>{r['request_id']}</b> &nbsp;|&nbsp; {r['destination']} &nbsp;|&nbsp;
                        {r['date_of_travel']} at {r['time_of_travel']} &nbsp;&nbsp;
                        <span class="{badge_class(r['status'])}">{STATUS_BADGE.get(r['status'], r['status'])}</span>
                    </div>
                    """, unsafe_allow_html=True)
                    with st.expander("View details"):
                        st.write(f"**Department:** {r['department']}  |  **Vehicle Type:** {r['vehicle_type']}  |  **Passengers:** {r['passenger_count']}")
                        st.write(f"**Purpose:** {r['purpose']}")
                        if r["special_request"]:
                            st.write(f"**Special Request:** {r['special_request']}")
                        if r["status"] == "Approved":
                            st.markdown(f"""
                            <div class="driver-box">
                            🚘 <b>Driver:</b> {r['driver_name'] or 'TBD'} &nbsp;|&nbsp;
                            📞 <b>Contact:</b> {r['driver_contact'] or 'TBD'} &nbsp;|&nbsp;
                            🔢 <b>Vehicle No.:</b> {r['vehicle_number'] or 'TBD'}
                            </div>
                            """, unsafe_allow_html=True)
                        elif r["status"] == "Rejected":
                            st.error("This request was rejected by the admin.")

# =========================================================
# 9. ADMIN DASHBOARD
# =========================================================
else:
    company_header("🔐 Admin Dashboard")
    st.caption(f"Logged in as {user['full_name']} — Admin")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "⏳ Pending User Approvals", "👥 All Users", "🚗 Pending Requisitions",
        "📊 Analytics", "📁 All Requisitions & Export",
    ])

    # ---------------- TAB 1: Pending User Approvals ----------------
    with tab1:
        st.subheader("New Account Requests")
        with st.spinner("Loading users..."):
            users_df = fetch_all_users()
        pending_users = users_df[users_df["status"] == "Pending"] if not users_df.empty else users_df

        if pending_users.empty:
            st.success("🎉 No pending user registrations.")
        else:
            for _, u in pending_users.iterrows():
                with st.expander(f"👤 {u['full_name']} — @{u['username']} ({u.get('department', '')})"):
                    st.write(f"**Designation:** {u.get('designation', '')}")
                    st.write(f"**Employee ID:** {u.get('employee_id', '')}")
                    st.write(f"**Mobile:** {u.get('mobile', '')}")
                    st.write(f"**Requested on:** {u.get('created_at', '')}")
                    c1, c2 = st.columns(2)
                    if c1.button("✅ Approve", key=f"appr_{u['username']}", type="primary", use_container_width=True):
                        update_user(u["username"], {"status": "Approved"})
                        st.success(f"{u['full_name']} approved.")
                        st.rerun()
                    if c2.button("❌ Reject", key=f"rej_{u['username']}", use_container_width=True):
                        update_user(u["username"], {"status": "Rejected"})
                        st.warning(f"{u['full_name']} rejected.")
                        st.rerun()

    # ---------------- TAB 2: All Users (management) ----------------
    with tab2:
        st.subheader("All User Accounts")
        if users_df.empty:
            st.info("No users yet.")
        else:
            display_cols = ["username", "full_name", "designation", "employee_id", "department",
                             "mobile", "role", "status", "created_at"]
            st.dataframe(users_df[[c for c in display_cols if c in users_df.columns]],
                         use_container_width=True, hide_index=True, height=300)

            st.markdown("##### Change a user's role or status")
            approved_usernames = users_df["username"].tolist()
            sel_user = st.selectbox("Select username", approved_usernames)
            if sel_user:
                rec = users_df[users_df["username"] == sel_user].iloc[0]
                c1, c2, c3 = st.columns(3)
                with c1:
                    new_role = st.selectbox("Role", ["user", "admin"],
                                             index=["user", "admin"].index(rec.get("role", "user")))
                with c2:
                    new_status = st.selectbox("Status", STATUS_OPTIONS,
                                               index=STATUS_OPTIONS.index(rec.get("status", "Pending")) if rec.get("status") in STATUS_OPTIONS else 0)
                with c3:
                    st.write("")
                    st.write("")
                    if st.button("💾 Save Changes", use_container_width=True):
                        update_user(sel_user, {"role": new_role, "status": new_status})
                        st.success(f"Updated {sel_user}.")
                        st.rerun()

                st.markdown("---")
                st.markdown("##### 🗑️ Delete User (e.g. employee left the company)")
                if sel_user == user["username"]:
                    st.warning("You can't delete your own currently logged-in account.")
                else:
                    st.caption(
                        f"This permanently removes **{rec.get('full_name', sel_user)}** (@{sel_user}) from the system. "
                        "Their past requisition history will remain in **All Requisitions & Export**, but they will "
                        "no longer be able to log in or submit new requests. This cannot be undone."
                    )
                    confirm_delete = st.checkbox(
                        f"I understand this will permanently delete @{sel_user}'s account.",
                        key=f"confirm_del_{sel_user}",
                    )
                    if st.button("🗑️ Delete This User", type="primary", disabled=not confirm_delete,
                                 use_container_width=True, key=f"del_btn_{sel_user}"):
                        try:
                            delete_user(sel_user)
                            st.success(f"@{sel_user} has been deleted.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to delete user: {e}")

    # ---------------- TAB 3: Pending Requisitions ----------------
    with tab3:
        st.subheader("Requisitions Awaiting Action")
        with st.spinner("Loading requisitions..."):
            df_all = fetch_all_requisitions()
        pending_df = df_all[df_all["status"] == "Pending"] if not df_all.empty else df_all

        if pending_df.empty:
            st.success("🎉 No pending requisitions — all caught up!")
        else:
            for _, r in pending_df.iterrows():
                with st.expander(f"🟡 {r['request_id']} — {r['applicant_name']} ({r['department']}) → {r['destination']}"):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.write(f"**Mobile:** {r['mobile_number']}")
                        st.write(f"**Date/Time:** {r['date_of_travel']} at {r['time_of_travel']}")
                        st.write(f"**Passengers:** {r['passenger_count']}")
                    with c2:
                        st.write(f"**Vehicle Type:** {r['vehicle_type']}")
                        st.write(f"**Purpose:** {r['purpose']}")
                        st.write(f"**Special Request:** {r['special_request'] or '—'}")

                    with st.form(f"action_{r['request_id']}"):
                        d1, d2, d3 = st.columns(3)
                        with d1:
                            driver_name = st.text_input("Driver Name", key=f"dn_{r['request_id']}")
                        with d2:
                            driver_contact = st.text_input("Driver Contact", key=f"dc_{r['request_id']}")
                        with d3:
                            vehicle_number = st.text_input("Vehicle Number", key=f"vn_{r['request_id']}")

                        b1, b2 = st.columns(2)
                        approve_clicked = b1.form_submit_button("✅ Approve", type="primary", use_container_width=True)
                        reject_clicked = b2.form_submit_button("❌ Reject", use_container_width=True)

                    if approve_clicked or reject_clicked:
                        new_status = "Approved" if approve_clicked else "Rejected"
                        if approve_clicked and not (driver_name and driver_contact and vehicle_number):
                            st.error("Please fill Driver Name, Driver Contact, and Vehicle Number before approving.")
                        else:
                            try:
                                update_requisition(r["request_id"], {
                                    "status": new_status,
                                    "driver_name": driver_name if approve_clicked else "",
                                    "driver_contact": driver_contact if approve_clicked else "",
                                    "vehicle_number": vehicle_number if approve_clicked else "",
                                    "approved_by": user["full_name"],
                                    "action_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                })
                                st.success(f"Request {r['request_id']} marked as {new_status}.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Update failed: {e}")

    # ---------------- TAB 4: Analytics ----------------
    with tab4:
        st.subheader("📊 Visual Analytics")
        if df_all.empty:
            st.info("No data yet.")
        else:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Requests", len(df_all))
            m2.metric("🟢 Approved", int((df_all["status"] == "Approved").sum()))
            m3.metric("🔴 Rejected", int((df_all["status"] == "Rejected").sum()))
            m4.metric("🟡 Pending", int((df_all["status"] == "Pending").sum()))

            c1, c2 = st.columns(2)
            with c1:
                dept_counts = df_all["department"].value_counts().reset_index()
                dept_counts.columns = ["Department", "Requests"]
                fig1 = px.bar(dept_counts, x="Department", y="Requests", color="Department", text="Requests",
                              title="Department-wise Vehicle Usage")
                fig1.update_layout(showlegend=False, height=380)
                st.plotly_chart(fig1, use_container_width=True)
            with c2:
                status_counts = df_all["status"].value_counts().reset_index()
                status_counts.columns = ["Status", "Count"]
                fig2 = px.pie(status_counts, names="Status", values="Count", hole=0.45, title="Status Breakdown",
                              color="Status",
                              color_discrete_map={"Approved": "#28a745", "Rejected": "#dc3545", "Pending": "#ffc107"})
                fig2.update_layout(height=380)
                st.plotly_chart(fig2, use_container_width=True)

            df_all["_dt"] = pd.to_datetime(df_all["date_of_travel"], errors="coerce")
            monthly = df_all.dropna(subset=["_dt"]).copy()
            if not monthly.empty:
                monthly["Month"] = monthly["_dt"].dt.to_period("M").astype(str)
                monthly_counts = monthly.groupby("Month").size().reset_index(name="Requests")
                fig3 = px.line(monthly_counts, x="Month", y="Requests", markers=True, title="Monthly Request Trend")
                fig3.update_layout(height=350)
                st.plotly_chart(fig3, use_container_width=True)

    # ---------------- TAB 5: All Requisitions + Filters + Export ----------------
    with tab5:
        st.subheader("📁 All Requisitions — Search, Filter & Export")
        if df_all.empty:
            st.info("No data yet.")
        else:
            f1, f2, f3 = st.columns(3)
            with f1:
                dept_filter = st.multiselect("Department", sorted(df_all["department"].dropna().unique().tolist()))
            with f2:
                status_filter = st.multiselect("Status", STATUS_OPTIONS)
            with f3:
                dest_filter = st.text_input("Destination contains")

            df_all["_dt"] = pd.to_datetime(df_all["date_of_travel"], errors="coerce")
            min_d = df_all["_dt"].min()
            max_d = df_all["_dt"].max()
            default_start = min_d.date() if pd.notnull(min_d) else date.today()
            default_end = max_d.date() if pd.notnull(max_d) else date.today()
            date_range = st.date_input("Date of Travel range", value=(default_start, default_end))

            filtered = df_all.copy()
            if dept_filter:
                filtered = filtered[filtered["department"].isin(dept_filter)]
            if status_filter:
                filtered = filtered[filtered["status"].isin(status_filter)]
            if dest_filter:
                filtered = filtered[filtered["destination"].str.contains(dest_filter, case=False, na=False)]
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_d, end_d = date_range
                filtered = filtered[(filtered["_dt"] >= pd.Timestamp(start_d)) & (filtered["_dt"] <= pd.Timestamp(end_d))]

            filtered_display = filtered.drop(columns=["_dt"], errors="ignore")
            st.dataframe(filtered_display, use_container_width=True, hide_index=True, height=340)

            filters_summary = (
                f"Departments: {', '.join(dept_filter) if dept_filter else 'All'} | "
                f"Status: {', '.join(status_filter) if status_filter else 'All'} | "
                f"Destination filter: {dest_filter or 'None'}"
            )

            colA, colB = st.columns(2)
            with colA:
                excel_bytes = build_excel_report(filtered_display)
                st.download_button("⬇️ Export to Excel (.xlsx)", data=excel_bytes,
                                    file_name="vehicle_requisition_report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True)
            with colB:
                pdf_bytes = build_pdf_report(filtered_display, filters_summary)
                st.download_button("⬇️ Export to PDF (.pdf)", data=pdf_bytes,
                                    file_name="vehicle_requisition_report.pdf",
                                    mime="application/pdf",
                                    use_container_width=True)