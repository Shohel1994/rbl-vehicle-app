"""
Smart Vehicle Management System — Supabase (PostgreSQL) Edition
==================================================================
Enterprise-grade backend: Supabase Postgres via the official `supabase-py`
SDK, replacing SQLite. Adds self-service user registration with admin
approval, in addition to requisition approval. No email — everything
happens live inside the app. Exports: Excel (.xlsx) and PDF (.pdf).

Author: Senior Python Developer (generated for Shohel Rana)
"""

import io
import os
import base64
import hashlib
import random
from datetime import datetime, date, timedelta
from secrets import token_urlsafe

import pandas as pd
import plotly.express as px
import streamlit as st
from supabase import create_client, Client
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from fpdf.fonts import FontFace
from streamlit_autorefresh import st_autorefresh
import extra_streamlit_components as stx

# =========================================================
# 1. COMPANY INFO & PAGE CONFIG
# =========================================================
COMPANY_NAME = "Renaissaince Barind Ltd."
COMPANY_ADDRESS = "Ishwardi EPZ, Pakshi, Pabna"

st.set_page_config(
    page_title="RBL VMS",
    page_icon="logo.png",
    layout="wide",
    initial_sidebar_state="expanded",
)
# সাইডবারে লোগো দেখানোর জন্য
st.sidebar.image("logo.png", use_container_width=True)
USERS_TABLE = "users"
REQUISITIONS_TABLE = "requisitions"
DRIVERS_TABLE = "drivers"
VEHICLES_TABLE = "vehicles"
SESSIONS_TABLE = "sessions"

SESSION_COOKIE_NAME = "rbl_vms_session"
SESSION_LIFETIME_DAYS = 30
LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo.png")

VEHICLE_TYPES = ["Private Car", "HIACE", "Pick-up Van", "Covered Van", "Truck", "Shipment Vehicle", "Other"]
DEPARTMENTS = [
    "Accounts", "Warehouse", "Factory Merchandising", "Commercial", "Floor Operation",
    "TSD", "QMS", "Production Planning & Control", "Cutting", "R & D", "Admin",
    "Technical", "Finishing", "Quality", "IE", "HR & Compliance", "Other",
]

# Account status (users table) — approval workflow for logins
USER_STATUS_OPTIONS = ["Pending", "Approved", "Rejected"]
ROLE_OPTIONS = ["user", "security_officer", "admin"]
ROLE_DISPLAY = {"user": "Employee", "security_officer": "Security Officer", "admin": "Admin"}

# Requisition status (requisitions table) — full trip lifecycle
REQ_STATUS_OPTIONS = ["Pending", "Approved", "Rejected", "On Trip", "Completed"]
STATUS_BADGE = {
    "Pending": "🟡 Pending", "Approved": "🟢 Approved", "Rejected": "🔴 Rejected",
    "On Trip": "🔵 On Trip", "Completed": "✅ Completed",
}

# =========================================================
# 2. STYLING (desktop / Windows browser + mobile / Android friendly)
# =========================================================
st.markdown("""
<style>
    /* ---- Hide default Streamlit branding/chrome ----
       Deliberately NOT touching [data-testid="collapsedControl"] (the sidebar
       hamburger toggle) — that lives outside <header> in current Streamlit
       versions, so hiding the header/menu/footer/toolbar below does not
       affect the mobile sidebar-open control. */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header { visibility: hidden; }
    [data-testid="stToolbar"] { visibility: hidden; }
    .stDeployButton { display: none; }
    [data-testid="collapsedControl"] { visibility: visible; }

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
    .badge-ontrip   { background:#D6E9FF; color:#084298; padding:4px 12px; border-radius:14px; font-weight:700; white-space:nowrap; }
    .badge-completed{ background:#D1F2EB; color:#0B5345; padding:4px 12px; border-radius:14px; font-weight:700; white-space:nowrap; }
    .driver-box {
        background:#EAF7EE; border:1px solid #B7E4C7; border-radius:10px; padding:10px 14px; margin-top:8px;
    }
    .company-banner {
        text-align:center; padding: 6px 8px 14px 8px;
    }
    .company-banner .logo-row {
        display:flex; align-items:center; justify-content:center; gap:12px; margin-bottom:2px;
    }
    .company-banner .logo-row img { height:44px; width:auto; }
    .company-banner h1 { margin: 0; font-size: 1.9rem; }
    .company-banner .addr { color:#555; font-weight:600; margin:0 0 4px 0; }
    .company-banner .tag  { color:#888; margin:0; font-size:0.95rem; }

    .sidebar-brand {
        display:flex; align-items:center; gap:10px; margin-bottom:4px;
    }
    .sidebar-brand img { height:32px; width:auto; }
    .sidebar-brand span { font-size:1.15rem; font-weight:700; line-height:1.2; }

    /* Let wide tables/dataframes scroll horizontally instead of squeezing on small screens */
    .stDataFrame, .stDataEditor { overflow-x: auto; }

    /* ---- Mobile / Android phone (narrow viewport) ---- */
    @media (max-width: 640px) {
        .main > div { padding-top: 0.6rem; padding-left: 0.6rem; padding-right: 0.6rem; }
        .company-banner h1 { font-size: 1.35rem; }
        .company-banner .logo-row img { height:32px; }
        .company-banner .addr { font-size: 0.85rem; }
        .company-banner .tag  { font-size: 0.8rem; }
        .req-card { padding: 10px 12px; font-size: 0.9rem; }
        div.stButton > button, div.stDownloadButton > button { width: 100%; font-size: 0.95rem; }
        [data-testid="stMetricValue"] { font-size: 1.3rem; }
        .badge-pending, .badge-approved, .badge-rejected { padding: 3px 9px; font-size: 0.8rem; }
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def get_logo_base64():
    """Read logo.png (same folder as app.py) once and cache it as a base64 data
    URI, so it can be embedded inline in HTML headers. Returns None if the file
    isn't present — callers fall back to text-only branding rather than crash
    or show a broken-image icon."""
    try:
        with open(LOGO_PATH, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        return f"data:image/png;base64,{encoded}"
    except FileNotFoundError:
        return None


def company_header(subtitle: str = ""):
    """Reusable company name + address banner, shown at the top of every dashboard."""
    logo_uri = get_logo_base64()
    logo_img = f'<img src="{logo_uri}" alt="logo">' if logo_uri else ""
    st.markdown(
        f"""
        <div class="company-banner">
            <div class="logo-row">
                {logo_img}
                <h1>{COMPANY_NAME}</h1>
            </div>
            <p class="addr">📍 {COMPANY_ADDRESS}</p>
            {f'<p class="tag">{subtitle}</p>' if subtitle else ''}
        </div>
        """,
        unsafe_allow_html=True,
    )


def badge_class(status: str) -> str:
    return {
        "Pending": "badge-pending", "Approved": "badge-approved", "Rejected": "badge-rejected",
        "On Trip": "badge-ontrip", "Completed": "badge-completed",
    }.get(status, "badge-pending")


def is_blank(val) -> bool:
    """True if val is None, NaN, or an empty/whitespace string.

    Pandas turns SQL NULLs (from nullable Supabase columns like approved_time,
    admin_note, start_km) into float NaN when loaded into a DataFrame — and NaN
    is truthy in plain Python (`if nan:` is True), so a bare `if val:` or
    `val or fallback` silently does the wrong thing for unset fields. Every
    place that reads an optional trip-lifecycle field goes through this check.
    """
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str) and not val.strip():
        return True
    return False


def fmt(val, default: str = "—") -> str:
    """Display-safe string formatting: shows `default` instead of a literal
    'nan'/'None' for unset nullable fields (see is_blank)."""
    return default if is_blank(val) else str(val)


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
        "action_timestamp", "approved_time", "admin_note", "start_km", "end_km", "total_km",
        "actual_exit_time", "actual_return_time",
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
        "action_timestamp", "approved_time", "admin_note", "start_km", "end_km", "total_km",
        "actual_exit_time", "actual_return_time",
    ])


def fetch_requisitions_by_status(status: str) -> pd.DataFrame:
    """Server-side filtered fetch used by the Security Officer panel — only pulls
    requisitions in the given trip-status (e.g. 'Approved' or 'On Trip'), so
    Pending/Rejected requests are never even requested, let alone shown."""
    sb = get_supabase_client()
    res = (
        sb.table(REQUISITIONS_TABLE)
        .select("*")
        .eq("status", status)
        .order("created_at", desc=True)
        .execute()
    )
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=[
        "id", "request_id", "created_at", "username", "applicant_name", "department", "mobile_number",
        "date_of_travel", "time_of_travel", "destination", "passenger_count", "vehicle_type", "purpose",
        "special_request", "status", "driver_name", "driver_contact", "vehicle_number", "approved_by",
        "action_timestamp", "approved_time", "admin_note", "start_km", "end_km", "total_km",
        "actual_exit_time", "actual_return_time",
    ])


# ------------------- DRIVERS & VEHICLES TABLE HELPERS -------------------
# These back the dynamic dropdowns in the requisition-approval form so admins
# maintain one source of truth instead of retyping names/numbers each time.
def fetch_all_drivers() -> pd.DataFrame:
    sb = get_supabase_client()
    res = sb.table(DRIVERS_TABLE).select("*").order("driver_name").execute()
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=["id", "driver_name", "driver_contact", "created_at"])


def add_driver(driver_name: str, driver_contact: str):
    sb = get_supabase_client()
    sb.table(DRIVERS_TABLE).insert({"driver_name": driver_name, "driver_contact": driver_contact}).execute()


def delete_driver(driver_id):
    sb = get_supabase_client()
    sb.table(DRIVERS_TABLE).delete().eq("id", driver_id).execute()


def fetch_all_vehicles() -> pd.DataFrame:
    sb = get_supabase_client()
    res = sb.table(VEHICLES_TABLE).select("*").order("vehicle_number").execute()
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=["id", "vehicle_number", "created_at"])


def add_vehicle(vehicle_number: str):
    sb = get_supabase_client()
    sb.table(VEHICLES_TABLE).insert({"vehicle_number": vehicle_number}).execute()


def delete_vehicle(vehicle_id):
    sb = get_supabase_client()
    sb.table(VEHICLES_TABLE).delete().eq("id", vehicle_id).execute()


# ------------------- SESSION (REMEMBER ME) HELPERS -------------------
# A "remember me" cookie stores only an opaque, unguessable token — never the
# username or password directly — so a leaked/inspected cookie can't be used
# to reconstruct credentials. The token maps to a username via this table and
# expires automatically, and is revoked (deleted) on explicit logout.
def create_session(username: str) -> str:
    sb = get_supabase_client()
    token = token_urlsafe(32)
    expires_at = (datetime.utcnow() + timedelta(days=SESSION_LIFETIME_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
    sb.table(SESSIONS_TABLE).insert({
        "token": token, "username": username, "expires_at": expires_at,
    }).execute()
    return token


def get_session_username(token: str):
    """Return the username for a still-valid session token, or None."""
    if not token:
        return None
    sb = get_supabase_client()
    res = sb.table(SESSIONS_TABLE).select("*").eq("token", token).limit(1).execute()
    if not res.data:
        return None
    row = res.data[0]
    try:
        # PostgREST reads timestamptz columns back in ISO 8601 with a 'T'
        # separator and often a timezone offset (e.g. "...T14:22:30.123+00:00"),
        # which differs from the space-separated string we wrote on insert —
        # normalize both shapes before parsing so expiry checks don't silently
        # fail (and reject) every real session.
        raw = row["expires_at"].replace("T", " ")
        expires_at = datetime.strptime(raw[:19], "%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError, KeyError, AttributeError):
        return None
    if expires_at < datetime.utcnow():
        return None
    return row["username"]


def delete_session(token: str):
    if not token:
        return
    sb = get_supabase_client()
    sb.table(SESSIONS_TABLE).delete().eq("token", token).execute()


def get_cookie_manager():
    """Returns a CookieManager. Deliberately NOT wrapped in st.cache_resource:
    that cache is shared globally across every visitor's session on the
    server, and caching a stateful per-browser cookie wrapper there would
    leak one user's session cookie into another user's script run. Streamlit's
    component protocol is already session-scoped on its own, so a fresh,
    cheap instantiation on every rerun is the correct (and documented)
    pattern for this library."""
    return stx.CookieManager(key="rbl_vms_cookie_manager")


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


def build_bootstrap_user_dict(username: str) -> dict:
    return {"username": username, "full_name": "Super Admin (Bootstrap)", "role": "admin",
            "department": "Management", "designation": "System Administrator", "mobile": ""}


def build_user_dict(record: dict) -> dict:
    """Shared by both password login and cookie-based session restore, so the
    session_state.auth_user shape never drifts between the two paths."""
    return {
        "username": record["username"],
        "full_name": record.get("full_name", record["username"]),
        "role": record.get("role", "user"),
        "department": record.get("department", ""),
        "designation": record.get("designation", ""),
        "mobile": record.get("mobile", ""),
    }


def attempt_login(username: str, password: str):
    boot_user, boot_pass = get_bootstrap_admin()
    if boot_user and username == boot_user and password == boot_pass:
        return build_bootstrap_user_dict(username)

    record = get_user_by_username(username)
    if not record:
        return None
    if record.get("status") != "Approved":
        st.error(f"⏳ Your account status is **{record.get('status', 'Pending')}**. Please wait for admin approval.")
        return "PENDING"
    if record.get("password") != hash_password(password):
        return None
    return build_user_dict(record)


def restore_user_from_username(username: str):
    """Used only for cookie-based 'remember me' restore — re-validates the
    account is still Approved (in case it was later revoked/rejected) before
    trusting the session token."""
    boot_user, _ = get_bootstrap_admin()
    if boot_user and username == boot_user:
        return build_bootstrap_user_dict(username)
    record = get_user_by_username(username)
    if record and record.get("status") == "Approved":
        return build_user_dict(record)
    return None


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
                remember_me = st.checkbox("Remember me on this device", value=True)
                submitted = st.form_submit_button("Login", type="primary", use_container_width=True)

            if submitted:
                result = attempt_login(username.strip(), password)
                if result == "PENDING":
                    pass  # error already shown
                elif result is None:
                    st.error("❌ Invalid username or password.")
                else:
                    session_token = None
                    if remember_me:
                        try:
                            session_token = create_session(result["username"])
                            cookie_manager.set(
                                SESSION_COOKIE_NAME, session_token, key="set_login_cookie",
                                expires_at=datetime.now() + timedelta(days=SESSION_LIFETIME_DAYS),
                            )
                        except Exception:
                            session_token = None  # Remember-me is best-effort; login still succeeds without it.
                    result["session_token"] = session_token
                    st.session_state.auth_user = result
                    st.rerun()

        with tab_register:
            st.caption("New employees or Security Officers can request an account here. "
                       "An admin must approve your account before you can log in.")

            reg_role = st.radio(
                "Register as", ["Employee", "Security Officer"],
                horizontal=True, key="reg_role_choice",
            )
            st.markdown("---")

            if reg_role == "Employee":
                st.caption("Your default password will be your **Employee ID**.")
                with st.form("register_employee_form", clear_on_submit=True):
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

            else:  # Security Officer registration — same simple flow as normal users
                st.caption("Choose your own username and password below, just like a regular account.")
                with st.form("register_security_form", clear_on_submit=True):
                    sec_full_name = st.text_input("Full Name *", key="sec_full_name")
                    sec_username = st.text_input("Username *", key="sec_username")
                    sec_password = st.text_input("Password *", type="password", key="sec_password")
                    sec_password_confirm = st.text_input("Confirm Password *", type="password", key="sec_password_confirm")
                    sec_submitted = st.form_submit_button("Submit Request", type="primary", use_container_width=True)

                if sec_submitted:
                    errors = []
                    if not sec_full_name.strip():
                        errors.append("Full Name is required.")
                    if not sec_username.strip():
                        errors.append("Username is required.")
                    if not sec_password:
                        errors.append("Password is required.")
                    elif len(sec_password) < 4:
                        errors.append("Password must be at least 4 characters.")
                    elif sec_password != sec_password_confirm:
                        errors.append("Passwords do not match.")
                    if not errors and get_user_by_username(sec_username.strip()):
                        errors.append("This username is already taken. Please choose another.")

                    if errors:
                        for e in errors:
                            st.error(e)
                    else:
                        register_user({
                            "username": sec_username.strip(),
                            "password": hash_password(sec_password),
                            "full_name": sec_full_name.strip(),
                            "designation": "Security Officer",
                            "employee_id": "",
                            "department": "Security",
                            "mobile": "",
                            "role": "security_officer",
                            "status": "Pending",
                        })
                        st.success(
                            "✅ Your Security Officer account request has been submitted! "
                            "Please wait for admin approval before signing in."
                        )


def logout_button():
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        token = st.session_state.get("auth_user", {}).get("session_token")
        if token:
            try:
                delete_session(token)
            except Exception:
                pass  # DB cleanup is best-effort; logout must still proceed either way.
        try:
            cookie_manager.delete(SESSION_COOKIE_NAME, key="delete_logout_cookie")
        except KeyError:
            pass  # No cookie was ever set for this session — nothing to remove.
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
    on_trip = int((df["status"] == "On Trip").sum()) if not df.empty else 0
    completed = int((df["status"] == "Completed").sum()) if not df.empty else 0

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Summary Statistics", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", size=9)
    pdf.set_fill_color(230, 230, 230)
    for label, value in [("Total", total), ("Pending", pending), ("Approved", approved),
                          ("On Trip", on_trip), ("Completed", completed), ("Rejected", rejected)]:
        pdf.cell(38, 7, f"{label}: {value}", border=1, align="C", fill=True)
    pdf.ln(12)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Requisition Records", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)

    # Column widths are sized generously for "Req ID" (the longest field, e.g.
    # "REQ-20260809072230-852") and every cell wraps automatically via fpdf2's
    # table() API — this is what actually prevents text from bleeding into the
    # next column, instead of truncating with "…" as the previous version did.
    headers = ["Req ID", "Applicant", "Department", "Date", "Time", "Destination",
               "Vehicle", "Status", "Driver", "Vehicle No.", "Total KM"]
    cols = ["request_id", "applicant_name", "department", "date_of_travel", "time_of_travel",
            "destination", "vehicle_type", "status", "driver_name", "vehicle_number", "total_km"]
    col_widths = [42, 24, 22, 18, 13, 24, 16, 16, 22, 28, 16]

    pdf.set_font("Helvetica", size=7)
    heading_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 98, 254))
    with pdf.table(col_widths=col_widths, text_align="LEFT", first_row_as_headings=True,
                   line_height=5, headings_style=heading_style, cell_fill_color=(245, 245, 245),
                   cell_fill_mode="ROWS") as table:
        header_row = table.row()
        for h in headers:
            header_row.cell(h)
        for _, r in df.iterrows():
            row = table.row()
            for col in cols:
                row.cell(fmt(r.get(col, ""), ""))

    return bytes(pdf.output())


# =========================================================
# 6. SESSION STATE INIT (with "Remember Me" cookie restore)
# =========================================================
# Instantiated exactly once per script run — the underlying component uses a
# fixed key, and Streamlit errors on duplicate keys within a single run, so
# every other place in this file that needs cookies reuses this same object
# rather than calling get_cookie_manager() again.
cookie_manager = get_cookie_manager()

if "auth_user" not in st.session_state:
    restored_user = None
    session_token = cookie_manager.get(SESSION_COOKIE_NAME)
    if session_token:
        remembered_username = get_session_username(session_token)
        if remembered_username:
            restored_user = restore_user_from_username(remembered_username)
        if restored_user:
            restored_user["session_token"] = session_token
            st.session_state.auth_user = restored_user
        else:
            # Token is missing/expired/revoked, or the account is no longer
            # Approved — clear the stale cookie so we don't keep retrying it.
            try:
                cookie_manager.delete(SESSION_COOKIE_NAME, key="delete_stale_cookie")
            except KeyError:
                pass

    if "auth_user" not in st.session_state:
        login_view()
        st.stop()

user = st.session_state.auth_user

# =========================================================
# 7. SIDEBAR
# =========================================================
logo_uri = get_logo_base64()
logo_img = f'<img src="{logo_uri}" alt="logo">' if logo_uri else ""
st.sidebar.markdown(
    f'<div class="sidebar-brand">{logo_img}<span>{COMPANY_NAME}</span></div>',
    unsafe_allow_html=True,
)
st.sidebar.caption(f"📍 {COMPANY_ADDRESS}")
st.sidebar.markdown("---")
st.sidebar.markdown(f"**{user['full_name']}**")
st.sidebar.caption(f"Role: {ROLE_DISPLAY.get(user['role'], user['role'].capitalize())}")

auto_refresh_on = st.sidebar.checkbox("🔄 Auto-refresh every 10s", value=True,
                                       help="Automatically reloads live data across the app. "
                                            "Turn off temporarily if you're filling out a long form.")
if st.sidebar.button("🔄 Refresh Now", use_container_width=True):
    st.rerun()
st.sidebar.markdown("---")
logout_button()

if auto_refresh_on:
    st_autorefresh(interval=10_000, key="global_autorefresh")

# =========================================================
# 8. EMPLOYEE DASHBOARD
# =========================================================
if user["role"] == "user":
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
                mobile_number = st.text_input("Mobile Number *", value=user.get("mobile", ""),
                                               placeholder="01XXXXXXXXX")
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

                        if r["status"] in ("Approved", "On Trip", "Completed"):
                            approved_time = r["time_of_travel"] if is_blank(r.get("approved_time")) else r.get("approved_time")
                            time_note = (
                                f" (rescheduled from {r['time_of_travel']})"
                                if not is_blank(r.get("approved_time")) and r.get("approved_time") != r["time_of_travel"]
                                else ""
                            )
                            st.markdown(f"""
                            <div class="driver-box">
                            🚘 <b>Driver:</b> {fmt(r['driver_name'], 'TBD')} &nbsp;|&nbsp;
                            📞 <b>Contact:</b> {fmt(r['driver_contact'], 'TBD')} &nbsp;|&nbsp;
                            🔢 <b>Vehicle No.:</b> {fmt(r['vehicle_number'], 'TBD')} &nbsp;|&nbsp;
                            🕒 <b>Approved Departure Time:</b> {approved_time}{time_note}
                            </div>
                            """, unsafe_allow_html=True)
                            if not is_blank(r.get("admin_note")):
                                st.caption(f"📝 Admin Note: {r['admin_note']}")

                        if r["status"] in ("On Trip", "Completed"):
                            st.write(f"**Gate Out (Actual Exit):** {fmt(r.get('actual_exit_time'))}  |  **Start KM:** {fmt(r.get('start_km'))}")
                        if r["status"] == "Completed":
                            st.write(f"**Gate In (Actual Return):** {fmt(r.get('actual_return_time'))}  |  **End KM:** {fmt(r.get('end_km'))}  |  **Total KM:** {fmt(r.get('total_km'))}")

                        if r["status"] == "Rejected":
                            st.error("This request was rejected by the admin.")
                            if not is_blank(r.get("admin_note")):
                                st.caption(f"📝 Reason: {r['admin_note']}")

# =========================================================
# 9. SECURITY OFFICER DASHBOARD — Vehicle Gate In / Gate Out Panel
# =========================================================
elif user["role"] == "security_officer":
    company_header("🛡️ Security Officer Dashboard — Vehicle Gate Panel")
    st.caption(f"Logged in as {user['full_name']} — Security Officer")

    tab_out, tab_in = st.tabs(["🚦 Ready to Depart (Approved Trips)", "🔁 Currently On Trip (Inbound Vehicles)"])

    # ---------------- TAB 1: Ready to Depart ----------------
    with tab_out:
        st.subheader("Approved Trips Awaiting Gate Out")
        with st.spinner("Loading approved trips..."):
            ready_df = fetch_requisitions_by_status("Approved")

        if ready_df.empty:
            st.info("No trips are currently approved and waiting to depart.")
        else:
            for _, r in ready_df.iterrows():
                approved_time = r["time_of_travel"] if is_blank(r.get("approved_time")) else r.get("approved_time")
                with st.expander(f"🟢 {r['applicant_name']} ({r['department']}) → {r['destination']}  |  Vehicle: {fmt(r['vehicle_number'], 'N/A')}"):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.write(f"**Applicant Name:** {r['applicant_name']}")
                        st.write(f"**Department:** {r['department']}")
                        st.write(f"**Vehicle:** {fmt(r['vehicle_number'], 'N/A')} ({r['vehicle_type']})")
                    with c2:
                        st.write(f"**Destination:** {r['destination']}")
                        st.write(f"**Requested Time:** {r['date_of_travel']} at {r['time_of_travel']}")
                        st.write(f"**Admin Approved Time:** {approved_time}")
                    if not is_blank(r.get("admin_note")):
                        st.caption(f"📝 Admin Notes: {r['admin_note']}")

                    with st.form(f"gateout_{r['request_id']}"):
                        start_km = st.number_input("Start KM (Odometer Reading) *", min_value=0.0, step=1.0,
                                                     format="%.1f", key=f"skm_{r['request_id']}")
                        depart_clicked = st.form_submit_button("🚦 Gate Out / Depart", type="primary", use_container_width=True)

                    if depart_clicked:
                        try:
                            update_requisition(r["request_id"], {
                                "start_km": float(start_km),
                                "actual_exit_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "status": "On Trip",
                            })
                            st.success(f"✅ Gate Out recorded for {r['applicant_name']} — vehicle is now On Trip.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to record Gate Out: {e}")

    # ---------------- TAB 2: Currently On Trip ----------------
    with tab_in:
        st.subheader("Vehicles Currently Outside the Gate")
        with st.spinner("Loading active trips..."):
            ontrip_df = fetch_requisitions_by_status("On Trip")

        if ontrip_df.empty:
            st.info("No vehicles are currently on a trip.")
        else:
            for _, r in ontrip_df.iterrows():
                with st.expander(f"🔵 {r['applicant_name']} ({r['department']}) → {r['destination']}  |  Vehicle: {fmt(r['vehicle_number'], 'N/A')}"):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.write(f"**Applicant Name:** {r['applicant_name']}")
                        st.write(f"**Vehicle:** {fmt(r['vehicle_number'], 'N/A')} ({r['vehicle_type']})")
                        st.write(f"**Destination:** {r['destination']}")
                    with c2:
                        st.write(f"**Gate Out Time:** {fmt(r.get('actual_exit_time'))}")
                        st.write(f"**Start KM:** {fmt(r.get('start_km'))}")

                    start_km_val = 0.0 if is_blank(r.get("start_km")) else float(r.get("start_km"))
                    with st.form(f"gatein_{r['request_id']}"):
                        end_km = st.number_input(
                            "End KM (Odometer Reading) *", min_value=start_km_val, step=1.0, format="%.1f",
                            help=f"Must be greater than or equal to Start KM ({start_km_val:.1f}).",
                            key=f"ekm_{r['request_id']}",
                        )
                        return_clicked = st.form_submit_button("🏁 Gate In / Complete", type="primary", use_container_width=True)

                    if return_clicked:
                        if end_km < start_km_val:
                            st.error("End KM cannot be less than Start KM.")
                        else:
                            total_km = round(float(end_km) - start_km_val, 1)
                            try:
                                update_requisition(r["request_id"], {
                                    "end_km": float(end_km),
                                    "total_km": total_km,
                                    "actual_return_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "status": "Completed",
                                })
                                st.success(f"✅ Gate In recorded for {r['applicant_name']}. Total distance: **{total_km} KM**")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Failed to record Gate In: {e}")

# =========================================================
# 10. ADMIN DASHBOARD
# =========================================================
elif user["role"] == "admin":
    company_header("🔐 Admin Dashboard")
    st.caption(f"Logged in as {user['full_name']} — Admin")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "⏳ Pending User Approvals", "👥 All Users", "🚗 Pending Requisitions",
        "📊 Analytics", "📁 All Requisitions & Export", "🚘 Manage Drivers & Vehicles",
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
                role_tag = "🛡️ Security Officer" if u.get("role") == "security_officer" else "👤 Employee"
                with st.expander(f"{role_tag} — {u['full_name']} (@{u['username']})"):
                    if u.get("role") == "security_officer":
                        st.write("**Role Requested:** Security Officer")
                    else:
                        st.write(f"**Designation:** {u.get('designation', '')}")
                        st.write(f"**Employee ID:** {u.get('employee_id', '')}")
                        st.write(f"**Department:** {u.get('department', '')}")
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
            users_display = users_df[[c for c in display_cols if c in users_df.columns]].copy()
            if "role" in users_display.columns:
                users_display["role"] = users_display["role"].map(lambda r: ROLE_DISPLAY.get(r, r))
            st.dataframe(users_display, use_container_width=True, hide_index=True, height=300)

            st.markdown("##### Change a user's role or status")
            approved_usernames = users_df["username"].tolist()
            sel_user = st.selectbox("Select username", approved_usernames)
            if sel_user:
                rec = users_df[users_df["username"] == sel_user].iloc[0]
                c1, c2, c3 = st.columns(3)
                with c1:
                    current_role = rec.get("role", "user")
                    new_role = st.selectbox("Role", ROLE_OPTIONS,
                                             index=ROLE_OPTIONS.index(current_role) if current_role in ROLE_OPTIONS else 0,
                                             format_func=lambda r: ROLE_DISPLAY.get(r, r))
                with c2:
                    new_status = st.selectbox("Status", USER_STATUS_OPTIONS,
                                               index=USER_STATUS_OPTIONS.index(rec.get("status", "Pending")) if rec.get("status") in USER_STATUS_OPTIONS else 0)
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

        # Fetched once for this tab render and shared across every pending-request
        # card below, so each dropdown reflects the same up-to-date driver/vehicle list.
        drivers_df = fetch_all_drivers()
        vehicles_df = fetch_all_vehicles()
        driver_contact_map = dict(zip(drivers_df["driver_name"], drivers_df["driver_contact"])) if not drivers_df.empty else {}
        driver_options = ["— Select Driver —"] + drivers_df["driver_name"].tolist() if not drivers_df.empty else []
        vehicle_options = ["— Select Vehicle —"] + vehicles_df["vehicle_number"].tolist() if not vehicles_df.empty else []

        if drivers_df.empty or vehicles_df.empty:
            st.warning(
                "⚠️ No drivers and/or vehicles are registered yet. Add them under the "
                "**🚘 Manage Drivers & Vehicles** tab before you can approve requests."
            )

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

                    # These two selects live OUTSIDE the form on purpose: widgets inside
                    # an st.form don't rerun the script until submit, so picking a driver
                    # wouldn't reveal their contact number until after clicking Approve.
                    # Outside the form, the contact updates the instant a driver is chosen.
                    d1, d2 = st.columns(2)
                    with d1:
                        selected_driver = st.selectbox(
                            "Driver Name", driver_options or ["No drivers available"],
                            key=f"drv_{r['request_id']}", disabled=not driver_options,
                        )
                    with d2:
                        bound_contact = driver_contact_map.get(selected_driver, "")
                        st.text_input("Driver Contact (auto-filled)", value=bound_contact, disabled=True,
                                      key=f"dc_disp_{r['request_id']}")
                    selected_vehicle = st.selectbox(
                        "Vehicle Number", vehicle_options or ["No vehicles available"],
                        key=f"veh_{r['request_id']}", disabled=not vehicle_options,
                    )

                    with st.form(f"action_{r['request_id']}"):
                        try:
                            default_time = datetime.strptime(r["time_of_travel"], "%H:%M").time()
                        except (ValueError, TypeError):
                            default_time = datetime.now().time()
                        approved_time = st.time_input(
                            "Approved Departure Time",
                            value=default_time,
                            help=f"Originally requested for {r['time_of_travel']}. Adjust if rescheduling.",
                            key=f"atime_{r['request_id']}",
                        )
                        admin_note = st.text_area(
                            "Admin Note / Remarks (optional)",
                            placeholder="e.g., Rescheduled due to vehicle availability, or reason for rejection",
                            key=f"note_{r['request_id']}",
                        )

                        b1, b2 = st.columns(2)
                        approve_clicked = b1.form_submit_button("✅ Approve", type="primary", use_container_width=True)
                        reject_clicked = b2.form_submit_button("❌ Reject", use_container_width=True)

                    if approve_clicked or reject_clicked:
                        new_status = "Approved" if approve_clicked else "Rejected"
                        driver_ready = driver_options and selected_driver != "— Select Driver —"
                        vehicle_ready = vehicle_options and selected_vehicle != "— Select Vehicle —"
                        if approve_clicked and not (driver_ready and vehicle_ready):
                            st.error("Please select a Driver and a Vehicle before approving.")
                        else:
                            try:
                                updates = {
                                    "status": new_status,
                                    "driver_name": selected_driver if approve_clicked else "",
                                    "driver_contact": driver_contact_map.get(selected_driver, "") if approve_clicked else "",
                                    "vehicle_number": selected_vehicle if approve_clicked else "",
                                    "approved_by": user["full_name"],
                                    "action_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "admin_note": admin_note.strip(),
                                }
                                if approve_clicked:
                                    updates["approved_time"] = approved_time.strftime("%H:%M")
                                update_requisition(r["request_id"], updates)
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
            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("Total Requests", len(df_all))
            m2.metric("🟡 Pending", int((df_all["status"] == "Pending").sum()))
            m3.metric("🟢 Approved", int((df_all["status"] == "Approved").sum()))
            m4.metric("🔵 On Trip", int((df_all["status"] == "On Trip").sum()))
            m5.metric("✅ Completed", int((df_all["status"] == "Completed").sum()))
            m6.metric("🔴 Rejected", int((df_all["status"] == "Rejected").sum()))

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
                              color_discrete_map={"Approved": "#28a745", "Rejected": "#dc3545", "Pending": "#ffc107",
                                                   "On Trip": "#0d6efd", "Completed": "#17a673"})
                fig2.update_layout(height=380)
                st.plotly_chart(fig2, use_container_width=True)

            if "total_km" in df_all.columns and df_all["total_km"].notna().any():
                total_km_all = pd.to_numeric(df_all["total_km"], errors="coerce").dropna().sum()
                st.metric("🛣️ Total KM Covered (Completed Trips)", f"{total_km_all:.1f} KM")

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
                status_filter = st.multiselect("Status", REQ_STATUS_OPTIONS)
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

    # ---------------- TAB 6: Manage Drivers & Vehicles ----------------
    with tab6:
        st.subheader("🚘 Manage Drivers & Vehicles")
        st.caption("These lists power the Driver and Vehicle dropdowns admins use when approving requisitions.")

        dcol, vcol = st.columns(2)

        # ---- Drivers ----
        with dcol:
            st.markdown("##### 👨‍✈️ Drivers")
            with st.form("add_driver_form", clear_on_submit=True):
                new_driver_name = st.text_input("Driver Name *")
                new_driver_contact = st.text_input("Driver Contact *", placeholder="017XXXXXXXX")
                add_driver_clicked = st.form_submit_button("➕ Add Driver", type="primary", use_container_width=True)

            if add_driver_clicked:
                if not new_driver_name.strip() or not new_driver_contact.strip():
                    st.error("Both Driver Name and Driver Contact are required.")
                else:
                    try:
                        add_driver(new_driver_name.strip(), new_driver_contact.strip())
                        st.success(f"✅ Driver '{new_driver_name.strip()}' added.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Failed to add driver: {e}")

            st.markdown("###### Current Drivers")
            drivers_df_mgmt = fetch_all_drivers()
            if drivers_df_mgmt.empty:
                st.info("No drivers added yet.")
            else:
                for _, d in drivers_df_mgmt.iterrows():
                    r1, r2 = st.columns([4, 1])
                    r1.write(f"**{d['driver_name']}** — {d['driver_contact']}")
                    if r2.button("🗑️", key=f"del_drv_{d['id']}", help="Delete this driver"):
                        try:
                            delete_driver(d["id"])
                            st.success(f"Deleted driver '{d['driver_name']}'.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to delete driver: {e}")

        # ---- Vehicles ----
        with vcol:
            st.markdown("##### 🚐 Vehicles")
            with st.form("add_vehicle_form", clear_on_submit=True):
                new_vehicle_number = st.text_input("Vehicle Number *", placeholder="e.g. DHK-METRO-GA-1234")
                add_vehicle_clicked = st.form_submit_button("➕ Add Vehicle", type="primary", use_container_width=True)

            if add_vehicle_clicked:
                if not new_vehicle_number.strip():
                    st.error("Vehicle Number is required.")
                else:
                    try:
                        add_vehicle(new_vehicle_number.strip())
                        st.success(f"✅ Vehicle '{new_vehicle_number.strip()}' added.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Failed to add vehicle: {e}")

            st.markdown("###### Current Vehicles")
            vehicles_df_mgmt = fetch_all_vehicles()
            if vehicles_df_mgmt.empty:
                st.info("No vehicles added yet.")
            else:
                for _, v in vehicles_df_mgmt.iterrows():
                    r1, r2 = st.columns([4, 1])
                    r1.write(f"**{v['vehicle_number']}**")
                    if r2.button("🗑️", key=f"del_veh_{v['id']}", help="Delete this vehicle"):
                        try:
                            delete_vehicle(v["id"])
                            st.success(f"Deleted vehicle '{v['vehicle_number']}'.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to delete vehicle: {e}")

# =========================================================
# 11. FALLBACK — unrecognized role
# =========================================================
else:
    st.error("⚠️ Your account role is not recognized. Please contact the Admin.")
