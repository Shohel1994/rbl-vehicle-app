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
import re
import time
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
import requests
import streamlit as st


# =========================================================
# TELEGRAM NOTIFICATION SYSTEM & HELPER FUNCTIONS
# =========================================================
def send_telegram_alert(message: str):
    """Sends real-time notifications to the Telegram group with secret fallback."""
    try:
        # st.secrets থেকে রিড করার সুরক্ষিত পদ্ধতি + ব্যাকআপ মান
        bot_token = st.secrets.get("TELEGRAM_BOT_TOKEN", "8868510704:AAGkOS_s70f7ARKpvLbP3OvDDNEzDcutqIY")
        chat_id = st.secrets.get("TELEGRAM_CHAT_ID", "-1004360578852")

        if not bot_token or not chat_id:
            return

        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        payload = {
            "chat_id": str(chat_id),
            "text": message,
            "parse_mode": "Markdown"
        }
        requests.post(url, json=payload, timeout=5)
    except Exception as e:
        print(f"Telegram Notification Error: {e}")


def insert_requisition(data: dict):
    """Inserts a new requisition row and fires the 'New Requisition' Telegram
    alert. Returns the database's own auto-incrementing `id` (a small
    integer, e.g. 42) so callers can show drivers/employees a short
    "Requisition #42" instead of the long internal request_id string
    (e.g. REQ-20260831103122-982) — that long string still exists and is
    still the real unique key used for every lookup/update/join in the app,
    it's just never meant to be read or memorized by a person. Returns None
    if the new id couldn't be read back (insert still succeeds either way)."""
    sb = get_supabase_client()
    res = sb.table(REQUISITIONS_TABLE).insert(data).execute()
    new_id = res.data[0].get("id") if res.data else None

    # Invalidate cached reads so every list/dashboard reflects this insert
    # on the very next rerun instead of waiting out the cache TTL.
    _clear_requisition_caches()

    # English Telegram Alert for New Requisition
    applicant = data.get("applicant_name", "N/A")
    dept = data.get("department", "N/A")
    dest = data.get("destination", "N/A")
    display_id = new_id if new_id is not None else data.get("request_id", "N/A")
    date_of_travel = data.get("date_of_travel", "N/A")
    time_of_travel = data.get("time_of_travel", "N/A")
    vehicle_type = data.get("vehicle_type", "N/A")
    passenger_count = data.get("passenger_count", "N/A")

    msg = (
        f"🚨 **New Vehicle Requisition Submitted!**\n\n"
        f"🆔 **Requisition #:** {display_id}\n"
        f"👤 **Applicant:** {applicant}\n"
        f"🏢 **Department:** {dept}\n"
        f"📍 **Destination:** {dest}\n"
        f"📅 **Date/Time:** {date_of_travel} at {time_of_travel}\n"
        f"🚐 **Vehicle Type:** {vehicle_type}\n"
        f"👥 **Passengers:** {passenger_count}"
    )
    send_telegram_alert(msg)
    return new_id


def update_requisition(request_id: str, updates: dict, notify: bool = True):
    """Writes `updates` to the requisitions row for `request_id`.

    `notify` controls whether this call fires the Telegram "Status Updated"
    alert (default True, unchanged from before). Set notify=False for
    updates that should stay silent — currently used by the Gate Officer's
    Gate Out / Gate In panel, since KM entries there are meant to
    be quiet record-keeping (Admin can always pull a report/export), while
    the Driver's own Start/End KM entries (via submit_driver_km) and Admin's
    Approve/Reject actions keep sending the alert as before.
    """
    sb = get_supabase_client()

    applicant = ""
    dest = ""
    short_id = None
    if notify:
        # Look up applicant/destination/id so the alert is informative even
        # though `updates` itself usually only carries status-related
        # fields. Skipped entirely when notify=False since nothing here is
        # needed if we're not sending a message.
        try:
            existing = sb.table(REQUISITIONS_TABLE).select("id, applicant_name, destination").eq(
                "request_id", request_id
            ).limit(1).execute()
            if existing.data:
                applicant = existing.data[0].get("applicant_name", "")
                dest = existing.data[0].get("destination", "")
                short_id = existing.data[0].get("id")
        except Exception:
            pass  # Alert enrichment is best-effort; the update itself must still proceed.

    sb.table(REQUISITIONS_TABLE).update(updates).eq("request_id", request_id).execute()

    # Invalidate cached reads immediately after every write, regardless of
    # `notify`, so status changes (approve/reject, gate in/out, driver KM
    # entry, admin edits) show up on the very next rerun instead of waiting
    # out the cache TTL.
    _clear_requisition_caches()

    if not notify:
        return

    # English Telegram Alert for Status Update
    status = updates.get("status", "Updated")
    driver = updates.get("driver_name", "")
    vehicle = updates.get("vehicle_number", "")
    display_id = short_id if short_id is not None else request_id

    msg = (
        f"📢 **Requisition Status Updated!**\n\n"
        f"🆔 **Requisition #:** {display_id}\n"
    )
    if applicant:
        msg += f"👤 **Applicant:** {applicant}\n"
    if dest:
        msg += f"📍 **Destination:** {dest}\n"
    msg += f"📌 **New Status:** {status}"
    if driver:
        msg += f"\n👨‍✈️ **Driver:** {driver}"
    if vehicle:
        msg += f"\n🚗 **Vehicle:** {vehicle}"

    send_telegram_alert(msg)


# =========================================================
# 1. COMPANY INFO & PAGE CONFIG
# =========================================================
COMPANY_NAME = "Renaissaince Barind Ltd."
COMPANY_ADDRESS = "Ishwardi EPZ, Pakshi, Pabna"

st.set_page_config(
    page_title="RBL VMS",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

USERS_TABLE = "users"
REQUISITIONS_TABLE = "requisitions"
DRIVERS_TABLE = "drivers"
VEHICLES_TABLE = "vehicles"
SESSIONS_TABLE = "sessions"
SHUTTLE_TEMPLATES_TABLE = "shuttle_templates"

SESSION_COOKIE_NAME = "rbl_vms_session"
SESSION_LIFETIME_DAYS = 30

VEHICLE_TYPES = ["HIACE", "Private Car", "Pick-up Van", "Covered Van", "Truck", "Shipment Vehicle", "Ambulance", "Other"]
DEPARTMENTS = [
    "Accounts", "Warehouse", "Factory Merchandising", "Commercial", "Floor Operation",
    "TSD", "QMS", "Production Planning & Control", "Cutting", "R & D", "Admin",
    "Technical", "Finishing", "Quality", "IE", "HR & Compliance", "Medical", "Other",
]

# Account status (users table) — approval workflow for logins
USER_STATUS_OPTIONS = ["Pending", "Approved", "Rejected"]
ROLE_OPTIONS = ["user", "gate_officer", "driver", "admin", "executive", "nurse"]
ROLE_DISPLAY = {
    "user": "Employee",
    "gate_officer": "Gate Officer",
    "driver": "Driver",
    "admin": "Admin",
    "executive": "Executive / Management",
    "nurse": "Nurse / Senior Nurse",
}

# Requisition status (requisitions table) — full trip lifecycle
REQ_STATUS_OPTIONS = ["Pending", "Approved", "Rejected", "On Trip", "Completed"]
STATUS_BADGE = {
    "Pending": "🟡 Pending", "Approved": "🟢 Approved", "Rejected": "🔴 Rejected",
    "On Trip": "🔵 On Trip", "Completed": "✅ Completed",
}


# =========================================================
# 2. HELPER FUNCTIONS & HEADER (No Logo Version)
# =========================================================
@st.cache_data(show_spinner=False)
def get_logo_base64():
    """Logo system completely disabled."""
    return None


def company_header(subtitle: str = ""):
    """Reusable company name + address banner without logo."""
    st.title(f"🚗 {COMPANY_NAME}")
    st.caption(f"📍 {COMPANY_ADDRESS}")
    if subtitle:
        st.subheader(subtitle)
    st.divider()


def badge_class(status: str) -> str:
    return {
        "Pending": "badge-pending", "Approved": "badge-approved", "Rejected": "badge-rejected",
        "On Trip": "badge-ontrip", "Completed": "badge-completed",
    }.get(status, "badge-pending")


def is_blank(val) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str) and not val.strip():
        return True
    return False


def fmt(val, default: str = "—") -> str:
    return default if is_blank(val) else str(val)


def short_req_id(value) -> str:
    """Short, human-friendly requisition label ('#42') built from the
    database's own auto-incrementing `id` column. The long internal
    request_id (e.g. 'REQ-20260831103122-982') still exists and is still
    the real unique key used for every lookup/update/join in the app — it's
    just never meant to be read or memorized by a driver or employee, only
    this short number is. Falls back to showing the raw value if it isn't
    a usable number (e.g. already blank, or an unexpected type)."""
    if is_blank(value):
        return "—"
    try:
        return f"#{int(float(value))}"
    except (TypeError, ValueError):
        return str(value)


def fmt_time_12h(value, default: str = "—") -> str:
    """Render a stored time/timestamp string in 12-hour AM/PM format
    (Bangladesh local convention). Storage format in Supabase is UNCHANGED —
    still 24-hour '%H:%M' or '%Y-%m-%d %H:%M:%S' — this only changes what's
    *displayed* on screen, in exports, and in the Duty Tracker table.

    Accepts bare 'HH:MM' / 'HH:MM:SS' (time_of_travel, approved_time) or full
    'YYYY-MM-DD HH:MM:SS' (actual_exit_time, actual_return_time,
    action_timestamp). timestamptz columns come back from Supabase/PostgREST
    as ISO 8601 with a trailing UTC/offset marker and often fractional
    seconds (e.g. "2026-08-30T09:25:19.123456+00:00") — that suffix is
    stripped here before parsing so the driver/employee-facing display shows
    a clean 12-hour time instead of the raw timestamp. This is purely
    cosmetic reformatting of the same wall-clock numbers that were written
    (no timezone conversion/math is applied). If the value doesn't match a
    known shape (e.g. it's genuinely something else), it's returned
    unchanged rather than raising — exports must never crash on this.
    """
    if is_blank(value):
        return default
    s = str(value).strip().replace("T", " ")
    # Strip a trailing timezone offset ("+00:00", "+06:00", "-05:00", "Z")
    # and any fractional seconds, e.g. "2026-08-30 09:25:19.123456+00:00"
    # -> "2026-08-30 09:25:19". Keeps the exact numbers as stored/typed.
    s_clean = re.sub(r"(\.\d+)?(Z|[+-]\d{2}:?\d{2})$", "", s, flags=re.IGNORECASE).strip()
    for candidate in (s_clean, s):
        for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%H:%M:%S", "%H:%M"):
            try:
                dt = datetime.strptime(candidate, f)
            except ValueError:
                continue
            return dt.strftime("%Y-%m-%d %I:%M %p") if "%Y" in f else dt.strftime("%I:%M %p")
    return s


def hash_password(raw: str) -> str:
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def time_input_12h(label: str, key_prefix: str, default_time=None):
    """A 12-hour AM/PM time-entry widget: one 'H:MM' text field plus an
    AM/PM dropdown (2 inputs total, down from 3 separate Hour/Minute/AM-PM
    selectboxes) — fewer things to click through to set an arbitrary time,
    e.g. type "4:45" and pick "PM" instead of opening three dropdowns.
    Returns a plain datetime.time object — exactly what st.time_input
    returns — so every existing call site that does `.strftime('%H:%M')`
    on the result keeps working unchanged. Invalid typed input falls back
    to `default_time` with an inline warning rather than crashing the form."""
    if default_time is None:
        default_time = datetime.now().time()
    default_12h = default_time.strftime("%I:%M %p")  # e.g. "09:05 AM"
    d_hm, d_ampm = default_12h.rsplit(" ", 1)  # -> "09:05", "AM"

    st.markdown(f"**{label}**")
    c1, c2 = st.columns([2, 1])
    with c1:
        hm_text = st.text_input(
            "Hour:Minute", value=d_hm, key=f"{key_prefix}_hm",
            placeholder="e.g. 4:45",
            help="Type the time as H:MM or HH:MM on a 12-hour clock, then pick AM or PM.",
        )
    with c2:
        ap = st.selectbox("AM/PM", ["AM", "PM"], index=0 if d_ampm == "AM" else 1,
                           key=f"{key_prefix}_ampm")

    # Accept "4:45", "04:45", or common mistyped separators like "4.45" /
    # "4 45" by normalizing to "H:MM" before parsing.
    raw = hm_text.strip()
    normalized = re.sub(r"^(\d{1,2})[.\s](\d{2})$", r"\1:\2", raw)
    try:
        parsed = datetime.strptime(normalized, "%I:%M")
        return datetime.strptime(f"{parsed.strftime('%I:%M')} {ap}", "%I:%M %p").time()
    except ValueError:
        st.error(
            f"⚠️ Couldn't understand '{hm_text}' as a time (expected H:MM, e.g. 4:45) — "
            f"using {d_hm} {d_ampm} for now."
        )
        return default_time


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


# ---------------------- CACHE INVALIDATION HELPERS ----------------------
# All read helpers below are wrapped in @st.cache_data so the app doesn't
# hit Supabase on every single rerun (auto-refresh tick, button click,
# widget change, etc.) — this is what makes the app feel fast/"patla".
# Every write helper (insert/update/delete) calls the matching `_clear_*`
# function immediately after its Supabase call succeeds, so nobody ever
# sees stale data — they just don't pay a network round-trip for reads
# that haven't changed since the last fetch.
def _clear_user_caches():
    fetch_all_users.clear()


def _clear_requisition_caches():
    fetch_all_requisitions.clear()
    fetch_requisitions_by_user.clear()
    fetch_requisitions_by_status.clear()
    fetch_requisitions_by_driver.clear()
    get_last_driver_end_km.clear()


def _clear_driver_caches():
    fetch_all_drivers.clear()


def _clear_vehicle_caches():
    fetch_all_vehicles.clear()


def _clear_shuttle_template_caches():
    fetch_all_shuttle_templates.clear()


# ---------------------- USERS TABLE HELPERS ----------------------
def get_user_by_username(username: str):
    sb = get_supabase_client()
    res = sb.table(USERS_TABLE).select("*").eq("username", username).limit(1).execute()
    return res.data[0] if res.data else None


def register_user(data: dict):
    sb = get_supabase_client()
    sb.table(USERS_TABLE).insert(data).execute()
    _clear_user_caches()


def update_user(username: str, updates: dict):
    sb = get_supabase_client()
    sb.table(USERS_TABLE).update(updates).eq("username", username).execute()
    _clear_user_caches()


def delete_user(username: str):
    """Permanently remove a user account (e.g. after they leave the company)."""
    sb = get_supabase_client()
    sb.table(USERS_TABLE).delete().eq("username", username).execute()
    _clear_user_caches()


@st.cache_data(ttl=90, show_spinner=False)
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


# NOTE: insert_requisition() and update_requisition() are defined once, above,
# in the "TELEGRAM NOTIFICATION SYSTEM & HELPER FUNCTIONS" section — they
# perform the Supabase write AND fire the Telegram alert. They are
# intentionally not redefined here; a second, alert-less definition at this
# point in the file would silently shadow (override) the ones above and the
# Telegram notifications would never fire, since Python keeps whichever `def`
# runs last for a given name.


def delete_requisition(request_id: str):
    """Permanently removes a requisition row. Used only by the Admin's
    Edit/Delete Trip tab for correcting outright mistakes — e.g. a
    duplicate entry, a test submission, or a request that should never
    have existed. This is NOT part of the normal trip lifecycle
    (Pending -> Approved -> On Trip -> Completed), which never deletes
    rows; use status changes for that instead."""
    sb = get_supabase_client()
    sb.table(REQUISITIONS_TABLE).delete().eq("request_id", request_id).execute()
    _clear_requisition_caches()


@st.cache_data(ttl=90, show_spinner=False)
def fetch_all_requisitions() -> pd.DataFrame:
    sb = get_supabase_client()
    res = sb.table(REQUISITIONS_TABLE).select("*").order("created_at", desc=True).execute()
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=[
        "id", "request_id", "created_at", "username", "applicant_name", "department", "mobile_number",
        "date_of_travel", "time_of_travel", "destination", "passenger_count", "vehicle_type", "purpose",
        "special_request", "status", "driver_name", "driver_contact", "vehicle_number", "approved_by",
        "action_timestamp", "approved_time", "admin_note", "start_km", "end_km", "total_km",
        "actual_exit_time", "actual_return_time", "driver_start_km", "driver_end_km", "driver_km_updated_at",
    ])


@st.cache_data(ttl=90, show_spinner=False)
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
        "actual_exit_time", "actual_return_time", "driver_start_km", "driver_end_km", "driver_km_updated_at",
    ])


@st.cache_data(ttl=90, show_spinner=False)
def fetch_requisitions_by_status(status: str) -> pd.DataFrame:
    """Server-side filtered fetch used by the Gate Officer panel — only pulls
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
        "actual_exit_time", "actual_return_time", "driver_start_km", "driver_end_km", "driver_km_updated_at",
    ])


# ------------------- DRIVER-SUBMITTED KM HELPERS (NEW) -------------------
# Additive helpers backing the Driver Dashboard (role branch, Section 9B) and
# the Admin's KM Variance Report (Tab 8). These never touch the Gate
# Officer's start_km / end_km / total_km fields — they read/write the
# separate driver_start_km / driver_end_km / driver_km_updated_at columns only.
def _normalize_driver_name(name) -> str:
    """Case-insensitive, whitespace-collapsed key for matching a Driver's own
    registered full_name against the Admin-assigned requisitions.driver_name.
    Both are free-typed by different people (the driver at registration, the
    admin when adding them under Manage Drivers & Vehicles / approving a
    request) — without this normalization, something as small as an extra
    space, a missing space (e.g. "Mr. Yusuf" vs "Mr.Yusuf"), or a different
    capitalization makes Supabase's exact .eq() match silently return
    nothing, and the driver's approved trips simply never appear on their
    own dashboard. ALL whitespace is stripped (not just leading/trailing),
    so any spacing variation around punctuation like "Mr." still matches."""
    if is_blank(name):
        return ""
    return "".join(str(name).casefold().split())


@st.cache_data(ttl=90, show_spinner=False)
def fetch_requisitions_by_driver(driver_name: str) -> pd.DataFrame:
    """All trips assigned to this driver, matched case-insensitively and
    ignoring leading/trailing whitespace (see _normalize_driver_name) —
    so a mismatch like "Yusuf" vs "yusuf " no longer hides trips from the
    Driver Dashboard. Fetches every requisition with a driver assigned and
    filters client-side rather than relying on Supabase's exact-match .eq(),
    since PostgREST has no case-insensitive-and-trimmed equality operator."""
    sb = get_supabase_client()
    res = (
        sb.table(REQUISITIONS_TABLE)
        .select("*")
        .not_.is_("driver_name", "null")
        .order("created_at", desc=True)
        .execute()
    )
    cols = [
        "id", "request_id", "created_at", "username", "applicant_name", "department", "mobile_number",
        "date_of_travel", "time_of_travel", "destination", "passenger_count", "vehicle_type", "purpose",
        "special_request", "status", "driver_name", "driver_contact", "vehicle_number", "approved_by",
        "action_timestamp", "approved_time", "admin_note", "start_km", "end_km", "total_km",
        "actual_exit_time", "actual_return_time", "driver_start_km", "driver_end_km", "driver_km_updated_at",
    ]
    if not res.data:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(res.data)
    target = _normalize_driver_name(driver_name)
    matched = df[df["driver_name"].map(_normalize_driver_name) == target]
    return matched.reset_index(drop=True)


@st.cache_data(ttl=90, show_spinner=False)
def get_last_driver_end_km(driver_name: str, vehicle_number: str) -> float:
    """The driver's own most recent End KM for this specific vehicle — used
    to auto-fill their next Start KM (still fully editable). Returns 0.0 if
    no prior driver-submitted End KM exists yet for this driver+vehicle.
    Driver name is matched case-insensitively and ignoring leading/trailing
    whitespace (same normalization as fetch_requisitions_by_driver) so a
    stray capitalization/space difference doesn't silently break the
    auto-fill either. Vehicle number is still matched exactly, since it's
    always chosen from the same fixed Manage Vehicles dropdown everywhere,
    so it can't drift the way a free-typed name can."""
    if is_blank(vehicle_number):
        return 0.0
    sb = get_supabase_client()
    res = (
        sb.table(REQUISITIONS_TABLE)
        .select("driver_name, driver_end_km, created_at")
        .eq("vehicle_number", vehicle_number)
        .not_.is_("driver_end_km", "null")
        .order("created_at", desc=True)
        .execute()
    )
    if not res.data:
        return 0.0
    target = _normalize_driver_name(driver_name)
    for row in res.data:
        if _normalize_driver_name(row.get("driver_name")) == target and row.get("driver_end_km") is not None:
            return float(row["driver_end_km"])
    return 0.0


def submit_driver_km(row: dict, driver_start_km=None, driver_end_km=None):
    """Saves the driver's own odometer entry AND drives the trip's lifecycle
    status — per business requirement, the Driver's KM entries are now what
    move a trip forward, not just the Gate Officer's Gate panel:
      - Start KM only  -> status Approved -> On Trip. actual_exit_time is
        stamped now (only if the Gate Officer hasn't already logged one).
      - End KM present -> status -> Completed. actual_return_time is stamped
        now (only if the Gate Officer hasn't already logged one), and
        total_km is filled from the driver's own distance if the Gate
        Officer hasn't recorded one.
    This reuses update_requisition() (rather than a raw Supabase call) so the
    Telegram alert keeps the EXACT same "📢 Requisition Status Updated!"
    format that the Gate Officer's Gate Out/Gate In has always used — only
    the trigger has moved, from the Gate Officer's entry to the Driver's own
    entry.
    `row` is the full existing requisition dict (from
    fetch_requisitions_by_driver), used only so we never clobber a timestamp
    or total the Gate Officer has already logged — the Gate Officer's own
    start_km/end_km columns are never written here, and the Gate Officer can
    still Gate Out/Gate In independently at any time for cross-verification
    (see the KM Variance Report tab)."""
    request_id = row["request_id"]
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    updates = {
        "driver_km_updated_at": now_str,
        # Re-affirm these (unchanged) values so update_requisition()'s
        # Telegram alert includes the Driver/Vehicle lines exactly like it
        # always has for Gate-Officer-triggered updates.
        "driver_name": row.get("driver_name", ""),
        "vehicle_number": row.get("vehicle_number", ""),
    }
    # IMPORTANT: `row` typically comes from a pandas DataFrame row (via
    # .to_dict()), where a missing numeric cell is Python float('nan'),
    # NOT None. `driver_start_km is not None` is True for NaN, so the old
    # check let `float(nan)` slip into `updates` — and float('nan') is NOT
    # valid JSON, so the Supabase client's request body serialization blew
    # up with "Out of range float values are not JSON compliant: nan" the
    # moment a driver tried to complete a trip that never had its own
    # Start KM recorded (e.g. one that Admin/Gate Officer had already
    # marked Completed directly, bypassing the Driver's Start KM step).
    # is_blank() correctly treats NaN, None, and "" all as "missing", so
    # this now safely skips writing a value instead of crashing.
    has_start = not is_blank(driver_start_km)
    has_end = not is_blank(driver_end_km)

    if has_start:
        updates["driver_start_km"] = float(driver_start_km)
    if has_end:
        updates["driver_end_km"] = float(driver_end_km)

    if has_end:
        updates["status"] = "Completed"
        if is_blank(row.get("actual_return_time")):
            updates["actual_return_time"] = now_str
        if is_blank(row.get("total_km")) and has_start:
            updates["total_km"] = round(float(driver_end_km) - float(driver_start_km), 1)
    elif has_start and row.get("status") == "Approved":
        updates["status"] = "On Trip"
        if is_blank(row.get("actual_exit_time")):
            updates["actual_exit_time"] = now_str

    update_requisition(request_id, updates)


def effective_km_fields(row) -> tuple:
    """(start_km, end_km, total_km) using the Driver's own odometer entries
    as the PRIMARY source of truth for standard distance reports and duty
    tracking, falling back to the Gate Officer's Gate-Out/Gate-In readings
    only when the driver hasn't logged their own numbers for that trip yet.
    The Gate Officer's start_km/end_km columns are never modified by this —
    read-only helper."""
    d_start, d_end = row.get("driver_start_km"), row.get("driver_end_km")
    if not is_blank(d_start) and not is_blank(d_end):
        s, e = float(d_start), float(d_end)
        return s, e, round(e - s, 1)
    s_start, s_end = row.get("start_km"), row.get("end_km")
    if not is_blank(s_start) and not is_blank(s_end):
        s, e = float(s_start), float(s_end)
        return s, e, round(e - s, 1)
    return None, None, 0.0


def compute_duty_hours(row) -> float:
    """How long a vehicle was actually out for a given trip, in hours.

    Uses the same Gate-Officer-logged Gate-Out/Gate-In timestamps
    (actual_exit_time / actual_return_time) as the Duty Tracker tab, so the
    numbers here always agree with that tab. A trip that hasn't Gated Out
    yet contributes 0 hours; a trip that's still 'On Trip' (Gated Out but
    not yet Gated In) counts up to right now, so an in-progress duty still
    shows up in live totals instead of being ignored until it's Completed.
    Read-only helper — never writes anything back to Supabase.
    """
    start_raw = row.get("actual_exit_time")
    if is_blank(start_raw):
        return 0.0
    start_dt = pd.to_datetime(start_raw, errors="coerce", utc=True, format="mixed")
    if pd.isna(start_dt):
        return 0.0

    end_raw = row.get("actual_return_time")
    if is_blank(end_raw):
        end_dt = pd.Timestamp.now(tz="UTC")
    else:
        end_dt = pd.to_datetime(end_raw, errors="coerce", utc=True, format="mixed")
        if pd.isna(end_dt):
            end_dt = pd.Timestamp.now(tz="UTC")

    hours = (end_dt - start_dt).total_seconds() / 3600.0
    return round(max(hours, 0.0), 2)


# ------------------- DRIVERS & VEHICLES TABLE HELPERS -------------------
# These back the dynamic dropdowns in the requisition-approval form so admins
# maintain one source of truth instead of retyping names/numbers each time.
@st.cache_data(ttl=90, show_spinner=False)
def fetch_all_drivers() -> pd.DataFrame:
    sb = get_supabase_client()
    res = sb.table(DRIVERS_TABLE).select("*").order("driver_name").execute()
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=["id", "driver_name", "driver_contact", "created_at"])


def add_driver(driver_name: str, driver_contact: str):
    sb = get_supabase_client()
    sb.table(DRIVERS_TABLE).insert({"driver_name": driver_name, "driver_contact": driver_contact}).execute()
    _clear_driver_caches()


def delete_driver(driver_id):
    sb = get_supabase_client()
    sb.table(DRIVERS_TABLE).delete().eq("id", driver_id).execute()
    _clear_driver_caches()


@st.cache_data(ttl=90, show_spinner=False)
def fetch_all_vehicles() -> pd.DataFrame:
    sb = get_supabase_client()
    res = sb.table(VEHICLES_TABLE).select("*").order("vehicle_number").execute()
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=["id", "vehicle_number", "created_at"])


def add_vehicle(vehicle_number: str):
    sb = get_supabase_client()
    sb.table(VEHICLES_TABLE).insert({"vehicle_number": vehicle_number}).execute()
    _clear_vehicle_caches()


def delete_vehicle(vehicle_id):
    sb = get_supabase_client()
    sb.table(VEHICLES_TABLE).delete().eq("id", vehicle_id).execute()
    _clear_vehicle_caches()


# ------------------- SHUTTLE TEMPLATE HELPERS (NEW) -------------------
# Backs the "🚌 Staff Shuttle" Admin tab: for recurring fixed-schedule
# routes (e.g. the 2 daily HIACE staff shuttle runs), Admin saves the
# route/purpose/vehicle-type/default driver+vehicle ONCE as a named
# template, then quick-submits today's trip in a couple of clicks instead
# of retyping everything every day. Independent of drivers/vehicles/
# requisitions tables — never modifies them.
@st.cache_data(ttl=90, show_spinner=False)
def fetch_all_shuttle_templates() -> pd.DataFrame:
    sb = get_supabase_client()
    res = sb.table(SHUTTLE_TEMPLATES_TABLE).select("*").order("template_name").execute()
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=[
        "id", "template_name", "applicant_name", "department", "destination", "purpose",
        "vehicle_type", "passenger_count", "default_time", "default_driver_name",
        "default_vehicle_number", "created_at",
    ])


def add_shuttle_template(data: dict):
    sb = get_supabase_client()
    sb.table(SHUTTLE_TEMPLATES_TABLE).insert(data).execute()
    _clear_shuttle_template_caches()


def delete_shuttle_template(template_id):
    sb = get_supabase_client()
    sb.table(SHUTTLE_TEMPLATES_TABLE).delete().eq("id", template_id).execute()
    _clear_shuttle_template_caches()


# ------------------- DAILY AUTO-GENERATED SHUTTLE REQUISITIONS (NEW) -------------------
# Fixed 6-trip daily shuttle schedule. These are created automatically
# (no button click needed) the first time ANY logged-in user loads the app
# each day, as Pending requisitions with driver/vehicle left blank — Admin
# just opens "Pending Requests" and picks Driver + Vehicle like any normal
# approval (including via the existing Bulk Assign feature). Tagged via
# admin_note so re-checking never depends on exact row order/count matching
# (safe even if Admin edits/deletes one of today's six later).
DAILY_SHUTTLE_TAG = "AUTO-DAILY-SHUTTLE"

DAILY_SHUTTLE_TEMPLATES = [
    {"destination": "Ishwardi", "time": "19:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
    {"destination": "Ishwardi", "time": "19:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
    {"destination": "Ishwardi", "time": "19:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
    {"destination": "Dashuria", "time": "19:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
    {"destination": "Ishwardi", "time": "20:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
    {"destination": "Bepza", "time": "20:15", "passenger_count": 1, "purpose": "Commercial Duty", "vehicle_type": "HIACE"},
]


def ensure_daily_shuttle_requisitions():
    """Auto-creates today's fixed shuttle requisitions (Pending, no driver/
    vehicle assigned) if they don't already exist. Runs at most once per
    browser session per day via session_state (not on every rerun/
    auto-refresh), and is itself idempotent via the admin_note tag check —
    so even if two people happen to trigger it around the same time, it
    won't double-create today's six. Any failure here is logged and
    swallowed, never shown to the user or allowed to block the rest of the
    app from loading."""
    today_str = str(date.today())
    if st.session_state.get("_daily_shuttle_checked_date") == today_str:
        return
    st.session_state["_daily_shuttle_checked_date"] = today_str

    try:
        sb = get_supabase_client()
        existing = (
            sb.table(REQUISITIONS_TABLE)
            .select("id")
            .eq("date_of_travel", today_str)
            .like("admin_note", f"%{DAILY_SHUTTLE_TAG}%")
            .execute()
        )
        already_count = len(existing.data) if existing.data else 0
        if already_count >= len(DAILY_SHUTTLE_TEMPLATES):
            return  # today's 6 are already in place

        for tpl in DAILY_SHUTTLE_TEMPLATES[already_count:]:
            data = {
                "request_id": generate_request_id(),
                "username": "",
                "applicant_name": "Staff Shuttle",
                "department": "Admin",
                "mobile_number": "",
                "date_of_travel": today_str,
                "time_of_travel": tpl["time"],
                "destination": tpl["destination"],
                "passenger_count": tpl["passenger_count"],
                "vehicle_type": tpl["vehicle_type"],
                "purpose": tpl["purpose"],
                "special_request": "",
                "status": "Pending",
                "driver_name": "",
                "driver_contact": "",
                "vehicle_number": "",
                "approved_by": "",
                "admin_note": DAILY_SHUTTLE_TAG,
            }
            insert_requisition(data)
    except Exception as e:
        print(f"Daily shuttle auto-creation error: {e}")


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
                            # IMPORTANT: cookie_manager.set() only *dispatches* a message
                            # to the browser's cookie-manager iframe component asking it
                            # to write document.cookie — it does not write synchronously.
                            # If we st.rerun() immediately, Streamlit tears the page down
                            # before the browser JS has a chance to actually persist the
                            # cookie, so "Remember me" silently never works. A short pause
                            # here gives the component time to finish the write before the
                            # rerun happens.
                            with st.spinner("Setting up your session..."):
                                time.sleep(1.0)
                        except Exception:
                            session_token = None  # Remember-me is best-effort; login still succeeds without it.
                    result["session_token"] = session_token
                    st.session_state.auth_user = result
                    st.rerun()

        with tab_register:
            st.caption("New employees, Gate Officers, Drivers, or Medical staff can request an "
                       "account here. An admin must approve your account before you can log in.")

            reg_role = st.radio(
                "Register as", ["Employee", "Gate Officer", "Driver", "Nurse / Senior Nurse"],
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

            elif reg_role == "Gate Officer":
                st.caption("Choose your own username and password below, just like a regular account.")
                with st.form("register_gate_form", clear_on_submit=True):
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
                            "designation": "Gate Officer",
                            "employee_id": "",
                            "department": "Gate",
                            "mobile": "",
                            "role": "gate_officer",
                            "status": "Pending",
                        })
                        st.success(
                            "✅ Your Gate Officer account request has been submitted! "
                            "Please wait for admin approval before signing in."
                        )

            elif reg_role == "Driver":
                st.caption(
                    "Choose your own username and password below. **Important:** enter your Full "
                    "Name exactly as it appears (or will appear) in the Admin's **Manage Drivers & "
                    "Vehicles** list — your assigned trips are matched by this name, so a mismatch "
                    "means your trips won't show up in your Driver Dashboard."
                )
                with st.form("register_driver_form", clear_on_submit=True):
                    drv_full_name = st.text_input("Full Name *", key="drv_full_name")
                    drv_username = st.text_input("Username *", key="drv_username")
                    drv_password = st.text_input("Password *", type="password", key="drv_password")
                    drv_password_confirm = st.text_input("Confirm Password *", type="password", key="drv_password_confirm")
                    drv_submitted = st.form_submit_button("Submit Request", type="primary", use_container_width=True)

                if drv_submitted:
                    errors = []
                    if not drv_full_name.strip():
                        errors.append("Full Name is required.")
                    if not drv_username.strip():
                        errors.append("Username is required.")
                    if not drv_password:
                        errors.append("Password is required.")
                    elif len(drv_password) < 4:
                        errors.append("Password must be at least 4 characters.")
                    elif drv_password != drv_password_confirm:
                        errors.append("Passwords do not match.")
                    if not errors and get_user_by_username(drv_username.strip()):
                        errors.append("This username is already taken. Please choose another.")

                    if errors:
                        for e in errors:
                            st.error(e)
                    else:
                        register_user({
                            "username": drv_username.strip(),
                            "password": hash_password(drv_password),
                            "full_name": drv_full_name.strip(),
                            "designation": "Driver",
                            "employee_id": "",
                            "department": "Transport",
                            "mobile": "",
                            "role": "driver",
                            "status": "Pending",
                        })
                        st.success(
                            "✅ Your Driver account request has been submitted! "
                            "Please wait for admin approval before signing in."
                        )

            else:  # reg_role == "Nurse / Senior Nurse"
                st.caption(
                    "Choose your own username and password below. Once approved, you'll get a "
                    "simple one-click Emergency dashboard for arranging a vehicle to carry a patient."
                )
                with st.form("register_nurse_form", clear_on_submit=True):
                    nur_full_name = st.text_input("Full Name *", key="nur_full_name")
                    nur_designation = st.radio("Designation *", ["Nurse", "Senior Nurse"],
                                                horizontal=True, key="nur_designation")
                    nur_username = st.text_input("Username *", key="nur_username")
                    nur_mobile = st.text_input("Mobile Number *", key="nur_mobile", placeholder="01XXXXXXXXX")
                    nur_password = st.text_input("Password *", type="password", key="nur_password")
                    nur_password_confirm = st.text_input("Confirm Password *", type="password", key="nur_password_confirm")
                    nur_submitted = st.form_submit_button("Submit Request", type="primary", use_container_width=True)

                if nur_submitted:
                    errors = []
                    if not nur_full_name.strip():
                        errors.append("Full Name is required.")
                    if not nur_username.strip():
                        errors.append("Username is required.")
                    if not nur_mobile.strip():
                        errors.append("Mobile Number is required.")
                    if not nur_password:
                        errors.append("Password is required.")
                    elif len(nur_password) < 4:
                        errors.append("Password must be at least 4 characters.")
                    elif nur_password != nur_password_confirm:
                        errors.append("Passwords do not match.")
                    if not errors and get_user_by_username(nur_username.strip()):
                        errors.append("This username is already taken. Please choose another.")

                    if errors:
                        for e in errors:
                            st.error(e)
                    else:
                        register_user({
                            "username": nur_username.strip(),
                            "password": hash_password(nur_password),
                            "full_name": nur_full_name.strip(),
                            "designation": nur_designation,
                            "employee_id": "",
                            "department": "Medical",
                            "mobile": nur_mobile.strip(),
                            "role": "nurse",
                            "status": "Pending",
                        })
                        st.success(
                            f"✅ Your {nur_designation} account request has been submitted! "
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
    pdf.cell(0, 6, f"Generated on: {datetime.now().strftime('%Y-%m-%d %I:%M %p')}",
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

    # Every cell wraps automatically via fpdf2's table() API — this is what
    # actually prevents text from bleeding into the next column, instead of
    # truncating with "…" as an earlier version did. "Req #" uses the
    # database's own short auto-incrementing id (e.g. #42) rather than the
    # long internal request_id string, which exists only for unique
    # lookups and isn't meant to be read by a person.
    headers = ["Req #", "Applicant", "Department", "Date", "Time", "Destination",
               "Vehicle", "Status", "Driver", "Vehicle No.", "Total KM"]
    cols = ["id", "applicant_name", "department", "date_of_travel", "time_of_travel",
            "destination", "vehicle_type", "status", "driver_name", "vehicle_number", "total_km"]
    col_widths = [16, 24, 22, 18, 13, 30, 16, 16, 22, 28, 16]

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
                val = r.get(col, "")
                if col == "time_of_travel":
                    val = fmt_time_12h(val, "")
                elif col == "id":
                    val = short_req_id(val)
                row.cell(fmt(val, ""))

    return bytes(pdf.output())


# =========================================================
# 5B. DUTY TRACKER — EXCEL + PDF REPORT GENERATORS
# =========================================================
# These are ADDITIVE helpers for the new "Duty Tracker & Analytics" tab
# (Section 10B below). They are intentionally separate from
# build_excel_report() / build_pdf_report() / ReportPDF above so the
# existing "All Requisitions & Export" tab (Tab 5) keeps behaving exactly
# as before — nothing here is called from, or changes, that code path.

DUTY_TRACKER_DISPLAY_COLS = [
    "Vehicle No", "Driver Name", "Start Time", "End Time",
    "Total KM", "Duty Duration (Hrs)", "Route / Purpose",
]


def sanitize_pdf_text(value) -> str:
    """Make any string safe to hand to FPDF's core 'Helvetica' font.

    Core (non-embedded) PDF fonts like Helvetica only support the Latin-1
    character set. Common "smart" punctuation that Python/pandas/Streamlit
    happily display — em dashes (—), en dashes (–), curly quotes ('' ""),
    ellipses (…), bullets (•) — falls outside that set and makes fpdf2 raise
    FPDFUnicodeEncodingException the moment it's written to a cell. This
    function swaps the common offenders for plain-ASCII equivalents, then
    uses a Latin-1 encode/decode round-trip as a final safety net so any
    other unsupported character (e.g. stray emoji, non-Latin scripts such as
    Bangla typed into a destination/purpose field) degrades to '?' instead of
    crashing the whole export.

    Note: this keeps the export crash-proof but Latin-1-only. If admin notes,
    destinations, or driver names need full Bangla/Unicode rendering in the
    PDF itself, that requires embedding a Unicode TTF font (e.g. via
    pdf.add_font(...)) instead of the core Helvetica font — a larger change,
    out of scope for this fix.
    """
    if value is None:
        return ""
    text = str(value)
    replacements = {
        "\u2014": "-",   # — em dash
        "\u2013": "-",   # – en dash
        "\u2015": "-",   # ― horizontal bar
        "\u2018": "'", "\u2019": "'",   # ‘ ’ curly single quotes
        "\u201c": '"', "\u201d": '"',   # “ ” curly double quotes
        "\u2026": "...",  # … ellipsis
        "\u2022": "-",   # • bullet
        "\u2192": "->",  # → arrow
        "\u00a0": " ",   # non-breaking space
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    # Final safety net: anything still outside Latin-1 becomes '?' rather
    # than raising, so the PDF always generates successfully.
    return text.encode("latin-1", "replace").decode("latin-1")


def build_duty_tracker_excel(detail_df: pd.DataFrame, summary_metrics: dict) -> bytes:
    """Formatted .xlsx export for the Duty Tracker: a 'Summary' sheet with the
    KPI cards' values, plus a 'Duty Log' sheet with the full filtered detail
    table. Plug in your own detail_df / summary_metrics from the tab below —
    both are plain pandas / dict objects, nothing Supabase-specific here.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary_df = pd.DataFrame(
            [{"Metric": k, "Value": v} for k, v in summary_metrics.items()]
        )
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        detail_df.to_excel(writer, index=False, sheet_name="Duty Log")

        # Light auto-fit so columns aren't clipped in Excel — purely cosmetic,
        # safe to remove if you don't want the extra openpyxl dependency calls.
        # `fillna("")` BEFORE `.astype(str)` matters on pandas >= 3.0: that
        # version stopped converting NaN/None to the literal string "nan" on
        # astype(str), leaving real NaN behind instead. A column that's
        # partially or entirely blank (e.g. admin_note with no note yet)
        # would then make .str.len().max() return NaN, and int(NaN) raises
        # "ValueError: cannot convert float NaN to integer" — filling blanks
        # with "" first guarantees every length is a real integer (0 for
        # blank cells).
        from openpyxl.utils import get_column_letter
        for sheet_name, sheet_df in (("Summary", summary_df), ("Duty Log", detail_df)):
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(sheet_df.columns, start=1):
                width = max(12, min(40, int(sheet_df[col].fillna("").astype(str).str.len().max() if not sheet_df.empty else 12) + 2))
                ws.column_dimensions[get_column_letter(i)].width = width

    return buf.getvalue()


class DutyTrackerPDF(FPDF):
    """Separate FPDF subclass (rather than reusing ReportPDF) so this report's
    title/branding can evolve independently of the existing requisition report."""

    def header(self):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(15, 98, 254)
        self.cell(0, 9, COMPANY_NAME, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.set_font("Helvetica", "", 10)
        self.set_text_color(90, 90, 90)
        self.cell(0, 6, COMPANY_ADDRESS, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.set_font("Helvetica", "B", 12)
        self.set_text_color(0, 0, 0)
        self.cell(0, 8, "Vehicle & Driver Duty Tracker Report", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")


def build_duty_tracker_pdf(detail_df: pd.DataFrame, summary_metrics: dict, filters_summary: str) -> bytes:
    """PDF containing the KPI summary table followed by the detailed duty log.
    `detail_df` must already have the DUTY_TRACKER_DISPLAY_COLS columns (see
    the tab below for how it's built from the requisitions DataFrame).

    Every string written to the PDF is passed through sanitize_pdf_text()
    first — see that function's docstring for why this is necessary with
    FPDF's core Helvetica font.
    """
    pdf = DutyTrackerPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, f"Generated on: {datetime.now().strftime('%Y-%m-%d %I:%M %p')}",
              new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.multi_cell(0, 6, sanitize_pdf_text(filters_summary))
    pdf.ln(2)

    # ---- KPI Summary block ----
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Summary Metrics", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", size=9)
    pdf.set_fill_color(230, 230, 230)
    for label, value in summary_metrics.items():
        safe_label = sanitize_pdf_text(label)
        safe_value = sanitize_pdf_text(value)
        pdf.cell(0, 7, f"{safe_label}: {safe_value}", border=1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(6)

    # ---- Detailed duty log table ----
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Detailed Duty Log", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)

    headers = DUTY_TRACKER_DISPLAY_COLS
    col_widths = [30, 28, 34, 34, 18, 26, 107]  # sums to ~277mm, fits A4 landscape

    pdf.set_font("Helvetica", size=7)
    heading_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 98, 254))
    with pdf.table(col_widths=col_widths, text_align="LEFT", first_row_as_headings=True,
                   line_height=5, headings_style=heading_style, cell_fill_color=(245, 245, 245),
                   cell_fill_mode="ROWS") as table:
        header_row = table.row()
        for h in headers:
            header_row.cell(sanitize_pdf_text(h))
        for _, r in detail_df.iterrows():
            row = table.row()
            for col in headers:
                row.cell(sanitize_pdf_text(fmt(r.get(col, ""), "")))

    return bytes(pdf.output())


# =========================================================
# 5C. MANAGEMENT / EXECUTIVE DASHBOARD — EXCEL + PDF REPORT GENERATORS
# =========================================================
# ADDITIVE helpers backing the new "Management Dashboard" (Section 10C).
# Deliberately kept separate from every export helper above so nothing here
# touches, shadows, or changes the behaviour of the existing Admin exports
# (Tab 5 "All Requisitions & Export", the Duty Tracker, or the KM Variance
# Report). This dashboard is READ-ONLY: it never calls insert_requisition(),
# update_requisition(), or any Supabase write helper.

def build_management_excel(kpis: dict, dept_df: pd.DataFrame, detail_df: pd.DataFrame,
                            dept_duty_df: pd.DataFrame = None) -> bytes:
    """Formatted .xlsx export for the Management Dashboard: a 'KPI Summary'
    sheet, a 'Department Usage' sheet, an optional 'Department Duty Summary'
    sheet (requests/completed/KM/duty-hours per department), and a
    'Detailed Data' sheet with the full filtered requisition rows for the
    selected date range. `dept_duty_df` is optional and defaults to None so
    any existing caller that doesn't pass it keeps working unchanged."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        kpi_df = pd.DataFrame([{"Metric": k, "Value": v} for k, v in kpis.items()])
        kpi_df.to_excel(writer, index=False, sheet_name="KPI Summary")
        dept_df.to_excel(writer, index=False, sheet_name="Department Usage")
        sheets = [("KPI Summary", kpi_df), ("Department Usage", dept_df)]
        if dept_duty_df is not None:
            dept_duty_df.to_excel(writer, index=False, sheet_name="Department Duty Summary")
            sheets.append(("Department Duty Summary", dept_duty_df))
        detail_df.to_excel(writer, index=False, sheet_name="Detailed Data")
        sheets.append(("Detailed Data", detail_df))

        # `fillna("")` BEFORE `.astype(str)` matters on pandas >= 3.0: that
        # version stopped converting NaN/None to the literal string "nan" on
        # astype(str), leaving real NaN behind instead. A column that's
        # partially or entirely blank (e.g. admin_note, or numeric fields
        # like total_km before any trip is Completed) would then make
        # .str.len().max() return NaN, and int(NaN) raises
        # "ValueError: cannot convert float NaN to integer" — exactly the
        # error this fixes. Filling blanks with "" first guarantees every
        # length is a real integer (0 for blank cells).
        from openpyxl.utils import get_column_letter
        for sheet_name, sheet_df in sheets:
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(sheet_df.columns, start=1):
                width = max(12, min(40, int(sheet_df[col].fillna("").astype(str).str.len().max() if not sheet_df.empty else 12) + 2))
                ws.column_dimensions[get_column_letter(i)].width = width

    return buf.getvalue()


class ManagementPDF(FPDF):
    """Separate FPDF subclass so the Executive summary report's branding can
    evolve independently of the other report types in this file."""

    def header(self):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(15, 98, 254)
        self.cell(0, 9, COMPANY_NAME, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.set_font("Helvetica", "", 10)
        self.set_text_color(90, 90, 90)
        self.cell(0, 6, COMPANY_ADDRESS, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.set_font("Helvetica", "B", 12)
        self.set_text_color(0, 0, 0)
        self.cell(0, 8, "Executive Summary Report", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")


def build_management_pdf(kpis: dict, dept_df: pd.DataFrame, filters_summary: str,
                          dept_duty_df: pd.DataFrame = None) -> bytes:
    """Summary-focused PDF: generation timestamp + active filters, a KPI
    block (Total Requisitions / Completed / Pending / Total KM / etc.), a
    Department-wise Usage table, and — when provided — a Department Duty
    Summary table (requests, completed trips, total KM, and total duty
    hours per department). `dept_duty_df` is optional and defaults to None
    so any existing caller that doesn't pass it keeps working unchanged.
    Kept deliberately light (no full row-level dump) since this is meant as
    an at-a-glance Executive summary — full detail is in the Excel export."""
    pdf = ManagementPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, f"Generated on: {datetime.now().strftime('%Y-%m-%d %I:%M %p')}",
              new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.multi_cell(0, 6, sanitize_pdf_text(filters_summary))
    pdf.ln(2)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Key Performance Indicators", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", size=10)
    pdf.set_fill_color(230, 230, 230)
    for label, value in kpis.items():
        pdf.cell(0, 8, sanitize_pdf_text(f"{label}: {value}"), border=1, fill=True,
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(6)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Department-wise Vehicle Usage", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)

    if dept_df is not None and not dept_df.empty:
        pdf.set_font("Helvetica", size=9)
        heading_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 98, 254))
        with pdf.table(col_widths=[110, 60], text_align="LEFT", first_row_as_headings=True,
                       line_height=6, headings_style=heading_style, cell_fill_color=(245, 245, 245),
                       cell_fill_mode="ROWS") as table:
            header_row = table.row()
            header_row.cell("Department")
            header_row.cell("Requests")
            for _, r in dept_df.iterrows():
                row = table.row()
                row.cell(sanitize_pdf_text(r.get("Department", "")))
                row.cell(sanitize_pdf_text(r.get("Requests", "")))
    else:
        pdf.set_font("Helvetica", size=10)
        pdf.cell(0, 8, "No data available for the selected date range.",
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    if dept_duty_df is not None:
        pdf.ln(6)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 8, "Department Duty & Distance Summary", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(1)

        if not dept_duty_df.empty:
            pdf.set_font("Helvetica", size=9)
            heading_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 98, 254))
            cols = ["Department", "Total Requests", "Completed Trips", "Total KM", "Total Duty Hours"]
            with pdf.table(col_widths=[52, 32, 32, 30, 34], text_align="LEFT", first_row_as_headings=True,
                           line_height=6, headings_style=heading_style, cell_fill_color=(245, 245, 245),
                           cell_fill_mode="ROWS") as table:
                header_row = table.row()
                for h in cols:
                    header_row.cell(h)
                for _, r in dept_duty_df.iterrows():
                    row = table.row()
                    for c in cols:
                        row.cell(sanitize_pdf_text(r.get(c, "")))
        else:
            pdf.set_font("Helvetica", size=10)
            pdf.cell(0, 8, "No duty data available for the selected date range.",
                      new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    return bytes(pdf.output())


def render_management_dashboard(df_all: pd.DataFrame):
    """Read-only Executive/Management overview: a live 'today' snapshot,
    top-level KPI cards, department-wise usage & duty-time breakdown, a
    date-range filter, and Excel/PDF export buttons. Shared by both the
    Admin Dashboard's 'Management Dashboard' tab and the standalone
    'executive' role view — same function, same numbers, so the two roles
    never see different totals for the same range."""
    st.subheader("📈 Executive / Management Overview")

    if df_all.empty:
        st.info("No requisition data available yet.")
        return

    # ---------------------------------------------------------------
    # LIVE "TODAY" SNAPSHOT — always reflects the current moment, using the
    # full unfiltered dataset, regardless of whatever date range is picked
    # below. This answers "what's happening right now / today" separately
    # from the historical, range-filtered KPIs further down.
    # ---------------------------------------------------------------
    today = date.today()
    st.markdown(f"##### 📅 Today's Live Snapshot — {today.strftime('%A, %d %B %Y')}")
    st.caption("These counts are live and independent of the date-range filter below.")

    created_dt = pd.to_datetime(df_all.get("created_at"), errors="coerce", utc=True, format="mixed")
    return_dt = pd.to_datetime(df_all.get("actual_return_time"), errors="coerce", utc=True, format="mixed")
    today_utc = pd.Timestamp.now(tz="UTC").normalize()

    pending_now = int((df_all["status"] == "Pending").sum())
    on_trip_now = int((df_all["status"] == "On Trip").sum())
    requested_today = int((created_dt.dt.normalize() == today_utc).sum())
    completed_today = int((return_dt.dt.normalize() == today_utc).sum())

    # Each metric is paired with an expander right below it — clicking it
    # reveals the actual requisitions behind that number, since a plain
    # st.metric can't be clicked directly in Streamlit.
    t1, t2, t3, t4 = st.columns(4)

    with t1:
        st.metric("🟡 Pending Right Now", pending_now)
        with st.expander(f"🔍 View {pending_now} pending"):
            pending_rows = df_all[df_all["status"] == "Pending"]
            if pending_rows.empty:
                st.caption("No pending requisitions right now.")
            else:
                cols = ["id", "applicant_name", "department", "destination",
                        "date_of_travel", "time_of_travel", "purpose"]
                disp = pending_rows[[c for c in cols if c in pending_rows.columns]].copy()
                if "id" in disp.columns:
                    disp["id"] = disp["id"].apply(short_req_id)
                    disp = disp.rename(columns={"id": "Req #"})
                if "time_of_travel" in disp.columns:
                    disp["time_of_travel"] = disp["time_of_travel"].apply(lambda v: fmt_time_12h(v, v))
                st.dataframe(disp, use_container_width=True, hide_index=True,
                             height=min(320, 45 + 35 * len(disp)))

    with t2:
        st.metric("🔵 Vehicles On Trip Now", on_trip_now)
        with st.expander(f"🔍 View {on_trip_now} on trip"):
            on_trip_rows = df_all[df_all["status"] == "On Trip"]
            if on_trip_rows.empty:
                st.caption("No vehicles are currently on a trip.")
            else:
                cols = ["id", "applicant_name", "driver_name", "vehicle_number",
                        "destination", "actual_exit_time"]
                disp = on_trip_rows[[c for c in cols if c in on_trip_rows.columns]].copy()
                if "id" in disp.columns:
                    disp["id"] = disp["id"].apply(short_req_id)
                    disp = disp.rename(columns={"id": "Req #"})
                if "actual_exit_time" in disp.columns:
                    disp["actual_exit_time"] = disp["actual_exit_time"].apply(lambda v: fmt_time_12h(v, v))
                st.dataframe(disp, use_container_width=True, hide_index=True,
                             height=min(320, 45 + 35 * len(disp)))

    with t3:
        st.metric("📥 Requests Submitted Today", requested_today)
        with st.expander(f"🔍 View {requested_today} submitted today"):
            submitted_today_rows = df_all[created_dt.dt.normalize() == today_utc]
            if submitted_today_rows.empty:
                st.caption("No requests submitted yet today.")
            else:
                cols = ["id", "applicant_name", "department", "destination",
                        "status", "created_at"]
                disp = submitted_today_rows[[c for c in cols if c in submitted_today_rows.columns]].copy()
                if "id" in disp.columns:
                    disp["id"] = disp["id"].apply(short_req_id)
                    disp = disp.rename(columns={"id": "Req #"})
                if "created_at" in disp.columns:
                    disp["created_at"] = disp["created_at"].apply(lambda v: fmt_time_12h(v, v))
                st.dataframe(disp, use_container_width=True, hide_index=True,
                             height=min(320, 45 + 35 * len(disp)))

    with t4:
        st.metric("✅ Trips Completed Today", completed_today)
        with st.expander(f"🔍 View {completed_today} completed today"):
            completed_today_rows = df_all[return_dt.dt.normalize() == today_utc]
            if completed_today_rows.empty:
                st.caption("No trips completed yet today.")
            else:
                disp = completed_today_rows.copy()
                disp["Total KM"] = disp.apply(lambda row: effective_km_fields(row)[2], axis=1)
                cols = ["id", "applicant_name", "driver_name", "vehicle_number",
                        "destination", "actual_return_time", "Total KM"]
                disp = disp[[c for c in cols if c in disp.columns]].copy()
                if "id" in disp.columns:
                    disp["id"] = disp["id"].apply(short_req_id)
                    disp = disp.rename(columns={"id": "Req #"})
                if "actual_return_time" in disp.columns:
                    disp["actual_return_time"] = disp["actual_return_time"].apply(lambda v: fmt_time_12h(v, v))
                st.dataframe(disp, use_container_width=True, hide_index=True,
                             height=min(320, 45 + 35 * len(disp)))

    work_df = df_all.copy()
    work_df["_dt"] = pd.to_datetime(work_df["date_of_travel"], errors="coerce")
    min_d = work_df["_dt"].min()
    max_d = work_df["_dt"].max()
    default_start = min_d.date() if pd.notnull(min_d) else date.today()
    default_end = max_d.date() if pd.notnull(max_d) else date.today()

    st.markdown("---")
    st.markdown("##### 🔎 Date Range Filter")
    dc1, dc2 = st.columns(2)
    with dc1:
        mgmt_start = st.date_input("Start Date", value=default_start, key="mgmt_start_date")
    with dc2:
        mgmt_end = st.date_input("End Date", value=default_end, key="mgmt_end_date")

    if mgmt_start > mgmt_end:
        st.error("⚠️ Start Date must be on or before End Date.")
        return

    filtered = work_df[
        (work_df["_dt"] >= pd.Timestamp(mgmt_start)) & (work_df["_dt"] <= pd.Timestamp(mgmt_end))
    ].copy()

    total_requisitions = len(filtered)
    completed_trips = int((filtered["status"] == "Completed").sum()) if not filtered.empty else 0
    pending_requests = int((filtered["status"] == "Pending").sum()) if not filtered.empty else 0

    if not filtered.empty:
        completed_rows = filtered[filtered["status"] == "Completed"]
        total_distance = (
            float(completed_rows.apply(lambda row: effective_km_fields(row)[2], axis=1).sum())
            if not completed_rows.empty else 0.0
        )
    else:
        total_distance = 0.0

    st.markdown("---")
    st.markdown("##### 📌 Key Performance Indicators")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📋 Total Requisitions", total_requisitions)
    k2.metric("✅ Completed Trips", completed_trips)
    k3.metric("🟡 Pending Requests", pending_requests)
    k4.metric("🛣️ Total Distance (KM)", f"{total_distance:.1f}")

    st.markdown("---")
    st.markdown("##### 📊 Department-wise Vehicle Usage")
    if filtered.empty:
        st.info("No data available for the selected date range.")
        dept_counts = pd.DataFrame(columns=["Department", "Requests"])
    else:
        dept_counts = filtered["department"].value_counts().reset_index()
        dept_counts.columns = ["Department", "Requests"]
        fig = px.bar(dept_counts, x="Department", y="Requests", color="Department", text="Requests",
                     title="Department-wise Vehicle Usage")
        fig.update_layout(showlegend=False, height=380)
        st.plotly_chart(fig, use_container_width=True, key="mgmt_dept_chart")

    # ---------------------------------------------------------------
    # DEPARTMENT-WISE DUTY TIME — how many requests each department made,
    # how many completed, how much distance was covered, and how long
    # vehicles were actually out (duty hours) for that department, within
    # the selected date range. Duty hours use the same actual_exit_time /
    # actual_return_time fields as the Duty Tracker tab (compute_duty_hours),
    # so an in-progress ("On Trip") duty still counts up to right now.
    # ---------------------------------------------------------------
    st.markdown("---")
    st.markdown("##### ⏱️ Department-wise Vehicle Usage Time (Duty Hours)")

    if filtered.empty:
        st.info("No data available for the selected date range.")
        dept_duty_df = pd.DataFrame(columns=["Department", "Total Requests", "Completed Trips",
                                              "Total KM", "Total Duty Hours"])
    else:
        calc = filtered.copy()
        calc["_km"] = calc.apply(lambda row: effective_km_fields(row)[2], axis=1)
        calc["_hrs"] = calc.apply(compute_duty_hours, axis=1)
        dept_duty_df = (
            calc.groupby("department")
            .agg(
                **{
                    "Total Requests": ("request_id", "count"),
                    "Completed Trips": ("status", lambda s: int((s == "Completed").sum())),
                    "Total KM": ("_km", "sum"),
                    "Total Duty Hours": ("_hrs", "sum"),
                }
            )
            .reset_index()
            .rename(columns={"department": "Department"})
        )
        dept_duty_df["Total KM"] = dept_duty_df["Total KM"].round(1)
        dept_duty_df["Total Duty Hours"] = dept_duty_df["Total Duty Hours"].round(1)
        dept_duty_df = dept_duty_df.sort_values("Total Requests", ascending=False).reset_index(drop=True)

        b1, b2 = st.columns(2)
        busiest_dept = dept_duty_df.loc[dept_duty_df["Total Requests"].idxmax(), "Department"]
        most_hours_dept = dept_duty_df.loc[dept_duty_df["Total Duty Hours"].idxmax(), "Department"]
        b1.metric("🏆 Busiest Department (by Requests)", busiest_dept)
        b2.metric("⏱️ Most Vehicle-Hours Department", most_hours_dept)

        st.dataframe(dept_duty_df, use_container_width=True, hide_index=True, height=280)

        fig_hrs = px.bar(dept_duty_df, x="Department", y="Total Duty Hours", color="Department",
                          text="Total Duty Hours", title="Department-wise Vehicle Duty Hours")
        fig_hrs.update_layout(showlegend=False, height=380)
        st.plotly_chart(fig_hrs, use_container_width=True, key="mgmt_dept_duty_chart")

    st.markdown("---")
    st.markdown("##### ⬇️ Download Reports")

    filtered_display = filtered.drop(columns=["_dt"], errors="ignore").copy()
    if "time_of_travel" in filtered_display.columns:
        filtered_display["time_of_travel"] = filtered_display["time_of_travel"].apply(lambda v: fmt_time_12h(v, v))

    kpis = {
        "Date Range": f"{mgmt_start.strftime('%Y-%m-%d')} to {mgmt_end.strftime('%Y-%m-%d')}",
        "Total Requisitions": total_requisitions,
        "Completed Trips": completed_trips,
        "Pending Requests": pending_requests,
        "Total Distance (KM)": f"{total_distance:.1f}",
    }
    filters_summary = f"Date Range: {mgmt_start.strftime('%Y-%m-%d')} to {mgmt_end.strftime('%Y-%m-%d')}"

    e1, e2 = st.columns(2)
    with e1:
        mgmt_excel = build_management_excel(kpis, dept_counts, filtered_display, dept_duty_df)
        st.download_button(
            "⬇️ Download Detailed Excel Report (.xlsx)", data=mgmt_excel,
            file_name="management_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="mgmt_excel_dl",
        )
    with e2:
        mgmt_pdf = build_management_pdf(kpis, dept_counts, filters_summary, dept_duty_df)
        st.download_button(
            "⬇️ Download Summary PDF Report (.pdf)", data=mgmt_pdf,
            file_name="management_summary_report.pdf",
            mime="application/pdf",
            use_container_width=True, key="mgmt_pdf_dl",
        )


# =========================================================
# 6. SESSION STATE INIT (with "Remember Me" cookie restore)
# =========================================================
# Instantiated exactly once per script run — the underlying component uses a
# fixed key, and Streamlit errors on duplicate keys within a single run, so
# every other place in this file that needs cookies reuses this same object
# rather than calling get_cookie_manager() again.
cookie_manager = get_cookie_manager()

if "auth_user" not in st.session_state:
    # extra_streamlit_components's CookieManager runs inside a browser iframe
    # component. On the very first script run after a fresh page load, that
    # component hasn't finished round-tripping "here are the browser's
    # cookies" back to Python yet — cookie_manager.get_all() returns None
    # (NOT an empty dict) while it's still waiting. If we treat that None as
    # "no cookie exists" and immediately show the login form, the remembered
    # session is thrown away on every single page load, which is exactly the
    # "Remember me doesn't work" symptom. An empty dict {} — as opposed to
    # None — genuinely means "component is ready, browser has no cookie".
    all_cookies = cookie_manager.get_all()

    # A SINGLE 300ms retry isn't always enough — on a slower device/browser
    # the cookie iframe can take longer than that to finish its first
    # round-trip, and once the single retry is used up, `all_cookies` is
    # still None but gets treated as "no cookie" (see `(all_cookies or {})`
    # below), so the login page flashes on screen before the NEXT page
    # load/refresh finally picks up the cookie and jumps to the main app.
    # Retrying up to 6 times (up to ~1.8s total) closes that gap so a
    # remembered device goes straight to the main dashboard, with only a
    # brief "Restoring your session..." placeholder instead of the login form.
    COOKIE_BOOTSTRAP_MAX_RETRIES = 6
    cookie_retries = st.session_state.get("_cookie_bootstrap_retries", 0)

    if all_cookies is None and cookie_retries < COOKIE_BOOTSTRAP_MAX_RETRIES:
        st.session_state["_cookie_bootstrap_retries"] = cookie_retries + 1
        st.info("🔄 Restoring your session...")
        # A fresh key each retry (rather than one static key) is what makes
        # this fire again on every subsequent short-lived rerun — reusing
        # the same key would hit that key's own `limit=1` and stop retrying
        # after the first attempt.
        st_autorefresh(interval=300, limit=1, key=f"cookie_bootstrap_refresh_{cookie_retries}")
        st.stop()

    restored_user = None
    session_token = (all_cookies or {}).get(SESSION_COOKIE_NAME)
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
ensure_daily_shuttle_requisitions()  # auto-creates today's 6 fixed shuttle requisitions, once per session/day

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

auto_refresh_on = st.sidebar.checkbox("🔄 Auto-refresh every 60s", value=True,
                                       help="Automatically reloads live data across the app. "
                                            "Turn off temporarily if you're filling out a long form, "
                                            "or use '🔄 Refresh Now' any time for an on-demand update.")
if st.sidebar.button("🔄 Refresh Now", use_container_width=True):
    st.rerun()
st.sidebar.markdown("---")
logout_button()

# Auto-refresh is skipped only for the Gate Officer role, since that
# dashboard is spent almost entirely typing Gate In/Out odometer entries —
# a background rerun mid-typing was causing focus loss / lost keystrokes
# there. Every other role (including Driver) gets the live auto-refresh.
# Everyone can still hit "🔄 Refresh Now" above for an on-demand update, and
# every write already clears the relevant cache so approvals/gate actions
# show up instantly regardless of this setting.
# Interval set to 60s, with the checkbox still OFF by default (see above) —
# so idle screens don't rerun in the background unless someone explicitly
# turns this on, but when they do, it checks for updates every 60s rather
# than every 120s.
# Every write path still calls its matching _clear_*_caches() immediately,
# so approvals/gate actions/driver KM entries still show up instantly for
# the person who made the change; this setting only controls how often
# *other* idle screens passively refresh to see someone else's changes.
NO_AUTOREFRESH_ROLES = {"gate_officer"}
if auto_refresh_on and user["role"] not in NO_AUTOREFRESH_ROLES:
    st_autorefresh(interval=60_000, key="global_autorefresh")

# =========================================================
# 8. EMPLOYEE DASHBOARD
# =========================================================
if user["role"] == "user":
    company_header("👤 Employee Dashboard")
    tab1, tab2 = st.tabs(["📋 New Requisition", "📍 My Requests / Live Status"])

    with tab1:
        st.subheader("Submit a New Vehicle Requisition")
        st.caption(
            "👤 Applicant Name, Department, and Mobile Number are auto-filled from your "
            "profile — no need to re-select or re-type them every time. You can still "
            "adjust them below if this trip needs different details."
        )
        with st.form("req_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                # Auto-filled directly from the logged-in session's profile data
                # (st.session_state.auth_user) so employees never have to
                # re-select their department or retype their name/mobile number
                # for every new requisition.
                applicant_name = st.text_input("Applicant Name *", value=user.get("full_name", ""))
                department = st.selectbox(
                    "Department *", DEPARTMENTS,
                    index=DEPARTMENTS.index(user["department"]) if user.get("department") in DEPARTMENTS else 0,
                )
                mobile_number = st.text_input("Mobile Number *", value=user.get("mobile", ""),
                                               placeholder="01XXXXXXXXX")
                passenger_count = st.number_input("Passenger Count *", min_value=1, max_value=50, value=1)
            with c2:
                date_of_travel = st.date_input("Date of Travel *", min_value=date.today())
                time_of_travel = time_input_12h("Time of Travel *", key_prefix="new_req_tt")
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
                        new_id = insert_requisition(data)
                        st.success(f"✅ Requisition submitted! Your Requisition number is **{short_req_id(new_id)}**")
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
                        <b>Requisition {short_req_id(r.get('id'))}</b> &nbsp;|&nbsp; {r['destination']} &nbsp;|&nbsp;
                        {r['date_of_travel']} at {fmt_time_12h(r['time_of_travel'])} &nbsp;&nbsp;
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
                                f" (rescheduled from {fmt_time_12h(r['time_of_travel'])})"
                                if not is_blank(r.get("approved_time")) and r.get("approved_time") != r["time_of_travel"]
                                else ""
                            )
                            st.markdown(f"""
                            <div class="driver-box">
                            🚘 <b>Driver:</b> {fmt(r['driver_name'], 'TBD')} &nbsp;|&nbsp;
                            📞 <b>Contact:</b> {fmt(r['driver_contact'], 'TBD')} &nbsp;|&nbsp;
                            🔢 <b>Vehicle No.:</b> {fmt(r['vehicle_number'], 'TBD')} &nbsp;|&nbsp;
                            🕒 <b>Approved Departure Time:</b> {fmt_time_12h(approved_time)}{time_note}
                            </div>
                            """, unsafe_allow_html=True)
                            if not is_blank(r.get("admin_note")):
                                st.caption(f"📝 Admin Note: {r['admin_note']}")

                        if r["status"] in ("On Trip", "Completed"):
                            st.write(f"**Gate Out (Actual Exit):** {fmt_time_12h(r.get('actual_exit_time'))}  |  **Start KM:** {fmt(r.get('start_km'))}")
                        if r["status"] == "Completed":
                            st.write(f"**Gate In (Actual Return):** {fmt_time_12h(r.get('actual_return_time'))}  |  **End KM:** {fmt(r.get('end_km'))}  |  **Total KM:** {fmt(r.get('total_km'))}")

                        if r["status"] == "Rejected":
                            st.error("This request was rejected by the admin.")
                            if not is_blank(r.get("admin_note")):
                                st.caption(f"📝 Reason: {r['admin_note']}")

# =========================================================
# 9. GATE OFFICER DASHBOARD — Vehicle Gate In / Gate Out Panel
# =========================================================
elif user["role"] == "gate_officer":
    company_header("🛡️ Gate Officer Dashboard — Vehicle Gate Panel")
    st.caption(f"Logged in as {user['full_name']} — Gate Officer")

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
                        st.write(f"**Requested Time:** {r['date_of_travel']} at {fmt_time_12h(r['time_of_travel'])}")
                        st.write(f"**Admin Approved Time:** {fmt_time_12h(approved_time)}")
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
                            }, notify=False)
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
                        st.write(f"**Gate Out Time:** {fmt_time_12h(r.get('actual_exit_time'))}")
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
                                }, notify=False)
                                st.success(f"✅ Gate In recorded for {r['applicant_name']}. Total distance: **{total_km} KM**")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Failed to record Gate In: {e}")

# =========================================================
# 9B. DRIVER DASHBOARD — My Assigned Trips & Odometer Entry (NEW)
# =========================================================
# Mirrors the Gate Officer panel's two-tab design exactly: a "Start
# Trip" tab (like the Gate Officer's Gate Out) that only takes Start KM, and
# a separate "End Trip" tab (like the Gate Officer's Gate In) that only
# takes End KM once the driver is back — instead of one combined form.
# Submitting Start KM alone moves the trip Approved -> On Trip; submitting
# End KM later moves it -> Completed (see submit_driver_km()).
elif user["role"] == "driver":
    company_header("🚙 Driver Dashboard — My Assigned Trips")
    st.caption(f"Logged in as {user['full_name']} — Driver")
    st.info(
        "This view only shows trips where the **Driver Name** on the requisition "
        "matches your registered Full Name. If a trip you drove isn't listed here, "
        "ask an Admin to check the name spelling in **Manage Drivers & Vehicles**."
    )

    tab_depart, tab_return = st.tabs(["🚦 Start Trip (Approved)", "🏁 End Trip (Return)"])

    with st.spinner("Loading your assigned trips..."):
        my_trips = fetch_requisitions_by_driver(user["full_name"])

    # ---------------- TAB 1: Start Trip (Start KM only) ----------------
    with tab_depart:
        st.subheader("Approved Trips Awaiting Your Start KM")
        start_trips = my_trips[my_trips["status"] == "Approved"] if not my_trips.empty else my_trips

        if start_trips.empty:
            st.info("You have no Approved trips waiting to start.")
        else:
            for _, r in start_trips.iterrows():
                auto_start = get_last_driver_end_km(user["full_name"], r.get("vehicle_number", ""))
                with st.expander(
                    f"🟢 Requisition {short_req_id(r.get('id'))} — {r['destination']}  |  "
                    f"Vehicle: {fmt(r.get('vehicle_number'), 'N/A')}"
                ):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.write(f"**Applicant:** {r['applicant_name']} ({r['department']})")
                        st.write(f"**Date/Time:** {r['date_of_travel']} at {fmt_time_12h(r['time_of_travel'])}")
                    with c2:
                        approved_time = r["time_of_travel"] if is_blank(r.get("approved_time")) else r.get("approved_time")
                        st.write(f"**Approved Departure Time:** {fmt_time_12h(approved_time)}")
                        st.write(f"**Vehicle Type:** {r['vehicle_type']}")
                    if not is_blank(r.get("admin_note")):
                        st.caption(f"📝 Admin Notes: {r['admin_note']}")

                    if auto_start and auto_start > 0:
                        st.caption(
                            f"↩️ Auto-filled from your last logged End KM for "
                            f"**{fmt(r.get('vehicle_number'))}** — change it below if needed."
                        )

                    with st.form(f"driver_start_{r['request_id']}"):
                        d_start_km = st.number_input(
                            "Start KM (Odometer Reading) *", min_value=0.0, step=1.0, format="%.1f",
                            value=float(auto_start), key=f"dstart_{r['request_id']}",
                        )
                        depart_clicked = st.form_submit_button(
                            "🚦 Start Trip / Depart", type="primary", use_container_width=True
                        )

                    if depart_clicked:
                        try:
                            submit_driver_km(r.to_dict(), driver_start_km=d_start_km, driver_end_km=None)
                            st.success(
                                f"✅ Trip started for {r['applicant_name']} — status is now On Trip. "
                                "A Telegram alert has been sent."
                            )
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to record Start KM: {e}")

    # ---------------- TAB 2: End Trip (End KM only) ----------------
    with tab_return:
        st.subheader("Trips Currently On the Road")

        def _needs_end(row):
            if row.get("status") == "On Trip":
                return True
            # Also surface trips already marked Completed via the Gate
            # Officer's own Gate-In, as long as the driver hasn't logged
            # their own End KM yet — so a fast Gate-Officer entry never
            # locks the driver out of finishing their own log (needed for
            # the KM Variance Report to have both sides).
            if row.get("status") == "Completed" and is_blank(row.get("driver_end_km")):
                return True
            return False

        end_trips = my_trips[my_trips.apply(_needs_end, axis=1)] if not my_trips.empty else my_trips

        if end_trips.empty:
            st.info("No trips are currently waiting for your End KM.")
        else:
            for _, r in end_trips.iterrows():
                # `is_blank()` here (not just checking for None) matters:
                # a trip that Admin/Gate Officer already marked Completed
                # directly — without the driver ever logging a Start KM —
                # has driver_start_km as pandas NaN, not None or 0. Treating
                # that as "no baseline yet" (has_own_start = False) is what
                # lets us show a Start KM field below instead of silently
                # defaulting to 0.0 and later trying to write NaN to
                # Supabase, which is exactly what caused the
                # "Out of range float values are not JSON compliant: nan"
                # crash on trips like this one.
                has_own_start = not is_blank(r.get("driver_start_km"))
                start_km_val = float(r.get("driver_start_km")) if has_own_start else 0.0
                with st.expander(
                    f"🔵 Requisition {short_req_id(r.get('id'))} — {r['destination']}  |  "
                    f"Vehicle: {fmt(r.get('vehicle_number'), 'N/A')}"
                ):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.write(f"**Applicant:** {r['applicant_name']} ({r['department']})")
                        st.write(f"**Vehicle:** {fmt(r.get('vehicle_number'), 'N/A')} ({r['vehicle_type']})")
                        st.write(f"**Destination:** {r['destination']}")
                    with c2:
                        st.write(f"**Your Start KM:** {fmt(r.get('driver_start_km'))}")
                        st.write(f"**Trip Started:** {fmt_time_12h(r.get('actual_exit_time'))}")

                    if not has_own_start:
                        st.warning(
                            "⚠️ You don't have a Start KM logged for this trip yet (it was likely "
                            "completed directly by Admin/Gate Officer). Please enter BOTH your "
                            "Start KM and End KM below so this trip has a proper distance on record."
                        )

                    with st.form(f"driver_end_{r['request_id']}"):
                        if not has_own_start:
                            d_start_km = st.number_input(
                                "Start KM (Odometer Reading) *", min_value=0.0, step=1.0, format="%.1f",
                                key=f"dend_start_{r['request_id']}",
                            )
                        else:
                            d_start_km = start_km_val
                        d_end_km = st.number_input(
                            "End KM (Odometer Reading) *", min_value=0.0, step=1.0, format="%.1f",
                            help=(
                                f"Must be greater than or equal to your Start KM ({start_km_val:.1f})."
                                if has_own_start else
                                "Must be greater than or equal to the Start KM you enter above."
                            ),
                            key=f"dend_{r['request_id']}",
                        )
                        return_clicked = st.form_submit_button(
                            "🏁 Complete Trip / Return", type="primary", use_container_width=True
                        )

                    if return_clicked:
                        if d_end_km < d_start_km:
                            st.error("End KM cannot be less than Start KM.")
                        else:
                            try:
                                submit_driver_km(
                                    r.to_dict(),
                                    driver_start_km=d_start_km,
                                    driver_end_km=d_end_km,
                                )
                                st.success(
                                    f"✅ Trip completed for {r['applicant_name']}. "
                                    f"Distance: **{d_end_km - d_start_km:.1f} KM**. "
                                    "A Telegram alert has been sent."
                                )
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Failed to record End KM: {e}")

# =========================================================
# 9C. NURSE / SENIOR NURSE — EMERGENCY DASHBOARD (NEW)
# =========================================================
# A deliberately minimal dashboard: Medical staff need to arrange a vehicle
# for a patient FAST, not fill out a multi-field form. The Emergency tab is
# a single button — one click submits a fully pre-filled Pending
# requisition (today's date/time, "Hospital / Emergency" destination,
# "HIACE" vehicle type, 1 passenger) and fires an extra high-visibility
# Telegram alert on top of the normal "New Requisition" one, so Admin/
# the Gate Officer notice it immediately. It reuses insert_requisition()
# (not a raw Supabase call) so it still shows up in Admin's Pending Requests
# queue for driver/vehicle assignment exactly like any other request.
elif user["role"] == "nurse":
    company_header("🚨 Medical Emergency Vehicle Request")
    st.caption(f"Logged in as {user['full_name']} — {user.get('designation') or 'Nurse'}")

    tab_emergency, tab_my_emergency_requests = st.tabs(
        ["🚨 Emergency Request", "📍 My Requests / Live Status"]
    )

    with tab_emergency:
        st.markdown("### 🚑 Need a vehicle right now to carry a patient?")
        st.write(
            "Tap the button below to submit an emergency vehicle request immediately — "
            "no form to fill in. Admin and the Gate Officer are notified right away to arrange "
            "a driver and vehicle."
        )
        if st.button("🚨 EMERGENCY — Request Vehicle for Patient Carry", type="primary",
                      use_container_width=True):
            request_id = generate_request_id()
            now = datetime.now()
            data = {
                "request_id": request_id,
                "username": user["username"],
                "applicant_name": user.get("full_name", ""),
                "department": user.get("department") or "Medical",
                "mobile_number": user.get("mobile", ""),
                "date_of_travel": str(now.date()),
                "time_of_travel": now.strftime("%H:%M"),
                "destination": "Hospital / Emergency",
                "passenger_count": 1,
                "vehicle_type": "HIACE",
                "purpose": "🚨 EMERGENCY — Patient Carry",
                "special_request": "",
                "status": "Pending",
                "driver_name": "",
                "driver_contact": "",
                "vehicle_number": "",
                "approved_by": "",
            }
            with st.spinner("Sending emergency request..."):
                try:
                    new_id = insert_requisition(data)
                    # Extra, high-visibility alert on top of insert_requisition()'s
                    # normal "New Vehicle Requisition Submitted!" message, so an
                    # emergency doesn't blend in with routine requests.
                    send_telegram_alert(
                        "🚨🚑 **EMERGENCY VEHICLE REQUEST — PATIENT CARRY** 🚑🚨\n\n"
                        f"🆔 **Requisition #:** {short_req_id(new_id)}\n"
                        f"👤 **Requested by:** {user.get('full_name', '')} "
                        f"({user.get('designation') or 'Nurse'})\n"
                        f"📞 **Contact:** {user.get('mobile', '')}\n"
                        f"⏰ **Time:** {now.strftime('%Y-%m-%d %I:%M %p')}\n\n"
                        "⚡ Please arrange a driver and vehicle IMMEDIATELY."
                    )
                    st.success(
                        f"✅ Emergency request **{short_req_id(new_id)}** sent! "
                        "Admin/Gate Officer have been alerted."
                    )
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ Failed to send emergency request: {e}")

    with tab_my_emergency_requests:
        st.subheader("My Requests — Live Status")
        with st.spinner("Loading your requests..."):
            my_df = fetch_requisitions_by_user(user["username"])
        if my_df.empty:
            st.info("You haven't submitted any requests yet.")
        else:
            for _, r in my_df.iterrows():
                with st.container():
                    st.markdown(f"""
                    <div class="req-card">
                        <b>Requisition {short_req_id(r.get('id'))}</b> &nbsp;|&nbsp; {r['destination']} &nbsp;|&nbsp;
                        {r['date_of_travel']} at {fmt_time_12h(r['time_of_travel'])} &nbsp;&nbsp;
                        <span class="{badge_class(r['status'])}">{STATUS_BADGE.get(r['status'], r['status'])}</span>
                    </div>
                    """, unsafe_allow_html=True)
                    with st.expander("View details"):
                        st.write(f"**Purpose:** {r['purpose']}")
                        if r["status"] in ("Approved", "On Trip", "Completed"):
                            st.markdown(f"""
                            <div class="driver-box">
                            🚘 <b>Driver:</b> {fmt(r['driver_name'], 'TBD')} &nbsp;|&nbsp;
                            📞 <b>Contact:</b> {fmt(r['driver_contact'], 'TBD')} &nbsp;|&nbsp;
                            🔢 <b>Vehicle No.:</b> {fmt(r['vehicle_number'], 'TBD')}
                            </div>
                            """, unsafe_allow_html=True)
                        if r["status"] == "Rejected":
                            st.error("This request was rejected by the admin.")
                            if not is_blank(r.get("admin_note")):
                                st.caption(f"📝 Reason: {r['admin_note']}")

# =========================================================
# 9D. EXECUTIVE / MANAGEMENT DASHBOARD (NEW)
# =========================================================
# Standalone, READ-ONLY role for senior management: top-level KPI cards,
# department-wise usage chart, a date-range filter, and Excel/PDF export
# buttons — reusing render_management_dashboard() so 'executive' and
# 'admin' (via the "Management Dashboard" tab below) always see identical
# numbers for the same filters. Executives get none of the approve/reject,
# user-management, or fleet-management capabilities of the Admin Dashboard.
elif user["role"] == "executive":
    company_header("📈 Executive Dashboard")
    st.caption(f"Logged in as {user['full_name']} — Executive / Management")
    with st.spinner("Loading requisition data..."):
        exec_df_all = fetch_all_requisitions()
    render_management_dashboard(exec_df_all)

# =========================================================
# 10. ADMIN DASHBOARD
# =========================================================
elif user["role"] == "admin":
    company_header("🔐 Admin Dashboard")
    st.caption(f"Logged in as {user['full_name']} — Admin")

    # Reordered per business requirement: "Pending Requests" is now the first
    # (default) tab an admin sees. Visual order comes purely from this label
    # list — each variable below is named for what it holds, not its position,
    # so the underlying tab bodies didn't need to be reshuffled in the file.
    tab_pending_req, tab_create_req, tab_shuttle, tab_users, tab_pending_users, tab_analytics, tab_export, \
        tab_edit_trip, tab_fleet, tab_duty, tab_variance, tab_management = st.tabs([
            "🚗 Pending Requests", "➕ Create Requisition", "🚌 Staff Shuttle", "👥 User List", "⏳ ID Requests",
            "📊 Analytics", "📁 All Requisitions & Export", "✏️ Edit / Delete Trip",
            "🚘 Manage Drivers & Vehicles",
            "🕒 Duty Tracker & Analytics", "📈 KM Variance Report",
            "🏆 Management Dashboard",
        ])

    # Hoisted above every tab body so both users_df and df_all are guaranteed
    # ready regardless of which tab runs first in code — previously these
    # were only fetched inside specific tab bodies (old Tab 1 / old Tab 3),
    # which would have broken once those tabs were reordered.
    with st.spinner("Loading users and requisitions..."):
        users_df = fetch_all_users()
        df_all = fetch_all_requisitions()

    # ---------------- Pending Requests (was Tab 3) ----------------
    with tab_pending_req:
        st.subheader("Requisitions Awaiting Action")
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

        # ---------------------------------------------------------------
        # BULK ASSIGN — approve several Pending requisitions at once with
        # ONE shared Driver + Vehicle + Departure Time, instead of opening
        # each one individually. Useful when multiple people are going on
        # the same shuttle/HIACE run together. This is purely additive: the
        # per-request Approve/Reject cards below still work exactly as
        # before for one-at-a-time decisions.
        # ---------------------------------------------------------------
        if not pending_df.empty and not drivers_df.empty and not vehicles_df.empty:
            with st.expander("🚐 Bulk Assign Vehicle & Driver (Multiple Requests at Once)"):
                st.caption(
                    "Select two or more Pending requests below, pick ONE Driver, Vehicle, and "
                    "Departure Time, and approve all of them together in a single click — handy "
                    "when several people are sharing the same trip."
                )
                bulk_option_map = {
                    f"{short_req_id(r.get('id'))} — {r['applicant_name']} ({r['department']}) → "
                    f"{r['destination']} @ {fmt_time_12h(r['time_of_travel'])}": r["request_id"]
                    for _, r in pending_df.iterrows()
                }
                bulk_selected_labels = st.multiselect(
                    "Select Pending Requests to Bulk-Approve", list(bulk_option_map.keys()),
                    key="bulk_assign_select",
                )

                bd1, bd2 = st.columns(2)
                with bd1:
                    bulk_driver = st.selectbox(
                        "Driver Name", driver_options or ["No drivers available"],
                        key="bulk_assign_driver", disabled=not driver_options,
                    )
                with bd2:
                    bulk_contact = driver_contact_map.get(bulk_driver, "")
                    st.text_input("Driver Contact (auto-filled)", value=bulk_contact, disabled=True,
                                  key="bulk_assign_driver_contact")
                bulk_vehicle = st.selectbox(
                    "Vehicle Number", vehicle_options or ["No vehicles available"],
                    key="bulk_assign_vehicle", disabled=not vehicle_options,
                )
                bulk_time = time_input_12h(
                    "Approved Departure Time (applied to every selected request)",
                    key_prefix="bulk_assign_time",
                )
                bulk_note = st.text_area(
                    "Admin Note / Remarks (optional, applied to every selected request)",
                    key="bulk_assign_note",
                )

                bulk_ready = (
                    len(bulk_selected_labels) >= 2
                    and bulk_driver != "— Select Driver —"
                    and bulk_vehicle != "— Select Vehicle —"
                )
                if st.button(
                    f"✅ Approve & Assign {len(bulk_selected_labels)} Selected Requests",
                    type="primary", use_container_width=True, disabled=not bulk_ready,
                    key="bulk_assign_submit",
                ):
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    bulk_updates = {
                        "status": "Approved",
                        "driver_name": bulk_driver,
                        "driver_contact": driver_contact_map.get(bulk_driver, ""),
                        "vehicle_number": bulk_vehicle,
                        "approved_by": user["full_name"],
                        "action_timestamp": now_str,
                        "approved_time": bulk_time.strftime("%H:%M"),
                        "admin_note": bulk_note.strip(),
                    }
                    done, failed = [], []
                    with st.spinner("Approving selected requests..."):
                        for label in bulk_selected_labels:
                            req_id = bulk_option_map[label]
                            try:
                                update_requisition(req_id, bulk_updates)
                                done.append(label)
                            except Exception as e:
                                failed.append(f"{label} ({e})")
                    if done:
                        st.success(
                            f"✅ Approved {len(done)} requests with Driver **{bulk_driver}** / "
                            f"Vehicle **{bulk_vehicle}**."
                        )
                    if failed:
                        st.error("⚠️ Failed:\n\n" + "\n".join(f"- {f}" for f in failed))
                    if done:
                        st.rerun()
                elif not bulk_ready and bulk_selected_labels:
                    st.caption("⚠️ Select at least 2 requests and choose a Driver and Vehicle to enable bulk approval.")

        if pending_df.empty:
            st.success("🎉 No pending requisitions — all caught up!")
        else:
            # Rendering each pending request as its own expander + form (with
            # driver/vehicle selects + text inputs) is the single heaviest
            # thing this tab does — every one of those widgets gets rebuilt
            # on every rerun. Capping the list to the most recent 50 keeps
            # the page fast even when the Pending queue grows into the
            # hundreds; pending_df is already sorted newest-first (it comes
            # from df_all, which is ordered by created_at desc), so this
            # never hides the oldest/most-overdue requests — those surface
            # first as the newer ones above them get approved/rejected.
            PENDING_DISPLAY_LIMIT = 50
            pending_display_df = pending_df.head(PENDING_DISPLAY_LIMIT)
            if len(pending_df) > PENDING_DISPLAY_LIMIT:
                st.info(
                    f"Showing the {PENDING_DISPLAY_LIMIT} most recent of {len(pending_df)} pending "
                    "requests for speed. Approve/reject these first, or use Bulk Assign above, "
                    "to bring the rest into view."
                )
            for _, r in pending_display_df.iterrows():
                with st.expander(f"🟡 Requisition {short_req_id(r.get('id'))} — {r['applicant_name']} ({r['department']}) → {r['destination']}"):
                    st.caption(f"Technical ID: `{r['request_id']}`")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.write(f"**Mobile:** {r['mobile_number']}")
                        st.write(f"**Date/Time:** {r['date_of_travel']} at {fmt_time_12h(r['time_of_travel'])}")
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
                        approved_time = time_input_12h(
                            f"Approved Departure Time (originally requested {fmt_time_12h(r['time_of_travel'])})",
                            key_prefix=f"atime_{r['request_id']}",
                            default_time=default_time,
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
                                st.success(f"Requisition {short_req_id(r.get('id'))} marked as {new_status}.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Update failed: {e}")

    # ---------------- Create Requisition (Admin) — NEW ----------------
    with tab_create_req:
        st.subheader("➕ Create a Requisition Directly (Admin)")
        st.caption(
            "Use this when you need to arrange a vehicle for someone yourself — a walk-in request, "
            "a phone call, or any case that didn't come through the normal Employee 'New "
            "Requisition' form. This creates the requisition **already Approved**, with the Driver "
            "and Vehicle assigned in the same step — it skips the Pending queue entirely."
        )

        create_drivers_df = fetch_all_drivers()
        create_vehicles_df = fetch_all_vehicles()
        create_driver_contact_map = dict(zip(create_drivers_df["driver_name"], create_drivers_df["driver_contact"])) if not create_drivers_df.empty else {}
        create_driver_options = ["— Select Driver —"] + create_drivers_df["driver_name"].tolist() if not create_drivers_df.empty else []
        create_vehicle_options = ["— Select Vehicle —"] + create_vehicles_df["vehicle_number"].tolist() if not create_vehicles_df.empty else []

        if create_drivers_df.empty or create_vehicles_df.empty:
            st.warning(
                "⚠️ No drivers and/or vehicles are registered yet. Add them under the "
                "**🚘 Manage Drivers & Vehicles** tab before creating a requisition here."
            )

        # Optional: pick an existing Approved user to auto-fill their profile
        # details. Kept OUTSIDE the form (same pattern as the driver/vehicle
        # selects in Pending Requests) so choosing someone here immediately
        # updates the defaults shown inside the form below, instead of only
        # taking effect after the form is submitted.
        existing_user_options = ["— Manual Entry —"] + (
            users_df[users_df["status"] == "Approved"]["username"].tolist() if not users_df.empty else []
        )
        selected_existing_user = st.selectbox(
            "Fill from an existing user (optional)", existing_user_options, key="create_req_existing_user",
        )
        prefill = {}
        if selected_existing_user != "— Manual Entry —" and not users_df.empty:
            match = users_df[users_df["username"] == selected_existing_user]
            if not match.empty:
                prefill = match.iloc[0].to_dict()

        with st.form("admin_create_req_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                ca_applicant_name = st.text_input("Applicant Name *", value=fmt(prefill.get("full_name"), ""))
                ca_department = st.selectbox(
                    "Department *", DEPARTMENTS,
                    index=DEPARTMENTS.index(prefill.get("department")) if prefill.get("department") in DEPARTMENTS else 0,
                    key="ca_department",
                )
                ca_mobile = st.text_input("Mobile Number *", value=fmt(prefill.get("mobile"), ""),
                                           placeholder="01XXXXXXXXX")
                ca_passenger_count = st.number_input("Passenger Count *", min_value=1, max_value=50, value=1,
                                                      key="ca_passenger_count")
            with c2:
                ca_date = st.date_input("Date of Travel *", min_value=date.today(), key="ca_date")
                ca_time = time_input_12h("Time of Travel *", key_prefix="create_req_tt")
                ca_destination = st.text_input("Destination *", key="ca_destination")
                ca_vehicle_type = st.selectbox("Vehicle Type Required *", VEHICLE_TYPES, key="ca_vehicle_type")

            ca_purpose = st.text_area("Purpose of Travel *", height=90, key="ca_purpose")
            ca_special_request = st.text_area("Special Request (optional)", height=70, key="ca_special_request")

            st.markdown("---")
            st.markdown("##### 🚗 Assign Driver & Vehicle (required to create)")
            d1, d2 = st.columns(2)
            with d1:
                ca_driver = st.selectbox(
                    "Driver Name *", create_driver_options or ["No drivers available"],
                    disabled=not create_driver_options, key="ca_driver",
                )
            with d2:
                ca_vehicle = st.selectbox(
                    "Vehicle Number *", create_vehicle_options or ["No vehicles available"],
                    disabled=not create_vehicle_options, key="ca_vehicle",
                )
            ca_admin_note = st.text_area("Admin Note (optional)", height=60, key="ca_admin_note")

            ca_submitted = st.form_submit_button("✅ Create & Approve Requisition", type="primary",
                                                  use_container_width=True)

        if ca_submitted:
            errors = []
            if not ca_applicant_name.strip():
                errors.append("Applicant Name is required.")
            if not ca_mobile.strip():
                errors.append("Mobile Number is required.")
            if not ca_destination.strip():
                errors.append("Destination is required.")
            if not ca_purpose.strip():
                errors.append("Purpose of Travel is required.")
            driver_ready = bool(create_driver_options) and ca_driver != "— Select Driver —"
            vehicle_ready = bool(create_vehicle_options) and ca_vehicle != "— Select Vehicle —"
            if not driver_ready:
                errors.append("Please select a Driver.")
            if not vehicle_ready:
                errors.append("Please select a Vehicle.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                request_id = generate_request_id()
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                data = {
                    "request_id": request_id,
                    "username": prefill.get("username", ""),
                    "applicant_name": ca_applicant_name.strip(),
                    "department": ca_department,
                    "mobile_number": ca_mobile.strip(),
                    "date_of_travel": str(ca_date),
                    "time_of_travel": ca_time.strftime("%H:%M"),
                    "destination": ca_destination.strip(),
                    "passenger_count": int(ca_passenger_count),
                    "vehicle_type": ca_vehicle_type,
                    "purpose": ca_purpose.strip(),
                    "special_request": ca_special_request.strip(),
                    "status": "Approved",
                    "driver_name": ca_driver,
                    "driver_contact": create_driver_contact_map.get(ca_driver, ""),
                    "vehicle_number": ca_vehicle,
                    "approved_by": user["full_name"],
                    "action_timestamp": now_str,
                    "approved_time": ca_time.strftime("%H:%M"),
                    "admin_note": ca_admin_note.strip(),
                }
                with st.spinner("Saving to Supabase..."):
                    try:
                        new_id = insert_requisition(data)
                        st.success(
                            f"✅ Requisition **{short_req_id(new_id)}** created and already Approved — "
                            f"Driver **{ca_driver}** / Vehicle **{ca_vehicle}**."
                        )
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Failed to save requisition: {e}")

    # ---------------- Staff Shuttle (Recurring Routes) — NEW ----------------
    with tab_shuttle:
        st.subheader("🚌 Staff Shuttle — Quick Submit")
        st.caption(
            "For fixed, recurring HIACE routes (e.g. the morning staff pickup, the evening "
            "7:15 PM drop-off) — save each route once as a template below, then submit "
            "tomorrow's trip with a single click instead of retyping everything every day."
        )

        shuttle_drivers_df = fetch_all_drivers()
        shuttle_vehicles_df = fetch_all_vehicles()
        shuttle_driver_options = shuttle_drivers_df["driver_name"].tolist() if not shuttle_drivers_df.empty else []
        shuttle_vehicle_options = shuttle_vehicles_df["vehicle_number"].tolist() if not shuttle_vehicles_df.empty else []
        shuttle_driver_contact_map = (
            dict(zip(shuttle_drivers_df["driver_name"], shuttle_drivers_df["driver_contact"]))
            if not shuttle_drivers_df.empty else {}
        )

        if not shuttle_driver_options or not shuttle_vehicle_options:
            st.warning(
                "⚠️ No drivers and/or vehicles are registered yet. Add them under the "
                "**🚘 Manage Drivers & Vehicles** tab before submitting a shuttle trip."
            )

        templates_df = fetch_all_shuttle_templates()

        # "Tomorrow" is recomputed fresh every time this tab renders (today's
        # real date + 1 day) — this is what makes the one-click button always
        # correct without anyone having to touch a date picker: run this
        # tonight (03/09/26) and it submits for 04/09/26; run it any other
        # night and it automatically submits for the following day.
        tomorrow_date = date.today() + timedelta(days=1)

        def _submit_shuttle_trip(tpl_row, trip_date, driver_name, vehicle_number):
            """Creates one Approved requisition from a shuttle template for
            the given date/driver/vehicle. Returns the new short req id, or
            raises on failure (caller handles the try/except + message)."""
            request_id = generate_request_id()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            default_time_val = fmt(tpl_row.get("default_time"), "08:00")
            data = {
                "request_id": request_id,
                "username": "",
                "applicant_name": fmt(tpl_row.get("applicant_name"), tpl_row["template_name"]),
                "department": fmt(tpl_row.get("department"), "Admin"),
                "mobile_number": "",
                "date_of_travel": str(trip_date),
                "time_of_travel": default_time_val,
                "destination": fmt(tpl_row.get("destination"), ""),
                "passenger_count": int(tpl_row.get("passenger_count") or 1),
                "vehicle_type": fmt(tpl_row.get("vehicle_type"), "HIACE"),
                "purpose": fmt(tpl_row.get("purpose"), tpl_row["template_name"]),
                "special_request": "",
                "status": "Approved",
                "driver_name": driver_name,
                "driver_contact": shuttle_driver_contact_map.get(driver_name, ""),
                "vehicle_number": vehicle_number,
                "approved_by": user["full_name"],
                "action_timestamp": now_str,
                "approved_time": default_time_val,
                "admin_note": f"Auto-submitted from shuttle template: {tpl_row['template_name']}",
            }
            new_id = insert_requisition(data)
            return short_req_id(new_id)

        st.markdown("##### 🚀 Quick Submit")
        if templates_df.empty:
            st.info("No shuttle templates saved yet — add one below under 'Manage Templates'.")
        else:
            # ---- ONE button for ALL fixed morning trips at once ----
            # This is the "4 trips, one click" button: every saved template
            # is submitted in a single go for tomorrow's date, each using
            # its own saved default time/driver/vehicle — nothing to pick.
            bulk_ready = bool(shuttle_driver_options) and bool(shuttle_vehicle_options)
            st.info(
                f"📅 **Tomorrow's date:** {tomorrow_date.strftime('%A, %d %B %Y')} "
                f"({tomorrow_date.strftime('%d/%m/%y')})"
            )
            if st.button(
                f"🌅 Submit ALL {len(templates_df)} Fixed Trips for Tomorrow ({tomorrow_date.strftime('%d/%m/%y')})",
                type="primary", use_container_width=True, disabled=not bulk_ready,
                key="shuttle_bulk_submit_all",
            ):
                created, failed = [], []
                with st.spinner("Saving all fixed trips to Supabase..."):
                    for _, tpl in templates_df.iterrows():
                        d_name = tpl.get("default_driver_name", "")
                        v_number = tpl.get("default_vehicle_number", "")
                        if d_name not in shuttle_driver_options or v_number not in shuttle_vehicle_options:
                            failed.append(f"{tpl['template_name']} (no default Driver/Vehicle saved)")
                            continue
                        try:
                            new_short_id = _submit_shuttle_trip(tpl, tomorrow_date, d_name, v_number)
                            created.append(f"{tpl['template_name']} → {new_short_id}")
                        except Exception as e:
                            failed.append(f"{tpl['template_name']} ({e})")
                if created:
                    st.success("✅ Created:\n\n" + "\n".join(f"- {c}" for c in created))
                if failed:
                    st.error("⚠️ Skipped:\n\n" + "\n".join(f"- {f}" for f in failed))
                if created:
                    st.balloons()
                    st.rerun()
            if not bulk_ready:
                st.caption("⚠️ Add at least one Driver and Vehicle before using the bulk button above.")

            st.markdown("---")
            st.caption("Or submit / adjust one trip at a time:")

            for _, tpl in templates_df.iterrows():
                tpl_id = tpl["id"]
                default_driver = tpl.get("default_driver_name", "")
                default_vehicle = tpl.get("default_vehicle_number", "")
                driver_ok = default_driver in shuttle_driver_options
                vehicle_ok = default_vehicle in shuttle_vehicle_options

                with st.container():
                    st.markdown(
                        f"**🚐 {tpl['template_name']}** — {fmt(tpl.get('destination'))} "
                        f"@ {fmt_time_12h(tpl.get('default_time'))} "
                        f"| Driver: {fmt(default_driver, '— none saved —')} "
                        f"| Vehicle: {fmt(default_vehicle, '— none saved —')}"
                    )
                    single_ready = driver_ok and vehicle_ok
                    if st.button(
                        f"✅ Submit for Tomorrow ({tomorrow_date.strftime('%d/%m/%y')}, {fmt_time_12h(tpl.get('default_time'))})",
                        key=f"shuttle_submit_tomorrow_{tpl_id}", type="primary",
                        use_container_width=True, disabled=not single_ready,
                    ):
                        with st.spinner("Saving to Supabase..."):
                            try:
                                new_short_id = _submit_shuttle_trip(tpl, tomorrow_date, default_driver, default_vehicle)
                                st.success(
                                    f"✅ **{tpl['template_name']}** requisition **{new_short_id}** "
                                    f"created for {tomorrow_date.strftime('%d/%m/%y')} — "
                                    f"Driver **{default_driver}**, Vehicle **{default_vehicle}**."
                                )
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Failed to save requisition: {e}")
                    if not single_ready:
                        st.caption(
                            "⚠️ This template has no valid saved default Driver/Vehicle — "
                            "edit it below under 'Manage Templates' first, or use the manual form below."
                        )

                    with st.expander("✏️ Change date / driver / vehicle for this one time"):
                        st.write(f"**Purpose:** {fmt(tpl.get('purpose'))}")
                        st.write(
                            f"**Vehicle Type:** {fmt(tpl.get('vehicle_type'))}  |  "
                            f"**Passengers:** {fmt(tpl.get('passenger_count'))}"
                        )

                        qc1, qc2, qc3 = st.columns(3)
                        with qc1:
                            q_date = st.date_input("Trip Date", value=tomorrow_date, key=f"shuttle_date_{tpl_id}")
                        with qc2:
                            driver_choices_this = shuttle_driver_options or ["No drivers available"]
                            q_driver = st.selectbox(
                                "Driver", driver_choices_this,
                                index=driver_choices_this.index(default_driver) if default_driver in driver_choices_this else 0,
                                key=f"shuttle_driver_{tpl_id}", disabled=not shuttle_driver_options,
                            )
                        with qc3:
                            vehicle_choices_this = shuttle_vehicle_options or ["No vehicles available"]
                            q_vehicle = st.selectbox(
                                "Vehicle", vehicle_choices_this,
                                index=vehicle_choices_this.index(default_vehicle) if default_vehicle in vehicle_choices_this else 0,
                                key=f"shuttle_vehicle_{tpl_id}", disabled=not shuttle_vehicle_options,
                            )

                        if st.button(
                            f"✅ Submit — {tpl['template_name']} for {q_date}",
                            key=f"shuttle_submit_{tpl_id}", use_container_width=True,
                        ):
                            if not shuttle_driver_options or not shuttle_vehicle_options:
                                st.error("Please add at least one Driver and Vehicle first.")
                            else:
                                with st.spinner("Saving to Supabase..."):
                                    try:
                                        new_short_id = _submit_shuttle_trip(tpl, q_date, q_driver, q_vehicle)
                                        st.success(
                                            f"✅ **{tpl['template_name']}** requisition **{new_short_id}** "
                                            f"created for {q_date} — Driver **{q_driver}**, Vehicle **{q_vehicle}**."
                                        )
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"❌ Failed to save requisition: {e}")
                    st.markdown("---")

        st.markdown("---")
        st.markdown("##### ⚙️ Manage Templates")
        with st.form("add_shuttle_template_form", clear_on_submit=True):
            t1, t2 = st.columns(2)
            with t1:
                new_tpl_name = st.text_input(
                    "Template Name *", placeholder="e.g. Evening Staff Shuttle (7:15 PM)"
                )
                new_tpl_destination = st.text_input(
                    "Destination / Route *", placeholder="e.g. Staff drop-off points, Ishwardi route"
                )
                new_tpl_vehicle_type = st.selectbox(
                    "Vehicle Type", VEHICLE_TYPES,
                    index=VEHICLE_TYPES.index("HIACE") if "HIACE" in VEHICLE_TYPES else 0,
                    key="new_tpl_vehicle_type",
                )
                new_tpl_passengers = st.number_input(
                    "Passenger Count", min_value=1, max_value=50, value=10, key="new_tpl_passengers"
                )
            with t2:
                new_tpl_purpose = st.text_input(
                    "Purpose *", placeholder="e.g. Staff Shuttle - Evening Drop-off"
                )
                new_tpl_time = time_input_12h("Default Departure Time *", key_prefix="new_shuttle_time")
                new_tpl_driver = st.selectbox(
                    "Default Driver", ["— None —"] + shuttle_driver_options, key="new_tpl_driver"
                )
                new_tpl_vehicle = st.selectbox(
                    "Default Vehicle", ["— None —"] + shuttle_vehicle_options, key="new_tpl_vehicle"
                )
            add_tpl_clicked = st.form_submit_button("➕ Save Template", type="primary", use_container_width=True)

        if add_tpl_clicked:
            tpl_errors = []
            if not new_tpl_name.strip():
                tpl_errors.append("Template Name is required.")
            if not new_tpl_destination.strip():
                tpl_errors.append("Destination / Route is required.")
            if not new_tpl_purpose.strip():
                tpl_errors.append("Purpose is required.")
            if tpl_errors:
                for e in tpl_errors:
                    st.error(e)
            else:
                try:
                    add_shuttle_template({
                        "template_name": new_tpl_name.strip(),
                        "applicant_name": new_tpl_name.strip(),
                        "department": "Admin",
                        "destination": new_tpl_destination.strip(),
                        "purpose": new_tpl_purpose.strip(),
                        "vehicle_type": new_tpl_vehicle_type,
                        "passenger_count": int(new_tpl_passengers),
                        "default_time": new_tpl_time.strftime("%H:%M"),
                        "default_driver_name": "" if new_tpl_driver == "— None —" else new_tpl_driver,
                        "default_vehicle_number": "" if new_tpl_vehicle == "— None —" else new_tpl_vehicle,
                    })
                    st.success(f"✅ Template '{new_tpl_name.strip()}' saved.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Failed to save template: {e}")

        st.markdown("###### Current Templates")
        if templates_df.empty:
            st.info("No templates yet.")
        else:
            for _, tpl in templates_df.iterrows():
                r1, r2 = st.columns([5, 1])
                r1.write(
                    f"**{tpl['template_name']}** — {fmt(tpl.get('destination'))} "
                    f"@ {fmt_time_12h(tpl.get('default_time'))}"
                )
                if r2.button("🗑️", key=f"del_shuttle_tpl_{tpl['id']}", help="Delete this template"):
                    try:
                        delete_shuttle_template(tpl["id"])
                        st.success(f"Deleted template '{tpl['template_name']}'.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Failed to delete template: {e}")

    # ---------------- User List (was Tab 2) ----------------
    with tab_users:
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

            st.markdown("##### Change a user's role, status, or profile info")
            approved_usernames = users_df["username"].tolist()
            sel_user = st.selectbox("Select username", approved_usernames)
            if sel_user:
                rec = users_df[users_df["username"] == sel_user].iloc[0]
                c1, c2 = st.columns(2)
                with c1:
                    current_role = rec.get("role", "user")
                    new_role = st.selectbox("Role", ROLE_OPTIONS,
                                             index=ROLE_OPTIONS.index(current_role) if current_role in ROLE_OPTIONS else 0,
                                             format_func=lambda r: ROLE_DISPLAY.get(r, r))
                with c2:
                    new_status = st.selectbox("Status", USER_STATUS_OPTIONS,
                                               index=USER_STATUS_OPTIONS.index(rec.get("status", "Pending")) if rec.get("status") in USER_STATUS_OPTIONS else 0)

                # Full Name and Department are editable too — this is what
                # actually fixes a mistake like an employee registering under
                # the wrong department (their New Requisition form auto-fills
                # from this exact field, so a wrong value here keeps
                # suggesting the wrong department every time they submit).
                c3, c4 = st.columns(2)
                with c3:
                    new_full_name = st.text_input("Full Name", value=fmt(rec.get("full_name"), ""))
                with c4:
                    current_dept = rec.get("department", "")
                    dept_options = DEPARTMENTS if current_dept in DEPARTMENTS else DEPARTMENTS + (
                        [current_dept] if current_dept else []
                    )
                    new_department = st.selectbox(
                        "Department", dept_options,
                        index=dept_options.index(current_dept) if current_dept in dept_options else 0,
                    )

                if st.button("💾 Save Changes", use_container_width=True):
                    update_user(sel_user, {
                        "role": new_role,
                        "status": new_status,
                        "full_name": new_full_name.strip(),
                        "department": new_department,
                    })
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

    # ---------------- ID Requests (was Tab 1) ----------------
    with tab_pending_users:
        st.subheader("New Account Requests")
        pending_users = users_df[users_df["status"] == "Pending"] if not users_df.empty else users_df

        if pending_users.empty:
            st.success("🎉 No pending user registrations.")
        else:
            for _, u in pending_users.iterrows():
                role_req = u.get("role")
                if role_req == "gate_officer":
                    role_tag = "🛡️ Gate Officer"
                elif role_req == "driver":
                    role_tag = "🚙 Driver"
                else:
                    role_tag = "👤 Employee"
                with st.expander(f"{role_tag} — {u['full_name']} (@{u['username']})"):
                    if role_req in ("gate_officer", "driver"):
                        st.write(f"**Role Requested:** {ROLE_DISPLAY.get(role_req, role_req)}")
                    else:
                        st.write(f"**Designation:** {u.get('designation', '')}")
                        st.write(f"**Employee ID:** {u.get('employee_id', '')}")
                        st.write(f"**Department:** {u.get('department', '')}")
                        st.write(f"**Mobile:** {u.get('mobile', '')}")
                    st.write(f"**Requested on:** {fmt_time_12h(u.get('created_at', ''), u.get('created_at', ''))}")
                    c1, c2 = st.columns(2)
                    if c1.button("✅ Approve", key=f"appr_{u['username']}", type="primary", use_container_width=True):
                        update_user(u["username"], {"status": "Approved"})
                        st.success(f"{u['full_name']} approved.")
                        st.rerun()
                    if c2.button("❌ Reject", key=f"rej_{u['username']}", use_container_width=True):
                        update_user(u["username"], {"status": "Rejected"})
                        st.warning(f"{u['full_name']} rejected.")
                        st.rerun()

    # ---------------- Analytics (was Tab 4) ----------------
    with tab_analytics:
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
                st.plotly_chart(fig1, use_container_width=True, key="analytics_dept_chart")
            with c2:
                status_counts = df_all["status"].value_counts().reset_index()
                status_counts.columns = ["Status", "Count"]
                fig2 = px.pie(status_counts, names="Status", values="Count", hole=0.45, title="Status Breakdown",
                              color="Status",
                              color_discrete_map={"Approved": "#28a745", "Rejected": "#dc3545", "Pending": "#ffc107",
                                                   "On Trip": "#0d6efd", "Completed": "#17a673"})
                fig2.update_layout(height=380)
                st.plotly_chart(fig2, use_container_width=True, key="analytics_status_pie")

            # Driver-verified KM (falls back to the Gate Officer's
            # Gate-In/Out readings only when a driver hasn't logged their
            # own numbers) is now the primary source for this metric, per
            # business requirement — the metric itself is unchanged, only
            # its source.
            completed_trips = df_all[df_all["status"] == "Completed"].copy()
            if not completed_trips.empty:
                completed_trips["_eff_km"] = completed_trips.apply(
                    lambda row: effective_km_fields(row)[2], axis=1
                )
                total_km_all = float(completed_trips["_eff_km"].sum())
                st.metric("🛣️ Total KM Covered (Completed Trips — Driver-verified)", f"{total_km_all:.1f} KM")

            df_all["_dt"] = pd.to_datetime(df_all["date_of_travel"], errors="coerce")
            monthly = df_all.dropna(subset=["_dt"]).copy()
            if not monthly.empty:
                monthly["Month"] = monthly["_dt"].dt.to_period("M").astype(str)
                monthly_counts = monthly.groupby("Month").size().reset_index(name="Requests")
                fig3 = px.line(monthly_counts, x="Month", y="Requests", markers=True, title="Monthly Request Trend")
                fig3.update_layout(height=350)
                st.plotly_chart(fig3, use_container_width=True, key="analytics_monthly_trend")

    # ---------------- All Requisitions & Export (was Tab 5) ----------------
    with tab_export:
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

            filtered_display = filtered.drop(columns=["_dt"], errors="ignore").copy()
            if "time_of_travel" in filtered_display.columns:
                filtered_display["time_of_travel"] = filtered_display["time_of_travel"].apply(
                    lambda v: fmt_time_12h(v, v)
                )
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

    # ---------------- Edit / Delete Trip (NEW) ----------------
    with tab_edit_trip:
        st.subheader("✏️ Edit or Delete a Requisition")
        st.caption(
            "Use this to fix a mistaken entry — e.g. a driver typed the wrong odometer "
            "reading, a wrong destination/time was saved, or a duplicate/test request needs "
            "to be removed entirely. Changes here go straight to the database."
        )

        if df_all.empty:
            st.info("No requisitions yet.")
        else:
            # -----------------------------------------------------------
            # BULK DELETE (NEW) — for clearing out MANY stuck/duplicate/
            # test requisitions at once. Deliberately searches driver_name
            # by free-text PARTIAL match against the raw requisition data,
            # rather than the "1️⃣ Select Driver" dropdown further below
            # (which only lists names from the Manage Drivers & Vehicles
            # master list via exact normalized matching). This matters
            # because a trip's driver_name is free-typed at the time it
            # was approved/created — if it was ever typed even slightly
            # differently from that driver's master-list entry, the
            # dropdown-based flow below won't find it even though the
            # trip is still sitting in the database. This search-based
            # box finds it regardless, so nothing stays permanently
            # invisible to Admin just because of a spelling mismatch.
            # -----------------------------------------------------------
            with st.expander("🧹 Bulk Delete Multiple Trips (search-based, catches name-mismatched trips too)"):
                st.caption(
                    "Search by driver name (partial match — this also finds trips whose driver_name "
                    "was typed slightly differently than in Manage Drivers & Vehicles) and/or filter "
                    "by status, then select as many as you need and delete them all together. "
                    "This permanently removes them — it cannot be undone."
                )
                bc1, bc2 = st.columns(2)
                with bc1:
                    bulk_del_driver_search = st.text_input(
                        "Driver name contains", key="bulk_del_driver_search",
                        placeholder="e.g. Manik",
                    )
                with bc2:
                    bulk_del_status_filter = st.multiselect(
                        "Status (optional)", REQ_STATUS_OPTIONS, key="bulk_del_status_filter",
                    )

                bulk_del_scope = df_all.copy()
                if bulk_del_driver_search.strip():
                    bulk_del_scope = bulk_del_scope[
                        bulk_del_scope["driver_name"].fillna("").astype(str).str.contains(
                            bulk_del_driver_search.strip(), case=False, na=False
                        )
                    ]
                if bulk_del_status_filter:
                    bulk_del_scope = bulk_del_scope[bulk_del_scope["status"].isin(bulk_del_status_filter)]

                if not bulk_del_driver_search.strip() and not bulk_del_status_filter:
                    st.caption("Type a driver name or pick a status above to see matching trips here.")
                elif bulk_del_scope.empty:
                    st.info("No requisitions match this search.")
                else:
                    bulk_del_option_map = {
                        f"{short_req_id(r.get('id'))} — Driver: {fmt(r.get('driver_name'), '(none)')} — "
                        f"{r['applicant_name']} → {r['destination']} "
                        f"({r['status']}, {r['date_of_travel']})": r["request_id"]
                        for _, r in bulk_del_scope.iterrows()
                    }
                    st.write(f"**{len(bulk_del_option_map)} matching requisition(s) found.**")
                    bulk_del_selected = st.multiselect(
                        "Select requisitions to permanently delete",
                        list(bulk_del_option_map.keys()),
                        key="bulk_del_selected",
                    )

                    if bulk_del_selected:
                        st.warning(
                            f"⚠️ You are about to permanently delete **{len(bulk_del_selected)}** "
                            "requisition(s). This cannot be undone."
                        )
                        confirm_bulk_del = st.checkbox(
                            f"I understand this will permanently delete {len(bulk_del_selected)} requisition(s).",
                            key="confirm_bulk_del",
                        )
                        if st.button(
                            f"🗑️ Delete {len(bulk_del_selected)} Selected Requisitions",
                            type="primary", disabled=not confirm_bulk_del,
                            use_container_width=True, key="bulk_del_submit",
                        ):
                            deleted, failed = [], []
                            with st.spinner("Deleting selected requisitions..."):
                                for label in bulk_del_selected:
                                    rid = bulk_del_option_map[label]
                                    try:
                                        delete_requisition(rid)
                                        deleted.append(label)
                                    except Exception as e:
                                        failed.append(f"{label} ({e})")
                            if deleted:
                                st.success(f"✅ Deleted {len(deleted)} requisition(s).")
                            if failed:
                                st.error("⚠️ Failed to delete:\n\n" + "\n".join(f"- {f}" for f in failed))
                            if deleted:
                                st.rerun()

            st.markdown("---")

            edit_drivers_df = fetch_all_drivers()
            edit_vehicles_df = fetch_all_vehicles()
            edit_driver_contact_map = (
                dict(zip(edit_drivers_df["driver_name"], edit_drivers_df["driver_contact"]))
                if not edit_drivers_df.empty else {}
            )
            edit_driver_choices = edit_drivers_df["driver_name"].tolist() if not edit_drivers_df.empty else []

            # ---- STEP 1: pick a Driver first ----
            # "Approved" trips are shown here as "🟡 Pending / Upcoming" from
            # the driver's own perspective — assigned to them but not yet
            # Gated Out. "All / No Driver Assigned" keeps the old behaviour
            # of browsing every requisition (e.g. still-Pending requests
            # that haven't been assigned a driver yet, or Rejected ones).
            driver_filter_choices = ["— All / No Driver Assigned —"] + edit_driver_choices
            selected_edit_driver = st.selectbox(
                "1️⃣ Select Driver", driver_filter_choices, key="edit_trip_driver_filter",
            )

            if selected_edit_driver == "— All / No Driver Assigned —":
                driver_scoped_df = df_all
            else:
                target_key = _normalize_driver_name(selected_edit_driver)
                driver_scoped_df = df_all[df_all["driver_name"].map(_normalize_driver_name) == target_key]

            if driver_scoped_df.empty:
                st.warning(
                    f"No requisitions found for **{selected_edit_driver}** — showing all "
                    "requisitions instead."
                )
                driver_scoped_df = df_all

            # ---- STEP 2: pick a trip category for that driver ----
            status_groups = {
                "🟡 Pending / Upcoming (Approved, not yet started)": "Approved",
                "🔵 On Trip": "On Trip",
                "✅ Completed": "Completed",
                "⏳ Still Pending (no driver assigned yet)": "Pending",
                "🔴 Rejected": "Rejected",
            }
            # Only offer categories that actually have at least one matching
            # row, so the dropdown doesn't show empty groups.
            available_groups = {
                label: status_val for label, status_val in status_groups.items()
                if not driver_scoped_df[driver_scoped_df["status"] == status_val].empty
            }
            if not available_groups:
                st.warning("No categorized trips found for this selection — showing all statuses instead.")
                available_groups = {
                    STATUS_BADGE.get(s, s): s
                    for s in sorted(driver_scoped_df["status"].dropna().unique().tolist())
                }

            selected_group_label = st.selectbox(
                "2️⃣ Select Trip Category", list(available_groups.keys()), key="edit_trip_status_filter",
            )
            category_df = driver_scoped_df[driver_scoped_df["status"] == available_groups[selected_group_label]]

            # ---- STEP 3: pick the specific requisition to edit ----
            edit_options = {
                f"{short_req_id(r.get('id'))} — {r['applicant_name']} → {r['destination']} ({r['date_of_travel']})": r["request_id"]
                for _, r in category_df.iterrows()
            }
            selected_label = st.selectbox(
                "3️⃣ Select Requisition to Edit", list(edit_options.keys()), key="edit_trip_select"
            )
            selected_request_id = edit_options[selected_label]
            row = df_all[df_all["request_id"] == selected_request_id].iloc[0]
            selected_short_id = short_req_id(row.get("id"))
            st.caption(f"Technical ID: `{selected_request_id}`")
            st.markdown("---")

            edit_driver_choices = edit_drivers_df["driver_name"].tolist() if not edit_drivers_df.empty else []
            edit_vehicle_choices = edit_vehicles_df["vehicle_number"].tolist() if not edit_vehicles_df.empty else []
            # Always keep the row's CURRENT driver/vehicle selectable even if
            # it's since been removed from Manage Drivers & Vehicles, so
            # opening this form never silently wipes out a valid historical
            # assignment just because the master list changed later.
            current_driver = fmt(row.get("driver_name"), "")
            current_vehicle = fmt(row.get("vehicle_number"), "")
            if current_driver and current_driver not in edit_driver_choices:
                edit_driver_choices = [current_driver] + edit_driver_choices
            if current_vehicle and current_vehicle not in edit_vehicle_choices:
                edit_vehicle_choices = [current_vehicle] + edit_vehicle_choices
            edit_driver_choices_display = ["— None —"] + edit_driver_choices
            edit_vehicle_choices_display = ["— None —"] + edit_vehicle_choices

            with st.form("edit_trip_form"):
                st.markdown("##### Trip Details")
                e1, e2 = st.columns(2)
                with e1:
                    et_applicant_name = st.text_input("Applicant Name", value=fmt(row.get("applicant_name"), ""))
                    et_department = st.selectbox(
                        "Department", DEPARTMENTS,
                        index=DEPARTMENTS.index(row.get("department")) if row.get("department") in DEPARTMENTS else 0,
                    )
                    et_mobile = st.text_input("Mobile Number", value=fmt(row.get("mobile_number"), ""))
                    et_passenger_count = st.number_input(
                        "Passenger Count", min_value=1, max_value=50,
                        value=int(row.get("passenger_count") or 1),
                    )
                    et_status = st.selectbox(
                        "Status", REQ_STATUS_OPTIONS,
                        index=REQ_STATUS_OPTIONS.index(row.get("status")) if row.get("status") in REQ_STATUS_OPTIONS else 0,
                    )
                with e2:
                    try:
                        et_date_default = datetime.strptime(str(row.get("date_of_travel")), "%Y-%m-%d").date()
                    except (ValueError, TypeError):
                        et_date_default = date.today()
                    et_date = st.date_input("Date of Travel", value=et_date_default)
                    try:
                        et_time_default = datetime.strptime(fmt(row.get("time_of_travel"), "09:00"), "%H:%M").time()
                    except ValueError:
                        et_time_default = datetime.now().time()
                    et_time = time_input_12h("Time of Travel", key_prefix="edit_trip_tt", default_time=et_time_default)
                    et_destination = st.text_input("Destination", value=fmt(row.get("destination"), ""))
                    et_vehicle_type = st.selectbox(
                        "Vehicle Type", VEHICLE_TYPES,
                        index=VEHICLE_TYPES.index(row.get("vehicle_type")) if row.get("vehicle_type") in VEHICLE_TYPES else 0,
                    )

                et_purpose = st.text_area("Purpose", value=fmt(row.get("purpose"), ""), height=80)
                et_special_request = st.text_area("Special Request", value=fmt(row.get("special_request"), ""), height=60)

                st.markdown("---")
                st.markdown("##### Driver & Vehicle Assignment")
                d1, d2 = st.columns(2)
                with d1:
                    et_driver = st.selectbox(
                        "Driver", edit_driver_choices_display,
                        index=edit_driver_choices_display.index(current_driver) if current_driver in edit_driver_choices_display else 0,
                    )
                with d2:
                    et_vehicle = st.selectbox(
                        "Vehicle Number", edit_vehicle_choices_display,
                        index=edit_vehicle_choices_display.index(current_vehicle) if current_vehicle in edit_vehicle_choices_display else 0,
                    )

                st.markdown("---")
                st.markdown("##### Odometer / KM Corrections")
                st.caption(
                    "Tick 'leave blank' to clear a value that hasn't actually been recorded, "
                    "instead of leaving a stray 0 that would throw off distance reports."
                )
                k1, k2 = st.columns(2)
                with k1:
                    st.markdown("**Gate Officer's readings**")
                    sk_blank = st.checkbox("Start KM — leave blank", value=is_blank(row.get("start_km")),
                                            key="et_sk_blank")
                    et_start_km = st.number_input(
                        "Start KM", min_value=0.0, step=1.0, format="%.1f",
                        value=float(row.get("start_km")) if not is_blank(row.get("start_km")) else 0.0,
                        disabled=sk_blank, key="et_start_km",
                    )
                    ek_blank = st.checkbox("End KM — leave blank", value=is_blank(row.get("end_km")),
                                            key="et_ek_blank")
                    et_end_km = st.number_input(
                        "End KM", min_value=0.0, step=1.0, format="%.1f",
                        value=float(row.get("end_km")) if not is_blank(row.get("end_km")) else 0.0,
                        disabled=ek_blank, key="et_end_km",
                    )
                with k2:
                    st.markdown("**Driver's own readings**")
                    dsk_blank = st.checkbox("Driver Start KM — leave blank", value=is_blank(row.get("driver_start_km")),
                                             key="et_dsk_blank")
                    et_driver_start_km = st.number_input(
                        "Driver Start KM", min_value=0.0, step=1.0, format="%.1f",
                        value=float(row.get("driver_start_km")) if not is_blank(row.get("driver_start_km")) else 0.0,
                        disabled=dsk_blank, key="et_driver_start_km",
                    )
                    dek_blank = st.checkbox("Driver End KM — leave blank", value=is_blank(row.get("driver_end_km")),
                                             key="et_dek_blank")
                    et_driver_end_km = st.number_input(
                        "Driver End KM", min_value=0.0, step=1.0, format="%.1f",
                        value=float(row.get("driver_end_km")) if not is_blank(row.get("driver_end_km")) else 0.0,
                        disabled=dek_blank, key="et_driver_end_km",
                    )

                et_admin_note = st.text_area("Admin Note", value=fmt(row.get("admin_note"), ""), height=60)
                et_notify = st.checkbox(
                    "📢 Send a Telegram notification about this correction", value=False, key="et_notify"
                )

                et_save_clicked = st.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)

            if et_save_clicked:
                updates = {
                    "applicant_name": et_applicant_name.strip(),
                    "department": et_department,
                    "mobile_number": et_mobile.strip(),
                    "passenger_count": int(et_passenger_count),
                    "status": et_status,
                    "date_of_travel": str(et_date),
                    "time_of_travel": et_time.strftime("%H:%M"),
                    "destination": et_destination.strip(),
                    "vehicle_type": et_vehicle_type,
                    "purpose": et_purpose.strip(),
                    "special_request": et_special_request.strip(),
                    "driver_name": "" if et_driver == "— None —" else et_driver,
                    "driver_contact": "" if et_driver == "— None —" else edit_driver_contact_map.get(et_driver, ""),
                    "vehicle_number": "" if et_vehicle == "— None —" else et_vehicle,
                    "start_km": None if sk_blank else float(et_start_km),
                    "end_km": None if ek_blank else float(et_end_km),
                    "driver_start_km": None if dsk_blank else float(et_driver_start_km),
                    "driver_end_km": None if dek_blank else float(et_driver_end_km),
                    "admin_note": et_admin_note.strip(),
                }
                # Recompute total_km using the same driver-first-then-gate-officer
                # priority as effective_km_fields() everywhere else in the app,
                # so a manual correction here stays consistent with every
                # report/dashboard that reads total_km.
                updates["total_km"] = effective_km_fields(updates)[2]

                try:
                    update_requisition(selected_request_id, updates, notify=et_notify)
                    st.success(f"✅ Requisition {selected_short_id} updated.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Failed to update: {e}")

            st.markdown("---")
            st.markdown("##### 🗑️ Delete This Requisition Permanently")
            st.caption(
                f"This permanently removes **{selected_short_id}** from the system — it will "
                "disappear from every report, export, and dashboard. This cannot be undone."
            )
            confirm_delete_trip = st.checkbox(
                f"I understand this will permanently delete {selected_short_id}.",
                key=f"confirm_del_trip_{selected_request_id}",
            )
            if st.button("🗑️ Delete This Requisition", type="primary", disabled=not confirm_delete_trip,
                         use_container_width=True, key=f"del_trip_btn_{selected_request_id}"):
                try:
                    delete_requisition(selected_request_id)
                    st.success(f"{selected_short_id} has been deleted.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Failed to delete: {e}")

    # ---------------- Manage Drivers & Vehicles (was Tab 6) ----------------
    with tab_fleet:
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

    # ---------------- Duty Tracker & Analytics (was Tab 7) ----------------
    with tab_duty:
        st.subheader("🕒 Vehicle & Driver Duty Tracker & Analytics Dashboard")
        st.caption(
            "Filter any custom date/time window plus a specific vehicle or driver to see live duty "
            "duration, distance, and trip counts — with CSV / Excel / PDF export."
        )

        # ---- Reuse the same requisitions dataset already fetched above ----
        # `df_all` is now hoisted above every admin tab, so we don't hit
        # Supabase again here. Swap this for your own DataFrame if you wire
        # this tab up standalone.
        duty_df_raw = df_all.copy()

        if duty_df_raw.empty:
            st.info("No requisition data yet — the duty tracker will populate once trips are logged.")
        else:
            # -------------------------------------------------------------
            # STEP 0 — Render every filter widget FIRST, before any variable
            # derived from them is used. This avoids NameError from reading
            # a widget's value before Streamlit has actually created it.
            # -------------------------------------------------------------
            st.markdown("##### 🔎 Filters")
            fc1, fc2 = st.columns(2)
            with fc1:
                st.markdown("**Start of Range**")
                filter_start_date = st.date_input("Start Date", value=date.today(), key="duty_start_date")
                filter_start_time = st.time_input("Start Time", value=datetime.strptime("07:00", "%H:%M").time(),
                                                   key="duty_start_time")
            with fc2:
                st.markdown("**End of Range**")
                filter_end_date = st.date_input("End Date", value=date.today() + timedelta(days=1), key="duty_end_date")
                filter_end_time = st.time_input("End Time", value=datetime.strptime("06:59", "%H:%M").time(),
                                                 key="duty_end_time")

            # -------------------------------------------------------------
            # STEP 1 — Build real start/end datetimes for every trip, as
            # UTC-aware timestamps. Supabase/PostgREST returns timestamptz
            # columns as ISO 8601 strings — sometimes with an explicit
            # offset, sometimes without — so `utc=True` normalizes BOTH
            # cases onto a single tz-aware ("datetime64[us, UTC]" / similar)
            # dtype. Without this, mixed naive/aware values make pandas
            # infer a naive dtype for one column and an aware dtype for
            # another, and mixing the two later raises exactly the
            # "Invalid comparison between dtype=datetime64[...,UTC] and
            # datetime" TypeError.
            # -------------------------------------------------------------
            # `format="mixed"` matters as much as `utc=True` here: Supabase/
            # PostgREST timestamptz values can come back with or without
            # fractional seconds, or with a trailing "Z" vs "+00:00" offset,
            # row to row. Without `format="mixed"`, pandas infers a single
            # format from the first non-null value and silently coerces every
            # differently-shaped row to NaT (even with errors="coerce") —
            # which then makes real trips vanish from the tracker with no
            # error at all. "mixed" parses each value independently.
            #
            # `.astype("datetime64[ns, UTC]")` right after parsing pins both
            # columns to a FIXED, known time-unit resolution. Without this,
            # pandas infers the unit from whatever data is present: a column
            # that is entirely blank (e.g. brand-new "Pending" requisitions
            # that have never been Gated Out — exactly what you get right
            # after a fresh data reset) comes back as second-resolution
            # ("datetime64[s, UTC]"), while pd.Timestamp.now() below defaults
            # to microsecond resolution. pandas 3.x treats that mismatch as
            # an unsafe implicit upcast and raises
            # "TypeError: Invalid value '...' for dtype 'datetime64[s, UTC]'"
            # the moment we assign into it — pinning both columns to "ns"
            # up front means the later assignment always matches exactly,
            # whether the underlying data is empty, partial, or full.
            duty_df_raw["_start_dt"] = pd.to_datetime(
                duty_df_raw.get("actual_exit_time"), errors="coerce", utc=True, format="mixed"
            ).astype("datetime64[ns, UTC]")
            duty_df_raw["_end_dt"] = pd.to_datetime(
                duty_df_raw.get("actual_return_time"), errors="coerce", utc=True, format="mixed"
            ).astype("datetime64[ns, UTC]")

            # A trip only has meaningful "duty duration" once the Gate
            # Officer has logged a Gate Out (actual_exit_time). If Gate In
            # hasn't happened yet (still "On Trip"), treat "now" (in UTC, to
            # match the column's dtype) as the running end time so
            # in-progress duty shows up too. Assigning a naive Timestamp
            # into a UTC-aware column is exactly what triggers the
            # "Invalid value ... for dtype 'datetime64[us, UTC]'" TypeError,
            # so we must assign an equally tz-aware Timestamp here.
            still_out_mask = duty_df_raw["_start_dt"].notna() & duty_df_raw["_end_dt"].isna()
            duty_df_raw.loc[still_out_mask, "_end_dt"] = pd.Timestamp.now(tz="UTC")

            # Only rows that actually left the gate are real "duty" records.
            duty_base = duty_df_raw[duty_df_raw["_start_dt"].notna()].copy()

            # -------------------------------------------------------------
            # STEP 2 — Combine the date/time widgets into naive datetimes,
            # then localize them to UTC so they can be compared directly
            # against the tz-aware `_start_dt` / `_end_dt` columns above.
            # (If your admin users think in a local timezone rather than
            # UTC, swap "UTC" below for that zone, e.g. "Asia/Dhaka", and
            # pandas will convert correctly at comparison time.)
            # -------------------------------------------------------------
            range_start = pd.Timestamp(datetime.combine(filter_start_date, filter_start_time)).tz_localize("UTC")
            range_end = pd.Timestamp(datetime.combine(filter_end_date, filter_end_time)).tz_localize("UTC")

            if range_start >= range_end:
                st.error("⚠️ The start date/time must be earlier than the end date/time.")
                st.stop()

            fc3, fc4 = st.columns(2)
            with fc3:
                # Sourced from the Manage Drivers & Vehicles master list (not
                # from whatever vehicle_number values happen to appear in
                # historical requisition rows), so this dropdown always
                # matches the same fleet list used everywhere else in the app
                # (e.g. the Admin's Pending Requests approval dropdown).
                duty_vehicles_df = fetch_all_vehicles()
                vehicle_choices = ["All Vehicles"] + sorted(
                    v for v in duty_vehicles_df["vehicle_number"].dropna().unique().tolist() if v
                ) if not duty_vehicles_df.empty else ["All Vehicles"]
                duty_vehicle_filter = st.selectbox("Vehicle Selection", vehicle_choices, key="duty_vehicle_filter")
            with fc4:
                # Same idea for drivers: pulled from the Manage Drivers &
                # Vehicles master list rather than the trip data, so it's
                # the single source of truth for driver identity everywhere
                # (Duty Tracker, Pending Requests approval, KM Variance, and
                # the Driver Dashboard's own trip matching).
                duty_drivers_df = fetch_all_drivers()
                driver_choices = ["All Drivers"] + sorted(
                    d for d in duty_drivers_df["driver_name"].dropna().unique().tolist() if d
                ) if not duty_drivers_df.empty else ["All Drivers"]
                duty_driver_filter = st.selectbox("Driver Selection", driver_choices, key="duty_driver_filter")

            # -------------------------------------------------------------
            # STEP 3 — Apply the date/time window + vehicle/driver filters.
            # A trip is included if its duty window OVERLAPS the selected
            # range at all (not just if it starts inside it) — this correctly
            # captures overnight duties like "07:00 today to 06:59 tomorrow".
            # Both sides are now UTC-aware, so this comparison is safe.
            # -------------------------------------------------------------
            duty_filtered = duty_base[
                (duty_base["_start_dt"] <= range_end) & (duty_base["_end_dt"] >= range_start)
            ].copy()

            if duty_vehicle_filter != "All Vehicles":
                duty_filtered = duty_filtered[
                    duty_filtered["vehicle_number"].astype(str).str.strip() == duty_vehicle_filter.strip()
                ]
            if duty_driver_filter != "All Drivers":
                # Case/whitespace-insensitive match (same normalization as
                # the Driver Dashboard) so picking a name from the master
                # list above reliably matches trip rows even if a requisition
                # was saved with a slightly different capitalization/spacing.
                duty_filtered = duty_filtered[
                    duty_filtered["driver_name"].map(_normalize_driver_name) == _normalize_driver_name(duty_driver_filter)
                ]

            # -------------------------------------------------------------
            # STEP 4 — Derived fields: duty duration in hours, total KM.
            # Total KM now comes from effective_km_fields(), which prefers
            # the Driver's own logged Start/End KM and falls back to the
            # Gate Officer's Gate-Out/Gate-In readings only when the driver
            # hasn't submitted their own numbers yet (business requirement:
            # Driver KM is the primary baseline metric for duty tracking).
            # -------------------------------------------------------------
            duty_filtered["_duration_hrs"] = (
                (duty_filtered["_end_dt"] - duty_filtered["_start_dt"]).dt.total_seconds() / 3600.0
            ).round(2)
            duty_filtered["_km"] = duty_filtered.apply(lambda row: effective_km_fields(row)[2], axis=1)

            st.markdown("---")
            st.markdown("##### 📌 Summary")

            total_km = float(duty_filtered["_km"].sum())
            total_duration_hrs = float(duty_filtered["_duration_hrs"].sum())
            total_trips = int(len(duty_filtered))
            duration_h = int(total_duration_hrs)
            duration_m = int(round((total_duration_hrs - duration_h) * 60))

            # "Active Driver & Assigned Vehicle" — meaningful only when the
            # filters have narrowed things down to one driver/vehicle; with
            # "All Drivers"/"All Vehicles" selected we instead surface the
            # busiest one within the filtered window.
            if duty_driver_filter != "All Drivers":
                active_driver_label = duty_driver_filter
            elif not duty_filtered.empty and duty_filtered["driver_name"].notna().any():
                active_driver_label = duty_filtered["driver_name"].value_counts().idxmax()
            else:
                active_driver_label = "—"

            if duty_vehicle_filter != "All Vehicles":
                active_vehicle_label = duty_vehicle_filter
            elif not duty_filtered.empty and duty_filtered["vehicle_number"].notna().any():
                active_vehicle_label = duty_filtered["vehicle_number"].value_counts().idxmax()
            else:
                active_vehicle_label = "—"

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("🛣️ Total Distance", f"{total_km:.1f} KM")
            k2.metric("⏱️ Total Duty Duration", f"{duration_h}h {duration_m}m")
            k3.metric("🚗 Total Trips", total_trips)
            k4.metric("👨‍✈️ Active Driver / Vehicle", f"{active_driver_label} / {active_vehicle_label}")

            # -------------------------------------------------------------
            # STEP 5 — Detailed table:
            # [Vehicle No, Driver Name, Start Time, End Time, Total KM,
            #  Duty Duration (Hours), Route / Purpose]
            # -------------------------------------------------------------
            st.markdown("---")
            st.markdown("##### 📋 Detailed Duty Log")

            if duty_filtered.empty:
                st.info("No trips match the selected filters.")
                detail_display = pd.DataFrame(columns=DUTY_TRACKER_DISPLAY_COLS)
            else:
                detail_display = pd.DataFrame({
                    "Vehicle No": duty_filtered["vehicle_number"].map(lambda v: fmt(v, "—")),
                    "Driver Name": duty_filtered["driver_name"].map(lambda v: fmt(v, "—")),
                    "Start Time": duty_filtered["_start_dt"].dt.strftime("%Y-%m-%d %I:%M %p"),
                    "End Time": duty_filtered["_end_dt"].dt.strftime("%Y-%m-%d %I:%M %p"),
                    "Total KM": duty_filtered["_km"].round(1),
                    "Duty Duration (Hrs)": duty_filtered["_duration_hrs"],
                    "Route / Purpose": duty_filtered["destination"].fillna("").astype(str)
                                        + " — " + duty_filtered["purpose"].fillna("").astype(str),
                }).reset_index(drop=True)
                st.dataframe(detail_display, use_container_width=True, hide_index=True, height=340)

            # -------------------------------------------------------------
            # STEP 6 — Multi-format export: CSV / Excel / PDF.
            # summary_metrics feeds both the Excel "Summary" sheet and the
            # PDF's KPI block — plug in any additional metrics here as needed.
            # -------------------------------------------------------------
            summary_metrics = {
                "Date/Time Range": f"{range_start.strftime('%Y-%m-%d %I:%M %p')} to {range_end.strftime('%Y-%m-%d %I:%M %p')}",
                "Vehicle Filter": duty_vehicle_filter,
                "Driver Filter": duty_driver_filter,
                "Total Distance (KM)": f"{total_km:.1f}",
                "Total Duty Duration": f"{duration_h}h {duration_m}m",
                "Total Trips": total_trips,
                "Active Driver": active_driver_label,
                "Assigned Vehicle": active_vehicle_label,
            }
            filters_summary_text = (
                f"Range: {summary_metrics['Date/Time Range']} | Vehicle: {duty_vehicle_filter} | "
                f"Driver: {duty_driver_filter}"
            )

            st.markdown("---")
            st.markdown("##### ⬇️ Export Duty Report")
            e1, e2, e3 = st.columns(3)
            with e1:
                csv_bytes = detail_display.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Download CSV", data=csv_bytes, file_name="duty_tracker_report.csv",
                    mime="text/csv", use_container_width=True,
                )
            with e2:
                duty_excel_bytes = build_duty_tracker_excel(detail_display, summary_metrics)
                st.download_button(
                    "⬇️ Download Excel (.xlsx)", data=duty_excel_bytes, file_name="duty_tracker_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with e3:
                duty_pdf_bytes = build_duty_tracker_pdf(detail_display, summary_metrics, filters_summary_text)
                st.download_button(
                    "⬇️ Download PDF (.pdf)", data=duty_pdf_bytes, file_name="duty_tracker_report.pdf",
                    mime="application/pdf", use_container_width=True,
                )

    # ---------------- KM Variance Report (NEW) ----------------
    with tab_variance:
        st.subheader("📈 KM Variance Report — Driver vs Gate Officer Readings")
        st.caption(
            "Compares the Driver's self-logged odometer readings against the "
            "Gate Officer's Gate-Out/Gate-In readings for the same trip. "
            "Variance = (Driver End − Driver Start) − (Gate End − Gate Start)."
        )

        if df_all.empty or "driver_start_km" not in df_all.columns:
            st.info(
                "No driver-submitted KM data yet. This report populates once Drivers "
                "start logging their Start/End KM from the Driver Dashboard."
            )
        else:
            variance_source = df_all[df_all["driver_start_km"].notna()].copy()

            if variance_source.empty:
                st.info(
                    "No driver-submitted KM entries yet. This report populates once "
                    "Drivers start logging their Start/End KM from the Driver Dashboard."
                )
            else:
                def _driver_dist(row):
                    if is_blank(row.get("driver_start_km")) or is_blank(row.get("driver_end_km")):
                        return None
                    return round(float(row["driver_end_km"]) - float(row["driver_start_km"]), 1)

                def _gate_dist(row):
                    if is_blank(row.get("start_km")) or is_blank(row.get("end_km")):
                        return None
                    return round(float(row["end_km"]) - float(row["start_km"]), 1)

                rows = []
                for _, r in variance_source.iterrows():
                    d_dist = _driver_dist(r)
                    s_dist = _gate_dist(r)
                    variance = round(d_dist - s_dist, 1) if (d_dist is not None and s_dist is not None) else None
                    rows.append({
                        "Trip ID": short_req_id(r.get("id")),
                        "Vehicle No": fmt(r.get("vehicle_number")),
                        "Driver Name": fmt(r.get("driver_name")),
                        "Driver Start KM": fmt(r.get("driver_start_km")),
                        "Driver End KM": fmt(r.get("driver_end_km")),
                        "Gate Start KM": fmt(r.get("start_km")),
                        "Gate End KM": fmt(r.get("end_km")),
                        "Variance (KM)": variance if variance is not None else "—",
                    })

                # Plain, standard st.dataframe — no background colors or
                # highlight styling, per requirement.
                variance_df = pd.DataFrame(rows)
                st.dataframe(variance_df, use_container_width=True, hide_index=True, height=400)

                colA, colB = st.columns(2)
                with colA:
                    var_excel = build_excel_report(variance_df)
                    st.download_button(
                        "⬇️ Export to Excel (.xlsx)", data=var_excel,
                        file_name="km_variance_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with colB:
                    csv_bytes = variance_df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        "⬇️ Export to CSV (.csv)", data=csv_bytes,
                        file_name="km_variance_report.csv", mime="text/csv",
                        use_container_width=True,
                    )

    # ---------------- Management Dashboard (NEW) ----------------
    with tab_management:
        # Admins get the same Executive/Management overview as the
        # standalone 'executive' role, reusing the exact same function so
        # both roles always see identical KPI numbers for the same filters.
        render_management_dashboard(df_all)

# =========================================================
# 11. FALLBACK — unrecognized role
# =========================================================
else:
    st.error("⚠️ Your account role is not recognized. Please contact the Admin.")
