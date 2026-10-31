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
import types
import json
import base64
import hashlib
import random
import re
import time
from datetime import datetime, date, timedelta, timezone
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
import httpx

# ---------------------------------------------------------------------------
# Supabase "ConnectionTerminated" fix: the Supabase client is cached for the
# whole life of the app, and its pooled HTTP/2 connection is often closed by
# the server (or by Windows/your network) while the app sits idle. The next
# request then fails with httpx.RemoteProtocolError. Here, such a request is
# simply retried once on a fresh connection. Safe requests (read / update /
# delete) are retried on any connection error; an INSERT (POST) is retried
# only if the connection could not even be opened, so a requisition can never
# be saved twice.
# ---------------------------------------------------------------------------
_httpx_original_send = httpx.Client.send


def _httpx_send_with_reconnect(self, request, *args, **kwargs):
    try:
        return _httpx_original_send(self, request, *args, **kwargs)
    except (httpx.RemoteProtocolError, httpx.ReadError, httpx.WriteError, httpx.ConnectError) as exc:
        if request.method.upper() == "POST" and not isinstance(exc, httpx.ConnectError):
            raise
        return _httpx_original_send(self, request, *args, **kwargs)


if not getattr(httpx.Client.send, "_rbl_reconnect_patched", False):
    _httpx_send_with_reconnect._rbl_reconnect_patched = True
    httpx.Client.send = _httpx_send_with_reconnect


# =========================================================
# BANGLADESH TIME (UTC+6) — the ONE clock this whole app uses
# =========================================================
# Streamlit Cloud (and most servers) run on UTC, so a bare datetime.now() /
# bd_today() there is 6 hours behind Bangladesh — a trip started at 7:00 AM
# was being saved as 1:00 AM. Every timestamp the app writes, every "today" /
# "tomorrow" it computes and every "now" it compares against now comes from
# these helpers instead, so the result is the same on a UTC cloud server and
# on a Bangladesh laptop. Bangladesh has no daylight-saving, so a fixed +6
# offset is exact (and needs no tz database, which Windows may not have).
BD_TZ = timezone(timedelta(hours=6))


def bd_now() -> datetime:
    """Current Bangladesh wall-clock time as a plain (naive) datetime."""
    return datetime.now(BD_TZ).replace(tzinfo=None)


def bd_today() -> date:
    """Today's date in Bangladesh."""
    return bd_now().date()


def bd_now_str() -> str:
    """Bangladesh 'now' formatted for saving to the database."""
    return bd_now().strftime("%Y-%m-%d %H:%M:%S")


def bd_now_ts() -> pd.Timestamp:
    """Bangladesh 'now' as a Timestamp labeled UTC. Supabase hands the saved
    Bangladesh wall-clock values back with a '+00:00' label (the numbers are
    still Bangladesh time), and pandas parses them with utc=True — so to
    compare or subtract against those columns, 'now' must carry the same
    label. Only the numbers matter; nothing is converted."""
    return pd.Timestamp(bd_now()).tz_localize("UTC")


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
    # Save created_at in Bangladesh time (instead of the database's own
    # now(), which is UTC) so "submitted today"/"Requested on" are correct.
    data = dict(data)
    data.setdefault("created_at", bd_now_str())
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
    updates that should stay silent (quiet record-keeping edits, e.g. Admin
    corrections or syncing the other requisitions of a group trip), while
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
    # `notify`, so status changes (approve/reject, trip start/end, driver KM
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
ROLE_OPTIONS = ["user", "driver", "admin", "executive", "nurse"]
ROLE_DISPLAY = {
    "user": "Employee",
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
            out = dt.strftime("%Y-%m-%d %I:%M %p") if "%Y" in f else dt.strftime("%I:%M %p")
            return drop_hour_zero(out)
    return s


def drop_hour_zero(text: str) -> str:
    """'2026-09-28 01:05 PM' -> '2026-09-28 1:05 PM' and '01:05 PM' -> '1:05 PM'
    (works on Windows too, unlike strftime's %-I). Only the hour's leading
    zero is removed; dates like '2026-09-08' are left alone."""
    return re.sub(r"(^|\s)0(\d:)", r"\1\2", str(text))


def fmt_clock_12h(value, default: str = "—") -> str:
    """Compact clock for dashboards: '1:00 PM' when the moment is today
    (Bangladesh date), otherwise '27-Sep 1:00 PM'."""
    full = fmt_time_12h(value, default)
    m = re.match(r"^(\d{4}-\d{2}-\d{2}) (\d{1,2}:\d{2} [AP]M)$", full)
    if not m:
        return full
    day, clock = m.groups()
    if day == str(bd_today()):
        return clock
    return f"{datetime.strptime(day, '%Y-%m-%d').strftime('%d-%b')} {clock}"


TIMESTAMP_COLUMNS = ["created_at", "action_timestamp", "actual_exit_time",
                     "actual_return_time", "driver_km_updated_at"]


def humanize_timestamp_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Copy of `df` with the raw database timestamp columns rewritten as
    12-hour Bangladesh-time text (e.g. '2026-09-28 1:00 PM') instead of raw
    ISO strings with a '+00:00' label — used for on-screen tables and exports."""
    out = df.copy()
    for c in TIMESTAMP_COLUMNS:
        if c in out.columns:
            out[c] = out[c].apply(lambda v: fmt_time_12h(v, ""))
    return out


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
        default_time = bd_now().time()
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
    fetch_active_requisitions.clear()
    fetch_requisitions_by_user.clear()
    fetch_requisitions_by_driver.clear()
    get_last_driver_end_km.clear()
    fetch_auto_duty_trips.clear()


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
    data = dict(data)
    data.setdefault("created_at", bd_now_str())  # Bangladesh time, not the DB's UTC now()
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
    return f"REQ-{bd_now().strftime('%Y%m%d%H%M%S')}-{random.randint(100, 999)}"


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


REQUISITION_COLUMNS = [
    "id", "request_id", "created_at", "username", "applicant_name", "department", "mobile_number",
    "date_of_travel", "time_of_travel", "destination", "passenger_count", "vehicle_type", "purpose",
    "special_request", "status", "driver_name", "driver_contact", "vehicle_number", "approved_by",
    "action_timestamp", "approved_time", "admin_note", "start_km", "end_km", "total_km",
    "actual_exit_time", "actual_return_time", "driver_start_km", "driver_end_km", "driver_km_updated_at",
    "trip_group_id",
]


def _execute_all_pages(make_query, page_size: int = 1000):
    """Runs a Supabase query page by page and returns ALL rows.

    Supabase/PostgREST silently returns at most 1000 rows per request. Without
    paging, once the requisitions table grew past 1000 rows the OLDEST rows were
    quietly cut off (the Export tab's start date jumped forward and old trips
    vanished from reports even though they were still in the database).
    `make_query` must return a fresh, ordered query each time it is called."""
    rows, start = [], 0
    while True:
        batch = make_query().range(start, start + page_size - 1).execute().data or []
        rows.extend(batch)
        if len(batch) < page_size or start > 500000:
            break
        start += page_size
    return types.SimpleNamespace(data=rows)


@st.cache_data(ttl=90, show_spinner=False)
def fetch_all_requisitions() -> pd.DataFrame:
    sb = get_supabase_client()
    res = _execute_all_pages(
        lambda: sb.table(REQUISITIONS_TABLE).select("*").order("created_at", desc=True).order("id", desc=True)
    )
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=REQUISITION_COLUMNS)


@st.cache_data(ttl=90, show_spinner=False)
def fetch_active_requisitions() -> pd.DataFrame:
    """Only Pending / Approved / On Trip requisitions — a small, constant-size
    slice (it does NOT grow as history piles up), used by Admin's landing page
    so it opens fast. Sections that need the full history use fetch_all_requisitions()."""
    sb = get_supabase_client()
    res = _execute_all_pages(
        lambda: sb.table(REQUISITIONS_TABLE)
        .select("*")
        .in_("status", ["Pending", "Approved", "On Trip"])
        .order("created_at", desc=True)
        .order("id", desc=True)
    )
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=REQUISITION_COLUMNS)


@st.cache_data(ttl=90, show_spinner=False)
def fetch_requisitions_by_user(username: str) -> pd.DataFrame:
    """Server-side filtered fetch — strict data isolation: only this user's rows are ever requested."""
    sb = get_supabase_client()
    res = _execute_all_pages(
        lambda: sb.table(REQUISITIONS_TABLE)
        .select("*")
        .eq("username", username)
        .order("created_at", desc=True)
        .order("id", desc=True)
    )
    return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=REQUISITION_COLUMNS)


# ------------------- DRIVER-SUBMITTED KM HELPERS (NEW) -------------------
# Additive helpers backing the Driver Dashboard (role branch, Section 9B).
# They read/write the driver_start_km / driver_end_km / driver_km_updated_at
# columns (older trips may also carry start_km / end_km readings, which are
# only ever read, never written, by these helpers).
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


def _driver_ilike_pattern(name) -> str:
    """SQL ILIKE pattern that is guaranteed to match every driver_name whose
    _normalize_driver_name() equals this driver's: all the name's non-space
    characters, in order, with anything allowed in between ("Mr. Yusuf" still
    matches "Mr.Yusuf"). Used ONLY to cut down how many rows are downloaded
    from Supabase — the exact normalized comparison is still applied after.
    Returns "" when no usable pattern exists (caller then skips narrowing)."""
    chars = [c for c in str(name or "").casefold() if not c.isspace() and c not in "%_*\\"]
    return ("%" + "%".join(chars) + "%") if chars else ""


@st.cache_data(ttl=90, show_spinner=False)
def fetch_requisitions_by_driver(driver_name: str) -> pd.DataFrame:
    """All trips assigned to this driver, matched case-insensitively and
    ignoring leading/trailing whitespace (see _normalize_driver_name) —
    so a mismatch like "Yusuf" vs "yusuf " no longer hides trips from the
    Driver Dashboard. Fetches every requisition with a driver assigned and
    filters client-side rather than relying on Supabase's exact-match .eq(),
    since PostgREST has no case-insensitive-and-trimmed equality operator."""
    sb = get_supabase_client()
    pattern = _driver_ilike_pattern(driver_name)
    try:
        def _drv_query():
            q = sb.table(REQUISITIONS_TABLE).select("*").not_.is_("driver_name", "null")
            if pattern:
                # Only this driver's rows travel over the network (fast on weak links).
                q = q.ilike("driver_name", pattern)
            return q.order("created_at", desc=True).order("id", desc=True)
        res = _execute_all_pages(_drv_query)
    except Exception:
        # Safety net: fall back to the original full fetch + client-side filter.
        res = _execute_all_pages(
            lambda: sb.table(REQUISITIONS_TABLE)
            .select("*")
            .not_.is_("driver_name", "null")
            .order("created_at", desc=True)
            .order("id", desc=True)
        )
    if not res.data:
        return pd.DataFrame(columns=REQUISITION_COLUMNS)
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
    pattern = _driver_ilike_pattern(driver_name)
    try:
        q = (
            sb.table(REQUISITIONS_TABLE)
            .select("driver_name, driver_end_km, created_at")
            .eq("vehicle_number", vehicle_number)
            .not_.is_("driver_end_km", "null")
        )
        if pattern:
            q = q.ilike("driver_name", pattern)
        res = q.order("created_at", desc=True).limit(50).execute()
    except Exception:
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


# ------------------- GROUP TRIP HELPERS (NEW) -------------------
# Requisitions that Admin bulk-assigns share one `trip_group_id`. They are ONE
# physical trip, so the Driver sees them as ONE card, enter
# Start/End KM once, and Telegram sends ONE message.
def generate_trip_group_id() -> str:
    return f"GRP-{bd_now().strftime('%Y%m%d%H%M%S')}-{random.randint(100, 999)}"


def group_trips(df: pd.DataFrame) -> list:
    """Split df rows into trips: a list of lists of rows. Rows sharing a
    trip_group_id come together; rows without one are a trip of their own."""
    groups, seen = [], {}
    if df is None or df.empty:
        return groups
    for _, r in df.iterrows():
        gid = r.get("trip_group_id")
        if is_blank(gid):
            groups.append([r])
        else:
            if gid not in seen:
                seen[gid] = []
                groups.append(seen[gid])
            seen[gid].append(r)
    return groups


def _sort_dt(value):
    """Timestamp text -> naive datetime for ordering. Reads the stored wall-clock
    numbers as they are (same approach as fmt_time_12h) and copes with every shape
    Supabase returns: 'YYYY-MM-DD HH:MM:SS', ISO 'T', fractional seconds, +00:00.
    Missing/unreadable values sort LAST."""
    if is_blank(value):
        return datetime.max
    s = str(value).strip().replace("T", " ")
    s = re.sub(r"(\.\d+)?(Z|[+-]\d{2}:?\d{2})$", "", s, flags=re.IGNORECASE).strip()
    for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            continue
    return datetime.max


def sort_by_approval_order(df: pd.DataFrame) -> pd.DataFrame:
    """Oldest approval first: the requisition the Admin approved FIRST for a driver
    comes first (#1), the next approved is #2, and so on. action_timestamp holds
    the moment of approval; ties (e.g. a batch auto-created in the same second)
    fall back to creation time and then id, so the order stays stable."""
    if df is None or df.empty:
        return df
    d = df.copy()
    keys = []
    for col in ("action_timestamp", "created_at"):
        if col in d.columns:
            d["_k_" + col] = d[col].map(_sort_dt)
            keys.append("_k_" + col)
    if "id" in d.columns:
        d["_k_id"] = pd.to_numeric(d["id"], errors="coerce")
        keys.append("_k_id")
    if keys:
        d = d.sort_values(keys, ascending=True, na_position="last", kind="mergesort")
    return d.drop(columns=[c for c in d.columns if c.startswith("_k_")])


def mark_group_duplicates(df: pd.DataFrame) -> pd.Series:
    """True for the 2nd+ requisition of the same group trip, so distance/hours
    of one physical trip are not counted several times in totals."""
    if df is None or df.empty or "trip_group_id" not in df.columns:
        return pd.Series(False, index=df.index if df is not None else None)
    g = df["trip_group_id"]
    has = g.notna() & (g.astype(str).str.strip() != "")
    return has & g.duplicated(keep="first")


def bulk_update_requisitions(request_ids: list, updates: dict):
    """One database call that updates many requisitions with the same values."""
    sb = get_supabase_client()
    sb.table(REQUISITIONS_TABLE).update(updates).in_("request_id", list(request_ids)).execute()
    _clear_requisition_caches()


def attach_requisitions_to_trip(request_ids: list, pending_rows: list, target, approved_by: str):
    """Add already-submitted Pending requisitions to a trip that is ALREADY
    Approved / On Trip (same driver, vehicle, time). They join the target's
    group (one is created if it has none), so the Driver sees ONE trip and
    enters KM once. If the target is already On Trip, the new requisitions
    inherit its start time and Start KM too. ONE Telegram message."""
    gid = target.get("trip_group_id")
    if is_blank(gid):
        gid = generate_trip_group_id()
        update_requisition(target["request_id"], {"trip_group_id": gid}, notify=False)
    approved_hhmm = target.get("approved_time") if not is_blank(target.get("approved_time")) \
        else target.get("time_of_travel")
    updates = {
        "status": target.get("status"),
        "driver_name": fmt(target.get("driver_name"), ""),
        "driver_contact": fmt(target.get("driver_contact"), ""),
        "vehicle_number": fmt(target.get("vehicle_number"), ""),
        "approved_by": approved_by,
        "action_timestamp": bd_now_str(),
        "approved_time": approved_hhmm,
        "trip_group_id": gid,
    }
    if target.get("status") == "On Trip":
        if not is_blank(target.get("actual_exit_time")):
            updates["actual_exit_time"] = str(target.get("actual_exit_time"))
        for f in ("start_km", "driver_start_km"):
            if not is_blank(target.get(f)):
                updates[f] = float(target.get(f))
    bulk_update_requisitions(request_ids, updates)

    lines = [
        f"• {short_req_id(r.get('id'))} — {r.get('applicant_name', '')} "
        f"({r.get('department', '')}) → {r.get('destination', '')}"
        for r in pending_rows
    ]
    send_telegram_alert(
        f"➕ **Added to an existing trip! ({len(pending_rows)} more requisition(s))**\n\n"
        + "\n".join(lines)
        + f"\n\n🔗 **Joined trip:** {short_req_id(target.get('id'))} ({target.get('status')})"
        + f"\n👨‍✈️ **Driver:** {updates['driver_name']}\n🚗 **Vehicle:** {updates['vehicle_number']}"
        + f"\n🕒 **Departure:** {fmt_time_12h(approved_hhmm)}\n✅ **Approved by:** {approved_by}"
    )


def send_bulk_assignment_alert(rows: list, driver: str, vehicle: str, approved_hhmm: str,
                               approved_by: str, note: str = ""):
    """ONE Telegram message for a whole bulk assignment."""
    lines = [
        f"• {short_req_id(r.get('id'))} — {r.get('applicant_name', '')} "
        f"({r.get('department', '')}) → {r.get('destination', '')}"
        for r in rows
    ]
    msg = (
        f"🚐 **Group Trip Assigned! ({len(rows)} requisitions, 1 vehicle)**\n\n"
        + "\n".join(lines)
        + f"\n\n👨‍✈️ **Driver:** {driver}\n🚗 **Vehicle:** {vehicle}"
        + f"\n🕒 **Departure:** {fmt_time_12h(approved_hhmm)}"
        + f"\n✅ **Approved by:** {approved_by}"
    )
    if note:
        msg += f"\n📝 **Note:** {note}"
    send_telegram_alert(msg)


def send_group_status_alert(rows: list, status: str, driver: str = "", vehicle: str = ""):
    """ONE Telegram message when a group trip starts / completes."""
    ids = ", ".join(short_req_id(r.get("id")) for r in rows)
    applicants = ", ".join(dict.fromkeys(str(r.get("applicant_name", "")) for r in rows))
    dests = ", ".join(dict.fromkeys(str(r.get("destination", "")) for r in rows))
    msg = (
        f"📢 **Group Trip Status Updated!**\n\n"
        f"🆔 **Requisitions:** {ids}\n"
        f"👤 **Applicants:** {applicants}\n"
        f"📍 **Destinations:** {dests}\n"
        f"📌 **New Status:** {status}"
    )
    if not is_blank(driver):
        msg += f"\n👨‍✈️ **Driver:** {driver}"
    if not is_blank(vehicle):
        msg += f"\n🚗 **Vehicle:** {vehicle}"
    send_telegram_alert(msg)


def _build_driver_km_updates(row: dict, driver_start_km=None, driver_end_km=None, now_str=None) -> dict:
    """The DB fields to write for one requisition when the driver enters KM.
    (Same rules as before: Start KM -> On Trip, End KM -> Completed.)
    is_blank() is used so a pandas NaN never reaches the JSON request."""
    now_str = now_str or bd_now_str()
    updates = {
        "driver_km_updated_at": now_str,
        "driver_name": row.get("driver_name", ""),
        "vehicle_number": row.get("vehicle_number", ""),
    }
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
    return updates


def submit_driver_km(row: dict, driver_start_km=None, driver_end_km=None):
    """Single-requisition driver KM entry (+ the usual Telegram alert)."""
    updates = _build_driver_km_updates(row, driver_start_km, driver_end_km)
    update_requisition(row["request_id"], updates)


def submit_driver_km_group(rows: list, driver_start_km=None, driver_end_km=None):
    """Driver enters Start/End KM ONCE for a whole group trip: the same values
    are saved on every requisition in the group and ONE Telegram message is
    sent. A group of one behaves exactly like submit_driver_km()."""
    if len(rows) == 1:
        return submit_driver_km(rows[0], driver_start_km, driver_end_km)
    now_str = bd_now_str()
    final_status = None
    for row in rows:
        updates = _build_driver_km_updates(row, driver_start_km, driver_end_km, now_str)
        update_requisition(row["request_id"], updates, notify=False)
        final_status = updates.get("status", final_status)
    first = rows[0]
    send_group_status_alert(rows, final_status or "Updated",
                            first.get("driver_name", ""), first.get("vehicle_number", ""))


# ------------------- EASY-EDIT INPUT HELPERS (NEW) -------------------
def parse_stored_datetime(value):
    """Stored timestamp string (plain or ISO with +00:00) -> naive datetime, or None."""
    if is_blank(value):
        return None
    s = str(value).strip().replace("T", " ")
    s = re.sub(r"(\.\d+)?(Z|[+-]\d{2}:?\d{2})$", "", s, flags=re.IGNORECASE).strip()
    for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            continue
    return None


def optional_datetime_input(label: str, key: str, existing_value):
    """Checkbox + date + 12-hour time. Returns 'YYYY-MM-DD HH:MM:SS' or None
    (None = not recorded / clear it). Use OUTSIDE st.form so the checkbox
    reacts instantly."""
    existing = parse_stored_datetime(existing_value)
    on = st.checkbox(f"{label} recorded", value=existing is not None, key=f"{key}_on")
    if not on:
        return None
    base = existing or bd_now()
    d = st.date_input(f"{label} — date", value=base.date(), key=f"{key}_d")
    t = time_input_12h(f"{label} — time", key_prefix=f"{key}_t", default_time=base.time())
    return datetime.combine(d, t).strftime("%Y-%m-%d %H:%M:%S")


def optional_time_input(label: str, key: str, existing_value, fallback_time=None):
    """Checkbox + 12-hour time. Returns 'HH:MM' or None."""
    existing = None
    if not is_blank(existing_value):
        for f in ("%H:%M:%S", "%H:%M"):
            try:
                existing = datetime.strptime(str(existing_value).strip(), f).time()
                break
            except ValueError:
                continue
    on = st.checkbox(f"{label} set", value=existing is not None, key=f"{key}_on")
    if not on:
        return None
    t = time_input_12h(label, key_prefix=f"{key}_t",
                       default_time=existing or fallback_time or bd_now().time())
    return t.strftime("%H:%M")


def km_text_input(label: str, value, key: str) -> str:
    if is_blank(value):
        shown = ""
    else:
        v = float(value)
        shown = str(int(v)) if v == int(v) else str(round(v, 1))
    return st.text_input(label, value=shown, key=key, placeholder="empty = blank")


def km_plain(v) -> str:
    """383939.0 -> '383939' ; 120.5 -> '120.5' (no useless '.0')."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ""
    return str(int(f)) if f == int(f) else str(round(f, 1))


def driver_km_input(label: str, value, key: str, help: str = None) -> float:
    """Odometer input for drivers. Whole-number readings are shown WITHOUT '.0'
    (e.g. 383939, not 383939.0) and get a numeric keypad on phones, so the
    driver can quickly edit just the last digits. A reading that really has a
    decimal part (e.g. 120.5) still falls back to a one-decimal field."""
    v = float(value or 0)
    if v == int(v):
        return float(st.number_input(label, min_value=0, step=1, value=int(v),
                                     format="%d", key=key, help=help))
    return float(st.number_input(label, min_value=0.0, step=1.0, value=v,
                                 format="%.1f", key=key, help=help))


def show_admin_note(grp) -> None:
    """Prominent, hard-to-miss Admin Note box for the driver screens."""
    import html as _html
    notes = []
    for x in grp:
        n = x.get("admin_note")
        if not is_blank(n):
            n = str(n).strip()
            if n and n not in [t for _, t in notes]:
                notes.append((short_req_id(x.get("id")), n))
    if not notes:
        return
    if len(notes) == 1:
        body = _html.escape(notes[0][1]).replace("\n", "<br>")
    else:
        body = "<br>".join(f"<b>{_html.escape(i)}:</b> {_html.escape(t)}" for i, t in notes)
    st.markdown(
        '<div style="background:#FFF4CE;border:2px solid #F59E0B;border-left:8px solid #D97706;'
        'border-radius:10px;padding:.8rem 1rem;margin:.4rem 0 .8rem;color:#78350F;font-size:1.05rem;">'
        f'<div style="font-weight:800;letter-spacing:.04em;">📝 ADMIN NOTE</div>{body}</div>',
        unsafe_allow_html=True,
    )


def parse_km_text(text):
    """-> (value or None, ok)."""
    t = (text or "").strip()
    if not t:
        return None, True
    try:
        v = float(t)
    except ValueError:
        return None, False
    return (v, True) if v >= 0 else (None, False)


def admin_edit_trip_panel(row, df_all: pd.DataFrame, drivers_df: pd.DataFrame, vehicles_df: pd.DataFrame):
    """Edit ANY requisition (any status): every detail, status, driver/vehicle,
    approved time, trip start/end date+time, KM readings, admin note.
    Not an st.form on purpose: checkboxes / fields react instantly."""
    rid = row["request_id"]
    short = short_req_id(row.get("id"))
    # Widget keys include a signature of the saved row, so if the data changes
    # (e.g. the driver finishes the trip) the form reloads the fresh values
    # instead of showing stale ones.
    sig = hashlib.md5(str(row.to_dict()).encode("utf-8")).hexdigest()[:8]

    def k(name):
        return f"et_{name}_{rid}_{sig}"

    driver_contact_map = dict(zip(drivers_df["driver_name"], drivers_df["driver_contact"])) if not drivers_df.empty else {}
    driver_choices = drivers_df["driver_name"].tolist() if not drivers_df.empty else []
    vehicle_choices = vehicles_df["vehicle_number"].tolist() if not vehicles_df.empty else []
    current_driver = fmt(row.get("driver_name"), "")
    current_vehicle = fmt(row.get("vehicle_number"), "")
    if current_driver and current_driver not in driver_choices:
        driver_choices = [current_driver] + driver_choices
    if current_vehicle and current_vehicle not in vehicle_choices:
        vehicle_choices = [current_vehicle] + vehicle_choices
    driver_display = ["— None —"] + driver_choices
    vehicle_display = ["— None —"] + vehicle_choices

    st.markdown("---")
    st.markdown(f"#### ✏️ Editing {short}  —  {STATUS_BADGE.get(row.get('status'), row.get('status'))}")
    st.caption(f"Technical ID: `{rid}`")

    group_rows = pd.DataFrame()
    gid = row.get("trip_group_id")
    if not is_blank(gid) and "trip_group_id" in df_all.columns:
        group_rows = df_all[df_all["trip_group_id"] == gid]
    in_group = len(group_rows) > 1
    apply_group = False
    if in_group:
        st.info("🚐 This requisition is part of a group trip with: "
                + ", ".join(short_req_id(x) for x in group_rows["id"]))
        apply_group = st.checkbox(
            "Apply Driver, Vehicle, KM and trip times to ALL requisitions of this group trip",
            value=True, key=k("apply_group"),
        )

    # ---------------- Status ----------------
    cur_status = row.get("status")
    et_status = st.selectbox(
        "Status", REQ_STATUS_OPTIONS,
        index=REQ_STATUS_OPTIONS.index(cur_status) if cur_status in REQ_STATUS_OPTIONS else 0,
        key=k("status"),
    )

    # ---------------- Trip details ----------------
    st.markdown("##### Trip Details")
    e1, e2 = st.columns(2)
    dept_list = DEPARTMENTS if row.get("department") in DEPARTMENTS or is_blank(row.get("department")) \
        else DEPARTMENTS + [row.get("department")]
    vtype_list = VEHICLE_TYPES if row.get("vehicle_type") in VEHICLE_TYPES or is_blank(row.get("vehicle_type")) \
        else VEHICLE_TYPES + [row.get("vehicle_type")]
    try:
        pc_default = min(max(int(float(row.get("passenger_count"))), 1), 50)
    except (TypeError, ValueError):
        pc_default = 1
    try:
        date_default = datetime.strptime(str(row.get("date_of_travel"))[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        date_default = bd_today()
    try:
        time_default = datetime.strptime(fmt(row.get("time_of_travel"), "09:00")[:5], "%H:%M").time()
    except ValueError:
        time_default = bd_now().time()

    with e1:
        et_applicant = st.text_input("Applicant Name", value=fmt(row.get("applicant_name"), ""), key=k("app"))
        et_department = st.selectbox(
            "Department", dept_list,
            index=dept_list.index(row.get("department")) if row.get("department") in dept_list else 0,
            key=k("dept"),
        )
        et_mobile = st.text_input("Mobile Number", value=fmt(row.get("mobile_number"), ""), key=k("mob"))
        et_passengers = st.number_input("Passenger Count", min_value=1, max_value=50, value=pc_default, key=k("pax"))
    with e2:
        et_date = st.date_input("Date of Travel", value=date_default, key=k("date"))
        et_time = time_input_12h("Time of Travel", key_prefix=k("tt"), default_time=time_default)
        et_destination = st.text_input("Destination", value=fmt(row.get("destination"), ""), key=k("dest"))
        et_vehicle_type = st.selectbox(
            "Vehicle Type", vtype_list,
            index=vtype_list.index(row.get("vehicle_type")) if row.get("vehicle_type") in vtype_list else 0,
            key=k("vtype"),
        )
    et_purpose = st.text_area("Purpose", value=fmt(row.get("purpose"), ""), height=80, key=k("purpose"))
    et_special = st.text_area("Special Request", value=fmt(row.get("special_request"), ""), height=60, key=k("special"))

    # ---------------- Driver & vehicle ----------------
    st.markdown("##### Driver & Vehicle")
    d1, d2 = st.columns(2)
    with d1:
        et_driver = st.selectbox(
            "Driver", driver_display,
            index=driver_display.index(current_driver) if current_driver in driver_display else 0,
            key=k("driver"),
        )
    with d2:
        et_vehicle = st.selectbox(
            "Vehicle Number", vehicle_display,
            index=vehicle_display.index(current_vehicle) if current_vehicle in vehicle_display else 0,
            key=k("vehicle"),
        )
    et_approved_time = optional_time_input(
        "Approved Departure Time", k("appr"), row.get("approved_time"), fallback_time=et_time,
    )

    # ---------------- Actual times ----------------
    st.markdown("##### 🕒 Actual Trip Times (Trip Start and Trip End)")
    t1, t2 = st.columns(2)
    with t1:
        et_exit = optional_datetime_input("Trip Start", k("exit"), row.get("actual_exit_time"))
    with t2:
        et_return = optional_datetime_input("Trip End", k("ret"), row.get("actual_return_time"))

    # ---------------- KM ----------------
    st.markdown("##### 🛣️ Odometer / KM  (empty box = blank / not recorded)")
    k1, k2 = st.columns(2)
    with k1:
        dsk_txt = km_text_input("Start KM", row.get("driver_start_km"), k("dsk"))
    with k2:
        dek_txt = km_text_input("End KM", row.get("driver_end_km"), k("dek"))
    if is_blank(row.get("driver_start_km")) and is_blank(row.get("driver_end_km")) \
            and not is_blank(row.get("start_km")) and not is_blank(row.get("end_km")):
        st.caption(
            f"ℹ️ This older trip already has readings on record — Start {km_plain(row.get('start_km'))}, "
            f"End {km_plain(row.get('end_km'))}. They stay in the reports unless you enter Start/End KM above."
        )

    et_note = st.text_area("Admin Note", value=fmt(row.get("admin_note"), ""), height=60, key=k("note"))
    et_notify = st.checkbox("📢 Send a Telegram notification about this correction", value=False, key=k("notify"))

    if st.button("💾 Save Changes", type="primary", use_container_width=True, key=k("save")):
        dsk, ok3 = parse_km_text(dsk_txt)
        dek, ok4 = parse_km_text(dek_txt)
        errors = []
        if not (ok3 and ok4):
            errors.append("KM values must be numbers (0 or more) or left empty.")
        elif dsk is not None and dek is not None and dek < dsk:
            errors.append("End KM cannot be less than Start KM.")
        if et_exit and et_return and et_return < et_exit:
            errors.append("Trip End time cannot be earlier than Trip Start time.")

        if errors:
            for msg in errors:
                st.error(msg)
        else:
            updates = {
                "applicant_name": et_applicant.strip(),
                "department": et_department,
                "mobile_number": et_mobile.strip(),
                "passenger_count": int(et_passengers),
                "status": et_status,
                "date_of_travel": str(et_date),
                "time_of_travel": et_time.strftime("%H:%M"),
                "destination": et_destination.strip(),
                "vehicle_type": et_vehicle_type,
                "purpose": et_purpose.strip(),
                "special_request": et_special.strip(),
                "driver_name": "" if et_driver == "— None —" else et_driver,
                "driver_contact": "" if et_driver == "— None —" else driver_contact_map.get(et_driver, ""),
                "vehicle_number": "" if et_vehicle == "— None —" else et_vehicle,
                "approved_time": et_approved_time,
                "actual_exit_time": et_exit,
                "actual_return_time": et_return,
                "driver_start_km": dsk,
                "driver_end_km": dek,
                "admin_note": et_note.strip(),
            }
            # Older trips may still carry start_km/end_km readings (never edited here):
            # keep them as the fallback so their total isn't wiped by this save.
            updates["total_km"] = effective_km_fields({
                "start_km": row.get("start_km"), "end_km": row.get("end_km"),
                "driver_start_km": dsk, "driver_end_km": dek,
            })[2]
            try:
                update_requisition(rid, updates, notify=et_notify)
                n_others = 0
                if apply_group:
                    shared_fields = ("driver_name", "driver_contact", "vehicle_number", "approved_time",
                                     "actual_exit_time", "actual_return_time",
                                     "driver_start_km", "driver_end_km", "total_km")
                    shared = {f: updates[f] for f in shared_fields}
                    for other_id in group_rows["request_id"]:
                        if other_id != rid:
                            update_requisition(other_id, shared, notify=False)
                            n_others += 1
                st.session_state["edit_flash"] = (
                    f"✅ Requisition {short} updated"
                    + (f" (and {n_others} other requisition(s) in its group)." if n_others else ".")
                )
                st.rerun()
            except Exception as e:
                st.error(f"❌ Failed to update: {e}")

    # ---------------- Delete ----------------
    st.markdown("---")
    st.markdown("##### 🗑️ Delete This Requisition Permanently")
    st.caption(f"This permanently removes **{short}** from every report, export and dashboard. Cannot be undone.")
    confirm_del = st.checkbox(f"I understand this will permanently delete {short}.", key=k("confirm_del"))
    if st.button("🗑️ Delete This Requisition", type="primary", disabled=not confirm_del,
                 use_container_width=True, key=k("del_btn")):
        try:
            delete_requisition(rid)
            st.session_state["edit_flash"] = f"🗑️ {short} has been deleted."
            st.rerun()
        except Exception as e:
            st.error(f"❌ Failed to delete: {e}")


def effective_km_fields(row) -> tuple:
    """(start_km, end_km, total_km) using the Driver's own odometer entries
    as the PRIMARY source of truth for standard distance reports and duty
    tracking, falling back to the older start_km/end_km readings (trips
    recorded before drivers entered their own KM) only when the driver hasn't
    logged their own numbers for that trip yet. Read-only helper."""
    d_start, d_end = row.get("driver_start_km"), row.get("driver_end_km")
    if not is_blank(d_start) and not is_blank(d_end):
        s, e = float(d_start), float(d_end)
        return s, e, round(e - s, 1)
    s_start, s_end = row.get("start_km"), row.get("end_km")
    if not is_blank(s_start) and not is_blank(s_end):
        s, e = float(s_start), float(s_end)
        return s, e, round(e - s, 1)
    return None, None, 0.0


def display_km_fields(row) -> tuple:
    """(start_km, end_km, total_km) for showing ONE trip on screen / in exports.
    Unlike effective_km_fields() it also shows a lone Start KM while the trip is
    still running. Driver's reading first, older start_km/end_km as fallback.
    Values are floats or None. Read-only helper."""
    def _pick(first, second):
        for v in (row.get(first), row.get(second)):
            if not is_blank(v):
                return float(v)
        return None
    s, e = _pick("driver_start_km", "start_km"), _pick("driver_end_km", "end_km")
    t = row.get("total_km")
    if not is_blank(t):
        t = float(t)
    else:
        t = round(e - s, 1) if (s is not None and e is not None) else None
    return s, e, t


def compute_duty_hours(row) -> float:
    """How long a vehicle was actually out for a given trip, in hours.

    Uses the same trip start/end timestamps
    (actual_exit_time / actual_return_time) as the Duty Tracker tab, so the
    numbers here always agree with that tab. A trip that hasn't started
    yet contributes 0 hours; a trip that's still 'On Trip' (started but
    not yet ended) counts up to right now, so an in-progress duty still
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
        end_dt = bd_now_ts()
    else:
        end_dt = pd.to_datetime(end_raw, errors="coerce", utc=True, format="mixed")
        if pd.isna(end_dt):
            end_dt = bd_now_ts()

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


# ------------------- DAILY AUTO-GENERATED REQUISITIONS (UPDATED) -------------------
# Every day, the first time ANY logged-in user loads the app AFTER 6:30 PM
# (Bangladesh time) the requisitions below are created automatically (no
# button click needed). Four groups:
#
#   1. Staff Shuttle   (4 trips)  -> created as PENDING (Admin picks Driver +
#                                    Vehicle in "Pending Requests", like always)
#   2. Commercial Duty (6 trips)  -> ROSTER group: ONE driver + ONE vehicle for
#                                    all 6 trips, chosen in the "🤖 Auto Duty
#                                    Roster" admin tab. The 6 trips are ALWAYS
#                                    created already APPROVED (never Pending).
#                                    If no driver/vehicle is selected the group
#                                    is simply not created until one is.
#   3. Night Shift     (5 trips)  -> ROSTER group, same behaviour (one driver +
#                                    one vehicle for all 5 trips, always Approved).
#   4. Female Worker Drop (4 trips) -> created as PENDING.
#
# A roster selection STAYS the same every day until Admin changes it (weekly
# rotation = just re-select the driver/vehicle once a week). A group whose
# "Active" box is un-ticked is skipped (not created) that day.
#
# Each auto requisition is tagged in admin_note as "<TAG>#<slot no>", so
# re-checking never depends on row order and never creates a duplicate (safe
# even if Admin edits/deletes one of today's rows later).
DAILY_SHUTTLE_TAG = "AUTO-DAILY-SHUTTLE"
DAILY_COMMERCIAL_TAG = "AUTO-DAILY-COMMERCIAL"
DAILY_NIGHT_SHIFT_TAG = "AUTO-DAILY-NIGHT-SHIFT"
DAILY_FEMALE_DROP_TAG = "AUTO-DAILY-FEMALE-DROP"

# Auto-creation only starts at/after this Bangladesh time (24h clock).
DAILY_AUTO_START_HOUR = 18
DAILY_AUTO_START_MINUTE = 30

# Table that stores the roster (driver + vehicle chosen for each slot).
AUTO_ROSTER_TABLE = "auto_duty_roster"
AUTO_ROSTER_SQL = """create table if not exists auto_duty_roster (
    slot_key text primary key,
    group_tag text not null,
    slot_no int not null,
    driver_name text default '',
    vehicle_number text default '',
    trip_time text default '',
    enabled boolean default true,
    updated_by text default '',
    updated_at timestamptz default now(),
    trips_json text default ''
);
alter table auto_duty_roster add column if not exists trips_json text default '';
alter table auto_duty_roster disable row level security;"""

# ---- Group 1: Staff Shuttle (Pending) ----
DAILY_SHUTTLE_TEMPLATES = [
    {"destination": "Ishwardi", "time": "19:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
    {"destination": "Ishwardi", "time": "19:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
    {"destination": "Dashuria", "time": "19:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
    {"destination": "Ishwardi", "time": "20:15", "passenger_count": 8, "purpose": "Staff Drop", "vehicle_type": "HIACE"},
]

# ---- Group 2: Commercial Duty (ROSTER) — number of slots = range(6) ----
DAILY_COMMERCIAL_TEMPLATES = [
    {"destination": "Bepza/others", "time": "20:15", "passenger_count": 1, "purpose": "Commercial Duty",
     "vehicle_type": "HIACE", "label": f"Commercial Duty #{i}"}
    for i in range(1, 7)
]

# ---- Group 3: Night Shift Vehicle (ROSTER) ----
# "department" is different for every slot, so it is set per row here.
# The time below is only the starting default — Admin can change the
# departure time of each slot in the "🤖 Auto Duty Roster" tab.
DAILY_NIGHT_SHIFT_TEMPLATES = [
    {"destination": "Bepza", "time": "18:30", "passenger_count": 1, "purpose": "Admin general duty",
     "vehicle_type": "HIACE", "department": dept, "label": f"{dept} (Night #{i})"}
    for i, dept in enumerate(
        ["Fuel Refill", "QC PICK/Drop", "Patient Drop", "QC PICK/Drop", "Patient Drop"], start=1
    )
]

# ---- Group 4: Female Worker Drop (Pending) ----
DAILY_FEMALE_DROP_TEMPLATES = [
    {"destination": area, "time": "19:30", "passenger_count": 10, "purpose": "Female worker drop", "vehicle_type": "HIACE"}
    for area in ["Ishwardi area", "Dashuria area", "Lalpur area", "Kustia area"]
]

# One dict per auto-created group.  roster=True -> driver/vehicle taken from the
# "🤖 Auto Duty Roster" tab (created Approved when both are selected).
DAILY_AUTO_GROUPS = [
    {"tag": DAILY_SHUTTLE_TAG, "templates": DAILY_SHUTTLE_TEMPLATES,
     "applicant": "Staff Shuttle", "department": "Admin", "roster": False,
     "title": "Staff Shuttle"},
    {"tag": DAILY_COMMERCIAL_TAG, "templates": DAILY_COMMERCIAL_TEMPLATES,
     "applicant": "Commercial", "department": "Commercial", "roster": True,
     "title": "💼 Commercial Duty"},
    {"tag": DAILY_NIGHT_SHIFT_TAG, "templates": DAILY_NIGHT_SHIFT_TEMPLATES,
     "applicant": "Night Shift Vehicle", "department": "Commercial", "roster": True,
     "title": "🌙 Night Shift Vehicle"},
    {"tag": DAILY_FEMALE_DROP_TAG, "templates": DAILY_FEMALE_DROP_TEMPLATES,
     "applicant": "Admin", "department": "Admin", "roster": False,
     "title": "Female Worker Drop"},
]


# ---- Roster (driver + vehicle per slot) helpers ----
@st.cache_data(ttl=60, show_spinner=False)
def fetch_auto_roster():
    """Returns ({slot_key: row}, error_text). error_text is '' when OK — it is
    only filled when the roster table does not exist yet / can't be read."""
    try:
        sb = get_supabase_client()
        res = sb.table(AUTO_ROSTER_TABLE).select("*").execute()
        return {r["slot_key"]: r for r in (res.data or [])}, ""
    except Exception as e:
        return {}, str(e)


def _clear_auto_roster_caches():
    fetch_auto_roster.clear()


@st.cache_data(ttl=60, show_spinner=False)
def fetch_auto_duty_trips(date_from: str, date_to: str) -> pd.DataFrame:
    """Auto-created Commercial Duty + Night Shift requisitions (any status) whose
    trip date is between date_from and date_to (inclusive, 'YYYY-MM-DD').
    Small, targeted query — does not load the whole history."""
    sb = get_supabase_client()
    frames = []
    for tag in (DAILY_COMMERCIAL_TAG, DAILY_NIGHT_SHIFT_TAG):
        res = (
            sb.table(REQUISITIONS_TABLE).select("*")
            .gte("date_of_travel", date_from).lte("date_of_travel", date_to)
            .like("admin_note", f"%{tag}%")
            .execute()
        )
        if res.data:
            frames.append(pd.DataFrame(res.data))
    if not frames:
        return pd.DataFrame(columns=REQUISITION_COLUMNS)
    df = pd.concat(frames, ignore_index=True).drop_duplicates(subset="request_id")
    return df.sort_values(["date_of_travel", "time_of_travel"], ascending=[False, True]).reset_index(drop=True)


def save_auto_roster_slots(rows: list):
    """Saves (upserts) many roster slots in ONE request."""
    sb = get_supabase_client()
    sb.table(AUTO_ROSTER_TABLE).upsert(rows, on_conflict="slot_key").execute()
    _clear_auto_roster_caches()


def _slot_marker_found(marker: str, notes: list) -> bool:
    """True if '<TAG>#<n>' appears in any note (and is not just a prefix of '#<n>0')."""
    pat = re.escape(marker) + r"(?!\d)"
    return any(re.search(pat, n) for n in notes)


def _hhmm_to_time(value, fallback="08:00"):
    """'HH:MM' / 'HH:MM:SS' -> datetime.time (falls back safely)."""
    for candidate in (value, fallback):
        if is_blank(candidate):
            continue
        for f in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(str(candidate).strip(), f).time()
            except ValueError:
                continue
    return datetime.strptime("08:00", "%H:%M").time()


def _default_group_trips(g: dict) -> list:
    """The built-in trips of a group (used until Admin edits them in the roster tab)."""
    out = []
    for i, tpl in enumerate(g["templates"], start=1):
        out.append({
            "tid": i,
            "applicant": tpl.get("applicant", g["applicant"]),
            "department": tpl.get("department", g["department"]),
            "destination": tpl["destination"],
            "time": tpl["time"],
            "passengers": int(tpl["passenger_count"]),
            "purpose": tpl["purpose"],
            "vehicle_type": tpl["vehicle_type"],
        })
    return out


def get_group_trips(g: dict, roster: dict):
    """-> (trips, next_tid) for a roster group. Uses the list Admin saved in the
    roster tab (trips_json); falls back to the built-in defaults when nothing
    has been saved yet (or the saved text is unreadable)."""
    defaults = _default_group_trips(g)
    row = roster.get(g["tag"]) or {}
    raw = row.get("trips_json")
    if isinstance(raw, str) and raw.strip():
        try:
            obj = json.loads(raw)
            trips = []
            for d in obj.get("trips", []):
                trips.append({
                    "tid": int(d["tid"]),
                    "applicant": str(d.get("applicant", "")).strip(),
                    "department": str(d.get("department", "")).strip(),
                    "destination": str(d.get("destination", "")).strip(),
                    "time": str(d.get("time", "")).strip()[:5],
                    "passengers": max(1, int(d.get("passengers", 1))),
                    "purpose": str(d.get("purpose", "")).strip(),
                    "vehicle_type": str(d.get("vehicle_type", "HIACE")).strip() or "HIACE",
                })
            next_tid = int(obj.get("next_tid") or (max([x["tid"] for x in trips], default=0) + 1))
            return trips, next_tid
        except Exception as e:
            print(f"Unreadable trips_json for {g['tag']} (using defaults): {e}")
    return defaults, len(defaults) + 1


def save_group_trips(tag: str, trips: list, next_tid: int, updated_by: str):
    """Saves the trip list of one roster group (other roster fields stay as they are)."""
    sb = get_supabase_client()
    sb.table(AUTO_ROSTER_TABLE).upsert({
        "slot_key": tag, "group_tag": tag, "slot_no": 0,
        "trips_json": json.dumps({"next_tid": next_tid, "trips": trips}, ensure_ascii=False),
        "updated_by": updated_by, "updated_at": bd_now_str(),
    }, on_conflict="slot_key").execute()
    _clear_auto_roster_caches()


def parse_trip_time(value):
    """'19:15' / '7:15 PM' -> 'HH:MM' (24h), or None if it can't be read."""
    s = str(value or "").strip().upper().replace(".", "")
    for f in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p", "%I %p"):
        try:
            return datetime.strptime(s, f).strftime("%H:%M")
        except ValueError:
            continue
    return None


def ensure_daily_shuttle_requisitions():
    """Auto-creates today's fixed requisitions (Staff Shuttle, Commercial Duty,
    Night Shift Vehicle, Female Worker Drop) if they don't already exist.
    Only runs once the Bangladesh clock reaches the start time (6:30 PM).

    Roster groups (Commercial / Night Shift): when the group has BOTH a driver and
    a vehicle selected in the "🤖 Auto Duty Roster" tab, every trip of the group
    is created already Approved with that driver + vehicle. They are never
    created as Pending — if no driver/vehicle is selected (or the group is
    switched off) nothing is created for that group.

    Runs at most once per browser session per day and is idempotent via the
    admin_note '<TAG>#<slot>' check. Any failure is logged and swallowed,
    never blocking the rest of the app."""
    now_bd = bd_now()
    if (now_bd.hour, now_bd.minute) < (DAILY_AUTO_START_HOUR, DAILY_AUTO_START_MINUTE):
        # Before 6:30 PM: do nothing, and do NOT mark today as "checked" so
        # the very next load after 6:30 PM still creates them.
        return

    today_str = str(bd_today())
    if st.session_state.get("_daily_shuttle_checked_date") == today_str:
        return
    st.session_state["_daily_shuttle_checked_date"] = today_str

    try:
        sb = get_supabase_client()

        # Roster is optional — if the table doesn't exist yet everything just
        # falls back to Pending requisitions (the old behaviour).
        roster = {}
        try:
            res_r = sb.table(AUTO_ROSTER_TABLE).select("*").execute()
            roster = {r["slot_key"]: r for r in (res_r.data or [])}
        except Exception as e:
            print(f"Auto roster not available (using Pending): {e}")

        contact_map = {}
        if any(r.get("driver_name") for r in roster.values()):
            try:
                res_d = sb.table(DRIVERS_TABLE).select("driver_name, driver_contact").execute()
                contact_map = {d["driver_name"]: (d.get("driver_contact") or "") for d in (res_d.data or [])}
            except Exception as e:
                print(f"Driver contact lookup failed: {e}")

        for g in DAILY_AUTO_GROUPS:
            tag = g["tag"]

            # --- which trips does this group create today? ---
            driver = vehicle = ""
            if g["roster"]:
                # ONE roster row per group: its driver + vehicle apply to every trip,
                # its trip list (editable in the roster tab) says WHAT trips exist.
                slot = roster.get(tag)
                if slot is None or slot.get("enabled") is False:
                    continue  # no roster saved yet / group switched off -> nothing created
                driver = (slot.get("driver_name") or "").strip()
                vehicle = (slot.get("vehicle_number") or "").strip()
                if not (driver and vehicle):
                    continue  # roster groups are NEVER created as Pending
                trips, _next = get_group_trips(g, roster)
                items = [(tr["tid"], {
                    "applicant": tr["applicant"], "department": tr["department"],
                    "destination": tr["destination"], "time": tr["time"],
                    "passenger_count": tr["passengers"], "purpose": tr["purpose"],
                    "vehicle_type": tr["vehicle_type"],
                }) for tr in trips]
            else:
                items = list(enumerate(g["templates"], start=1))
            if not items:
                continue

            existing = (
                sb.table(REQUISITIONS_TABLE)
                .select("admin_note")
                .eq("date_of_travel", today_str)
                .like("admin_note", f"%{tag}%")
                .execute()
            )
            notes = [str(r.get("admin_note") or "") for r in (existing.data or [])]
            # Rows made by the OLD version of this feature carried just the bare
            # tag (no '#n'); count them so they are not created a second time.
            legacy_count = sum(1 for n in notes if n.strip() == tag)

            for slot_id, tpl in items:
                marker = f"{tag}#{slot_id}"
                if _slot_marker_found(marker, notes) or slot_id <= legacy_count:
                    continue

                trip_time = tpl["time"]
                data = {
                    "request_id": generate_request_id(),
                    "username": "",
                    "applicant_name": tpl.get("applicant", g["applicant"]),
                    "department": tpl.get("department", g["department"]),
                    "mobile_number": "",
                    "date_of_travel": today_str,
                    "time_of_travel": trip_time,
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
                    "admin_note": marker,
                }
                if driver and vehicle:
                    data.update({
                        "status": "Approved",
                        "driver_name": driver,
                        "driver_contact": contact_map.get(driver, ""),
                        "vehicle_number": vehicle,
                        "approved_by": "Auto Roster",
                        "action_timestamp": bd_now_str(),
                        "approved_time": trip_time,
                    })
                insert_requisition(data)
    except Exception as e:
        print(f"Daily auto-requisition error: {e}")


# ------------------- AUTO "SUBMIT ALL FIXED TRIPS FOR TOMORROW" (NEW) -------------------
# Does exactly what the Admin's "🌅 Submit ALL Fixed Trips for Tomorrow" button in
# the 🚌 Staff Shuttle tab does — but automatically: the first time ANY device /
# ANY logged-in user loads the app after 9:55 PM (Bangladesh time), every saved
# shuttle template is submitted for TOMORROW.
#   * Template has a valid default Driver + Vehicle -> created Approved and goes
#     straight to the driver.
#   * Template has no (valid) default Driver/Vehicle -> NOT created (these trips
#     are never created as Pending). Save the template with a default Driver +
#     Vehicle and it is submitted automatically.
# It never creates a duplicate: if tomorrow already has a requisition for that
# template (submitted by hand or by this function) the template is skipped.
# Only runs between 9:55 PM and midnight, so after 12:00 AM it can never submit
# for the day-after-tomorrow by mistake.
TOMORROW_SHUTTLE_AUTO_HOUR = 21
TOMORROW_SHUTTLE_AUTO_MINUTE = 55


def ensure_tomorrow_shuttle_submitted():
    now_bd = bd_now()
    if (now_bd.hour, now_bd.minute) < (TOMORROW_SHUTTLE_AUTO_HOUR, TOMORROW_SHUTTLE_AUTO_MINUTE):
        return  # too early (before 9:55 PM) — and after midnight hour is 0, so also a no-op

    tomorrow = bd_today() + timedelta(days=1)
    tomorrow_str = str(tomorrow)
    if st.session_state.get("_tomorrow_shuttle_checked_date") == tomorrow_str:
        return

    try:
        sb = get_supabase_client()
        tpl_res = sb.table(SHUTTLE_TEMPLATES_TABLE).select("*").order("template_name").execute()
        templates = tpl_res.data or []
        if templates:
            drv_res = sb.table(DRIVERS_TABLE).select("driver_name, driver_contact").execute()
            veh_res = sb.table(VEHICLES_TABLE).select("vehicle_number").execute()
            contact_map = {d["driver_name"]: (d.get("driver_contact") or "") for d in (drv_res.data or [])}
            vehicle_set = {v["vehicle_number"] for v in (veh_res.data or [])}

            existing = (
                sb.table(REQUISITIONS_TABLE)
                .select("admin_note")
                .eq("date_of_travel", tomorrow_str)
                .like("admin_note", "%shuttle template:%")
                .execute()
            )
            notes = [str(r.get("admin_note") or "") for r in (existing.data or [])]

            for tpl in templates:
                name = str(tpl.get("template_name") or "").strip()
                if not name:
                    continue
                if any(f"shuttle template: {name}" in n for n in notes):
                    continue  # already submitted for tomorrow (by hand or automatically)

                d_name = fmt(tpl.get("default_driver_name"), "")
                v_number = fmt(tpl.get("default_vehicle_number"), "")
                have_both = d_name in contact_map and v_number in vehicle_set
                if not have_both:
                    continue  # never create a Pending trip — needs a default Driver + Vehicle
                time_val = fmt(tpl.get("default_time"), "08:00")
                data = {
                    "request_id": generate_request_id(),
                    "username": "",
                    "applicant_name": fmt(tpl.get("applicant_name"), name),
                    "department": fmt(tpl.get("department"), "Admin"),
                    "mobile_number": "",
                    "date_of_travel": tomorrow_str,
                    "time_of_travel": time_val,
                    "destination": fmt(tpl.get("destination"), ""),
                    "passenger_count": int(tpl.get("passenger_count") or 1),
                    "vehicle_type": fmt(tpl.get("vehicle_type"), "HIACE"),
                    "purpose": fmt(tpl.get("purpose"), name),
                    "special_request": "",
                    "status": "Approved",
                    "driver_name": d_name,
                    "driver_contact": contact_map.get(d_name, ""),
                    "vehicle_number": v_number,
                    "approved_by": "Auto (Shuttle 9:55 PM)",
                    "admin_note": f"Auto-submitted from shuttle template: {name} (auto 9:55 PM)",
                }
                data["action_timestamp"] = bd_now_str()
                data["approved_time"] = time_val
                insert_requisition(data)
        # Mark as done only after a clean run, so a temporary network error
        # is retried on the next page load instead of being skipped for the day.
        st.session_state["_tomorrow_shuttle_checked_date"] = tomorrow_str
    except Exception as e:
        print(f"Auto tomorrow-shuttle error: {e}")


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
            st.caption("New employees, Drivers, or Medical staff can request an "
                       "account here. An admin must approve your account before you can log in.")

            reg_role = st.radio(
                "Register as", ["Employee", "Driver", "Nurse / Senior Nurse"],
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
    pdf.cell(0, 6, f"Generated on: {bd_now().strftime('%Y-%m-%d %I:%M %p')}",
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

    # Grand total of the Total KM column shown above.
    grand_km = pd.to_numeric(df["total_km"], errors="coerce").sum() if ("total_km" in df.columns and not df.empty) else 0
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 8, f"Total KM: {grand_km:.1f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    return bytes(pdf.output())


# =========================================================
# 5A. EMPLOYEE SELF-SERVICE: "MY REQUISITION HISTORY" EXCEL + PDF
# =========================================================
# Used only by the Employee dashboard's "My Requests / Live Status" tab.
# It receives the DataFrame from fetch_requisitions_by_user(), which is
# already filtered on the server to the logged-in user's own rows, so a
# user can never export anybody else's data. Nothing else in the app calls
# these helpers, so no other screen is affected.
def build_my_history_df(df: pd.DataFrame) -> pd.DataFrame:
    """Clean, human-readable copy of a user's requisitions for export."""
    rows = []
    for _, r in df.iterrows():
        s_km, e_km, t_km = display_km_fields(r)
        approved_time = r.get("approved_time")
        rows.append({
            "Req #": short_req_id(r.get("id")),
            "Requested On": fmt_time_12h(r.get("created_at"), ""),
            "Date of Travel": fmt(r.get("date_of_travel"), ""),
            "Time of Travel": fmt_time_12h(r.get("time_of_travel"), ""),
            "Destination": fmt(r.get("destination"), ""),
            "Passengers": fmt(r.get("passenger_count"), ""),
            "Vehicle Type": fmt(r.get("vehicle_type"), ""),
            "Purpose": fmt(r.get("purpose"), ""),
            "Special Request": fmt(r.get("special_request"), ""),
            "Status": fmt(r.get("status"), ""),
            "Driver": fmt(r.get("driver_name"), ""),
            "Driver Contact": fmt(r.get("driver_contact"), ""),
            "Vehicle No.": fmt(r.get("vehicle_number"), ""),
            "Approved Departure": fmt_time_12h(approved_time, "") if not is_blank(approved_time) else "",
            "Trip Start": fmt_time_12h(r.get("actual_exit_time"), ""),
            "Trip End": fmt_time_12h(r.get("actual_return_time"), ""),
            "Start KM": "" if s_km is None else s_km,
            "End KM": "" if e_km is None else e_km,
            "Total KM": "" if t_km is None else t_km,
            "Admin Note": fmt(r.get("admin_note"), ""),
        })
    return pd.DataFrame(rows)


def _my_history_counts(hist_df: pd.DataFrame) -> dict:
    counts = {"Total": len(hist_df)}
    for status in REQ_STATUS_OPTIONS:
        counts[status] = int((hist_df["Status"] == status).sum()) if not hist_df.empty else 0
    km = pd.to_numeric(hist_df["Total KM"], errors="coerce").sum() if not hist_df.empty else 0
    counts["Total KM"] = round(float(km), 1)
    return counts


def build_my_history_excel(hist_df: pd.DataFrame, user: dict) -> bytes:
    counts = _my_history_counts(hist_df)
    summary_rows = [
        ("Company", COMPANY_NAME), ("Name", user.get("full_name", "")),
        ("Username", user.get("username", "")), ("Department", user.get("department", "")),
        ("Generated on (BD time)", bd_now().strftime("%Y-%m-%d %I:%M %p")),
    ] + [(k, v) for k, v in counts.items()]
    summary_df = pd.DataFrame(summary_rows, columns=["Item", "Value"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        hist_df.to_excel(writer, index=False, sheet_name="My Requisitions")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        _autofit_columns(writer, [("My Requisitions", hist_df), ("Summary", summary_df)])
        # Bold, coloured header row + frozen header, so it looks like a report.
        from openpyxl.styles import Font, PatternFill, Alignment
        for name in ("My Requisitions", "Summary"):
            ws = writer.sheets[name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="0F62FE")
                cell.alignment = Alignment(horizontal="center", vertical="center")
    return buf.getvalue()


class MyHistoryPDF(ReportPDF):
    def header(self):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(15, 98, 254)
        self.cell(0, 9, sanitize_pdf_text(COMPANY_NAME), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_font("Helvetica", "", 10)
        self.set_text_color(90, 90, 90)
        self.cell(0, 6, sanitize_pdf_text(COMPANY_ADDRESS), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_font("Helvetica", "B", 12)
        self.set_text_color(0, 0, 0)
        self.cell(0, 8, "My Vehicle Requisition History", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)


def build_my_history_pdf(hist_df: pd.DataFrame, user: dict) -> bytes:
    counts = _my_history_counts(hist_df)
    pdf = MyHistoryPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, sanitize_pdf_text(
        f"Name: {user.get('full_name', '')}   |   Department: {user.get('department', '')}   |   "
        f"Username: {user.get('username', '')}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 6, f"Generated on: {bd_now().strftime('%Y-%m-%d %I:%M %p')} (BD time)",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(2)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Summary", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", size=9)
    pdf.set_fill_color(230, 230, 230)
    for label in ["Total"] + REQ_STATUS_OPTIONS:
        pdf.cell(38, 7, f"{label}: {counts[label]}", border=1, align="C", fill=True)
    pdf.ln(12)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Requisition Records", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)

    headers = ["Req #", "Date", "Time", "Destination", "Purpose", "Pax", "Vehicle Type",
               "Status", "Driver", "Vehicle No.", "Total KM", "Admin Note"]
    cols = ["Req #", "Date of Travel", "Time of Travel", "Destination", "Purpose", "Passengers",
            "Vehicle Type", "Status", "Driver", "Vehicle No.", "Total KM", "Admin Note"]
    col_widths = [13, 19, 16, 34, 40, 10, 19, 17, 25, 23, 14, 27]

    pdf.set_font("Helvetica", size=7)
    heading_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 98, 254))
    with pdf.table(col_widths=col_widths, text_align="LEFT", first_row_as_headings=True,
                   line_height=5, headings_style=heading_style, cell_fill_color=(245, 245, 245),
                   cell_fill_mode="ROWS") as table:
        header_row = table.row()
        for h in headers:
            header_row.cell(h)
        for _, r in hist_df.iterrows():
            row = table.row()
            for col in cols:
                row.cell(sanitize_pdf_text(r.get(col, "")))

    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 8, f"Total KM: {counts['Total KM']:.1f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
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
    "Start KM", "End KM", "Total KM", "Duty Duration (Hrs)", "Route / Purpose", "Admin Note",
]

# Summary grouping choices offered in the Duty Tracker. "Driver + Date" is the
# default so every row carries its date (who worked how many hours on which day).
DUTY_SUMMARY_GROUP_OPTIONS = [
    "Driver + Date (daily)", "Driver (total)", "Date (all drivers)", "Vehicle (total)",
]


def effective_trip_date(df: pd.DataFrame) -> pd.Series:
    """The date a trip REALLY belongs to: the date of the trip's actual
    start (actual_exit_time) when the trip has started, otherwise the requested
    date_of_travel. So a requisition made for 4 Oct that actually left on 7 Oct
    night is reported on 7 Oct. Returns midnight timestamps (naive). Read-only."""
    base = pd.to_datetime(df["date_of_travel"], errors="coerce")
    if "actual_exit_time" not in df.columns:
        return base
    ex = pd.to_datetime(df["actual_exit_time"], errors="coerce", utc=True, format="mixed")
    ex = ex.dt.tz_localize(None).dt.normalize()
    return ex.astype("datetime64[ns]").fillna(base.astype("datetime64[ns]"))


def assign_duty_dates(df: pd.DataFrame, gap_hours: float = 6.0, max_shift_hours: float = 18.0) -> pd.DataFrame:
    """Adds a `_duty_date` column (YYYY-MM-DD) = the date the driver's DUTY
    SESSION started, so a night shift is reported as ONE duty on ONE date
    instead of being split across two calendar dates.

    A session = a run of trips by the same driver where the next trip starts
    less than `gap_hours` after the previous one ended. Example (Siam):
    7 Oct 6:30 PM -> 8 Oct 4:00 AM commercial duty, then the morning staff drop
    ending 9:00 AM = one session dated 7 Oct. He rests at the garage, comes back
    the evening of 8 Oct = a NEW session dated 8 Oct. Needs `_start_dt` / `_end_dt`.
    `max_shift_hours` is a safety cap: a trip starting more than this many hours
    after the session began always starts a NEW session, so one stuck/forgotten
    trip (e.g. ended a day late) can never glue several days into one duty.
    Read-only: nothing is written anywhere."""
    out = df.copy()
    if out.empty:
        out["_duty_date"] = pd.Series(dtype=str)
        return out
    key = out["driver_name"].map(_normalize_driver_name)
    if "vehicle_number" in out.columns:
        veh = out["vehicle_number"].fillna("").astype(str).str.strip().str.lower()
        key = key.where(key.astype(str).str.strip() != "", veh)
    out["_skey"] = key
    out = out.sort_values(["_skey", "_start_dt"], kind="stable")
    gap = pd.Timedelta(hours=float(gap_hours))
    cap = pd.Timedelta(hours=float(max_shift_hours))
    duty_dates = {}
    for _, grp in out.groupby("_skey", sort=False):
        sess_start, running_end = None, None
        for idx, st_, en_ in zip(grp.index, grp["_start_dt"], grp["_end_dt"]):
            if sess_start is None or (st_ - running_end) > gap or (st_ - sess_start) > cap:
                sess_start, running_end = st_, en_
            else:
                running_end = max(running_end, en_)
            duty_dates[idx] = sess_start.strftime("%Y-%m-%d")
    out["_duty_date"] = pd.Series(duty_dates)
    return out.drop(columns=["_skey"]).sort_index()


def duty_span_hours(starts, ends, idle_gap_hours: float = 6.0) -> float:
    """Duty time (TWH) of ONE driver (or vehicle) for ONE day.

    = first trip start -> last trip end, but any break with NO trip for MORE than
    `idle_gap_hours` is cut out and NOT counted. Example: a 7:00-9:00 AM trip and
    a 6:30 PM-8:30 AM night duty on the same date is 2 h + 14 h = 16 h, not the
    25.5 h you get by counting the whole 9.5 h rest in between. Shorter breaks
    between trips (<= idle_gap_hours) still count as duty time, as before.
    Read-only helper; `starts` / `ends` are datetime sequences."""
    pairs = sorted(
        ((s_, e_) for s_, e_ in zip(starts, ends) if pd.notna(s_) and pd.notna(e_)),
        key=lambda p: p[0],
    )
    gap = pd.Timedelta(hours=float(idle_gap_hours))
    total = pd.Timedelta(0)
    blk_s = blk_e = None
    for s_, e_ in pairs:
        if blk_s is None:
            blk_s, blk_e = s_, e_
        elif (s_ - blk_e) > gap:
            total += blk_e - blk_s
            blk_s, blk_e = s_, e_
        else:
            blk_e = max(blk_e, e_)
    if blk_s is not None:
        total += blk_e - blk_s
    return total.total_seconds() / 3600.0


def build_duty_summary(df: pd.DataFrame, group_option: str, idle_gap_hours: float = 6.0) -> pd.DataFrame:
    """Roll the filtered duty rows up into a summary table.

    `df` is the already-filtered duty DataFrame from the Duty Tracker tab; it
    must carry the helper columns _start_dt, _end_dt, _km, _duration_hrs plus
    driver_name / vehicle_number. Every grouping also reports the First Start
    and Last End time (12-hour AM/PM) so the summary always shows *when* the
    duty happened, not just how long it was. The "Date" is the calendar date
    the trip started on. Read-only helper: nothing is written anywhere.
    """
    key_map = {
        "Driver + Date (daily)": ["Date", "Driver Name"],
        "Driver (total)": ["Driver Name"],
        "Date (all drivers)": ["Date"],
        "Vehicle (total)": ["Vehicle No"],
    }
    keys = key_map.get(group_option, ["Date", "Driver Name"])
    out_cols = keys + ["Trips", "Total KM", "Run Hours", "Total Duty Hours", "Total Duty Time (h m)",
                       "First Start", "Last End"]
    if df is None or df.empty:
        return pd.DataFrame(columns=out_cols)

    work = df.copy()
    work["Date"] = work["_duty_date"] if "_duty_date" in work.columns else work["_start_dt"].dt.strftime("%Y-%m-%d")
    work["Driver Name"] = work["driver_name"].map(lambda v: fmt(v, "—"))
    work["Vehicle No"] = work["vehicle_number"].map(lambda v: fmt(v, "—"))

    g = (
        work.groupby(keys)
        .agg(
            Trips=("_km", "size"),
            km_sum=("_km", "sum"),
            hrs_sum=("_duration_hrs", "sum"),
            first_start=("_start_dt", "min"),
            last_end=("_end_dt", "max"),
        )
        .reset_index()
    )
    # DUTY TIME = first trip start -> last trip end, per driver (or vehicle)
    # per day, so short gaps between trips are INCLUDED (e.g. 6:50 AM -> 9:00 PM),
    # but a break with no trip for more than `idle_gap_hours` is NOT counted.
    # "Run Hours" stays the plain sum of each trip's own Start->End time.
    entity = "Vehicle No" if "Vehicle No" in keys else "Driver Name"
    day = pd.DataFrame(
        [(d_, e_, duty_span_hours(grp["_start_dt"], grp["_end_dt"], idle_gap_hours))
         for (d_, e_), grp in work.groupby(["Date", entity])],
        columns=["Date", entity, "duty_hrs"],
    )
    duty = day.groupby(keys)["duty_hrs"].sum().reset_index()
    g = g.merge(duty, on=keys, how="left")
    g["duty_hrs"] = g["duty_hrs"].fillna(0.0)

    g["Total KM"] = g["km_sum"].round(1)
    g["Run Hours"] = g["hrs_sum"].round(2)
    g["Total Duty Hours"] = g["duty_hrs"].round(2)
    g["Total Duty Time (h m)"] = g["duty_hrs"].map(
        lambda h: f"{int(round(h * 60)) // 60}h {int(round(h * 60)) % 60}m"
    )
    g["First Start"] = g["first_start"].dt.strftime("%d-%b-%Y %I:%M %p").map(drop_hour_zero)
    g["Last End"] = g["last_end"].dt.strftime("%d-%b-%Y %I:%M %p").map(drop_hour_zero)

    if "Date" in keys:
        g = g.sort_values(["Date", "duty_hrs"], ascending=[True, False])
    else:
        g = g.sort_values("duty_hrs", ascending=False)
    return g[out_cols].reset_index(drop=True)


def build_suspicious_trips(df: pd.DataFrame, min_hrs: float = 3.0, max_kmh: float = 10.0,
                           always_hrs: float = 6.0) -> pd.DataFrame:
    """Trips worth a second look, one row per physical trip (requisitions
    bundled in a group trip appear once, with all their Req #s):
      * EVERY trip that ran longer than `always_hrs` hours, whatever its KM, and
      * shorter trips (at least `min_hrs` hours) that covered very FEW km (or no
        km at all) — the classic sign of a forgotten / late "End Trip".
    Needs the Duty Tracker helper columns _start_dt, _end_dt, _km. Read-only."""
    cols = ["Req #", "Driver", "Vehicle", "Start", "End", "Hours", "KM", "Avg km/h", "Why flagged"]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    rv = df.copy()
    rv["_km_num"] = pd.to_numeric(rv["_km"], errors="coerce")
    g = rv["trip_group_id"] if "trip_group_id" in rv.columns else pd.Series(None, index=rv.index, dtype=object)
    has = g.notna() & (g.astype(str).str.strip() != "")
    solo = pd.Series(["row-" + str(i) for i in range(len(rv))], index=rv.index)
    rv["_tkey"] = g.astype(str).where(has, solo)
    trips = rv.groupby("_tkey", sort=False).agg(
        start=("_start_dt", "min"), end=("_end_dt", "max"), km=("_km_num", "max"),
        driver=("driver_name", "first"), vehicle=("vehicle_number", "first"),
        ids=("id", lambda s_: ", ".join(short_req_id(v) for v in s_)),
        still_open=("status", lambda s_: bool((s_ == "On Trip").any())),
    ).reset_index(drop=True)
    trips["hrs"] = (trips["end"] - trips["start"]).dt.total_seconds() / 3600.0
    trips["kmh"] = trips["km"] / trips["hrs"].where(trips["hrs"] > 0)
    slow = trips["km"].isna() | (trips["km"] <= 0) | (trips["kmh"] < float(max_kmh))
    flagged = trips[
        (trips["hrs"] >= float(always_hrs)) | ((trips["hrs"] >= float(min_hrs)) & slow)
    ].copy()
    if flagged.empty:
        return pd.DataFrame(columns=cols)

    def _why(r):
        if r["still_open"]:
            return "Still On Trip — End Trip not recorded yet"
        if pd.isna(r["km"]) or r["km"] <= 0:
            return "No KM recorded for a long trip"
        if r["km"] / r["hrs"] < float(max_kmh):
            return f"Only {r['km']:.0f} km in {r['hrs']:.1f} h — maybe End Trip was done late"
        return f"Trip ran {r['hrs']:.1f} h (longer than {float(always_hrs):g} h) — please check"

    flagged = flagged.sort_values("hrs", ascending=False)
    return pd.DataFrame({
        "Req #": flagged["ids"],
        "Driver": flagged["driver"].map(lambda v: fmt(v, "—")),
        "Vehicle": flagged["vehicle"].map(lambda v: fmt(v, "—")),
        "Start": flagged["start"].dt.strftime("%d-%b %I:%M %p").map(drop_hour_zero),
        "End": flagged["end"].dt.strftime("%d-%b %I:%M %p").map(drop_hour_zero),
        "Hours": flagged["hrs"].round(1),
        "KM": flagged["km"].round(1),
        "Avg km/h": flagged["kmh"].round(1),
        "Why flagged": flagged.apply(_why, axis=1),
    }).reset_index(drop=True)


def build_long_duty_list(df: pd.DataFrame, limit_hrs: float = 12.0, idle_gap_hours: float = 6.0) -> pd.DataFrame:
    """Driver-days whose duty time (first trip start -> last trip end without
    long idle breaks, exactly the 'Total Duty Hours' of the Driver + Date summary) is MORE than
    `limit_hrs`, with the Req #s of that duty and its longest trip so the cause
    is easy to find. Read-only."""
    cols = ["Date", "Driver Name", "Total Duty Time (h m)", "Total Duty Hours", "Trips", "Total KM",
            "First Start", "Last End", "Longest trip", "Requisitions"]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    summ = build_duty_summary(df, "Driver + Date (daily)", idle_gap_hours)
    over = summ[summ["Total Duty Hours"] > float(limit_hrs)].copy()
    if over.empty:
        return pd.DataFrame(columns=cols)

    w = df.copy()
    w["Date"] = w["_duty_date"] if "_duty_date" in w.columns else w["_start_dt"].dt.strftime("%Y-%m-%d")
    w["Driver Name"] = w["driver_name"].map(lambda v: fmt(v, "—"))
    w["_h"] = (w["_end_dt"] - w["_start_dt"]).dt.total_seconds() / 3600.0
    w = w.sort_values("_start_dt")
    reqs, longest = {}, {}
    for key, grp in w.groupby(["Date", "Driver Name"]):
        reqs[key] = ", ".join(short_req_id(v) for v in grp["id"])
        top = grp.loc[grp["_h"].idxmax()]
        longest[key] = f"{short_req_id(top['id'])} ({top['_h']:.1f} h)"
    keys = list(zip(over["Date"], over["Driver Name"]))
    over["Requisitions"] = [reqs.get(k, "") for k in keys]
    over["Longest trip"] = [longest.get(k, "") for k in keys]
    return over[cols].sort_values("Total Duty Hours", ascending=False).reset_index(drop=True)


MATRIX_METRICS = ["TWH", "SWH", "Km"]
MATRIX_SUMMARY_TOPS = ("Total", "Average")


def build_driver_daily_matrix(df: pd.DataFrame, all_driver_names=None, idle_gap_hours: float = 6.0) -> pd.DataFrame:
    """Driver x Date table in the sheet format: for every date three columns
    TWH | SWH | Km, then a 'Total' block and an 'Average' block (per-day
    average = total / number of date columns), and a GRAND TOTAL row.

      TWH = Total Working Hours = DUTY time (first trip start -> last trip end
            that day; short gaps between trips included, but a break with no
            trip for more than `idle_gap_hours` is NOT counted).
      SWH = trip running hours (sum of each trip's own Start -> End time).
      Km  = distance run that day (driver KM first, older start/end KM as fallback).

    Returns a DataFrame with a 2-level column index (date label, metric) and
    the driver name as the index. Drivers from the master list are included
    even with no trips (blank row). Date = the day the trip STARTED."""
    work = df.copy() if df is not None else pd.DataFrame()
    display = {}
    for n in (all_driver_names or []):
        display.setdefault(_normalize_driver_name(n), n)

    days, g = [], pd.DataFrame()
    if not work.empty:
        work["_dkey"] = work["driver_name"].map(_normalize_driver_name)
        work["_day_sort"] = work["_duty_date"] if "_duty_date" in work.columns else work["_start_dt"].dt.strftime("%Y-%m-%d")
        g = work.groupby(["_dkey", "_day_sort"]).agg(
            _s=("_start_dt", "min"), _e=("_end_dt", "max"),
            swh=("_duration_hrs", "sum"), km=("_km", "sum"),
        ).reset_index()
        _twh = {
            k: duty_span_hours(grp["_start_dt"], grp["_end_dt"], idle_gap_hours)
            for k, grp in work.groupby(["_dkey", "_day_sort"])
        }
        g["twh"] = [_twh[(a, b)] for a, b in zip(g["_dkey"], g["_day_sort"])]
        days = sorted(g["_day_sort"].unique())
        for k, n in work.groupby("_dkey")["driver_name"].first().items():
            display.setdefault(k, fmt(n, "—"))

    if not display:
        return pd.DataFrame()

    keys = list(display.keys())
    n_days = max(len(days), 1)
    cols = {}
    for d in days:
        label = datetime.strptime(d, "%Y-%m-%d").strftime("%d/%m/%y")
        sub = g[g["_day_sort"] == d].set_index("_dkey")
        cols[(label, "TWH")] = sub["twh"].reindex(keys).round(1)
        cols[(label, "SWH")] = sub["swh"].reindex(keys).round(1)
        cols[(label, "Km")] = sub["km"].reindex(keys).round(1)

    if not g.empty:
        tot = g.groupby("_dkey")[["twh", "swh", "km"]].sum().reindex(keys).fillna(0.0)
    else:
        tot = pd.DataFrame(0.0, index=keys, columns=["twh", "swh", "km"])
    for m, c in zip(MATRIX_METRICS, ["twh", "swh", "km"]):
        cols[("Total", m)] = tot[c].round(1)
    for m, c in zip(MATRIX_METRICS, ["twh", "swh", "km"]):
        cols[("Average", m)] = (tot[c] / n_days).round(1)

    out = pd.DataFrame(cols, index=keys)
    out.columns = pd.MultiIndex.from_tuples(list(cols.keys()))
    out.index = [display[k] for k in keys]
    out = out.sort_index(key=lambda ix: ix.str.lower())
    out.index.name = "Driver Name"
    grand = out.sum(min_count=1).round(1)  # sums are linear, so Average column sums stay correct
    out.loc["GRAND TOTAL"] = grand
    return out


def _matrix_blocks(matrix_df: pd.DataFrame, dates_per_block: int = 6):
    """Split the matrix's top-level labels into PDF-friendly blocks: groups of
    `dates_per_block` dates, then one block with Total + Average."""
    tops = list(dict.fromkeys(matrix_df.columns.get_level_values(0)))
    date_tops = [t for t in tops if t not in MATRIX_SUMMARY_TOPS]
    blocks = [date_tops[i:i + dates_per_block] for i in range(0, len(date_tops), dates_per_block)]
    summary = [t for t in tops if t in MATRIX_SUMMARY_TOPS]
    if summary:
        blocks.append(summary)
    return blocks


def _matrix_num(v) -> str:
    """Blank for missing, '187' for whole numbers, '9.9' otherwise."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    v = float(v)
    return str(int(v)) if v == int(v) else f"{v:.1f}"


def _write_matrix_sheet(wb, matrix_df: pd.DataFrame, position: int = 1):
    """Writes the Driver KM Matrix sheet by hand so it looks like the sheet
    format: merged date headers over TWH | SWH | Km, borders, centered cells,
    bold GRAND TOTAL row."""
    from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
    from openpyxl.utils import get_column_letter
    ws = wb.create_sheet("Driver KM Matrix", position)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    head_fill = PatternFill("solid", fgColor="D9D9D9")

    cols = list(matrix_df.columns)
    ws.cell(1, 1, "Driver Name")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    for j, (top, metric) in enumerate(cols, start=2):
        ws.cell(2, j, metric)
        if metric == MATRIX_METRICS[0]:
            ws.cell(1, j, top)
            ws.merge_cells(start_row=1, start_column=j, end_row=1, end_column=j + len(MATRIX_METRICS) - 1)
    for ri, (name, row) in enumerate(matrix_df.iterrows(), start=3):
        ws.cell(ri, 1, name)
        for j, v in enumerate(row.tolist(), start=2):
            ws.cell(ri, j, None if pd.isna(v) else float(v))

    last_row, last_col = 2 + len(matrix_df), 1 + len(cols)
    for r in range(1, last_row + 1):
        for c in range(1, last_col + 1):
            cell = ws.cell(r, c)
            cell.border = border
            cell.alignment = center
            if r <= 2:
                cell.fill = head_fill
                cell.font = Font(bold=True)
            elif r == last_row:
                cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 18
    for c in range(2, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 8
    ws.freeze_panes = "B3"


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


def _autofit_columns(writer, sheets):
    """Light auto-fit so columns aren't clipped in Excel — purely cosmetic.
    `fillna("")` BEFORE `.astype(str)` matters on pandas >= 3.0: that version
    stopped converting NaN/None to the literal string "nan" on astype(str),
    leaving real NaN behind instead, which makes .str.len().max() return NaN
    and int(NaN) raise. Filling blanks with "" first guarantees real integers."""
    from openpyxl.utils import get_column_letter
    for sheet_name, sheet_df in sheets:
        ws = writer.sheets[sheet_name]
        for i, col in enumerate(sheet_df.columns, start=1):
            width = max(12, min(40, int(sheet_df[col].fillna("").astype(str).str.len().max() if not sheet_df.empty else 12) + 2))
            ws.column_dimensions[get_column_letter(i)].width = width


def build_duty_tracker_excel(detail_df: pd.DataFrame, summary_metrics: dict,
                              driver_summary_df: pd.DataFrame = None,
                              matrix_df: pd.DataFrame = None) -> bytes:
    """Formatted .xlsx export for the Duty Tracker: a 'Summary' sheet with the
    KPI cards' values, an optional 'Duty Summary' sheet, an optional 'Driver
    KM Matrix' sheet (driver x date Run KM), plus a 'Duty Log' sheet with the
    full filtered detail table. The optional args default to None so any
    older caller keeps working unchanged.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary_df = pd.DataFrame(
            [{"Metric": k, "Value": v} for k, v in summary_metrics.items()]
        )
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        sheets = [("Summary", summary_df)]
        if driver_summary_df is not None:
            driver_summary_df.to_excel(writer, index=False, sheet_name="Duty Summary")
            sheets.append(("Duty Summary", driver_summary_df))
        detail_df.to_excel(writer, index=False, sheet_name="Duty Log")
        sheets.append(("Duty Log", detail_df))

        _autofit_columns(writer, sheets)

        # Driver x Date matrix (TWH | SWH | Km) written last, then placed as the 2nd sheet.
        if matrix_df is not None and not matrix_df.empty:
            _write_matrix_sheet(writer.book, matrix_df, position=1)

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


def build_duty_tracker_pdf(detail_df: pd.DataFrame, summary_metrics: dict, filters_summary: str,
                            driver_summary_df: pd.DataFrame = None,
                            matrix_df: pd.DataFrame = None) -> bytes:
    """PDF containing the KPI summary table, an optional duty summary, an
    optional Driver-wise Daily Run KM matrix, followed by the detailed duty
    log and its Total KM. `detail_df` must already have the
    DUTY_TRACKER_DISPLAY_COLS columns. The optional args default to None so
    older callers keep working.

    Every string written to the PDF is passed through sanitize_pdf_text()
    first — see that function's docstring for why this is necessary with
    FPDF's core Helvetica font.
    """
    pdf = DutyTrackerPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, f"Generated on: {bd_now().strftime('%Y-%m-%d %I:%M %p')}",
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

    # ---- Duty summary (grouped by Driver+Date / Driver / Date / Vehicle) ----
    if driver_summary_df is not None and not driver_summary_df.empty:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 8, "Duty Summary", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(1)
        pdf.set_font("Helvetica", size=8)
        drv_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 98, 254))
        summary_cols = list(driver_summary_df.columns)
        equal_w = round(277 / len(summary_cols), 1)
        with pdf.table(col_widths=[equal_w] * len(summary_cols), text_align="LEFT",
                       first_row_as_headings=True, line_height=6, headings_style=drv_style,
                       cell_fill_color=(245, 245, 245), cell_fill_mode="ROWS") as drv_table:
            drv_header = drv_table.row()
            for h in summary_cols:
                drv_header.cell(sanitize_pdf_text(h))
            for _, dr in driver_summary_df.iterrows():
                drow = drv_table.row()
                for col in summary_cols:
                    drow.cell(sanitize_pdf_text(fmt(dr.get(col, ""), "")))
        pdf.ln(6)

    # ---- Driver-wise Daily Duty & Run KM (TWH | SWH | Km per date) ----
    if matrix_df is not None and not matrix_df.empty:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 8, "Driver-wise Daily Duty Hours & Run KM", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "I", 8)
        pdf.cell(0, 5, "TWH = total working (duty) hours, first start to last end, long idle breaks excluded  |  "
                       "SWH = trip running hours  |  Km = distance run",
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(1)
        hstyle = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 98, 254))
        bold = FontFace(emphasis="BOLD")
        first_w = 34
        for tops in _matrix_blocks(matrix_df):
            ncols = len(tops) * len(MATRIX_METRICS)
            w = min(round((277 - first_w) / ncols, 2), 18)
            pdf.set_font("Helvetica", size=8)
            with pdf.table(col_widths=[first_w] + [w] * ncols, text_align="CENTER",
                           first_row_as_headings=False, line_height=6,
                           cell_fill_color=(245, 245, 245), cell_fill_mode="ROWS") as mt:
                r1 = mt.row()
                r1.cell("", style=hstyle)
                for t in tops:
                    r1.cell(sanitize_pdf_text(t), colspan=len(MATRIX_METRICS), style=hstyle)
                r2 = mt.row()
                r2.cell("Driver Name", style=hstyle)
                for _ in tops:
                    for m in MATRIX_METRICS:
                        r2.cell(m, style=hstyle)
                for name, mr in matrix_df.iterrows():
                    is_total = (name == "GRAND TOTAL")
                    mrow = mt.row()
                    mrow.cell(sanitize_pdf_text(name), style=bold if is_total else None)
                    for t in tops:
                        for m in MATRIX_METRICS:
                            mrow.cell(_matrix_num(mr[(t, m)]), style=bold if is_total else None)
            pdf.ln(4)
        pdf.ln(2)

    # ---- Detailed duty log table ----
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Detailed Duty Log", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)

    headers = DUTY_TRACKER_DISPLAY_COLS
    # 10 columns: Vehicle, Driver, Start Time, End Time, Start KM, End KM,
    # Total KM, Duty Hrs, Route/Purpose, Admin Note — sums to 277mm, fits A4 landscape.
    col_widths = [26, 26, 28, 28, 14, 14, 14, 18, 55, 54]

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

    total_km_val = pd.to_numeric(detail_df.get("Total KM"), errors="coerce").sum() if not detail_df.empty else 0
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 8, f"Total KM (all rows above): {total_km_val:.1f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

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

        _autofit_columns(writer, sheets)

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
    pdf.cell(0, 6, f"Generated on: {bd_now().strftime('%Y-%m-%d %I:%M %p')}",
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
    today = bd_today()
    st.markdown(f"##### 📅 Today's Live Snapshot — {today.strftime('%A, %d %B %Y')}")
    st.caption("These counts are live and independent of the date-range filter below.")

    created_dt = pd.to_datetime(df_all.get("created_at"), errors="coerce", utc=True, format="mixed")
    return_dt = pd.to_datetime(df_all.get("actual_return_time"), errors="coerce", utc=True, format="mixed")
    # created_at and actual_return_time are both saved by the app in
    # Bangladesh time (see insert_requisition / bd_now_str), so "today" is
    # simply today's Bangladesh date on the same clock.
    today_stored = bd_now_ts().normalize()

    pending_now = int((df_all["status"] == "Pending").sum())
    on_trip_now = int((df_all["status"] == "On Trip").sum())
    requested_today = int((created_dt.dt.normalize() == today_stored).sum())
    completed_today = int((return_dt.dt.normalize() == today_stored).sum())

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
                    disp["actual_exit_time"] = disp["actual_exit_time"].apply(fmt_clock_12h)
                    disp = disp.rename(columns={"actual_exit_time": "Start Time"})
                st.dataframe(disp, use_container_width=True, hide_index=True,
                             height=min(320, 45 + 35 * len(disp)))

    with t3:
        st.metric("📥 Requests Submitted Today", requested_today)
        with st.expander(f"🔍 View {requested_today} submitted today"):
            submitted_today_rows = df_all[created_dt.dt.normalize() == today_stored]
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
                    disp["created_at"] = disp["created_at"].apply(fmt_clock_12h)
                    disp = disp.rename(columns={"created_at": "Submitted At"})
                st.dataframe(disp, use_container_width=True, hide_index=True,
                             height=min(320, 45 + 35 * len(disp)))

    with t4:
        st.metric("✅ Trips Completed Today", completed_today)
        with st.expander(f"🔍 View {completed_today} completed today"):
            completed_today_rows = df_all[return_dt.dt.normalize() == today_stored]
            if completed_today_rows.empty:
                st.caption("No trips completed yet today.")
            else:
                disp = completed_today_rows.copy()
                disp["Total KM"] = disp.apply(lambda row: effective_km_fields(row)[2], axis=1)
                cols = ["id", "applicant_name", "driver_name", "vehicle_number",
                        "destination", "actual_exit_time", "actual_return_time", "Total KM"]
                disp = disp[[c for c in cols if c in disp.columns]].copy()
                if "id" in disp.columns:
                    disp["id"] = disp["id"].apply(short_req_id)
                    disp = disp.rename(columns={"id": "Req #"})
                # Start Time (when the trip left) and End Time (when it came
                # back), shown as e.g. "1:00 PM" — a trip that started on an
                # earlier day shows its date too, e.g. "27-Sep 11:30 PM".
                if "actual_exit_time" in disp.columns:
                    disp["actual_exit_time"] = disp["actual_exit_time"].apply(fmt_clock_12h)
                if "actual_return_time" in disp.columns:
                    disp["actual_return_time"] = disp["actual_return_time"].apply(fmt_clock_12h)
                disp = disp.rename(columns={"actual_exit_time": "Start Time",
                                            "actual_return_time": "End Time"})
                st.dataframe(disp, use_container_width=True, hide_index=True,
                             height=min(320, 45 + 35 * len(disp)))

    work_df = df_all.copy()
    work_df["_dt"] = effective_trip_date(work_df)
    min_d = work_df["_dt"].min()
    max_d = work_df["_dt"].max()
    default_start = min_d.date() if pd.notnull(min_d) else bd_today()
    default_end = max_d.date() if pd.notnull(max_d) else bd_today()

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
            float(completed_rows.apply(lambda row: effective_km_fields(row)[2], axis=1)
                  .where(~mark_group_duplicates(completed_rows), 0.0).sum())
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

    filtered_display = humanize_timestamp_columns(filtered.drop(columns=["_dt"], errors="ignore"))
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
    #
    # IMPORTANT: on many versions of the cookie component the very first run
    # returns an EMPTY dict {} (its default), not None, while the browser's real
    # cookies are still on the way. Treating that {} as "no cookie" is what made
    # the login page flash for a moment and then jump to the first page by itself.
    # So an empty result is treated as "not ready yet" too, for the first few
    # short retries only (a truly new/logged-out browser waits ~1.2 s once, then
    # sees the login form as usual).
    COOKIE_BOOTSTRAP_MAX_RETRIES = 4
    cookie_retries = st.session_state.get("_cookie_bootstrap_retries", 0)

    if (not all_cookies) and cookie_retries < COOKIE_BOOTSTRAP_MAX_RETRIES:
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
ensure_daily_shuttle_requisitions()  # auto-creates today's shuttle / commercial / night-shift / female-drop requisitions after 6:30 PM (BD time), once per session/day (roster groups come out Approved)
ensure_tomorrow_shuttle_submitted()  # after 9:55 PM (BD time): auto-submits ALL saved Staff Shuttle templates for tomorrow (same as the "Submit ALL Fixed Trips for Tomorrow" button), once per session/day

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

auto_refresh_on = st.sidebar.checkbox("🔄 Auto-refresh every 90s", value=True,
                                       help="Automatically reloads live data across the app every 90 seconds. "
                                            "Turn off temporarily if you're filling out a long form, "
                                            "or use '🔄 Refresh Now' any time for an on-demand update.")
if st.sidebar.button("🔄 Refresh Now", use_container_width=True):
    st.rerun()
st.sidebar.markdown("---")
logout_button()

# Auto-refresh applies to every role. Everyone can still hit "🔄 Refresh Now"
# above for an on-demand update, and every write path already clears its matching
# cache immediately, so approvals / driver KM entries show up instantly for the
# person who made the change; the timer below only controls how often *other*
# idle screens passively refresh to see someone else's changes.
NO_AUTOREFRESH_ROLES = set()   # no role is exempt
# Every auto-refresh re-sends the whole page, which is slow on weak mobile data,
# so ALL roles refresh every 120 seconds (the sidebar's "Refresh Now" button is
# always there for an instant manual refresh).
AUTOREFRESH_INTERVAL_MS = 120_000
if auto_refresh_on and user["role"] not in NO_AUTOREFRESH_ROLES:
    st_autorefresh(interval=AUTOREFRESH_INTERVAL_MS, key="global_autorefresh")


def section_nav(key: str, labels: list):
    """Drop-in replacement for st.tabs() that REMEMBERS the selected section.
    st.tabs() always jumps back to the first tab whenever the script reruns
    (auto-refresh, st.rerun(), cache refresh). This uses a keyed st.radio, so
    the choice lives in st.session_state and survives every rerun. It returns
    one True/False per label, so the old `with tab_x:` blocks only needed to
    become `if tab_x:`. Bonus: only the selected section's code runs on each
    rerun (st.tabs ran ALL tabs every time), so the app is also faster."""
    if st.session_state.get(key) not in labels:
        st.session_state[key] = labels[0]
    choice = st.radio("Section", labels, horizontal=True, key=key, label_visibility="collapsed")
    return tuple(choice == l for l in labels)

# =========================================================
# 8. EMPLOYEE DASHBOARD
# =========================================================
if user["role"] == "user":
    company_header("👤 Employee Dashboard")
    tab1, tab2 = section_nav("emp_section", ["📋 New Requisition", "📍 My Requests / Live Status"])

    if tab1:
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
                date_of_travel = st.date_input("Date of Travel *", min_value=bd_today())
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

    if tab2:
        st.subheader("My Requests — Live Status")
        with st.spinner("Loading your requests..."):
            my_df = fetch_requisitions_by_user(user["username"])
        if my_df.empty:
            st.info("You haven't submitted any requisitions yet.")
        else:
            # ---- Download my complete history (Excel + PDF) — own data only ----
            st.markdown("##### 📥 Download My Request History")
            st.caption("Your complete requisition history (all dates, all statuses) — only your own requests.")
            try:
                my_hist_df = build_my_history_df(my_df)
                _safe_user = re.sub(r"[^A-Za-z0-9_-]", "_", str(user["username"]))
                _stamp = bd_now().strftime("%Y%m%d_%H%M")
                hd1, hd2 = st.columns(2)
                with hd1:
                    st.download_button(
                        "⬇️ Download Excel (.xlsx)", data=build_my_history_excel(my_hist_df, user),
                        file_name=f"my_requisition_history_{_safe_user}_{_stamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="my_hist_xlsx",
                    )
                with hd2:
                    st.download_button(
                        "⬇️ Download PDF (.pdf)", data=build_my_history_pdf(my_hist_df, user),
                        file_name=f"my_requisition_history_{_safe_user}_{_stamp}.pdf",
                        mime="application/pdf", use_container_width=True, key="my_hist_pdf",
                    )
            except Exception as e:
                st.warning(f"Could not prepare the download files right now: {e}")
            st.markdown("---")

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

                        _s_km, _e_km, _t_km = display_km_fields(r)
                        if r["status"] in ("On Trip", "Completed"):
                            st.write(f"**Trip Started:** {fmt_time_12h(r.get('actual_exit_time'))}  |  **Start KM:** {km_plain(_s_km) or '—'}")
                        if r["status"] == "Completed":
                            st.write(f"**Trip Ended:** {fmt_time_12h(r.get('actual_return_time'))}  |  **End KM:** {km_plain(_e_km) or '—'}  |  **Total KM:** {km_plain(_t_km) if _t_km is not None else '—'}")

                        if r["status"] == "Rejected":
                            st.error("This request was rejected by the admin.")
                            if not is_blank(r.get("admin_note")):
                                st.caption(f"📝 Reason: {r['admin_note']}")

# =========================================================
# 9B. DRIVER DASHBOARD — My Assigned Trips & Odometer Entry (NEW)
# =========================================================
# A two-tab design: a "Start Trip" tab that only takes Start KM, and a
# separate "End Trip" tab that only takes End KM once the driver is back —
# instead of one combined form.
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

    tab_depart, tab_return = section_nav("driver_section", ["🚦 Start Trip (Approved)", "🏁 End Trip (Return)"])

    with st.spinner("Loading your assigned trips..."):
        my_trips = fetch_requisitions_by_driver(user["full_name"])

    # ---------------- TAB 1: Start Trip (Start KM only) ----------------
    if tab_depart:
        st.subheader("Approved Trips Awaiting Your Start KM")
        st.caption("🔢 Trips are listed in the order the Admin approved them — #1 was approved first, then #2, #3 ...")
        start_trips = my_trips[my_trips["status"] == "Approved"] if not my_trips.empty else my_trips

        if start_trips.empty:
            st.info("You have no Approved trips waiting to start.")
        else:
            for n, grp in enumerate(group_trips(sort_by_approval_order(start_trips)), start=1):
                first = grp[0]
                gkey = first["request_id"]
                multi = len(grp) > 1
                auto_start = get_last_driver_end_km(user["full_name"], first.get("vehicle_number", ""))
                ids_label = ", ".join(short_req_id(x.get("id")) for x in grp)
                dests = " / ".join(dict.fromkeys(str(x["destination"]) for x in grp))
                if multi:
                    title = (f"#{n} · 🟢 Group Trip — Requisitions {ids_label} — {dests}  |  "
                             f"Vehicle: {fmt(first.get('vehicle_number'), 'N/A')}")
                else:
                    title = (f"#{n} · 🟢 Requisition {short_req_id(first.get('id'))} — {first['destination']}  |  "
                             f"Vehicle: {fmt(first.get('vehicle_number'), 'N/A')}")
                if any(not is_blank(x.get("admin_note")) for x in grp):
                    title += "  |  📝 ADMIN NOTE"
                with st.expander(title):
                    if multi:
                        st.info(f"🚐 These {len(grp)} requisitions are ONE trip — enter your Start KM only once below.")
                        for r in grp:
                            st.write(f"**{short_req_id(r.get('id'))}** — {r['applicant_name']} ({r['department']}) "
                                     f"→ {r['destination']}  |  {r['date_of_travel']} at {fmt_time_12h(r['time_of_travel'])}")
                    else:
                        c1, c2 = st.columns(2)
                        with c1:
                            st.write(f"**Applicant:** {first['applicant_name']} ({first['department']})")
                            st.write(f"**Date/Time:** {first['date_of_travel']} at {fmt_time_12h(first['time_of_travel'])}")
                        with c2:
                            approved_time = first["time_of_travel"] if is_blank(first.get("approved_time")) else first.get("approved_time")
                            st.write(f"**Approved Departure Time:** {fmt_time_12h(approved_time)}")
                            st.write(f"**Vehicle Type:** {first['vehicle_type']}")
                    show_admin_note(grp)

                    if auto_start and auto_start > 0:
                        st.caption(
                            f"↩️ Auto-filled from your last logged End KM for "
                            f"**{fmt(first.get('vehicle_number'))}** — change it below if needed."
                        )

                    with st.form(f"driver_start_{gkey}"):
                        d_start_km = driver_km_input(
                            "Start KM (Odometer Reading) *", auto_start, key=f"dstart_i_{gkey}",
                        )
                        depart_clicked = st.form_submit_button(
                            "🚦 Start Trip / Depart", type="primary", use_container_width=True
                        )

                    if depart_clicked:
                        try:
                            submit_driver_km_group([x.to_dict() for x in grp],
                                                   driver_start_km=d_start_km, driver_end_km=None)
                            st.success(
                                f"✅ Trip started ({len(grp)} requisition(s)) — status is now On Trip. "
                                "A Telegram alert has been sent."
                            )
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to record Start KM: {e}")

    # ---------------- TAB 2: End Trip (End KM only) ----------------
    if tab_return:
        st.subheader("Trips Currently On the Road")

        def _needs_end(row):
            if row.get("status") == "On Trip":
                return True
            # Also surface trips already marked Completed by Admin, as long as
            # the driver hasn't logged their own End KM yet.
            if row.get("status") == "Completed" and is_blank(row.get("driver_end_km")):
                return True
            return False

        end_trips = my_trips[my_trips.apply(_needs_end, axis=1)] if not my_trips.empty else my_trips

        if end_trips.empty:
            st.info("No trips are currently waiting for your End KM.")
        else:
            for n, grp in enumerate(group_trips(sort_by_approval_order(end_trips)), start=1):
                first = grp[0]
                gkey = first["request_id"]
                multi = len(grp) > 1
                own_starts = [float(x["driver_start_km"]) for x in grp if not is_blank(x.get("driver_start_km"))]
                has_own_start = bool(own_starts)
                start_km_val = own_starts[0] if has_own_start else 0.0
                ids_label = ", ".join(short_req_id(x.get("id")) for x in grp)
                dests = " / ".join(dict.fromkeys(str(x["destination"]) for x in grp))
                if multi:
                    title = (f"#{n} · 🔵 Group Trip — Requisitions {ids_label} — {dests}  |  "
                             f"Vehicle: {fmt(first.get('vehicle_number'), 'N/A')}")
                else:
                    title = (f"#{n} · 🔵 Requisition {short_req_id(first.get('id'))} — {first['destination']}  |  "
                             f"Vehicle: {fmt(first.get('vehicle_number'), 'N/A')}")
                if any(not is_blank(x.get("admin_note")) for x in grp):
                    title += "  |  📝 ADMIN NOTE"
                with st.expander(title):
                    if multi:
                        st.info(f"🚐 These {len(grp)} requisitions are ONE trip — enter your End KM only once below.")
                        for r in grp:
                            st.write(f"**{short_req_id(r.get('id'))}** — {r['applicant_name']} ({r['department']}) → {r['destination']}")
                        st.write(f"**Your Start KM:** {fmt(first.get('driver_start_km'))}  |  "
                                 f"**Trip Started:** {fmt_time_12h(first.get('actual_exit_time'))}")
                    else:
                        c1, c2 = st.columns(2)
                        with c1:
                            st.write(f"**Applicant:** {first['applicant_name']} ({first['department']})")
                            st.write(f"**Vehicle:** {fmt(first.get('vehicle_number'), 'N/A')} ({first['vehicle_type']})")
                            st.write(f"**Destination:** {first['destination']}")
                        with c2:
                            st.write(f"**Your Start KM:** {fmt(first.get('driver_start_km'))}")
                            st.write(f"**Trip Started:** {fmt_time_12h(first.get('actual_exit_time'))}")

                    show_admin_note(grp)

                    if not has_own_start:
                        st.warning(
                            "⚠️ You don't have a Start KM logged for this trip yet (it was likely "
                            "completed directly by Admin). Please enter BOTH your "
                            "Start KM and End KM below so this trip has a proper distance on record."
                        )

                    with st.form(f"driver_end_{gkey}"):
                        if not has_own_start:
                            d_start_km = driver_km_input(
                                "Start KM (Odometer Reading) *", 0, key=f"dend_start_i_{gkey}",
                            )
                        else:
                            d_start_km = start_km_val
                        if has_own_start:
                            st.caption(
                                f"✏️ End KM is pre-filled with your Start KM ({km_plain(start_km_val)}). "
                                "Just change the last 2-3 digits to the current odometer reading."
                            )
                        d_end_km = driver_km_input(
                            "End KM (Odometer Reading) *",
                            start_km_val if has_own_start else 0,
                            key=f"dend_i_{gkey}",
                            help=(
                                f"Must be greater than your Start KM ({km_plain(start_km_val)})."
                                if has_own_start else
                                "Must be greater than or equal to the Start KM you enter above."
                            ),
                        )
                        return_clicked = st.form_submit_button(
                            "🏁 Complete Trip / Return", type="primary", use_container_width=True
                        )

                    if return_clicked:
                        if d_end_km < d_start_km:
                            st.error("End KM cannot be less than Start KM.")
                        elif has_own_start and d_end_km == d_start_km:
                            # End KM now starts pre-filled with the Start KM, so make sure
                            # the driver didn't press the button without updating it.
                            st.error(
                                "End KM is still the same as Start KM — please change the last "
                                "digits to the current odometer reading, then submit."
                            )
                        else:
                            try:
                                submit_driver_km_group(
                                    [x.to_dict() for x in grp],
                                    driver_start_km=d_start_km,
                                    driver_end_km=d_end_km,
                                )
                                st.success(
                                    f"✅ Trip completed ({len(grp)} requisition(s)). "
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
# Telegram alert on top of the normal "New Requisition" one, so Admin
# notices it immediately. It reuses insert_requisition()
# (not a raw Supabase call) so it still shows up in Admin's Pending Requests
# queue for driver/vehicle assignment exactly like any other request.
elif user["role"] == "nurse":
    company_header("🚨 Medical Emergency Vehicle Request")
    st.caption(f"Logged in as {user['full_name']} — {user.get('designation') or 'Nurse'}")

    tab_emergency, tab_my_emergency_requests = section_nav("nurse_section",
        ["🚨 Emergency Request", "📍 My Requests / Live Status"]
    )

    if tab_emergency:
        st.markdown("### 🚑 Need a vehicle right now to carry a patient?")
        st.write(
            "Tap the button below to submit an emergency vehicle request immediately — "
            "no form to fill in. Admin is notified right away to arrange "
            "a driver and vehicle."
        )
        if st.button("🚨 EMERGENCY — Request Vehicle for Patient Carry", type="primary",
                      use_container_width=True):
            request_id = generate_request_id()
            now = bd_now()
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
                        "Admin has been alerted."
                    )
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ Failed to send emergency request: {e}")

    if tab_my_emergency_requests:
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
    tab_pending_req, tab_create_req, tab_shuttle, tab_roster, tab_users, tab_pending_users, tab_analytics, tab_export, \
        tab_edit_trip, tab_fleet, tab_duty, tab_management = section_nav("admin_section", [
            "🚗 Pending Requests", "➕ Create Requisition", "🚌 Staff Shuttle", "🤖 Auto Duty Roster", "👥 User List", "⏳ ID Requests",
            "📊 Analytics", "📁 All Requisitions & Export", "✏️ Edit / Delete Trip",
            "🚘 Manage Drivers & Vehicles",
            "🕒 Duty Tracker & Analytics",
            "🏆 Management Dashboard",
        ])

    # Hoisted above every tab body so both users_df and df_all are guaranteed
    # ready regardless of which tab runs first in code — previously these
    # were only fetched inside specific tab bodies (old Tab 1 / old Tab 3),
    # which would have broken once those tabs were reordered.
    # Load ONLY what the section that is open actually uses (each section's code
    # runs only when it is selected, and so should its data):
    #   * Pending Requests (the landing page) -> just the few active trips
    #   * Analytics / Export / Edit / Duty / Management -> full history
    #   * Create Requisition / User List / ID Requests -> the users table
    #   * Staff Shuttle / Manage Drivers & Vehicles -> nothing up front
    users_df = pd.DataFrame()
    df_all = pd.DataFrame(columns=REQUISITION_COLUMNS)
    _needs_full_history = tab_analytics or tab_export or tab_edit_trip or tab_duty or tab_management
    _needs_users = tab_create_req or tab_users or tab_pending_users
    if _needs_full_history or tab_pending_req or _needs_users:
        with st.spinner("Loading..."):
            if _needs_users:
                users_df = fetch_all_users()
            if _needs_full_history:
                df_all = fetch_all_requisitions()
            elif tab_pending_req:
                df_all = fetch_active_requisitions()

    # ---------------- Pending Requests (was Tab 3) ----------------
    if tab_pending_req:
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
            with st.expander("➕ Add to an Already-Approved Trip (same driver & vehicle)"):
                st.caption(
                    "Use this when a trip is ALREADY Approved (or On Trip) and another request comes in "
                    "for the same time — e.g. Commercial's HIACE at 11:30 AM is approved with Manik, then "
                    "HR asks for 11:30 AM too. Pick the new request(s) and the existing trip: they join it, "
                    "Manik sees ONE trip and enters KM once."
                )
                open_trips = df_all[df_all["status"].isin(["Approved", "On Trip"])]
                if open_trips.empty:
                    st.info("No Approved / On Trip trips to join right now.")
                else:
                    attach_pending_rows = {r["request_id"]: r for _, r in pending_df.iterrows()}
                    attach_option_map = {
                        f"{short_req_id(r.get('id'))} — {r['applicant_name']} ({r['department']}) → "
                        f"{r['destination']} @ {fmt_time_12h(r['time_of_travel'])}": r["request_id"]
                        for _, r in pending_df.iterrows()
                    }
                    attach_selected = st.multiselect(
                        "1. New Pending request(s) to add", list(attach_option_map.keys()), key="attach_select",
                    )
                    target_map = {}
                    for _, t in open_trips.iterrows():
                        t_time = t["time_of_travel"] if is_blank(t.get("approved_time")) else t.get("approved_time")
                        target_map[
                            f"{short_req_id(t.get('id'))} — {t['status']} — Driver: {fmt(t.get('driver_name'))} / "
                            f"Vehicle: {fmt(t.get('vehicle_number'))} @ {fmt_time_12h(t_time)} — "
                            f"{t['applicant_name']} → {t['destination']}"
                        ] = t["request_id"]
                    attach_target_label = st.selectbox(
                        "2. Join this existing trip", list(target_map.keys()), key="attach_target",
                    )
                    if st.button(
                        f"➕ Add {len(attach_selected)} request(s) to this trip", type="primary",
                        use_container_width=True, disabled=not attach_selected, key="attach_submit",
                    ):
                        a_ids = [attach_option_map[l] for l in attach_selected]
                        a_target = df_all[df_all["request_id"] == target_map[attach_target_label]].iloc[0]
                        try:
                            with st.spinner("Adding to the trip..."):
                                attach_requisitions_to_trip(
                                    a_ids, [attach_pending_rows[i] for i in a_ids], a_target, user["full_name"],
                                )
                            st.success(
                                f"✅ Added {len(a_ids)} request(s) to the trip — Driver "
                                f"**{fmt(a_target.get('driver_name'))}** / Vehicle **{fmt(a_target.get('vehicle_number'))}**."
                            )
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to add to the trip: {e}")

            with st.expander("🚐 Bulk Assign Vehicle & Driver (Multiple Requests at Once)"):
                st.caption(
                    "Select two or more Pending requests, pick ONE Driver, Vehicle and Departure Time, "
                    "and approve them together. They become ONE group trip: the team gets a single "
                    "combined Telegram message, and the Driver enters Start/End KM only once for all of them."
                )
                pending_rows = {r["request_id"]: r for _, r in pending_df.iterrows()}
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
                    ids = [bulk_option_map[l] for l in bulk_selected_labels]
                    approved_hhmm = bulk_time.strftime("%H:%M")
                    bulk_updates = {
                        "status": "Approved",
                        "driver_name": bulk_driver,
                        "driver_contact": driver_contact_map.get(bulk_driver, ""),
                        "vehicle_number": bulk_vehicle,
                        "approved_by": user["full_name"],
                        "action_timestamp": bd_now_str(),
                        "approved_time": approved_hhmm,
                        "admin_note": bulk_note.strip(),
                        "trip_group_id": generate_trip_group_id(),
                    }
                    try:
                        with st.spinner("Approving selected requests..."):
                            bulk_update_requisitions(ids, bulk_updates)
                            send_bulk_assignment_alert(
                                [pending_rows[i] for i in ids], bulk_driver, bulk_vehicle,
                                approved_hhmm, user["full_name"], bulk_note.strip(),
                            )
                        st.success(
                            f"✅ Approved {len(ids)} requests as ONE group trip — Driver **{bulk_driver}** / "
                            f"Vehicle **{bulk_vehicle}**."
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(
                            f"❌ Bulk approval failed: {e}\n\n"
                            "If this mentions `trip_group_id`, run this once in the Supabase SQL Editor: "
                            "`alter table requisitions add column if not exists trip_group_id text;`"
                        )
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
            PENDING_PAGE_SIZE = 15
            PENDING_DISPLAY_LIMIT = st.session_state.get("pending_show_n", PENDING_PAGE_SIZE)
            pending_display_df = pending_df.head(PENDING_DISPLAY_LIMIT)
            if len(pending_df) > PENDING_DISPLAY_LIMIT:
                st.info(
                    f"Showing the {PENDING_DISPLAY_LIMIT} most recent of {len(pending_df)} pending "
                    "requests, so the page opens quickly."
                )
                if st.button(f"⬇️ Show {PENDING_PAGE_SIZE} more", key="pending_show_more"):
                    st.session_state["pending_show_n"] = PENDING_DISPLAY_LIMIT + PENDING_PAGE_SIZE
                    st.rerun()
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
                            default_time = bd_now().time()
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
                                    "action_timestamp": bd_now_str(),
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
    if tab_create_req:
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
        # Each option reads "Full Name (username) · Department". The search box of the
        # dropdown matches ANY part of that text, so the Admin can type either the
        # person's real name or their login username and still find them.
        user_label_map = {}
        if not users_df.empty:
            _approved = users_df[users_df["status"] == "Approved"]
            _rows = []
            for _, _u in _approved.iterrows():
                _full = fmt(_u.get("full_name"), "(no name)")
                _lbl = f"{_full} ({_u['username']})"
                _dept = fmt(_u.get("department"), "")
                if _dept:
                    _lbl += f" · {_dept}"
                _rows.append((_full.lower(), _lbl, _u["username"]))
            for _, _lbl, _uname in sorted(_rows):
                user_label_map[_lbl] = _uname
        existing_user_options = ["— Manual Entry —"] + list(user_label_map.keys())
        selected_user_label = st.selectbox(
            "Fill from an existing user — type the Full Name or the Username to search (optional)",
            existing_user_options, key="create_req_existing_user_v2",
        )
        selected_existing_user = user_label_map.get(selected_user_label, "— Manual Entry —")
        prefill = {}
        if selected_existing_user != "— Manual Entry —" and not users_df.empty:
            match = users_df[users_df["username"] == selected_existing_user]
            if not match.empty:
                prefill = match.iloc[0].to_dict()

        # Department of the picked user -> position in the list (ignores spaces / capital
        # letters, so "quality " still matches "Quality").
        _pref_dept = str(prefill.get("department") or "").strip().lower()
        _pref_dept_idx = next((i for i, d in enumerate(DEPARTMENTS) if d.lower() == _pref_dept), 0)

        with st.form("admin_create_req_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                ca_applicant_name = st.text_input("Applicant Name *", value=fmt(prefill.get("full_name"), ""))
                # The key includes the picked user: a widget with a fixed key remembers its
                # old value and ignores a new `index`, which is why Department stayed on
                # "Accounts" while Name and Mobile were refilled correctly.
                ca_department = st.selectbox(
                    "Department *", DEPARTMENTS,
                    index=_pref_dept_idx,
                    key=f"ca_department_{selected_existing_user}",
                )
                ca_mobile = st.text_input("Mobile Number *", value=fmt(prefill.get("mobile"), ""),
                                           placeholder="01XXXXXXXXX")
                ca_passenger_count = st.number_input("Passenger Count *", min_value=1, max_value=50, value=1,
                                                      key="ca_passenger_count")
            with c2:
                ca_date = st.date_input("Date of Travel *", min_value=bd_today(), key="ca_date")
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
                now_str = bd_now_str()
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
    if tab_shuttle:
        st.subheader("🚌 Staff Shuttle — Quick Submit")
        st.caption(
            "For fixed, recurring HIACE routes (e.g. the morning staff pickup, the evening "
            "7:15 PM drop-off) — save each route once as a template below, then submit "
            "tomorrow's trip with a single click instead of retyping everything every day."
        )
        st.info(
            f"🤖 **Automatic:** every night after "
            f"{TOMORROW_SHUTTLE_AUTO_HOUR % 12 or 12}:{TOMORROW_SHUTTLE_AUTO_MINUTE:02d} "
            f"{'PM' if TOMORROW_SHUTTLE_AUTO_HOUR >= 12 else 'AM'} (Bangladesh time), as soon as anyone opens the app on "
            "any device, ALL saved templates below are submitted for tomorrow automatically and go straight "
            "to the driver as **Approved**. A template needs a default Driver + Vehicle for this — "
            "templates without them are NOT submitted (never left as Pending). "
            "Templates already submitted for tomorrow are never duplicated, so the manual button is still safe to use."
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
        tomorrow_date = bd_today() + timedelta(days=1)

        def _submit_shuttle_trip(tpl_row, trip_date, driver_name, vehicle_number):
            """Creates one Approved requisition from a shuttle template for
            the given date/driver/vehicle. Returns the new short req id, or
            raises on failure (caller handles the try/except + message)."""
            request_id = generate_request_id()
            now_str = bd_now_str()
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
                    # Past 9:55 PM? Submit the newly saved template for tomorrow right away.
                    st.session_state.pop("_tomorrow_shuttle_checked_date", None)
                    ensure_tomorrow_shuttle_submitted()
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
                if (tpl.get("default_driver_name") not in shuttle_driver_options
                        or tpl.get("default_vehicle_number") not in shuttle_vehicle_options):
                    r1.caption("⚠️ No valid default Driver/Vehicle — this template will NOT be auto-submitted. "
                               "Delete it and save again with a Driver + Vehicle.")
                if r2.button("🗑️", key=f"del_shuttle_tpl_{tpl['id']}", help="Delete this template"):
                    try:
                        delete_shuttle_template(tpl["id"])
                        st.success(f"Deleted template '{tpl['template_name']}'.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Failed to delete template: {e}")

    # ---------------- Auto Duty Roster (NEW) ----------------
    # Admin picks ONE Driver + ONE Vehicle for each roster group (Commercial Duty
    # = all 6 trips, Night Shift = all 5 trips). The choice is saved and every
    # evening's auto requisitions use it (created already Approved) until Admin
    # changes it — e.g. the weekly rotation.
    if tab_roster:
        st.subheader("🤖 Auto Duty Roster — Commercial Duty & Night Shift")
        st.caption(
            "Select ONE Driver and ONE Vehicle for each group and press Save — they are used for ALL trips "
            "of that group and STAY the same every day until you change them. After "
            f"{DAILY_AUTO_START_HOUR % 12 or 12}:{DAILY_AUTO_START_MINUTE:02d} "
            f"{'PM' if DAILY_AUTO_START_HOUR >= 12 else 'AM'} (Bangladesh time) the app creates each day's "
            "requisitions automatically, always **Approved** and sent straight to the driver. "
            "A group with no Driver/Vehicle selected (or un-ticked Active) is NOT created. "
            "Under each group you can also edit its trips (time, destination, department, applicant), add new trips or remove trips."
        )

        if st.session_state.get("roster_flash"):
            st.success(st.session_state.pop("roster_flash"))

        roster_map, roster_err = fetch_auto_roster()
        if roster_err:
            st.error(
                "⚠️ The roster table is not set up yet. Open **Supabase → SQL Editor**, run the SQL below "
                "once, then refresh this page."
            )
            st.code(AUTO_ROSTER_SQL, language="sql")
            st.caption(f"Technical detail: {roster_err}")
        else:
            roster_drivers_df = fetch_all_drivers()
            roster_vehicles_df = fetch_all_vehicles()
            roster_driver_opts = roster_drivers_df["driver_name"].tolist() if not roster_drivers_df.empty else []
            roster_vehicle_opts = roster_vehicles_df["vehicle_number"].tolist() if not roster_vehicles_df.empty else []
            roster_contact_map = (
                dict(zip(roster_drivers_df["driver_name"], roster_drivers_df["driver_contact"]))
                if not roster_drivers_df.empty else {}
            )
            NOT_SET = "— Not set (trips NOT created) —"

            if not roster_driver_opts or not roster_vehicle_opts:
                st.warning(
                    "⚠️ No drivers or vehicles saved yet. Add them first under "
                    "**🚘 Manage Drivers & Vehicles**, then come back here."
                )

            for g in DAILY_AUTO_GROUPS:
                if not g["roster"]:
                    continue
                g_tag = g["tag"]
                g_trips, g_next_tid = get_group_trips(g, roster_map)
                saved = roster_map.get(g_tag, {})
                saved_driver = (saved.get("driver_name") or "").strip()
                saved_vehicle = (saved.get("vehicle_number") or "").strip()
                saved_enabled = saved.get("enabled")
                saved_enabled = True if saved_enabled is None else bool(saved_enabled)

                drv_opts = [NOT_SET] + roster_driver_opts
                if saved_driver and saved_driver not in drv_opts:
                    drv_opts.insert(1, saved_driver)
                veh_opts = [NOT_SET] + roster_vehicle_opts
                if saved_vehicle and saved_vehicle not in veh_opts:
                    veh_opts.insert(1, saved_vehicle)

                st.markdown(f"#### {g['title']}")
                st.caption(f"Applies to all {len(g_trips)} trip(s) of this group each day.")
                if not (saved_driver and saved_vehicle):
                    st.warning("⚠️ No Driver/Vehicle saved for this group — its trips are NOT being created. Select and Save below.")
                elif not saved_enabled:
                    st.warning("⏸️ This group is switched off (Active un-ticked) — its trips are NOT being created.")

                # ---- one Driver + one Vehicle for the whole group ----
                with st.form(f"roster_form_{g_tag}"):
                    c1, c2, c3 = st.columns([2.6, 2.2, 1])
                    sel_driver = c1.selectbox(
                        "Driver", drv_opts,
                        index=drv_opts.index(saved_driver) if saved_driver in drv_opts else 0,
                        key=f"roster_drv_{g_tag}",
                    )
                    sel_vehicle = c2.selectbox(
                        "Vehicle", veh_opts,
                        index=veh_opts.index(saved_vehicle) if saved_vehicle in veh_opts else 0,
                        key=f"roster_veh_{g_tag}",
                    )
                    sel_active = c3.checkbox("Active", value=saved_enabled, key=f"roster_on_{g_tag}")
                    save_clicked = st.form_submit_button(
                        f"💾 Save {g['title']} Driver & Vehicle", type="primary", use_container_width=True,
                    )

                if save_clicked:
                    drv_val = "" if sel_driver == NOT_SET else sel_driver
                    veh_val = "" if sel_vehicle == NOT_SET else sel_vehicle
                    if bool(drv_val) != bool(veh_val):
                        st.error("❌ Select BOTH Driver and Vehicle (or leave both as 'Not set').")
                    else:
                        try:
                            save_auto_roster_slots([{
                                "slot_key": g_tag, "group_tag": g_tag, "slot_no": 0,
                                "driver_name": drv_val, "vehicle_number": veh_val,
                                "enabled": bool(sel_active),
                                "updated_by": user["full_name"], "updated_at": bd_now_str(),
                            }])
                            # Existing (Pending / any) requisitions are never touched. If it is
                            # already past the auto-start time, create any of today's trips of
                            # this group that are still missing right now.
                            st.session_state.pop("_daily_shuttle_checked_date", None)
                            ensure_daily_shuttle_requisitions()
                            st.session_state["roster_flash"] = (
                                f"✅ {g['title']} Driver & Vehicle saved. They will be used for all "
                                f"{len(g_trips)} trip(s) every day until you change them."
                            )
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to save roster: {e}")

                # ---- the trips themselves: edit / add / remove ----
                with st.expander(
                    f"✏️ Edit the {len(g_trips)} trip(s) of this group — change time, destination, department, "
                    "applicant name · add new trips · remove trips", expanded=False,
                ):
                    st.caption(
                        "Edit any cell directly. **Add a trip:** click the empty row at the bottom of the table (＋). "
                        "**Remove a trip:** tick 🗑️ Delete on that row. Then press the Save button under the table. "
                        "Changes apply from the next auto-creation (trips already created today are not changed — "
                        "use the list at the bottom of this page for those)."
                    )
                    trips_sig = hashlib.md5(json.dumps(g_trips, sort_keys=True).encode("utf-8")).hexdigest()[:8]
                    trips_df = pd.DataFrame([{
                        "🗑️ Delete": False,
                        "Applicant Name": tr["applicant"],
                        "Department": tr["department"],
                        "Destination": tr["destination"],
                        "Time (HH:MM, 24h)": tr["time"],
                        "Passengers": tr["passengers"],
                        "Purpose": tr["purpose"],
                        "Vehicle Type": tr["vehicle_type"],
                        "_tid": tr["tid"],
                    } for tr in g_trips], columns=[
                        "🗑️ Delete", "Applicant Name", "Department", "Destination", "Time (HH:MM, 24h)",
                        "Passengers", "Purpose", "Vehicle Type", "_tid",
                    ])
                    default_tpl = g["templates"][0]
                    with st.form(f"trips_form_{g_tag}"):
                        edited_trips = st.data_editor(
                            trips_df, key=f"trips_editor_{g_tag}_{trips_sig}", num_rows="dynamic",
                            hide_index=True, use_container_width=True,
                            column_config={
                                "🗑️ Delete": st.column_config.CheckboxColumn("🗑️ Delete", default=False),
                                "Applicant Name": st.column_config.TextColumn("Applicant Name", default=g["applicant"]),
                                "Department": st.column_config.TextColumn("Department", default=g["department"]),
                                "Destination": st.column_config.TextColumn("Destination", default=default_tpl["destination"]),
                                "Time (HH:MM, 24h)": st.column_config.TextColumn(
                                    "Time (HH:MM, 24h)", default=default_tpl["time"], help="e.g. 19:15 (7:15 PM also works)"),
                                "Passengers": st.column_config.NumberColumn("Passengers", min_value=1, step=1, default=1),
                                "Purpose": st.column_config.TextColumn("Purpose", default=default_tpl["purpose"]),
                                "Vehicle Type": st.column_config.SelectboxColumn(
                                    "Vehicle Type", options=VEHICLE_TYPES, default="HIACE"),
                                "_tid": None,
                            },
                        )
                        save_trips_clicked = st.form_submit_button(
                            f"💾 Save {g['title']} Trips", type="primary", use_container_width=True,
                        )

                    if save_trips_clicked:
                        new_trips, errors = [], []
                        next_tid = g_next_tid
                        for n, (_, er) in enumerate(edited_trips.iterrows(), start=1):
                            if bool(er.get("🗑️ Delete")):
                                continue
                            dest_txt = str(er.get("Destination") or "").strip()
                            appl_txt = str(er.get("Applicant Name") or "").strip()
                            if not dest_txt and not appl_txt:
                                continue  # blank row
                            if not dest_txt:
                                errors.append(f"Row {n}: Destination is required.")
                                continue
                            hhmm = parse_trip_time(er.get("Time (HH:MM, 24h)"))
                            if not hhmm:
                                errors.append(f"Row {n}: Time '{er.get('Time (HH:MM, 24h)')}' is not valid — use HH:MM, e.g. 19:15.")
                                continue
                            tid_val = er.get("_tid")
                            if pd.isna(tid_val):
                                tid_val = next_tid
                                next_tid += 1
                            try:
                                pax = max(1, int(er.get("Passengers") or 1))
                            except (TypeError, ValueError):
                                pax = 1
                            vt = str(er.get("Vehicle Type") or "").strip() or "HIACE"
                            new_trips.append({
                                "tid": int(tid_val),
                                "applicant": appl_txt or g["applicant"],
                                "department": str(er.get("Department") or "").strip() or g["department"],
                                "destination": dest_txt,
                                "time": hhmm,
                                "passengers": pax,
                                "purpose": str(er.get("Purpose") or "").strip() or default_tpl["purpose"],
                                "vehicle_type": vt,
                            })
                        if errors:
                            for msg_err in errors:
                                st.error(f"❌ {msg_err}")
                        else:
                            try:
                                save_group_trips(g_tag, new_trips, next_tid, user["full_name"])
                                # Past the auto-start time? Create any new trip of today right away.
                                st.session_state.pop("_daily_shuttle_checked_date", None)
                                ensure_daily_shuttle_requisitions()
                                st.session_state["roster_flash"] = (
                                    f"✅ {g['title']}: {len(new_trips)} trip(s) saved. "
                                    "They will be created automatically every day."
                                )
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Failed to save trips: {e}")
                                st.caption(
                                    "If the message mentions 'trips_json', run this once in Supabase → SQL Editor, "
                                    "then try again:"
                                )
                                st.code(
                                    "alter table auto_duty_roster add column if not exists trips_json text default '';",
                                    language="sql",
                                )

            st.info(
                "ℹ️ Existing Pending requisitions are never changed by this page. Staff Shuttle and Female Worker "
                "Drop requisitions are created as **Pending** every day — approve them from "
                "**🚗 Pending Requests** as usual."
            )

        # ------------------------------------------------------------------
        # Commercial Duty + Night Shift trips: list, pick one from the dropdown,
        # and edit EVERYTHING about it in the vertical edit form (same form as
        # "✏️ Edit / Delete Trip": status, driver, vehicle, times, KM, note ...).
        # ------------------------------------------------------------------
        st.markdown("---")
        st.markdown("### 📋 Commercial Duty & Night Shift Trips — View / Edit")
        if st.session_state.get("edit_flash"):
            st.success(st.session_state.pop("edit_flash"))

        lf1, lf2, lf3 = st.columns([2, 2, 3])
        with lf1:
            list_group = st.selectbox(
                "Group", ["All (Commercial + Night Shift)", "💼 Commercial Duty", "🌙 Night Shift Vehicle"],
                key="auto_trips_group",
            )
        with lf2:
            list_status = st.multiselect(
                "Status", REQ_STATUS_OPTIONS, default=["Approved"], key="auto_trips_status",
                help="Default shows the Approved trips. Add other statuses (On Trip, Completed ...) or clear it to see all.",
            )
        with lf3:
            list_dates = st.date_input(
                "Trip date (from — to)", value=(bd_today(), bd_today()), key="auto_trips_dates",
            )
        if isinstance(list_dates, (tuple, list)):
            d_from = list_dates[0]
            d_to = list_dates[1] if len(list_dates) == 2 else list_dates[0]
        else:
            d_from = d_to = list_dates

        try:
            auto_trips_df = fetch_auto_duty_trips(str(d_from), str(d_to))
        except Exception as e:
            auto_trips_df = pd.DataFrame(columns=REQUISITION_COLUMNS)
            st.error(f"❌ Could not load the trips: {e}")

        def _trip_group_label(note):
            note = str(note or "")
            if DAILY_COMMERCIAL_TAG in note:
                return "💼 Commercial Duty"
            if DAILY_NIGHT_SHIFT_TAG in note:
                return "🌙 Night Shift Vehicle"
            return "—"

        if not auto_trips_df.empty:
            auto_trips_df = auto_trips_df.copy()
            auto_trips_df["_group"] = auto_trips_df["admin_note"].map(_trip_group_label)
            if list_group != "All (Commercial + Night Shift)":
                auto_trips_df = auto_trips_df[auto_trips_df["_group"] == list_group]
            if list_status:
                auto_trips_df = auto_trips_df[auto_trips_df["status"].isin(list_status)]

        if auto_trips_df.empty:
            st.info("No Commercial Duty / Night Shift trips match these filters.")
        else:
            st.caption(f"{len(auto_trips_df)} trip(s) found.")
            st.dataframe(
                pd.DataFrame({
                    "Req #": auto_trips_df["id"].map(short_req_id),
                    "Group": auto_trips_df["_group"],
                    "Date": auto_trips_df["date_of_travel"],
                    "Time": auto_trips_df["time_of_travel"].map(lambda v: fmt_time_12h(v, "")),
                    "Duty / Dept": auto_trips_df["department"].fillna(""),
                    "Destination": auto_trips_df["destination"].fillna(""),
                    "Status": auto_trips_df["status"].map(lambda s: STATUS_BADGE.get(s, s)),
                    "Driver": auto_trips_df["driver_name"].fillna(""),
                    "Vehicle": auto_trips_df["vehicle_number"].fillna(""),
                }),
                use_container_width=True, hide_index=True,
                height=min(80 + 35 * len(auto_trips_df), 420),
            )

            trip_options = {}
            for _, _r in auto_trips_df.iterrows():
                _label = (
                    f"{short_req_id(_r.get('id'))} | {_r['_group']} | "
                    f"{STATUS_BADGE.get(_r['status'], _r['status'])} | "
                    f"{_r['date_of_travel']} {fmt_time_12h(_r['time_of_travel'], '')} | "
                    f"{fmt(_r.get('department'), '')} → {fmt(_r.get('destination'), '')} | "
                    f"🚘 {fmt(_r.get('driver_name'), '—')}"
                )
                trip_options[_label] = _r["request_id"]
            picked = st.selectbox(
                f"Select a trip to edit ({len(trip_options)} found)",
                ["— Select a trip —"] + list(trip_options.keys()), key="auto_trips_pick",
            )
            if picked != "— Select a trip —":
                picked_row = auto_trips_df[auto_trips_df["request_id"] == trip_options[picked]].iloc[0]
                admin_edit_trip_panel(
                    picked_row, auto_trips_df, fetch_all_drivers(), fetch_all_vehicles(),
                )

    # ---------------- User List (was Tab 2) ----------------
    if tab_users:
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
                    # An old account whose role no longer exists is still shown correctly
                    # instead of silently flipping to "Employee" on the next save.
                    role_choices = ROLE_OPTIONS + ([current_role] if current_role and current_role not in ROLE_OPTIONS else [])
                    new_role = st.selectbox("Role", role_choices,
                                             index=role_choices.index(current_role) if current_role in role_choices else 0,
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
    if tab_pending_users:
        st.subheader("New Account Requests")
        pending_users = users_df[users_df["status"] == "Pending"] if not users_df.empty else users_df

        if pending_users.empty:
            st.success("🎉 No pending user registrations.")
        else:
            for _, u in pending_users.iterrows():
                role_req = u.get("role")
                if role_req == "driver":
                    role_tag = "🚙 Driver"
                else:
                    role_tag = "👤 Employee"
                with st.expander(f"{role_tag} — {u['full_name']} (@{u['username']})"):
                    if role_req == "driver":
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
    if tab_analytics:
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

            # Driver-verified KM (falls back to the older start/end readings
            # only when a driver hasn't logged their own numbers) is the primary source for this metric, per
            # business requirement — the metric itself is unchanged, only
            # its source.
            completed_trips = df_all[df_all["status"] == "Completed"].copy()
            if not completed_trips.empty:
                completed_trips["_eff_km"] = completed_trips.apply(
                    lambda row: effective_km_fields(row)[2], axis=1
                )
                total_km_all = float(completed_trips["_eff_km"].where(~mark_group_duplicates(completed_trips), 0.0).sum())
                st.metric("🛣️ Total KM Covered (Completed Trips — Driver-verified)", f"{total_km_all:.1f} KM")

            df_all["_dt"] = effective_trip_date(df_all)
            monthly = df_all.dropna(subset=["_dt"]).copy()
            if not monthly.empty:
                monthly["Month"] = monthly["_dt"].dt.to_period("M").astype(str)
                monthly_counts = monthly.groupby("Month").size().reset_index(name="Requests")
                fig3 = px.line(monthly_counts, x="Month", y="Requests", markers=True, title="Monthly Request Trend")
                fig3.update_layout(height=350)
                st.plotly_chart(fig3, use_container_width=True, key="analytics_monthly_trend")

    # ---------------- All Requisitions & Export (was Tab 5) ----------------
    if tab_export:
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

            df_all["_dt"] = effective_trip_date(df_all)
            min_d = df_all["_dt"].min()
            max_d = df_all["_dt"].max()
            default_start = min_d.date() if pd.notnull(min_d) else bd_today()
            default_end = max_d.date() if pd.notnull(max_d) else bd_today()
            date_range = st.date_input("Trip date range (actual start date)", value=(default_start, default_end))

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

            filtered_display = humanize_timestamp_columns(filtered.drop(columns=["_dt"], errors="ignore"))
            if "time_of_travel" in filtered_display.columns:
                filtered_display["time_of_travel"] = filtered_display["time_of_travel"].apply(
                    lambda v: fmt_time_12h(v, v)
                )
            # Total KM for the table AND the PDF: driver-first (falls back to
            # the older start/end readings), so driver-only trips don't show blank.
            if not filtered.empty:
                filtered_display["total_km"] = filtered.apply(lambda row: effective_km_fields(row)[2], axis=1)
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
    if tab_edit_trip:
        st.subheader("✏️ Edit or Delete a Requisition")
        if st.session_state.get("edit_flash"):
            st.success(st.session_state.pop("edit_flash"))
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

            st.markdown("##### 🔎 1. Find the trip (any status — Pending, Approved, On Trip, Completed, Rejected)")
            ef1, ef2, ef3 = st.columns([2, 2, 3])
            with ef1:
                edit_status_filter = st.multiselect("Status (empty = all)", REQ_STATUS_OPTIONS, key="edit_find_status")
            with ef2:
                known_drivers = sorted({str(d).strip() for d in df_all["driver_name"].dropna().tolist() if str(d).strip()})
                edit_driver_filter = st.selectbox("Driver", ["All Drivers"] + known_drivers, key="edit_find_driver")
            with ef3:
                edit_search = st.text_input(
                    "Search", key="edit_find_search",
                    placeholder="Req # (e.g. 42), name, destination, vehicle, purpose...",
                )
            edit_date_on = st.checkbox("Filter by trip date (actual start date)", key="edit_find_date_on")
            edit_date_range = None
            if edit_date_on:
                edit_date_range = st.date_input("Trip date range (actual start date)", value=(bd_today(), bd_today()),
                                                key="edit_find_date_range")

            edit_scope = df_all.copy()
            if edit_status_filter:
                edit_scope = edit_scope[edit_scope["status"].isin(edit_status_filter)]
            if edit_driver_filter != "All Drivers":
                _tk = _normalize_driver_name(edit_driver_filter)
                edit_scope = edit_scope[edit_scope["driver_name"].map(_normalize_driver_name) == _tk]
            if edit_date_on and edit_date_range:
                _edate = effective_trip_date(edit_scope).dt.date
                if isinstance(edit_date_range, (tuple, list)) and len(edit_date_range) == 2:
                    edit_scope = edit_scope[(_edate >= edit_date_range[0]) & (_edate <= edit_date_range[1])]
                elif isinstance(edit_date_range, (tuple, list)) and len(edit_date_range) == 1:
                    edit_scope = edit_scope[_edate == edit_date_range[0]]
            if edit_search.strip():
                _needle = edit_search.strip().lstrip("#").lower()
                _blob = pd.Series("", index=edit_scope.index)
                for _c in ("applicant_name", "destination", "vehicle_number", "driver_name",
                           "purpose", "department", "request_id"):
                    if _c in edit_scope.columns:
                        _blob = _blob + " " + edit_scope[_c].fillna("").astype(str)
                _mask = _blob.str.lower().str.contains(_needle, regex=False)
                if _needle.isdigit() and "id" in edit_scope.columns:
                    _ids = pd.to_numeric(edit_scope["id"], errors="coerce").fillna(-1).astype(int).astype(str)
                    _mask = _mask | (_ids == _needle)
                edit_scope = edit_scope[_mask]

            if edit_scope.empty:
                st.info("No requisitions match these filters.")
            else:
                EDIT_MAX_OPTIONS = 300
                if len(edit_scope) > EDIT_MAX_OPTIONS:
                    st.caption(f"Showing the newest {EDIT_MAX_OPTIONS} of {len(edit_scope)} matches — narrow the filters to find older ones.")
                edit_options = {}
                for _, _r in edit_scope.head(EDIT_MAX_OPTIONS).iterrows():
                    _label = (
                        f"{short_req_id(_r.get('id'))} | {STATUS_BADGE.get(_r['status'], _r['status'])} | "
                        f"{_r['date_of_travel']} {fmt_time_12h(_r['time_of_travel'], '')} | "
                        f"{_r['applicant_name']} → {_r['destination']} | 🚘 {fmt(_r.get('driver_name'), '—')}"
                    )
                    edit_options[_label] = _r["request_id"]
                selected_label = st.selectbox(
                    f"2. Select the requisition to edit ({len(edit_scope)} found)",
                    list(edit_options.keys()), key="edit_trip_select",
                )
                selected_request_id = edit_options[selected_label]
                selected_row = df_all[df_all["request_id"] == selected_request_id].iloc[0]
                admin_edit_trip_panel(selected_row, df_all, edit_drivers_df, edit_vehicles_df)

    # ---------------- Manage Drivers & Vehicles (was Tab 6) ----------------
    if tab_fleet:
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
    if tab_duty:
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
                filter_start_date = st.date_input("Start Date", value=bd_today(), key="duty_start_date")
                # 12-hour clock with an AM/PM dropdown (same widget used in the
                # requisition forms) instead of the 24-hour st.time_input.
                filter_start_time = time_input_12h(
                    "Start Time", key_prefix="duty_start",
                    default_time=datetime.strptime("05:30", "%H:%M").time(),
                )
            with fc2:
                st.markdown("**End of Range**")
                filter_end_date = st.date_input("End Date", value=bd_today() + timedelta(days=1), key="duty_end_date")
                filter_end_time = time_input_12h(
                    "End Time", key_prefix="duty_end",
                    default_time=datetime.strptime("05:29", "%H:%M").time(),
                )

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
            # that have never started — exactly what you get right
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

            # A trip only has meaningful "duty duration" once the trip
            # has started (actual_exit_time). If it hasn't ended yet
            # (still "On Trip"), treat "now" (in UTC, to
            # match the column's dtype) as the running end time so
            # in-progress duty shows up too. Assigning a naive Timestamp
            # into a UTC-aware column is exactly what triggers the
            # "Invalid value ... for dtype 'datetime64[us, UTC]'" TypeError,
            # so we must assign an equally tz-aware Timestamp here.
            still_out_mask = duty_df_raw["_start_dt"].notna() & duty_df_raw["_end_dt"].isna()
            duty_df_raw.loc[still_out_mask, "_end_dt"] = bd_now_ts()

            # Only rows that actually started are real "duty" records.
            duty_base = duty_df_raw[duty_df_raw["_start_dt"].notna()].copy()

            # -------------------------------------------------------------
            # STEP 2 — Combine the date/time widgets into naive datetimes,
            # then localize them to UTC so they can be compared directly
            # against the tz-aware `_start_dt` / `_end_dt` columns above.
            # (Saved times are Bangladesh wall-clock values that Supabase labels
            # "+00:00"; the range typed here is Bangladesh time too, so it is
            # labeled "UTC" the same way and compared number-for-number. No
            # timezone conversion is wanted or applied.)
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

            # ---- More filters + report options ----
            # Department / Status / Route-or-Purpose text narrow down the trips
            # themselves; "Group summary by" and "Sort log by" control how the
            # Summary table and Detailed Duty Log are arranged. Everything
            # below (KPIs, summary, log, CSV / Excel / PDF exports) follows
            # whatever is picked here.
            fc5, fc6, fc7 = st.columns(3)
            with fc5:
                dept_choices = sorted(
                    d for d in duty_base["department"].dropna().astype(str).unique().tolist() if d.strip()
                ) if "department" in duty_base.columns else []
                duty_dept_filter = st.multiselect("Department", dept_choices, key="duty_dept_filter")
            with fc6:
                status_choices = sorted(
                    s_ for s_ in duty_base["status"].dropna().astype(str).unique().tolist() if s_.strip()
                ) if "status" in duty_base.columns else []
                duty_status_filter = st.multiselect("Trip Status", status_choices, key="duty_status_filter")
            with fc7:
                duty_text_filter = st.text_input(
                    "Route / Purpose contains", key="duty_text_filter", placeholder="e.g. Bepza",
                )

            fc8, fc9 = st.columns(2)
            with fc8:
                duty_group_option = st.selectbox(
                    "Group summary by", DUTY_SUMMARY_GROUP_OPTIONS, key="duty_group_option",
                    help="Driver + Date shows how many hours each driver worked on each date.",
                )
            with fc9:
                duty_sort_option = st.selectbox(
                    "Sort Detailed Log by",
                    ["Start Time — oldest first", "Start Time — newest first", "Driver Name (A–Z)",
                     "Vehicle No (A–Z)", "Duty Hours — highest first", "Total KM — highest first"],
                    key="duty_sort_option",
                )

            fc10, fc11, fc12 = st.columns(3)
            with fc10:
                duty_shift_aware = st.checkbox(
                    "🌙 Night-shift aware dates", value=True, key="duty_shift_aware",
                    help="ON: a night shift (e.g. 6:30 PM to 9:00 AM next morning) is ONE duty counted on the "
                         "date it STARTED. OFF: every trip goes to the calendar date it started on.",
                )
            with fc11:
                duty_gap_hours = st.number_input(
                    "Break with no trip, longer than (hours)", min_value=1.0, max_value=24.0, value=6.0,
                    step=0.5, key="duty_gap_hours",
                    help="If a driver has NO trip for more than this many hours, that break is NOT counted in "
                         "duty hours (TWH), and (with night-shift dates ON) the next trip starts a new duty.",
                )
            with fc12:
                duty_max_shift = st.number_input(
                    "Longest single duty (hours)", min_value=10.0, max_value=36.0, value=18.0,
                    step=1.0, key="duty_max_shift", disabled=not duty_shift_aware,
                    help="Safety limit: a trip starting later than this after the duty began is always counted "
                         "as a NEW duty, so one forgotten/stuck trip cannot merge several days into one.",
                )

            # -------------------------------------------------------------
            # STEP 3 — Apply the date/time window + vehicle/driver filters.
            # A trip is included if its duty window OVERLAPS the selected
            # range at all (not just if it starts inside it) — this correctly
            # captures overnight duties like "05:30 today to 05:29 tomorrow".
            # Both sides are now UTC-aware, so this comparison is safe.
            # -------------------------------------------------------------
            # Night-shift aware duty dates: computed on ALL trips (before the
            # date-range filter) so a shift is never cut in half by the filter.
            if duty_shift_aware:
                duty_base = assign_duty_dates(duty_base, duty_gap_hours, duty_max_shift)
            duty_filtered = duty_base[
                (duty_base["_start_dt"] <= range_end) & (duty_base["_end_dt"] >= range_start)
            ].copy()
            if duty_shift_aware and not duty_filtered.empty:
                # Keep only duties whose DUTY DATE (the day the shift started) is
                # inside the chosen dates, so a night shift that began before the
                # range never shows up as a stray earlier-date row.
                duty_filtered = duty_filtered[
                    (duty_filtered["_duty_date"] >= str(filter_start_date))
                    & (duty_filtered["_duty_date"] <= str(filter_end_date))
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
            if duty_dept_filter:
                duty_filtered = duty_filtered[duty_filtered["department"].astype(str).isin(duty_dept_filter)]
            if duty_status_filter:
                duty_filtered = duty_filtered[duty_filtered["status"].astype(str).isin(duty_status_filter)]
            if duty_text_filter.strip():
                _needle = duty_text_filter.strip()
                duty_filtered = duty_filtered[
                    duty_filtered["destination"].fillna("").astype(str).str.contains(_needle, case=False, regex=False)
                    | duty_filtered["purpose"].fillna("").astype(str).str.contains(_needle, case=False, regex=False)
                ]

            # -------------------------------------------------------------
            # STEP 4 — Derived fields: duty duration in hours, total KM.
            # Total KM now comes from effective_km_fields(), which prefers
            # the Driver's own logged Start/End KM and falls back to the
            # older start/end readings only when the driver
            # hasn't submitted their own numbers yet (business requirement:
            # Driver KM is the primary baseline metric for duty tracking).
            # -------------------------------------------------------------
            duty_filtered["_duration_hrs"] = (
                (duty_filtered["_end_dt"] - duty_filtered["_start_dt"]).dt.total_seconds() / 3600.0
            ).round(2)
            duty_filtered["_km"] = duty_filtered.apply(lambda row: effective_km_fields(row)[2], axis=1)
            # Start/End odometer readings, using the same driver-first,
            # older-reading-fallback priority as the KM total above.
            duty_filtered["_start_km"] = duty_filtered.apply(lambda row: effective_km_fields(row)[0], axis=1)
            duty_filtered["_end_km"] = duty_filtered.apply(lambda row: effective_km_fields(row)[1], axis=1)
            # Requisitions bundled in one group trip are ONE physical run: count its
            # distance / hours once (the 2nd+ requisition shows 0 in those two columns).
            _dup_group_rows = mark_group_duplicates(duty_filtered)
            duty_filtered.loc[_dup_group_rows, ["_km", "_duration_hrs"]] = 0.0

            # Apply the chosen sort order to the Detailed Duty Log.
            duty_filtered["_sort_driver"] = duty_filtered["driver_name"].fillna("").astype(str).str.lower()
            duty_filtered["_sort_vehicle"] = duty_filtered["vehicle_number"].fillna("").astype(str).str.lower()
            _sort_map = {
                "Start Time — oldest first": ("_start_dt", True),
                "Start Time — newest first": ("_start_dt", False),
                "Driver Name (A–Z)": ("_sort_driver", True),
                "Vehicle No (A–Z)": ("_sort_vehicle", True),
                "Duty Hours — highest first": ("_duration_hrs", False),
                "Total KM — highest first": ("_km", False),
            }
            _sort_col, _sort_asc = _sort_map.get(duty_sort_option, ("_start_dt", True))
            duty_filtered = duty_filtered.sort_values(_sort_col, ascending=_sort_asc, kind="stable")

            # Trips that started but never ended (no Trip End / End KM
            # yet) are counted up to "right now", so an old forgotten "On Trip"
            # entry keeps adding hours every minute and blows up the totals.
            # Flag them so Admin can complete/fix/delete them in Edit / Delete Trip.
            still_open_count = int(
                (duty_filtered["status"] == "On Trip").sum()
            ) if "status" in duty_filtered.columns else 0
            if still_open_count:
                st.warning(
                    f"⚠️ {still_open_count} trip(s) in this range are still marked **On Trip** (no return "
                    "recorded), so their duty hours keep counting up to the current time. If any of them "
                    "are old/forgotten, complete or remove them from the **✏️ Edit / Delete Trip** tab "
                    "to get accurate duty-hour totals."
                )

            # -------------------------------------------------------------
            # 🚩 NEEDS REVIEW — two quick lists (with Requisition numbers) so
            # Admin can find wrong entries at a glance and fix them in the
            # ✏️ Edit / Delete Trip tab. Follows the filters chosen above.
            # -------------------------------------------------------------
            if not duty_filtered.empty:
                st.markdown("##### 🚩 Needs Review")
                rv0, rv1, rv2, rv3 = st.columns(4)
                with rv0:
                    rv_always_hrs = st.number_input(
                        "List EVERY trip longer than (hours)", min_value=1.0, max_value=48.0, value=6.0,
                        step=0.5, key="rv_always_hrs",
                        help="Any trip that ran longer than this is listed, whatever its KM.",
                    )
                with rv1:
                    rv_min_hrs = st.number_input(
                        "Also list trips of at least (hours)…", min_value=1.0, max_value=24.0, value=3.0,
                        step=0.5, key="rv_min_hrs",
                    )
                with rv2:
                    rv_max_kmh = st.number_input(
                        "…if average speed is below (km/h)", min_value=1.0, max_value=60.0, value=10.0,
                        step=1.0, key="rv_max_kmh",
                        help="A trip that took many hours but covered very few km usually means the driver "
                             "pressed End Trip late (or forgot).",
                    )
                with rv3:
                    rv_duty_limit = st.number_input(
                        "Duty longer than (hours)", min_value=4.0, max_value=36.0, value=12.0,
                        step=0.5, key="rv_duty_limit",
                    )

                _sus = build_suspicious_trips(duty_filtered, rv_min_hrs, rv_max_kmh, rv_always_hrs)
                _longduty = build_long_duty_list(duty_filtered, rv_duty_limit, duty_gap_hours)

                st.markdown(f"**1️⃣ Long trips / long trips with few km — {len(_sus)} trip(s)**")
                if _sus.empty:
                    st.success("✅ No trips to check in this range.")
                else:
                    st.caption("Open the Req # in **✏️ Edit / Delete Trip** and correct the Trip End time (or KM).")
                    st.dataframe(_sus, use_container_width=True, hide_index=True)

                st.markdown(f"**2️⃣ Duty longer than {rv_duty_limit:g} hours — {len(_longduty)} driver-day(s)**")
                if _longduty.empty:
                    st.success(f"✅ No driver was on duty for more than {rv_duty_limit:g} hours in this range.")
                else:
                    st.caption("Check the **Longest trip** first — a late End Trip there is the usual reason.")
                    st.dataframe(_longduty, use_container_width=True, hide_index=True)
                st.markdown("---")

            # Explain an empty result instead of just showing zeros. Only trips
            # that have actually started (trip start time / the driver's Start
            # KM) count as duty — Pending / Approved trips that haven't started
            # yet never appear here, which is the most common reason for zero.
            if duty_filtered.empty:
                try:
                    _d0, _d1 = range_start.date(), range_end.date()
                    _sched = df_all.copy()
                    _sched["_tdate"] = pd.to_datetime(_sched["date_of_travel"], errors="coerce").dt.date
                    _sched = _sched[(_sched["_tdate"] >= _d0) & (_sched["_tdate"] <= _d1)]
                    _not_started = _sched[_sched["status"].isin(["Pending", "Approved"])]
                    _on_trip_sched = _sched[_sched["status"] == "On Trip"]
                    st.info(
                        f"ℹ️ No started trips found between {range_start.strftime('%d-%b-%Y %I:%M %p')} and "
                        f"{range_end.strftime('%d-%b-%Y %I:%M %p')}. Trips scheduled in these dates: "
                        f"{len(_sched)} — not started yet (Pending/Approved): {len(_not_started)}, "
                        f"On Trip: {len(_on_trip_sched)}. Only trips with a recorded start time / driver Start KM "
                        "count as duty, and any active Vehicle/Driver/Department/Status/Route filters also apply."
                    )
                except Exception:
                    st.info("ℹ️ No started trips match the selected date/time range and filters.")

            st.markdown("---")
            st.markdown("##### 📌 Summary")

            total_km = float(duty_filtered["_km"].sum())
            run_hrs_total = float(duty_filtered["_duration_hrs"].sum())  # sum of trip run times
            # Duty time = first start -> last end per driver per day (gaps included)
            total_duration_hrs = float(
                build_duty_summary(duty_filtered, "Driver + Date (daily)", duty_gap_hours)["Total Duty Hours"].sum()
            ) if not duty_filtered.empty else 0.0
            run_h = int(run_hrs_total)
            run_m = int(round((run_hrs_total - run_h) * 60))
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

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("🛣️ Total Distance", f"{total_km:.1f} KM")
            k2.metric("⏱️ Total Duty Time", f"{duration_h}h {duration_m}m",
                      help="First trip start to last trip end, per driver per day. Breaks with no trip longer than the limit above are not counted.")
            k3.metric("🚙 Run Time (trips only)", f"{run_h}h {run_m}m")
            k4.metric("🚗 Total Trips", total_trips)
            k5.metric("👨‍✈️ Active Driver / Vehicle", f"{active_driver_label} / {active_vehicle_label}")

            # -------------------------------------------------------------
            # DUTY SUMMARY — trips, KM and duty hours, grouped as chosen in
            # "Group summary by" (default: Driver + Date, so every row shows
            # which date the hours belong to, plus First Start / Last End
            # times in 12-hour AM/PM).
            # -------------------------------------------------------------
            st.markdown("---")
            st.markdown(f"##### 👨‍✈️ Duty Summary — {duty_group_option} (কে কত ঘণ্টা ডিউটি করেছে)")

            driver_summary_df = build_duty_summary(duty_filtered, duty_group_option, duty_gap_hours)
            if driver_summary_df.empty:
                st.info("No trips match the selected filters.")
            else:
                st.dataframe(driver_summary_df, use_container_width=True, hide_index=True,
                             height=min(380, 45 + 35 * len(driver_summary_df)))

            # -------------------------------------------------------------
            # DRIVER-WISE DAILY RUN KM — driver x date matrix with Total KM,
            # per-day average, total hours and average hours per day.
            # -------------------------------------------------------------
            matrix_names = (
                duty_drivers_df["driver_name"].tolist()
                if duty_driver_filter == "All Drivers" and not duty_drivers_df.empty else []
            )
            matrix_df = build_driver_daily_matrix(duty_filtered, matrix_names, duty_gap_hours)

            st.markdown("---")
            st.markdown("##### 🗓️ Driver-wise Daily Duty & Run KM (TWH | SWH | Km)")
            st.caption(
                "**TWH** = total working (duty) hours, first trip start → last trip end that day (a break with no trip longer than the limit above is not counted)  |  "
                "**SWH** = trip running hours  |  **Km** = distance run. "
                "Average = total ÷ number of date columns. Date = the day the trip started."
            )
            if matrix_df.empty:
                st.info("No data for the selected range.")
            else:
                st.dataframe(matrix_df, use_container_width=True,
                             height=min(420, 45 + 35 * len(matrix_df)))

            # -------------------------------------------------------------
            # STEP 5 — Detailed table:
            # [Vehicle No, Driver Name, Start Time, End Time, Start KM,
            #  End KM, Total KM, Duty Duration (Hours), Route / Purpose,
            #  Admin Note]
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
                    "Start Time": duty_filtered["_start_dt"].dt.strftime("%Y-%m-%d %I:%M %p").map(drop_hour_zero),
                    "End Time": duty_filtered["_end_dt"].dt.strftime("%Y-%m-%d %I:%M %p").map(drop_hour_zero),
                    "Start KM": duty_filtered["_start_km"].astype(float).round(1),
                    "End KM": duty_filtered["_end_km"].astype(float).round(1),
                    "Total KM": duty_filtered["_km"].round(1),
                    "Duty Duration (Hrs)": duty_filtered["_duration_hrs"],
                    "Route / Purpose": duty_filtered["destination"].fillna("").astype(str)
                                        + " — " + duty_filtered["purpose"].fillna("").astype(str),
                    "Admin Note": (
                        duty_filtered["admin_note"].fillna("").astype(str)
                        if "admin_note" in duty_filtered.columns else ""
                    ),
                }).reset_index(drop=True)
                st.dataframe(detail_display, use_container_width=True, hide_index=True, height=340)

            # -------------------------------------------------------------
            # STEP 6 — Multi-format export: CSV / Excel / PDF.
            # summary_metrics feeds both the Excel "Summary" sheet and the
            # PDF's KPI block — plug in any additional metrics here as needed.
            # -------------------------------------------------------------
            summary_metrics = {
                "Date/Time Range": drop_hour_zero(range_start.strftime('%Y-%m-%d %I:%M %p')) + " to " + drop_hour_zero(range_end.strftime('%Y-%m-%d %I:%M %p')),
                "Vehicle Filter": duty_vehicle_filter,
                "Driver Filter": duty_driver_filter,
                "Department Filter": ", ".join(duty_dept_filter) if duty_dept_filter else "All",
                "Status Filter": ", ".join(duty_status_filter) if duty_status_filter else "All",
                "Route / Purpose Filter": duty_text_filter.strip() or "None",
                "Summary Grouped By": duty_group_option,
                "Log Sorted By": duty_sort_option,
                "Total Distance (KM)": f"{total_km:.1f}",
                "Total Duty Time (first start to last end, long idle breaks excluded)": f"{duration_h}h {duration_m}m",
                "Run Time (trips only)": f"{run_h}h {run_m}m",
                "Total Trips": total_trips,
                "Active Driver": active_driver_label,
                "Assigned Vehicle": active_vehicle_label,
            }
            filters_summary_text = (
                f"Range: {summary_metrics['Date/Time Range']} | Vehicle: {duty_vehicle_filter} | "
                f"Driver: {duty_driver_filter} | Department: {summary_metrics['Department Filter']} | "
                f"Status: {summary_metrics['Status Filter']} | Grouped by: {duty_group_option}"
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
                duty_excel_bytes = build_duty_tracker_excel(detail_display, summary_metrics, driver_summary_df, matrix_df)
                st.download_button(
                    "⬇️ Download Excel (.xlsx)", data=duty_excel_bytes, file_name="duty_tracker_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with e3:
                duty_pdf_bytes = build_duty_tracker_pdf(detail_display, summary_metrics, filters_summary_text, driver_summary_df, matrix_df)
                st.download_button(
                    "⬇️ Download PDF (.pdf)", data=duty_pdf_bytes, file_name="duty_tracker_report.pdf",
                    mime="application/pdf", use_container_width=True,
                )

    # ---------------- Management Dashboard (NEW) ----------------
    if tab_management:
        # Admins get the same Executive/Management overview as the
        # standalone 'executive' role, reusing the exact same function so
        # both roles always see identical KPI numbers for the same filters.
        render_management_dashboard(df_all)

# =========================================================
# 11. FALLBACK — unrecognized role
# =========================================================
else:
    st.error("⚠️ Your account role is not recognized. Please contact the Admin.")
