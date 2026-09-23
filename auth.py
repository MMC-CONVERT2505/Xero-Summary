


"""
MMC Convert - Xero Conversion Summary
---------------------------------------------------------------
Flow:
  GET  /                -> Login page (Xero logo + "Login with Xero")
  GET  /login             -> redirects to Xero consent screen
  GET  /callback           -> exchanges code for tokens, saves to DB, redirects to /name-entry
  GET  /name-entry         -> ask user to confirm/edit the display name for this connection
  POST /name-entry         -> saves display name, redirects to /dashboard
  GET  /dashboard          -> single page, no sidebar. Date range selector + "Generate Summary"
  POST /logout             -> clears all saved tokens from the DB, redirects to /
  GET  /api/org-info       -> lightweight call, populates header on page load
  GET  /api/summary        -> called via JS when "Generate Summary" is clicked; returns JSON
  POST /api/export-excel   -> generates and returns an .xlsx file of the last summary

Install:
  pip install fastapi uvicorn httpx python-dotenv jinja2 openpyxl python-multipart --break-system-packages
Run:
  uvicorn auth:app --reload --port 8000
"""

import os
import io
import re
import time
import uuid
import random
import asyncio
import sqlite3
import secrets
from datetime import date, datetime
import httpx
from fastapi import FastAPI, Request, HTTPException, Query, Body, Form
from fastapi.responses import RedirectResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font

load_dotenv()

CLIENT_ID = os.getenv("XERO_CLIENT_ID")
CLIENT_SECRET = os.getenv("XERO_CLIENT_SECRET")
REDIRECT_URI = os.getenv("XERO_REDIRECT_URI")
SCOPES = os.getenv("XERO_SCOPES")

AUTHORIZE_URL = "https://login.xero.com/identity/connect/authorize"
TOKEN_URL = "https://identity.xero.com/connect/token"
CONNECTIONS_URL = "https://api.xero.com/connections"
API_BASE = "https://api.xero.com/api.xro/2.0"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "xero_tokens.db")

# Path prefix this app is served under, e.g. "/xero-summary" when running
# behind nginx alongside other tools on the same domain. Empty string means
# it's served at the domain root (local dev / its own subdomain).
BASE_PATH = os.getenv("APP_BASE_PATH", "").rstrip("/")

# Temporary in-memory holder for token_data while user picks an org on
# /select-org. Keyed by a random "pending" id, cleared once consumed.
_pending_selections: dict = {}

app = FastAPI()
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))
templates.env.globals["base_path"] = BASE_PATH
# --------------------------------------------------------------------
# SQLite persistence layer
# --------------------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS xero_tokens (
            tenant_id     TEXT PRIMARY KEY,
            tenant_name   TEXT,
            display_name  TEXT,
            access_token  TEXT NOT NULL,
            refresh_token TEXT NOT NULL,
            expires_at    REAL NOT NULL,
            updated_at    REAL NOT NULL
        )
        """
    )
    try:
        conn.execute("ALTER TABLE xero_tokens ADD COLUMN display_name TEXT")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


init_db()


def save_token(tenant_id, tenant_name, access_token, refresh_token, expires_in, display_name=None):
    conn = sqlite3.connect(DB_PATH)
    existing = conn.execute(
        "SELECT display_name FROM xero_tokens WHERE tenant_id = ?", (tenant_id,)
    ).fetchone()
    final_display_name = display_name if display_name is not None else (existing[0] if existing else None)

    conn.execute(
        """
        INSERT INTO xero_tokens (tenant_id, tenant_name, display_name, access_token, refresh_token, expires_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(tenant_id) DO UPDATE SET
            tenant_name=excluded.tenant_name,
            display_name=excluded.display_name,
            access_token=excluded.access_token,
            refresh_token=excluded.refresh_token,
            expires_at=excluded.expires_at,
            updated_at=excluded.updated_at
        """,
        (tenant_id, tenant_name, final_display_name, access_token, refresh_token, time.time() + expires_in, time.time()),
    )
    conn.commit()
    conn.close()


def update_display_name(tenant_id: str, display_name: str):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "UPDATE xero_tokens SET display_name = ? WHERE tenant_id = ?",
        (display_name, tenant_id),
    )
    conn.commit()
    conn.close()


def get_token_row(tenant_id):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT * FROM xero_tokens WHERE tenant_id = ?", (tenant_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def list_token_rows():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT tenant_id, tenant_name, display_name, updated_at FROM xero_tokens").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def clear_all_tokens():
    """Wipes every saved connection - used on logout so the next login starts fresh."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM xero_tokens")
    conn.commit()
    conn.close()


def delete_token_row(tenant_id: str):
    """Removes a single tenant's saved connection from the local DB."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM xero_tokens WHERE tenant_id = ?", (tenant_id,))
    conn.commit()
    conn.close()


# --------------------------------------------------------------------
# Rate-limit aware GET wrapper
# --------------------------------------------------------------------

async def xero_get(client: httpx.AsyncClient, url: str, headers: dict, params: dict = None, max_retries: int = 5):
    for _ in range(max_retries):
        resp = await client.get(url, headers=headers, params=params)
        if resp.status_code == 429:
            # Xero sends two very different flavours of 429 under the same
            # status code: a short per-minute throttle (fine to sleep and
            # retry) and the app's 5000-calls/day limit being exhausted,
            # where Retry-After can be tens of thousands of seconds (hours).
            # Blindly sleeping for that long makes a live request look
            # "hung" for the rest of the day, so fail fast instead and let
            # the caller surface a clear message.
            if resp.headers.get("X-Rate-Limit-Problem") == "day":
                retry_after = int(resp.headers.get("Retry-After", "0"))
                hours = round(retry_after / 3600, 1)
                raise RuntimeError(
                    f"Xero daily API limit reached for this app (5000 calls/day). "
                    f"Try again in about {hours} hours (resets {resp.headers.get('Expires', 'later today/tomorrow')})."
                )
            retry_after = int(resp.headers.get("Retry-After", "5"))
            # Random jitter (0-3s on top of Xero's own Retry-After) so
            # concurrent callers that all got rate-limited at the same
            # moment don't all wake up and retry in the same instant,
            # immediately re-triggering the same 429 for everyone (a
            # thundering herd that can burn through every retry and still
            # fail, even though the request itself was fine).
            await asyncio.sleep(retry_after + 1 + random.uniform(0, 3))
            continue
        if resp.status_code in (502, 503, 504):
            # Transient failures on Xero's own infrastructure (gateway
            # timeout/unavailable), not something wrong with our request -
            # Xero sends no Retry-After for these, so back off a short fixed
            # delay with jitter and try again rather than killing the whole
            # run over what's usually a momentary blip.
            await asyncio.sleep(3 + random.uniform(0, 3))
            continue
        resp.raise_for_status()
        return resp
    resp.raise_for_status()
    return resp


def count_taxed_line_items(line_items) -> int:
    """Counts LineItems/JournalLines with actual tax charged
    (TaxAmount > 0) - NOT based on TaxType, since some TaxTypes (e.g.
    "EXEMPTEXPENSES", "BASEXCLUDED") are assigned even when TaxAmount is
    0.00, which would wrongly count as "taxed" if we went by TaxType
    alone. Kept for logging/visibility only - no longer used to scale
    the Total Lines buffer (see document_line_buffer below for why)."""
    return sum(
        1 for li in (line_items or [])
        if (li.get("TaxAmount") or 0) > 0
    )


def document_line_buffer(line_items) -> int:
    """Total Lines approximation buffer for ONE document (an Invoice,
    Bill, Credit Note, Bill Credit, BankTransaction, Overpayment, or
    Prepayment) - replacing the old approach of adding +2 for EVERY
    taxed line item.

    Why the old approach was wrong: it scaled with how many LINE ITEMS a
    document had, but the GL-level postings it was trying to approximate
    - a tax line, and a control-account (AR/AP/bank) line - each post
    ONCE PER DOCUMENT, not once per line item. On a file where invoices
    average ~1.7 line items each, that distinction barely shows. On a
    file where invoices average ~11.7 line items each (all commonly
    taxed), the old formula added roughly 2 lines PER LINE ITEM instead
    of 2 lines per DOCUMENT - overshooting the true total by roughly 2x.

    New rule, applied once per document regardless of how many line
    items it has:
      - has >=1 line item, at least one with TaxAmount > 0 -> +2
        (1 tax-account line + 1 control-account line)
      - has >=1 line item, none with TaxAmount > 0          -> +1
        (just the control-account line)
      - has no line items                                    -> +0

    "Taxed" is decided by TaxAmount > 0, NOT by TaxType - some TaxTypes
    (e.g. "EXEMPTEXPENSES", "BASEXCLUDED") get assigned even when
    TaxAmount is 0.00, and a 0.00 tax amount never posts a real GL tax
    line, so TaxType alone would wrongly flag those as taxed.
    """
    items = line_items or []
    if not items:
        return 0
    has_tax = any((li.get("TaxAmount") or 0) > 0 for li in items)
    return 2 if has_tax else 1


def valid_line_items(line_items):
    """Drops LineItems/JournalLines that don't represent real money
    movement - no "LineAmount" field at all, or a LineAmount of 0 -
    before they're counted, taxed-checked, or fed into
    document_line_buffer(). Returns (kept_items, missing_amount_count,
    zero_amount_count) so callers can log exactly how many lines were
    ignored and why."""
    kept = []
    missing = 0
    zero = 0
    for li in (line_items or []):
        if "LineAmount" not in li:
            missing += 1
            continue
        if (li.get("LineAmount") or 0) == 0:
            zero += 1
            continue
        kept.append(li)
    return kept, missing, zero


def parse_xero_date(d: str):
    """Handles both ISO ('2024-01-15T00:00:00') and the older
    Microsoft-JSON format ('/Date(1705276800000+0000)/')."""
    if not d:
        return None
    if d.startswith("/Date("):
        millis = int(d.split("(")[1].split("+")[0].split("-")[0].rstrip(")/"))
        return datetime.utcfromtimestamp(millis / 1000).date()
    return datetime.strptime(d.split("T")[0], "%Y-%m-%d").date()


# --------------------------------------------------------------------
# OAuth flow
# --------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def login_page(request: Request):
    return templates.TemplateResponse(request, "login.html", {})


@app.get("/login")
def login(switch_account: bool = False):
    state = secrets.token_urlsafe(16)
    params = {
        "response_type": "code",
        "client_id": CLIENT_ID,
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
        "state": state,
    }
    # Xero caches the identity session cookie in the browser. Without a
    # `prompt` param it silently reuses that session and skips straight to
    # the previously-used account. Passing prompt=login/select_account
    # forces Xero's account chooser so the user can pick a different email.
    if switch_account:
        params["prompt"] = "login"  # use "select_account" instead if Xero shows a multi-account picker for the user's tenant setup

    url = f"{AUTHORIZE_URL}?{httpx.QueryParams(params)}"
    return RedirectResponse(url)


@app.get("/callback")
async def callback(request: Request):
    code = request.query_params.get("code")
    error = request.query_params.get("error")

    if error:
        raise HTTPException(400, f"Xero returned error: {error}")
    if not code:
        raise HTTPException(400, "No authorization code returned")

    async with httpx.AsyncClient() as client:
        resp = await client.post(
            TOKEN_URL,
            data={"grant_type": "authorization_code", "code": code, "redirect_uri": REDIRECT_URI},
            auth=(CLIENT_ID, CLIENT_SECRET),
        )
    if resp.status_code != 200:
        raise HTTPException(
            400,
            f"Token exchange failed: {resp.text}. Go to / and login again "
            f"(don't reload this callback URL — the code is single-use)."
        )
    token_data = resp.json()

    async with httpx.AsyncClient() as client:
        conn_resp = await client.get(
            CONNECTIONS_URL, headers={"Authorization": f"Bearer {token_data['access_token']}"}
        )
    conn_resp.raise_for_status()
    tenants = conn_resp.json()
    print(f"[LOGIN] Xero returned {len(tenants)} organisations")

    if not tenants:
        raise HTTPException(400, "No Xero organisation connected")

    if len(tenants) == 1:
        tenant_id = tenants[0]["tenantId"]
        tenant_name = tenants[0]["tenantName"]
        save_token(
            tenant_id=tenant_id,
            tenant_name=tenant_name,
            access_token=token_data["access_token"],
            refresh_token=token_data["refresh_token"],
            expires_in=token_data["expires_in"],
        )
        return RedirectResponse(url=f"{BASE_PATH}/name-entry?tenant_id={tenant_id}")

    pending_id = secrets.token_urlsafe(16)
    _pending_selections[pending_id] = {"token_data": token_data, "tenants": tenants}
    return templates.TemplateResponse(
        request, "select_org.html", {"pending": pending_id, "tenants": tenants}
    )


@app.post("/select-org")
async def select_org(pending: str = Form(...), tenant_id: str = Form(...)):
    entry = _pending_selections.pop(pending, None)
    if not entry:
        raise HTTPException(400, "Selection expired, go to / and login again")
    token_data = entry["token_data"]
    tenants = entry["tenants"]
    chosen = next((t for t in tenants if t["tenantId"] == tenant_id), None)
    if not chosen:
        raise HTTPException(400, "Invalid organisation selected")
    save_token(
        tenant_id=chosen["tenantId"],
        tenant_name=chosen["tenantName"],
        access_token=token_data["access_token"],
        refresh_token=token_data["refresh_token"],
        expires_in=token_data["expires_in"],
    )
    tid = chosen["tenantId"]
    return RedirectResponse(url=f"{BASE_PATH}/name-entry?tenant_id={tid}", status_code=303)


@app.get("/name-entry", response_class=HTMLResponse)
async def name_entry(request: Request, tenant_id: str = None):
    if not tenant_id:
        return RedirectResponse(url=f"{BASE_PATH}/")
    row = get_token_row(tenant_id)
    if not row:
        return RedirectResponse(url=f"{BASE_PATH}/")
    return templates.TemplateResponse(
        request,
        "name_entry.html",
        {"tenant_id": tenant_id, "tenant_name": row["tenant_name"]},
    )

@app.post("/name-entry")
async def name_entry_submit(tenant_id: str = Form(...), display_name: str = Form(...)):
    row = get_token_row(tenant_id)
    if not row:
        raise HTTPException(400, "Invalid tenant")
    update_display_name(tenant_id, display_name.strip())
    return RedirectResponse(url=f"{BASE_PATH}/dashboard?tenant_id={tenant_id}", status_code=303)


async def get_valid_access_token(tenant_id: str) -> str:
    row = get_token_row(tenant_id)
    if not row:
        raise HTTPException(401, "No token found for this tenant. Login again via /")

    if time.time() > row["expires_at"] - 60:
        async with httpx.AsyncClient() as client:
            resp = await client.post(
                TOKEN_URL,
                data={"grant_type": "refresh_token", "refresh_token": row["refresh_token"]},
                auth=(CLIENT_ID, CLIENT_SECRET),
            )
        resp.raise_for_status()
        new_token = resp.json()
        save_token(
            tenant_id=tenant_id,
            tenant_name=row["tenant_name"],
            access_token=new_token["access_token"],
            refresh_token=new_token["refresh_token"],
            expires_in=new_token["expires_in"],
        )
        return new_token["access_token"]

    return row["access_token"]


# --------------------------------------------------------------------
# Dashboard page (just renders the shell; no data fetched on load)
# --------------------------------------------------------------------

@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request, tenant_id: str = None):
    if not tenant_id:
        return RedirectResponse(url=f"{BASE_PATH}/")
    row = get_token_row(tenant_id)
    if not row:
        return RedirectResponse(url=f"{BASE_PATH}/")
    if not row.get("display_name"):
        return RedirectResponse(url=f"{BASE_PATH}/name-entry?tenant_id={tenant_id}")
    return templates.TemplateResponse(
        request,
        "dashboard.html",
        {"tenant_id": tenant_id, "tenant_name": row["display_name"]},
    )


@app.post("/logout")
async def logout(tenant_id: str = Form(...)):
    """
    Fully disconnects this Xero organisation:
      1. Revokes the connection on Xero's side (DELETE /connections/{id})
      2. Removes the saved tokens from the local DB
    This guarantees the next login always shows a fresh Xero consent screen,
    with no risk of stale/mismatched organisation data.
    """
    try:
        access_token = await get_valid_access_token(tenant_id)

        async with httpx.AsyncClient() as client:
            conn_resp = await client.get(
                CONNECTIONS_URL,
                headers={"Authorization": f"Bearer {access_token}"},
            )
        conn_resp.raise_for_status()
        connections = conn_resp.json()

        connection = next((c for c in connections if c.get("tenantId") == tenant_id), None)

        if connection:
            connection_id = connection["id"]
            async with httpx.AsyncClient() as client:
                del_resp = await client.delete(
                    f"{CONNECTIONS_URL}/{connection_id}",
                    headers={"Authorization": f"Bearer {access_token}"},
                )
            print(f"[LOGOUT] Xero disconnect status: {del_resp.status_code}")
        else:
            print(f"[LOGOUT] No matching Xero connection found for tenant_id={tenant_id}")

    except Exception as e:
        print(f"[LOGOUT] Error while revoking Xero connection: {e}")

    delete_token_row(tenant_id)
    return RedirectResponse(url=f"{BASE_PATH}/", status_code=303)


@app.get("/api/org-info")
async def api_org_info(tenant_id: str):
    """Lightweight call - just Organisation info, used to populate header on page load."""
    access_token = await get_valid_access_token(tenant_id)
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Xero-tenant-id": tenant_id,
        "Accept": "application/json",
    }
    async with httpx.AsyncClient(timeout=30) as client:
        org_resp = await xero_get(client, f"{API_BASE}/Organisation", headers)
    org = org_resp.json().get("Organisations", [{}])[0]

    row = get_token_row(tenant_id)
    display_name = row.get("display_name") if row else None

    return JSONResponse({
        "company_name": display_name or org.get("Name"),
        "financial_year_end_day": org.get("FinancialYearEndDay"),
        "financial_year_end_month": org.get("FinancialYearEndMonth"),
    })


# --------------------------------------------------------------------
# Helper: Financial Year start calculation
# --------------------------------------------------------------------

def compute_fy_start(fy_end_day: int, fy_end_month: int, today: date) -> date:
    fy_end_this_year = date(today.year, fy_end_month, fy_end_day)

    if today <= fy_end_this_year:
        fy_start_year = today.year - 1
    else:
        fy_start_year = today.year

    fy_end_prev = date(fy_start_year, fy_end_month, fy_end_day)
    if fy_end_month == 12 and fy_end_day == 31:
        return date(fy_start_year + 1, 1, 1)
    from datetime import timedelta
    return fy_end_prev + timedelta(days=1)


def subtract_years(d: date, years: int) -> date:
    try:
        return d.replace(year=d.year - years)
    except ValueError:
        return d.replace(year=d.year - years, day=28)


# --------------------------------------------------------------------
# "Inception to Date" - finds the actual earliest transaction date in
# this org, so the frontend's "Inception to Date" range button doesn't
# have to guess a start date. Cheap by design: for each document type,
# only page 1 sorted oldest-first (order=Date ASC) is fetched - one
# record's Date is all that's needed per type, not the full dataset.
# The overall inception date is the earliest across every type checked.
# --------------------------------------------------------------------
@app.get("/api/inception-date")
async def api_inception_date(tenant_id: str):
    access_token = await get_valid_access_token(tenant_id)
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Xero-tenant-id": tenant_id,
        "Accept": "application/json",
    }
    # Every endpoint that has its own "Date" field and can realistically
    # be the oldest record in an org - covers the vast majority of cases
    # without fetching every transaction type that exists.
    endpoints = [
        "Invoices", "CreditNotes", "BankTransactions", "ManualJournals",
        "Overpayments", "Prepayments",
    ]
    earliest = None
    earliest_source = None
    async with httpx.AsyncClient(timeout=30) as client:
        for ep in endpoints:
            try:
                resp = await xero_get(
                    client, f"{API_BASE}/{ep}", headers,
                    params={"page": 1, "order": "Date ASC"},
                )
                data = resp.json().get(ep, [])
                if data:
                    d = parse_xero_date(data[0].get("Date"))
                    if d and (earliest is None or d < earliest):
                        earliest = d
                        earliest_source = ep
            except Exception as e:
                print(f"[INCEPTION] {ep} lookup failed (skipping): {e}")
                continue

    if earliest is None:
        # No transactions of any checked type found - fall back to a
        # safe, obviously-not-real placeholder rather than failing the
        # request outright.
        earliest = date(2000, 1, 1)
        earliest_source = None

    print(f"[INCEPTION] tenant_id={tenant_id} inception_date={earliest} (source={earliest_source})")
    return JSONResponse({"inception_date": earliest.isoformat(), "source": earliest_source})


# --------------------------------------------------------------------
# Main summary endpoint - called only when "Generate Summary" is clicked
# --------------------------------------------------------------------

async def generate_summary_data(
    tenant_id: str,
    from_date: str = None,
    to_date: str = None,
    range_preset: str = None,  # "1ytd" | "2ytd" | "3ytd" | None (custom)
    skip_attachments: bool = False,  # True = skip the slow per-record attachment lookups entirely (Xero has no bulk-attachments endpoint, so this is the only real speed lever for large orgs)
    job_id: str = None,  # when set, every print() below also updates SUMMARY_JOBS[job_id]["progress"] for polling clients - see the shadowed print() just below
) -> dict:
    # Shadows the builtin print() for the rest of THIS function and every
    # nested async def inside it (Python closures resolve free variables
    # by lexical scope, not call order, so this applies retroactively to
    # every "print(...)" call already written below without touching any
    # of them individually - the whole summary generation already logs
    # its progress via print(), so this is what turns that into
    # pollable job progress for free instead of a parallel refactor).
    import builtins as _builtins

    def print(*args, **kwargs):
        _builtins.print(*args, **kwargs)
        if job_id and job_id in SUMMARY_JOBS:
            SUMMARY_JOBS[job_id]["progress"] = " ".join(str(a) for a in args)

    access_token = await get_valid_access_token(tenant_id)
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Xero-tenant-id": tenant_id,
        "Accept": "application/json",
    }

    async def refresh_headers():
        # Xero access tokens last ~30 min. Large orgs (thousands of
        # invoices/bank txns across dozens of paginated calls) can take
        # longer than that to fully fetch, so the token can expire mid-run.
        # get_valid_access_token() only actually hits Xero when the token
        # is close to expiring, so calling this every page is cheap.
        fresh_token = await get_valid_access_token(tenant_id)
        headers["Authorization"] = f"Bearer {fresh_token}"

    def log_excluded(category: str, excluded: list, id_field: str = "identifier"):
        # excluded: list of (identifier, status) tuples for records skipped
        # because they're DELETED/VOIDED - printed here so it's visible in
        # the console exactly which records didn't count, and why.
        deleted = [i for i, st in excluded if st == "DELETED"]
        voided = [i for i, st in excluded if st == "VOIDED"]
        print(f"[SUMMARY] {category} excluded = {len(excluded)}  (DELETED={len(deleted)}, VOIDED={len(voided)})")
        if deleted:
            print(f"[SUMMARY]   {category} DELETED {id_field}s: {deleted}")
        if voided:
            print(f"[SUMMARY]   {category} VOIDED {id_field}s: {voided}")

    print(f"\n[SUMMARY] Starting generation for tenant_id={tenant_id}")

    async with httpx.AsyncClient(timeout=60) as client:
        # --- Organisation info ---
        print("[SUMMARY] Fetching Organisation...")
        org_resp = await xero_get(client, f"{API_BASE}/Organisation", headers)
        org = org_resp.json().get("Organisations", [{}])[0]

        _row = get_token_row(tenant_id)
        _display_name = _row.get("display_name") if _row else None
        if _display_name:
            org["Name"] = _display_name

        base_currency = org.get("BaseCurrency")
        fy_end_day = org.get("FinancialYearEndDay")
        fy_end_month = org.get("FinancialYearEndMonth")

        today = date.today()

        # Attachment candidates are collected (endpoint, record_id) across
        # EVERY document type as they're fetched below - Invoices/Bills,
        # CreditNotes, BankTransactions, ManualJournals, PurchaseOrders,
        # Quotes, Overpayments and Prepayments (the full set of Xero
        # endpoints that support HasAttachments/the /Attachments
        # sub-resource) - and only actually hit ONCE, all together, in a
        # single high-concurrency batch at the very end (see
        # fetch_all_attachment_counts below the last document fetch).
        # Previously each type's attachments were fetched separately right
        # after that type's own document fetch, at a low concurrency (4) -
        # serialising the whole run behind N sequential low-concurrency
        # batches, which is what made attachment-fetching slow.
        all_attachment_candidates = []  # list of (endpoint, record_id)

        def add_attachment_candidates(endpoint: str, ids: list):
            all_attachment_candidates.extend((endpoint, i) for i in ids)

        # Multi-currency TRANSACTION count (distinct from multi_currency_accounts,
        # which counts bank/credit-card ACCOUNTS set up in a foreign currency).
        # A transaction counts here when its own CurrencyCode differs from the
        # org's BaseCurrency, tallied across every counted (non-VOIDED/DELETED)
        # Invoice, Bill, Credit Note, and Bank Transaction as they're fetched -
        # no extra API calls needed.
        multi_currency_txn_count = 0

        def note_currency(doc: dict):
            nonlocal multi_currency_txn_count
            cur = doc.get("CurrencyCode")
            if cur and cur != base_currency:
                multi_currency_txn_count += 1

        async def fetch_all_attachment_counts() -> int:
            if not all_attachment_candidates:
                return 0
            # Xero's own API guidance caps recommended concurrent
            # connections at ~5-10 - going much higher (this used to be 20)
            # makes every concurrent request hit the per-minute rate limit
            # at once, they all get the same Retry-After and all retry
            # together, hit the limit again, and so on - a thundering herd
            # that burns through xero_get's 5 retries and then silently
            # returns 0 for records that DO have attachments, undercounting
            # the true total. 8 concurrent + jittered retries below stays
            # inside that safe range while still being noticeably faster
            # than the very conservative 5 this started at.
            sem = asyncio.Semaphore(8)
            failed = []

            async def fetch_one(endpoint, record_id):
                async with sem:
                    try:
                        att_resp = await xero_get(
                            client, f"{API_BASE}/{endpoint}/{record_id}/Attachments", headers,
                        )
                        return len(att_resp.json().get("Attachments", []))
                    except Exception as e:
                        failed.append((endpoint, record_id, str(e)))
                        return 0

            print(f"[SUMMARY] Fetching attachment counts for {len(all_attachment_candidates)} records across all document types...")
            results = await asyncio.gather(*[fetch_one(ep, i) for ep, i in all_attachment_candidates])
            if failed:
                print(
                    f"[SUMMARY] WARNING: {len(failed)} of {len(all_attachment_candidates)} attachment "
                    f"lookups failed (counted as 0 - total attachment count is an UNDERCOUNT). "
                    f"First few: {failed[:5]}"
                )
            return sum(results)

        # --- Resolve date range ---
        if range_preset in ("1ytd", "2ytd", "3ytd") and fy_end_day and fy_end_month:
            fy_start = compute_fy_start(fy_end_day, fy_end_month, today)
            years_back = int(range_preset[0])
            start = subtract_years(fy_start, years_back)
            end = today
        elif from_date and to_date:
            start = datetime.strptime(from_date, "%Y-%m-%d").date()
            end = datetime.strptime(to_date, "%Y-%m-%d").date()
        else:
            if fy_end_day and fy_end_month:
                fy_start = compute_fy_start(fy_end_day, fy_end_month, today)
                start = subtract_years(fy_start, 1)
            else:
                start = date(today.year - 1, 1, 1)
            end = today

        date_filter = f"Date >= DateTime({start.year},{start.month},{start.day}) AND Date <= DateTime({end.year},{end.month},{end.day})"
        # Deleted/Voided records never affect the ledger and should never be
        # migrated - excluded directly in the API "where" query (not just
        # client-side after fetch) so they're never even fetched. Different
        # endpoints support different subsets of these statuses (e.g.
        # Overpayments/Prepayments can be VOIDED but never DELETED via this
        # API), so three variants are built here for the endpoints below.
        date_filter_no_deleted_voided = f'{date_filter} AND Status!="DELETED" AND Status!="VOIDED"'
        date_filter_no_deleted = f'{date_filter} AND Status!="DELETED"'
        date_filter_no_voided = f'{date_filter} AND Status!="VOIDED"'

        # --- Currencies ---
        print("[SUMMARY] Fetching Currencies...")
        cur_resp = await xero_get(client, f"{API_BASE}/Currencies", headers)
        enabled_currencies = [c.get("Code") for c in cur_resp.json().get("Currencies", [])]

        # --- Accounts (Chart of Accounts) ---
        print("[SUMMARY] Fetching Accounts...")
        acc_resp = await xero_get(client, f"{API_BASE}/Accounts", headers)
        all_fetched_accounts = acc_resp.json().get("Accounts", [])
        accounts = [a for a in all_fetched_accounts if a.get("Status") == "ACTIVE"]
        print(
            f"[SUMMARY] Accounts fetched={len(all_fetched_accounts)}, "
            f"ACTIVE={len(accounts)}, "
            f"excluded(non-ACTIVE)={len(all_fetched_accounts) - len(accounts)}"
        )

        bank_count = sum(1 for a in accounts if a.get("BankAccountType") == "BANK")
        cc_count = sum(1 for a in accounts if a.get("BankAccountType") == "CREDITCARD")
        multi_currency_accounts = [
            {"name": a.get("Name"), "code": a.get("Code"), "currency": a.get("CurrencyCode")}
            for a in accounts
            if a.get("CurrencyCode") and a.get("CurrencyCode") != base_currency
        ]

        bank_account_ids = {
            a["AccountID"] for a in accounts if a.get("BankAccountType") == "BANK"
        }
        credit_card_account_ids = {
            a["AccountID"] for a in accounts if a.get("BankAccountType") == "CREDITCARD"
        }

        # Debug: show exactly how Xero classifies each bank-like account,
        # so we can see why the bank vs credit-card split differs from the
        # accounting firm's own report.
        print(f"[SUMMARY] Accounts: total={len(accounts)}, BANK={bank_count}, CREDITCARD={cc_count}")
        for a in accounts:
            if a.get("BankAccountType") in ("BANK", "CREDITCARD"):
                print(
                    f"    -> {a.get('Name')!r}  type={a.get('BankAccountType')}  "
                    f"currency={a.get('CurrencyCode')}  AccountID={a.get('AccountID')}"
                )

        # --- Tracking Categories (Class/Job/Location) ---
        print("[SUMMARY] Fetching Tracking Categories...")
        track_resp = await xero_get(client, f"{API_BASE}/TrackingCategories", headers)
        tracking_categories = track_resp.json().get("TrackingCategories", [])

        # --- Items (Tracked Inventory) ---
        print("[SUMMARY] Fetching Items...")
        items_resp = await xero_get(client, f"{API_BASE}/Items", headers)
        items = items_resp.json().get("Items", [])
        tracked_items = [i for i in items if i.get("IsTrackedAsInventory")]

        # --- Payroll: Employees count + Payroll count (flattened Payslips
        # across every PayRun in the selected date range).
        #
        # Two Payroll API versions, on purpose (same finding as
        # fetch_payroll_only.py): payroll.xro/2.0 (camelCase) for
        # /Employees, but payroll.xro/1.0 (PascalCase, legacy) for
        # /PayRuns - the 2.0 PayRuns list 500s on its last page for some
        # orgs (confirmed genuine Xero-side bug), and 2.0's per-PayRun
        # detail endpoint 405s ("Method not allowed for the current
        # customer jurisdiction"). 1.0 has neither problem, and 1.0's
        # per-PayRun detail returns a "Payslips" array - flattening that
        # across every PayRun in range gives the "Payroll" count.
        PAYROLL_BASE_V1 = "https://api.xero.com/payroll.xro/1.0"
        payroll_headers = {"Authorization": f"Bearer {access_token}", "Xero-tenant-id": tenant_id, "Accept": "application/json"}
        print("[SUMMARY] Checking Payroll...")
        payroll_enabled = False
        employee_count = 0
        payroll_count = 0
        try:
            # Xero's Payroll API returns 200 OK on /Employees even for orgs
            # that have never touched Payroll (itemCount=0, empty list) - a
            # 200 status alone does NOT mean payroll is actually in use.
            # "enabled" is decided below from the real employee count.
            pr_resp = await client.get(f"{PAYROLL_BASE_V1.replace('1.0', '2.0')}/Employees", headers=payroll_headers, params={"page": 1})
            payroll_api_reachable = pr_resp.status_code == 200
        except Exception:
            payroll_api_reachable = False

        if payroll_api_reachable:
            try:
                # Employees - paginated, v2.0.
                page = 1
                page_count = 1
                while page <= page_count:
                    resp = await xero_get(client, "https://api.xero.com/payroll.xro/2.0/Employees", payroll_headers, params={"page": page})
                    body = resp.json()
                    page_count = (body.get("pagination") or {}).get("pageCount", page_count)
                    employee_count += len(body.get("employees", []))
                    page += 1
                print(f"[SUMMARY] Payroll Employees count = {employee_count}")

                # "Enabled" = actually has employees set up, not just a 200
                # from an API that returns 200 either way.
                payroll_enabled = employee_count > 0

                # PayRuns - paginated, v1.0 (see comment above for why).
                # Skipped entirely when there are no employees - an org with
                # no employees has never run payroll, so this would just be
                # empty pages burning API calls for nothing.
                all_payruns = []
                if payroll_enabled:
                    page = 1
                    while True:
                        resp = await xero_get(client, f"{PAYROLL_BASE_V1}/PayRuns", payroll_headers, params={"page": page})
                        data = resp.json().get("PayRuns", [])
                        if not data:
                            break
                        all_payruns.extend(data)
                        page += 1

                def parse_payroll_date(d):
                    if not d:
                        return None
                    if d.startswith("/Date("):
                        millis = int(d.split("(")[1].split("+")[0].split("-")[0].rstrip(")/"))
                        return datetime.utcfromtimestamp(millis / 1000).date()
                    return parse_xero_date(d)

                payruns_in_range = [
                    pr for pr in all_payruns
                    if (d := parse_payroll_date(pr.get("PaymentDate"))) and start <= d <= end
                ]
                print(f"[SUMMARY] PayRuns in range = {len(payruns_in_range)} (of {len(all_payruns)} total)")

                # Per-PayRun detail (v1.0) - flatten "Payslips" across every
                # PayRun in range to get the "Payroll" count.
                sem = asyncio.Semaphore(5)

                async def fetch_payrun_payslip_count(pr):
                    async with sem:
                        try:
                            resp = await xero_get(
                                client, f"{PAYROLL_BASE_V1}/PayRuns/{pr.get('PayRunID')}", payroll_headers,
                            )
                        except Exception:
                            return 0
                        detail_list = resp.json().get("PayRuns", [])
                        if not detail_list:
                            return 0
                        return len(detail_list[0].get("Payslips", []) or [])

                if payruns_in_range:
                    payslip_counts = await asyncio.gather(*[fetch_payrun_payslip_count(pr) for pr in payruns_in_range])
                    payroll_count = sum(payslip_counts)
                print(f"[SUMMARY] Payroll count (flattened Payslips across PayRuns in range) = {payroll_count} (added to Total Lines)")
            except Exception as e:
                print(f"[SUMMARY] Payroll count fetch failed (payroll enabled but detail unavailable): {e}")

        # --- Invoices: SINGLE PASS for both Invoice (ACCREC) and Bill (ACCPAY) ---
        # Previously this endpoint was fetched 3 separate times (once for
        # attachments, once for ACCREC, once for ACCPAY). Attachments counting
        # has been removed entirely (Xero has no bulk-attachments endpoint, so
        # it wasn't reducible - it's dropped instead), and ACCREC/ACCPAY are
        # now counted together in one pass, cutting this endpoint's page
        # fetches from 3x down to 1x.
        async def fetch_invoices_and_bills():
            invoice_total = 0
            bill_total = 0
            invoice_line_items_total = 0
            bill_line_items_total = 0
            invoice_taxed_line_items_total = 0
            bill_taxed_line_items_total = 0
            invoice_doc_buffer_total = 0  # NEW: per-document buffer (see document_line_buffer)
            bill_doc_buffer_total = 0     # NEW
            # invoice_tax_lookup: InvoiceID -> has this invoice/bill got at
            # least one taxed LineItem. Built here (no extra API call) so
            # count_payments() can look up whether the Invoice/Bill a
            # Payment is settling was taxed - the /Payments list endpoint
            # itself only returns an empty nested Invoice.LineItems, so
            # this is the only way to know without a second fetch.
            invoice_tax_lookup = {}
            invoice_status_breakdown = {}
            bill_status_breakdown = {}
            excluded_invoices = []  # (InvoiceNumber/InvoiceID, Status) for DELETED/VOIDED ACCREC
            excluded_bills = []     # same, for ACCPAY
            invoice_missing_amount_count = 0  # LineItems with no LineAmount field at all
            invoice_zero_amount_count = 0     # LineItems with LineAmount == 0
            bill_missing_amount_count = 0
            bill_zero_amount_count = 0
            attachment_candidates = []  # (InvoiceID) for records with HasAttachments=true
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Invoices - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Invoices", headers,
                    params={"page": page, "where": date_filter_no_deleted_voided, "summaryOnly": "false"},
                )
                data = resp.json().get("Invoices", [])
                if not data:
                    print(f"[SUMMARY] Invoices - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] Invoices - page {page} returned {len(data)} records")

                for inv in data:
                    st = inv.get("Status", "UNKNOWN")
                    inv_type = inv.get("Type")
                  # Exclude VOIDED and DELETED - these never affect the
                    # ledger and should never be migrated. Everything else
                    # (AUTHORISED, PAID, DRAFT, SUBMITTED) is counted so this
                    # stays as close as possible to what Xero's own "All"
                    # invoices list shows.
                    kept_li, li_missing, li_zero = valid_line_items(inv.get("LineItems"))
                    if inv_type == "ACCREC":
                        invoice_status_breakdown[st] = invoice_status_breakdown.get(st, 0) + 1
                        if st not in ("VOIDED", "DELETED"):
                            invoice_missing_amount_count += li_missing
                            invoice_zero_amount_count += li_zero
                            invoice_total += 1
                            invoice_line_items_total += len(kept_li)
                            invoice_taxed_line_items_total += count_taxed_line_items(kept_li)
                            invoice_doc_buffer_total += document_line_buffer(kept_li)
                            note_currency(inv)
                        else:
                            excluded_invoices.append((inv.get("InvoiceNumber") or inv.get("InvoiceID"), st))
                    elif inv_type == "ACCPAY":
                        bill_status_breakdown[st] = bill_status_breakdown.get(st, 0) + 1
                        if st not in ("VOIDED", "DELETED"):
                            bill_missing_amount_count += li_missing
                            bill_zero_amount_count += li_zero
                            bill_total += 1
                            bill_line_items_total += len(kept_li)
                            bill_taxed_line_items_total += count_taxed_line_items(kept_li)
                            bill_doc_buffer_total += document_line_buffer(kept_li)
                            note_currency(inv)
                        else:
                            excluded_bills.append((inv.get("InvoiceNumber") or inv.get("InvoiceID"), st))

                    if inv.get("InvoiceID") and inv_type in ("ACCREC", "ACCPAY") and st not in ("VOIDED", "DELETED"):
                        invoice_tax_lookup[inv["InvoiceID"]] = count_taxed_line_items(kept_li) > 0

                    # Attachment detection piggybacks on this same pass - no
                    # extra Invoices fetch needed. Xero has no bulk-attachments
                    # endpoint, so the only unavoidable extra calls are one
                    # per record that actually has an attachment (below).
                    if inv.get("HasAttachments") and inv.get("InvoiceID"):
                        attachment_candidates.append(inv["InvoiceID"])

                page += 1
                await asyncio.sleep(0.05)

            print(f"[SUMMARY] Invoice (ACCREC) count = {invoice_total}  (by status: {invoice_status_breakdown})")
            print(f"[SUMMARY] Bill (ACCPAY) count = {bill_total}  (by status: {bill_status_breakdown})")
            log_excluded("Invoice (ACCREC)", excluded_invoices)
            log_excluded("Bill (ACCPAY)", excluded_bills)
            print(f"[SUMMARY] Invoice LineItems total = {invoice_line_items_total} (taxed={invoice_taxed_line_items_total}, doc_buffer={invoice_doc_buffer_total})")
            print(f"[SUMMARY] Invoice LineItems ignored = {invoice_missing_amount_count + invoice_zero_amount_count}  (no LineAmount field={invoice_missing_amount_count}, LineAmount=0={invoice_zero_amount_count})")
            print(f"[SUMMARY] Bill LineItems ignored = {bill_missing_amount_count + bill_zero_amount_count}  (no LineAmount field={bill_missing_amount_count}, LineAmount=0={bill_zero_amount_count})")
            print(f"[SUMMARY] Bill LineItems total = {bill_line_items_total} (taxed={bill_taxed_line_items_total}, doc_buffer={bill_doc_buffer_total})")
            # Invoice TOTAL, same fix as Bill TOTAL below - each Invoice
            # line-item contributes 1 line, PLUS (per invoice, not per
            # line) +2 if that invoice has any taxed line (1 tax-account
            # line + 1 Accounts Receivable/"receivable" control-account
            # line) or +1 if untaxed-only. E.g. 5 invoices, 5 line items
            # total, all taxed -> 5 + (5 invoices x 2) = 15. Confirmed
            # against a real Xero Accounts Receivable report - each
            # invoice payment shows its own "Receivable Payment" row there,
            # same pattern as "Payable Payment" for Bills.
            print(f"[SUMMARY] Invoice TOTAL (LineItems + tax/receivable buffer) = {invoice_line_items_total} + {invoice_doc_buffer_total} = {invoice_line_items_total + invoice_doc_buffer_total}")
            # Bill TOTAL, spelled out: each Bill line-item contributes 1
            # line, PLUS (per bill, not per line) +2 if that bill has any
            # taxed line (1 tax-account line + 1 Accounts Payable/"payable"
            # control-account line) or +1 if untaxed-only. E.g. 5 bills,
            # 5 line items total, all taxed -> 5 + (5 bills x 2) = 15.
            print(f"[SUMMARY] Bill TOTAL (LineItems + tax/payable buffer) = {bill_line_items_total} + {bill_doc_buffer_total} = {bill_line_items_total + bill_doc_buffer_total}")
            print(f"[SUMMARY] Invoices/Bills with attachments to check = {len(attachment_candidates)}")
            add_attachment_candidates("Invoices", attachment_candidates)

            return (
                invoice_total, bill_total,
                invoice_line_items_total, bill_line_items_total,
                invoice_taxed_line_items_total, bill_taxed_line_items_total,
                invoice_doc_buffer_total, bill_doc_buffer_total,
                invoice_tax_lookup,
            )

        # --- CreditNotes: SINGLE PASS for both Credit Memo and Bill Credit ---
        async def fetch_credit_notes():
            memo_total = 0
            credit_total = 0
            memo_line_items_total = 0
            credit_line_items_total = 0
            memo_taxed_line_items_total = 0
            credit_taxed_line_items_total = 0
            memo_doc_buffer_total = 0    # NEW
            credit_doc_buffer_total = 0  # NEW
            memo_status_breakdown = {}
            credit_status_breakdown = {}
            excluded_memos = []
            excluded_credits = []
            # Allocations = the "Receivable/Payable Credit Note Allocation"
            # lines seen in Xero's Account Transactions report - each is a
            # separate GL posting made when a Credit Memo/Bill Credit is
            # applied against an Invoice/Bill, distinct from the credit
            # note's own LineItems. summaryOnly=false above means each
            # CreditNote already comes back with its "Allocations" array,
            # so no extra fetch is needed to count these.
            memo_allocation_count = 0
            credit_allocation_count = 0
            memo_missing_amount_count = 0
            memo_zero_amount_count = 0
            credit_missing_amount_count = 0
            credit_zero_amount_count = 0
            # credit_note_tax_lookup: CreditNoteID -> has this credit note got
            # at least one taxed LineItem. Same pattern as invoice_tax_lookup -
            # lets count_payments() below check whether the credit note an
            # ARCREDITPAYMENT/APCREDITPAYMENT refund is against was taxed,
            # without a second fetch (the /Payments endpoint's own nested
            # Invoice.LineItems always comes back empty).
            credit_note_tax_lookup = {}
            credit_note_attachment_candidates = []
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] CreditNotes - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/CreditNotes", headers,
                    params={"page": page, "where": date_filter_no_deleted_voided, "summaryOnly": "false"},
                )
                data = resp.json().get("CreditNotes", [])
                if not data:
                    print(f"[SUMMARY] CreditNotes - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] CreditNotes - page {page} returned {len(data)} records")

                for n in data:
                    st = n.get("Status", "UNKNOWN")
                    note_type = n.get("Type")
                    kept_li, li_missing, li_zero = valid_line_items(n.get("LineItems"))
                    if n.get("CreditNoteID") and note_type in ("ACCRECCREDIT", "ACCPAYCREDIT") and st not in ("VOIDED", "DELETED"):
                        credit_note_tax_lookup[n["CreditNoteID"]] = count_taxed_line_items(kept_li) > 0
                    if n.get("HasAttachments") and n.get("CreditNoteID"):
                        credit_note_attachment_candidates.append(n["CreditNoteID"])
                    if note_type == "ACCRECCREDIT":
                        memo_status_breakdown[st] = memo_status_breakdown.get(st, 0) + 1
                        if st not in ("VOIDED", "DELETED"):
                            memo_missing_amount_count += li_missing
                            memo_zero_amount_count += li_zero
                            memo_total += 1
                            memo_line_items_total += len(kept_li)
                            memo_taxed_line_items_total += count_taxed_line_items(kept_li)
                            memo_doc_buffer_total += document_line_buffer(kept_li)
                            memo_allocation_count += len(n.get("Allocations", []) or [])
                            note_currency(n)
                        else:
                            excluded_memos.append((n.get("CreditNoteNumber") or n.get("CreditNoteID"), st))
                    elif note_type == "ACCPAYCREDIT":
                        credit_status_breakdown[st] = credit_status_breakdown.get(st, 0) + 1
                        if st not in ("VOIDED", "DELETED"):
                            credit_missing_amount_count += li_missing
                            credit_zero_amount_count += li_zero
                            credit_total += 1
                            credit_line_items_total += len(kept_li)
                            credit_taxed_line_items_total += count_taxed_line_items(kept_li)
                            credit_doc_buffer_total += document_line_buffer(kept_li)
                            credit_allocation_count += len(n.get("Allocations", []) or [])
                            note_currency(n)
                        else:
                            excluded_credits.append((n.get("CreditNoteNumber") or n.get("CreditNoteID"), st))

                page += 1
                await asyncio.sleep(0.05)

            print(f"[SUMMARY] Credit Memo (ACCRECCREDIT) count = {memo_total}  (by status: {memo_status_breakdown})")
            print(f"[SUMMARY] Bill Credit (ACCPAYCREDIT) count = {credit_total}  (by status: {credit_status_breakdown})")
            log_excluded("Credit Memo (ACCRECCREDIT)", excluded_memos)
            log_excluded("Bill Credit (ACCPAYCREDIT)", excluded_credits)
            print(f"[SUMMARY] Credit Memo LineItems total = {memo_line_items_total} (taxed={memo_taxed_line_items_total}, doc_buffer={memo_doc_buffer_total})")
            print(f"[SUMMARY] Bill Credit LineItems total = {credit_line_items_total} (taxed={credit_taxed_line_items_total}, doc_buffer={credit_doc_buffer_total})")
            print(f"[SUMMARY] Credit Memo LineItems ignored = {memo_missing_amount_count + memo_zero_amount_count}  (no LineAmount field={memo_missing_amount_count}, LineAmount=0={memo_zero_amount_count})")
            print(f"[SUMMARY] Bill Credit LineItems ignored = {credit_missing_amount_count + credit_zero_amount_count}  (no LineAmount field={credit_missing_amount_count}, LineAmount=0={credit_zero_amount_count})")
            print(f"[SUMMARY] Receivable Credit Note Allocation count = {memo_allocation_count}")
            print(f"[SUMMARY] Payable Credit Note Allocation count = {credit_allocation_count}")
            print(f"[SUMMARY] CreditNotes with attachments to check = {len(credit_note_attachment_candidates)}")
            add_attachment_candidates("CreditNotes", credit_note_attachment_candidates)
            return (
                memo_total, credit_total, memo_line_items_total, credit_line_items_total,
                memo_taxed_line_items_total, credit_taxed_line_items_total,
                memo_doc_buffer_total, credit_doc_buffer_total,
                memo_allocation_count, credit_allocation_count,
                credit_note_tax_lookup,
            )

        # We fetch every BankTransaction once, bucket each one by its account
        # type, and log anything we can't classify. Two separate helper
        # functions are exposed (count_bank_transactions / count_credit_card_transactions)
        # but they share this single fetch so we never hit the API twice.
        async def fetch_all_bank_transactions():
            all_txns = []
            excluded_txns = []
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] BankTransactions - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/BankTransactions", headers,
                    params={"page": page, "where": date_filter_no_deleted, "summaryOnly": "false"},
                )
                data = resp.json().get("BankTransactions", [])
                if not data:
                    print(f"[SUMMARY] BankTransactions - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] BankTransactions - page {page} returned {len(data)} records")
                # The API "where" filter above is a best-effort bandwidth
                # optimization - it isn't always honoured by Xero, so DELETED
                # records must still be dropped here client-side to guarantee
                # they never reach the counts.
                for txn in data:
                    st = (txn.get("Status") or "").upper()
                    if st == "DELETED":
                        excluded_txns.append((txn.get("BankTransactionID"), st))
                        continue
                    all_txns.append(txn)
                page += 1
                await asyncio.sleep(0.05)
            log_excluded("BankTransactions", excluded_txns)
            print(f"[SUMMARY] BankTransactions fetched total = {len(all_txns)}")
            return all_txns

        # Types that represent a transfer between the company's own bank
        # accounts (surfaced as a BankTransaction on each side) rather than
        # a real transaction against a third party - excluded entirely from
        # the Bank Transactions count/LineItems/totals.
        BANK_TXN_TRANSFER_TYPES = {"RECEIVE-TRANSFER", "SPEND-TRANSFER"}

        def count_bank_transactions(all_txns):
            total = 0
            line_items_total = 0
            taxed_line_items_total = 0
            doc_buffer_total = 0  # NEW
            unclassified = 0
            spend_total = 0.0
            receive_total = 0.0
            ignored_transfer_count = 0
            missing_amount_count = 0
            zero_amount_count = 0
            for txn in all_txns:
                txn_type = (txn.get("Type") or "").upper()
                if txn_type in BANK_TXN_TRANSFER_TYPES:
                    ignored_transfer_count += 1
                    continue
                acc_id = txn.get("BankAccount", {}).get("AccountID")
                if acc_id in bank_account_ids:
                    kept_li, li_missing, li_zero = valid_line_items(txn.get("LineItems"))
                    missing_amount_count += li_missing
                    zero_amount_count += li_zero
                    total += 1
                    line_items_total += len(kept_li)
                    taxed_line_items_total += count_taxed_line_items(kept_li)
                    doc_buffer_total += document_line_buffer(kept_li)
                    note_currency(txn)
                    amount = txn.get("Total") or 0
                    if txn_type.startswith("SPEND"):
                        spend_total += amount
                    elif txn_type.startswith("RECEIVE"):
                        receive_total += amount
                elif acc_id not in credit_card_account_ids:
                    # Account isn't in either set -> log it so we can see why
                    unclassified += 1
                    if unclassified <= 10:
                        print(
                            f"[SUMMARY] Unclassified bank txn account: "
                            f"AccountID={acc_id}  Name={txn.get('BankAccount', {}).get('Name')!r}"
                        )
            if unclassified:
                print(f"[SUMMARY] WARNING: {unclassified} bank txns had accounts in neither BANK nor CREDITCARD set")
            if ignored_transfer_count:
                print(f"[SUMMARY] Ignored {ignored_transfer_count} bank txns with Type in {sorted(BANK_TXN_TRANSFER_TYPES)} (own-account transfers)")
            print(f"[SUMMARY] Bank transactions = {total}, LineItems total = {line_items_total} (taxed={taxed_line_items_total}, doc_buffer={doc_buffer_total})")
            print(f"[SUMMARY] Bank transactions LineItems ignored = {missing_amount_count + zero_amount_count}  (no LineAmount field={missing_amount_count}, LineAmount=0={zero_amount_count})")
            print(f"[SUMMARY] Money SPEND (SPEND* types) = {spend_total:,.2f}, Money RECEIVE (RECEIVE* types) = {receive_total:,.2f}")
            return total, line_items_total, taxed_line_items_total, doc_buffer_total

        def count_credit_card_transactions(all_txns):
            total = 0
            line_items_total = 0
            taxed_line_items_total = 0
            doc_buffer_total = 0  # NEW
            missing_amount_count = 0
            zero_amount_count = 0
            for txn in all_txns:
                if (txn.get("Type") or "").upper() in BANK_TXN_TRANSFER_TYPES:
                    continue
                if txn.get("BankAccount", {}).get("AccountID") in credit_card_account_ids:
                    kept_li, li_missing, li_zero = valid_line_items(txn.get("LineItems"))
                    missing_amount_count += li_missing
                    zero_amount_count += li_zero
                    total += 1
                    line_items_total += len(kept_li)
                    taxed_line_items_total += count_taxed_line_items(kept_li)
                    doc_buffer_total += document_line_buffer(kept_li)
                    note_currency(txn)
            print(f"[SUMMARY] Credit card transactions = {total}, LineItems total = {line_items_total} (taxed={taxed_line_items_total}, doc_buffer={doc_buffer_total})")
            print(f"[SUMMARY] Credit card LineItems ignored = {missing_amount_count + zero_amount_count}  (no LineAmount field={missing_amount_count}, LineAmount=0={zero_amount_count})")
            return total, line_items_total, taxed_line_items_total, doc_buffer_total

        def parse_xero_date(d: str):
            """Handles both ISO ('2024-01-15T00:00:00') and the older
            Microsoft-JSON format ('/Date(1705276800000+0000)/')."""
            if not d:
                return None
            if d.startswith("/Date("):
                millis = int(d.split("(")[1].split("+")[0].split("-")[0].rstrip(")/"))
                return datetime.utcfromtimestamp(millis / 1000).date()
            return datetime.strptime(d.split("T")[0], "%Y-%m-%d").date()

        async def count_manual_journals():
            total = 0
            line_items_total = 0
            taxed_line_items_total = 0
            doc_buffer_total = 0  # NEW: +1 per journal that has any TaxAmount > 0 line
            fetched = 0
            parse_failures = 0
            missing_amount_count = 0
            zero_amount_count = 0
            attachment_candidates = []
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] ManualJournals - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/ManualJournals", headers,
                    params={"page": page},
                )
                data = resp.json().get("ManualJournals", [])
                if not data:
                    print(f"[SUMMARY] ManualJournals - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] ManualJournals - page {page} returned {len(data)} records")
                fetched += len(data)
                for mj in data:
                    d = mj.get("Date")
                    try:
                        mj_date = parse_xero_date(d)
                        if mj_date and start <= mj_date <= end:
                            total += 1
                            # ManualJournals use "JournalLines", not "LineItems".
                            kept_li, li_missing, li_zero = valid_line_items(mj.get("JournalLines"))
                            missing_amount_count += li_missing
                            zero_amount_count += li_zero
                            line_items_total += len(kept_li)
                            journal_taxed_count = count_taxed_line_items(kept_li)
                            taxed_line_items_total += journal_taxed_count
                            # +1 buffer per journal based on TaxAmount alone
                            # (not the full +2/+1 document_line_buffer scale,
                            # since JournalLines are already real GL postings -
                            # this just adds the one tax-account line a taxed
                            # journal implies but doesn't carry as its own line).
                            if journal_taxed_count > 0:
                                doc_buffer_total += 1
                            if mj.get("HasAttachments") and mj.get("ManualJournalID"):
                                attachment_candidates.append(mj["ManualJournalID"])
                    except Exception as e:
                        parse_failures += 1
                        if parse_failures <= 5:
                            print(f"[SUMMARY] Could not parse ManualJournal date {d!r}: {e}")
                page += 1
                await asyncio.sleep(0.05)
            print(
                f"[SUMMARY] ManualJournals fetched={fetched}, "
                f"in-range={total}, parse_failures={parse_failures}"
            )
            print(f"[SUMMARY] ManualJournal JournalLines total = {line_items_total} (taxed={taxed_line_items_total}, doc_buffer={doc_buffer_total})")
            print(f"[SUMMARY] ManualJournal JournalLines ignored = {missing_amount_count + zero_amount_count}  (no LineAmount field={missing_amount_count}, LineAmount=0={zero_amount_count})")
            print(f"[SUMMARY] ManualJournals with attachments to check = {len(attachment_candidates)}")
            add_attachment_candidates("ManualJournals", attachment_candidates)
            return total, line_items_total, taxed_line_items_total, doc_buffer_total

        # --- BankTransfers - a genuinely separate endpoint from
        # BankTransactions (it references two linked BankTransaction IDs
        # but isn't one itself). Verified against Xero's OpenAPI spec: this
        # endpoint has NO "page" param at all - Xero returns everything in
        # a single call - and the BankTransfer object has no LineItems
        # field, so this only contributes a record count, not to the
        # LineItems total.
        async def fetch_bank_transfers():
            await refresh_headers()
            print("[SUMMARY] BankTransfers - fetching (single call, not paginated)...")
            resp = await xero_get(
                client, f"{API_BASE}/BankTransfers", headers,
                params={"where": date_filter},
            )
            data = resp.json().get("BankTransfers", [])
            transfer_total = sum(t.get("Amount") or 0 for t in data)
            print(f"[SUMMARY] BankTransfers returned {len(data)} records")
            print(f"[SUMMARY] Money TRANSFER (between own bank accounts) = {transfer_total:,.2f}")
            return len(data)

        # --- Payments - covers invoice payments, bill payments,
        # credit-note allocations, and overpayment/prepayment payments,
        # all via the "PaymentType" field on this one endpoint. Verified
        # against Xero's OpenAPI spec: Payment has NO LineItems field, so
        # this only contributes a record count, not to the LineItems total.
        async def count_payments(invoice_tax_lookup, credit_note_tax_lookup):
            total = 0
            type_breakdown = {}
            # The /Payments list endpoint's nested Invoice.LineItems always
            # comes back empty (verified against a live raw dump), so tax
            # can't be checked from the Payment record itself. Instead,
            # look up the InvoiceID against invoice_tax_lookup (built for
            # free while fetching Invoices/Bills above - no extra call).
            taxed_invoice_payments = 0
            untaxed_invoice_payments = 0
            no_invoice_reference = 0
            excluded_payments = []
            # Receivable/Payable Credit Note Refund tax-line buffer: for
            # ARCREDITPAYMENT/APCREDITPAYMENT records, +1 line ONLY if the
            # credit note being refunded had a taxed line - same "1 record,
            # 1 tax line" rule as Invoice/Bill's own buffer, via
            # credit_note_tax_lookup (built in fetch_credit_notes, since
            # this endpoint's nested Invoice.LineItems is also always empty).
            receivable_credit_note_refund_taxed_count = 0
            payable_credit_note_refund_taxed_count = 0
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Payments - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Payments", headers,
                    params={"page": page, "where": date_filter_no_deleted},
                )
                data = resp.json().get("Payments", [])
                if not data:
                    print(f"[SUMMARY] Payments - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] Payments - page {page} returned {len(data)} records")
                for p in data:
                    # Same client-side safety filter as BankTransactions - the
                    # API "where" filter is a best-effort optimization only.
                    st = (p.get("Status") or "").upper()
                    if st == "DELETED":
                        excluded_payments.append((p.get("PaymentID"), st))
                        continue
                    pt = p.get("PaymentType", "UNKNOWN")
                    type_breakdown[pt] = type_breakdown.get(pt, 0) + 1
                    total += 1

                    invoice_id = p.get("Invoice", {}).get("InvoiceID")
                    if invoice_id in invoice_tax_lookup:
                        if invoice_tax_lookup[invoice_id]:
                            taxed_invoice_payments += 1
                        else:
                            untaxed_invoice_payments += 1
                    else:
                        # Not settling an Invoice/Bill we fetched - e.g. a
                        # payment against a CreditNote/Overpayment/
                        # Prepayment, or the underlying doc is outside the
                        # date range this run fetched.
                        no_invoice_reference += 1

                    if pt == "ARCREDITPAYMENT" and credit_note_tax_lookup.get(invoice_id):
                        receivable_credit_note_refund_taxed_count += 1
                    elif pt == "APCREDITPAYMENT" and credit_note_tax_lookup.get(invoice_id):
                        payable_credit_note_refund_taxed_count += 1
                page += 1
                await asyncio.sleep(0.05)
            print(f"[SUMMARY] Payments count = {total}  (by type: {type_breakdown})")
            log_excluded("Payments", excluded_payments, id_field="PaymentID")
            print(
                f"[SUMMARY] Payments tax check - "
                f"taxed_invoice_payments={taxed_invoice_payments}, "
                f"untaxed_invoice_payments={untaxed_invoice_payments}, "
                f"no_invoice_reference={no_invoice_reference}"
            )
            # Same fix as Invoice TOTAL/Bill TOTAL above, applied to
            # Payments split by type: ACCRECPAYMENT = an Invoice being
            # paid (shows as its own "Receivable Payment" row in Xero's
            # Accounts Receivable report), ACCPAYPAYMENT = a Bill being
            # paid ("Payable Payment" row in Accounts Payable). Each is a
            # two-sided posting - 1 bank-side line + 1 receivable/payable-
            # side line - so each type's count is doubled here. This is
            # visibility only: the combined 2x total was already correct
            # (payment_payable_buffer = payment_count covers both types
            # together); this just breaks that total down by type.
            invoice_payment_count = type_breakdown.get("ACCRECPAYMENT", 0)
            bill_payment_count = type_breakdown.get("ACCPAYPAYMENT", 0)
            print(
                f"[SUMMARY] Invoice Payment (ACCRECPAYMENT) = {invoice_payment_count} (bank-side) "
                f"+ {invoice_payment_count} (receivable-side) = {invoice_payment_count * 2}"
            )
            print(
                f"[SUMMARY] Bill Payment (ACCPAYPAYMENT) = {bill_payment_count} (bank-side) "
                f"+ {bill_payment_count} (payable-side) = {bill_payment_count * 2}"
            )
            # Receivable/Payable Credit Note Refund - a Credit Memo/Bill
            # Credit refunded out through the bank (PaymentType
            # ARCREDITPAYMENT/APCREDITPAYMENT). Counted here since it's
            # only visible via /Payments, not /CreditNotes or /BankTransactions.
            receivable_credit_note_refund_count = type_breakdown.get("ARCREDITPAYMENT", 0)
            payable_credit_note_refund_count = type_breakdown.get("APCREDITPAYMENT", 0)
            print(
                f"[SUMMARY] Receivable Credit Note Refund (ARCREDITPAYMENT) = {receivable_credit_note_refund_count}"
            )
            print(
                f"[SUMMARY] Payable Credit Note Refund (APCREDITPAYMENT) = {payable_credit_note_refund_count}"
            )
            print(
                f"[SUMMARY] Receivable Credit Note Refund taxed (tax-line buffer, same rule as Invoice/Bill) = {receivable_credit_note_refund_taxed_count} of {receivable_credit_note_refund_count}"
            )
            print(
                f"[SUMMARY] Payable Credit Note Refund taxed (tax-line buffer, same rule as Invoice/Bill) = {payable_credit_note_refund_taxed_count} of {payable_credit_note_refund_count}"
            )
            return (
                total, taxed_invoice_payments, untaxed_invoice_payments, no_invoice_reference,
                receivable_credit_note_refund_count, payable_credit_note_refund_count,
                receivable_credit_note_refund_taxed_count, payable_credit_note_refund_taxed_count,
            )

        # --- Overpayments - its own document type (LineItems, SubTotal,
        # etc.), distinct from the SPEND-OVERPAYMENT/RECEIVE-OVERPAYMENT
        # records already counted inside BankTransactions above (those
        # capture the bank movement; this is the source document). Verified
        # against Xero's OpenAPI spec: paginated, HAS LineItems.
        async def fetch_overpayments():
            total = 0
            line_items_total = 0
            taxed_line_items_total = 0
            doc_buffer_total = 0  # NEW
            type_breakdown = {}
            excluded_overpayments = []
            missing_amount_count = 0
            zero_amount_count = 0
            overpayment_attachment_candidates = []
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Overpayments - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Overpayments", headers,
                    params={"page": page, "where": date_filter_no_voided},
                )
                data = resp.json().get("Overpayments", [])
                if not data:
                    print(f"[SUMMARY] Overpayments - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] Overpayments - page {page} returned {len(data)} records")
                for o in data:
                    # Same client-side safety filter as BankTransactions - the
                    # API "where" filter is a best-effort optimization only.
                    st = (o.get("Status") or "").upper()
                    if st == "VOIDED":
                        excluded_overpayments.append((o.get("OverpaymentID"), st))
                        continue
                    ty = o.get("Type", "UNKNOWN")
                    type_breakdown[ty] = type_breakdown.get(ty, 0) + 1
                    total += 1
                    kept_li, li_missing, li_zero = valid_line_items(o.get("LineItems"))
                    missing_amount_count += li_missing
                    zero_amount_count += li_zero
                    line_items_total += len(kept_li)
                    taxed_line_items_total += count_taxed_line_items(kept_li)
                    doc_buffer_total += document_line_buffer(kept_li)
                    if o.get("HasAttachments") and o.get("OverpaymentID"):
                        overpayment_attachment_candidates.append(o["OverpaymentID"])
                page += 1
                await asyncio.sleep(0.05)
            print(f"[SUMMARY] Overpayments count = {total}  (by type: {type_breakdown})")
            print(f"[SUMMARY] Overpayments LineItems total = {line_items_total} (taxed={taxed_line_items_total}, doc_buffer={doc_buffer_total})")
            print(f"[SUMMARY] Overpayments LineItems ignored = {missing_amount_count + zero_amount_count}  (no LineAmount field={missing_amount_count}, LineAmount=0={zero_amount_count})")
            print(f"[SUMMARY] Overpayments with attachments to check = {len(overpayment_attachment_candidates)}")
            add_attachment_candidates("Overpayments", overpayment_attachment_candidates)
            log_excluded("Overpayments", excluded_overpayments, id_field="OverpaymentID")
            return total, line_items_total, taxed_line_items_total, doc_buffer_total

        # --- Prepayments - its own document type, distinct from the
        # SPEND-PREPAYMENT/RECEIVE-PREPAYMENT records already counted
        # inside BankTransactions above. Verified against Xero's OpenAPI
        # spec: paginated, HAS LineItems.
        async def fetch_prepayments():
            total = 0
            line_items_total = 0
            taxed_line_items_total = 0
            doc_buffer_total = 0  # NEW
            type_breakdown = {}
            excluded_prepayments = []
            missing_amount_count = 0
            zero_amount_count = 0
            prepayment_attachment_candidates = []
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Prepayments - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Prepayments", headers,
                    params={"page": page, "where": date_filter_no_voided},
                )
                data = resp.json().get("Prepayments", [])
                if not data:
                    print(f"[SUMMARY] Prepayments - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] Prepayments - page {page} returned {len(data)} records")
                for pp in data:
                    # Same client-side safety filter as BankTransactions - the
                    # API "where" filter is a best-effort optimization only.
                    st = (pp.get("Status") or "").upper()
                    if st == "VOIDED":
                        excluded_prepayments.append((pp.get("PrepaymentID"), st))
                        continue
                    ty = pp.get("Type", "UNKNOWN")
                    type_breakdown[ty] = type_breakdown.get(ty, 0) + 1
                    total += 1
                    kept_li, li_missing, li_zero = valid_line_items(pp.get("LineItems"))
                    missing_amount_count += li_missing
                    zero_amount_count += li_zero
                    line_items_total += len(kept_li)
                    taxed_line_items_total += count_taxed_line_items(kept_li)
                    doc_buffer_total += document_line_buffer(kept_li)
                    if pp.get("HasAttachments") and pp.get("PrepaymentID"):
                        prepayment_attachment_candidates.append(pp["PrepaymentID"])
                page += 1
                await asyncio.sleep(0.05)
            print(f"[SUMMARY] Prepayments count = {total}  (by type: {type_breakdown})")
            print(f"[SUMMARY] Prepayments LineItems total = {line_items_total} (taxed={taxed_line_items_total}, doc_buffer={doc_buffer_total})")
            print(f"[SUMMARY] Prepayments LineItems ignored = {missing_amount_count + zero_amount_count}  (no LineAmount field={missing_amount_count}, LineAmount=0={zero_amount_count})")
            print(f"[SUMMARY] Prepayments with attachments to check = {len(prepayment_attachment_candidates)}")
            add_attachment_candidates("Prepayments", prepayment_attachment_candidates)
            log_excluded("Prepayments", excluded_prepayments, id_field="PrepaymentID")
            return total, line_items_total, taxed_line_items_total, doc_buffer_total

        # --- Purchase Orders (Addons section - kept separate from the core
        # transactions count, never rolled into total_lines) ---
        async def count_purchase_orders():
            total = 0
            status_breakdown = {}
            excluded_pos = []
            po_attachment_candidates = []
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] PurchaseOrders - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/PurchaseOrders", headers,
                    params={"page": page, "where": date_filter_no_deleted},
                )
                data = resp.json().get("PurchaseOrders", [])
                if not data:
                    print(f"[SUMMARY] PurchaseOrders - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] PurchaseOrders - page {page} returned {len(data)} records")
                for po in data:
                    st = po.get("Status", "UNKNOWN")
                    status_breakdown[st] = status_breakdown.get(st, 0) + 1
                    if st != "DELETED":
                        total += 1
                        if po.get("HasAttachments") and po.get("PurchaseOrderID"):
                            po_attachment_candidates.append(po["PurchaseOrderID"])
                    else:
                        excluded_pos.append((po.get("PurchaseOrderNumber") or po.get("PurchaseOrderID"), st))
                page += 1
                await asyncio.sleep(0.05)
            print(f"[SUMMARY] Purchase Order count = {total}  (by status: {status_breakdown})")
            print(f"[SUMMARY] Purchase Orders with attachments to check = {len(po_attachment_candidates)}")
            add_attachment_candidates("PurchaseOrders", po_attachment_candidates)
            log_excluded("Purchase Order", excluded_pos)
            return total

        # --- Quotes, shown as "Sales Order" in Addons (Xero has no native
        # Sales Order object - Quotes is the closest equivalent, converting
        # to an Invoice once accepted). Also kept separate from total_lines.
        async def count_quotes():
            total = 0
            status_breakdown = {}
            excluded_quotes = []
            quote_attachment_candidates = []
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Quotes - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Quotes", headers,
                    params={"page": page, "where": date_filter_no_deleted},
                )
                data = resp.json().get("Quotes", [])
                if not data:
                    print(f"[SUMMARY] Quotes - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] Quotes - page {page} returned {len(data)} records")
                for q in data:
                    st = q.get("Status", "UNKNOWN")
                    status_breakdown[st] = status_breakdown.get(st, 0) + 1
                    if st != "DELETED":
                        total += 1
                        if q.get("HasAttachments") and q.get("QuoteID"):
                            quote_attachment_candidates.append(q["QuoteID"])
                    else:
                        excluded_quotes.append((q.get("QuoteNumber") or q.get("QuoteID"), st))
                page += 1
                await asyncio.sleep(0.05)
            print(f"[SUMMARY] Quote (Sales Order) count = {total}  (by status: {status_breakdown})")
            print(f"[SUMMARY] Quotes with attachments to check = {len(quote_attachment_candidates)}")
            add_attachment_candidates("Quotes", quote_attachment_candidates)
            log_excluded("Quote (Sales Order)", excluded_quotes)
            return total

        # --- Repeating Invoices (Addons). These are templates, not dated
        # transactions - RepeatingInvoices has no "Date" field and its GET
        # endpoint isn't paginated (a single fetch returns every template),
        # so this count is a live "active templates now" snapshot and is
        # NOT affected by the selected conversion period.
        async def count_repeating_invoices():
            await refresh_headers()
            print("[SUMMARY] RepeatingInvoices - fetching...")
            resp = await xero_get(
                client, f"{API_BASE}/RepeatingInvoices", headers,
                params={"where": 'Status!="DELETED" AND Status!="VOIDED"'},
            )
            data = resp.json().get("RepeatingInvoices", [])
            print(f"[SUMMARY] RepeatingInvoices returned {len(data)} records")

            total = 0
            status_breakdown = {}
            excluded_repeating = []
            for ri in data:
                st = ri.get("Status", "UNKNOWN")
                status_breakdown[st] = status_breakdown.get(st, 0) + 1
                if st not in ("VOIDED", "DELETED"):
                    total += 1
                else:
                    excluded_repeating.append((ri.get("RepeatingInvoiceID"), st))
            print(f"[SUMMARY] Recurring Invoices count = {total}  (by status: {status_breakdown})")
            log_excluded("Recurring Invoices", excluded_repeating, id_field="RepeatingInvoiceID")
            return total

        (
            invoice_count, bill_count,
            invoice_line_items_total, bill_line_items_total,
            invoice_taxed_line_items_total, bill_taxed_line_items_total,
            invoice_doc_buffer_total, bill_doc_buffer_total,
            invoice_tax_lookup,
        ) = await fetch_invoices_and_bills()
        (
            credit_memo_count, bill_credit_count,
            credit_memo_line_items_total, bill_credit_line_items_total,
            credit_memo_taxed_line_items_total, bill_credit_taxed_line_items_total,
            credit_memo_doc_buffer_total, bill_credit_doc_buffer_total,
            receivable_credit_note_allocation_count, payable_credit_note_allocation_count,
            credit_note_tax_lookup,
        ) = await fetch_credit_notes()

        # Bank + credit-card share one fetch, then split into two counts.
        all_bank_txns = await fetch_all_bank_transactions()
        bank_txn_count, bank_line_items_total, bank_taxed_line_items_total, bank_doc_buffer_total = count_bank_transactions(all_bank_txns)
        credit_card_txn_count, credit_card_line_items_total, credit_card_taxed_line_items_total, credit_card_doc_buffer_total = count_credit_card_transactions(all_bank_txns)

        bank_txn_attachment_candidates = [
            t["BankTransactionID"] for t in all_bank_txns
            if t.get("HasAttachments") and t.get("BankTransactionID")
        ]
        print(f"[SUMMARY] BankTransactions with attachments to check = {len(bank_txn_attachment_candidates)}")
        add_attachment_candidates("BankTransactions", bank_txn_attachment_candidates)

        manual_journal_count, manual_journal_line_items_total, manual_journal_taxed_line_items_total, manual_journal_doc_buffer_total = await count_manual_journals()

        purchase_order_count = await count_purchase_orders()
        sales_order_count = await count_quotes()
        recurring_invoice_count = await count_repeating_invoices()

        # --- "Bank" is broader than just BankTransactions: BankTransfers,
        # Payments, Overpayments and Prepayments are all bank-related money
        # movement too, so they're folded into ONE combined Bank figure -
        # both the "Bank" row shown in the UI and its contribution to Total
        # Lines. Nothing else (Invoices, Bills, Credit Notes, Manual
        # Journals, Credit Card) is touched - this only expands what "Bank"
        # covers.
        bank_transfer_count = await fetch_bank_transfers()  # no LineItems
        (
            payment_count, taxed_invoice_payments, untaxed_invoice_payments, no_invoice_reference_payments,
            receivable_credit_note_refund_count, payable_credit_note_refund_count,
            receivable_credit_note_refund_taxed_count, payable_credit_note_refund_taxed_count,
        ) = await count_payments(invoice_tax_lookup, credit_note_tax_lookup)
        overpayment_count, overpayment_line_items_total, overpayment_taxed_line_items_total, overpayment_doc_buffer_total = await fetch_overpayments()
        prepayment_count, prepayment_line_items_total, prepayment_taxed_line_items_total, prepayment_doc_buffer_total = await fetch_prepayments()

        # Every document type has now been fetched and had its
        # HasAttachments candidates added to all_attachment_candidates
        # (Invoices/Bills, CreditNotes, BankTransactions, ManualJournals,
        # PurchaseOrders, Quotes, Overpayments, Prepayments) - fetch every
        # candidate's real attachment count in ONE combined high-
        # concurrency batch now, instead of one slow low-concurrency batch
        # per document type.
        attachments_skipped = skip_attachments
        if skip_attachments:
            print(f"[SUMMARY] Skipping attachment count fetch (skip_attachments=true) - {len(all_attachment_candidates)} candidates NOT checked.")
            attachment_count = 0
        else:
            attachment_count = await fetch_all_attachment_counts()
            print(f"[SUMMARY] Total attachment count (all document types combined) = {attachment_count}")

        # combined_bank_records: RECORD count across bank-related endpoints -
        # used for the "Bank" transaction-count display only, never for
        # Total Lines (see combined_bank_line_total below). BankTransfers
        # are deliberately EXCLUDED here - a transfer just moves money
        # between two of the company's own bank accounts, it isn't a
        # bank transaction against a third party, so it shouldn't inflate
        # the "Bank" count (it's still shown separately as its own
        # "Bank Transfers" row in the console breakdown below).
        combined_bank_records = (
            bank_txn_count + payment_count
            + overpayment_count + prepayment_count
        )

        # combined_bank_line_total: the LINE count contribution to Total
        # Lines. BankTransfers and Payments genuinely have no LineItems
        # field (verified against Xero's OpenAPI spec) and so fall back to
        # a per-record proxy; BankTransactions/Overpayments/Prepayments use
        # ONLY their real LineItems sum, exactly once.
        #
        # payment_payable_buffer: a Payment is a two-sided GL posting - it
        # shows up as its own line in the Bank account AND as a separate
        # "Payable Payment"/"Receivable Payment" line in the Accounts
        # Payable/Receivable ledger (confirmed against a real Xero
        # Accounts Payable report - each payment has its own "Payable
        # Payment" row there, distinct from the bank-side entry). So each
        # Payment contributes 2 lines total, not 1 - the payment_count
        # already counted below is the bank-side line; this is the extra
        # payable/receivable-side line.
        payment_payable_buffer = payment_count
        # bank_transfer_buffer: a BankTransfer is also a two-sided posting -
        # it moves money out of one bank account and into another, so each
        # transfer contributes 2 lines total, not 1 (e.g. 1 transfer -> 2).
        # No LineItems to buffer via document_line_buffer, so this is a
        # flat +1 per record instead.
        bank_transfer_buffer = bank_transfer_count
        bank_transfer_line_total = bank_transfer_count + bank_transfer_buffer

        #new line added by me
        payable_credit_note_allocation_buffer = payable_credit_note_allocation_count
        payable_credit_note_allocation_line_total = (
            payable_credit_note_allocation_count + payable_credit_note_allocation_buffer
        )
        receivable_credit_note_allocation_buffer = receivable_credit_note_allocation_count
        receivable_credit_note_allocation_line_total = (
            receivable_credit_note_allocation_count + receivable_credit_note_allocation_buffer
        )

        # Receivable/Payable Credit Note Refund buffer - NOT a flat +1
        # anymore (that double-counted: the refund's own bank-side +
        # receivable/payable-side lines are ALREADY inside payment_count +
        # payment_payable_buffer above, since these are ordinary Payment
        # records). Same "1 record, 1 tax line" rule as Invoice/Bill
        # instead - +1 ONLY if the credit note being refunded was taxed
        # (credit_note_tax_lookup, built in fetch_credit_notes). This is a
        # genuinely separate line (the tax-account posting), so it's safe
        # to add to Total Lines without double-counting.
        receivable_credit_note_refund_buffer = receivable_credit_note_refund_taxed_count
        receivable_credit_note_refund_line_total = (
            receivable_credit_note_refund_count + receivable_credit_note_refund_buffer
        )
        payable_credit_note_refund_buffer = payable_credit_note_refund_taxed_count
        payable_credit_note_refund_line_total = (
            payable_credit_note_refund_count + payable_credit_note_refund_buffer
        )
#---------------------------------------------------------------------------------

        # Receivable/Payable Credit Note buffer: same tax-based rule as
        # Invoice/Bill (see document_line_buffer) - +2 per document if it
        # has any taxed line (1 tax-account line + 1 AR/AP control-account
        # line), +1 if untaxed-only, +0 if no line items. Reuses
        # credit_memo_doc_buffer_total/bill_credit_doc_buffer_total (already
        # computed this way in fetch_credit_notes) instead of a flat +1 per
        # record. NOT added again to Total Lines here - that buffer is
        # already folded in once via total_doc_buffer/the Credit Memo/Bill
        # Credit rows below; these two are display-only rows so the board
        # can show them under the same label as the reference report.
        receivable_credit_note_buffer = credit_memo_doc_buffer_total
        receivable_credit_note_line_total = credit_memo_count + receivable_credit_note_buffer
        payable_credit_note_buffer = bill_credit_doc_buffer_total
        payable_credit_note_line_total = bill_credit_count + payable_credit_note_buffer

        combined_bank_line_items = (
            bank_line_items_total + overpayment_line_items_total + prepayment_line_items_total
        )
        combined_bank_line_total = (
            combined_bank_line_items + bank_transfer_line_total + payment_count + payment_payable_buffer
        )
        combined_bank_taxed_line_items = (
            bank_taxed_line_items_total + overpayment_taxed_line_items_total + prepayment_taxed_line_items_total
        )
        # combined_bank_doc_buffer_total: sum of the NEW per-document
        # buffer (see document_line_buffer) across BankTransactions,
        # Overpayments, and Prepayments - the only bank-related categories
        # that actually have a LineItems array to apply it to.
        # BankTransfers have no LineItems at all, so there's no per-document
        # tax/control-account distinction to buffer for them - their
        # contribution stays exactly the 1-line-per-record proxy already
        # folded into combined_bank_line_total above. Payments also have no
        # LineItems, but DO get their own buffer (payment_payable_buffer
        # above) since they're a confirmed two-sided posting.
        combined_bank_doc_buffer_total = (
            bank_doc_buffer_total + overpayment_doc_buffer_total + prepayment_doc_buffer_total
        )
        print(
            f"[SUMMARY] Bank (combined) - BankTransactions={bank_txn_count}+{bank_line_items_total}, "
            f"BankTransfers={bank_transfer_count} (excluded from Bank count - own-account transfer, not a bank transaction) "
            f"+ {bank_transfer_buffer} buffer (two-sided posting) = {bank_transfer_line_total} (in Total Lines), "
            f"Payments={payment_count} (bank-side) + {payment_payable_buffer} (payable/receivable-side) = {payment_count + payment_payable_buffer}, "
            f"Overpayments={overpayment_count}+{overpayment_line_items_total}, "
            f"Prepayments={prepayment_count}+{prepayment_line_items_total} "
            f"-> display_records={combined_bank_records} (BankTransactions+Payments+Overpayments+Prepayments, no BankTransfers), "
            f"line_total(for Total Lines)={combined_bank_line_total}, "
            f"taxed_line_items={combined_bank_taxed_line_items}, "
            f"doc_buffer={combined_bank_doc_buffer_total}"
        )

        # /Journals (General Ledger) is gated behind Xero's Advanced pricing
        # tier plus a security assessment and use-case approval - most apps
        # get a 401 here. Falls back to a LineItems-sum total (computed
        # above from the endpoints we do have access to) so the summary
        # still returns something useful instead of a 500 - Bank's
        # contribution here is combined_bank_line_total above (real
        # LineItems for BankTransactions/Overpayments/Prepayments, plus a
        # 1-line-per-record proxy for BankTransfers/Payments only, counted
        # exactly once); every other category is unchanged, still
        # LineItems-only. If/when this app gets Journals access,
        # general_ledger_line_count will populate and total_lines below
        # switches to the true ledger-level number automatically.

        #real code
        # line_items_fallback_total = (
        #     invoice_line_items_total + bill_line_items_total
        #     + credit_memo_line_items_total + bill_credit_line_items_total
        #     + combined_bank_line_total + credit_card_line_items_total
        #     + manual_journal_line_items_total
        #     + 0  # sales_receipts, always 0
        # )
        #----------------------------------------------------

        #added by me 

        # NOTE: receivable_credit_note_line_total/payable_credit_note_line_total
        # are NOT added here - their buffer (credit_memo_doc_buffer_total/
        # bill_credit_doc_buffer_total) is already counted once via the
        # Credit Memo/Bill Credit LineItems rows + total_doc_buffer below.
        # Adding them here too would double-count the same buffer.
        #
        # Receivable/Payable Credit Note Refund: only the BUFFER
        # (receivable_credit_note_refund_buffer/payable_credit_note_refund_buffer)
        # is added, NOT the record count - ARCREDITPAYMENT/APCREDITPAYMENT
        # records are ordinary Payments, so their count is already inside
        # payment_count + payment_payable_buffer above via combined_bank_line_total.
        # Adding the count again here would double-count it.
        line_items_fallback_total = (
            invoice_line_items_total + bill_line_items_total
            + credit_memo_line_items_total + bill_credit_line_items_total
            + combined_bank_line_total + credit_card_line_items_total
            + manual_journal_line_items_total + manual_journal_doc_buffer_total
            + payable_credit_note_allocation_line_total
            + receivable_credit_note_allocation_line_total
            + receivable_credit_note_refund_buffer + payable_credit_note_refund_buffer
            + payroll_count
            + 0  # sales_receipts, always 0
        )

        #----------------------------------------

        # Approximation buffer ONLY applied to the fallback (never to real
        # /Journals data). FIX: previously this multiplied the TOTAL COUNT
        # OF TAXED LINE ITEMS by a flat constant (+2 each) - which scales
        # with how many line items a document has, not with how many
        # DOCUMENTS exist. The GL-level postings this buffer approximates
        # (a tax-account line, and a control-account/AR/AP/bank line) each
        # post ONCE PER DOCUMENT, not once per line item - so that old
        # approach overshot massively on any file where documents average
        # many line items each (confirmed: correct on a file averaging
        # ~1.7 lines/invoice, ~2x too high on a file averaging ~11.7
        # lines/invoice).
        #
        # New rule (see document_line_buffer): applied ONCE per document,
        # not once per line item -
        #   - document has >=1 line item, at least one taxed -> +2
        #     (1 tax-account line + 1 control-account line)
        #   - document has >=1 line item, none taxed          -> +1
        #     (just the control-account line)
        #   - document has no line items                       -> +0
        #
        # Applied to every category with a real LineItems array (Invoice,
        # Bill, Credit Memo, Bill Credit, BankTransaction, CreditCard,
        # Overpayment, Prepayment). Manual Journals are DELIBERATELY
        # excluded - a ManualJournal's JournalLines ARE already actual
        # GL-level debit/credit postings (that's what a manual journal is),
        # not document-level lines needing translation to GL - adding this
        # buffer there would double-count real GL lines that already exist.
        # BankTransfers/Payments are also excluded - they have no LineItems
        # at all, so there's no taxed/untaxed distinction to buffer; their
        # contribution is already the flat 1-line-per-record proxy above.
        total_doc_buffer = (
            invoice_doc_buffer_total + bill_doc_buffer_total
            + credit_memo_doc_buffer_total + bill_credit_doc_buffer_total
            + combined_bank_doc_buffer_total + credit_card_doc_buffer_total
        )
        line_items_fallback_total_buffered = (
            line_items_fallback_total + total_doc_buffer
        )

        # TOTAL_LINES_BUFFER: flat overall buffer added on top of the final
        # Total Lines figure, regardless of source (real /Journals count or
        # the LineItems-based fallback) - not tied to any one category.
        TOTAL_LINES_BUFFER = 50

        # Total Lines is the LineItems-sum + per-document tax/control-account
        # buffer (line_items_fallback_total_buffered), NOT the raw /Journals
        # GL rollup. Verified against a client's own Xero-exported report
        # (~50k) - it matched the buffered document-approximation, not the
        # real GL JournalLines count (which ran ~84k on the same data,
        # because Xero's ledger posts more granular double-entry lines per
        # document - tax-rate splits, tracking-category splits, multi-
        # currency realised gain/loss, etc - than this buffer approximates).
        # /Journals is also gated behind Xero's Advanced tier + security
        # assessment (most apps don't have access) and a full paginated
        # fetch of it is one of the slower parts of a run, so it's no longer
        # called at all here.
        general_ledger_line_count = line_items_fallback_total_buffered
        general_ledger_line_count += TOTAL_LINES_BUFFER
        print(
            f"[SUMMARY] + Total Lines buffer (flat, overall) = "
            f"{TOTAL_LINES_BUFFER} -> Total Lines = {general_ledger_line_count}"
        )

        print(
            f"[SUMMARY] Total Lines (LineItems-based fallback) - "
            f"invoices={invoice_line_items_total}, bills={bill_line_items_total}, "
            f"credit_memos={credit_memo_line_items_total}, bill_credits={bill_credit_line_items_total}, "
            f"bank(combined)={combined_bank_line_total}, credit_card={credit_card_line_items_total}, "
            f"manual_journals={manual_journal_line_items_total}, "
            f"total_doc_buffer={total_doc_buffer}, "
            f"raw_total={line_items_fallback_total}, buffered_total={line_items_fallback_total_buffered}"
        )
        print(f"[SUMMARY] Total Lines (value used in response) = {general_ledger_line_count}")

        # Explicit total for Manual Journals - +1 per journal that has any
        # TaxAmount > 0 JournalLine (NOT the full +2/+1 document_line_buffer
        # scale used for Invoice/Bill/etc, since JournalLines are already
        # real GL debit/credit postings, not document-level lines needing
        # translation to GL - only the tax-account line is added here).
        print(
            f"[SUMMARY] Manual Journals TOTAL (taxed journals +1 buffer each) - "
            f"records={manual_journal_count}, JournalLines={manual_journal_line_items_total} "
            f"+ buffer={manual_journal_doc_buffer_total} = {manual_journal_line_items_total + manual_journal_doc_buffer_total}"
        )
        print(
            f"[SUMMARY] Bank Transfers TOTAL (flat +1 buffer per record, two-sided posting) - "
            f"records={bank_transfer_count} + buffer={bank_transfer_buffer} = {bank_transfer_line_total}"
        )
        print(
            f"[SUMMARY] Receivable Credit Note Allocation TOTAL (flat +1 buffer per record) - "
            f"records={receivable_credit_note_allocation_count} + buffer={receivable_credit_note_allocation_buffer} = {receivable_credit_note_allocation_line_total}"
        )
        print(
            f"[SUMMARY] Payable Credit Note Allocation TOTAL (flat +1 buffer per record) - "
            f"records={payable_credit_note_allocation_count} + buffer={payable_credit_note_allocation_buffer} = {payable_credit_note_allocation_line_total}"
        )
        print(
            f"[SUMMARY] Receivable Credit Note Refund TOTAL (tax-line buffer only, same rule as Invoice/Bill - "
            f"record count already inside Payments total, not re-added) - "
            f"records={receivable_credit_note_refund_count} + buffer={receivable_credit_note_refund_buffer} (display total={receivable_credit_note_refund_line_total})"
        )
        print(
            f"[SUMMARY] Payable Credit Note Refund TOTAL (tax-line buffer only, same rule as Invoice/Bill - "
            f"record count already inside Payments total, not re-added) - "
            f"records={payable_credit_note_refund_count} + buffer={payable_credit_note_refund_buffer} (display total={payable_credit_note_refund_line_total})"
        )
        print(
            f"[SUMMARY] Receivable Credit Note TOTAL (same tax-based buffer as Invoice/Bill, display-only - "
            f"already counted once via the Credit Memo row) - "
            f"records={credit_memo_count} + buffer={receivable_credit_note_buffer} = {receivable_credit_note_line_total}"
        )
        print(
            f"[SUMMARY] Payable Credit Note TOTAL (same tax-based buffer as Invoice/Bill, display-only - "
            f"already counted once via the Bill Credit row) - "
            f"records={bill_credit_count} + buffer={payable_credit_note_buffer} = {payable_credit_note_line_total}"
        )

        # Consolidated table - same numbers as the print()s above, just
        # laid out so the whole Total Lines breakdown is readable at a
        # glance instead of scattered across many log lines. "Total" per
        # row = that category's actual contribution to Total Lines (not
        # always LineItems+Buffer - e.g. Bank Transfers/Payments have no
        # LineItems, so their Total comes from the record-based proxy
        # instead), and the rows sum to buffered_total below.
        table_rows = [
            ("Invoice", invoice_count, invoice_line_items_total, invoice_doc_buffer_total,
             invoice_line_items_total + invoice_doc_buffer_total),
            ("Bill", bill_count, bill_line_items_total, bill_doc_buffer_total,
             bill_line_items_total + bill_doc_buffer_total),
            ("Credit Memo", credit_memo_count, credit_memo_line_items_total, credit_memo_doc_buffer_total,
             credit_memo_line_items_total + credit_memo_doc_buffer_total),
            ("Bill Credit", bill_credit_count, bill_credit_line_items_total, bill_credit_doc_buffer_total,
             bill_credit_line_items_total + bill_credit_doc_buffer_total),
            ("Bank Transactions", bank_txn_count, bank_line_items_total, bank_doc_buffer_total,
             bank_line_items_total + bank_doc_buffer_total),
            ("Credit Card", credit_card_txn_count, credit_card_line_items_total, credit_card_doc_buffer_total,
             credit_card_line_items_total + credit_card_doc_buffer_total),
            ("Manual Journals", manual_journal_count, manual_journal_line_items_total, manual_journal_doc_buffer_total,
             manual_journal_line_items_total + manual_journal_doc_buffer_total),
            ("Bank Transfers", bank_transfer_count, 0, bank_transfer_buffer, bank_transfer_line_total),
            ("Payments", payment_count, 0, payment_payable_buffer, payment_count + payment_payable_buffer),
            ("Overpayments", overpayment_count, overpayment_line_items_total, overpayment_doc_buffer_total,
             overpayment_line_items_total + overpayment_doc_buffer_total),
            ("Prepayments", prepayment_count, prepayment_line_items_total, prepayment_doc_buffer_total,
             prepayment_line_items_total + prepayment_doc_buffer_total),
            ("Receivable Credit Note Allocation", receivable_credit_note_allocation_count, 0,
             receivable_credit_note_allocation_buffer, receivable_credit_note_allocation_line_total),
            ("Payable Credit Note Allocation", payable_credit_note_allocation_count, 0,
             payable_credit_note_allocation_buffer, payable_credit_note_allocation_line_total),
            ("Receivable Credit Note", credit_memo_count, 0,
             receivable_credit_note_buffer, receivable_credit_note_line_total),
            ("Payable Credit Note", bill_credit_count, 0,
             payable_credit_note_buffer, payable_credit_note_line_total),
            ("Receivable Credit Note Refund", receivable_credit_note_refund_count, 0,
             receivable_credit_note_refund_buffer, receivable_credit_note_refund_line_total),
            ("Payable Credit Note Refund", payable_credit_note_refund_count, 0,
             payable_credit_note_refund_buffer, payable_credit_note_refund_line_total),
            ("Payroll", employee_count, payroll_count, 0, payroll_count),
        ]

        col = f"{'Category':<20}{'Records':>10}{'LineItems':>12}{'Buffer':>10}{'Total':>10}"
        rule = "-" * len(col)
        print(f"\n[SUMMARY] ===== Total Lines Breakdown =====")
        print(col)
        print(rule)
        for label, records, line_items, buffer_val, total in table_rows:
            print(f"{label:<20}{records:>10}{line_items:>12}{buffer_val:>10}{total:>10}")
        print(rule)
        print(f"{'TOTAL LINES (raw)':<20}{'':>10}{'':>12}{'':>10}{line_items_fallback_total:>10}")
        print(f"{'TOTAL LINES (buffered)':<20}{'':>10}{'':>12}{'':>10}{line_items_fallback_total_buffered:>10}")
        print(f"{'TOTAL LINES (used)':<20}{'':>10}{'':>12}{'':>10}{general_ledger_line_count:>10}")
        print(rule + "\n")

    print(f"[SUMMARY] Multi-currency transactions (Invoices/Bills/CreditNotes/BankTxns with CurrencyCode != {base_currency}) = {multi_currency_txn_count}")
    print("[SUMMARY] Done. Sending response.\n")

    return {
        "company_name": org.get("Name"),
        "base_currency": base_currency,
        "financial_year_end": f"{fy_end_day}/{fy_end_month}" if fy_end_day else "N/A",
        "conversion_period": f"{start.strftime('%d/%m/%Y')} - {end.strftime('%d/%m/%Y')}",
        "enabled_currencies": enabled_currencies,

        "chart_of_accounts": {
            "total": len(accounts),
            "bank_accounts": bank_count,
            "credit_cards": cc_count,
            "multi_currency_accounts": multi_currency_accounts,
            "multi_currency_transactions": multi_currency_txn_count,
        },

        "class_job_location": {
            "enabled": len(tracking_categories) > 0,
            "categories": [c.get("Name") for c in tracking_categories],
            # Xero's TrackingCategory has no "type" field distinguishing
            # Class/Job/Location - it's just a name the org picked. Best-
            # effort keyword match on the category name, so "Yes" here can
            # be wrong if the org used an unrelated name - the raw
            # "categories" list above is the source of truth to double-check.
            "class": any("class" in (c.get("Name") or "").lower() for c in tracking_categories),
            "job": any("job" in (c.get("Name") or "").lower() for c in tracking_categories),
            "location": any(
                kw in (c.get("Name") or "").lower()
                for c in tracking_categories
                for kw in ("location", "region", "branch", "site")
            ),
        },

        "tracked_inventory": {
            "enabled": len(tracked_items) > 0,
            "count": len(tracked_items),
        },

        "payroll": {
            "enabled": payroll_enabled,
            "employee_count": employee_count,
            # "Payroll" count = flattened Payslips across every PayRun in
            # the selected date range (see the fetch above for why).
            "payroll_count": payroll_count,
        },

        "attachments": {
            # Combined across Invoices/Bills, CreditNotes, BankTransactions
            # and ManualJournals - only hits the per-record /Attachments
            # endpoint for records that actually flagged HasAttachments,
            # never for the full dataset.
            "count": attachment_count,
            "skipped": attachments_skipped,
        },

        "transactions_in_period": {
            "invoices": invoice_count,
            "bills": bill_count,
            "credit_memos": credit_memo_count,
            "bill_credits": bill_credit_count,
            # "Bank" = combined records across BankTransactions, Payments,
            # Overpayments and Prepayments - all bank-related money movement
            # against third parties. BankTransfers are deliberately excluded
            # (own-account transfer, not a bank transaction against a third
            # party). See combined_bank_records above.
            "bank_transactions": combined_bank_records,
            "credit_card_transactions": credit_card_txn_count,
            "manual_journals": manual_journal_count,
            "sales_receipts": 0,
            "receivable_credit_note_allocations": receivable_credit_note_allocation_count,
            "payable_credit_note_allocations": payable_credit_note_allocation_count,
            "receivable_credit_notes": credit_memo_count,
            "payable_credit_notes": bill_credit_count,
            "receivable_credit_note_refunds": receivable_credit_note_refund_count,
            "payable_credit_note_refunds": payable_credit_note_refund_count,
            # LineItems-sum based, NOT the raw /Journals GL rollup - see the
            # note above count_general_ledger_lines()'s old call site for why.
            # Bank's contribution is combined_bank_line_total (real LineItems
            # for BankTransactions/Overpayments/Prepayments + a
            # 1-line-per-record proxy for BankTransfers/Payments only, each
            # transaction counted exactly once), plus a per-DOCUMENT
            # tax/control-account approximation buffer (see
            # document_line_buffer above - +2 if a document has at least one
            # taxed line, +1 if it has lines but none taxed; NOT per line
            # item, since those GL postings happen once per document
            # regardless of line-item count). Every per-category count above
            # is unaffected (not buffered).
            "total_lines": general_ledger_line_count,
        },

        # Separate section, shown below Transactions on the dashboard.
        # Deliberately NOT folded into transactions_in_period.total_lines.
        "addons": {
            "purchase_orders": purchase_order_count,
            "sales_orders": sales_order_count,
            "recurring_invoices": recurring_invoice_count,
        },
    }


# --------------------------------------------------------------------
# Job runner - /api/summary now starts a background job instead of
# blocking on the whole (potentially many-minute) fetch, so the request
# returns almost instantly and the client polls for progress/result
# instead of holding a connection open long enough to hit a reverse-
# proxy timeout (502) or a browser request timeout.
# --------------------------------------------------------------------
SUMMARY_JOBS: dict = {}  # job_id -> {"status", "progress", "result", "error", "created_at"}
_SUMMARY_JOB_TTL_SECONDS = 3600  # stale jobs older than this get swept on each new job start


def _cleanup_old_summary_jobs():
    now = time.time()
    stale = [jid for jid, j in SUMMARY_JOBS.items() if now - j.get("created_at", now) > _SUMMARY_JOB_TTL_SECONDS]
    for jid in stale:
        del SUMMARY_JOBS[jid]


async def _run_summary_job(job_id: str, tenant_id: str, from_date, to_date, range_preset, skip_attachments):
    try:
        data = await generate_summary_data(
            tenant_id, from_date, to_date, range_preset, skip_attachments, job_id=job_id,
        )
        SUMMARY_JOBS[job_id]["status"] = "done"
        SUMMARY_JOBS[job_id]["result"] = data
    except Exception as e:
        print(f"[SUMMARY] Job {job_id} failed: {e}")
        SUMMARY_JOBS[job_id]["status"] = "error"
        SUMMARY_JOBS[job_id]["error"] = str(e)


@app.get("/api/summary/start")
async def start_summary_job(
    tenant_id: str,
    from_date: str = Query(None),
    to_date: str = Query(None),
    range_preset: str = Query(None),
    skip_attachments: bool = Query(False),
):
    _cleanup_old_summary_jobs()
    job_id = uuid.uuid4().hex
    SUMMARY_JOBS[job_id] = {
        "status": "running",
        "progress": "Starting…",
        "result": None,
        "error": None,
        "created_at": time.time(),
    }
    asyncio.create_task(_run_summary_job(job_id, tenant_id, from_date, to_date, range_preset, skip_attachments))
    return JSONResponse({"job_id": job_id})


@app.get("/api/summary/status")
async def summary_job_status(job_id: str):
    job = SUMMARY_JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found (it may have expired - start a new one via /api/summary/start)")
    return JSONResponse({
        "status": job["status"],
        "progress": job["progress"],
        "result": job["result"],
        "error": job["error"],
    })


# Kept for backwards compatibility (direct callers, scripts, tests that
# still hit /api/summary synchronously) - the dashboard itself now uses
# /api/summary/start + /api/summary/status instead.
@app.get("/api/summary")
async def api_summary(
    tenant_id: str,
    from_date: str = Query(None),
    to_date: str = Query(None),
    range_preset: str = Query(None),
    skip_attachments: bool = Query(False),
):
    data = await generate_summary_data(tenant_id, from_date, to_date, range_preset, skip_attachments)
    return JSONResponse(data)


# --------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------
def safe_filename(name: str) -> str:
    name = (name or "").strip() or "Xero_Conversion_Summary"
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = re.sub(r'\s+', "_", name)
    return name

@app.post("/api/export-excel")
async def export_excel(payload: dict = Body(...)):
    data = payload.get("data", {})
    file_ref = payload.get("file_ref", "")
    company_name = data.get("company_name", "Xero_Conversion_Summary")
    filename_base = safe_filename(file_ref if file_ref and file_ref != "–" else company_name)
    filename = f"{filename_base}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    bold = Font(bold=True)
    txn = data.get("transactions_in_period", {})
    addons = data.get("addons", {})
    rows = [
        ("Company Name", data.get("company_name")),
        ("Conversion Period", data.get("conversion_period")),
        ("Financial Year End", data.get("financial_year_end")),
        ("Base Currency", data.get("base_currency")),
        ("Enabled Currencies", ", ".join(data.get("enabled_currencies", []))),
        ("", ""),
        ("Chart of Accounts", ""),
        ("Total Accounts", data.get("chart_of_accounts", {}).get("total")),
        ("Bank Accounts", data.get("chart_of_accounts", {}).get("bank_accounts")),
        ("Credit Cards", data.get("chart_of_accounts", {}).get("credit_cards")),
        ("Multi-Currency Accounts", len(data.get("chart_of_accounts", {}).get("multi_currency_accounts", []))),
        ("Multi-Currency Transactions", data.get("chart_of_accounts", {}).get("multi_currency_transactions")),
        ("", ""),
        ("Features", ""),
        ("Class / Job / Location", "Yes" if data.get("class_job_location", {}).get("enabled") else "No"),
        ("  Class", "Yes" if data.get("class_job_location", {}).get("class") else "No"),
        ("  Job", "Yes" if data.get("class_job_location", {}).get("job") else "No"),
        ("  Location", "Yes" if data.get("class_job_location", {}).get("location") else "No"),
        ("Tracked Inventory", "Yes" if data.get("tracked_inventory", {}).get("enabled") else "No"),
        ("Payroll", "Yes" if data.get("payroll", {}).get("enabled") else "No"),
        ("Payroll Count", data.get("payroll", {}).get("payroll_count")),
        ("Employees", data.get("payroll", {}).get("employee_count")),
        ("Attachments", data.get("attachments", {}).get("count", 0)),
        ("", ""),
        ("Number of Transactions in Conversion Period", ""),
        ("Total Lines", txn.get("total_lines")),
        ("Bank", txn.get("bank_transactions")),
        ("Credit Card", txn.get("credit_card_transactions")),
        ("Invoice", txn.get("invoices")),
        ("Credit Memo", txn.get("credit_memos")),
        ("Sales Receipt", txn.get("sales_receipts")),
        ("Bill", txn.get("bills")),
        ("Bill Credit", txn.get("bill_credits")),
        ("Manual Journal", txn.get("manual_journals")),
        ("Receivable Credit Note Allocation", txn.get("receivable_credit_note_allocations")),
        ("Payable Credit Note Allocation", txn.get("payable_credit_note_allocations")),
        ("Receivable Credit Note", txn.get("receivable_credit_notes")),
        ("Payable Credit Note", txn.get("payable_credit_notes")),
        ("Receivable Credit Note Refund", txn.get("receivable_credit_note_refunds")),
        ("Payable Credit Note Refund", txn.get("payable_credit_note_refunds")),
        ("", ""),
        ("Addons", ""),
        ("Purchase Order", addons.get("purchase_orders")),
        ("Sales Order", addons.get("sales_orders")),
        ("Recurring Invoices", addons.get("recurring_invoices")),
    ]

    for label, value in rows:
        ws.append([label, value])

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=1):
        cell = row[0]
        if cell.value in (
            "Chart of Accounts", "Features",
            "Number of Transactions in Conversion Period", "Company Name",
            "Addons",
        ):
            cell.font = bold

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 40

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )









# """
# MMC Convert - Xero Conversion Summary
# ---------------------------------------------------------------
# Flow:
#   GET  /                -> Login page (Xero logo + "Login with Xero")
#   GET  /login             -> redirects to Xero consent screen
#   GET  /callback           -> exchanges code for tokens, saves to DB, redirects to /name-entry
#   GET  /name-entry         -> ask user to confirm/edit the display name for this connection
#   POST /name-entry         -> saves display name, redirects to /dashboard
#   GET  /dashboard          -> single page, no sidebar. Date range selector + "Generate Summary"
#   POST /logout             -> clears all saved tokens from the DB, redirects to /
#   GET  /api/org-info       -> lightweight call, populates header on page load
#   GET  /api/summary        -> called via JS when "Generate Summary" is clicked; returns JSON
#   POST /api/export-excel   -> generates and returns an .xlsx file of the last summary

# Install:
#   pip install fastapi uvicorn httpx python-dotenv jinja2 openpyxl python-multipart --break-system-packages
# Run:
#   uvicorn auth:app --reload --port 8000
# """

# import os
# import io
# import re
# import time
# import asyncio
# import sqlite3
# import secrets
# from datetime import date, datetime
# import httpx
# from fastapi import FastAPI, Request, HTTPException, Query, Body, Form
# from fastapi.responses import RedirectResponse, HTMLResponse, JSONResponse, StreamingResponse
# from fastapi.staticfiles import StaticFiles
# from fastapi.templating import Jinja2Templates
# from dotenv import load_dotenv
# from openpyxl import Workbook
# from openpyxl.styles import Font

# load_dotenv()

# CLIENT_ID = os.getenv("XERO_CLIENT_ID")
# CLIENT_SECRET = os.getenv("XERO_CLIENT_SECRET")
# REDIRECT_URI = os.getenv("XERO_REDIRECT_URI")
# SCOPES = os.getenv("XERO_SCOPES")

# AUTHORIZE_URL = "https://login.xero.com/identity/connect/authorize"
# TOKEN_URL = "https://identity.xero.com/connect/token"
# CONNECTIONS_URL = "https://api.xero.com/connections"
# API_BASE = "https://api.xero.com/api.xro/2.0"

# BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# DB_PATH = os.path.join(BASE_DIR, "xero_tokens.db")

# # Path prefix this app is served under, e.g. "/xero-summary" when running
# # behind nginx alongside other tools on the same domain. Empty string means
# # it's served at the domain root (local dev / its own subdomain).
# BASE_PATH = os.getenv("APP_BASE_PATH", "").rstrip("/")

# # Temporary in-memory holder for token_data while user picks an org on
# # /select-org. Keyed by a random "pending" id, cleared once consumed.
# _pending_selections: dict = {}

# app = FastAPI()
# app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
# templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))
# templates.env.globals["base_path"] = BASE_PATH


# # --------------------------------------------------------------------
# # SQLite persistence layer
# # --------------------------------------------------------------------

# def init_db():
#     conn = sqlite3.connect(DB_PATH)
#     conn.execute(
#         """
#         CREATE TABLE IF NOT EXISTS xero_tokens (
#             tenant_id     TEXT PRIMARY KEY,
#             tenant_name   TEXT,
#             display_name  TEXT,
#             access_token  TEXT NOT NULL,
#             refresh_token TEXT NOT NULL,
#             expires_at    REAL NOT NULL,
#             updated_at    REAL NOT NULL
#         )
#         """
#     )
#     try:
#         conn.execute("ALTER TABLE xero_tokens ADD COLUMN display_name TEXT")
#     except sqlite3.OperationalError:
#         pass
#     conn.commit()
#     conn.close()


# init_db()


# def save_token(tenant_id, tenant_name, access_token, refresh_token, expires_in, display_name=None):
#     conn = sqlite3.connect(DB_PATH)
#     existing = conn.execute(
#         "SELECT display_name FROM xero_tokens WHERE tenant_id = ?", (tenant_id,)
#     ).fetchone()
#     final_display_name = display_name if display_name is not None else (existing[0] if existing else None)

#     conn.execute(
#         """
#         INSERT INTO xero_tokens (tenant_id, tenant_name, display_name, access_token, refresh_token, expires_at, updated_at)
#         VALUES (?, ?, ?, ?, ?, ?, ?)
#         ON CONFLICT(tenant_id) DO UPDATE SET
#             tenant_name=excluded.tenant_name,
#             display_name=excluded.display_name,
#             access_token=excluded.access_token,
#             refresh_token=excluded.refresh_token,
#             expires_at=excluded.expires_at,
#             updated_at=excluded.updated_at
#         """,
#         (tenant_id, tenant_name, final_display_name, access_token, refresh_token, time.time() + expires_in, time.time()),
#     )
#     conn.commit()
#     conn.close()


# def update_display_name(tenant_id: str, display_name: str):
#     conn = sqlite3.connect(DB_PATH)
#     conn.execute(
#         "UPDATE xero_tokens SET display_name = ? WHERE tenant_id = ?",
#         (display_name, tenant_id),
#     )
#     conn.commit()
#     conn.close()


# def get_token_row(tenant_id):
#     conn = sqlite3.connect(DB_PATH)
#     conn.row_factory = sqlite3.Row
#     row = conn.execute("SELECT * FROM xero_tokens WHERE tenant_id = ?", (tenant_id,)).fetchone()
#     conn.close()
#     return dict(row) if row else None


# def list_token_rows():
#     conn = sqlite3.connect(DB_PATH)
#     conn.row_factory = sqlite3.Row
#     rows = conn.execute("SELECT tenant_id, tenant_name, display_name, updated_at FROM xero_tokens").fetchall()
#     conn.close()
#     return [dict(r) for r in rows]


# def clear_all_tokens():
#     """Wipes every saved connection - used on logout so the next login starts fresh."""
#     conn = sqlite3.connect(DB_PATH)
#     conn.execute("DELETE FROM xero_tokens")
#     conn.commit()
#     conn.close()


# def delete_token_row(tenant_id: str):
#     """Removes a single tenant's saved connection from the local DB."""
#     conn = sqlite3.connect(DB_PATH)
#     conn.execute("DELETE FROM xero_tokens WHERE tenant_id = ?", (tenant_id,))
#     conn.commit()
#     conn.close()


# # --------------------------------------------------------------------
# # Rate-limit aware GET wrapper
# # --------------------------------------------------------------------

# async def xero_get(client: httpx.AsyncClient, url: str, headers: dict, params: dict = None, max_retries: int = 5):
#     for _ in range(max_retries):
#         resp = await client.get(url, headers=headers, params=params)
#         if resp.status_code == 429:
#             retry_after = int(resp.headers.get("Retry-After", "5"))
#             await asyncio.sleep(retry_after + 1)
#             continue
#         resp.raise_for_status()
#         return resp
#     resp.raise_for_status()
#     return resp


# # --------------------------------------------------------------------
# # OAuth flow
# # --------------------------------------------------------------------

# @app.get("/", response_class=HTMLResponse)
# async def login_page(request: Request):
#     return templates.TemplateResponse(request, "login.html", {})


# @app.get("/login")
# def login(switch_account: bool = False):
#     state = secrets.token_urlsafe(16)
#     params = {
#         "response_type": "code",
#         "client_id": CLIENT_ID,
#         "redirect_uri": REDIRECT_URI,
#         "scope": SCOPES,
#         "state": state,
#     }
#     # Xero caches the identity session cookie in the browser. Without a
#     # `prompt` param it silently reuses that session and skips straight to
#     # the previously-used account. Passing prompt=login/select_account
#     # forces Xero's account chooser so the user can pick a different email.
#     if switch_account:
#         params["prompt"] = "login"  # use "select_account" instead if Xero shows a multi-account picker for the user's tenant setup

#     url = f"{AUTHORIZE_URL}?{httpx.QueryParams(params)}"
#     return RedirectResponse(url)


# @app.get("/callback")
# async def callback(request: Request):
#     code = request.query_params.get("code")
#     error = request.query_params.get("error")

#     if error:
#         raise HTTPException(400, f"Xero returned error: {error}")
#     if not code:
#         raise HTTPException(400, "No authorization code returned")

#     async with httpx.AsyncClient() as client:
#         resp = await client.post(
#             TOKEN_URL,
#             data={"grant_type": "authorization_code", "code": code, "redirect_uri": REDIRECT_URI},
#             auth=(CLIENT_ID, CLIENT_SECRET),
#         )
#     if resp.status_code != 200:
#         raise HTTPException(
#             400,
#             f"Token exchange failed: {resp.text}. Go to / and login again "
#             f"(don't reload this callback URL — the code is single-use)."
#         )
#     token_data = resp.json()

#     async with httpx.AsyncClient() as client:
#         conn_resp = await client.get(
#             CONNECTIONS_URL, headers={"Authorization": f"Bearer {token_data['access_token']}"}
#         )
#     conn_resp.raise_for_status()
#     tenants = conn_resp.json()
#     print(f"[LOGIN] Xero returned {len(tenants)} organisations")

#     if not tenants:
#         raise HTTPException(400, "No Xero organisation connected")

#     if len(tenants) == 1:
#         tenant_id = tenants[0]["tenantId"]
#         tenant_name = tenants[0]["tenantName"]
#         save_token(
#             tenant_id=tenant_id,
#             tenant_name=tenant_name,
#             access_token=token_data["access_token"],
#             refresh_token=token_data["refresh_token"],
#             expires_in=token_data["expires_in"],
#         )
#         return RedirectResponse(url=f"{BASE_PATH}/name-entry?tenant_id={tenant_id}")

#     pending_id = secrets.token_urlsafe(16)
#     _pending_selections[pending_id] = {"token_data": token_data, "tenants": tenants}
#     return templates.TemplateResponse(
#         request, "select_org.html", {"pending": pending_id, "tenants": tenants}
#     )


# @app.post("/select-org")
# async def select_org(pending: str = Form(...), tenant_id: str = Form(...)):
#     entry = _pending_selections.pop(pending, None)
#     if not entry:
#         raise HTTPException(400, "Selection expired, go to / and login again")
#     token_data = entry["token_data"]
#     tenants = entry["tenants"]
#     chosen = next((t for t in tenants if t["tenantId"] == tenant_id), None)
#     if not chosen:
#         raise HTTPException(400, "Invalid organisation selected")
#     save_token(
#         tenant_id=chosen["tenantId"],
#         tenant_name=chosen["tenantName"],
#         access_token=token_data["access_token"],
#         refresh_token=token_data["refresh_token"],
#         expires_in=token_data["expires_in"],
#     )
#     tid = chosen["tenantId"]
#     return RedirectResponse(url=f"{BASE_PATH}/name-entry?tenant_id={tid}", status_code=303)


# @app.get("/name-entry", response_class=HTMLResponse)
# async def name_entry(request: Request, tenant_id: str = None):
#     if not tenant_id:
#         return RedirectResponse(url=f"{BASE_PATH}/")
#     row = get_token_row(tenant_id)
#     if not row:
#         return RedirectResponse(url=f"{BASE_PATH}/")
#     return templates.TemplateResponse(
#         request,
#         "name_entry.html",
#         {"tenant_id": tenant_id, "tenant_name": row["tenant_name"]},
#     )


# @app.post("/name-entry")
# async def name_entry_submit(tenant_id: str = Form(...), display_name: str = Form(...)):
#     row = get_token_row(tenant_id)
#     if not row:
#         raise HTTPException(400, "Invalid tenant")
#     update_display_name(tenant_id, display_name.strip())
#     return RedirectResponse(url=f"{BASE_PATH}/dashboard?tenant_id={tenant_id}", status_code=303)


# async def get_valid_access_token(tenant_id: str) -> str:
#     row = get_token_row(tenant_id)
#     if not row:
#         raise HTTPException(401, "No token found for this tenant. Login again via /")

#     if time.time() > row["expires_at"] - 60:
#         async with httpx.AsyncClient() as client:
#             resp = await client.post(
#                 TOKEN_URL,
#                 data={"grant_type": "refresh_token", "refresh_token": row["refresh_token"]},
#                 auth=(CLIENT_ID, CLIENT_SECRET),
#             )
#         resp.raise_for_status()
#         new_token = resp.json()
#         save_token(
#             tenant_id=tenant_id,
#             tenant_name=row["tenant_name"],
#             access_token=new_token["access_token"],
#             refresh_token=new_token["refresh_token"],
#             expires_in=new_token["expires_in"],
#         )
#         return new_token["access_token"]

#     return row["access_token"]


# # --------------------------------------------------------------------
# # Dashboard page (just renders the shell; no data fetched on load)
# # --------------------------------------------------------------------

# @app.get("/dashboard", response_class=HTMLResponse)
# async def dashboard(request: Request, tenant_id: str = None):
#     if not tenant_id:
#         return RedirectResponse(url=f"{BASE_PATH}/")
#     row = get_token_row(tenant_id)
#     if not row:
#         return RedirectResponse(url=f"{BASE_PATH}/")
#     if not row.get("display_name"):
#         return RedirectResponse(url=f"{BASE_PATH}/name-entry?tenant_id={tenant_id}")
#     return templates.TemplateResponse(
#         request,
#         "dashboard.html",
#         {"tenant_id": tenant_id, "tenant_name": row["display_name"]},
#     )


# @app.post("/logout")
# async def logout(tenant_id: str = Form(...)):
#     """
#     Fully disconnects this Xero organisation:
#       1. Revokes the connection on Xero's side (DELETE /connections/{id})
#       2. Removes the saved tokens from the local DB
#     This guarantees the next login always shows a fresh Xero consent screen,
#     with no risk of stale/mismatched organisation data.
#     """
#     try:
#         access_token = await get_valid_access_token(tenant_id)

#         async with httpx.AsyncClient() as client:
#             conn_resp = await client.get(
#                 CONNECTIONS_URL,
#                 headers={"Authorization": f"Bearer {access_token}"},
#             )
#         conn_resp.raise_for_status()
#         connections = conn_resp.json()

#         connection = next((c for c in connections if c.get("tenantId") == tenant_id), None)

#         if connection:
#             connection_id = connection["id"]
#             async with httpx.AsyncClient() as client:
#                 del_resp = await client.delete(
#                     f"{CONNECTIONS_URL}/{connection_id}",
#                     headers={"Authorization": f"Bearer {access_token}"},
#                 )
#             print(f"[LOGOUT] Xero disconnect status: {del_resp.status_code}")
#         else:
#             print(f"[LOGOUT] No matching Xero connection found for tenant_id={tenant_id}")

#     except Exception as e:
#         print(f"[LOGOUT] Error while revoking Xero connection: {e}")

#     delete_token_row(tenant_id)
#     return RedirectResponse(url=f"{BASE_PATH}/", status_code=303)


# @app.get("/api/org-info")
# async def api_org_info(tenant_id: str):
#     """Lightweight call - just Organisation info, used to populate header on page load."""
#     access_token = await get_valid_access_token(tenant_id)
#     headers = {
#         "Authorization": f"Bearer {access_token}",
#         "Xero-tenant-id": tenant_id,
#         "Accept": "application/json",
#     }
#     async with httpx.AsyncClient(timeout=30) as client:
#         org_resp = await xero_get(client, f"{API_BASE}/Organisation", headers)
#     org = org_resp.json().get("Organisations", [{}])[0]

#     row = get_token_row(tenant_id)
#     display_name = row.get("display_name") if row else None

#     return JSONResponse({
#         "company_name": display_name or org.get("Name"),
#         "financial_year_end_day": org.get("FinancialYearEndDay"),
#         "financial_year_end_month": org.get("FinancialYearEndMonth"),
#     })


# # --------------------------------------------------------------------
# # Helper: Financial Year start calculation
# # --------------------------------------------------------------------

# def compute_fy_start(fy_end_day: int, fy_end_month: int, today: date) -> date:
#     fy_end_this_year = date(today.year, fy_end_month, fy_end_day)

#     if today <= fy_end_this_year:
#         fy_start_year = today.year - 1
#     else:
#         fy_start_year = today.year

#     fy_end_prev = date(fy_start_year, fy_end_month, fy_end_day)
#     if fy_end_month == 12 and fy_end_day == 31:
#         return date(fy_start_year + 1, 1, 1)
#     from datetime import timedelta
#     return fy_end_prev + timedelta(days=1)


# def subtract_years(d: date, years: int) -> date:
#     try:
#         return d.replace(year=d.year - years)
#     except ValueError:
#         return d.replace(year=d.year - years, day=28)


# # --------------------------------------------------------------------
# # Main summary endpoint - called only when "Generate Summary" is clicked
# # --------------------------------------------------------------------

# @app.get("/api/summary")
# async def api_summary(
#     tenant_id: str,
#     from_date: str = Query(None),
#     to_date: str = Query(None),
#     range_preset: str = Query(None),  # "1ytd" | "2ytd" | "3ytd" | None (custom)
# ):
#     access_token = await get_valid_access_token(tenant_id)
#     headers = {
#         "Authorization": f"Bearer {access_token}",
#         "Xero-tenant-id": tenant_id,
#         "Accept": "application/json",
#     }

#     async def refresh_headers():
#         # Xero access tokens last ~30 min. Large orgs (thousands of
#         # invoices/bank txns across dozens of paginated calls) can take
#         # longer than that to fully fetch, so the token can expire mid-run.
#         # get_valid_access_token() only actually hits Xero when the token
#         # is close to expiring, so calling this every page is cheap.
#         fresh_token = await get_valid_access_token(tenant_id)
#         headers["Authorization"] = f"Bearer {fresh_token}"

#     print(f"\n[SUMMARY] Starting generation for tenant_id={tenant_id}")

#     async with httpx.AsyncClient(timeout=60) as client:
#         # --- Organisation info ---
#         print("[SUMMARY] Fetching Organisation...")
#         org_resp = await xero_get(client, f"{API_BASE}/Organisation", headers)
#         org = org_resp.json().get("Organisations", [{}])[0]

#         _row = get_token_row(tenant_id)
#         _display_name = _row.get("display_name") if _row else None
#         if _display_name:
#             org["Name"] = _display_name

#         base_currency = org.get("BaseCurrency")
#         fy_end_day = org.get("FinancialYearEndDay")
#         fy_end_month = org.get("FinancialYearEndMonth")

#         today = date.today()

#         # --- Resolve date range ---
#         if range_preset in ("1ytd", "2ytd", "3ytd") and fy_end_day and fy_end_month:
#             fy_start = compute_fy_start(fy_end_day, fy_end_month, today)
#             years_back = int(range_preset[0])
#             start = subtract_years(fy_start, years_back)
#             end = today
#         elif from_date and to_date:
#             start = datetime.strptime(from_date, "%Y-%m-%d").date()
#             end = datetime.strptime(to_date, "%Y-%m-%d").date()
#         else:
#             if fy_end_day and fy_end_month:
#                 fy_start = compute_fy_start(fy_end_day, fy_end_month, today)
#                 start = subtract_years(fy_start, 1)
#             else:
#                 start = date(today.year - 1, 1, 1)
#             end = today

#         date_filter = f"Date >= DateTime({start.year},{start.month},{start.day}) AND Date <= DateTime({end.year},{end.month},{end.day})"

#         # --- Currencies ---
#         print("[SUMMARY] Fetching Currencies...")
#         cur_resp = await xero_get(client, f"{API_BASE}/Currencies", headers)
#         enabled_currencies = [c.get("Code") for c in cur_resp.json().get("Currencies", [])]

#         # --- Accounts (Chart of Accounts) ---
#         print("[SUMMARY] Fetching Accounts...")
#         acc_resp = await xero_get(client, f"{API_BASE}/Accounts", headers)
#         all_fetched_accounts = acc_resp.json().get("Accounts", [])
#         accounts = [a for a in all_fetched_accounts if a.get("Status") == "ACTIVE"]
#         print(
#             f"[SUMMARY] Accounts fetched={len(all_fetched_accounts)}, "
#             f"ACTIVE={len(accounts)}, "
#             f"excluded(non-ACTIVE)={len(all_fetched_accounts) - len(accounts)}"
#         )

#         bank_count = sum(1 for a in accounts if a.get("BankAccountType") == "BANK")
#         cc_count = sum(1 for a in accounts if a.get("BankAccountType") == "CREDITCARD")
#         multi_currency_accounts = [
#             {"name": a.get("Name"), "code": a.get("Code"), "currency": a.get("CurrencyCode")}
#             for a in accounts
#             if a.get("CurrencyCode") and a.get("CurrencyCode") != base_currency
#         ]

#         bank_account_ids = {
#             a["AccountID"] for a in accounts if a.get("BankAccountType") == "BANK"
#         }
#         credit_card_account_ids = {
#             a["AccountID"] for a in accounts if a.get("BankAccountType") == "CREDITCARD"
#         }

#         # Debug: show exactly how Xero classifies each bank-like account,
#         # so we can see why the bank vs credit-card split differs from the
#         # accounting firm's own report.
#         print(f"[SUMMARY] Accounts: total={len(accounts)}, BANK={bank_count}, CREDITCARD={cc_count}")
#         for a in accounts:
#             if a.get("BankAccountType") in ("BANK", "CREDITCARD"):
#                 print(
#                     f"    -> {a.get('Name')!r}  type={a.get('BankAccountType')}  "
#                     f"currency={a.get('CurrencyCode')}  AccountID={a.get('AccountID')}"
#                 )

#         # --- Tracking Categories (Class/Job/Location) ---
#         print("[SUMMARY] Fetching Tracking Categories...")
#         track_resp = await xero_get(client, f"{API_BASE}/TrackingCategories", headers)
#         tracking_categories = track_resp.json().get("TrackingCategories", [])

#         # --- Items (Tracked Inventory) ---
#         print("[SUMMARY] Fetching Items...")
#         items_resp = await xero_get(client, f"{API_BASE}/Items", headers)
#         items = items_resp.json().get("Items", [])
#         tracked_items = [i for i in items if i.get("IsTrackedAsInventory")]

#         # --- Payroll check ---
#         print("[SUMMARY] Checking Payroll...")
#         payroll_enabled = False
#         try:
#             pr_resp = await client.get(
#                 "https://api.xero.com/payroll.xro/2.0/Employees",
#                 headers={"Authorization": f"Bearer {access_token}", "Xero-tenant-id": tenant_id},
#             )
#             payroll_enabled = pr_resp.status_code == 200
#         except Exception:
#             payroll_enabled = False

#         # --- Invoices: SINGLE PASS for both Invoice (ACCREC) and Bill (ACCPAY) ---
#         # Previously this endpoint was fetched 3 separate times (once for
#         # attachments, once for ACCREC, once for ACCPAY). Attachments counting
#         # has been removed entirely (Xero has no bulk-attachments endpoint, so
#         # it wasn't reducible - it's dropped instead), and ACCREC/ACCPAY are
#         # now counted together in one pass, cutting this endpoint's page
#         # fetches from 3x down to 1x.
#         async def fetch_invoices_and_bills():
#             invoice_total = 0
#             bill_total = 0
#             invoice_status_breakdown = {}
#             bill_status_breakdown = {}
#             attachment_candidates = []  # (InvoiceID) for records with HasAttachments=true
#             page = 1
#             while True:
#                 await refresh_headers()
#                 print(f"[SUMMARY] Invoices - page {page}...")
#                 resp = await xero_get(
#                     client, f"{API_BASE}/Invoices", headers,
#                     params={"page": page, "where": date_filter},
#                 )
#                 data = resp.json().get("Invoices", [])
#                 if not data:
#                     print(f"[SUMMARY] Invoices - page {page} empty, stopping.")
#                     break
#                 print(f"[SUMMARY] Invoices - page {page} returned {len(data)} records")

#                 for inv in data:
#                     st = inv.get("Status", "UNKNOWN")
#                     inv_type = inv.get("Type")
#                     # Exclude VOIDED and DELETED - these never affect the
#                     # ledger and should never be migrated. Everything else
#                     # (AUTHORISED, PAID, DRAFT, SUBMITTED) is counted so this
#                     # stays as close as possible to what Xero's own "All"
#                     # invoices list shows.
#                     if inv_type == "ACCREC":
#                         invoice_status_breakdown[st] = invoice_status_breakdown.get(st, 0) + 1
#                         if st not in ("VOIDED", "DELETED"):
#                             invoice_total += 1
#                     elif inv_type == "ACCPAY":
#                         bill_status_breakdown[st] = bill_status_breakdown.get(st, 0) + 1
#                         if st not in ("VOIDED", "DELETED"):
#                             bill_total += 1

#                     # Attachment detection piggybacks on this same pass - no
#                     # extra Invoices fetch needed. Xero has no bulk-attachments
#                     # endpoint, so the only unavoidable extra calls are one
#                     # per record that actually has an attachment (below).
#                     if inv.get("HasAttachments") and inv.get("InvoiceID"):
#                         attachment_candidates.append(inv["InvoiceID"])

#                 page += 1
#                 await asyncio.sleep(0.25)

#             print(f"[SUMMARY] Invoice (ACCREC) count = {invoice_total}  (by status: {invoice_status_breakdown})")
#             print(f"[SUMMARY] Bill (ACCPAY) count = {bill_total}  (by status: {bill_status_breakdown})")
#             print(f"[SUMMARY] Invoices/Bills with attachments to check = {len(attachment_candidates)}")

#             # Only hit the per-record attachment endpoint for records that
#             # actually flagged HasAttachments - never for the full dataset.
#             attachment_total = 0
#             if attachment_candidates:
#                 sem = asyncio.Semaphore(4)

#                 async def fetch_one(invoice_id):
#                     async with sem:
#                         try:
#                             att_resp = await xero_get(
#                                 client, f"{API_BASE}/Invoices/{invoice_id}/Attachments", headers,
#                             )
#                             return len(att_resp.json().get("Attachments", []))
#                         except Exception:
#                             return 0

#                 results = await asyncio.gather(*[fetch_one(i) for i in attachment_candidates])
#                 attachment_total = sum(results)

#             print(f"[SUMMARY] Invoice/Bill attachment count = {attachment_total}")
#             return invoice_total, bill_total, attachment_total

#         # --- CreditNotes: SINGLE PASS for both Credit Memo and Bill Credit ---
#         async def fetch_credit_notes():
#             memo_total = 0
#             credit_total = 0
#             memo_status_breakdown = {}
#             credit_status_breakdown = {}
#             page = 1
#             while True:
#                 await refresh_headers()
#                 print(f"[SUMMARY] CreditNotes - page {page}...")
#                 resp = await xero_get(
#                     client, f"{API_BASE}/CreditNotes", headers,
#                     params={"page": page, "where": date_filter},
#                 )
#                 data = resp.json().get("CreditNotes", [])
#                 if not data:
#                     print(f"[SUMMARY] CreditNotes - page {page} empty, stopping.")
#                     break
#                 print(f"[SUMMARY] CreditNotes - page {page} returned {len(data)} records")

#                 for n in data:
#                     st = n.get("Status", "UNKNOWN")
#                     note_type = n.get("Type")
#                     if note_type == "ACCRECCREDIT":
#                         memo_status_breakdown[st] = memo_status_breakdown.get(st, 0) + 1
#                         if st not in ("VOIDED", "DELETED"):
#                             memo_total += 1
#                     elif note_type == "ACCPAYCREDIT":
#                         credit_status_breakdown[st] = credit_status_breakdown.get(st, 0) + 1
#                         if st not in ("VOIDED", "DELETED"):
#                             credit_total += 1

#                 page += 1
#                 await asyncio.sleep(0.25)

#             print(f"[SUMMARY] Credit Memo (ACCRECCREDIT) count = {memo_total}  (by status: {memo_status_breakdown})")
#             print(f"[SUMMARY] Bill Credit (ACCPAYCREDIT) count = {credit_total}  (by status: {credit_status_breakdown})")
#             return memo_total, credit_total

#         # We fetch every BankTransaction once, bucket each one by its account
#         # type, and log anything we can't classify. Two separate helper
#         # functions are exposed (count_bank_transactions / count_credit_card_transactions)
#         # but they share this single fetch so we never hit the API twice.
#         async def fetch_all_bank_transactions():
#             all_txns = []
#             page = 1
#             while True:
#                 await refresh_headers()
#                 print(f"[SUMMARY] BankTransactions - page {page}...")
#                 resp = await xero_get(
#                     client, f"{API_BASE}/BankTransactions", headers,
#                     params={"page": page, "where": date_filter},
#                 )
#                 data = resp.json().get("BankTransactions", [])
#                 if not data:
#                     print(f"[SUMMARY] BankTransactions - page {page} empty, stopping.")
#                     break
#                 print(f"[SUMMARY] BankTransactions - page {page} returned {len(data)} records")
#                 all_txns.extend(data)
#                 page += 1
#                 await asyncio.sleep(0.25)
#             print(f"[SUMMARY] BankTransactions fetched total = {len(all_txns)}")
#             return all_txns

#         def count_bank_transactions(all_txns):
#             total = 0
#             unclassified = 0
#             for txn in all_txns:
#                 acc_id = txn.get("BankAccount", {}).get("AccountID")
#                 if acc_id in bank_account_ids:
#                     total += 1
#                 elif acc_id not in credit_card_account_ids:
#                     # Account isn't in either set -> log it so we can see why
#                     unclassified += 1
#                     if unclassified <= 10:
#                         print(
#                             f"[SUMMARY] Unclassified bank txn account: "
#                             f"AccountID={acc_id}  Name={txn.get('BankAccount', {}).get('Name')!r}"
#                         )
#             if unclassified:
#                 print(f"[SUMMARY] WARNING: {unclassified} bank txns had accounts in neither BANK nor CREDITCARD set")
#             print(f"[SUMMARY] Bank transactions = {total}")
#             return total

#         def count_credit_card_transactions(all_txns):
#             total = sum(
#                 1 for txn in all_txns
#                 if txn.get("BankAccount", {}).get("AccountID") in credit_card_account_ids
#             )
#             print(f"[SUMMARY] Credit card transactions = {total}")
#             return total

#         def parse_xero_date(d: str):
#             """Handles both ISO ('2024-01-15T00:00:00') and the older
#             Microsoft-JSON format ('/Date(1705276800000+0000)/')."""
#             if not d:
#                 return None
#             if d.startswith("/Date("):
#                 millis = int(d.split("(")[1].split("+")[0].split("-")[0].rstrip(")/"))
#                 return datetime.utcfromtimestamp(millis / 1000).date()
#             return datetime.strptime(d.split("T")[0], "%Y-%m-%d").date()

#         async def count_manual_journals():
#             total = 0
#             fetched = 0
#             parse_failures = 0
#             page = 1
#             while True:
#                 await refresh_headers()
#                 print(f"[SUMMARY] ManualJournals - page {page}...")
#                 resp = await xero_get(
#                     client, f"{API_BASE}/ManualJournals", headers,
#                     params={"page": page},
#                 )
#                 data = resp.json().get("ManualJournals", [])
#                 if not data:
#                     print(f"[SUMMARY] ManualJournals - page {page} empty, stopping.")
#                     break
#                 print(f"[SUMMARY] ManualJournals - page {page} returned {len(data)} records")
#                 fetched += len(data)
#                 for mj in data:
#                     d = mj.get("Date")
#                     try:
#                         mj_date = parse_xero_date(d)
#                         if mj_date and start <= mj_date <= end:
#                             total += 1
#                     except Exception as e:
#                         parse_failures += 1
#                         if parse_failures <= 5:
#                             print(f"[SUMMARY] Could not parse ManualJournal date {d!r}: {e}")
#                 page += 1
#                 await asyncio.sleep(0.25)
#             print(
#                 f"[SUMMARY] ManualJournals fetched={fetched}, "
#                 f"in-range={total}, parse_failures={parse_failures}"
#             )
#             return total

#         invoice_count, bill_count, attachment_count = await fetch_invoices_and_bills()
#         credit_memo_count, bill_credit_count = await fetch_credit_notes()

#         # Bank + credit-card share one fetch, then split into two counts.
#         all_bank_txns = await fetch_all_bank_transactions()
#         bank_txn_count = count_bank_transactions(all_bank_txns)
#         credit_card_txn_count = count_credit_card_transactions(all_bank_txns)

#         manual_journal_count = await count_manual_journals()

#     print("[SUMMARY] Done. Sending response.\n")

#     return JSONResponse({
#         "company_name": org.get("Name"),
#         "base_currency": base_currency,
#         "financial_year_end": f"{fy_end_day}/{fy_end_month}" if fy_end_day else "N/A",
#         "conversion_period": f"{start.strftime('%d/%m/%Y')} - {end.strftime('%d/%m/%Y')}",
#         "enabled_currencies": enabled_currencies,

#         "chart_of_accounts": {
#             "total": len(accounts),
#             "bank_accounts": bank_count,
#             "credit_cards": cc_count,
#             "multi_currency_accounts": multi_currency_accounts,
#         },

#         "class_job_location": {
#             "enabled": len(tracking_categories) > 0,
#             "categories": [c.get("Name") for c in tracking_categories],
#         },

#         "tracked_inventory": {
#             "enabled": len(tracked_items) > 0,
#             "count": len(tracked_items),
#         },

#         "payroll": {
#             "enabled": payroll_enabled,
#         },

#         "attachments": {
#             "count": attachment_count,  # Only checked on Invoices/Bills now,
#                                          # not BankTransactions/CreditNotes/
#                                          # ManualJournals - keeps this cheap.
#         },

#         "transactions_in_period": {
#             "invoices": invoice_count,
#             "bills": bill_count,
#             "credit_memos": credit_memo_count,
#             "bill_credits": bill_credit_count,
#             "bank_transactions": bank_txn_count,
#             "credit_card_transactions": credit_card_txn_count,
#             "manual_journals": manual_journal_count,
#             "sales_receipts": 0,
#             "total_lines": (
#                 invoice_count + bill_count + credit_memo_count
#                 + bill_credit_count + bank_txn_count + credit_card_txn_count
#                 + manual_journal_count + 0
#             ),
#         },
#     })


# # --------------------------------------------------------------------
# # Excel export
# # --------------------------------------------------------------------

# def safe_filename(name: str) -> str:
#     name = (name or "").strip() or "Xero_Conversion_Summary"
#     name = re.sub(r'[\\/*?:"<>|]', "", name)
#     name = re.sub(r'\s+', "_", name)
#     return name


# @app.post("/api/export-excel")
# async def export_excel(payload: dict = Body(...)):
#     data = payload.get("data", {})
#     file_ref = payload.get("file_ref", "")
#     company_name = data.get("company_name", "Xero_Conversion_Summary")

#     filename_base = safe_filename(file_ref if file_ref and file_ref != "–" else company_name)
#     filename = f"{filename_base}.xlsx"

#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Summary"

#     bold = Font(bold=True)

#     txn = data.get("transactions_in_period", {})

#     rows = [
#         ("Company Name", data.get("company_name")),
#         ("Conversion Period", data.get("conversion_period")),
#         ("Financial Year End", data.get("financial_year_end")),
#         ("Base Currency", data.get("base_currency")),
#         ("Enabled Currencies", ", ".join(data.get("enabled_currencies", []))),
#         ("", ""),
#         ("Chart of Accounts", ""),
#         ("Total Accounts", data.get("chart_of_accounts", {}).get("total")),
#         ("Bank Accounts", data.get("chart_of_accounts", {}).get("bank_accounts")),
#         ("Credit Cards", data.get("chart_of_accounts", {}).get("credit_cards")),
#         ("Multi-Currency Accounts", len(data.get("chart_of_accounts", {}).get("multi_currency_accounts", []))),
#         ("", ""),
#         ("Features", ""),
#         ("Class / Job / Location", "Yes" if data.get("class_job_location", {}).get("enabled") else "No"),
#         ("Tracked Inventory", "Yes" if data.get("tracked_inventory", {}).get("enabled") else "No"),
#         ("Payroll", "Yes" if data.get("payroll", {}).get("enabled") else "No"),
#         ("Attachments", data.get("attachments", {}).get("count", 0)),
#         ("", ""),
#         ("Number of Transactions in Conversion Period", ""),
#         ("Total Lines", txn.get("total_lines")),
#         ("Bank", txn.get("bank_transactions")),
#         ("Credit Card", txn.get("credit_card_transactions")),
#         ("Invoice", txn.get("invoices")),
#         ("Credit Memo", txn.get("credit_memos")),
#         ("Sales Receipt", txn.get("sales_receipts")),
#         ("Bill", txn.get("bills")),
#         ("Bill Credit", txn.get("bill_credits")),
#         ("Manual Journal", txn.get("manual_journals")),
#     ]

#     for label, value in rows:
#         ws.append([label, value])

#     for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=1):
#         cell = row[0]
#         if cell.value in (
#             "Chart of Accounts", "Features",
#             "Number of Transactions in Conversion Period", "Company Name",
#         ):
#             cell.font = bold

#     ws.column_dimensions["A"].width = 34
#     ws.column_dimensions["B"].width = 40

#     buffer = io.BytesIO()
#     wb.save(buffer)
#     buffer.seek(0)

#     return StreamingResponse(
#         buffer,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": f'attachment; filename="{filename}"'},
#     )