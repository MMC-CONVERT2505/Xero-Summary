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
            retry_after = int(resp.headers.get("Retry-After", "5"))
            await asyncio.sleep(retry_after + 1)
            continue
        resp.raise_for_status()
        return resp
    resp.raise_for_status()
    return resp


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
# Main summary endpoint - called only when "Generate Summary" is clicked
# --------------------------------------------------------------------

@app.get("/api/summary")
async def api_summary(
    tenant_id: str,
    from_date: str = Query(None),
    to_date: str = Query(None),
    range_preset: str = Query(None),  # "1ytd" | "2ytd" | "3ytd" | None (custom)
):
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

        # --- Payroll check ---
        print("[SUMMARY] Checking Payroll...")
        payroll_enabled = False
        try:
            pr_resp = await client.get(
                "https://api.xero.com/payroll.xro/2.0/Employees",
                headers={"Authorization": f"Bearer {access_token}", "Xero-tenant-id": tenant_id},
            )
            payroll_enabled = pr_resp.status_code == 200
        except Exception:
            payroll_enabled = False

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
            invoice_status_breakdown = {}
            bill_status_breakdown = {}
            attachment_candidates = []  # (InvoiceID) for records with HasAttachments=true
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Invoices - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Invoices", headers,
                    params={"page": page, "where": date_filter, "summaryOnly": "false"},
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
                    if inv_type == "ACCREC":
                        invoice_status_breakdown[st] = invoice_status_breakdown.get(st, 0) + 1
                        if st not in ("VOIDED", "DELETED"):
                            invoice_total += 1
                            invoice_line_items_total += len(inv.get("LineItems", []) or [])
                    elif inv_type == "ACCPAY":
                        bill_status_breakdown[st] = bill_status_breakdown.get(st, 0) + 1
                        if st not in ("VOIDED", "DELETED"):
                            bill_total += 1
                            bill_line_items_total += len(inv.get("LineItems", []) or [])

                    # Attachment detection piggybacks on this same pass - no
                    # extra Invoices fetch needed. Xero has no bulk-attachments
                    # endpoint, so the only unavoidable extra calls are one
                    # per record that actually has an attachment (below).
                    if inv.get("HasAttachments") and inv.get("InvoiceID"):
                        attachment_candidates.append(inv["InvoiceID"])

                page += 1
                await asyncio.sleep(0.25)

            print(f"[SUMMARY] Invoice (ACCREC) count = {invoice_total}  (by status: {invoice_status_breakdown})")
            print(f"[SUMMARY] Bill (ACCPAY) count = {bill_total}  (by status: {bill_status_breakdown})")
            print(f"[SUMMARY] Invoice LineItems total = {invoice_line_items_total}")
            print(f"[SUMMARY] Bill LineItems total = {bill_line_items_total}")
            print(f"[SUMMARY] Invoices/Bills with attachments to check = {len(attachment_candidates)}")

            # Only hit the per-record attachment endpoint for records that
            # actually flagged HasAttachments - never for the full dataset.
            attachment_total = 0
            if attachment_candidates:
                sem = asyncio.Semaphore(4)

                async def fetch_one(invoice_id):
                    async with sem:
                        try:
                            att_resp = await xero_get(
                                client, f"{API_BASE}/Invoices/{invoice_id}/Attachments", headers,
                            )
                            return len(att_resp.json().get("Attachments", []))
                        except Exception:
                            return 0

                results = await asyncio.gather(*[fetch_one(i) for i in attachment_candidates])
                attachment_total = sum(results)

            print(f"[SUMMARY] Invoice/Bill attachment count = {attachment_total}")
            return invoice_total, bill_total, attachment_total, invoice_line_items_total, bill_line_items_total

        # --- CreditNotes: SINGLE PASS for both Credit Memo and Bill Credit ---
        async def fetch_credit_notes():
            memo_total = 0
            credit_total = 0
            memo_line_items_total = 0
            credit_line_items_total = 0
            memo_status_breakdown = {}
            credit_status_breakdown = {}
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] CreditNotes - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/CreditNotes", headers,
                    params={"page": page, "where": date_filter, "summaryOnly": "false"},
                )
                data = resp.json().get("CreditNotes", [])
                if not data:
                    print(f"[SUMMARY] CreditNotes - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] CreditNotes - page {page} returned {len(data)} records")

                for n in data:
                    st = n.get("Status", "UNKNOWN")
                    note_type = n.get("Type")
                    if note_type == "ACCRECCREDIT":
                        memo_status_breakdown[st] = memo_status_breakdown.get(st, 0) + 1
                        if st not in ("VOIDED", "DELETED"):
                            memo_total += 1
                            memo_line_items_total += len(n.get("LineItems", []) or [])
                    elif note_type == "ACCPAYCREDIT":
                        credit_status_breakdown[st] = credit_status_breakdown.get(st, 0) + 1
                        if st not in ("VOIDED", "DELETED"):
                            credit_total += 1
                            credit_line_items_total += len(n.get("LineItems", []) or [])

                page += 1
                await asyncio.sleep(0.25)

            print(f"[SUMMARY] Credit Memo (ACCRECCREDIT) count = {memo_total}  (by status: {memo_status_breakdown})")
            print(f"[SUMMARY] Bill Credit (ACCPAYCREDIT) count = {credit_total}  (by status: {credit_status_breakdown})")
            print(f"[SUMMARY] Credit Memo LineItems total = {memo_line_items_total}")
            print(f"[SUMMARY] Bill Credit LineItems total = {credit_line_items_total}")
            return memo_total, credit_total, memo_line_items_total, credit_line_items_total

        # We fetch every BankTransaction once, bucket each one by its account
        # type, and log anything we can't classify. Two separate helper
        # functions are exposed (count_bank_transactions / count_credit_card_transactions)
        # but they share this single fetch so we never hit the API twice.
        async def fetch_all_bank_transactions():
            all_txns = []
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] BankTransactions - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/BankTransactions", headers,
                    params={"page": page, "where": date_filter, "summaryOnly": "false"},
                )
                data = resp.json().get("BankTransactions", [])
                if not data:
                    print(f"[SUMMARY] BankTransactions - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] BankTransactions - page {page} returned {len(data)} records")
                all_txns.extend(data)
                page += 1
                await asyncio.sleep(0.25)
            print(f"[SUMMARY] BankTransactions fetched total = {len(all_txns)}")
            return all_txns

        def count_bank_transactions(all_txns):
            total = 0
            line_items_total = 0
            unclassified = 0
            for txn in all_txns:
                acc_id = txn.get("BankAccount", {}).get("AccountID")
                if acc_id in bank_account_ids:
                    total += 1
                    line_items_total += len(txn.get("LineItems", []) or [])
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
            print(f"[SUMMARY] Bank transactions = {total}, LineItems total = {line_items_total}")
            return total, line_items_total

        def count_credit_card_transactions(all_txns):
            total = 0
            line_items_total = 0
            for txn in all_txns:
                if txn.get("BankAccount", {}).get("AccountID") in credit_card_account_ids:
                    total += 1
                    line_items_total += len(txn.get("LineItems", []) or [])
            print(f"[SUMMARY] Credit card transactions = {total}, LineItems total = {line_items_total}")
            return total, line_items_total

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
            fetched = 0
            parse_failures = 0
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
                            line_items_total += len(mj.get("JournalLines", []) or [])
                    except Exception as e:
                        parse_failures += 1
                        if parse_failures <= 5:
                            print(f"[SUMMARY] Could not parse ManualJournal date {d!r}: {e}")
                page += 1
                await asyncio.sleep(0.25)
            print(
                f"[SUMMARY] ManualJournals fetched={fetched}, "
                f"in-range={total}, parse_failures={parse_failures}"
            )
            print(f"[SUMMARY] ManualJournal JournalLines total = {line_items_total}")
            return total, line_items_total

        # --- General Ledger (Journals) - the actual source of "Total Lines".
        # Every transaction type (Invoice, Bill, BankTransaction,
        # ManualJournal, Payment, etc.) posts its double-entry ledger effect
        # here, so this is a completely separate fetch from everything
        # above - it can't reuse Invoices/BankTransactions/CreditNotes/
        # ManualJournals data, it's Xero's own ledger-level rollup.
        # Pagination is offset=JournalNumber based (NOT page-number based
        # like every other endpoint in this file), and /Journals has no
        # "where" filter support, so date-range filtering is client-side
        # per JournalDate - same pattern as count_manual_journals() above.
        async def count_general_ledger_lines():
            total_journals = 0
            line_items_total = 0
            parse_failures = 0
            offset = 0
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Journals - offset {offset}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Journals", headers,
                    params={"offset": offset},
                )
                data = resp.json().get("Journals", [])
                if not data:
                    print(f"[SUMMARY] Journals - offset {offset} empty, stopping.")
                    break
                print(f"[SUMMARY] Journals - offset {offset} returned {len(data)} records")

                for j in data:
                    d = j.get("JournalDate")
                    try:
                        j_date = parse_xero_date(d)
                        if j_date and start <= j_date <= end:
                            total_journals += 1
                            line_items_total += len(j.get("JournalLines", []) or [])
                    except Exception as e:
                        parse_failures += 1
                        if parse_failures <= 5:
                            print(f"[SUMMARY] Could not parse Journal date {d!r}: {e}")

                last_journal_number = data[-1].get("JournalNumber")
                if last_journal_number is None:
                    print(
                        "[SUMMARY] WARNING: Journals page had no JournalNumber on "
                        "the last record, stopping to avoid an infinite loop."
                    )
                    break
                offset = last_journal_number + 1
                await asyncio.sleep(0.25)

            print(
                f"[SUMMARY] Journals in-range={total_journals}, "
                f"parse_failures={parse_failures}"
            )
            print(f"[SUMMARY] General Ledger JournalLines total = {line_items_total}")
            return line_items_total

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
            print(f"[SUMMARY] BankTransfers returned {len(data)} records")
            return len(data)

        # --- Payments - covers invoice payments, bill payments,
        # credit-note allocations, and overpayment/prepayment payments,
        # all via the "PaymentType" field on this one endpoint. Verified
        # against Xero's OpenAPI spec: Payment has NO LineItems field, so
        # this only contributes a record count, not to the LineItems total.
        async def count_payments():
            total = 0
            type_breakdown = {}
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Payments - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Payments", headers,
                    params={"page": page, "where": date_filter},
                )
                data = resp.json().get("Payments", [])
                if not data:
                    print(f"[SUMMARY] Payments - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] Payments - page {page} returned {len(data)} records")
                for p in data:
                    pt = p.get("PaymentType", "UNKNOWN")
                    type_breakdown[pt] = type_breakdown.get(pt, 0) + 1
                    total += 1
                page += 1
                await asyncio.sleep(0.25)
            print(f"[SUMMARY] Payments count = {total}  (by type: {type_breakdown})")
            return total

        # --- Overpayments - its own document type (LineItems, SubTotal,
        # etc.), distinct from the SPEND-OVERPAYMENT/RECEIVE-OVERPAYMENT
        # records already counted inside BankTransactions above (those
        # capture the bank movement; this is the source document). Verified
        # against Xero's OpenAPI spec: paginated, HAS LineItems.
        async def fetch_overpayments():
            total = 0
            line_items_total = 0
            type_breakdown = {}
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Overpayments - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Overpayments", headers,
                    params={"page": page, "where": date_filter},
                )
                data = resp.json().get("Overpayments", [])
                if not data:
                    print(f"[SUMMARY] Overpayments - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] Overpayments - page {page} returned {len(data)} records")
                for o in data:
                    ty = o.get("Type", "UNKNOWN")
                    type_breakdown[ty] = type_breakdown.get(ty, 0) + 1
                    total += 1
                    line_items_total += len(o.get("LineItems", []) or [])
                page += 1
                await asyncio.sleep(0.25)
            print(f"[SUMMARY] Overpayments count = {total}  (by type: {type_breakdown})")
            print(f"[SUMMARY] Overpayments LineItems total = {line_items_total}")
            return total, line_items_total

        # --- Prepayments - its own document type, distinct from the
        # SPEND-PREPAYMENT/RECEIVE-PREPAYMENT records already counted
        # inside BankTransactions above. Verified against Xero's OpenAPI
        # spec: paginated, HAS LineItems.
        async def fetch_prepayments():
            total = 0
            line_items_total = 0
            type_breakdown = {}
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Prepayments - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Prepayments", headers,
                    params={"page": page, "where": date_filter},
                )
                data = resp.json().get("Prepayments", [])
                if not data:
                    print(f"[SUMMARY] Prepayments - page {page} empty, stopping.")
                    break
                print(f"[SUMMARY] Prepayments - page {page} returned {len(data)} records")
                for pp in data:
                    ty = pp.get("Type", "UNKNOWN")
                    type_breakdown[ty] = type_breakdown.get(ty, 0) + 1
                    total += 1
                    line_items_total += len(pp.get("LineItems", []) or [])
                page += 1
                await asyncio.sleep(0.25)
            print(f"[SUMMARY] Prepayments count = {total}  (by type: {type_breakdown})")
            print(f"[SUMMARY] Prepayments LineItems total = {line_items_total}")
            return total, line_items_total

        # --- Purchase Orders (Addons section - kept separate from the core
        # transactions count, never rolled into total_lines) ---
        async def count_purchase_orders():
            total = 0
            status_breakdown = {}
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] PurchaseOrders - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/PurchaseOrders", headers,
                    params={"page": page, "where": date_filter},
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
                page += 1
                await asyncio.sleep(0.25)
            print(f"[SUMMARY] Purchase Order count = {total}  (by status: {status_breakdown})")
            return total

        # --- Quotes, shown as "Sales Order" in Addons (Xero has no native
        # Sales Order object - Quotes is the closest equivalent, converting
        # to an Invoice once accepted). Also kept separate from total_lines.
        async def count_quotes():
            total = 0
            status_breakdown = {}
            page = 1
            while True:
                await refresh_headers()
                print(f"[SUMMARY] Quotes - page {page}...")
                resp = await xero_get(
                    client, f"{API_BASE}/Quotes", headers,
                    params={"page": page, "where": date_filter},
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
                page += 1
                await asyncio.sleep(0.25)
            print(f"[SUMMARY] Quote (Sales Order) count = {total}  (by status: {status_breakdown})")
            return total

        # --- Repeating Invoices (Addons). These are templates, not dated
        # transactions - RepeatingInvoices has no "Date" field and its GET
        # endpoint isn't paginated (a single fetch returns every template),
        # so this count is a live "active templates now" snapshot and is
        # NOT affected by the selected conversion period.
        async def count_repeating_invoices():
            await refresh_headers()
            print("[SUMMARY] RepeatingInvoices - fetching...")
            resp = await xero_get(client, f"{API_BASE}/RepeatingInvoices", headers)
            data = resp.json().get("RepeatingInvoices", [])
            print(f"[SUMMARY] RepeatingInvoices returned {len(data)} records")

            total = 0
            status_breakdown = {}
            for ri in data:
                st = ri.get("Status", "UNKNOWN")
                status_breakdown[st] = status_breakdown.get(st, 0) + 1
                if st not in ("VOIDED", "DELETED"):
                    total += 1
            print(f"[SUMMARY] Recurring Invoices count = {total}  (by status: {status_breakdown})")
            return total

        (
            invoice_count, bill_count, attachment_count,
            invoice_line_items_total, bill_line_items_total,
        ) = await fetch_invoices_and_bills()
        (
            credit_memo_count, bill_credit_count,
            credit_memo_line_items_total, bill_credit_line_items_total,
        ) = await fetch_credit_notes()

        # Bank + credit-card share one fetch, then split into two counts.
        all_bank_txns = await fetch_all_bank_transactions()
        bank_txn_count, bank_line_items_total = count_bank_transactions(all_bank_txns)
        credit_card_txn_count, credit_card_line_items_total = count_credit_card_transactions(all_bank_txns)

        manual_journal_count, manual_journal_line_items_total = await count_manual_journals()
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
        payment_count = await count_payments()              # no LineItems
        overpayment_count, overpayment_line_items_total = await fetch_overpayments()
        prepayment_count, prepayment_line_items_total = await fetch_prepayments()

        combined_bank_records = (
            bank_txn_count + bank_transfer_count + payment_count
            + overpayment_count + prepayment_count
        )
        combined_bank_line_items = (
            bank_line_items_total + overpayment_line_items_total + prepayment_line_items_total
            # BankTransfers/Payments have no LineItems field (verified
            # against Xero's OpenAPI spec), so they only add records above.
        )
        combined_bank_total = combined_bank_records + combined_bank_line_items
        print(
            f"[SUMMARY] Bank (combined) - BankTransactions={bank_txn_count}+{bank_line_items_total}, "
            f"BankTransfers={bank_transfer_count}, Payments={payment_count}, "
            f"Overpayments={overpayment_count}+{overpayment_line_items_total}, "
            f"Prepayments={prepayment_count}+{prepayment_line_items_total} "
            f"-> records={combined_bank_records}, line_items={combined_bank_line_items}, "
            f"combined_total={combined_bank_total}"
        )

        # /Journals (General Ledger) is gated behind Xero's Advanced pricing
        # tier plus a security assessment and use-case approval - most apps
        # get a 401 here. Falls back to a LineItems-sum total (computed
        # above from the endpoints we do have access to) so the summary
        # still returns something useful instead of a 500 - Bank's
        # contribution here is the combined_bank_total above (records +
        # LineItems across all 5 bank-related endpoints); every other
        # category is unchanged, still LineItems-only. If/when this app
        # gets Journals access, general_ledger_line_count will populate and
        # total_lines below switches to the true ledger-level number
        # automatically.
        line_items_fallback_total = (
            invoice_line_items_total + bill_line_items_total
            + credit_memo_line_items_total + bill_credit_line_items_total
            + combined_bank_total + credit_card_line_items_total
            + manual_journal_line_items_total
            + 0  # sales_receipts, always 0
        )
        try:
            general_ledger_line_count = await count_general_ledger_lines()
        except Exception as e:
            print(
                f"[SUMMARY] Journals fetch failed - this endpoint requires Xero's "
                f"Advanced tier + security assessment + use-case approval, most "
                f"apps don't have access. Falling back to LineItems-sum total. Error: {e}"
            )
            general_ledger_line_count = line_items_fallback_total

        print(
            f"[SUMMARY] Total Lines (LineItems-based fallback) - "
            f"invoices={invoice_line_items_total}, bills={bill_line_items_total}, "
            f"credit_memos={credit_memo_line_items_total}, bill_credits={bill_credit_line_items_total}, "
            f"bank(combined)={combined_bank_total}, credit_card={credit_card_line_items_total}, "
            f"manual_journals={manual_journal_line_items_total}, total={line_items_fallback_total}"
        )
        print(f"[SUMMARY] Total Lines (value used in response) = {general_ledger_line_count}")

    print("[SUMMARY] Done. Sending response.\n")

    return JSONResponse({
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
        },

        "class_job_location": {
            "enabled": len(tracking_categories) > 0,
            "categories": [c.get("Name") for c in tracking_categories],
        },

        "tracked_inventory": {
            "enabled": len(tracked_items) > 0,
            "count": len(tracked_items),
        },

        "payroll": {
            "enabled": payroll_enabled,
        },

        "attachments": {
            "count": attachment_count,  # Only checked on Invoices/Bills now,
                                         # not BankTransactions/CreditNotes/
                                         # ManualJournals - keeps this cheap.
        },

        "transactions_in_period": {
            "invoices": invoice_count,
            "bills": bill_count,
            "credit_memos": credit_memo_count,
            "bill_credits": bill_credit_count,
            # "Bank" = combined records across BankTransactions,
            # BankTransfers, Payments, Overpayments and Prepayments - all
            # bank-related money movement, not just the raw BankTransactions
            # count. See combined_bank_records above.
            "bank_transactions": combined_bank_records,
            "credit_card_transactions": credit_card_txn_count,
            "manual_journals": manual_journal_count,
            "sales_receipts": 0,
            # General Ledger (Journals) based when available - Xero's
            # /Journals endpoint is the true ledger-level rollup (every
            # transaction type's double-entry effect), which is what the
            # reference report's "Total Lines" actually measures. Falls
            # back to a LineItems-sum if this app doesn't have Journals
            # access (see count_general_ledger_lines() try/except above);
            # Bank's contribution to that fallback is combined_bank_total
            # (records + LineItems across all 5 bank-related endpoints).
            # Every other per-category count above is unaffected.
            "total_lines": general_ledger_line_count,
        },

        # Separate section, shown below Transactions on the dashboard.
        # Deliberately NOT folded into transactions_in_period.total_lines.
        "addons": {
            "purchase_orders": purchase_order_count,
            "sales_orders": sales_order_count,
            "recurring_invoices": recurring_invoice_count,
        },
    })


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
        ("", ""),
        ("Features", ""),
        ("Class / Job / Location", "Yes" if data.get("class_job_location", {}).get("enabled") else "No"),
        ("Tracked Inventory", "Yes" if data.get("tracked_inventory", {}).get("enabled") else "No"),
        ("Payroll", "Yes" if data.get("payroll", {}).get("enabled") else "No"),
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