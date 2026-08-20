"""
fetch_payroll_only.py
---------------------------------------------------------------
Companion script for MMC Convert (auth.py).

Dumps raw JSON for the Payroll AU/NZ/UK endpoints this app has scope
access to, into ONE timestamped folder, plus a counts summary.

Endpoints covered (payroll.xro/2.0, separate from the Accounting API
used by auth.py - camelCase field names, own pagination shape):
  - /Employees   (paginated)
  - /PayRuns     (paginated - summary level only: totals per pay run,
                  NOT itemised payslips. Per-PayRun/per-Payslip detail
                  endpoints returned 405 "Method not allowed for the
                  current customer jurisdiction" when tested against
                  this AU org - Xero restricts that drill-down for some
                  jurisdictions/API versions, so it's not fetched here.)
  - /Timesheets  (paginated)

Requires the connected token to have been granted the payroll.employees,
payroll.payruns and payroll.timesheets scopes (see XERO_SCOPES in the
.env two levels up) - if the org was connected before those scopes were
added, reconnect via auth.py's /login flow first (a new consent screen
will ask for Payroll permission).

Does NOT modify auth.py - reuses its existing get_valid_access_token
and get_token_row functions as-is (this hits the Payroll API directly,
not xero_get/API_BASE, since Payroll has its own base URL and response
shape).

Usage:
  python fetch_payroll_only.py --tenant-id <TENANT_ID>

  Optional:
    --from-date       YYYY-MM-DD (filters PayRuns/Timesheets by date, client-side)
    --to-date         YYYY-MM-DD
    --output-dir      Parent folder to create the new run-folder inside (default: current directory)

Example:
  python fetch_payroll_only.py --tenant-id 829023a2-b1f1-461c-9849-bdfdbda648be --from-date 2022-07-01 --to-date 2024-06-30
"""

import os
import re
import json
import asyncio
import argparse
from datetime import datetime

import httpx

from auth import get_valid_access_token, get_token_row

PAYROLL_BASE = "https://api.xero.com/payroll.xro/2.0"


def safe_folder_name(name: str) -> str:
    name = (name or "").strip() or "Xero_Payroll"
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = re.sub(r"\s+", "_", name)
    return name


def parse_xero_date(d):
    if not d:
        return None
    return datetime.strptime(d.split("T")[0], "%Y-%m-%d").date()


async def fresh_headers(tenant_id: str, headers: dict):
    fresh_token = await get_valid_access_token(tenant_id)
    headers["Authorization"] = f"Bearer {fresh_token}"


async def payroll_get(client, tenant_id, headers, endpoint, params=None, max_retries=5):
    for attempt in range(max_retries):
        await fresh_headers(tenant_id, headers)
        resp = await client.get(f"{PAYROLL_BASE}/{endpoint}", headers=headers, params=params)
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", "5"))
            await asyncio.sleep(retry_after + 1)
            continue
        if resp.status_code >= 500:
            # Transient Xero-side errors - back off and retry rather than
            # failing the whole run over one flaky page.
            print(f"[PAYROLL] {endpoint} returned {resp.status_code}, retrying (attempt {attempt + 1}/{max_retries})...")
            await asyncio.sleep(3 * (attempt + 1))
            continue
        resp.raise_for_status()
        return resp
    resp.raise_for_status()
    return resp


# --------------------------------------------------------------------
# /Employees - paginated
# --------------------------------------------------------------------
async def fetch_employees(client, tenant_id, headers):
    all_employees = []
    page = 1
    page_count = 1
    while page <= page_count:
        print(f"[PAYROLL] Employees - page {page}...")
        try:
            resp = await payroll_get(client, tenant_id, headers, "Employees", params={"page": page})
        except httpx.HTTPStatusError as e:
            print(f"[PAYROLL] WARNING: Employees - page {page} failed after retries ({e}). Stopping here - results may be INCOMPLETE.")
            break
        body = resp.json()
        # Payroll AU returns 400 Bad Request (not an empty list) if you
        # request a page beyond pagination.pageCount, so pageCount from
        # the response drives the loop instead of "stop on empty".
        page_count = (body.get("pagination") or {}).get("pageCount", page_count)
        data = body.get("employees", [])
        print(f"[PAYROLL] Employees - page {page} returned {len(data)} records (pageCount={page_count})")
        all_employees.extend(data)
        page += 1
        await asyncio.sleep(0.25)
    print(f"[PAYROLL] Total Employees fetched = {len(all_employees)}")
    return all_employees


# --------------------------------------------------------------------
# /Deductions and /EarningsRates - payroll settings, small single-page
# lists (deduction/earnings-rate types configured for this org, not
# per-employee or per-pay-period records).
# --------------------------------------------------------------------
async def fetch_deductions(client, tenant_id, headers):
    print("[PAYROLL] Deductions - fetching...")
    resp = await payroll_get(client, tenant_id, headers, "Deductions", params={"page": 1})
    data = resp.json().get("deductions", [])
    print(f"[PAYROLL] Deductions returned {len(data)} records")
    return data


async def fetch_earnings_rates(client, tenant_id, headers):
    print("[PAYROLL] EarningsRates - fetching...")
    resp = await payroll_get(client, tenant_id, headers, "EarningsRates", params={"page": 1})
    data = resp.json().get("earningsRates", [])
    print(f"[PAYROLL] EarningsRates returned {len(data)} records")
    return data


# --------------------------------------------------------------------
# /PayRuns - paginated, summary level (totalCost/totalPay per pay run,
# no itemised payslip lines - see module docstring).
# --------------------------------------------------------------------
async def fetch_payruns(client, tenant_id, headers, from_date, to_date):
    all_payruns = []
    page = 1
    page_count = 1
    incomplete = False
    while page <= page_count:
        print(f"[PAYROLL] PayRuns - page {page}...")
        try:
            resp = await payroll_get(client, tenant_id, headers, "PayRuns", params={"page": page})
        except httpx.HTTPStatusError as e:
            # Xero's own API has returned a persistent 500 on this page in
            # testing (generic "An error occurred in Xero" - not something
            # fixable client-side, retries don't help). Don't crash the
            # whole run over one bad page - keep what was fetched so far
            # and flag it clearly instead.
            print(f"[PAYROLL] WARNING: PayRuns - page {page} failed after retries ({e}). Stopping here - results are INCOMPLETE.")
            incomplete = True
            break
        body = resp.json()
        page_count = (body.get("pagination") or {}).get("pageCount", page_count)
        data = body.get("payRuns", [])
        print(f"[PAYROLL] PayRuns - page {page} returned {len(data)} records (pageCount={page_count})")
        all_payruns.extend(data)
        page += 1
        await asyncio.sleep(0.25)
    print(f"[PAYROLL] Total PayRuns fetched = {len(all_payruns)}" + (" (INCOMPLETE - see warning above)" if incomplete else ""))

    if from_date and to_date:
        start = datetime.strptime(from_date, "%Y-%m-%d").date()
        end = datetime.strptime(to_date, "%Y-%m-%d").date()
        filtered = [
            pr for pr in all_payruns
            if (d := parse_xero_date(pr.get("paymentDate"))) and start <= d <= end
        ]
        print(f"[PAYROLL] PayRuns in range {start}..{end} = {len(filtered)}")
        return filtered
    return all_payruns


# --------------------------------------------------------------------
# /Timesheets - paginated
# --------------------------------------------------------------------
async def fetch_timesheets(client, tenant_id, headers, from_date, to_date):
    all_timesheets = []
    page = 1
    page_count = 1
    while page <= page_count:
        print(f"[PAYROLL] Timesheets - page {page}...")
        try:
            resp = await payroll_get(client, tenant_id, headers, "Timesheets", params={"page": page})
        except httpx.HTTPStatusError as e:
            print(f"[PAYROLL] WARNING: Timesheets - page {page} failed after retries ({e}). Stopping here - results may be INCOMPLETE.")
            break
        body = resp.json()
        page_count = (body.get("pagination") or {}).get("pageCount", page_count)
        data = body.get("timesheets", [])
        print(f"[PAYROLL] Timesheets - page {page} returned {len(data)} records (pageCount={page_count})")
        all_timesheets.extend(data)
        page += 1
        await asyncio.sleep(0.25)
    print(f"[PAYROLL] Total Timesheets fetched = {len(all_timesheets)}")

    if from_date and to_date:
        start = datetime.strptime(from_date, "%Y-%m-%d").date()
        end = datetime.strptime(to_date, "%Y-%m-%d").date()
        filtered = [
            ts for ts in all_timesheets
            if (d := parse_xero_date(ts.get("startDate"))) and start <= d <= end
        ]
        print(f"[PAYROLL] Timesheets in range {start}..{end} = {len(filtered)}")
        return filtered
    return all_timesheets


def save_json(run_dir, filename, data):
    path = os.path.join(run_dir, filename)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    return path


async def run_all(tenant_id: str, from_date: str, to_date: str):
    access_token = await get_valid_access_token(tenant_id)
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Xero-tenant-id": tenant_id,
        "Accept": "application/json",
    }
    async with httpx.AsyncClient(timeout=60) as client:
        employees = await fetch_employees(client, tenant_id, headers)
        deductions = await fetch_deductions(client, tenant_id, headers)
        earnings_rates = await fetch_earnings_rates(client, tenant_id, headers)
        payruns = await fetch_payruns(client, tenant_id, headers, from_date, to_date)
        timesheets = await fetch_timesheets(client, tenant_id, headers, from_date, to_date)
    return employees, deductions, earnings_rates, payruns, timesheets


def main():
    parser = argparse.ArgumentParser(description="Fetch Payroll (Employees/PayRuns/Timesheets) from Xero into one folder.")
    parser.add_argument("--tenant-id", required=True, help="Xero tenant_id (must already be connected with payroll scopes via auth.py's login flow)")
    parser.add_argument("--from-date", default=None, help="Custom range start, YYYY-MM-DD (filters PayRuns by paymentDate, Timesheets by startDate)")
    parser.add_argument("--to-date", default=None, help="Custom range end, YYYY-MM-DD")
    parser.add_argument("--output-dir", default=".", help="Parent folder to create the new run-folder inside")
    args = parser.parse_args()

    row = get_token_row(args.tenant_id)
    if not row:
        raise SystemExit(
            f"No saved token found for tenant_id={args.tenant_id}. "
            f"Login via auth.py's / page first (this script reuses that saved token)."
        )

    employees, deductions, earnings_rates, payruns, timesheets = asyncio.run(
        run_all(args.tenant_id, args.from_date, args.to_date)
    )

    company_name = row.get("display_name") or row.get("tenant_name") or "Xero_Payroll"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"{safe_folder_name(company_name)}_Payroll_{timestamp}"
    run_dir = os.path.join(args.output_dir, folder_name)
    os.makedirs(run_dir, exist_ok=True)

    saved_paths = {
        "employees": save_json(run_dir, "employees_raw.json", employees),
        "deductions": save_json(run_dir, "deductions_raw.json", deductions),
        "earnings_rates": save_json(run_dir, "earnings_rates_raw.json", earnings_rates),
        "payruns": save_json(run_dir, "payruns_raw.json", payruns),
        "timesheets": save_json(run_dir, "timesheets_raw.json", timesheets),
    }

    payrun_status_breakdown = {}
    total_cost = 0.0
    total_pay = 0.0
    for pr in payruns:
        st = pr.get("payRunStatus", "UNKNOWN")
        payrun_status_breakdown[st] = payrun_status_breakdown.get(st, 0) + 1
        total_cost += pr.get("totalCost") or 0
        total_pay += pr.get("totalPay") or 0

    counts = {
        "employees": {"record_count": len(employees)},
        "deductions": {"record_count": len(deductions)},
        "earnings_rates": {"record_count": len(earnings_rates)},
        "payruns": {
            "record_count": len(payruns),
            "status_breakdown": payrun_status_breakdown,
            "total_cost": round(total_cost, 2),
            "total_pay": round(total_pay, 2),
        },
        "timesheets": {"record_count": len(timesheets)},
    }
    counts_path = save_json(run_dir, "counts_summary.json", counts)

    print()
    print("=" * 70)
    print(f"Employees      : records={len(employees)}")
    print(f"Deductions     : records={len(deductions)}")
    print(f"EarningsRates  : records={len(earnings_rates)}")
    print(f"PayRuns        : records={len(payruns)}  statuses={payrun_status_breakdown}  total_cost={total_cost:,.2f}  total_pay={total_pay:,.2f}")
    print(f"Timesheets     : records={len(timesheets)}")
    print("=" * 70)
    print()
    print(f"All files saved in: {run_dir}")
    for label, path in saved_paths.items():
        print(f"  {label:20s} -> {path}")
    print(f"  {'counts_summary':20s} -> {counts_path}")


if __name__ == "__main__":
    main()
