"""
Standalone debug script - checks the Invoices endpoint with the SAME date
filter your tool uses, but NO status filter, so you can see every status
that falls inside the date range.

Reuses auth.py's DB/token functions so you don't need to login again.

Usage:
    python check_invoices.py <tenant_id> <from_date> <to_date>

Example:
    python check_invoices.py a7b9d89d-1468-4393-9f4a-dfb7d8297a55 2020-04-01 2026-07-17
"""

import sys
import asyncio
import httpx
from datetime import datetime

from auth import get_valid_access_token, API_BASE, xero_get


async def main(tenant_id: str, from_date: str, to_date: str):
    start = datetime.strptime(from_date, "%Y-%m-%d").date()
    end = datetime.strptime(to_date, "%Y-%m-%d").date()
    date_filter = f"Date >= DateTime({start.year},{start.month},{start.day}) AND Date <= DateTime({end.year},{end.month},{end.day})"
    print(f"Using date filter: {date_filter}")
    print()

    access_token = await get_valid_access_token(tenant_id)
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Xero-tenant-id": tenant_id,
        "Accept": "application/json",
    }

    all_invoices = []

    async with httpx.AsyncClient(timeout=60) as client:
        print("=" * 70)
        print(f"FETCHING: Invoices between {from_date} and {to_date} - no status filter")
        print("=" * 70)
        page = 1
        while True:
            print(f"Page {page}...")
            resp = await xero_get(
                client, f"{API_BASE}/Invoices", headers,
                params={"page": page, "where": date_filter},
            )
            data = resp.json().get("Invoices", [])
            if not data:
                print(f"Page {page} empty, stopping.")
                break
            print(f"Page {page} returned {len(data)} records")
            all_invoices.extend(data)
            page += 1

    print()
    print(f"TOTAL invoices fetched (no filter at all) = {len(all_invoices)}")
    print()

    # --- Breakdown by Type (ACCREC = sales invoice, ACCPAY = bill) ---
    type_counts = {}
    for inv in all_invoices:
        t = inv.get("Type", "UNKNOWN")
        type_counts[t] = type_counts.get(t, 0) + 1
    print(f"Breakdown by Type: {type_counts}")

    # --- Breakdown by Status ---
    status_counts = {}
    for inv in all_invoices:
        s = inv.get("Status", "UNKNOWN")
        status_counts[s] = status_counts.get(s, 0) + 1
    print(f"Breakdown by Status: {status_counts}")

    # --- Breakdown by Type + Status combined ---
    combo_counts = {}
    for inv in all_invoices:
        key = (inv.get("Type", "UNKNOWN"), inv.get("Status", "UNKNOWN"))
        combo_counts[key] = combo_counts.get(key, 0) + 1
    print()
    print("Breakdown by (Type, Status):")
    for key, count in sorted(combo_counts.items()):
        print(f"   {key} -> {count}")

    # --- List every DRAFT invoice (both ACCREC and ACCPAY) ---
    draft_invoices = [inv for inv in all_invoices if inv.get("Status") == "DRAFT"]
    print(f"DRAFT count = {len(draft_invoices)}")

    for status in ("VOIDED", "SUBMITTED", "DELETED"):
        matching = [inv for inv in all_invoices if inv.get("Status") == status]
        print(f"{status} count = {len(matching)}")

    print()
    print("=" * 70)
    print("FINAL: AUTHORISED + PAID only (what a 'real transactions' count would be)")
    print("=" * 70)
    finalized = [inv for inv in all_invoices if inv.get("Status") in ("AUTHORISED", "PAID")]
    finalized_accrec = sum(1 for inv in finalized if inv.get("Type") == "ACCREC")
    finalized_accpay = sum(1 for inv in finalized if inv.get("Type") == "ACCPAY")
    print(f"AUTHORISED+PAID total = {len(finalized)}")
    print(f"  ACCREC (Invoice) = {finalized_accrec}")
    print(f"  ACCPAY (Bill)    = {finalized_accpay}")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python check_invoices.py <tenant_id> <from_date YYYY-MM-DD> <to_date YYYY-MM-DD>")
        sys.exit(1)
    asyncio.run(main(sys.argv[1], sys.argv[2], sys.argv[3]))