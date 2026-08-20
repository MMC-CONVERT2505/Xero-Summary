"""
fetch_invoices_only.py
---------------------------------------------------------------
Companion script for MMC Convert (auth.py).

Fetches every Invoice/Bill (/Invoices, covers both ACCREC and ACCPAY
via the "Type" field) for a tenant and dumps the raw JSON into a
timestamped folder, plus a small counts summary.

Deleted/Voided exclusion is applied TWICE:
  1. In the API "where" query (Status!="DELETED" AND Status!="VOIDED")
     - a bandwidth optimization, but Xero doesn't always honour it.
  2. Client-side, on every record actually returned - this is the
     real guarantee that no DELETED/VOIDED record ends up in the
     output file, regardless of whether step 1 worked.

Does NOT modify auth.py - reuses its existing get_valid_access_token,
xero_get, API_BASE, get_token_row functions as-is.

Usage:
  python fetch_invoices_only.py --tenant-id <TENANT_ID>

  Optional:
    --from-date       YYYY-MM-DD (custom range start)
    --to-date         YYYY-MM-DD (custom range end)
    --output-dir      Parent folder to create the new run-folder inside (default: current directory)

Example:
  python fetch_invoices_only.py --tenant-id 1314cf60-8470-46a1-a23b-b5d28ac77fa7 --from-date 2025-04-01 --to-date 2026-08-12
"""

import os
import re
import json
import asyncio
import argparse
from datetime import datetime

import httpx

# --- Reuse auth.py's existing logic untouched (must run from same folder) ---
from auth import get_valid_access_token, xero_get, get_token_row, API_BASE

EXCLUDED_STATUSES = {"DELETED", "VOIDED"}


def safe_folder_name(name: str) -> str:
    name = (name or "").strip() or "Xero_Invoices"
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = re.sub(r"\s+", "_", name)
    return name


def build_date_filter(from_date: str, to_date: str) -> str:
    base = None
    if from_date and to_date:
        start = datetime.strptime(from_date, "%Y-%m-%d").date()
        end = datetime.strptime(to_date, "%Y-%m-%d").date()
        base = f"Date >= DateTime({start.year},{start.month},{start.day}) AND Date <= DateTime({end.year},{end.month},{end.day})"
    status_filter = 'Status!="DELETED" AND Status!="VOIDED"'
    return f"{base} AND {status_filter}" if base else status_filter


async def fresh_headers(tenant_id: str, headers: dict):
    fresh_token = await get_valid_access_token(tenant_id)
    headers["Authorization"] = f"Bearer {fresh_token}"


async def fetch_invoices(client, tenant_id, headers, where_filter):
    all_invoices = []
    skipped_status = 0
    excluded = []  # (InvoiceNumber/InvoiceID, Status) for every skipped record
    page = 1
    while True:
        await fresh_headers(tenant_id, headers)
        params = {"page": page, "summaryOnly": "false"}
        if where_filter:
            params["where"] = where_filter
        print(f"[INVOICES] Invoices - page {page}...")
        resp = await xero_get(client, f"{API_BASE}/Invoices", headers, params=params)
        data = resp.json().get("Invoices", [])
        if not data:
            print(f"[INVOICES] Invoices - page {page} empty, stopping.")
            break
        print(f"[INVOICES] Invoices - page {page} returned {len(data)} records")
        for inv in data:
            st = (inv.get("Status") or "").upper()
            if st in EXCLUDED_STATUSES:
                skipped_status += 1
                excluded.append((inv.get("InvoiceNumber") or inv.get("InvoiceID"), st))
                continue
            all_invoices.append(inv)
        page += 1
        await asyncio.sleep(0.25)
    if skipped_status:
        deleted = [i for i, st in excluded if st == "DELETED"]
        voided = [i for i, st in excluded if st == "VOIDED"]
        print(f"[INVOICES] Skipped {skipped_status} DELETED/VOIDED records (client-side safety filter)  DELETED={len(deleted)}, VOIDED={len(voided)}")
        if deleted:
            print(f"[INVOICES]   DELETED InvoiceNumbers/IDs: {deleted}")
        if voided:
            print(f"[INVOICES]   VOIDED InvoiceNumbers/IDs: {voided}")
    print(f"[INVOICES] Total Invoices kept = {len(all_invoices)}")
    return all_invoices


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
    where_filter = build_date_filter(from_date, to_date)

    async with httpx.AsyncClient(timeout=60) as client:
        invoices = await fetch_invoices(client, tenant_id, headers, where_filter)

    return invoices


def main():
    parser = argparse.ArgumentParser(description="Fetch all Invoices/Bills from Xero into one JSON file (DELETED/VOIDED excluded).")
    parser.add_argument("--tenant-id", required=True, help="Xero tenant_id (must already be connected via auth.py's login flow)")
    parser.add_argument("--from-date", default=None, help="Custom range start, YYYY-MM-DD")
    parser.add_argument("--to-date", default=None, help="Custom range end, YYYY-MM-DD")
    parser.add_argument("--output-dir", default=".", help="Parent folder to create the new run-folder inside")
    args = parser.parse_args()

    row = get_token_row(args.tenant_id)
    if not row:
        raise SystemExit(
            f"No saved token found for tenant_id={args.tenant_id}. "
            f"Login via auth.py's / page first (this script reuses that saved token)."
        )

    invoices = asyncio.run(run_all(args.tenant_id, args.from_date, args.to_date))

    company_name = row.get("display_name") or row.get("tenant_name") or "Xero_Invoices"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"{safe_folder_name(company_name)}_Invoices_{timestamp}"
    run_dir = os.path.join(args.output_dir, folder_name)
    os.makedirs(run_dir, exist_ok=True)

    invoices_path = save_json(run_dir, "invoices_raw.json", invoices)

    status_breakdown = {}
    type_breakdown = {}
    for inv in invoices:
        st = inv.get("Status", "UNKNOWN")
        ty = inv.get("Type", "UNKNOWN")
        status_breakdown[st] = status_breakdown.get(st, 0) + 1
        type_breakdown[ty] = type_breakdown.get(ty, 0) + 1

    counts = {
        "invoices": {
            "record_count": len(invoices),
            "status_breakdown": status_breakdown,
            "type_breakdown": type_breakdown,
        },
    }
    counts_path = save_json(run_dir, "counts_summary.json", counts)

    print()
    print("=" * 70)
    print(f"Invoices : records={len(invoices)}  types={type_breakdown}  statuses={status_breakdown}")
    print("=" * 70)
    print()
    print(f"All files saved in: {run_dir}")
    print(f"  {'invoices':20s} -> {invoices_path}")
    print(f"  {'counts_summary':20s} -> {counts_path}")


if __name__ == "__main__":
    main()
