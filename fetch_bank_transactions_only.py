"""
fetch_bank_transactions_only.py
---------------------------------------------------------------
Companion script for MMC Convert (auth.py).

Dumps raw JSON for every bank/payment-related Xero endpoint into ONE
timestamped folder, plus a combined counts summary. Each endpoint is
fetched exactly once (no endpoint is ever re-fetched for a sub-type -
e.g. BankTransactions is one paginated pass covering SPEND, RECEIVE,
SPEND-OVERPAYMENT, RECEIVE-OVERPAYMENT, SPEND-PREPAYMENT and
RECEIVE-PREPAYMENT all via the "Type" field on each record).

Endpoints covered (verified against Xero's OpenAPI spec):
  - /BankTransactions  (paginated, HAS LineItems)
  - /BankTransfers     (NOT paginated - single call, NO LineItems)
  - /Payments          (paginated, NO LineItems - PaymentType field
                         covers invoice/bill/credit-note/overpayment/
                         prepayment payments, all in one endpoint)
  - /Overpayments      (paginated, HAS LineItems - own document type,
                         distinct from the SPEND/RECEIVE-OVERPAYMENT
                         BankTransactions records)
  - /Prepayments       (paginated, HAS LineItems - own document type,
                         distinct from the SPEND/RECEIVE-PREPAYMENT
                         BankTransactions records)

Payroll is NOT included here - it's a completely separate API
(payroll.xro/2.0, with different base URLs per country: AU/UK/NZ) and
is out of scope for this bank/payment-transactions dump. A lightweight
probe (same /Employees check auth.py already does) is saved for
reference only.

Does NOT modify auth.py - reuses its existing get_valid_access_token,
xero_get, API_BASE, get_token_row functions as-is.

Usage:
  python fetch_bank_transactions_only.py --tenant-id <TENANT_ID>

  Optional:
    --from-date       YYYY-MM-DD (custom range start)
    --to-date         YYYY-MM-DD (custom range end)
    --output-dir      Parent folder to create the new run-folder inside (default: current directory)

Example:
  python fetch_bank_transactions_only.py --tenant-id 1314cf60-8470-46a1-a23b-b5d28ac77fa7 --from-date 2025-04-01 --to-date 2026-08-12
"""

import os
import re
import json
import asyncio
import argparse
from datetime import date, datetime

import httpx

# --- Reuse auth.py's existing logic untouched (must run from same folder) ---
from auth import get_valid_access_token, xero_get, get_token_row, API_BASE


def safe_folder_name(name: str) -> str:
    name = (name or "").strip() or "Xero_Bank_Transactions"
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = re.sub(r"\s+", "_", name)
    return name


def build_date_filter(from_date: str, to_date: str) -> str | None:
    if not (from_date and to_date):
        return None
    start = datetime.strptime(from_date, "%Y-%m-%d").date()
    end = datetime.strptime(to_date, "%Y-%m-%d").date()
    return f"Date >= DateTime({start.year},{start.month},{start.day}) AND Date <= DateTime({end.year},{end.month},{end.day})"


async def fresh_headers(tenant_id: str, headers: dict):
    fresh_token = await get_valid_access_token(tenant_id)
    headers["Authorization"] = f"Bearer {fresh_token}"


# --------------------------------------------------------------------
# /BankTransactions - paginated, HAS LineItems
# --------------------------------------------------------------------
async def fetch_bank_transactions(client, tenant_id, headers, date_filter):
    all_txns = []
    page = 1
    while True:
        await fresh_headers(tenant_id, headers)
        params = {"summaryOnly": "false", "page": page}
        if date_filter:
            params["where"] = date_filter
        print(f"[BANK] BankTransactions - page {page}...")
        resp = await xero_get(client, f"{API_BASE}/BankTransactions", headers, params=params)
        data = resp.json().get("BankTransactions", [])
        if not data:
            print(f"[BANK] BankTransactions - page {page} empty, stopping.")
            break
        print(f"[BANK] BankTransactions - page {page} returned {len(data)} records")
        all_txns.extend(data)
        page += 1
        await asyncio.sleep(0.25)
    print(f"[BANK] Total BankTransactions fetched = {len(all_txns)}")
    return all_txns


# --------------------------------------------------------------------
# /BankTransfers - NOT paginated (Xero returns everything in one call,
# no "page" param exists for this endpoint), NO LineItems
# --------------------------------------------------------------------
async def fetch_bank_transfers(client, tenant_id, headers, date_filter):
    await fresh_headers(tenant_id, headers)
    params = {}
    if date_filter:
        params["where"] = date_filter
    print("[TRANSFERS] BankTransfers - fetching (single call, not paginated)...")
    resp = await xero_get(client, f"{API_BASE}/BankTransfers", headers, params=params)
    data = resp.json().get("BankTransfers", [])
    print(f"[TRANSFERS] BankTransfers returned {len(data)} records")
    return data


# --------------------------------------------------------------------
# /Payments - paginated, NO LineItems. PaymentType covers invoice
# payments, bill payments, credit-note allocations, and
# overpayment/prepayment payments - all in this single endpoint.
# --------------------------------------------------------------------
async def fetch_payments(client, tenant_id, headers, date_filter):
    all_payments = []
    page = 1
    while True:
        await fresh_headers(tenant_id, headers)
        params = {"page": page}
        if date_filter:
            params["where"] = date_filter
        print(f"[PAYMENTS] Payments - page {page}...")
        resp = await xero_get(client, f"{API_BASE}/Payments", headers, params=params)
        data = resp.json().get("Payments", [])
        if not data:
            print(f"[PAYMENTS] Payments - page {page} empty, stopping.")
            break
        print(f"[PAYMENTS] Payments - page {page} returned {len(data)} records")
        all_payments.extend(data)
        page += 1
        await asyncio.sleep(0.25)
    print(f"[PAYMENTS] Total Payments fetched = {len(all_payments)}")
    return all_payments


# --------------------------------------------------------------------
# /Overpayments - paginated, HAS LineItems. Own document type, distinct
# from the SPEND-OVERPAYMENT/RECEIVE-OVERPAYMENT BankTransactions
# records (those capture the bank movement, this is the source doc).
# --------------------------------------------------------------------
async def fetch_overpayments(client, tenant_id, headers, date_filter):
    all_ops = []
    page = 1
    while True:
        await fresh_headers(tenant_id, headers)
        params = {"page": page}
        if date_filter:
            params["where"] = date_filter
        print(f"[OVERPAY] Overpayments - page {page}...")
        resp = await xero_get(client, f"{API_BASE}/Overpayments", headers, params=params)
        data = resp.json().get("Overpayments", [])
        if not data:
            print(f"[OVERPAY] Overpayments - page {page} empty, stopping.")
            break
        print(f"[OVERPAY] Overpayments - page {page} returned {len(data)} records")
        all_ops.extend(data)
        page += 1
        await asyncio.sleep(0.25)
    print(f"[OVERPAY] Total Overpayments fetched = {len(all_ops)}")
    return all_ops


# --------------------------------------------------------------------
# /Prepayments - paginated, HAS LineItems. Own document type, distinct
# from the SPEND-PREPAYMENT/RECEIVE-PREPAYMENT BankTransactions records.
# --------------------------------------------------------------------
async def fetch_prepayments(client, tenant_id, headers, date_filter):
    all_pps = []
    page = 1
    while True:
        await fresh_headers(tenant_id, headers)
        params = {"page": page}
        if date_filter:
            params["where"] = date_filter
        print(f"[PREPAY] Prepayments - page {page}...")
        resp = await xero_get(client, f"{API_BASE}/Prepayments", headers, params=params)
        data = resp.json().get("Prepayments", [])
        if not data:
            print(f"[PREPAY] Prepayments - page {page} empty, stopping.")
            break
        print(f"[PREPAY] Prepayments - page {page} returned {len(data)} records")
        all_pps.extend(data)
        page += 1
        await asyncio.sleep(0.25)
    print(f"[PREPAY] Total Prepayments fetched = {len(all_pps)}")
    return all_pps


# --------------------------------------------------------------------
# Payroll - lightweight best-effort probe only (NOT a full transaction
# dump). Payroll is a separate API with country-specific base URLs
# (payroll.xro/2.0 for AU/NZ/UK each differ) - fully implementing
# PayRuns/Payslips fetching is out of scope here.
# --------------------------------------------------------------------
async def probe_payroll(client, headers):
    print("[PAYROLL] Probing /payroll.xro/2.0/Employees (enabled check only)...")
    try:
        resp = await client.get(
            "https://api.xero.com/payroll.xro/2.0/Employees",
            headers=headers,
        )
        enabled = resp.status_code == 200
        count = len(resp.json().get("Employees", [])) if enabled else 0
    except Exception as e:
        enabled = False
        count = 0
        print(f"[PAYROLL] Probe failed: {e}")
    print(f"[PAYROLL] Payroll enabled={enabled}, employee_count={count} (transactions NOT fetched - out of scope)")
    return {"enabled": enabled, "employee_count": count, "note": "Full payroll transaction dump not implemented - separate country-specific API."}


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
    date_filter = build_date_filter(from_date, to_date)

    async with httpx.AsyncClient(timeout=60) as client:
        bank_txns = await fetch_bank_transactions(client, tenant_id, headers, date_filter)
        bank_transfers = await fetch_bank_transfers(client, tenant_id, headers, date_filter)
        payments = await fetch_payments(client, tenant_id, headers, date_filter)
        overpayments = await fetch_overpayments(client, tenant_id, headers, date_filter)
        prepayments = await fetch_prepayments(client, tenant_id, headers, date_filter)
        payroll_probe = await probe_payroll(client, headers)

    return bank_txns, bank_transfers, payments, overpayments, prepayments, payroll_probe


def main():
    parser = argparse.ArgumentParser(description="Fetch all bank/payment-related endpoints from Xero into one folder.")
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

    bank_txns, bank_transfers, payments, overpayments, prepayments, payroll_probe = asyncio.run(
        run_all(args.tenant_id, args.from_date, args.to_date)
    )

    company_name = row.get("display_name") or row.get("tenant_name") or "Xero_Bank_Transactions"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"{safe_folder_name(company_name)}_BankTransactions_{timestamp}"
    run_dir = os.path.join(args.output_dir, folder_name)
    os.makedirs(run_dir, exist_ok=True)

    saved_paths = {
        "bank_transactions": save_json(run_dir, "bank_transactions_raw.json", bank_txns),
        "bank_transfers": save_json(run_dir, "bank_transfers_raw.json", bank_transfers),
        "payments": save_json(run_dir, "payments_raw.json", payments),
        "overpayments": save_json(run_dir, "overpayments_raw.json", overpayments),
        "prepayments": save_json(run_dir, "prepayments_raw.json", prepayments),
        "payroll_probe": save_json(run_dir, "payroll_probe.json", payroll_probe),
    }

    # --- Counts ---
    bank_txn_type_breakdown = {}
    for t in bank_txns:
        ty = t.get("Type", "UNKNOWN")
        bank_txn_type_breakdown[ty] = bank_txn_type_breakdown.get(ty, 0) + 1
    bank_txn_line_items = sum(len(t.get("LineItems", []) or []) for t in bank_txns)

    payment_type_breakdown = {}
    for p in payments:
        pt = p.get("PaymentType", "UNKNOWN")
        payment_type_breakdown[pt] = payment_type_breakdown.get(pt, 0) + 1

    overpayment_type_breakdown = {}
    for o in overpayments:
        ty = o.get("Type", "UNKNOWN")
        overpayment_type_breakdown[ty] = overpayment_type_breakdown.get(ty, 0) + 1
    overpayment_line_items = sum(len(o.get("LineItems", []) or []) for o in overpayments)

    prepayment_type_breakdown = {}
    for pp in prepayments:
        ty = pp.get("Type", "UNKNOWN")
        prepayment_type_breakdown[ty] = prepayment_type_breakdown.get(ty, 0) + 1
    prepayment_line_items = sum(len(pp.get("LineItems", []) or []) for pp in prepayments)

    # Only BankTransactions/Overpayments/Prepayments have a LineItems
    # array - Payments and BankTransfers are single-amount records with
    # no line-item breakdown, so they contribute 0 to this sum.
    combined_line_item_total = bank_txn_line_items + overpayment_line_items + prepayment_line_items

    counts = {
        "bank_transactions": {
            "record_count": len(bank_txns),
            "line_item_count": bank_txn_line_items,
            "type_breakdown": bank_txn_type_breakdown,
        },
        "bank_transfers": {
            "record_count": len(bank_transfers),
            "line_item_count": 0,  # BankTransfer has no LineItems field
        },
        "payments": {
            "record_count": len(payments),
            "line_item_count": 0,  # Payment has no LineItems field
            "payment_type_breakdown": payment_type_breakdown,
        },
        "overpayments": {
            "record_count": len(overpayments),
            "line_item_count": overpayment_line_items,
            "type_breakdown": overpayment_type_breakdown,
        },
        "prepayments": {
            "record_count": len(prepayments),
            "line_item_count": prepayment_line_items,
            "type_breakdown": prepayment_type_breakdown,
        },
        "payroll": payroll_probe,
        "combined_line_item_total": combined_line_item_total,
    }
    counts_path = save_json(run_dir, "counts_summary.json", counts)

    print()
    print("=" * 70)
    print(f"BankTransactions : records={len(bank_txns)}  line_items={bank_txn_line_items}  types={bank_txn_type_breakdown}")
    print(f"BankTransfers    : records={len(bank_transfers)}  (no LineItems)")
    print(f"Payments         : records={len(payments)}  (no LineItems)  types={payment_type_breakdown}")
    print(f"Overpayments     : records={len(overpayments)}  line_items={overpayment_line_items}  types={overpayment_type_breakdown}")
    print(f"Prepayments      : records={len(prepayments)}  line_items={prepayment_line_items}  types={prepayment_type_breakdown}")
    print(f"Payroll (probe)  : {payroll_probe}")
    print("-" * 70)
    print(f"Combined LineItems total (BankTransactions + Overpayments + Prepayments) = {combined_line_item_total}")
    print("=" * 70)
    print()
    print(f"All files saved in: {run_dir}")
    for label, path in saved_paths.items():
        print(f"  {label:20s} -> {path}")
    print(f"  {'counts_summary':20s} -> {counts_path}")


if __name__ == "__main__":
    main()
