MMC Convert - Xero Conversion Summary
======================================

SETUP
-----
1. pip install -r requirements.txt --break-system-packages
2. Copy .env.example to .env and fill in your Xero app's Client ID / Secret
   (Redirect URI in your Xero Developer app must exactly match XERO_REDIRECT_URI)
3. Run: uvicorn auth:app --reload --port 8000
4. Open http://localhost:8000 in your browser

FLOW
----
1. Login page shows MMC Convert logo + "Login with Xero"
2. After Xero login, you land on /dashboard (no data fetched yet)
3. Pick a date range preset (1+YTD / 2+YTD / 3+YTD) or Custom dates
4. Click "Generate Summary" -> this is the ONLY time data is fetched from Xero
5. Results show inline on the same page: company info, chart of accounts
   (simple line-by-line list), features (tracking/inventory/payroll),
   and transaction counts within the selected period.

NOTES
-----
- 1+YTD / 2+YTD / 3+YTD are calculated from the org's Financial Year End
  (fetched live from Xero's /Organisation endpoint) minus 1/2/3 years, through today.
- Tokens persist in xero_tokens.db (SQLite) so re-login isn't needed until the
  refresh token itself expires (~60 days of inactivity).
- Xero API rate limits (429) are auto-retried using the Retry-After header.
