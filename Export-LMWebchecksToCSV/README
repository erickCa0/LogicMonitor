LM Uptime Export – Quick Start

Use Export-LMUptimeCounts.ps1 to export LogicMonitor Uptime test counts into CSV files for reporting.

1. Prerequisites

Install the LogicMonitor PowerShell module (one-time setup):

Install-Module Logic.Monitor -Scope CurrentUser

You'll also need:

Your LogicMonitor portal name (the subdomain before .logicmonitor.com)
Authentication using either:
LMv1 Access ID/Access Key
Bearer Token
Saved LogicMonitor credentials
An existing Connect-LMAccount session
2. Run the Script
Using LMv1 Credentials
$env:LM_ACCESS_ID = "your-access-id"
$env:LM_ACCESS_KEY = "your-access-key"

.\Export-LMUptimeCounts.ps1 `
  -AccountName "your-portal" `
  -StartMonth 2026-01 `
  -EndMonth 2026-05 `
  -OutputDir .\exports\lm-uptime
Using a Bearer Token
$env:LM_BEARER_TOKEN = "your-token"

.\Export-LMUptimeCounts.ps1 `
  -AccountName "your-portal" `
  -BearerToken $env:LM_BEARER_TOKEN `
  -StartMonth 2026-05 `
  -EndMonth 2026-05 `
  -OutputDir .\exports\lm-uptime

Note: If your portal is acme.logicmonitor.com, use:

-AccountName "acme"
3. Review the Results

The script creates several CSV files. Most users only need:

monthly_uptime_test_counts.csv – Test counts by check
monthly_uptime_totals.csv – Monthly totals

Additional files provide device details, datasource instances, and warnings.

4. Choose a Counting Mode
Estimated (Recommended)

Fastest option. Calculates expected test counts based on polling intervals and enabled locations.

-Mode Estimated
Both (Estimated + Actual)

Also retrieves stored datapoints from LogicMonitor for comparison. Slower.

-Mode Both
5. Common Options
Task	Parameter
Web checks only	-Type uptimewebcheck
Ping checks only	-Type uptimepingcheck
Internal checks only	-IsInternal $true
External checks only	-IsInternal $false
Export CSV + JSON	-OutputFormat Both
6. Troubleshooting

If counts seem incorrect:

Check uptime_warnings.csv.
By default, missing polling intervals assume 5 minutes.
Override if needed:
-DefaultIntervalMinutes 10

Or specify custom polling interval properties:

-PollingIntervalProperties @("my.custom.interval","pollingInterval")
7. If the Script Won't Run

Scripts are blocked

Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

LogicMonitor module not found

Install-Module Logic.Monitor -Scope CurrentUser
Import-Module Logic.Monitor

Authentication errors

Verify:

-AccountName is only the portal subdomain.
Your credentials have read access.
Tokens or keys haven't expired.
8. Test Before Running Large Reports

Run a single month first:

.\Export-LMUptimeCounts.ps1 `
  -AccountName "your-portal" `
  -BearerToken $env:LM_BEARER_TOKEN `
  -StartMonth 2026-05 `
  -EndMonth 2026-05 `
  -OutputDir .\exports\lm-uptime-test

Then review:

.\exports\lm-uptime-test\monthly_uptime_totals.csv

If the results look correct, rerun with your full date range.
