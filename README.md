# Block-SenderDomains: Setup Guide

A drag-and-drop workflow for adding sender domains to your M365 Tenant Block
List. Drop emails into a desktop folder; a scheduled task runs at midnight
ET, extracts sender domains, blocks them indefinitely, and deletes the
processed files.

## What you'll end up with

- A folder at `Desktop\BlockSender` where you drag emails to be blocked
- A scheduled task that runs the script daily at midnight Eastern time
- Domains added to your Tenant Block List with no expiration
- A log file (`_blocklist.log`) inside the watch folder for audit
- A popup warning if you ever drop an email from a protected free-email
  domain (gmail, outlook.com, etc.) so nothing gets blocked accidentally

## One-time setup

There are three pieces: install the PowerShell module, register an app in
Entra ID for certificate-based auth, and create the scheduled task.

### Step 1 - Install the Exchange Online PowerShell module

In an elevated PowerShell window:

    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force

### Step 2 - Create a self-signed certificate

In the same PowerShell window (does NOT need to be elevated for this part):

    $cert = New-SelfSignedCertificate `
        -Subject "CN=SSLLC-BlockListAutomation" `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -KeyAlgorithm RSA `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears(10)

    Export-Certificate -Cert $cert -FilePath "$env:USERPROFILE\Desktop\SSLLC-BlockList.cer"

    Write-Host "Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green

Save the thumbprint - you'll paste it into the script in Step 4.

### Step 3 - Register the app in Entra ID

Go to https://entra.microsoft.com and sign in as a Global Admin.

1. **Identity > Applications > App registrations > New registration**
   - Name: `SSLLC Block List Automation`
   - Supported account types: **Accounts in this organizational directory only (single tenant)**
   - Redirect URI: leave blank
   - Click **Register**

2. On the Overview page, copy down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

3. **Certificates & secrets > Certificates > Upload certificate**
   - Upload the `SSLLC-BlockList.cer` file from your Desktop
   - Click **Add**

4. **API permissions > Add a permission**
   - Click **APIs my organization uses**
   - Search for and select **Office 365 Exchange Online**
   - Choose **Application permissions**
   - Check **Exchange.ManageAsApp**
   - Click **Add permissions**
   - Then click **Grant admin consent for [your tenant]** and confirm

5. Assign the Security Administrator role to the app:
   - Go to **Identity > Roles & admins > Roles & admins**
   - Find and click **Security Administrator**
   - Click **Add assignments**
   - Click **No member selected**, search for `SSLLC Block List Automation`,
     select it, click **Select**, then **Next** and **Assign**

### Step 4 - Set required environment variables

The script reads its four configuration values from environment variables.
Set them once with `setx` in an elevated Command Prompt (they persist across
reboots and sessions). Replace the placeholder values with your actual IDs and
thumbprint from Steps 2–3:

    setx BLOCKSENDER_TENANT_ID       "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    setx BLOCKSENDER_APP_ID          "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    setx BLOCKSENDER_CERT_THUMBPRINT "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
    setx BLOCKSENDER_ORGANIZATION    "yourcompany.onmicrosoft.com"

- `BLOCKSENDER_TENANT_ID` – Directory (tenant) ID from Entra
- `BLOCKSENDER_APP_ID` – Application (client) ID from Entra
- `BLOCKSENDER_CERT_THUMBPRINT` – Thumbprint from Step 2
- `BLOCKSENDER_ORGANIZATION` – Your `.onmicrosoft.com` domain

After running `setx`, open a **new** PowerShell window so the variables are
visible. To verify:

    $env:BLOCKSENDER_TENANT_ID
    $env:BLOCKSENDER_APP_ID
    $env:BLOCKSENDER_CERT_THUMBPRINT
    $env:BLOCKSENDER_ORGANIZATION

If unsure what your `.onmicrosoft.com` domain is: in the M365 admin center,
go to **Settings > Domains** - it is the `*.onmicrosoft.com` entry.

### Step 5 - Save and test the script

1.  Create folder `C:\Scripts` if it doesn't exist
2.  Save `Block-SenderDomains.ps1` to `C:\Scripts\`

Drop a junk email into `Desktop\BlockSender` (drag from New Outlook into
the folder; it should save as `.eml`). Then run:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Block-SenderDomains.ps1"

Expected outcome:

- Console shows the domain being queued and blocked
- The `.eml` file is deleted from the folder
- `_blocklist.log` shows a record of the run
- The domain appears in the M365 Defender portal under
  **Policies & rules > Threat policies > Tenant Allow/Block Lists > Domains & addresses**

If anything fails, check `_blocklist.log` for the error.

### Step 6 - Schedule it for midnight ET

Open **Task Scheduler** (`taskschd.msc`).

1.  Right-click **Task Scheduler Library > Create Task** (NOT "Basic Task")

2.  **General** tab:
    - Name: `Block Sender Domains`
    - Description: `Daily run: blocks domains from emails dropped in Desktop\BlockSender`
    - Select **Run only when user is logged on**
      (required because the script may show a popup and uses Outlook COM)
    - Check **Run with highest privileges**
    - Configure for: your Windows version

3.  **Triggers** tab > **New**:
    - Begin the task: **On a schedule**
    - **Daily**, start: today's date at `12:00:00 AM`
    - **IMPORTANT:** Check the **Synchronize across time zones** box if you
      see it. Then change the time zone selector to **(UTC-05:00) Eastern
      Time (US & Canada)**. This ensures it runs at midnight Eastern
      regardless of DST or whether you travel.
    - Recur every: `1 days`
    - Enabled: checked
    - Click **OK**

4.  **Actions** tab > **New**:
    - Action: **Start a program**
    - Program/script: `powershell.exe`
    - Add arguments:

          -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Scripts\Block-SenderDomains.ps1"

    - Click **OK**

5.  **Conditions** tab:
    - Uncheck **Start the task only if the computer is on AC power** (if
      you're on a laptop and want it to run on battery)
    - Check **Wake the computer to run this task** if you want it to fire
      even when the machine is asleep

6.  **Settings** tab:
    - Check **Allow task to be run on demand**
    - Check **Run task as soon as possible after a scheduled start is missed**
      (so if your laptop is off at midnight, it runs when you next log on)
    - Check **If the task is already running, then the following rule applies:
      Do not start a new instance**
    - Click **OK**

You'll be prompted for your password to save the task.

## Daily workflow

1. See spam or unwanted email in Outlook
2. Drag it from Outlook into `Desktop\BlockSender`
3. Done. At midnight ET the script runs, blocks the domain, deletes the file.

If you want to fire it manually instead of waiting:

- Open Task Scheduler
- Find the task, right-click > **Run**

Or run from PowerShell:

    Start-ScheduledTask -TaskName "Block Sender Domains"

## Verifying things work

After a run, check:

- **The folder**: processed `.eml`/`.msg` files should be gone
- **The log**: `Desktop\BlockSender\_blocklist.log` shows what happened
- **M365 Defender portal**:
  https://security.microsoft.com/tenantAllowBlockList?viewid=Domain
  New entries appear with notes like "Auto-blocked YYYY-MM-DD via drag-drop script"

## Troubleshooting

**"Connect-ExchangeOnline : The term ... is not recognized"**
The ExchangeOnlineManagement module isn't installed. Run Step 1 again.

**"AADSTS70011: The provided value for the input parameter 'scope' is not valid"**
The app permissions aren't granted or admin consent wasn't clicked. Recheck
Step 3, item 4.

**"Insufficient privileges to complete the operation"**
The Security Administrator role wasn't assigned to the service principal.
Recheck Step 3, item 5. Note: it can take a few minutes to propagate after
assignment.

**Script runs but nothing gets blocked, log says "Could not extract sender"**
New Outlook sometimes drags emails as `.url` shortcuts instead of real
files. Workaround: open the email, then File > Save As > save as `.eml`
to the watch folder.

**The .msg parsing fails with COM errors**
Outlook COM requires Classic Outlook or Office desktop bits to be
installed. If you're on a New-Outlook-only machine, drag as `.eml` instead
(the script handles both).

**Scheduled task shows "Last Run Result: 0x1" or similar**
Check the log file. Most often it's a missing or empty environment variable
(the log will name which one), or the cert isn't where the script expects it.
The cert lives in `Cert:\CurrentUser\My` - so the scheduled task must run as the same
user who created the cert.

**Cert expires**
The cert is valid for 10 years from creation. When it does expire, repeat
Step 2 to make a new one, upload the new `.cer` to the app registration
in Entra (Certificates & secrets), and update the `BLOCKSENDER_CERT_THUMBPRINT`
environment variable to the new thumbprint (`setx BLOCKSENDER_CERT_THUMBPRINT "NEW-THUMBPRINT"`).

## Notes on the protected domain list

The script ships with a list of common free-email/consumer domains it
refuses to block (gmail.com, outlook.com, yahoo.com, etc.). If you drop
an email from one of these, the script:

- Logs a warning
- Skips the block submission
- Pops up a dialog after the run summarizing what was skipped
- Leaves the source file in the folder so you can review

If you want to block a specific bad sender from a free-email domain,
either:

- Block the full sender address manually in the M365 admin center, or
- Edit the script to add full-address blocking logic for those cases

To customize the protected list, edit the `$ProtectedDomains` array
near the top of the script.

## Tenant Block List limits

The Tenant Allow/Block List has a cap of 500 sender entries. The script
will start failing when you hit that ceiling. If/when that happens, do a
periodic cleanup in the admin center, or write a companion script to
remove entries older than X months.
