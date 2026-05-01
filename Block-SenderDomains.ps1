<#
.SYNOPSIS
    Watches a desktop folder for dropped .msg/.eml files, extracts sender
    domains, adds them to the M365 Tenant Block List indefinitely, and
    deletes the processed files.

.DESCRIPTION
    Designed to run as a daily Scheduled Task. Authenticates to Exchange
    Online via certificate-based app auth (no interactive prompts). Skips
    common free-email/protected domains and surfaces them as a popup
    warning so the user can review manually.

.NOTES
    Author : Sustainable Science, LLC
    Requires: ExchangeOnlineManagement module, Entra app registration
              with Exchange.ManageAsApp permission and Security Admin role.

    Required environment variables (set once; the script exits with an
    ERROR if any are missing or empty):

        BLOCKSENDER_TENANT_ID       - Directory (tenant) ID from Entra
        BLOCKSENDER_APP_ID          - Application (client) ID from Entra
        BLOCKSENDER_CERT_THUMBPRINT - Thumbprint of the cert in Cert:\CurrentUser\My
        BLOCKSENDER_ORGANIZATION    - Your tenant's .onmicrosoft.com domain

    Set them permanently with setx (run once in an elevated Command Prompt):

        setx BLOCKSENDER_TENANT_ID       "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
        setx BLOCKSENDER_APP_ID          "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
        setx BLOCKSENDER_CERT_THUMBPRINT "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
        setx BLOCKSENDER_ORGANIZATION    "yourcompany.onmicrosoft.com"

    Or for the current PowerShell session only:

        $env:BLOCKSENDER_TENANT_ID       = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
        $env:BLOCKSENDER_APP_ID          = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
        $env:BLOCKSENDER_CERT_THUMBPRINT = "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
        $env:BLOCKSENDER_ORGANIZATION    = "yourcompany.onmicrosoft.com"
#>

# ====== CONFIGURATION ======
$WatchFolder    = "$env:USERPROFILE\Desktop\BlockSender"
$LogFile        = "$env:USERPROFILE\Desktop\BlockSender\_blocklist.log"
$TenantId       = $env:BLOCKSENDER_TENANT_ID
$AppId          = $env:BLOCKSENDER_APP_ID
$CertThumbprint = $env:BLOCKSENDER_CERT_THUMBPRINT
$Organization   = $env:BLOCKSENDER_ORGANIZATION

# Free email / common domains we never want to block wholesale.
# If a sender from one of these domains is dropped, the script logs
# a warning and pops a dialog instead of submitting the block.
$ProtectedDomains = @(
    'gmail.com','googlemail.com','outlook.com','hotmail.com','live.com','msn.com',
    'yahoo.com','yahoo.co.uk','aol.com','icloud.com','me.com','mac.com',
    'proton.me','protonmail.com','gmx.com','mail.com','zoho.com','yandex.com',
    'comcast.net','verizon.net','att.net','sbcglobal.net','cox.net'
)
# ===========================

# Ensure folder exists
if (-not (Test-Path $WatchFolder)) {
    New-Item -ItemType Directory -Path $WatchFolder -Force | Out-Null
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $line
    Write-Host $line
}

function Get-SenderFromMsg {
    param([string]$Path)
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $msg = $namespace.OpenSharedItem($Path)
        $sender = $null

        # SMTP address is most reliable; fall back to SenderEmailAddress
        if ($msg.SenderEmailType -eq "EX") {
            try { $sender = $msg.Sender.GetExchangeUser().PrimarySmtpAddress } catch {}
        }
        if (-not $sender) { $sender = $msg.SenderEmailAddress }

        # Also pull from headers as a fallback (handles forwarded spam better)
        $headerFrom = $null
        try {
            $headers = $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
            if ($headers -match '(?im)^From:.*<([^>]+)>') { $headerFrom = $matches[1] }
            elseif ($headers -match '(?im)^From:\s*([^\s]+@[^\s]+)') { $headerFrom = $matches[1] }
        } catch {}

        $msg.Close(0)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msg) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null

        # Prefer header-derived sender if present (catches spoofed display addresses)
        if ($headerFrom) { return $headerFrom }
        return $sender
    } catch {
        Write-Log "Failed to parse $Path : $_" "ERROR"
        return $null
    }
}

function Get-SenderFromEml {
    param([string]$Path)
    try {
        $content = Get-Content $Path -Raw -ErrorAction Stop
        # Grab only the header block (everything before first blank line)
        $headerBlock = ($content -split "`r?`n`r?`n", 2)[0]
        if ($headerBlock -match '(?im)^From:.*<([^>]+)>') { return $matches[1] }
        if ($headerBlock -match '(?im)^From:\s*([^\s]+@[^\s]+)') { return $matches[1] }
    } catch {
        Write-Log "Failed to parse $Path : $_" "ERROR"
    }
    return $null
}

function Get-EmailDomain {
    param([string]$Email)
    if ($Email -match '@([^\s>]+)') { return $matches[1].ToLower().Trim() }
    return $null
}

# ====== MAIN ======
Write-Log "=== Run started ==="

# Validate required environment variables before attempting any operations
$_missingVars = @()
if (-not $TenantId)       { $_missingVars += 'BLOCKSENDER_TENANT_ID' }
if (-not $AppId)          { $_missingVars += 'BLOCKSENDER_APP_ID' }
if (-not $CertThumbprint) { $_missingVars += 'BLOCKSENDER_CERT_THUMBPRINT' }
if (-not $Organization)   { $_missingVars += 'BLOCKSENDER_ORGANIZATION' }

if ($_missingVars.Count -gt 0) {
    Write-Log "Missing required environment variable(s): $($_missingVars -join ', '). Set them and re-run." "ERROR"
    exit 1
}

$files = Get-ChildItem -Path $WatchFolder -File | Where-Object { $_.Extension -in '.msg','.eml' }

if ($files.Count -eq 0) {
    Write-Log "No files to process. Exiting."
    return
}

Write-Log "Found $($files.Count) file(s) to process."

# Collect domains first, then connect once
$domainsToBlock = @{}     # domain -> list of source files
$skippedProtected = @{}   # protected domains we didn't block
$failedFiles = @()

foreach ($file in $files) {
    $sender = if ($file.Extension -eq '.msg') {
        Get-SenderFromMsg -Path $file.FullName
    } else {
        Get-SenderFromEml -Path $file.FullName
    }

    if (-not $sender) {
        Write-Log "Could not extract sender from $($file.Name)" "WARN"
        $failedFiles += $file.FullName
        continue
    }

    $domain = Get-EmailDomain -Email $sender
    if (-not $domain) {
        Write-Log "Could not extract domain from sender '$sender' in $($file.Name)" "WARN"
        $failedFiles += $file.FullName
        continue
    }

    if ($ProtectedDomains -contains $domain) {
        Write-Log "PROTECTED DOMAIN: '$domain' (from $sender) in $($file.Name) - skipping" "WARN"
        if (-not $skippedProtected.ContainsKey($domain)) { $skippedProtected[$domain] = @() }
        $skippedProtected[$domain] += $file.FullName
        continue
    }

    if (-not $domainsToBlock.ContainsKey($domain)) { $domainsToBlock[$domain] = @() }
    $domainsToBlock[$domain] += $file.FullName
    Write-Log "Queued domain '$domain' (from $sender, file $($file.Name))"
}

if ($domainsToBlock.Count -eq 0) {
    Write-Log "Nothing to submit to block list."
} else {
    # Connect to Exchange Online
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline `
            -AppId $AppId `
            -CertificateThumbprint $CertThumbprint `
            -Organization $Organization `
            -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connected to Exchange Online."
    } catch {
        Write-Log "Failed to connect to Exchange Online: $_" "ERROR"
        return
    }

    # Pull existing block list once for dedup
    try {
        $existing = Get-TenantAllowBlockListItems -ListType Sender -Block |
                    Select-Object -ExpandProperty Value
        $existingSet = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]$existing, [System.StringComparer]::OrdinalIgnoreCase)
    } catch {
        Write-Log "Failed to retrieve existing block list: $_" "ERROR"
        $existingSet = [System.Collections.Generic.HashSet[string]]::new()
    }

    foreach ($domain in $domainsToBlock.Keys) {
        if ($existingSet.Contains($domain)) {
            Write-Log "'$domain' already on block list - skipping submission."
            # Still delete the source files since the intent is satisfied
            foreach ($f in $domainsToBlock[$domain]) {
                Remove-Item $f -Force -ErrorAction SilentlyContinue
                Write-Log "Deleted $f"
            }
            continue
        }

        try {
            New-TenantAllowBlockListItems `
                -ListType Sender `
                -Block `
                -Entries $domain `
                -NoExpiration `
                -Notes "Auto-blocked $(Get-Date -Format 'yyyy-MM-dd') via drag-drop script" `
                -ErrorAction Stop | Out-Null
            Write-Log "BLOCKED: $domain"

            foreach ($f in $domainsToBlock[$domain]) {
                Remove-Item $f -Force -ErrorAction SilentlyContinue
                Write-Log "Deleted $f"
            }
        } catch {
            Write-Log "Failed to block '$domain': $_" "ERROR"
        }
    }

    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

# Surface protected-domain warnings as a popup so they aren't missed
if ($skippedProtected.Count -gt 0) {
    Add-Type -AssemblyName PresentationFramework
    $msg = "The following protected domains were NOT blocked:`n`n"
    foreach ($d in $skippedProtected.Keys) {
        $msg += "  $d  ($($skippedProtected[$d].Count) message(s))`n"
    }
    $msg += "`nReview the source files in:`n$WatchFolder`n`nDelete them manually after review, or block the full sender address via the M365 admin center."
    [System.Windows.MessageBox]::Show($msg, "Block List: Protected Domains Skipped", "OK", "Warning") | Out-Null
}

Write-Log "=== Run complete ==="