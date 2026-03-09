Import-Module ActiveDirectory

# =========================
# Configuration
# =========================
$MailDomain        = "bottleops.xyz"
$DistributionList  = "employees@bottleops.xyz"
$DefaultOU         = "OU=Users,OU=Corp,DC=bottleops,DC=xyz"
$DefaultFolders    = @("Sent", "Drafts", "Trash", "Junk", "Archive")
$HmailAdminUser    = "Administrator"
$DefaultAdPassword = 'P@$$word123'

# =========================
# Helper Functions
# =========================
function Read-PlaintextPassword {
    param([string]$Prompt)

    $secure = Read-Host $Prompt -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)

    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Get-HMailConnection {
    $adminPass = Read-PlaintextPassword "Enter hMailServer Administrator password"
    $hmail = New-Object -ComObject hMailServer.Application
    $hmail.Authenticate($HmailAdminUser, $adminPass)
    return $hmail
}

function Get-HMailAccountByAddress {
    param($Domain, [string]$Address)

    try {
        return $Domain.Accounts.ItemByAddress($Address)
    }
    catch {
        return $null
    }
}

function Get-HMailDistributionListByAddress {
    param($Domain, [string]$Address)

    try {
        return $Domain.DistributionLists.ItemByAddress($Address)
    }
    catch {
        return $null
    }
}

function Ensure-ImapFolders {
    param($Account, [string[]]$FolderNames)

    $created = @()
    $existing = @{}

    for ($i = 0; $i -lt $Account.IMAPFolders.Count; $i++) {
        $folder = $Account.IMAPFolders.Item($i)
        $existing[$folder.Name.ToLower()] = $true
    }

    foreach ($folderName in $FolderNames) {
        if (-not $existing.ContainsKey($folderName.ToLower())) {
            $Account.IMAPFolders.Add($folderName) | Out-Null
            Write-Host "Created folder: $folderName"
            $created += $folderName
        }
        else {
            Write-Host "Folder already exists: $folderName"
        }
    }

    return $created
}

function Get-HMailDomainByName {
    param(
        $Hmail,
        [string]$DomainName
    )

    if (-not $Hmail) {
        return $null
    }

    for ($i = 0; $i -lt $Hmail.Domains.Count; $i++) {
        $d = $Hmail.Domains.Item($i)
        if ($d.Name -ieq $DomainName) {
            return $d
        }
    }

    return $null
}

function Add-MailboxToDistributionList {
    param(
        $Domain,
        [string]$DistributionListAddress,
        [string]$RecipientAddress
    )

    $list = Get-HMailDistributionListByAddress -Domain $Domain -Address $DistributionListAddress
    if (-not $list) {
        throw "Distribution list not found: $DistributionListAddress"
    }

    for ($i = 0; $i -lt $list.Recipients.Count; $i++) {
        $recipient = $list.Recipients.Item($i)
        if ($recipient.RecipientAddress -ieq $RecipientAddress) {
            Write-Host "Distribution membership already exists: $RecipientAddress"
            return $false
        }
    }

    $newRecipient = $list.Recipients.Add()
    $newRecipient.RecipientAddress = $RecipientAddress
    $newRecipient.Save()

    Write-Host "Added to distribution list: $RecipientAddress -> $DistributionListAddress"
    return $true
}

# =========================
# Input
# =========================
Write-Host ""
Write-Host "=== BottleOps User Provisioning ==="
Write-Host ""

$FirstName = Read-Host "Enter first name"
$LastName  = Read-Host "Enter last name"
$Username  = Read-Host "Enter username (sAMAccountName)"

if ([string]::IsNullOrWhiteSpace($FirstName) -or
    [string]::IsNullOrWhiteSpace($LastName) -or
    [string]::IsNullOrWhiteSpace($Username)) {
    throw "First name, last name, and username are required."
}

$DisplayName   = "$FirstName $LastName"
$UserPrincipal = "$Username@$MailDomain"
$MailboxAddr   = "$Username@$MailDomain"

Write-Host ""
Write-Host "Press Enter to use default AD password: $DefaultAdPassword"
$AdPasswordPlain = Read-PlaintextPassword "Enter new AD password"
if ([string]::IsNullOrWhiteSpace($AdPasswordPlain)) {
    $AdPasswordPlain = $DefaultAdPassword
}

$AdPasswordSecure = ConvertTo-SecureString $AdPasswordPlain -AsPlainText -Force

# =========================
# Tracking
# =========================
$AdStatus = ""
$MailboxStatus = ""
$DistroStatus = ""
$CreatedFolders = @()

# =========================
# Step 1 - AD User
# =========================
Write-Host ""
Write-Host "=== Step 1: Active Directory ==="

try {
    $existingUser = Get-ADUser -Filter "SamAccountName -eq '$Username'" -ErrorAction Stop

    if ($existingUser) {
        Write-Host "AD user already exists: $Username"
        $AdStatus = "Already existed"
    }
}
catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    # Not typically hit by Get-ADUser -Filter, but harmless to keep explicit
    $existingUser = $null
}
catch {
    if ($_.Exception.Message -like "*Cannot find an object*") {
        $existingUser = $null
    }
    elseif ($_.FullyQualifiedErrorId -like "*ActiveDirectoryServer*") {
        throw
    }
    else {
        $existingUser = $null
    }
}

if (-not $existingUser) {
    try {
        New-ADUser `
            -Name $DisplayName `
            -GivenName $FirstName `
            -Surname $LastName `
            -DisplayName $DisplayName `
            -SamAccountName $Username `
            -UserPrincipalName $UserPrincipal `
            -Path $DefaultOU `
            -AccountPassword $AdPasswordSecure `
            -Enabled $true `
            -ChangePasswordAtLogon $true `
            -ErrorAction Stop

        Write-Host "Created AD user: $DisplayName ($Username)"
        $AdStatus = "Created"
    }
    catch {
        throw "Failed to create AD user '$Username'. $($_.Exception.Message)"
    }
}

# =========================
# Step 2 - hMailServer Connection
# =========================
Write-Host ""
Write-Host "=== Step 2: hMailServer Connection ==="

try {
    $hmail = Get-HMailConnection
    $domain = Get-HMailDomainByName -Hmail $hmail -DomainName $MailDomain

    if (-not $domain) {
        throw "Mail domain not found in hMailServer: $MailDomain"
    }
    
    Write-Host "Connected to hMailServer domain: $MailDomain"
}
catch {
    throw "Failed to connect to hMailServer. $($_.Exception.Message)"
}

# =========================
# Step 3 - Mailbox
# =========================
Write-Host ""
Write-Host "=== Step 3: Mailbox ==="

try {
    $account = Get-HMailAccountByAddress -Domain $domain -Address $MailboxAddr

    if ($account) {
        Write-Host "Mailbox already exists: $MailboxAddr"
        $MailboxStatus = "Already existed"
    }
    else {
        $account = $domain.Accounts.Add()
        $account.Address = $MailboxAddr

        # Placeholder only. Operational authentication should come from AD.
        $account.Password = "TempMailbox123!"
        $account.Active = $true
        $account.MaxSize = 0
        $account.Save()

        Write-Host "Created mailbox: $MailboxAddr"
        $MailboxStatus = "Created"
    }
}
catch {
    throw "Failed during mailbox step for '$MailboxAddr'. $($_.Exception.Message)"
}

# =========================
# Step 4 - IMAP Folders
# =========================
Write-Host ""
Write-Host "=== Step 4: Default Folders ==="

try {
    $CreatedFolders = Ensure-ImapFolders -Account $account -FolderNames $DefaultFolders
}
catch {
    throw "Failed while creating IMAP folders for '$MailboxAddr'. $($_.Exception.Message)"
}

# =========================
# Step 5 - Distribution List
# =========================
Write-Host ""
Write-Host "=== Step 5: Distribution List ==="

try {
    $added = Add-MailboxToDistributionList `
        -Domain $domain `
        -DistributionListAddress $DistributionList `
        -RecipientAddress $MailboxAddr

    if ($added) {
        $DistroStatus = "Added"
    }
    else {
        $DistroStatus = "Already a member"
    }
}
catch {
    throw "Failed during distribution list step. $($_.Exception.Message)"
}

# =========================
# Summary
# =========================
Write-Host ""
Write-Host "=== Provisioning Summary ==="
Write-Host "Display Name: $DisplayName"
Write-Host "Username:     $Username"
Write-Host "UPN:          $UserPrincipal"
Write-Host "Mailbox:      $MailboxAddr"
Write-Host "AD User:      $AdStatus"
Write-Host "Mailbox:      $MailboxStatus"

if ($CreatedFolders.Count -gt 0) {
    Write-Host "Folders:      Created -> $($CreatedFolders -join ', ')"
}
else {
    Write-Host "Folders:      All already existed"
}

Write-Host "Distribution: $DistroStatus ($DistributionList)"
Write-Host ""
Write-Host "Provisioning complete."
