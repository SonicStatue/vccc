Import-Module ActiveDirectory

# =========================
# Local Configuration
# =========================
$MailDomain        = "bottleops.xyz"
$DefaultOU         = "OU=Users,OU=Corp,DC=bottleops,DC=xyz"
$DistributionList  = "employees@bottleops.xyz"
$DefaultFolders    = @("Sent", "Drafts", "Trash", "Junk", "Archive")
$HmailAdminUser    = "Administrator"
$DefaultAdPassword = 'P@$$word123'

# =========================
# Helpers
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
    if (-not $hmail) {
        throw "Failed to create hMailServer COM object."
    }

    $hmail.Authenticate($HmailAdminUser, $adminPass)

    if ($hmail.Domains.Count -lt 1) {
        throw "Connected to hMailServer, but no domains were returned."
    }

    Write-Host "Connected to hMailServer successfully."
    return $hmail
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

function Get-HMailAccountByAddress {
    param(
        $Domain,
        [string]$Address
    )

    if (-not $Domain) {
        return $null
    }

    for ($i = 0; $i -lt $Domain.Accounts.Count; $i++) {
        $acct = $Domain.Accounts.Item($i)
        if ($acct.Address -ieq $Address) {
            return $acct
        }
    }

    return $null
}

function Get-HMailDistributionListByAddress {
    param(
        $Domain,
        [string]$Address
    )

    if (-not $Domain) {
        return $null
    }

    for ($i = 0; $i -lt $Domain.DistributionLists.Count; $i++) {
        $list = $Domain.DistributionLists.Item($i)
        if ($list.Address -ieq $Address) {
            return $list
        }
    }

    return $null
}

function Ensure-ImapFolders {
    param(
        $Account,
        [string[]]$FolderNames
    )

    $created = @()
    $existing = @{}

    for ($i = 0; $i -lt $Account.IMAPFolders.Count; $i++) {
        $folder = $Account.IMAPFolders.Item($i)
        $existing[$folder.Name.ToLower()] = $true
    }

    foreach ($folderName in $FolderNames) {
        if (-not $existing.ContainsKey($folderName.ToLower())) {
            $null = $Account.IMAPFolders.Add($folderName)
            Write-Host "Created folder: $folderName"
            $created += $folderName
        }
        else {
            Write-Host "Folder already exists: $folderName"
        }
    }

    return $created
}

function Add-MailboxToDistributionList {
    param(
        $DistributionListObject,
        [string]$RecipientAddress
    )

    if (-not $DistributionListObject) {
        throw "Distribution list object was null."
    }

    for ($i = 0; $i -lt $DistributionListObject.Recipients.Count; $i++) {
        $recipient = $DistributionListObject.Recipients.Item($i)
        if ($recipient.RecipientAddress -ieq $RecipientAddress) {
            Write-Host "Distribution membership already exists: $RecipientAddress"
            return $false
        }
    }

    $newRecipient = $DistributionListObject.Recipients.Add()
    $newRecipient.RecipientAddress = $RecipientAddress
    $newRecipient.Save()

    Write-Host "Added to distribution list: $RecipientAddress"
    return $true
}

# =========================
# Input
# =========================
Write-Host ""
Write-Host "=== BottleOps Local User Provisioning ==="
Write-Host ""

$FirstName = Read-Host "Enter first name"
$LastName  = Read-Host "Enter last name"
$Username  = Read-Host "Enter username (sAMAccountName)"

if ([string]::IsNullOrWhiteSpace($FirstName) -or
    [string]::IsNullOrWhiteSpace($LastName) -or
    [string]::IsNullOrWhiteSpace($Username)) {
    throw "First name, last name, and username are required."
}

$DisplayName = "$FirstName $LastName"
$UserPrincipalName = "$Username@$MailDomain"
$MailboxAddress = "$Username@$MailDomain"

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
$FolderStatus = ""
$DistroStatus = ""
$CreatedFolders = @()

# =========================
# Step 1 - AD User
# =========================
Write-Host ""
Write-Host "=== Step 1: Active Directory User ==="

try {
    $existingUser = Get-ADUser -Identity $Username -ErrorAction SilentlyContinue
}
catch {
    $existingUser = $null
}

if ($existingUser) {
    Write-Host "AD user already exists: $Username"
    $AdStatus = "Already existed"
}
else {
    try {
        New-ADUser `
            -Name $DisplayName `
            -GivenName $FirstName `
            -Surname $LastName `
            -DisplayName $DisplayName `
            -SamAccountName $Username `
            -UserPrincipalName $UserPrincipalName `
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
# Step 2 - hMailServer Local COM Connection
# =========================
Write-Host ""
Write-Host "=== Step 2: hMailServer ==="
Write-Host "MailDomain is: [$MailDomain]"

try {
    $adminPass = Read-PlaintextPassword "Enter hMailServer Administrator password"

    $hmail = New-Object -ComObject hMailServer.Application
    if (-not $hmail) {
        throw "Failed to create hMailServer COM object."
    }

    $hmail.Authenticate($HmailAdminUser, $adminPass)
    Write-Host "Connected to hMailServer successfully."

    Write-Host "Domain count returned by hMailServer: $($hmail.Domains.Count)"

    $domain = $null

    for ($i = 0; $i -lt $hmail.Domains.Count; $i++) {
        $d = $hmail.Domains.Item($i)
        Write-Host ("Found hMailServer domain [{0}]: [{1}]" -f $i, $d.Name)

        if ($d.Name -ieq $MailDomain) {
            $domain = $d
            Write-Host "Matched target mail domain: [$MailDomain]"
            break
        }
    }

    if (-not $domain) {
        throw "Mail domain not found: $MailDomain"
    }

    Write-Host "Using mail domain: $($domain.Name)"
}
catch {
    throw "Failed during hMailServer connection/domain lookup. $($_.Exception.Message)"
}
# =========================
# Step 3 - Mailbox
# =========================
Write-Host ""
Write-Host "=== Step 3: Mailbox ==="

try {
    $account = Get-HMailAccountByAddress -Domain $domain -Address $MailboxAddress

    if ($account) {
        Write-Host "Mailbox already exists: $MailboxAddress"
        $MailboxStatus = "Already existed"
    }
    else {
        $account = $domain.Accounts.Add()
        $account.Address = $MailboxAddress

        # Placeholder only. Actual user auth is expected to come from AD.
        $account.Password = "TempMailbox123!"
        $account.Active = $true
        $account.MaxSize = 0
        $account.Save()

        Write-Host "Created mailbox: $MailboxAddress"
        $MailboxStatus = "Created"
    }
}
catch {
    throw "Failed during mailbox creation/check for '$MailboxAddress'. $($_.Exception.Message)"
}

# =========================
# Step 4 - Default IMAP Folders
# =========================
Write-Host ""
Write-Host "=== Step 4: Default Folders ==="

try {
    $CreatedFolders = Ensure-ImapFolders -Account $account -FolderNames $DefaultFolders

    if ($CreatedFolders.Count -gt 0) {
        $FolderStatus = "Created: $($CreatedFolders -join ', ')"
    }
    else {
        $FolderStatus = "All already existed"
    }
}
catch {
    throw "Failed while ensuring folders for '$MailboxAddress'. $($_.Exception.Message)"
}

# =========================
# Step 5 - Distribution List
# =========================
Write-Host ""
Write-Host "=== Step 5: Distribution List ==="

try {
    $distList = Get-HMailDistributionListByAddress -Domain $domain -Address $DistributionList

    if (-not $distList) {
        throw "Distribution list not found: $DistributionList"
    }

    $added = Add-MailboxToDistributionList -DistributionListObject $distList -RecipientAddress $MailboxAddress

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
Write-Host "UPN:          $UserPrincipalName"
Write-Host "Mailbox:      $MailboxAddress"
Write-Host "AD User:      $AdStatus"
Write-Host "Mailbox:      $MailboxStatus"
Write-Host "Folders:      $FolderStatus"
Write-Host "Distribution: $DistroStatus ($DistributionList)"
Write-Host ""
Write-Host "Provisioning complete."
