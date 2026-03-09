# hMailServer IMAP Folder Provisioning Script

$domainName = "bottleops.xyz"

Write-Host ""
$username = Read-Host "Enter mailbox username (without domain)"

$mailboxAddress = "$username@$domainName"

Write-Host ""
$securePass = Read-Host "Enter hMailServer Administrator password" -AsSecureString
$adminPass = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
)

$foldersToCreate = @(
    "Sent",
    "Drafts",
    "Trash",
    "Junk",
    "Archive"
)

Write-Host ""
Write-Host "Connecting to hMailServer..."

$hmail = New-Object -ComObject hMailServer.Application
$hmail.Authenticate("Administrator", $adminPass)

$domain = $hmail.Domains.ItemByName($domainName)
$account = $domain.Accounts.ItemByAddress($mailboxAddress)

if (-not $account) {
    Write-Host ""
    Write-Host "Mailbox not found: $mailboxAddress"
    exit
}

Write-Host ""
Write-Host "Mailbox found. Checking folders..."

$existing = @{}

for ($i = 0; $i -lt $account.IMAPFolders.Count; $i++) {
    $folder = $account.IMAPFolders.Item($i)
    $existing[$folder.Name.ToLower()] = $true
}

foreach ($folderName in $foldersToCreate) {

    if (-not $existing.ContainsKey($folderName.ToLower())) {

        $newFolder = $account.IMAPFolders.Add()
        $newFolder.Name = $folderName
        $newFolder.Save()

        Write-Host "Created folder: $folderName"
    }
    else {
        Write-Host "Already exists: $folderName"
    }
}

Write-Host ""
Write-Host "Provisioning complete for $mailboxAddress"
