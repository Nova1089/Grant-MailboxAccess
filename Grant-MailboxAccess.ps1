<#
This script grants access to a mailbox to a provided list of users.
#>

# functions
function Show-Introduction
{
    Write-Host "This script grants access to a mailbox to a list of users." -ForegroundColor DarkCyan
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule($moduleName)
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..."
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required."
        $confirmInstall = Read-Host -Prompt "Would you like to install it? (y/n)"
    }
    while ($confirmInstall -inotmatch "(?<!\S)y(?!\S)") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Throw "Please run script with admin privileges. 
            1. Open Powershell as admin.
            2. CD into script directory.
            3. Run .\scriptname.ps1"
    }
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Read-Host -Prompt "Failed to connect to Exchange Online. Press Enter to try again"
        }
    }
}

function Prompt-MailboxIdentifier
{
    do
    {
        $mailboxIdentifier = Read-Host "Enter name or email of mailbox"
        $mailbox = TryGet-Mailbox -mailboxIdentifier $mailboxIdentifier -tellWhenFound
    }
    while ($null -eq $mailbox)

    return $mailbox.UserPrincipalName
}

function TryGet-Mailbox($mailboxIdentifier, [switch]$tellWhenFound)
{
    $mailbox = Get-EXOMailbox -Identity $mailboxIdentifier -ErrorAction SilentlyContinue

    if ($null -eq $mailbox)
    {
        Write-Warning "User or mailbox not found: $mailboxIdentifier."
        return $null
    }

    if ($tellWhenFound)
    {
        Write-Host "Found -- Display Name: $($mailbox.DisplayName) -- Email: $($mailbox.PrimarySmtpAddress) -- Type: $($mailbox.RecipientTypeDetails)`n" -ForegroundColor DarkCyan
    }
    return $mailbox
}

function Prompt-UserListInputMethod
{
    Write-Host "Choose user input method:"
    do
    {
        $choice = Read-Host ("[1] Provide text file. (Users listed by full name or email, separated by new line.)`n" +
            "[2] Enter user list manually.`n")        
    }
    while ($choice -notmatch '(?<!\S)[12](?!\S)') # regex matches a 1 or 2 but allows whitespace

    return [int]$choice
}

function Get-UsersFromTXT
{
    do 
    {
        $path = Read-Host "Enter path to txt file. (i.e. C:\UserList.txt)"
        $userList = Get-Content -Path $path -ErrorAction SilentlyContinue 
        
        if ($null -eq $userList)
        {
            Write-Warning "File not found or contents are empty."
            $keepGoing = $true
            continue
        }
        else
        {
            Write-Host "User list found." -ForegroundColor DarkCyan
            $keepGoing = $false
        }

        $finalUserList = New-Object -TypeName System.Collections.Generic.List[string]
        $i = 0
        foreach ($user in $userList)
        {
            if (($null -eq $user) -or ("" -eq $user)) { continue }            
            
            if ($null -eq (TryGet-Mailbox $user))
            {                
                $keepGoing = Prompt-YesOrNo "Would you like to fix the file and try again?"
                if ($keepGoing) { break }
            }
            else
            {
                $finalUserList.Add($user)
            }
            $i++
            Write-Progress -Activity "Looking up users..." -Status "$i users checked."
        }
    }
    while ($keepGoing)

    return $finalUserList
}

function Prompt-YesOrNo($question)
{
    do
    {
        $response = Read-Host "$question y/n"
    }
    while ($response -inotmatch '(?<!\S)[yn](?!\S)') # regex matches y or n but allows spaces

    if ($response -imatch '(?<!\S)y(?!\S)') # regex matches a y but allows spaces
    {
        return $true
    }
    return $false   
}

function Get-UsersManually
{
    $userList = New-Object -TypeName System.Collections.Generic.List[string]

    while ($true)
    {
        $response = Read-Host "Enter a user (full name or email) or type `"done`""
        if ($response -imatch '(?<!\S)done(?!\S)') { break } # regex matches the word done but allows spaces
        if ($null -eq (TryGet-Mailbox $response -tellWhenFound)) { continue }
        $userList.Add($response)
    }

    return $userList
}

function Prompt-PermissionsToGrant
{
    do
    {
        $grantFullAccess = Prompt-YesOrNo "Grant `"full access`" (read and manage)?"
        $shouldGrantSendAs = Prompt-YesOrNo "Grant `"send as`" access?"

        if (($grantFullAccess -eq $false) -and ($shouldGrantSendAs -eq $false))
        {
            Write-Warning "No permissions were selected to grant."
            $keepGoing = $true
        }
        else
        {
            $keepGoing = $false
        }
    }
    while ($keepGoing)

    return [PSCustomObject]@{
        grantFullAccess = $grantFullAccess
        shouldGrantSendAs     = $shouldGrantSendAs
    }
}

function Grant-AccessToMailbox($mailboxIdentifier, $userList, [bool]$grantFullAccess, [bool]$shouldGrantSendAs)
{
    $i = 0
    foreach ($user in $userList)
    {        
        Write-Progress -Activity "Granting access to mailbox..." -Status "$i users granted."
        if ($grantFullAccess)
        {
            Add-MailboxPermission -Identity $mailboxIdentifier -User $user -AccessRights FullAccess -Confirm:$false -WarningAction SilentlyContinue | Out-Null
        }

        if ($shouldGrantSendAs)
        {
            Add-RecipientPermission -Identity $mailboxIdentifier -Trustee $user -AccessRights SendAs -Confirm:$false -WarningAction SilentlyContinue | Out-Null
        }
        $i++        
    }
    Write-Progress -Activity "Granting access to mailbox..." -Status "$i users granted."
    Write-Host "Finished granting access to $i users. (If they didn't already have the access.)" -ForegroundColor DarkCyan
}

# main
Show-Introduction
Use-Module("ExchangeOnlineManagement")
TryConnect-ExchangeOnline

$mailboxIdentifier = Prompt-MailboxIdentifier
$userListInputMethod = Prompt-UserListInputMethod
switch ($userListInputMethod)
{
    1 { $userList = Get-UsersFromTXT }
    2 { $userList = Get-UsersManually }
}

$permissionSelections = Prompt-PermissionsToGrant
Grant-AccessToMailbox -mailboxIdentifier $mailboxIdentifier -userList $userList -grantFullAccess $permissionSelections.grantFullAccess -shouldGrantSendAs $permissionSelections.shouldGrantSendAs

Read-Host -Prompt "Press Enter to exit"