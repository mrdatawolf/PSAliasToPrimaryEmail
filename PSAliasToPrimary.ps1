param (
    [string]$Admin
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This script requires PowerShell 7 or higher. Please upgrade your PowerShell version."
    exit
}

# Function to test and install modules if necessary
function Test-ModuleInstallation {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "The $ModuleName module is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Force -Scope CurrentUser
        
        return $false
    } else {
        Write-Host "Importing $ModuleName..." -ForegroundColor Green
        try {
            Import-Module $ModuleName -ErrorAction Stop
        } catch {
            Write-Host "Failed to import $ModuleName. Please ensure all dependencies are installed." -ForegroundColor Red
            return $false
        }
    }

    return $true
}

# Clear existing modules to avoid conflicts
Remove-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
Remove-Module -Name Microsoft.Graph -ErrorAction SilentlyContinue

# List of required modules
$modules = @("ExchangeOnlineManagement", "Microsoft.Graph.Users", "Microsoft.Graph.Authentication", "Az.Accounts", "Az.Resources")
foreach ($module in $modules) {
    $result = Test-ModuleInstallation -ModuleName $module
    if (-not $result) {
        Write-Host "Please restart the script after installing the required modules." -ForegroundColor Red
        exit
    }
}

# Function to connect to Exchange Online with MFA
function Connect-Exchange {
    param (
        [string]$Admin
    )
    $UserCredential = Get-Credential -UserName $Admin
    Connect-ExchangeOnline -UserPrincipalName $UserCredential.UserName -Credential $UserCredential -ShowProgress $true
}

# Function to connect to Microsoft Graph with device code authentication
function Connect-Graph {
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -UseDeviceAuthentication
}

# Function to update email addresses and SIP address
function Update-EmailAddresses {
    param (
        [string]$Username,
        [string]$Alias
    )

    $Mailbox = Get-Mailbox -Identity $Username
    $EmailAddresses = $Mailbox.EmailAddresses

    if ($EmailAddresses -notcontains "SMTP:$Alias") {
        # Add the alias if it does not exist
        $EmailAddresses += "smtp:$Alias"
    }

    # Set the new primary email address to the alias
    $EmailAddresses = $EmailAddresses -replace "SMTP:", "smtp:"
    $EmailAddresses = $EmailAddresses -replace "smtp:$Alias", "SMTP:$Alias"

    # Update the mailbox with the new email addresses
    Set-Mailbox -Identity $Username -EmailAddresses $EmailAddresses

    Write-Host "Primary email address changed successfully for $Username."
}

# Function to update MFA settings
function Update-MfaSettings {
    param (
        [string]$UserId
    )

    # Require user to provide contact methods again
    $body = @{
        "method" = "resetPassword"
    }
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/users/$UserId/authentication/methods/resetPassword" -Body ($body | ConvertTo-Json)

    # Delete all existing app passwords
    $appPasswords = Get-MgUserAuthenticationMethod -UserId $UserId | Where-Object {$_.OdataType -eq "#microsoft.graph.passwordAuthenticationMethod"}
    foreach ($appPassword in $appPasswords) {
        Remove-MgUserAuthenticationMethod -UserId $UserId -AuthenticationMethodId $appPassword.Id
    }

    Write-Host "MFA settings updated successfully for $UserId."
}

# Function to prompt for user input and process the changes
function Invoke-UserInput {
    do {
        $Username = Read-Host -Prompt "Enter the username (UPN) of the user"
        $Alias = Read-Host -Prompt "Enter the new primary email alias"

        Update-EmailAddresses -Username $Username -Alias $Alias

        # Get the user ID from the username
        #$User = Get-MgUser -Filter "userPrincipalName eq '$Alias'"
        #if ($User) {
        #    $UserId = $User.Id
        #    Update-MfaSettings -UserId $UserId
        #} else {
        #    Write-Host "User not found: $Alias" -ForegroundColor Red
        #}

        $Continue = Read-Host -Prompt "Do you want to process another user? (yes/no)"
    } while ($Continue -eq "yes")
}

# Main script execution
Connect-Exchange -Admin $Admin
Connect-Graph
Invoke-UserInput
Disconnect-ExchangeOnline -Confirm:$false