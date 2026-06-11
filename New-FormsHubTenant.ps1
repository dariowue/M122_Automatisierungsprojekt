#logging functions
function Log-Action($message) {
    Write-Host $message
}

function Log-Info($message) {
    Write-Host "[INFO] $message" -ForegroundColor Cyan
}

function Log-Success($message) {
    Write-Host "[OK] $message" -ForegroundColor Green
}

function Log-Warn($message) {
    Write-Host "[WARN] $message" -ForegroundColor Yellow
}

function Log-Error($message) {
    Write-Host "[ERROR] $message" -ForegroundColor Red
}

#helper functions
function Get-Confirmation() {
    if ($confirmation) {
        $input = $(Write-Host "[WARN] Press enter to execute the action, enter 'n' to cancel " -ForegroundColor Yellow -NoNewLine; Read-Host) 
        if ($input -eq "n") {
            throw "Action was cancelled"
        }
        else {
            return
        }
    }
    return
}

function Get-IdWithRetry($name, $type, $upn) {
    Log-Info "Retrieving ${type} ID for ${name}"

    $id = $null
    for ($i = 1; $i -le 10; $i++) {
        try {
            switch($type) {
                "group" {
                    $id = (Get-MgGroup -Filter "displayName eq '${name}'" -ErrorAction Stop).Id
                }
                "user" {
                    $id = (Get-MgUser -UserId $upn -ErrorAction Stop).Id
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($id)) {
                Log-Success "Successfully retrieved ${type} ID for ${name}"
                return $id
            }

            Log-Info "${type} not available yet, waiting on Microsoft Graph"
            Start-Sleep -Seconds 2
        }
        catch {
            Log-Info "${type} not available yet, waiting on Microsoft Graph"
            Start-Sleep -Seconds 2
        }
    }
    throw "Tenant ${type} ID lookup for ${name} failed"
}

function Add-GroupMember($groupId, $groupName, $userId, $userName) {
    Log-Info "Adding ${userName} to ${groupName} group"

    New-MgGroupMember `
    -GroupId ${groupId} `
    -DirectoryObjectId ${userId} `
    -ErrorAction Stop

    Log-Success "Successfully added ${userName} to ${groupName} group"
}

#main functions
function New-TenantSecurityGroup($customer) {
    Log-Action "Creating tenant security group"
    Get-Confirmation

    $group = New-MgGroup `
    -DisplayName "tenant-${customer}" `
    -Description "Access to the ${customer} data" `
    -MailEnabled:$false `
    -MailNickname "tenant-${customer}" `
    -SecurityEnabled:$true `
    -ErrorAction Stop 

    if (-not $group.id) {
        throw "Tenant security group creation for ${customer} failed"
    }

    Log-Success "Successfully created group tenant-${customer} with id ${group.id}"
}

function Add-TenantGroupMembers($customer, $mailDomain) {
    Log-Action "Adding users to tenant security group"
    Get-Confirmation
    $groupName = "tenant-${customer}"

    $groupId = Get-IdWithRetry $groupName "group"

    $cockpitUserId = Get-IdWithRetry "Cloud Service User" "user" "mail@${env}.internal.com"

    Log-Info "Waiting for group to be fully available in Microsoft Graph before adding members"
    Start-Sleep -Seconds 10

    Add-GroupMember $groupId $groupName $cockpitUserId "Cloud Service User"

   $systemUserId = Get-IdWithRetry "Cloud System User" "user" "mail@${env}.internal.com"

    Log-Info "Waiting for changes to be fully processed in Microsoft Graph before adding next member"
    Start-Sleep -Seconds 10

    Add-GroupMember $groupId $groupName $systemUserId "Cloud System User"
}

function New-TenantServiceUser($customer, $password, $env, $mailDomain) {
    Log-Action "Creating tenant service user"
    Get-Confirmation

    $userPassword = @{
        Password = $password
        ForceChangePasswordNextSignIn = $false
    }

    $user = New-MgUser `
        -DisplayName "Cloud ${customer} Forms Service User ${env}" `
        -UserPrincipalName "mail@${env}.internal.com" `
        -MailNickname "${customer}" `
        -AccountEnabled `
        -PasswordProfile $userPassword `
        -ErrorAction Stop

    if (-not $user.Id) {
        throw "Tenant service user creation failed"
    }

    Log-Success "Successfully created user Cloud ${customer} Forms Service User ${env} with id ${user.Id}"
}

function Set-PasswordPolicy($customer, $env, $mailDomain) {
    Log-Action "Set password policy for tenant service user"
    Get-Confirmation

    $serviceUserId = Get-IdWithRetry "Cloud ${customer} Forms Service User ${env}" "user" "mail@${env}.internal.com"

    Update-MgUser -UserId $serviceUserId -PasswordPolicies "DisablePasswordExpiration" -ErrorAction Stop
    Log-Success "Successfully set password policy for tenant service user"
}

function Add-ServiceUserGroups ($customer, $env, $mailDomain) {
    Log-Action "Adding tenant service user to groups"
    Get-Confirmation
    Log-Info "Waiting for the data to be fully prepared and consistent in Microsoft Graph"
    Start-Sleep -Seconds 10
    $serviceUserId = Get-IdWithRetry "Cloud ${customer} Forms Service User ${env}" "user" "mail@${env}.internal.com"
    
    $groupNames = @(
    "${env}-internal",
    "tenant-${customer}",
    "internal-auditing-record-write",
    "internal-storage-command-write",
    "internal-storage-delivery-read",
    "internal-storage-delivery-write",
    "internal-storage-file-read",
    "internal-storage-file-write",
    "internal-storage-file-delete",
    "internal-storage-temp-read",
    "internal-storage-temp-write",
    "internal-storage-temp-delete"
    )

    foreach ($groupName in $groupNames) {
        $groupId = Get-IdWithRetry $groupName "group"
        Add-GroupMember $groupId $groupName $serviceUserId "Cloud ${customer} Forms Service User ${env}"
    }
}

function Export-OnePasswordTemplate ($customer, $password, $env, $mailDomain) {
    Log-Action "Preparing 1Password template for the tenant service user"
    Get-Confirmation

    $login = [PSCustomObject]@{
        title    = "Azure AD - Cloud ${customer} Formshub Service User ${env}"
        username = "mail@${env}.internal.com"
        password = $password
        url      = "${env}.${customer}.internal.com"
        tags     = "azure_ad;cloud_app;${customer};${env}"
    }

    $login |
    Select-Object title, username, password, url, tags |
    ConvertTo-Csv -NoTypeInformation |
    Select-Object -Skip 1 |
    Set-Content -Path ".\1password-import.csv" -Encoding UTF8
}

#main 
$inputEnv = $false
$inputCustomer = $false
$inputPassword = $false
$confirmation = $false
$mailDomain = ""

do {
    $env = Read-Host "Enter the environment (dev, test, preprod, prod)"
    if ($env -eq "dev" -or $env -eq "test" -or $env -eq "preprod") {
        $inputEnv = $true
        Log-Success "Environment $env was found"
        $mailDomain = "${env}.internal.com"
    }
    elseif ($env -eq "prod") {
        Log-Warn "The environment prod was selected"
        $mailDomain = "internal.com"
        $confirmation = $true
        $inputEnv = $true
    }
    else {
        Log-Error "Invalid environment entered"
    }
} while (!$inputEnv)

do {
    $customer = Read-Host "Enter the customer name" 
    if (![string]::IsNullOrWhiteSpace($customer)) {
        $inputCustomer = $true
        Log-Success "Customer $customer successfully set"
    }
    else {
        Log-Error "Customer input is empty"
    }
} while (!$inputCustomer)

do {
    $password = Read-Host "Enter a password"

    if (
        $password.Length -lt 24 -or
        $password -notmatch '[A-Z]' -or
        $password -notmatch '[a-z]' -or
        $password -notmatch '\d' -or
        $password -notmatch '[^a-zA-Z0-9]'
    ) {
        Log-Error "Invalid Password! Must be atleast 24 characters and contain a mix of uppercase, lowercase, digits and special characters." 
    } else {
        $inputPassword = $true
    }
    
} while (!$inputPassword)

$tenantId = "*************************************"
$clientId = "*************************************"
$clientSecret = "*************************************"

$secureClientSecret = ConvertTo-SecureString $clientSecret -AsPlainText -Force
$clientSecretCredential = New-Object System.Management.Automation.PSCredential($clientId, $secureClientSecret)

Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $clientSecretCredential -NoWelcome

try {
    New-TenantSecurityGroup $customer
    Add-TenantGroupMembers $customer $mailDomain
    New-TenantServiceUser $customer $password $env $mailDomain
    Set-PasswordPolicy $customer $env $mailDomain
    Add-ServiceUserGroups $customer $env $mailDomain
    Export-OnePasswordTemplate $customer $password $env $mailDomain
    Log-Success "Script was executed successfully"
    Log-Info "Disconnecting from Microsoft Graph"
    Disconnect-MgGraph
}
catch {
    Log-Error $_
    Log-Error "Script execution failed"
    Log-Info "Disconnecting from Microsoft Graph"
    Disconnect-MgGraph
    exit 1
}