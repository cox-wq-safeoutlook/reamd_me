<#
.SYNOPSIS
    Non-interactive (app-only) connect to Exchange Online and/or Microsoft Graph.

.DESCRIPTION
    Supports the two unattended methods Microsoft actually supports:
      - certificate app-only auth (works anywhere; use -CertificatePath on Linux)
      - managed identity (Azure-hosted only)
    Falls back to device-code for a human at a keyboard.

    ROPC (username + password) is deliberately not implemented: it is blocked by
    MFA/Conditional Access and is not supported by the Exchange Online V3 module.
    See docs/exchange-graph-powershell.md section 1.4.

.EXAMPLE
    ./Connect-M365.ps1 -AppId $env:EXO_APP_ID -Organization contoso.onmicrosoft.com `
        -TenantId $env:TENANT_ID -CertificatePath /secure/automation.pfx -Service Both

.EXAMPLE
    ./Connect-M365.ps1 -Organization contoso.onmicrosoft.com -ManagedIdentity -Service Exchange
#>
[CmdletBinding(DefaultParameterSetName = 'Certificate')]
param(
    [Parameter(ParameterSetName = 'Certificate', Mandatory)]
    [string] $AppId,

    [Parameter(ParameterSetName = 'Certificate')]
    [string] $TenantId,

    [Parameter(Mandatory)]
    [string] $Organization,

    [Parameter(ParameterSetName = 'Certificate')]
    [string] $CertificatePath,

    [Parameter(ParameterSetName = 'Certificate')]
    [securestring] $CertificatePassword,

    [Parameter(ParameterSetName = 'Certificate')]
    [string] $CertificateThumbprint,

    [Parameter(ParameterSetName = 'ManagedIdentity', Mandatory)]
    [switch] $ManagedIdentity,

    [Parameter(ParameterSetName = 'ManagedIdentity')]
    [string] $ManagedIdentityAccountId,

    [Parameter(ParameterSetName = 'DeviceCode', Mandatory)]
    [switch] $DeviceCode,

    [ValidateSet('Exchange', 'Graph', 'Both')]
    [string] $Service = 'Exchange',

    [string[]] $GraphScopes = @('User.Read.All'),

    [string[]] $CommandName
)

$ErrorActionPreference = 'Stop'

function Connect-Exchange {
    $common = @{ ShowBanner = $false }
    if ($CommandName) { $common['CommandName'] = $CommandName }

    switch ($PSCmdlet.ParameterSetName) {
        'Certificate' {
            $common += @{ AppId = $AppId; Organization = $Organization }
            if ($CertificateThumbprint) {
                if (-not $IsWindows) {
                    throw '-CertificateThumbprint is Windows-only; use -CertificatePath on this OS.'
                }
                $common['CertificateThumbprint'] = $CertificateThumbprint
            }
            elseif ($CertificatePath) {
                $common['Certificate'] = Get-CertificateObject
            }
            else { throw 'Provide -CertificatePath or -CertificateThumbprint.' }
        }
        'ManagedIdentity' {
            $common += @{ ManagedIdentity = $true; Organization = $Organization }
            if ($ManagedIdentityAccountId) { $common['ManagedIdentityAccountId'] = $ManagedIdentityAccountId }
        }
        'DeviceCode' { $common['Device'] = $true }
    }

    Connect-ExchangeOnline @common
    Get-ConnectionInformation | Select-Object Organization, UserPrincipalName, TokenStatus, State
}

function Connect-Graph {
    switch ($PSCmdlet.ParameterSetName) {
        'Certificate' {
            $tenant = if ($TenantId) { $TenantId } else { $Organization }
            $mg = @{ ClientId = $AppId; TenantId = $tenant; NoWelcome = $true }
            if ($CertificateThumbprint) { $mg['CertificateThumbprint'] = $CertificateThumbprint }
            else { $mg['Certificate'] = Get-CertificateObject }
            Connect-MgGraph @mg
        }
        'ManagedIdentity' {
            if ($ManagedIdentityAccountId) { Connect-MgGraph -Identity -ClientId $ManagedIdentityAccountId -NoWelcome }
            else { Connect-MgGraph -Identity -NoWelcome }
        }
        'DeviceCode' { Connect-MgGraph -Scopes $GraphScopes -UseDeviceAuthentication -NoWelcome }
    }
    Get-MgContext | Select-Object ClientId, TenantId, AuthType, Scopes
}

function Get-CertificateObject {
    if (-not (Test-Path $CertificatePath)) { throw "Certificate not found: $CertificatePath" }
    if ($CertificatePassword) {
        [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $CertificatePassword)
    }
    else {
        [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
    }
}

if ($Service -in 'Exchange', 'Both') {
    Import-Module ExchangeOnlineManagement
    Connect-Exchange
}
if ($Service -in 'Graph', 'Both') {
    Import-Module Microsoft.Graph.Authentication
    Connect-Graph
}
