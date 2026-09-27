<#
.SYNOPSIS
    Read-only snapshot of Exchange Online mail flow config: transport rules,
    inbound/outbound connectors, accepted and remote domains.

.DESCRIPTION
    Run this before changing anything. Produces JSON per object type plus the
    native transport-rule collection export (the supported restore format) and a
    short console summary highlighting connectors that look unrestricted.

    Requires an existing Exchange Online session (see Connect-M365.ps1).

.EXAMPLE
    ./Export-MailFlowConfig.ps1 -OutputDirectory ./mailflow-backup
#>
[CmdletBinding()]
param(
    [string] $OutputDirectory = (Join-Path (Get-Location) ("mailflow-{0:yyyyMMdd-HHmmss}" -f (Get-Date)))
)

$ErrorActionPreference = 'Stop'

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    throw 'No Exchange Online session. Run Connect-M365.ps1 first.'
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$sets = [ordered]@{
    'transport-rules'     = { Get-TransportRule }
    'inbound-connectors'  = { Get-InboundConnector }
    'outbound-connectors' = { Get-OutboundConnector }
    'accepted-domains'    = { Get-AcceptedDomain }
    'remote-domains'      = { Get-RemoteDomain }
}

foreach ($name in $sets.Keys) {
    $data = & $sets[$name]
    $path = Join-Path $OutputDirectory "$name.json"
    $data | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding utf8
    Write-Output ("{0,-22} {1,4} object(s) -> {2}" -f $name, @($data).Count, $path)
}

# Native, restorable transport rule export
$collection = Export-TransportRuleCollection
[IO.File]::WriteAllBytes((Join-Path $OutputDirectory 'transport-rules.xml'), $collection.FileData)

Write-Output "`nRules by state:"
Get-TransportRule | Group-Object State | ForEach-Object { "  {0}: {1}" -f $_.Name, $_.Count }

Write-Output "`nInbound connectors without IP or certificate restriction:"
Get-InboundConnector |
    Where-Object { -not $_.RestrictDomainsToIPAddresses -and -not $_.RestrictDomainsToCertificate } |
    Select-Object Name, Enabled, ConnectorType, ConnectorSource, @{n='SenderDomains';e={$_.SenderDomains -join ','}} |
    Format-Table -AutoSize

Write-Output "Outbound connectors relaying to smart hosts:"
Get-OutboundConnector |
    Where-Object { $_.SmartHosts } |
    Select-Object Name, Enabled, IsTransportRuleScoped, TlsSettings, @{n='SmartHosts';e={$_.SmartHosts -join ','}} |
    Format-Table -AutoSize

Write-Output "Snapshot written to $OutputDirectory"
