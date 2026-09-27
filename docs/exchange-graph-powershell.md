# Exchange Online + Microsoft Graph PowerShell reference

Scope: how to sign in non-interactively (fast, unattended), and the cmdlets that
matter for mail flow rules (transport rules) and connectors, plus the Graph
equivalents. Everything here is from the current Microsoft Learn docs; sources
are linked at the bottom.

Modules used:

| Module | Version installed here | What it gives you |
| --- | --- | --- |
| `ExchangeOnlineManagement` | 3.10.1 | `Connect-ExchangeOnline`, `Connect-IPPSSession`, `Get-EXO*` REST cmdlets, and (after connect) the full server-side Exchange cmdlet surface: `*-TransportRule`, `*-InboundConnector`, `*-OutboundConnector`, ... |
| `Microsoft.Graph.Authentication` (+ workload submodules) | 2.40.0 | `Connect-MgGraph`, `Invoke-MgGraphRequest`, and the generated `*-Mg*` cmdlets |

Install:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module Microsoft.Graph -Scope CurrentUser -Force   # or just Microsoft.Graph.Authentication
```

Before you connect, `Get-Command -Module ExchangeOnlineManagement` shows only ~37
cmdlets. The mail-flow cmdlets (`New-TransportRule`, `Set-OutboundConnector`, ...)
are *proxy* cmdlets injected into the session by the service when
`Connect-ExchangeOnline` succeeds — that is why `Get-Help New-TransportRule`
fails until you are connected (use `-LoadCmdletHelp` on connect if you want
local help files).

---

## 1. Non-interactive sign-in (the "speed" part)

### 1.1 Decision table

| Where the script runs | Exchange Online | Microsoft Graph |
| --- | --- | --- |
| Azure VM / Function / Automation with managed identity | `Connect-ExchangeOnline -ManagedIdentity -Organization contoso.onmicrosoft.com` | `Connect-MgGraph -Identity` |
| Any server/container you control | app-only certificate: `Connect-ExchangeOnline -AppId .. -Organization .. -CertificateThumbprint`/`-Certificate`/`-CertificateFilePath` | `Connect-MgGraph -ClientId .. -TenantId .. -CertificateThumbprint ..` |
| CI (GitHub Actions, ADO) | app-only certificate, or an access token you fetched yourself | federated identity credential → token → `Connect-MgGraph -AccessToken` |
| Quick admin work at a keyboard | `Connect-ExchangeOnline -UserPrincipalName you@contoso.com` (browser) or `-Device` | `Connect-MgGraph -Scopes ...` or `-UseDeviceAuthentication` |
| Username + password (ROPC) | **not usable in practice** — see 1.4 | possible but discouraged — see 1.4 |

The fastest reliable pattern for unattended work is **app-only certificate auth**
(CBA). It needs no user, no MFA prompt, no password in a vault, and it is the
only officially supported unattended method for Exchange Online besides managed
identity.

### 1.2 App-only certificate auth for Exchange Online

One-time setup:

1. Register a Microsoft Entra application (single tenant).
2. API permissions → **Office 365 Exchange Online** → *Application permissions* →
   `Exchange.ManageAsApp` → grant admin consent.
3. Assign the app an Entra **directory role** that carries the Exchange RBAC you
   need — `Exchange Administrator` (full mail flow control) or
   `Global Reader` / `Exchange Recipient Administrator` for read-only work. The
   role in the token is what builds the session's RBAC; without a role the
   connect succeeds but every cmdlet is access-denied.
4. Upload the public key (`.cer`) of a self-signed or CA certificate to the app.

Connect:

```powershell
# Windows, certificate in the local store
Connect-ExchangeOnline -AppId <app-guid> -Organization contoso.onmicrosoft.com `
    -CertificateThumbprint "012THISISADEMOTHUMBPRINT" -ShowBanner:$false

# Any OS, certificate object in memory (preferred on Linux — -CertificateThumbprint is Windows-only)
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
            '/secure/automation-cert.pfx', $pfxPassword)
Connect-ExchangeOnline -AppId <app-guid> -Organization contoso.onmicrosoft.com `
    -Certificate $cert -ShowBanner:$false

# Certificate file on disk (needs the password; fine if the password comes from a vault)
Connect-ExchangeOnline -AppId <app-guid> -Organization contoso.onmicrosoft.com `
    -CertificateFilePath /secure/automation-cert.pfx -CertificatePassword $pfxPassword
```

Notes that bite people:

* `-CertificateThumbprint` is **Windows only**. On Linux/macOS use `-Certificate`
  or `-CertificateFilePath`.
* `-Organization` must be the tenant's `.onmicrosoft.com` domain, not a vanity
  domain, and not the tenant GUID.
* App-only does **not** work for `New-UnifiedGroup`, `Remove-UnifiedGroup`,
  `Add-UnifiedGroupLinks`, `Remove-UnifiedGroupLinks` (use Graph groups instead),
  and is unsupported for Purview eDiscovery cmdlets in
  Security & Compliance PowerShell.
* Sovereign clouds: add `-ExchangeEnvironmentName O365USGovGCCHigh` /
  `O365USGovDoD` / `O365China`.
* Delegated/partner work: `-DelegatedOrganization customer.onmicrosoft.com`
  (GDAP + app consent is the supported route).

### 1.3 Managed identity (Azure-hosted automation)

```powershell
Connect-ExchangeOnline -ManagedIdentity -Organization contoso.onmicrosoft.com
Connect-ExchangeOnline -ManagedIdentity -Organization contoso.onmicrosoft.com -ManagedIdentityAccountId <user-assigned-mi-client-id>
Connect-MgGraph -Identity
Connect-MgGraph -Identity -ClientId <user-assigned-mi-client-id>
```

The managed identity's service principal still needs `Exchange.ManageAsApp` +
an Entra role for Exchange, and Graph application permissions for Graph.

### 1.4 ROPC (username/password grant) — what it actually does and why it's a trap

ROPC = OAuth 2.0 *Resource Owner Password Credentials*: a single POST that
exchanges a username and password for tokens, with no browser.

```http
POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

client_id=<public-client-app-guid>
&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default
&username=admin@contoso.com
&password=<password>
&grant_type=password
```

Then `Connect-MgGraph -AccessToken (ConvertTo-SecureString $token -AsPlainText -Force)`.

Microsoft's documented constraints (all of these are hard blockers, not advice):

* Incompatible with MFA — if the account is required to do MFA, the request is
  **blocked**, not downgraded. Microsoft explicitly recommends not using ROPC.
* Work/school accounts in an Entra tenant only; personal accounts (including
  personal accounts invited as guests) can't use it. Tenant-specific or
  `organizations` endpoint only, never `common`.
* Accounts without a password (FIDO, passwordless Authenticator, SMS sign-in)
  can't use it.
* Not supported for federated identity (AD FS) unless Home Realm Discovery has
  `AllowCloudPasswordValidation = true` with password hash sync;
  pass-through authentication does work.
* Passwords with leading/trailing whitespace fail.
* Conditional Access that requires MFA or a compliant device will block it, and
  resources are increasingly rejecting non-MFA tokens outright.

For **Exchange Online specifically**: the V3 module does not take a ROPC token.
`Connect-ExchangeOnline -Credential $cred` (basic-style username/password) only
works for non-MFA accounts and is on the way out; modern tenants with security
defaults or CA will fail. Do not build automation on it.

**Use instead, in order of preference:** managed identity → certificate app-only
(CBA) → federated identity credential in CI → device code for a human.
`-Device` / `-InlineCredential` exist for interactive-but-headless cases:

```powershell
Connect-ExchangeOnline -Device                      # device code, no local browser
Connect-ExchangeOnline -UserPrincipalName you@contoso.com -InlineCredential
Connect-MgGraph -UseDeviceAuthentication
```

### 1.5 Speed tips for large jobs

* Use the REST-backed `Get-EXO*` cmdlets instead of the classic ones for bulk
  reads — `Get-EXOMailbox`, `Get-EXORecipient`, `Get-EXOMailboxStatistics`,
  `Get-EXOMailboxPermission`, `Get-EXOCasMailbox`,
  `Get-EXOMailboxFolderPermission`, `Get-EXOMailboxFolderStatistics`,
  `Get-EXORecipientPermission`, `Get-EXOMobileDeviceStatistics`. They support
  `-PropertySets` / `-Properties` so you only pull the fields you need.
* `Connect-ExchangeOnline -ShowBanner:$false -SkipLoadingFormatData` shaves
  connect time; `-CommandName Get-TransportRule,Set-TransportRule,...` loads only
  the proxies you use, which is a big win on cold start.
* `Get-ConnectionInformation` tells you what sessions exist;
  `Disconnect-ExchangeOnline -Confirm:$false` at the end of every script.
* Graph: batch with `Invoke-MgGraphRequest` against `/$batch`, and prefer
  `-All` + `$select` over client-side filtering.

---

## 2. Mail flow rules (transport rules)

Cmdlets: `Get-TransportRule`, `New-TransportRule`, `Set-TransportRule`,
`Remove-TransportRule`, `Enable-TransportRule`, `Disable-TransportRule`,
`Get-TransportRuleAction`, `Get-TransportRulePredicate`,
`Export-TransportRuleCollection`, `Import-TransportRuleCollection`.

Mental model: a rule is **priority + conditions + exceptions + actions**.

* Every condition parameter has a mirrored exception parameter prefixed
  `ExceptIf` — `SubjectContainsWords` / `ExceptIfSubjectContainsWords`,
  `SenderDomainIs` / `ExceptIfSenderDomainIs`, and so on. That is why the syntax
  block is hundreds of parameters long: it is ~the same list twice.
* `Get-TransportRulePredicate` and `Get-TransportRuleAction` enumerate the legal
  conditions/actions in *your* tenant — run them instead of guessing parameter
  names.
* A rule with **no** conditions or exceptions applies to every message in the
  org. With a destructive action (`DeleteMessage`, `RejectMessageReasonText`)
  that is an org-wide outage. Always create with `-Enabled $false` first, or
  scope it hard.
* Rules using directory attributes (`SenderADAttributeContainsWords`,
  `ExceptIfFromMemberOf`, ...) only evaluate for senders/recipients that exist in
  the directory.
* Priority: `-Priority 0` is evaluated first. `Set-TransportRule -Priority`
  renumbers the rest.
* `-Mode Enforce | Audit | AuditAndNotify` — use `Audit` to dry-run a rule in
  production before enforcing.
* Matching happens after MIME decoding; you cannot match raw Base64.

Common shapes:

```powershell
# Disabled-by-default rule: prepend a banner to mail from outside
New-TransportRule -Name "External sender banner" -Enabled $false `
    -FromScope NotInOrganization -SentToScope InOrganization `
    -ApplyHtmlDisclaimerLocation Prepend `
    -ApplyHtmlDisclaimerText '<div style="background:#ffd">External sender.</div>' `
    -ApplyHtmlDisclaimerFallbackAction Wrap

# Audit mode first, then enforce
Set-TransportRule "External sender banner" -Mode Audit
Enable-TransportRule "External sender banner"

# Route a domain out through a connector (see §3)
New-TransportRule -Name "Partner via smart host" `
    -RecipientDomainIs partner.example -RouteMessageOutboundConnector "Partner Outbound"

# Inspect / back up / restore
Get-TransportRule | Format-Table Name,State,Mode,Priority,Description -Auto
Get-TransportRule "External sender banner" | Format-List
$file = Export-TransportRuleCollection
Set-Content -Path rules.xml -Value $file.FileData -Encoding Byte
Import-TransportRuleCollection -FileData ([IO.File]::ReadAllBytes('rules.xml'))
```

`Export-TransportRuleCollection` / `Import-TransportRuleCollection` is the
supported backup/restore path — take an export before touching rules.

Graph equivalent: **there isn't one.** Transport rules are not exposed in
Microsoft Graph today; Exchange Online PowerShell is the API.

---

## 3. Connectors

Cmdlets: `Get-InboundConnector`, `New-InboundConnector`, `Set-InboundConnector`,
`Remove-InboundConnector`, `Get-OutboundConnector`, `New-OutboundConnector`,
`Set-OutboundConnector`, `Remove-OutboundConnector`,
`Validate-OutboundConnector`, plus `Get-AcceptedDomain`,
`Get-RemoteDomain`, `Get-HostedConnectionFilterPolicy` for the surrounding
config.

### 3.1 Inbound connector — mail *into* the tenant

Key parameters (`New-InboundConnector`):

| Parameter | Meaning |
| --- | --- |
| `-SenderDomains` (required) | domains this connector applies to, `*.contoso.com` allowed |
| `-SenderIPAddresses` | source IPs/CIDRs; requires `ConnectorType OnPremises` semantics |
| `-RestrictDomainsToIPAddresses` | reject mail from `SenderDomains` that arrives from any other IP |
| `-RestrictDomainsToCertificate` | require the TLS cert name in `TlsSenderCertificateName` |
| `-RequireTls` | force TLS on inbound |
| `-TlsSenderCertificateName` | certificate subject/SAN the partner must present |
| `-ConnectorType` | `Partner` or `OnPremises` |
| `-CloudServicesMailEnabled` | `$true` only for hybrid/on-prem — preserves internal headers and treats mail as internal |
| `-TreatMessagesAsInternal` | mark as internal (hybrid split-domain) |
| `-AssociatedAcceptedDomains` | limit the connector to specific accepted domains |
| `-EFSkipIPs`, `-EFSkipLastIP`, `-EFSkipMailGateway`, `-EFUsers`, `-EFTestMode` | Enhanced Filtering ("skip listing") so EOP looks past your third-party gateway to the real source IP |
| `-ScanAndDropRecipients` | recipients whose mail is scanned then dropped |
| `-Enabled` | on/off |

```powershell
New-InboundConnector -Name "Contoso Inbound" -SenderDomains *.contoso.com `
    -SenderIPAddresses 192.168.0.1/25 -RestrictDomainsToIPAddresses $true

# Enhanced Filtering when a third-party gateway sits in front of EOP
Set-InboundConnector "Gateway Inbound" -EFSkipLastIP $true -EFUsers all -EFTestMode $false
```

### 3.2 Outbound connector — mail *out of* the tenant

Key parameters (`New-OutboundConnector`):

| Parameter | Meaning |
| --- | --- |
| `-RecipientDomains` | domains routed through this connector |
| `-AllAcceptedDomains` | hybrid: route mail for all accepted domains |
| `-SmartHosts` | FQDN/IP of the next hop (mutually exclusive with `-UseMXRecord $true`) |
| `-UseMXRecord` | `$true` = normal MX routing, `$false` = use `SmartHosts` |
| `-TlsSettings` | `EncryptionOnly`, `CertificateValidation`, `DomainValidation` |
| `-TlsDomain` | the domain to validate when `TlsSettings DomainValidation` |
| `-SmtpDaneMode`, `-MtaStsMode` | DANE / MTA-STS enforcement |
| `-IsTransportRuleScoped` | connector only used when a transport rule names it (`RouteMessageOutboundConnector`) |
| `-RouteAllMessagesViaOnPremises` | hybrid: send everything to on-prem first |
| `-SenderRewritingEnabled` | SRS rewriting on relay |
| `-CloudServicesMailEnabled` | hybrid header preservation |
| `-TestMode`, `-Enabled`, `-ConnectorType`, `-Comment` | validation/lifecycle |

```powershell
New-OutboundConnector -Name "Partner Outbound" -RecipientDomains *.contoso.com `
    -TlsSettings DomainValidation -TlsDomain *.contoso.com

# Smart-host relay used only by a transport rule
New-OutboundConnector -Name "Archive Relay" -UseMXRecord $false `
    -SmartHosts relay.example.net -IsTransportRuleScoped $true -Enabled $true

Validate-OutboundConnector -Identity "Partner Outbound" -Recipients user@contoso.com
```

Operational notes:

* Inbound connector creation is **blocked in E5 developer subscriptions** —
  support can't lift it.
* Connectors are the classic misconfiguration path for mail relay abuse: an
  inbound connector with wide `SenderDomains` and no IP/cert restriction, or an
  outbound connector that relays for a domain you don't own. Always pair
  `SenderDomains` with `SenderIPAddresses` + `RestrictDomainsToIPAddresses $true`
  or a certificate restriction, and review
  `Get-InboundConnector | ? {$_.ConnectorSource -ne 'HybridWizard'}` periodically.
* `ConnectorSource` tells you who made it: `Default`, `HybridWizard`,
  `Migrated`, `AdminUI`.

---

## 4. Graph side: what to use it for

Graph does **not** cover transport rules or connectors. Use it for the identity
and mailbox-object work around them:

```powershell
Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $thumb -NoWelcome

Get-MgUser -All -Property id,userPrincipalName,accountEnabled
Get-MgGroup -Filter "mailEnabled eq true"
Get-MgUserMailboxSetting -UserId user@contoso.com
Get-MgUserMailFolderMessageRule -UserId user@contoso.com -MailFolderId Inbox   # per-user inbox rules
Invoke-MgGraphRequest -Method GET -Uri 'v1.0/users?$select=id,mail&$top=999'
Get-MgContext; Disconnect-MgGraph
```

Per-user **inbox rules** are a different thing from org **transport rules**:
Graph `messageRules` (or `Get-InboxRule`/`New-InboundRule` in EXO) is per-mailbox;
`Get-TransportRule` is org-wide. Malicious-forwarding investigations need both.

Useful application permissions: `Mail.Read`, `MailboxSettings.Read`,
`User.Read.All`, `Group.Read.All`, `Directory.Read.All`,
`Exchange.ManageAsApp` (Office 365 Exchange Online, not Graph).

---

## 5. Repeatable session skeleton

```powershell
#requires -Modules ExchangeOnlineManagement
$ErrorActionPreference = 'Stop'
try {
    Connect-ExchangeOnline -AppId $env:EXO_APP_ID -Organization $env:EXO_ORG `
        -Certificate $cert -ShowBanner:$false `
        -CommandName Get-TransportRule,Set-TransportRule,New-TransportRule,
                     Get-InboundConnector,Get-OutboundConnector,Set-OutboundConnector
    # ... work, with -WhatIf on every mutating call until reviewed ...
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
```

Rules of thumb before any change: `Export-TransportRuleCollection` for a backup,
`-WhatIf` on the real command, create disabled or in `Audit` mode, then enable.

---

## Sources

* App-only auth for EXO: <https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2>
* Connect to Exchange Online PowerShell: <https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell>
* Managed identity for EXO: <https://learn.microsoft.com/en-us/powershell/exchange/connect-exo-powershell-managed-identity>
* `New-TransportRule`: <https://learn.microsoft.com/en-us/powershell/module/exchange/new-transportrule>
* `New-InboundConnector`: <https://learn.microsoft.com/en-us/powershell/module/exchange/new-inboundconnector>
* `New-OutboundConnector`: <https://learn.microsoft.com/en-us/powershell/module/exchange/new-outboundconnector>
* Graph app-only access: <https://learn.microsoft.com/en-us/graph/auth-v2-service>
* Graph PowerShell auth: <https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands>
* ROPC grant and how to migrate off it: <https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth-ropc>
