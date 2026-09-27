# reamd_me
classic

## Microsoft 365 PowerShell

- [Exchange Online + Microsoft Graph PowerShell reference](docs/exchange-graph-powershell.md) —
  non-interactive sign-in (app-only certificate, managed identity, why ROPC is a
  dead end), transport rules, and connectors.
- `scripts/Connect-M365.ps1` — app-only / managed-identity connect helper.
- `scripts/Export-MailFlowConfig.ps1` — read-only snapshot of rules, connectors
  and domains; run before changing mail flow.
