# Invoke-Unarchive
Exchange script to move contents back from personal archives to primary mailboxes

Blog: https://eightwone.com/2021/08/14/unarchiving-mailbox-items/

Notes:
- Requires Exchange Server 2013 SP1 (or later) or Exchange Online
- Microsoft Exchange Web Services (EWS) Managed API 2.21 or up is required. This is available as NuGet package, e.g.
  Install-Package exchange.webservices.managed.api
  See https://eightwone.com/2020/10/05/ews-webservices-managed-api on how to add NuGet as package provider.
- OAuth requires MSAL library (Microsoft.Identity.Client.dll). This is also available as a NuGet package, e.g.
  Install-Package Microsoft.Identity.Client -ProviderName NuGet
- Search order for DLL's is script Folder then installed packages.
- Script has not been tested with mixed locality, e.g. primary mailboxes on-premises with Exchange Online archives.

Example:
    .\Invoke-Unarchive.ps1 -Identity john@contoso.com -Server outlook.office365.com -Impersonation -Secret $MySecret -TenantId $TenantId -ClientId $ClientId -Verbose

Invokes unarchiving of contents from personal archive to primary mailbox for john@contoso.com. Authentication is OAuth using provided Tenant ID, Client ID and secret.
