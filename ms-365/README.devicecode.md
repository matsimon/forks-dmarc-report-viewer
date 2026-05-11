# Access MS 365 Mail Access with DavMail IMAP Proxy using DeviceCode authentication

This is another Docker-Compose example that shows how to set up DavMail and the 
DMARC-Report-Viewer together so it possible to read reports from an MS mail account.

In this scenario we are accessing a shared mailbox using delegated permissions using
a DeviceCode authentication. This is unlikely useful for a permanent deployment but
rather to get some data loaded in one-off cases.

## Instructions
1. You have Docker and Docker-Compose installed
2. Your MS account admin has given his consent
3. You have an app registered (single tenant) in Entra ID
4. In the app registration set this Redirect URI configuration:
   - Manage > Authentication > Redirect URI configuration > + Add Redirect URI
   - Select **Mobile and desktop application**
   - Enable this predefined URI `https://login.microsoftonline.com/common/oauth2/nativeclient`
6. In the app registration, enable public client flows:
   - Manage > Authentication > Settings
   - Enable `Allow public client flows`
8. You have set up **delegated** Graph API permissions for (Manage > API Permissions):
    - `Mail.ReadWrite`       - If the DMARC mailbox is a regular mailbox and directly accessed
    - `Mail.ReadWrite.Shared` - If the DMARC mailbox is a shared mailbox accessed only by delegated permissions
    - `offline_access`
    - `openid`
    - `profile`
9. Copy [docker-compose.devicecode.yml](docker-compose.devicecode.yml) and [davmail.properties](davmail.devicecode.properties)
10. Replace the placeholder values in the `davmail.devicecode.properties` file, mostly:
    - davmail.oauth.tenantId: Tenant ID of the Entra tenant
    - davmail.oauth.clientId: Application ID of your app registration
12. Replace the placeholder values in the `docker-compose.devicecode.yml` file, mostly:
    - IMAP_USER
      - Regular mailboxes: Just enter primary mail address
      - Shared mailboxes with delegated user permissions: `delegateduser@example.net/sharedmailbox@example.net`
    - HTTP_SERVER_USER
    - HTTP_SERVER_PASSWORD
14. Start the containers by running the command `docker compose --file docker-compose.devicdecode.yml`
15. Navigate to https://login.microsoft.com/device and enter the code shown in the output
16. Sign in as the user that has delegated permissions

## Possible way to create the app registration

In order to set up the app registration in a more reproducible way, here is some code that should help you set up the app registration using Powershell.
(Very much work in progress!)

TODO: Redirect URIs are not yet defined.

```powershell

Connect-MgGraph -Scopes Application.ReadWrite.All,AppRoleAssignment.ReadWrite.All

# Set a display name and register the application
$AppDisplayName = "DMARC Report Viewer"

$app = New-MgApplication -DisplayName $AppDisplayName -SignInAudience AzureADMyOrg

# Enable public client flows
Update-MgApplication `
  -ApplicationId $app.Id `
  -RequiredResourceAccess @(
    @{
      ResourceAppId  = $graphSp.AppId
      ResourceAccess = $graphResourceAccess
    },
    @{
      ResourceAppId  = $exoSp.AppId
      ResourceAccess = $exoResourceAccess
    }
  )

# Grant the delegated API permissions
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

$graphScopes = @(
  "openid",
  "profile",
  "offline_access"
  "Mail.ReadWrite",
  "Mail.ReadWrite.Shared"
)

$graphResourceAccess = $graphScopes | ForEach-Object {
  $scope = $graphSp.Oauth2PermissionScopes | Where-Object Value -eq $_
  @{
    Id   = $scope.Id
    Type = "Scope"
  }
}

```

