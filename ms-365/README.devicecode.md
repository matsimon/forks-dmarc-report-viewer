# Access MS 365 Mail Access with DavMail IMAP Proxy using DeviceCode authentication

This is another Docker-Compose example that shows how to set up DavMail and the 
DMARC-Report-Viewer together so it possible to read reports from an MS mail account.

In this scenario we are accessing a shared mailbox using delegated permissions using
a DeviceCode authentication. This is unlikely useful for a permanent deployment but
rather to get some data loaded in one-off cases.

## Instructions

1. You have Docker and Docker-Compose installed
2. Your MS account admin has given his consent
3. Use the Powershell code below to register and app in Entra ID, then
4. Set the redirect URI configuration as required
   * Enable `Allow public client flows`
   * Grant delegated Graph API permissions for the following scopes
     - `Mail.ReadWrite` - Required, if the DMARC mailbox is a regular mailbox and directly accessed
     - `Mail.ReadWrite.Shared` - Required, if the DMARC mailbox is a shared mailbox accessed by delegated permissions
     - `offline_access`
     - `openid`
     - `profile`
   * Create a service principal for the app registration
   * Restrict the use to assigned users and assign a user (or group) so that they can use it
9. Copy [docker-compose.devicecode.yml](docker-compose.devicecode.yml) and [davmail.devicecode.properties](davmail.devicecode.properties)
10. Replace the placeholder values in the `davmail.devicecode.properties` file, mostly:
    - `davmail.oauth.tenantId`: Tenant ID of the Entra tenant
    - `davmail.oauth.clientId`: Application ID of your app registration
11. Replace the placeholder values in the `docker-compose.devicecode.yml` file, mostly:
    - `IMAP_USER`, depending on your use case:
      - Regular mailboxes: Just enter the primary mail address
      - Shared mailboxes with delegated user permissions: `delegateduser@example.net/sharedmailbox@example.net`
    - `HTTP_SERVER_USER`
    - `HTTP_SERVER_PASSWORD`
12. Start the containers by running the command `docker compose --file docker-compose.devicdecode.yml`
13. Navigate to `https://login.microsoft.com/device` and enter the code shown in the output
14. Sign in as the user that has delegated permissions

## Tips

**Q:** Are DMARC messages sorted into a subfolder?
**A:** Add IMAP_FOLDER: "YOURFOLDER" to `docker-compose.devicecode.yml`

## Possible way to create the app registration

In order to set up the app registration in a more reproducible way, here is some code that should help you set 
up the app registration using Powershell.

```powershell
# Connect to Graph as admin
Connect-MgGraph -Scopes Application.ReadWrite.All,AppRoleAssignment.ReadWrite.All

# Set a display name and register the application
$AppDisplayName = "DMARC Report Viewer"

$app = New-MgApplication -DisplayName $AppDisplayName -SignInAudience AzureADMyOrg

# Enable public client flows
Update-MgApplication -ApplicationId $app.Id -IsFallbackPublicClient

# Set the redirect URI
$redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

Update-MgApplication -ApplicationId $app.Id `
    -PublicClient @{
        RedirectUris = @($redirectUri)
    }

# Grant the delegated API permissions
$graphAppId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph

# Get Graph service principal
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"

# Find the permission IDs we need
$scopeNames = @("Mail.ReadWrite", "Mail.ReadWrite.Shared", "User.Read", "openid", "profile", "offline_access")
$scopes = $graphSp.Oauth2PermissionScopes | Where-Object { $scopeNames -contains $_.Value }

# Build the required resource access
$requiredAccess = @{
    ResourceAppId = $graphAppId
    ResourceAccess = $scopes | ForEach-Object {
        @{
            Id = $_.Id
            Type = "Scope"
        }
    }
}

# Apply the scopes
Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @($requiredAccess)

# Grant admin consent
# Check if a service principal already exists for our app reg, if not, create it
$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'"
if (-not $sp) {
    $sp = New-MgServicePrincipal -AppId $app.AppId
}

# Check if grant already exists
$existingGrant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)' and resourceId eq '$($graphSp.Id)'"

if ($existingGrant) {
    # Update existing grant with all scopes
    Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $existingGrant.Id `
        -Scope ($scopeNames -join " ")
} else {
    # Create new grant with all scopes at once
    New-MgOauth2PermissionGrant `
        -ClientId $sp.Id `
        -ConsentType "AllPrincipals" `
        -ResourceId $graphSp.Id `
        -Scope ($scopeNames -join " ")
}

# Require explicit assignment
Update-MgServicePrincipal -ServicePrincipalId $sp.Id -AppRoleAssignmentRequired

# Assign a UPN that has delegated permissions using its UPN
$user = Get-MgUser -Filter "userPrincipalName eq 'youruser@example.onmicrosoft.com'"

New-MgServicePrincipalAppRoleAssignment `
    -ServicePrincipalId $sp.Id `
    -PrincipalId $user.Id `
    -ResourceId $sp.Id `
    -AppRoleId "00000000-0000-0000-0000-000000000000" # default role

# Show the App ID and the tenant ID which you need in our config files
Write-Host "App ID (davmail.oauth.clientId):" $app.AppId
Write-Host "Tenant ID (davmail.oauth.tenantId):" (Get-MgOrganization).Id
```

