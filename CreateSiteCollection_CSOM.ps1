Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 

$siteUrl = “https://mytenant.sharepoint.com/sites/mysitecollection”
$username = "admin@mytenant.onmicrosoft.com"
$password = Read-Host -Prompt "Enter password" -AsSecureString 
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
$ctx.Credentials = $credentials

$webCreationInformation = New-Object Microsoft.SharePoint.Client.WebCreationInformation
$webCreationInformation.Url = $SiteUrl      #"MyTeamSite"
$webCreationInformation.Title = $SiteTitle     #"My Team Site"
$webCreationInformation.WebTemplate = $SiteTemplate      #"STS#0"
$newWeb = $ctx.Web.Webs.Add($webCreationInformation)

$ctx.Load($newWeb) 
$ctx.ExecuteQuery()
Write-Host "Title" $newWeb.Title
