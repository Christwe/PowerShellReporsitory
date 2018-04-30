$TenantAdminUrl = "https://christwe-admin.sharepoint.com"
$ReplicatedBaseTemplate = @("851","101")     #https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listtemplatetype.aspx, https://www.codeproject.com/articles/765404/list-of-sharepoint-lists-basetemplatetype

$Creds = Get-Credential
Connect-SPOService -url $TenantAdminUrl -Credential $Creds

Get-SPOSite | where {$_.Template -like "*PUBLISHING*"} | ForEach-Object {
    $_.url
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($_.Url) 
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName, $Creds.Password)

    $web = $context.Web
    $lists = $web.Lists
    $context.Load($web)
    $context.Load($lists)
    $context.ExecuteQuery()

    #Must filter on Basetype = DocumentLibrary to insure we don't try to sync a list by accident
    $Lists | where {$_.BaseType -eq "DocumentLibrary"} | ForEach-Object {
        if($ReplicatedBaseTemplate -contains $_.BaseTemplate)
        {
            $_.Title
            $ListFolder = $_.RootFolder
            $context.Load($ListFolder)
            $context.ExecuteQuery()
            $LibUrl = $ListFolder.ServerRelativeUrl
            Add-SPOTenantCdnOrigin -CdnType Private -OriginUrl $liburl -Confirm:$false
        }
    }
}
Disconnect-SPOService