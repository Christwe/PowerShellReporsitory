<#

 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
 
 This sample assumes that you are familiar with the programming language being demonstrated and the 
 tools used to create and debug procedures. Microsoft support professionals can help explain the 
 functionality of a particular procedure, but they will not modify these examples to provide added 
 functionality or construct procedures to meet your specific needs. if you have limited programming 
 experience, you may want to contact a Microsoft Certified Partner or the Microsoft fee-based consulting 
 line at (800) 936-5200. 

 #>

 # Install SharePoint online shell https://www.microsoft.com/en-us/download/details.aspx?id=35588
 
$InputFile = "C:\Users\christwe\Desktop\users.txt"
$tenant = "weaverinc"        # This is the Tenant Name. Value must be enclosed in double quotes, Example: "Contoso01"
$SetRoleDefinition = "Read"   #Choose from Full Control,Design,Edit,Contribute,Read,Limited Access,System.LimitedView,System.LimitedEdit,Create new subsites,View Only
$OutputPath = "$env:temp"
$OutputFile = "Output.txt"


$SPOAdminCredential = Get-Credential

$Date = Get-Date
Set-Content -Path "$OutputPath\$OutputFile" -Value $Date

import-module microsoft.online.sharepoint.powershell -DisableNameChecking

Connect-SPOService -Url "https://$tenant-admin.sharepoint.com" -Credential $SPOAdminCredential

$UserNameFile = Get-Content -Path $InputFile 

Function RemoveUserFromAcl
{
    Param(
        $Object,
        $LoginName,
        $UserAccount,
        $Context
    )

    $RoleAssignments = $Object.RoleAssignments
    $Context.Load($RoleAssignments)
    $Context.ExecuteQuery()
    ForEach ($Role in $RoleAssignments)
    {
        $Context.Load($Role.Member)
        $Context.Load($Role.RoleDefinitionBindings)
        $Context.ExecuteQuery()

        if($Role.Member.LoginName -like "*$LoginName")
        {
            $ra = $Object.RoleAssignments.GetByPrincipal($UserAccount)
            $ra.RoleDefinitionBindings.RemoveAll()
            $ra.Update()
            $UserAccount.Update()
            $Context.Load($ra)
            $Context.Load($UserAccount)
            $Context.ExecuteQuery()
            $Result = 2
            Break
        }
    }
    Return $Result
}

Function RemoveUserFromGroups
{
    Param(
        $Groups,
        $WebObject,
        $LoginName
    )
    $Groups | ForEach-Object {
        ForEach($Login in Get-SPOUser -Site $WebObject.Url -Group $_.Title)
        {
            if($Login.LoginName -eq $LoginName)
            {
                $Result = 1
                Remove-SPOUser -LoginName $Login.LoginName -Site $WebObject.Url -Group $_.Title
            }
        }
    }
    Return $Result
}

Function AddUsertoAcl
{
    Param(
        $Object,
        $Context,
        $NewRoleDefinition,
        $UserAccount
    )
    $SetAccess = $Object.RoleDefinitions.GetByName($NewRoleDefinition)
    $SetRole = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($Context)
    $SetRole.Add($SetAccess)
    $SetPermission = $Object.RoleAssignments.Add($UserAccount,$SetRole)
    $Context.Load($SetPermission)
    $Context.Load($Object)
    try {
        $Context.ExecuteQuery()
        Return 0
    }
    Catch{Return 1}
}

function Get-SPOWebs
{
    param(
        $Object,
        $Context
    )
    $Context.Load($Object.Webs)
    $Context.ExecuteQuery()
    $Object.Webs | ForEach-Object {
       Get-SPOWebs -Object $_ -Context $Context
       $_
  }
}

Function Get-ListItems   #https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.list.getitems.aspx
{
    Param(
        [Microsoft.SharePoint.Client.ClientContext]$Context, 
        $List
    )
    $ListItems = @()
    $Query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    Do
    {
        $Items = $List.GetItems($Query)
        $Context.Load($Items)
        $Context.ExecuteQuery()
        $Query.ListItemCollectionPosition = $Items.ListItemCollectionPosition
        ForEach($Item in $Items)
        {
            $ListItems += $Item
        }
    }
    While($Query.ListItemCollectionPosition -ne $null)
    Return $ListItems
}

#https://sharepoint.stackexchange.com/questions/126221/spo-retrieve-hasuniqueroleassignements-property-using-powershell
Function Invoke-LoadMethod      
{
    param(
        [Microsoft.SharePoint.Client.ClientObject]$Object = $(throw "Please provide a Client Object"),
        [string]$PropertyName
    ) 
   $ctx = $Object.Context
   $load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
   $type = $Object.GetType()
   $clientLoad = $load.MakeGenericMethod($type) 


   $Parameter = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
   $Expression = [System.Linq.Expressions.Expression]::Lambda(
            [System.Linq.Expressions.Expression]::Convert(
                [System.Linq.Expressions.Expression]::PropertyOrField($Parameter,$PropertyName),
                [System.Object]
            ),
            $($Parameter)
   )
   $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
   $ExpressionArray.SetValue($Expression, 0)
   $clientLoad.Invoke($ctx,@($Object,$ExpressionArray))
}

Function AddUsertoListItemAcl
{
    Param (
        $Context,
        $RoleDefinition,
        $UserAccount,
        $ListItem
    )
    $Role = $Context.web.RoleDefinitions.GetByName($RoleDefinition)
    $RoleDB = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($Context)
    $RoleDB.Add($Role)
          
    $UserPerms = $ListItem.RoleAssignments.Add($UserAccount,$RoleDB)
    $ListItem.Update()
    $Context.Load($ListItem)
    $Context.ExecuteQuery()
}

Get-SPOSite -Limit All | where{($_.Template -ne "SRCHCEN#0") -or ($_.Template -ne "SPSMSITEHOST#0") -or ($_.Template -ne "POINTPUBLISHINGHUB#0")} | ForEach-object {
    $UserRemoved = $null
    $SPOContext = New-Object Microsoft.SharePoint.Client.ClientContext($_.Url) 
    $SPOContext.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SPOAdminCredential.UserName, $SPOAdminCredential.Password)

    $Site = $SPOContext.Web
    $SPOContext.Load($Site)
    try{
        $SPOContext.ExecuteQuery()
        $SiteUrl = $Site.Url
        Add-Content -Path "$OutputPath\$OutputFile" -Value "Scanning Site Collection $SiteUrl"
    }
    Catch{Add-Content -Path "$OutputPath\$OutputFile" -Value "ERROR: Enumerating site Collection"}

    ForEach ($UserName in $UserNameFile)
    {
        $error.Clear()
        $User = $Site.EnsureUser($UserName)
        $SPOContext.Load($User)
        try{
            $SPOContext.ExecuteQuery()
            Add-Content -Path "$OutputPath\$OutputFile" -Value "`t Scanning for user: $UserName"
        }
        Catch{Add-Content -Path "$OutputPath\$OutputFile" -Value "`t ERROR: $UserName is invalid error with Login"}
        
        if(!$Error)
        {
            $SiteGroups = $Site.SiteGroups
            $SPOContext.Load($SiteGroups)
            Try{$SPOContext.ExecuteQuery()}
            Catch{Add-Content -Path "$OutputPath\$OutputFile" -Value "`t ERROR: Enumerating site groups"}

            if($Site.hasUniqueRoleAssignments)
            {
                $UserRemoved = $null
                $UserRemoved = RemoveUserFromGroups -Groups $SiteGroups -WebObject $Site -LoginName $UserName
                $UserRemoved = RemoveUserFromAcl -Object $Site -LoginName $UserName -UserAccount $User -Context $SPOContext
                if($UserRemoved){
                    AddUsertoAcl -Object $Site -Context $SPOContext -NewRoleDefinition $SetRoleDefinition -UserAccount $User
                    Add-Content -Path "$OutputPath\$OutputFile" -Value "`t Removed\Added user to ACL for Site Collection and Groups"
                }
            }

            $Lists = $Site.Lists
            $SPOContext.Load($Lists)
            Try{$SPOContext.ExecuteQuery()}
            Catch{Add-Content -Path "$OutputPath\$OutputFile" -Value "ERROR: Loading all Lists"}
            ForEach ($List in $Lists)
            {
                if(!($List.Hidden))
                {   
                    $ListTitle = $List.Title     
                    Add-Content -Path "$OutputPath\$OutputFile" -Value "`t Scanning list: $ListTitle"
                    Invoke-LoadMethod -Object $List -PropertyName "HasUniqueRoleAssignments"
                    $SPOContext.ExecuteQuery()
                    if($List.HasUniqueRoleAssignments)
                    {
                        $UserRemoved = $null
                        $UserRemoved = RemoveUserFromAcl -Object $List -LoginName $UserName -UserAccount $User -Context $SPOContext
                        if($UserRemoved)
                        {
                            AddUsertoAcl -Object $List -Context $SPOContext -NewRoleDefinition $SetRoleDefinition -UserAccount $User
                            Add-Content -Path "$OutputPath\$OutputFile" -Value "`t Removed\Added user to ACL for List "
                        }
                    }
                    
                    $ListItems = Get-ListItems -Context $SPOContext -List $List
                    ForEach ($Item in $ListItems)
                    {
                        Invoke-LoadMethod -Object $Item -PropertyName "HasUniqueRoleAssignments"
                        #Invoke-LoadMethod -Object $Item -PropertyName "RoleDefinitions"
                        $SPOContext.ExecuteQuery()
                        if($Item.hasUniqueRoleAssignments)
                        {
                            $UserRemoved = $null
                            $UserRemoved = RemoveUserFromAcl -Object $Item -LoginName $UserName -UserAccount $User -Context $SPOContext
                            if($UserRemoved)
                            {
                                AddUsertoListItemAcl -Context $SPOContext -RoleDefinition $SetRoleDefinition -UserAccount $User -ListItem $Item
                                Add-Content -Path "$OutputPath\$OutputFile" -Value "`t Removed\Added user to ACL for List Item"
                            }
                        }
                    }
                }      
            }
    
            $SPOContext.Load($Site.Webs)
            Try{$SPOContext.ExecuteQuery()}
            Catch{Add-Content -Path "$OutputPath\$OutputFile" -Value "ERROR: Loading all Webs"}

            Get-SPOWebs -Object $Site -Context $SPOContext | ForEach-Object {
                $WebUrl = $_.url
                Add-Content -Path "$OutputPath\$OutputFile" -Value "`t`t Scanning subsite: $WebUrl"

                if($_.hasUniqueRoleAssignments)
                {
                    $UserRemoved = $null
                    $UserRemoved = RemoveUserFromAcl -Object $_ -LoginName $UserName -UserAccount $User -Context $SPOContext
                    if($UserRemoved)
                    {
                        AddUsertoAcl -Object $_ -Context $SPOContext -NewRoleDefinition $SetRoleDefinition -UserAccount $User
                        Add-Content -Path "$OutputPath\$OutputFile" -Value "`t`t Removed\Added user to ACL for Web"
                    }
                }
        
                $Lists = $_.Lists
                $SPOContext.Load($Lists)
                Try{$SPOContext.ExecuteQuery()}
                Catch{Add-Content -Path "$OutputPath\$OutputFile" -Value "ERROR: Loading all Lists"}
    
                ForEach ($List in $Lists)
                {
                    if(!($List.Hidden))
                    {        
                        $ListTitle = $List.Title
                        Add-Content -Path "$OutputPath\$OutputFile" -Value "`t`t Scanning list: $ListTitle"
                        Invoke-LoadMethod -Object $List -PropertyName "HasUniqueRoleAssignments"
                        $SPOContext.ExecuteQuery()
                        if($List.HasUniqueRoleAssignments)
                        {
                            $UserRemoved = $null
                            $UserRemoved = RemoveUserFromAcl -Object $List -LoginName $UserName -UserAccount $User -Context $SPOContext
                            if($UserRemoved)
                            {
                                AddUsertoAcl -Object $List -Context $SPOContext -NewRoleDefinition $SetRoleDefinition -UserAccount $User
                                Add-Content -Path "$OutputPath\$OutputFile" -Value "`t`t Removed\Added user to ACL for List"
                            }
                        }
                        $ListItems = Get-ListItems -Context $SPOContext -List $List
                        ForEach ($Item in $ListItems)
                        {
                            Invoke-LoadMethod -Object $Item -PropertyName "HasUniqueRoleAssignments"
                            #Invoke-LoadMethod -Object $Item -PropertyName "RoleDefinitions"
                            $SPOContext.ExecuteQuery()
                            if($Item.hasUniqueRoleAssignments)
                            {
                                $UserRemoved = $null
                                $UserRemoved = RemoveUserFromAcl -Object $Item -LoginName $UserName -UserAccount $User -Context $SPOContext
                                if($UserRemoved)
                                {
                                    AddUsertoListItemAcl -Context $SPOContext -RoleDefinition $SetRoleDefinition -UserAccount $User -ListItem $Item
                                    Add-Content -Path "$OutputPath\$OutputFile" -Value "`t Removed\Added user to ACL for List Item"
                                }
                            }
                        }                    
                    }
                } 
            }
        }
    }
}
$Date = Get-Date
Add-Content -Path "$OutputPath\$OutputFile" -Value $Date