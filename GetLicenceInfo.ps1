<#
.SYNOPSIS  
    Enumerates Tenant and User level licenses, provides answer to targeted questions
.DESCRIPTION  
    Make sure PC has Internet access or MSOnline module is already installed https://docs.microsoft.com/en-us/powershell/module/Azuread/?view=azureadps-2.0
    Make sure you run this script as an administrator on the PC
    Make sure PowerShell is RunAs Admin for best experience
.NOTES

This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED 
TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS for A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free right to use and modify 
the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or 
trademarks to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in 
which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, 
including attorneys fees, that arise or result from the use or distribution of the Sample Code.
#>

$ErrorActionPreference = "SilentlyContinue"

$OutputFolder = "$env:Temp\LicenseData"
$TenantLicenseFile = "Tenantlicense.csv"
$UserLicenseFile = "UserLicense.csv"
$GlobalAdminCredential = Get-Credential -Message "Enter Global Admin credentials"

#Trust repository so we can automate install of MSOnline
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted 

#PowerShellGet requires NuGet
If(!(Get-PackageProvider -Name NuGet))
{
    Install-PackageProvider -Name NuGet -Force  
}

#Verify MSOnline is installed or Update\Install MSOnline module 
if (Get-InstalledModule -Name MSOnline)
{
    Update-Module MSOnline
    Import-Module MSOnline
}Else
{
    Install-Module MSOnline
    Import-Module MSOnline
}
$Error.Clear()
try{Connect-MsolService -Credential $GlobalAdminCredential}
Catch{Write-Host "Failed to log into O365" -ForegroundColor Red}

If(!$Error)
{
    if(!(Test-Path -Path $OutputFolder -PathType Container))
    {
        New-Item -Path $OutputFolder -ItemType "Directory" | Out-Null
    }
    Set-Content -Path "$OutputFolder\$TenantLicenseFile" -Value "Plan,ActiveUnits,ConsumedUnits,LockedOutUnits,Service,Status"

    Get-MsolAccountSku | ForEach-Object {
        $Plan = $_.SkuPartNumber
        $ActiveUnits = $_.ActiveUnits
        $ConsumedUnits = $_.ConsumedUnits
        $LockedOutUnits = $_.LockedOutUnits
        Add-Content -Path "$OutputFolder\$TenantLicenseFile" -Value "$Plan,$ActiveUnits,$ConsumedUnits,$LockedOutUnits,,"
        $_.ServiceStatus | ForEach-Object {
            $Service = $_.ServicePlan.ServiceName
            $ServiceStatus = $_.ProvisioningStatus
            Add-Content -Path "$OutputFolder\$TenantLicenseFile" -Value ",,,,$Service,$ServiceStatus"
        }
    }

    Add-Content -Path "$OutputFolder\$TenantLicenseFile" -Value "NOTE: ProvisioninsStatus of PendingInput means it requires Configuration"
    
    Set-Content -Path "$OutputFolder\$UserLicenseFile" -Value "User,Plan,Service,ServiceStatus"
    Set-Content -Path "$OutputFolder\Answers.txt" -Value $(Get-Date)
    $E3 = $false
    $E1 = $false
    $E5 = $false
    $E1notE5User = 0
    $E1notE3User = 0
    Get-MsolUser | ForEach-Object {
        $User = $_.UserPrincipalName
        if($_.IsLicensed)
        {
            $_.Licenses | ForEach-Object {
                $Plan = $_.AccountSkuId
                Switch -Wildcard ($Plan){
                    "*ENTERPRISEPREMIUM" {$E5 = $true}
                    "*ENTERPRISEPACK" {$E3 = $true}
                    "*STANDARDPACK" {$E1 = $true}
                }
                
                ForEach($Service in $_.ServiceStatus)
                {
                    $ServiceName = $Service.ServicePlan.ServiceType
                    $ServiceStatus = $Service.ProvisioningStatus
                    Add-Content -Path "$OutputFolder\$UserLicenseFile" -Value "$User,$Plan,$ServiceName,$ServiceStatus"
                }
            }
        }Else
        {
            Add-Content -Path "$OutputFolder\$UserLicenseFile" -Value "$User,,,,"
        }
        if($E1 -and !$E5)
        {
            $E1notE5User += 1
        }
        If($E1 -and !$E3)
        {
            $E1notE3User += 1
        }
    }
    Add-Content -Path "$OutputFolder\Answers.txt" -Value "There are $E1notE3user users who have E3 but not E1 SKU's"
    Add-Content -Path "$OutputFolder\Answers.txt" -Value "There are $E1notE5user users who have E5 but not E1 SKU's"
}Else
{
    Write-Host $Error[0].Exception -ForegroundColor Red
}
Invoke-Item $OutputFolder

#https://blogs.technet.microsoft.com/treycarlee/2014/12/09/powershell-licensing-skus-in-office-365/
#ENTERPRISEPREMIUM – E5 License
#ENTERPRISEPACK – E3 License
#STANDARDPACH - E1 License