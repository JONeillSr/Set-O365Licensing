<#
Dkawula_test@moldedfiberglass.com
DKawula_Admin@moldedfiberglass.com
#>

# Provide access to the common parameters
[CmdletBinding()]

Param
(
    # Specify the desired action
    [Parameter(Mandatory = $true, HelpMessage = 'Specify the action to take. Valid values are Enable, Disable, and Both.')]
    [ValidateSet("Enable", "Disable", "Switch")]
    [String[]]$LicensesDesiredAction = @("Enable", "Disable", "Switch"),

    # Specify the tenant name
    [Parameter(Mandatory = $false, HelpMessage = 'Specify the Azure AD tenant name where users and licenses exist')]
    [String[]]$TenantName = "moldedfiberglass:",

    # Specify the ID of the licenses to add
    [Parameter(Mandatory = $True, HelpMessage = "Specify the license to work with by its SKU ID.")]
    [String[]]$LicenseNameAdd = "SPE_E5",

    # Specify the ID of the licenses to remove
    [Parameter(Mandatory = $True, HelpMessage = "Specify the license to work with by its SKU ID.")]
    [String[]]$LicenseNameRemove = "EMS",

    # Specify the license plan ID to work with
    [Paramater(Mandatory = $False, HelpMessage = "Specify the license plan (service) to work with using its ID.")]
    [String[]]$LicenseService = "EXCHANGE_S_ENTERPRISE"
)

function Set-License {
    # Get all users from connected Azure AD tenant
    $AZUsers = Get-MsolUser -All

    # Add additional property to only those users assigned the specified license
    $AZusers | Add-Member -MemberType ScriptProperty -Name "LicenseName" -Value {
        ($this.Licenses).where( { $_.AccountSkuID -match $LicenseName }).accountskuID } -force

    # If a license service specified, add additional property to only those users where the specified license has the specified service enabled
    If ($LicenseService) {
        $AZUsers | Add-Member -MemberType ScriptProperty -Name "LicenseServiceProvisioningStatus" -Value {
            $lic = ($this.Licenses).where( { $_.AccountSkuID -match $LicenseName })
            ($lic.servicestatus).where( { $_.serviceplan.servicename -match $LicenseService }).ProvisioningStatus } -force
    }
    
    # Disable specific service plan within a license
    ForEach ($AZUser in $AZUsers) {
        If ($AZUser.LicenseName -match $LicenseName -and $AZUser.LicenseServiceProvisioningStatus -eq 'Success') {
            Write-Host $AzUser.UserPrincipalName
            # $LicenseOptions = New-MsolLicenseOptions -AccountSkuId moldedfiberglass:SPE_E5 -DisabledPlans EXCHANGE_S_ENTERPRISE
            # Set-MsolUserLicense -UserPrincipalName $AZUser -LicenseOptions $LicenseOptions
        }
    }

    # Remove license from user
    ForEach ($AzUser in $AZUsers) {
        If ($AZUser.LicenseName -match $LicenseName) {
            Set-MsolUserLicense -UserPrincipalName $AZUser.UserPrincipalName -RemoveLicenses $FullLicenseName
        }
        else {
            Write-Host "User Not assigned matching license, no licenses removed."
        }
    }

    # Assign license to user
    ForEach ($AzUser in $AZUsers) {
        # Switching licenses
        If ($DesiredAction -eq "Switch") {
            try {
                Set-MsolUserLicense -UserPrincipalName $AZUser.UserPrincipalName -AddLicenses $FullLicenseNameAdd -RemoveLicenses $FullLicenseNameRemove
            }
            catch {
                Write-Host "User Not assigned matching license, no licenses removed."
            }
        }
    }
}

#Create full license name string from TenantName and LicenseName parameters
If ($TenantName) {
    $FullLicenseNameAdd = $TenantName + $LicenseNameAdd
    $FullLicenseNameRemove = $TenantName + $LicenseNameRemove
}

If ($LicensesDesiredAction -eq "Switch") {
    Set-License
}


<#
$AZUsers | Where-Object { $_.M365License -match "SPE_E5" } | Select-Object -Property Displayname, M365License, M365ProvisioningStatus | Export-CSV .\Start.CSV -NoTypeInformation

$AZusers | Where-Object { $_.M365License -match "SPE_E5" -and $_.M365ProvisioningStatus -ne 'Success' }
#>