<#
.SYNOPSIS 
    Dump all AAD entries to a CSV for further processing, save the CSV in files uploaded to SharePoint

.DESCRIPTION
    Dumps out all Azure Active Directory entries, expanding those that are normally buried in nexted objects so
    that they are more usable in downstream analytics.

    The file is then uploaded to a SharePoint document library as a latest and a dated version.

.INPUTS
    None

.OUTPUTS
    2 CSV Files in the specificed SharePoint document library, one marked as fname-Latest.csv and the other fname-yyyymmdd.csv

.LINK
    https://github.com/nhsengland/ICTAzureAutomation/blob/master/scripts/public/AAD-GetAllUsers.ps1

.NOTES
    If you need to detect whether running locally (with the ISE addin), either check for the $psISE variable
    or check if the machine="Client" and/or the user="Client".

    PREREQUISITES
    - Azure Automation
    - Azure AD (MSOL) PowerShell Module
    - Office 365 PnP PowerShell Module
    - Azure Automation Assets:
      - Login Credential with suitable rights to AAD and SP
      - Various variables holding domain/tenant and other more sensitive information - so no need to keep them in this script
    - Azure Automation ISE add-on (if you want to test locally) [Install-Module -Name AzureAutomationAuthoringToolkit; Install-AzureAutomationIseAddOn;]

    TODO
    - Add group membership checks (NSHE and NHSE Partners at least)
    - List users admin roles
    - Get services from licenses other than E3

    NOTE: This script should be run under Azure Automation
          However, it does work locally if you have the AzureAutomationAuthoringToolkit module installed and active
    --------------------------------------------------------------------------------
    Author: Julian Knight, Head of Corporate ICT Technology & Security, NHS England
    History:
        2016-05-14: Fixed, uploads to SP, works locally and under AA
        2016-05-05: Initial Version
 #>

#$VerbosePreference = 'Continue' # Uncomment to get verbose output

$invocation = (Get-Variable MyInvocation).Value
$cmdName = $invocation.MyCommand.Name
$strt = get-date

Write-Output "Starting $cmdName at $strt"

#region "Setup" # ----------------------------------------------------------------

$tenant = Get-AutomationVariable -Name 'tenant'                              # your tenant name (fetched from an Azure Automation resource)
$spoRootURL = "https://{0}.sharepoint.com/" -f $tenant
$spoAdminURL = "https://{0}-admin.sharepoint.com/" -f $tenant
$aadDefaultDomain = "{0}.onmicrosoft.com" -f $tenant
$aadCustomDomain = Get-AutomationVariable -Name 'aadCustomDomain'            # a custom domain used for user identities, e.g. custom.co.uk
$aadGenericPrefix = Get-AutomationVariable -Name 'aadGenericPrefix'          # An ID prefix used to identify generic accounts, assumes a dot between this and the rest of the ID, e.g. aadGenericPrefix.acct1@tenantname.onmicrosoft.om
$aadCustomDomain2 = Get-AutomationVariable -Name 'aadCustomDomain2'          # a custom domain used for user identities, e.g. custom2.net
$aadCustomDomain2Name = Get-AutomationVariable -Name 'aadCustomDomain2Name'  # A display name to match with the custom domain2

$usersOut = New-Object System.Collections.Generic.List[System.Object]

$ICTo365AuditSite = Get-AutomationVariable -Name 'ICT-o365AuditSite'         # The site to upload output CSV files to
$ICTo365AuditFolder = Get-AutomationVariable -Name 'ICT-o365AuditFolder'     # The folder to upload output CSV files to

# Temporary output & file names
$outPath = $env:TEMP   # C:\Users\Client\Temp
$outFile = Join-Path -Path $outPath -ChildPath "AAD-AllUsers-Latest.csv"
$DateStr = Get-Date -format "yyyyMMdd" # Get current date for adding to file names
$outFileDated = Join-Path -Path $outPath -ChildPath "AAD-AllUsers-$DateStr.csv"

# Check if any of the temporary CSV files exist
if ( Test-Path -Path (Join-Path -Path $outPath -ChildPath "*.csv") ) {
    # CSV files exist in the output path so this script is looping!
    Throw "CSV files exist in output folder - script is looping!"
}

#endregion "Setup" # -------------------------------------------------------------


#region "Functions" # ------------------------------------------------------------

function Get-MemberTypeExtended { 
    Param (
        # User Object from MSOL-GetUser
        [parameter(Mandatory=$true)]
        [Microsoft.Online.Administration.User]$user
    )

    # NOTE that these all have to be in the right order! #

    if ($user.UserPrincipalName -match '#EXT#') { # NB: Not Licensed
        $UserType1 = 'Guest'
    }
    if ($user.UserPrincipalName.StartsWith('SMO-')) { # NB: Not Licensed
        $UserType1 = 'Site Mailbox'
    }
    if ($user.UserPrincipalName.StartsWith("$aadGenericPrefix.")) {
        $UserType1 = 'Generic'
    }
    if ( ($user.UserPrincipalName.StartsWith('FinReturns.')) -or ($user.UserPrincipalName -eq "finance@$aadCustomDomain") ) {
        $UserType1 = 'Finance'
    }
    if ($user.UserPrincipalName.StartsWith('Audit.') -and $user.UserPrincipalName.EndsWith("@$aadDefaultDomain")) {
        $UserType1 = 'Ext Audit'
    }
    if ($user.UserPrincipalName.Contains('incident')) {
        $UserType1 = 'EPRR'
    }
    if ($user.UserPrincipalName.EndsWith( "@$aadDefaultDomain" )) {
        $UserType1 = 'Partner'
    }
    if ($user.UserPrincipalName.EndsWith( "@$aadCustomDomain2" )) {
        $UserType1 = $aadCustomDomain2Name
    }

    # ----------------------- #
    try {
        .\AAD-CustomMemberTypes.ps1
    } catch {
        Write-Verbose '.\AAD-CustomMemberTypes.ps1 Failed to run'
    }
    # ----------------------- #

    if ( ($user.IsLicensed -eq $false) -and ($user.ProxyAddr -ne '') -and ($user.UserPrincipalName -notmatch '#EXT#') ) {
        $UserType1 = 'Shared Mailbox' # NB: Not Licensed
    }
    if ( ($user.UserType1 -eq '') -or ($user.UserType1 -eq $null) ) {
        $UserType1 = 'Member'
    }

    return $UserType1
} # ---- End of function Get-MemberTypeExtended ---- #

function Track-LicensesAndServices { 
    Param (
        # User Object from MSOL-GetUser
        [parameter(Mandatory=$true)]
        $licenses, #[Microsoft.Online.Administration.UserLicense]

        # Updated User Object
        [parameter(Mandatory=$true)]
        $cols # [Microsoft.Online.Administration.User]
    )

    for ($i=0; $i -lt $($licenses.count); $i++) {
        $license = $licenses[$i]

        # Note that $licenses = $user.Licenses is an object, $cols.Licenses is a ;-separated string

        #AAD Basic
        if ($license.AccountSkuId -match 'AAD_BASIC') {
            # Service Names: AAD_BASIC
            $cols.IsAADb = $true
        }

        #CRMENTERPRISE IsCrmEnt
        if ($license.AccountSkuId -match 'CRMENTERPRISE') {
            # Service Names: CRMENTERPRISE, DMENTERPRISE, NBENTERPRISE, PARATURE_ENTERPRISE
            $cols.IsCrmEnt = $true
        }

        # NB: CRMINSTANCE: Not interested with that as it isn't user level licensing

        #CRMPLAN2 IsCrmP2
        if ($license.AccountSkuId -match 'CRMPLAN2') {
            # Service Names: CRMPLAN2
            $cols.IsCrmP2 = $true
        }

        #CRMSTANDARD IsCrmStd
        if ($license.AccountSkuId -match 'CRMSTANDARD') {
            # Service Names: CRMSTANDARD, MDM_SALES_COLLABORATION, NBPROFESSIONALFORCRM
            $cols.IsCrmStd = $true
        }

        # NB: CRMSTORAGE: Not interested with that as it isn't user level licensing
        # NB: CRMTESTINSTANCE: Not interested with that as it isn't user level licensing

        #ENTERPRISEPACK (E3)
        if ($license.AccountSkuId -match 'ENTERPRISEPACK') {
            # Service Names: EXCHANGE_S_ENTERPRISE, INTUNE_O365, MCOSTANDARD, OFFICESUBSCRIPTION, RMS_S_ENTERPRISE, SHAREPOINTENTERPRISE, SHAREPOINTWAC, SWAY, YAMMER_ENTERPRISE

            $cols.IsE3 = $true

            # Find current users service plans & provisioning status
            foreach ($servicePlan in $license.ServiceStatus) {
                # Set a flag for each service active on this account
                # Values represent provisioning status rather than just true/false (see Show-ServiceStatus function for details)

                if ( ($servicePlan.ServicePlan.ServiceName -eq 'EXCHANGE_S_ENTERPRISE') ) {
                    $cols.HasExchange = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'INTUNE_O365') ) {
                    $cols.HasIntune = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }                
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'MCOSTANDARD') ) {
                    $cols.HasLync = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'OFFICESUBSCRIPTION') ) {
                    $cols.HasOffice = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'RMS_S_ENTERPRISE') ) {
                    $cols.HasRMS = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'SHAREPOINTENTERPRISE') ) {
                    $cols.HasSP = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'SHAREPOINTWAC') ) {
                    $cols.HasWAC = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'SWAY') ) {
                    $cols.HasSway = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'YAMMER_ENTERPRISE') ) {
                    $cols.HasYam = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }
            }
        }

        #PLANNERSTANDALONE IsPlan
        if ($license.AccountSkuId -match 'PLANNERSTANDALONE') {
            $cols.IsPlan = $true
        }

        #POWER_BI_PRO IsBiPro
        if ($license.AccountSkuId -match 'POWER_BI_PRO') {
            $cols.IsBiPro = $true
        }

        #POWER_BI_STANDARD IsBiStd
        if ($license.AccountSkuId -match 'POWER_BI_STANDARD') {
            $cols.IsBiStd = $true
        }

        #PROJECTCLIENT IsProj
        if ($license.AccountSkuId -match 'PROJECTCLIENT') {
            $cols.IsProj = $true
        }

        #SHAREPOINTENTERPRISE (SharePoint P2)
        if ($license.AccountSkuId -match 'SHAREPOINTENTERPRISE') {
            # Service Names: INTUNE_O365, SHAREPOINTENTERPRISE

            $cols.IsP2 = $true

            if ( ($servicePlan.ServicePlan.ServiceName -eq 'SHAREPOINTENTERPRISE') ) {
                $cols.HasSP = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
            }

            if ( ($servicePlan.ServicePlan.ServiceName -eq 'INTUNE_O365') ) {
                $cols.HasIntune = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
            }                
        }

        #VISIOCLIENT IsVisio
        if ($license.AccountSkuId -match 'VISIOCLIENT') {
            $cols.IsVisio = $true
        }

    }

} # ---- End of function Track-LicensesAndServices ---- #

# Return the service status for a licensed O365 service
function Show-ServiceStatus {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        #[Parameter(Mandatory=$true)]
        [ValidateSet('Success', 'Disabled', 'PendingActivation', 'PendingInput', 'Error')]
        $status
    )
 
    switch ($status) {
        'Success' {
            'Y'
            break
        }
        'Disabled' {
            'N'
            break
        }
        'PendingActivation' {
            'P'
            break
        }
        'PendingInput' {
            'P'
            break
        }
        default {
            Write-Verbose "Unknown license service status $status"
            'N'
        }
    }
} # ---- End of function Show-ServiceStatus ---- #

# Build a new user object
function Build-UserRow {
    Param (
        # User Object from MSOL-GetUser
        [parameter(Mandatory=$true)]
        [Microsoft.Online.Administration.User]$user
    )

    # We want most of the AAD user object but some bits need expanding, also get rid of rubbish (note that you HAVE to use -Property * or the exclude doesn't work!
    $cols = $user | Select-Object -Property * -ExcludeProperty `
        AlternateEmailAddresses,AlternateMobilePhones,Licenses,AlternativeSecurityIds,ServiceInformation,ProxyAddresses,DirSyncProvisioningErrors,
        Errors,ExtensionData,ImmutableId,IndirectLicenseErrors,IsBlackberryUser,LiveId,PasswordResetNotRequiredDuringActivate,
        StrongAuthenticationMethods,StrongAuthenticationPhoneAppDetails,StrongAuthenticationProofupTime,StrongAuthenticationRequirements,StrongAuthenticationUserDetails
    
    #region "ExtraCols"
    # Flatten some of the more complex data for easier processing
    Add-Member -InputObject $cols -MemberType NoteProperty -Name AlternateEmailAddresses -Value ($user.AlternateEmailAddresses -join ';')
    Add-Member -InputObject $cols -MemberType NoteProperty -Name AlternateMobilePhones -Value ($user.AlternateMobilePhones -join ';')
    Add-Member -InputObject $cols -MemberType NoteProperty -Name Licenses -Value ($user.Licenses.AccountSkuId -join ';')
    Add-Member -InputObject $cols -MemberType NoteProperty -Name AltSecurityIDs -Value ($user.AlternativeSecurityIds.Length -join ';')
    Add-Member -InputObject $cols -MemberType NoteProperty -Name ServiceInfo -Value ($user.ServiceInformation.ServiceInstance -join ';')
    Add-Member -InputObject $cols -MemberType NoteProperty -Name ProxyAddr -Value ($user.ProxyAddresses -join ';')
    Add-Member -InputObject $cols -MemberType NoteProperty -Name DirSyncProvisioningErrors -Value ($user.DirSyncProvisioningErrors -join ';')

    # Add detailed member type info
    Add-Member -InputObject $cols -MemberType NoteProperty -Name UserTypeExtended -Value (Get-MemberTypeExtended -user $user)

    # How many days since last password change and since token was last refreshed
    # WARNING: This data does not seem very reliable!
    try {
        Add-Member -InputObject $cols -MemberType NoteProperty -Name PwAge -Value (New-TimeSpan -Start $user.LastPasswordChangeTimestamp).Days
    } catch {
        Add-Member -InputObject $cols -MemberType NoteProperty -Name PwAge -Value (New-TimeSpan -Start $user.WhenCreated).Days
    }
    try {
        Add-Member -InputObject $cols -MemberType NoteProperty -Name StsAge -Value (New-TimeSpan -Start $user.StsRefreshTokensValidFrom).Days
    } catch {
        Add-Member -InputObject $cols -MemberType NoteProperty -Name StsAge -Value (New-TimeSpan -Start $user.WhenCreated).Days
    }

    # --- Track license and service allocation (a license may contain several services) --- #
    # License flags - track which accounts have which licenses assigned - default values first
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsE3 -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsP2 -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsAADb -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsCrmEnt -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsCrmP2 -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsCrmStd -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsPlan -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsBiPro -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsBiStd -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsProj -Value $false
    Add-Member -InputObject $cols -MemberType NoteProperty -Name IsVisio -Value $false

    # Service flags - track which accounts have which services active - default values first
    # Values represent provisioning status rather than just true/false (see Show-ServiceStatus function for details)
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasExchange -Value 'N'
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasLync -Value 'N'
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasOffice -Value 'N'
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasRMS -Value 'N'
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasSP -Value 'N'
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasWAC -Value 'N'
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasSway -Value 'N'
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasYam -Value 'N'
    Add-Member -InputObject $cols -MemberType NoteProperty -Name HasInTune -Value 'N'
    #endregion "ExtraCols"

    # Make sure that street address doesn't contain a \n or \r (they can make a mess of the CSV output)
    $cols.StreetAddress.replace("`r"," ").replace("`n"," ")

    # For each license assigned to this acct...
    Track-LicensesAndServices -licenses $user.Licenses -cols $cols
    
    return $cols
} # ---- End of function Build-UserRow ---- #

#endregion "Functions" # ---------------------------------------------------------



#region "Login" # ----------------------------------------------------------------

$credential = Get-AutomationPSCredential -Name 'O365'
if ($Credential -eq $null) {
    throw "Could not retrieve 'O365' credential asset. Check that you created this asset in the Automation service."
}     

#Import-Module -Name MSOnline

# Connect to Azure AD
Connect-MsolService -Credential $credential

#endregion "Login" # -------------------------------------------------------------


Get-MsolUser -All | ForEach-Object `
    -Begin {
        Write-Verbose 'Start processing users'
    } `
    -Process {
        $usersOut.add( (Build-UserRow -User $_) )
    } `
    -End {
        Write-Verbose 'Finished processing users'
    }

Write-Output ("Number of records created: {0}" -f ($usersOut | measure).Count )
#$usersOut | gm
#$usersOut | ft

# Save as temp file
$usersOut | Export-Csv -Path $outFile -NoTypeInformation -Force

# Import PnP commands

# Connect to SP Web
Connect-SPOnline -Url ("{0}{1}" -f $spoRootURL, $ICTo365AuditSite) -Credentials $credential

# Load temp file to SP (using PnP)
Write-Output "Writing $outFile to $ICTo365AuditFolder"
Add-SPOFile -Path $outFile -Folder $ICTo365AuditFolder

# Rename local file to dated version
Write-Output "Moving $outFile to $outFileDated"
Move-Item -Path $outFile -Destination $outFileDated

# Create a dated copy for further analysis (e.g. changes over time)
Write-Output "Writing $outFileDated to $ICTo365AuditFolder"
Add-SPOFile -Path $outFileDated -Folder $ICTo365AuditFolder

# Tidy Up
Write-Output "Tidying up (Disconnecting from SPOnline)"
Disconnect-SPOnline
# Let the end of the script destroy the files so we can 
# check if the script is looping (which can be a problem with AA)
#Remove-Item $outFile -Force -ErrorAction SilentlyContinue # just in case the first upload failed
#Remove-Item $outFileDated -Force -ErrorAction SilentlyContinue

$strt = get-date
Write-Output "Ending $cmdName at $strt"

<#
Get-MsolUser -MaxResults 1 | Get-Member


   TypeName: Microsoft.Online.Administration.User

Name                                   MemberType Definition                                                                                                      
----                                   ---------- ----------                                                                                                      
Equals                                 Method     bool Equals(System.Object obj)                                                                                  
GetHashCode                            Method     int GetHashCode()                                                                                               
GetType                                Method     type GetType()                                                                                                  
ToString                               Method     string ToString()                                                                                               
  AlternateEmailAddresses                Property   System.Collections.Generic.List[string] AlternateEmailAddresses {get;set;}                                      
  AlternateMobilePhones                  Property   System.Collections.Generic.List[string] AlternateMobilePhones {get;set;}                                        
  AlternativeSecurityIds                 Property   System.Collections.Generic.List[Microsoft.Online.Administration.AlternativeSecurityId] AlternativeSecurityIds...
BlockCredential                        Property   System.Nullable[bool] BlockCredential {get;set;}                                                                
City                                   Property   string City {get;set;}                                                                                          
CloudExchangeRecipientDisplayType      Property   System.Nullable[int] CloudExchangeRecipientDisplayType {get;set;}                                               
Country                                Property   string Country {get;set;}                                                                                       
Department                             Property   string Department {get;set;}                                                                                    
DirSyncProvisioningErrors              Property   System.Collections.Generic.List[Microsoft.Online.Administration.DirSyncProvisioningError] DirSyncProvisioning...
DisplayName                            Property   string DisplayName {get;set;}                                                                                   
  Errors                                 Property   System.Collections.Generic.List[Microsoft.Online.Administration.ValidationError] Errors {get;set;}              
  ExtensionData                          Property   System.Runtime.Serialization.ExtensionDataObject ExtensionData {get;set;}                                       
Fax                                    Property   string Fax {get;set;}                                                                                           
FirstName                              Property   string FirstName {get;set;}                                                                                     
  ImmutableId                            Property   string ImmutableId {get;set;}                                                                                   
  IndirectLicenseErrors                  Property   System.Collections.Generic.List[Microsoft.Online.Administration.IndirectLicenseError] IndirectLicenseErrors {...
  IsBlackberryUser                       Property   System.Nullable[bool] IsBlackberryUser {get;set;}                                                               
IsLicensed                             Property   System.Nullable[bool] IsLicensed {get;set;}                                                                     
LastDirSyncTime                        Property   System.Nullable[datetime] LastDirSyncTime {get;set;}                                                            
LastName                               Property   string LastName {get;set;}                                                                                      
LastPasswordChangeTimestamp            Property   System.Nullable[datetime] LastPasswordChangeTimestamp {get;set;}                                                
LicenseReconciliationNeeded            Property   System.Nullable[bool] LicenseReconciliationNeeded {get;set;}                                                    
Licenses                               Property   System.Collections.Generic.List[Microsoft.Online.Administration.UserLicense] Licenses {get;set;}                
  LiveId                                 Property   string LiveId {get;set;}                                                                                        
MobilePhone                            Property   string MobilePhone {get;set;}                                                                                   
MSExchRecipientTypeDetails             Property   System.Nullable[long] MSExchRecipientTypeDetails {get;set;}                                                     
ObjectId                               Property   System.Nullable[guid] ObjectId {get;set;}                                                                       
Office                                 Property   string Office {get;set;}                                                                                        
OverallProvisioningStatus              Property   Microsoft.Online.Administration.ProvisioningStatus OverallProvisioningStatus {get;set;}                         
PasswordNeverExpires                   Property   System.Nullable[bool] PasswordNeverExpires {get;set;}                                                           
  PasswordResetNotRequiredDuringActivate Property   System.Nullable[bool] PasswordResetNotRequiredDuringActivate {get;set;}                                         
PhoneNumber                            Property   string PhoneNumber {get;set;}                                                                                   
PortalSettings                         Property   System.Xml.XmlElement PortalSettings {get;set;}                                                                 
PostalCode                             Property   string PostalCode {get;set;}                                                                                    
PreferredLanguage                      Property   string PreferredLanguage {get;set;}                                                                             
  ProxyAddresses                         Property   System.Collections.Generic.List[string] ProxyAddresses {get;set;}                                               
ServiceInformation                     Property   System.Collections.Generic.List[Microsoft.Online.Administration.ServiceInformation] ServiceInformation {get;s...
SignInName                             Property   string SignInName {get;set;}                                                                                    
SoftDeletionTimestamp                  Property   System.Nullable[datetime] SoftDeletionTimestamp {get;set;}                                                      
State                                  Property   string State {get;set;}                                                                                         
StreetAddress                          Property   string StreetAddress {get;set;}                                                                                 
  StrongAuthenticationMethods            Property   System.Collections.Generic.List[Microsoft.Online.Administration.StrongAuthenticationMethod] StrongAuthenticat...
  StrongAuthenticationPhoneAppDetails    Property   System.Collections.Generic.List[Microsoft.Online.Administration.StrongAuthenticationPhoneAppDetail] StrongAut...
  StrongAuthenticationProofupTime        Property   System.Nullable[long] StrongAuthenticationProofupTime {get;set;}                                                
  StrongAuthenticationRequirements       Property   System.Collections.Generic.List[Microsoft.Online.Administration.StrongAuthenticationRequirement] StrongAuthen...
  StrongAuthenticationUserDetails        Property   Microsoft.Online.Administration.StrongAuthenticationUserDetails StrongAuthenticationUserDetails {get;set;}      
StrongPasswordRequired                 Property   System.Nullable[bool] StrongPasswordRequired {get;set;}                                                         
StsRefreshTokensValidFrom              Property   System.Nullable[datetime] StsRefreshTokensValidFrom {get;set;}                                                  
Title                                  Property   string Title {get;set;}                                                                                         
UsageLocation                          Property   string UsageLocation {get;set;}                                                                                 
UserLandingPageIdentifierForO365Shell  Property   string UserLandingPageIdentifierForO365Shell {get;set;}                                                         
UserPrincipalName                      Property   string UserPrincipalName {get;set;}                                                                             
UserThemeIdentifierForO365Shell        Property   string UserThemeIdentifierForO365Shell {get;set;}                                                               
UserType                               Property   System.Nullable[Microsoft.Online.Administration.UserType] UserType {get;set;}                                   
ValidationStatus                       Property   System.Nullable[Microsoft.Online.Administration.ValidationStatus] ValidationStatus {get;set;}                   
WhenCreated                            Property   System.Nullable[datetime] WhenCreated {get;set;} 
#>
