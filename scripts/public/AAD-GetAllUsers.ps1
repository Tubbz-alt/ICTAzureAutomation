<#
.SYNOPSIS 
    Dump all AAD entries to a CSV for further processing

.DESCRIPTION

.NOTES
    TODO
    - Add group membership checks (NSHE and NHSE Partners at least)
    - List users admin roles
    - Get services from licenses other than E3

    NOTE: This script must be run under Azure Automation Only
    --------------------------------------------------------------------------------
    Author: Julian Knight, Head of Corporate ICT Technology & Security, NHS England
    History:
        2016-05-05: Initial Version
 #>

$VerbosePreference = 'Continue' # Uncomment to get verbose output

$invocation = (Get-Variable MyInvocation).Value
$cmdName = $invocation.MyCommand.Name
$strt = get-date

Write-Verbose ""
Write-Verbose "Starting $cmdName at $strt"
Write-Verbose ""

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
        [Microsoft.Online.Administration.UserLicense]$licenses,

        # Updated User Object
        [parameter(Mandatory=$true)]
        [Microsoft.Online.Administration.User]$cols
    )

    for ($i=0; $i -lt $($licenses.count); $i++) {
        $license = $licenses[$i]

        # Note that $licenses = $user.Licenses is an object, $cols.Licenses is a ;-separated string

        # E3
        if ($cols.Licenses -match ':ENTERPRISEPACK') {
            $cols.IsE3 = $true

            # Find current users service plans & provisioning status
            foreach ($servicePlan in $license.ServiceStatus) {
                # Set a flag for each service active on this account
                # Values represent provisioning status rather than just true/false (see Show-ServiceStatus function for details)

                if ( ($servicePlan.ServicePlan.ServiceName -eq 'EXCHANGE_S_ENTERPRISE') ) {
                    $cols.HasExchange = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
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
                if ( ($servicePlan.ServicePlan.ServiceName -eq 'INTUNE_O365') ) {
                    $cols.HasIntune = Show-ServiceStatus -status $servicePlan.ProvisioningStatus
                }                
            }
        }

        # SharePoint P2
        if ($cols.Licenses -match 'SHAREPOINTENTERPRISE') {
            $cols.IsP2 = $true
            $cols.HasSP = 'Y'
        }

        # AAD Basic
        if ($cols.Licenses -match 'AAD_BASIC') {
            $cols.IsAADb = $true
        }

        #CRMENTERPRISE IsCrmEnt
        if ($cols.Licenses -match 'CRMENTERPRISE') {
            $cols.IsCrmEnt = $true
        }

        #CRMPLAN2 IsCrmP2
        if ($cols.Licenses -match 'CRMPLAN2') {
            $cols.IsCrmP2 = $true
        }

        #CRMSTANDARD IsCrmStd
        if ($cols.Licenses -match 'CRMSTANDARD') {
            $cols.IsCrmStd = $true
        }

        #PLANNERSTANDALONE IsPlan
        if ($cols.Licenses -match 'PLANNERSTANDALONE') {
            $cols.IsPlan = $true
        }

        #POWER_BI_PRO IsBiPro
        if ($cols.Licenses -match 'POWER_BI_PRO') {
            $cols.IsBiPro = $true
        }

        #POWER_BI_STANDARD IsBiStd
        if ($cols.Licenses -match 'POWER_BI_STANDARD') {
            $cols.IsBiStd = $true
        }

        #PROJECTCLIENT IsProj
        if ($cols.Licenses -match 'PROJECTCLIENT') {
            $cols.IsProj = $true
        }

        #VISIOCLIENT IsVisio
        if ($cols.Licenses -match 'VISIOCLIENT') {
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

    # We want most of the AAD user object but some bits need expanding, also get rid of rubbish
    $cols = $user | Select-Object -ExcludeProperty `
        AlternateEmailAddresses,AlternateMobilePhones,Licenses,AlternativeSecurityIds,ServiceInformation,ProxyAddresses,
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

    # Add detailed member type info
    Add-Member -InputObject $cols -MemberType NoteProperty -Name UserTypeExtended -Value (Get-MemberTypeExtended -user $user)

    # How many days since last password change and since token was last refreshed
    # WARNING: This data does not seem very reliable!
    Add-Member -InputObject $cols -MemberType NoteProperty -Name PwAge -Value (New-TimeSpan -Start $user.LastPasswordChangeTimestamp).Days -ErrorAction SilentlyContinue
    Add-Member -InputObject $cols -MemberType NoteProperty -Name StsAge -Value (New-TimeSpan -Start $user.StsRefreshTokensValidFrom).Days -ErrorAction SilentlyContinue

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


Get-MsolUser -MaxResults 5 | ForEach-Object `
    -Begin {
        Write-Verbose 'Start processing users'
    } `
    -Process {
        $usersOut.add( (Build-UserRow -User $_) )
    } `
    -End {
        Write-Verbose 'Finished processing users'
    }

$strt = get-date
Write-Verbose ""
Write-Verbose "Ending $cmdName at $strt"
Write-Verbose ""

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