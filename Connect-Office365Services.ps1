<#
    .SYNOPSIS
    Connect-Office365Services

    PowerShell script defining functions to connect to Office 365 online services
    or Exchange On-Premises. Call manually or alternatively embed or call from $profile
    (Shell or ISE) to make functions available in your session. If loaded from
    PowerShell_ISE, menu items are defined for the functions. To surpress creation of
    menu items, hold 'Shift' while Powershell ISE loads.

    Michel de Rooij
    michel@eightwone.com
    http://eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.86, February 16th, 2018

    KNOWN LIMITATIONS:
    - When specifying PSSessionOptions for MFA, authentication fails (OAuth).
      Therefor, no PSSessionOptions are used for MFA.

    .LINK
    http://eightwone.com

    Revision History
    ---------------------------------------------------------------------
    1.2	    Community release
    1.3     Updated required version of Online Sign-In Assistant
    1.4	    Added (in-code) AzureEnvironment (Connect-AzureAD)
            Added version reporting for modules
    1.5     Added support for Exchange Online PowerShell module w/MFA
            Added IE proxy config support
            Small cosmetic changes in output
    1.51    Fixed PSSession for non-MFA EXO logons
            Changed credential entering logic for MFA
    1.6     Added support for the Skype for Business PowerShell module w/MFA
            Added support for the SharePoint Online PowerShell module w/MFA
    1.61    Fixed MFA choice bug
    1.7     Added AzureAD PowerShell Module support
            For disambiguation, renamed Connect-AzureAD to Connect-AzureActiveDirectory
    1.71    Added AzureADPreview PowerShell Module Support
    1.72    Changed credential non-prompting condition for AzureAD
    1.75    Added support for MFA-enabled Security & Compliance Center
            Added module version checks (online when possible) setting OnlineModuleVersionChecks
            Switched AzureADv1 to PS gallery version
            Removed Sign-In Assistant checks
            Added Set-Office365Environment to switch to other region, e.g. Germany, China etc.
    1.76    Fixed version number checks for SfB & SP
    1.77    Fixed version number checks for determining MFA status
            Changed default for online version checks to $false
    1.78    Added usage of most recently dated ExO MFA module found (in case multiple are found)
            Changed connecting to SCC with MFA to using the underlying New-ExO cmdlet
    1.80    Added Microsoft Teams PowerShell Module support
            Added Connect-MSTeams function
            Cleared default PSSessionOptions
            Some code rewrite (splatting)
    1.81    Added support for ExO module 16.00.2020.000 w/MFA -Credential support
    1.82    Bug fix SharePoint module version check
    1.83    Removed Credentials option for ExO/MFA connect
    1.84    Added Exchange ADAL loading support
    1.85    Fixed menu creation in ISE
    1.86    Updated version check for AzureADPreview (2.0.0.154)
            Added automatic module updating (Admin mode, OnlineModuleAutoUpdate & OnlineModuleVersionChecks)

    .DESCRIPTION
    The functions are listed below. Note that functions may call eachother, for example to
    connect to Exchange Online the Office 365 Credentials the user is prompted to enter these credentials.
    Also, the credentials are persistent in the current session, there is no need to re-enter credentials
    when connecting to Exchange Online Protection for example. Should different credentials be required,
    call Get-Office365Credentials or Get-OnPremisesCredentials again.

    - Connect-AzureActiveDirectory	    Connects to Azure Active Directory
    - Connect-AzureRMS           	    Connects to Azure Rights Management
    - Connect-ExchangeOnline     	    Connects to Exchange Online
    - Connect-SkypeOnline        	    Connects to Skype for Business Online
    - Connect-EOP                	    Connects to Exchange Online Protection
    - Connect-ComplianceCenter   	    Connects to Compliance Center
    - Connect-SharePointOnline   	    Connects to SharePoint Online
    - Connect-MSTeams                       Connects to Microsoft Teams
    - Get-Office365Credentials    	    Gets Office 365 credentials
    - Connect-ExchangeOnPremises 	    Connects to Exchange On-Premises
    - Get-OnPremisesCredentials    	    Gets On-Premises credentials
    - Get-ExchangeOnPremisesFQDNGets	    FQDN for Exchange On-Premises
    - Get-Office365Tenant		    Gets Office 365 tenant name
    - Set-Office365Environment		    Configures Uri's and region to use

    .EXAMPLES
    .\Microsoft.PowerShell_profile.ps1
    Defines functions in current shell or ISE session (when $profile contains functions or is replaced with script).
#>

#Requires -Version 3.0

Write-Host 'Loading Connect-Office365Services v1.83'

$local:ExoPSSessionModuleVersion_Recommended = '16.00.2020.000'
$local:HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
$local:OnlineModuleVersionChecks = $false
$local:OnlineModuleAutoUpdate = $false

$local:Functions = @(
    'Connect|Exchange Online|Connect-ExchangeOnline',
    'Connect|Exchange Online Protection|Connect-EOP',
    'Connect|Exchange Compliance Center|Connect-ComplianceCenter',
    'Connect|Azure AD (v1)|Connect-MSOnline|MSOnline|Azure Active Directory (v1)|https://www.powershellgallery.com/packages/MSOnline|1.1.166.0',
    'Connect|Azure AD (v2)|Connect-AzureAD|AzureAD|Azure Active Directory (v2)|https://www.powershellgallery.com/packages/azuread|2.0.0.155',
    'Connect|Azure AD (v2 Preview)|Connect-AzureAD|AzureADPreview|Azure Active Directory (v2 Preview)|https://www.powershellgallery.com/packages/AzureADPreview|2.0.0.154',
    'Connect|Azure RMS|Connect-AzureRMS|AADRM|Azure RMS|https://www.microsoft.com/en-us/download/details.aspx?id=30339',
    'Connect|Skype for Business Online|Connect-SkypeOnline|SkypeOnlineConnector|Skype for Business Online|https://www.microsoft.com/en-us/download/details.aspx?id=39366|7.0.0.0',
    'Connect|SharePoint Online|Connect-SharePointOnline|Microsoft.Online.Sharepoint.PowerShell|SharePoint Online|https://www.microsoft.com/en-us/download/details.aspx?id=35588|16.0.6906.0',
    'Connect|Microsoft Teams|Connect-MSTeams|MicrosoftTeams|Microsoft Teams|https://www.powershellgallery.com/packages/MicrosoftTeams|0.9.0'
    'Settings|Office 365 Credentials|Get-Office365Credentials',
    'Connect|Exchange On-Premises|Connect-ExchangeOnPremises',
    'Settings|On-Premises Credentials|Get-OnPremisesCredentials',
    'Settings|Exchange On-Premises FQDN|Get-ExchangeOnPremisesFQDN'
)

$local:CreateISEMenu = $psISE -and -not [System.Windows.Input.Keyboard]::IsKeyDown( [System.Windows.Input.Key]::LeftShift)
If ( $local:CreateISEMenu) {Write-Host 'ISE detected, adding ISE menu options'}

# Local Exchange session options
$global:SessionExchangeOptions = New-PSSessionOption

function global:Set-Office365Environment {
    param(
        [ValidateSet('Germany', 'China', 'AzurePPE', 'USGovernment', 'Default')]
        [string]$Environment
    )
    Switch ( $Environment) {
        'Germany' {
            $global:ConnectionEndpointUri = 'https://outlook.office.de/PowerShell-LiveID'
            $global:SCCConnectionEndpointUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.de/common'
            $global:SharePointRegion = 'Germany'
            $global:AzureEnvironment = 'AzureGermanyCloud'
        }
        'China' {
            $global:ConnectionEndpointUri = 'https://partner.outlook.cn/PowerShell-LiveID'
            $global:SCCConnectionEndpointUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:AzureADAuthorizationEndpointUri = 'https://login.chinacloudapi.cn/common'
            $global:SharePointRegion = 'China'
            $global:AzureEnvironment = 'AzureChinaCloud'
        }
        'AzurePPE' {
            $global:ConnectionEndpointUri = ''
            $global:SCCConnectionEndpointUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:AzureADAuthorizationEndpointUri = ''
            $global:SharePointRegion = ''
            $global:AzureEnvironment = 'AzurePPE'
        }
        'USGovernment' {
            $global:ConnectionEndpointUri = 'https://outlook.office365.com/PowerShell-LiveId'
            $global:SCCConnectionEndpointUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:AzureADAuthorizationEndpointUri = 'https://login-us.microsoftonline.com/'
            $global:SharePointRegion = 'ITAR'
            $global:AzureEnvironment = 'AzureUSGovernment'
        }
        default {
            $global:ConnectionEndpointUri = 'https://outlook.office365.com/PowerShell-LiveId'
            $global:SCCConnectionEndpointUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
            $global:AzureADAuthorizationEndpointUri = 'https://login.windows.net/common'
            $global:SharePointRegion = 'Default'
            $global:AzureEnvironment = 'AzureCloud'
        }
    }
    Write-Host ('Environment set to {0}' -f $global:AzureEnvironment)
}
function global:Get-MultiFactorAuthenticationUsage {
    $Answer = Read-host  -Prompt 'Would you like to use Multi-Factor Authentication? (y/N) '
    Switch ($Answer.ToUpper()) {
        'Y' { $rval = $true }
        Default { $rval = $false}
    }
    return $rval
}

function global:Connect-ExchangeOnline {
    If ( !($global:Office365Credentials)) { Get-Office365Credentials }
    If ( $global:Office365CredentialsMFA) {
        Write-Host "Connecting to Exchange Online using $($global:Office365Credentials.username) with MFA .."
        $global:Session365 = New-ExoPSSession -ConnectionUri $global:ConnectionEndpointUri -UserPrincipalName ($global:Office365Credentials).UserName -AzureADAuthorizationEndpointUri $global:AzureADAuthorizationEndpointUri -PSSessionOption $global:SessionExchangeOptions
    }
    Else {
        Write-Host "Connecting to Exchange Online using $($global:Office365Credentials.username) .."
        $global:Session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $global:ConnectionEndpointUri -Credential $global:Office365Credentials -Authentication Basic -AllowRedirection -SessionOption $global:SessionExchangeOptions
    }
    If ( $global:Session365 ) {Import-PSSession -Session $global:Session365 -AllowClobber}
}

function global:Connect-ExchangeOnPremises {
    If ( !($global:OnPremisesCredentials)) { Get-OnPremisesCredentials }
    If ( !($global:ExchangeOnPremisesFQDN)) { Get-ExchangeOnPremisesFQDN }
    Write-Host "Connecting to Exchange On-Premises $($global:ExchangeOnPremisesFQDN) using $($global:OnPremisesCredentials.username) .."
    $global:SessionExchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($global:ExchangeOnPremisesFQDN)/PowerShell" -Credential $global:OnPremisesCredentials -Authentication Kerberos -AllowRedirection -SessionOption $global:SessionExchangeOptions
    If ( $global:SessionExchange) {Import-PSSession -Session $global:SessionExchange -AllowClobber}
}

Function global:Get-ExchangeOnPremisesFQDN {
    $global:ExchangeOnPremisesFQDN = Read-Host -Prompt 'Enter Exchange On-Premises endpoint, e.g. exchange1.contoso.com'
}

function global:Connect-ComplianceCenter {
    If ( !($global:Office365Credentials)) { Get-Office365Credentials }
    If ( $global:Office365CredentialsMFA) {
        Write-Host "Connecting to Office 365 Security & Compliance Center using $($global:Office365Credentials.username) with MFA .."
        $global:Session365 = New-ExoPSSession -ConnectionUri $global:SCCConnectionEndpointUri -UserPrincipalName ($global:Office365Credentials).UserName -AzureADAuthorizationEndpointUri $global:AzureADAuthorizationEndpointUri -PSSessionOption $local:SessionExchangeOptions
        New-IPPSSession -ConnectionUri $global:ConnectionEndpointUri -UserPrincipalName ($global:Office365Credentials).UserName -AzureADAuthorizationEndpointUri $global:AzureADAuthorizationEndpointUri -PSSessionOption $local:SessionExchangeOptions
    }
    Else {
        Write-Host "Connecting to Office 365 Security & Compliance Center using $($global:Office365Credentials.username) .."
        $global:SessionCC = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.compliance.protection.outlook.com/powershell-liveid/' -Credential $global:Office365Credentials -Authentication Basic -AllowRedirection
        If ( $global:SessionCC ) {Import-PSSession -Session $global:SessionCC -AllowClobber}
    }
}

function global:Connect-EOP {
    If ( !($global:Office365Credentials)) { Get-Office365Credentials }
    Write-Host  -InputObject "Connecting to Exchange Online Protection using $($global:Office365Credentials.username) .."
    $global:SessionEOP = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.protection.outlook.com/powershell-liveid/' -Credential $global:Office365Credentials -Authentication Basic -AllowRedirection
    If ( $global:SessionEOP ) {Import-PSSession -Session $global:SessionEOP -AllowClobber}
}

function global:Connect-MSTeams {
    If ( !($global:Office365Credentials)) { Get-Office365Credentials }
    If (($global:Office365Credentials).username -like '*.onmicrosoft.com') {
        $global:Office365Tenant = ($global:Office365Credentials).username.Substring(($global:Office365Credentials).username.IndexOf('@') + 1).Replace('.onmicrosoft.com', '')
    }
    If ( $global:Office365CredentialsMFA) {
        Write-Host "Connecting to Microsoft Teams using $($global:Office365Credentials.username) with MFA .."
        $Parms = @{'AccountId' = ($global:Office365Credentials).username}
    }
    Else {
        Write-Host "Connecting to Microsoft Teams using $($global:Office365Credentials.username) .."
        $Parms = @{Credential = $global:Office365Credentials }
    }
    If ( $global:Office365Tenant) { $Parms['TenantId'] = $global:Office365Tenant }
    Connect-MicrosoftTeams @Parms
}

function global:Connect-AzureActiveDirectory {
    If ( !(Get-Module -Name AzureAD)) {Import-Module -Name AzureAD -ErrorAction SilentlyContinue}
    If ( !(Get-Module -Name AzureADPreview)) {Import-Module -Name AzureADPreview -ErrorAction SilentlyContinue}
    If ( (Get-Module -Name AzureAD) -or (Get-Module -Name AzureADPreview)) {
        If ( !($global:Office365Credentials)) { Get-Office365Credentials }
        If ( $global:Office365CredentialsMFA) {
            Write-Host 'Connecting to Azure Active Directory with MFA ..'
            $Parms = @{'AzureEnvironment' = $global:AzureEnvironment}
        }
        Else {
            Write-Host "Connecting to Azure Active Directory using $($global:Office365Credentials.username) .."
            $Parms = @{'Credential' = $global:Office365Credentials; 'AzureEnvironment' = $global:AzureEnvironment}
        }
        Connect-AzureAD @Parms
    }
    Else {
        If ( !(Get-Module -Name MSOnline)) {Import-Module -Name MSOnline -ErrorAction SilentlyContinue}
        If ( Get-Module -Name MSOnline) {
            If ( !($global:Office365Credentials)) { Get-Office365Credentials }
            Write-Host "Connecting to Azure Active Directory using $($global:Office365Credentials.username) .."
            Connect-MsolService -Credential $global:Office365Credentials -AzureEnvironment $global:AzureEnvironment
        }
        Else {Write-Error -Message 'Cannot connect to Azure Active Directory - problem loading module.'}
    }
}

function global:Connect-AzureRMS {
    If ( !(Get-Module -Name AADRM)) {Import-Module -Name AADRM -ErrorAction SilentlyContinue}
    If ( Get-Module -Name AADRM) {
        If ( !($global:Office365Credentials)) { Get-Office365Credentials }
        Write-Host "Connecting to Azure RMS using $($global:Office365Credentials.username) .."
        Connect-AadrmService -Credential $global:Office365Credentials
    }
    Else {Write-Error -Message 'Cannot connect to Azure RMS - problem loading module.'}
}

function global:Connect-SkypeOnline {
    If ( !(Get-Module -Name SkypeOnlineConnector)) {Import-Module -Name SkypeOnlineConnector -ErrorAction SilentlyContinue}
    If ( Get-Module -Name SkypeOnlineConnector) {
        If ( !($global:Office365Credentials)) { Get-Office365Credentials }
        If ( $global:Office365CredentialsMFA) {
            Write-Host "Connecting to Skype for Business Online using $($global:Office365Credentials.username) with MFA .."
            $Parms = @{'Username' = ($global:Office365Credentials).username}
        }
        Else {
            Write-Host "Connecting to Skype for Business Online using $($global:Office365Credentials.username) .."
            $Parms = @{'Credential' = $global:Office365Credentials}
        }
        $global:SessionSFB = New-CsOnlineSession @Parms
        If ( $global:SessionSFB ) {Import-PSSession -Session $global:SessionSFB -AllowClobber}
    }
    Else {
        Write-Error -Message 'Cannot connect to Skype for Business Online - problem loading module.'
    }
}

function global:Connect-SharePointOnline {
    If ( !(Get-Module -Name Microsoft.Online.Sharepoint.PowerShell)) {Import-Module -Name Microsoft.Online.Sharepoint.PowerShell -ErrorAction SilentlyContinue}
    If ( Get-Module -Name Microsoft.Online.Sharepoint.PowerShell) {
        If ( !($global:Office365Credentials)) { Get-Office365Credentials }
        If (($global:Office365Credentials).username -like '*.onmicrosoft.com') {
            $global:Office365Tenant = ($global:Office365Credentials).username.Substring(($global:Office365Credentials).username.IndexOf('@') + 1).Replace('.onmicrosoft.com', '')
        }
        Else {
            If ( !($global:Office365Tenant)) { Get-Office365Tenant }
        }
        If ( $global:Office365CredentialsMFA) {
            Write-Host 'Connecting to SharePoint Online with MFA ..'
            $Parms = @{'url' = "https://$($global:Office365Tenant)-admin.sharepoint.com"; 'Region' = $global:SharePointRegion}
        }
        Else {
            Write-Host "Connecting to SharePoint Online using $($global:Office365Credentials.username) .."
            $Parms = @{'url' = "https://$($global:Office365Tenant)-admin.sharepoint.com"; 'Credential' = $global:Office365Credentials; 'Region' = $global:SharePointRegion}
        }
        Connect-SPOService @Parms
    }
    Else {
        Write-Error -Message 'Cannot connect to SharePoint Online - problem loading module.'
    }
}

Function global:Get-Office365Credentials {
    $global:Office365Credentials = $host.ui.PromptForCredential('Office 365 Credentials', 'Please enter your Office 365 credentials', '', '')
    If ( (Get-Module -Name 'Microsoft.Exchange.Management.ExoPowershellModule') -or (Get-Module -Name 'MicrosoftTeams') -or
        ((Get-Module -Name 'SkypeOnlineConnector' -ListAvailable) -and [System.Version]((Get-Module -Name 'SkypeOnlineConnector' -ListAvailable).Version.Build) -ge [System.Version]'7.0' ) -or
        ((Get-Module -Name 'Microsoft.Online.Sharepoint.PowerShell' -ListAvailable) -and [System.Version]((Get-Module -Name 'Microsoft.Online.Sharepoint.PowerShell' -ListAvailable).Version.Build) -ge [System.Version]'16.0' )) {
        $global:Office365CredentialsMFA = Get-MultiFactorAuthenticationUsage
    }
    Else {
        $global:Office365CredentialsMFA = $false
    }
}

Function global:Get-OnPremisesCredentials {
    $global:OnPremisesCredentials = $host.ui.PromptForCredential('On-Premises Credentials', 'Please Enter Your On-Premises Credentials', '', '')
}

Function global:Get-ExchangeOnPremisesFQDN {
    $global:ExchangeOnPremisesFQDN = Read-Host -Prompt 'Enter Exchange On-Premises endpoint, e.g. exchange1.contoso.com'
}

Function global:Get-Office365Tenant {
    $global:Office365Tenant = Read-Host -Prompt 'Enter tenant ID, e.g. contoso for contoso.onmicrosoft.com'
}

function global:Connect-Office365 {
    Connect-AzureActiveDirectory
    Connect-AzureRMS
    Connect-ExchangeOnline
    Connect-MSTeams
    Connect-SkypeOnline
    Connect-EOP
    Connect-ComplianceCenter
    Connect-SharePointOnline
}

# Initialize environment
Set-Office365Environment -AzureEnvironment 'Default'

#Scan for Exchange & SCC MFA PowerShell module presence
$local:ExchangeMFAModule = 'Microsoft.Exchange.Management.ExoPowershellModule'
$local:ExchangeADALModule = 'Microsoft.IdentityModel.Clients.ActiveDirectory'
$local:ModuleList = @(Get-ChildItem -Path "$($env:LOCALAPPDATA)\Apps\2.0" -Filter "$($local:ExchangeMFAModule).manifest" -Recurse ) | Sort LastWriteTime -Desc | Select -First 1
If ( $local:ModuleList) {
    $local:ModuleName = Join-path -Path $local:ModuleList[0].Directory.FullName -ChildPath "$($local:ExchangeMFAModule).dll"
    $local:ModuleVersion = (Get-Item -Path $local:ModuleName).VersionInfo.ProductVersion
    Write-Host "Exchange Multi-Factor Authentication PowerShell Module installed (version $($local:ModuleVersion))" -ForegroundColor Green
    if ( [System.Version]$local:ModuleVersion -lt [System.Version]$local:ExoPSSessionModuleVersion_Recommended) {
        Write-Host ('It is highly recommended to update the ExoPSSession module to version {0} or higher' -f $local:ExoPSSessionModuleVersion_Recommended) -ForegroundColor Red
    }
    Import-Module -FullyQualifiedName $local:ModuleName -Force
    $local:ModuleName = Join-path -Path $local:ModuleList[0].Directory.FullName -ChildPath "$($local:ExchangeADALModule).dll"
    If( Test-Path -Path $local:ModuleName) {
        $local:ModuleVersion = (Get-Item -Path $local:ModuleName).VersionInfo.FileVersion
        Write-Host "Exchange supporting ADAL module found (version $($local:ModuleVersion))" -ForegroundColor Green
        Add-Type -Path $local:ModuleName
    }
}
Else {
    Write-Verbose -Message 'Exchange Multi-Factor Authentication PowerShell Module is not installed.`nYou can download the module from EAC (Hybrid page) or via http://bit.ly/ExOPSModule'
}

ForEach ( $local:Function in $local:Functions) {
    $local:Item = ($local:Function).split('|')
    If ( !($local:Item[3]) -or ( Get-Module -Name $local:Item[3] -ListAvailable)) {
        If ( $local:CreateISEMenu) {
            $local:MenuObj = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus | Where-Object -FilterScript { $_.DisplayName -eq $local:Item[0] }
            If ( !( $local:MenuObj)) {
                Try {$local:MenuObj = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add( $local:Item[0], $null, $null)}
                Catch {Write-Warning -Message $_}
            }
            Try {
                $local:RemoveItems = $local:MenuObj.Submenus |  Where-Object -FilterScript { $_.DisplayName -eq $local:Item[1] -or $_.Action -eq $local:Item[2] }
                $null = $local:RemoveItems |
                    ForEach-Object -Process { $local:MenuObj.Submenus.Remove( $_) }
                $null = $local:MenuObj.SubMenus.Add( $local:Item[1], [ScriptBlock]::Create( $local:Item[2]), $null)
            }
            Catch {Write-Warning -Message $_}
        }
        If ( $local:Item[3]) {
            $local:Module = Get-Module -Name $local:Item[3] -ListAvailable
            $local:Version = ($local:Module).Version[0]
            Write-Host "$($local:Item[4]) module installed (version $($local:Version))" -ForegroundColor Green -NoNewline
            If ( $local:HasInternetAccess -and $local:OnlineModuleVersionChecks) {
                Try {
                    $OnlineModule = Find-Module -Name $local:Item[3] -Repository PSGallery -ErrorAction Stop
                    $outdated = [System.Version]$local:Version -lt [System.Version]$OnlineModule.version
                    If( $outdated -and $local:OnlineModuleAutoUpdate) {
                        $ThisPrincipal= new-object System.Security.principal.windowsprincipal( [System.Security.Principal.WindowsIdentity]::GetCurrent())
                        if( $ThisPrincipal.IsInRole("Administrators")) { 
                            Update-Module -Name $local:Item[3] -Repository PSGallery -ErrorAction Stop -Confirm:$false
                            Update-Module -Name $local:Item[3] -RequiredVersion $local:Version -Repository PSGallery -ErrorAction Stop -Confirm:$false
                            Write-Host ' UPDATED' -ForegroundColor YELLOW
                        }
			Else {
			    Write-Host ' OUTDATED' -ForegroundColor Red
			}
                    }
                    Write-Host ' (Current)' -ForegroundColor Green
                }
                Catch {
		    Write-Host ''
                }
            }
	    Else {
                # Check if we have a last known version
                If ( $local:Item[6]) {
                    $outdated = [System.Version]$local:Version -lt [System.Version]$local:item[6]
                    If( $outdated -and $local:OnlineModuleAutoUpdate) {
                        Write-Host ' OUTDATED' -ForegroundColor Red
		    }
                    Else {
                        Write-Host ''
                    }
                }
		Else {
                    Write-Host ''
                }
	    }
        }
    }
    Else {
        Write-Host "$($local:Item[4]) module not detected, link: $($local:Item[5])" -ForegroundColor Yellow
    }
}

