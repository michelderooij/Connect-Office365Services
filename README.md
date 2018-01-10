# Connect-Office365Services

PowerShell script defining functions to connect to Office 365 online services
or Exchange On-Premises. Call manually or alternatively embed or call from $profile
(Shell or ISE) to make functions available in your session. If loaded from
PowerShell_ISE, menu items are defined for the functions. To surpress creation of
menu items, hold 'Shift' while Powershell ISE loads.

## Getting Started

After execution, the following helper functions will be available:

* Connect-AzureActiveDirectory	    Connects to Azure Active Directory
* Connect-AzureRMS           	    Connects to Azure Rights Management
* Connect-ExchangeOnline     	    Connects to Exchange Online
* Connect-SkypeOnline        	    Connects to Skype for Business Online
* Connect-EOP                	    Connects to Exchange Online Protection
* Connect-ComplianceCenter   	    Connects to Compliance Center
* Connect-SharePointOnline   	    Connects to SharePoint Online
* Connect-MSTeams                   Connects to Microsoft Teams
* Get-Office365Credentials    	    Gets Office 365 credentials
* Connect-ExchangeOnPremises 	    Connects to Exchange On-Premises
* Get-OnPremisesCredentials    	    Gets On-Premises credentials
* Get-ExchangeOnPremisesFQDNGets    FQDN for Exchange On-Premises
* Get-Office365Tenant		    Gets Office 365 tenant name
* Set-Office365Environment          Configures Uri's and region to use

### Prerequisites

Script requires PowerShell 3.0 and higher and any modules installed you would like to use for connecting to Office 365 workloads.

### Installing

Call the script ad-hoc, or load from PowerShell profile, e.g.

Store the script in the same location as your profile; default location is
$ENV:UserProfile\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1)

Then, create a PowerShell profile when you have not done so yet:

notepad $Profile

and insert a line to load the script (making helper functions available):

& (Join-Path $PSScriptRoot "Connect-Office365Services.ps1")

Next time you open PowerShell, the script should load. 

## Contributing

N/A

## Versioning

Initial version published on GitHub is 1.84. Changelog is contained in the script.

## Authors

* Michel de Rooij [initial work] https://github.com/michelderooij

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

## Acknowledgments

N/A
 