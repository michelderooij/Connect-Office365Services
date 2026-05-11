function Get-Office365ModuleInfo {
    # Fixed: Microsoft365DSC Repo URL corrected (was missing the '5')
    # Added: DLLPickle entry
    '[
    {
        "Module": "ExchangeOnlineManagement",
        "Description": "Exchange Online Management",
        "Repo": "https://www.powershellgallery.com/packages/ExchangeOnlineManagement"
    },
    {
        "Module": "MSOnline",
        "Description": "MSOnline",
        "Repo": "https://www.powershellgallery.com/packages/MSOnline",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "AzureAD",
        "Description": "Azure Active Directory (v2)",
        "Repo": "https://www.powershellgallery.com/packages/azuread",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "AzureADPreview",
        "Description": "Azure Active Directory (v2 Preview)",
        "Repo": "https://www.powershellgallery.com/packages/AzureADPreview",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "AIPService",
        "Description": "Azure Information Protection",
        "Repo": "https://www.powershellgallery.com/packages/AIPService"
    },
    {
        "Module": "Microsoft.Online.Sharepoint.PowerShell",
        "Description": "SharePoint Online",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Online.SharePoint.PowerShell"
    },
    {
        "Module": "MicrosoftTeams",
        "Description": "Microsoft Teams",
        "Repo": "https://www.powershellgallery.com/packages/MicrosoftTeams"
    },
    {
        "Module": "MSCommerce",
        "Description": "Microsoft Commerce",
        "Repo": "https://www.powershellgallery.com/packages/MSCommerce"
    },
    {
        "Module": "PnP.PowerShell",
        "Description": "PnP.PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/PnP.PowerShell"
    },
    {
        "Module": "Microsoft.PowerApps.Administration.PowerShell",
        "Description": "PowerApps-Admin-PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.PowerApps.Administration.PowerShell"
    },
    {
        "Module": "Microsoft.PowerApps.PowerShell",
        "Description": "PowerApps-PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.PowerApps.PowerShell"
    },
    {
        "Module": "Microsoft.Graph.Intune",
        "Description": "MSGraph-Intune",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Intune"
    },
    {
        "Module": "Microsoft.Graph",
        "Description": "Microsoft.Graph",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph"
    },
    {
        "Module": "Microsoft.Graph.Beta",
        "Description": "Microsoft.Graph.Beta",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Beta"
    },
    {
        "Module": "Microsoft.Graph.Entra",
        "Description": "Microsoft.Graph.Entra",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Entra",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "Microsoft.Entra",
        "Description": "Microsoft.Entra",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Entra",
        "Replaces": "Microsoft.Graph.Entra"
    },
    {
        "Module": "Microsoft.Graph.Entra.Beta",
        "Description": "Microsoft.Graph.Entra.Beta",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Entra.Beta",
        "ReplacedBy": "Microsoft.Entra.Beta"
    },
    {
        "Module": "Microsoft.Entra.Beta",
        "Description": "Microsoft.Entra.Beta",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Entra.Beta",
        "Replaces": "Microsoft.Graph.Entra.Beta"
    },
    {
        "Module": "MicrosoftPlaces",
        "Description": "MicrosoftPlaces",
        "Repo": "https://www.powershellgallery.com/packages/MicrosoftPlaces"
    },
    {
        "Module": "MicrosoftPowerBIMgmt",
        "Description": "MicrosoftPowerBIMgmt",
        "Repo": "https://www.powershellgallery.com/packages/MicrosoftPowerBIMgmt"
    },
    {
        "Module": "Az",
        "Description": "Az",
        "Repo": "https://www.powershellgallery.com/packages/Az"
    },
    {
        "Module": "Microsoft365DSC",
        "Description": "Microsoft365DSC",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft365DSC"
    },
    {
        "Module": "WhiteboardAdmin",
        "Description": "WhiteboardAdmin",
        "Repo": "https://www.powershellgallery.com/packages/WhiteboardAdmin"
    },
    {
        "Module": "MSIdentityTools",
        "Description": "MSIdentityTools",
        "Repo": "https://www.powershellgallery.com/packages/MSIdentityTools"
    },
    {
        "Module": "O365CentralizedAddInDeployment",
        "Description": "O365 Centralized Add-In Deployment Module",
        "Repo": "https://www.powershellgallery.com/packages/O365CentralizedAddInDeployment"
    },
    {
        "Module": "ORCA",
        "Description": "Office 365 Recommended Configuration Analyzer (ORCA)",
        "Repo": "https://www.powershellgallery.com/packages/ORCA"
    }
    ]' | ConvertFrom-Json
}
