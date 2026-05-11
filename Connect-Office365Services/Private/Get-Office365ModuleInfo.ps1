function Get-Office365ModuleInfo {
    '[
    {
        "Module": "ExchangeOnlineManagement",
        "Description": "Exchange Online Management",
        "Repo": "https://www.powershellgallery.com/packages/ExchangeOnlineManagement"
    },
    {
        "Module": "MSOnline",
        "Description": "Azure Active Directory (v1)",
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
        "Module": "Microsoft.Online.SharePoint.PowerShell",
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
        "Description": "Microsoft 365 Patterns and Practices (PnP)",
        "Repo": "https://www.powershellgallery.com/packages/PnP.PowerShell"
    },
    {
        "Module": "Microsoft.PowerApps.Administration.PowerShell",
        "Description": "PowerApps and Flow Administration",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.PowerApps.Administration.PowerShell"
    },
    {
        "Module": "Microsoft.PowerApps.PowerShell",
        "Description": "PowerApps and Flow",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.PowerApps.PowerShell"
    },
    {
        "Module": "Microsoft.Graph.Intune",
        "Description": "Microsoft Intune Graph API",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Intune"
    },
    {
        "Module": "Microsoft.Graph",
        "Description": "Microsoft Graph PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph"
    },
    {
        "Module": "Microsoft.Graph.Beta",
        "Description": "Microsoft Graph PowerShell Beta",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Beta"
    },
    {
        "Module": "Microsoft.Graph.Entra",
        "Description": "Microsoft Graph Entra (deprecated)",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Entra",
        "ReplacedBy": "Microsoft.Entra"
    },
    {
        "Module": "Microsoft.Entra",
        "Description": "Microsoft Entra PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Entra",
        "Replaces": "Microsoft.Graph.Entra"
    },
    {
        "Module": "Microsoft.Graph.Entra.Beta",
        "Description": "Microsoft Graph Entra Beta (deprecated)",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Graph.Entra.Beta",
        "ReplacedBy": "Microsoft.Entra.Beta"
    },
    {
        "Module": "Microsoft.Entra.Beta",
        "Description": "Microsoft Entra PowerShell Beta",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft.Entra.Beta",
        "Replaces": "Microsoft.Graph.Entra.Beta"
    },
    {
        "Module": "MicrosoftPlaces",
        "Description": "Microsoft Places",
        "Repo": "https://www.powershellgallery.com/packages/MicrosoftPlaces"
    },
    {
        "Module": "MicrosoftPowerBIMgmt",
        "Description": "Microsoft Power BI",
        "Repo": "https://www.powershellgallery.com/packages/MicrosoftPowerBIMgmt"
    },
    {
        "Module": "Az",
        "Description": "Azure PowerShell",
        "Repo": "https://www.powershellgallery.com/packages/Az"
    },
    {
        "Module": "Microsoft365DSC",
        "Description": "Microsoft 365 Desired State Configuration",
        "Repo": "https://www.powershellgallery.com/packages/Microsoft365DSC"
    },
    {
        "Module": "WhiteboardAdmin",
        "Description": "Microsoft Whiteboard Administration",
        "Repo": "https://www.powershellgallery.com/packages/WhiteboardAdmin"
    },
    {
        "Module": "MSIdentityTools",
        "Description": "Microsoft Identity Tools",
        "Repo": "https://www.powershellgallery.com/packages/MSIdentityTools"
    },
    {
        "Module": "O365CentralizedAddInDeployment",
        "Description": "O365 Centralized Add-In Deployment Module",
        "Repo": "https://www.powershellgallery.com/packages/O365CentralizedAddInDeployment"
    },
    {
        "Module": "ORCA",
        "Description": "Microsoft Defender for Office 365 Recommended Configuration Analyzer (ORCA)",
        "Repo": "https://www.powershellgallery.com/packages/ORCA"
    }
    ]' | ConvertFrom-Json
}
