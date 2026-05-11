function Connect-Office365 {
    # Fixed: removed calls to Connect-AzureActiveDirectory and Connect-AzureRMS
    # which no longer exist in the current codebase
    Connect-ExchangeOnline
    Connect-MSTeams
    Connect-IPPSSession
    Connect-SharePointOnline
}
