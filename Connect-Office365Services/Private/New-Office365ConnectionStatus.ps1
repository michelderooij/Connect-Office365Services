function New-Office365ConnectionStatus {
    param(
        [Parameter(Mandatory)]
        [string]$Service,

        [Parameter(Mandatory)]
        [string]$StateKey,

        [Parameter(Mandatory)]
        [bool]$Connected,

        [string]$Detail,

        [string]$ErrorMessage,

        [switch]$IncludeSuccess
    )

    $script:myOffice365Services[$StateKey] = $Connected

    $local:status = [PSCustomObject][ordered]@{
        Service   = $Service
        Connected = $Connected
        Detail    = $Detail
        Error     = $ErrorMessage
    }

    if ($ErrorMessage) {
        Write-Error -Message $ErrorMessage
    }

    if (-not $Connected -or $ErrorMessage -or $IncludeSuccess) {
        $local:status
    }
}
