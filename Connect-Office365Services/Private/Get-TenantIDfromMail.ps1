function Get-TenantIDfromMail {
    param(
        [string]$mail
    )
    # Validate email structure and domain before using in URI
    $local:HostnamePattern = '^([a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$'
    $local:parts = $mail -split '@'
    If( $local:parts.Count -ne 2 -or [string]::IsNullOrEmpty( $local:parts[1])) {
        Write-Warning 'E-mail address invalid, cannot determine Tenant ID'
        return $null
    }
    $domainPart= $local:parts[1]
    If( $domainPart -notmatch $local:HostnamePattern) {
        Write-Warning 'E-mail address invalid, cannot determine Tenant ID'
        return $null
    }
    $res= $null
    Try {
        $res= (Invoke-RestMethod -Uri ('https://login.microsoftonline.com/{0}/v2.0/.well-known/openid-configuration' -f $domainPart) -ErrorAction Stop).jwks_uri.split('/')[3]
        If(!( $res)) {
            Write-Warning 'Could not determine Tenant ID using e-mail address'
        }
    }
    Catch {
        Write-Warning ('Could not determine Tenant ID: {0}' -f $_.Exception.Message)
    }
    return $res
}
