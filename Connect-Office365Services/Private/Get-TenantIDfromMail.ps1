function Get-TenantIDfromMail {
    param(
        [string]$mail
    )
    $domainPart= ($mail -split '@')[1]
    If( $domainPart) {
        Try {
            $res= (Invoke-RestMethod -Uri ('https://login.microsoftonline.com/{0}/v2.0/.well-known/openid-configuration' -f $domainPart) -ErrorAction Stop).jwks_uri.split('/')[3]
            If(!( $res)) {
                Write-Warning 'Could not determine Tenant ID using e-mail address'
                $res= $null
            }
        }
        Catch {
            Write-Warning ('Could not determine Tenant ID: {0}' -f $_.Exception.Message)
            $res= $null
        }
    }
    Else {
        Write-Warning 'E-mail address invalid, cannot determine Tenant ID'
        $res= $null
    }
    return $res
}
