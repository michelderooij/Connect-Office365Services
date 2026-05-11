# Compare-TextVersionNumber: version comparison similar to [System.Version]'s CompareTo method
# Returns: 1 = CompareTo is newer, 0 = Equal, -1 = Version is newer
function Compare-TextVersionNumber {
    param(
        [string]$Version,
        [string]$CompareTo
    )
    $res= 0
    $null= $Version -match '^(?<version>[\d\.]+)(\-)?([a-zA-Z]*(?<preview>[\d]*))?$'
    $VersionVer= [System.Version]($matches.Version)
    If( $matches.Preview) {
        # Suffix .0 to satisfy [System.Version] as '#' alone won't initialize
        $VersionPreviewVer= [System.Version]('{0}.0' -f $matches.Preview)
    }
    Else {
        $VersionPreviewVer= [System.Version]'99999.99999'
    }
    $null= $CompareTo -match '^(?<version>[\d\.]+)(\-)?([a-zA-Z]*(?<preview>[\d]*))?$'
    $CompareToVer= [System.Version]($matches.Version)
    If( $matches.Preview) {
        $CompareToPreviewVer= [System.Version]('{0}.0' -f $matches.Preview)
    }
    Else {
        $CompareToPreviewVer= [System.Version]'99999.99999'
    }

    If( $VersionVer -gt $CompareToVer) {
        $res= -1
    }
    Else {
        If( $VersionVer -lt $CompareToVer) {
            $res= 1
        }
        Else {
            # Equal base version — check preview tag
            If( $VersionPreviewVer -gt $CompareToPreviewVer) {
                $res= -1
            }
            Else {
                If( $VersionPreviewVer -lt $CompareToPreviewVer) {
                    $res= 1
                }
                Else {
                    # Truly equal
                    $res= 0
                }
            }
        }
    }
    $res
}
