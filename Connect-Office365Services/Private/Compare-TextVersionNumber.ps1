# Compare-TextVersionNumber: version comparison similar to [System.Version]'s CompareTo method
# Returns: 1 = CompareTo is newer, 0 = Equal, -1 = Version is newer
#
# Prerelease handling:
#   A '-tag' suffix (e.g. -Preview, -Preview1, -alpha, -beta2) marks a prerelease.
#   Any prerelease is always older than the stable release of the same base version.
#   When a prerelease has a numeric suffix the number is used for ordering within
#   the same base version (Preview1 < Preview2).  A tag with no trailing number
#   (e.g. plain "-Preview") is treated as ordinal 0, so:
#       3.10.0-Preview < 3.10.0-Preview1 < 3.10.0-Preview2 < 3.10.0
function Compare-TextVersionNumber {
    param(
        [string]$Version,
        [string]$CompareTo
    )
    $res = 0

    $null = $Version -match '^(?<version>[\d\.]+)(-(?<tag>[a-zA-Z]+)(?<preview>\d*))?$'
    $VersionVer = [System.Version]($Matches['version'])
    if ( $Matches.ContainsKey('tag')) {
        # Prerelease: use the numeric suffix as ordinal; absent number → 0
        $local:n = if ($Matches['preview']) { $Matches['preview'] } else { '0' }
        $VersionPreviewVer = [System.Version]('{0}.0' -f $local:n)
    }
    else {
        # Stable release ranks above any prerelease of the same base version
        $VersionPreviewVer = [System.Version]'99999.99999'
    }

    $null = $CompareTo -match '^(?<version>[\d\.]+)(-(?<tag>[a-zA-Z]+)(?<preview>\d*))?$'
    $CompareToVer = [System.Version]($Matches['version'])
    if ( $Matches.ContainsKey('tag')) {
        $local:n = if ($Matches['preview']) { $Matches['preview'] } else { '0' }
        $CompareToPreviewVer = [System.Version]('{0}.0' -f $local:n)
    }
    else {
        $CompareToPreviewVer = [System.Version]'99999.99999'
    }

    if ( $VersionVer -gt $CompareToVer) {
        $res = -1
    }
    else {
        if ( $VersionVer -lt $CompareToVer) {
            $res = 1
        }
        else {
            # Equal base version — order by prerelease ordinal
            if ( $VersionPreviewVer -gt $CompareToPreviewVer) {
                $res = -1
            }
            else {
                if ( $VersionPreviewVer -lt $CompareToPreviewVer) {
                    $res = 1
                }
                else {
                    $res = 0
                }
            }
        }
    }
    $res
}
