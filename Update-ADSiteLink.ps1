<#
    .SYNOPSIS
        The goal of this script is to update the AD Site Links according to information provided in a CSV file.

    .DESCRIPTION
        The goal of this script is to update the AD Site Links according to information provided in a CSV file.

    .PARAMETER AdSitesCsv
        Specifies the CSV file containing the list of DNS  servers to be updated
        The CSV file should be without headers

    .EXAMPLE
        ./Update-DnsForwarder.ps1 -DnsServersCsv ./DnsServers.csv -AddressToTest "www.bursky.net" -DnsForwarder 9.9.9.9

        This example will query all DNS servers from the CSV file DnsServers.csv, trying to resolve the address WWW.BURSKY.NET using the DNS forwarder 9.9.9.9

    .NOTES
        NAME:   Update-AdSiteLink.ps1
        AUTHOR: Peter Bursky
        DATE:   21/5/2020
        EMAIL:  peter@bursky.net

        REQUIREMENTS:
        -Administrator permissions

        VERSION HISTORY:
        0.1 2020.05.21
            Initial Version.
    .link
        https://github.com/PeterBursky/powershell-admin
#>

[CmdletBinding()]
PARAM (
    [Parameter(Mandatory = $true, HelpMessage = "You must specify the CSV file containg the AD Site Link details")]
    [String]$AdSitesCsv
)

Write-Verbose "Importing CSV file containg the AD Site Configurations"
$adSiteList = Import-Csv -Path $AdSitesCsv -Header LinkName,Site1,Site2,Cost
$csvFileName = [System.IO.Path]::GetFileNameWithoutExtension($AdSitesCsv) + "_output.csv"
$csvText = "LinkName,Result,Details"
$csvText | Out-File -FilePath $csvFileName

foreach ($adSite in $adSiteList) {
    Write-Verbose "Processing record: $adSite"
    try {
        $result = Get-ADReplicationSite -Identity $adSite.Site1
        $result = Get-ADReplicationSite -Identity $adSite.Site2
    } #TRY
    catch {
        Write-Verbose "AN error occurred:"
        Write-Verbose $_
    }
} #FOREACH
