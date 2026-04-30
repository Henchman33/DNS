Get-WinDNSRecords -IncludeDetails -Prettify | Out-HtmlView -Filtering

<#
Get-WinDNSRecords -IncludeDetails -Prettify | Out-HtmlView -Filtering
NAME
    Get-WinDNSRecords

SYNOPSIS
    Gets all the DNS records from all the zones within a forest


SYNTAX
    Get-WinDNSRecords [[-IncludeZone] ] [[-ExcludeZone] ] [-IncludeDetails] [-Prettify] [-IncludeDNSRecords] [-AsHashtable] []


DESCRIPTION
    Gets all the DNS records from all the zones within a forest


PARAMETERS
    -IncludeZone
        Limit the output of DNS records to specific zones

    -ExcludeZone
        Limit the output of dNS records to only zones not in the exclude list

    -IncludeDetails []
        Adds additional information such as creation time, changed time

    -Prettify []
        Converts arrays into strings connected with comma

    -IncludeDNSRecords []
        Include full DNS records just in case one would like to further process them

    -AsHashtable []
        Outputs the results as a hashtable instead of an array


        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>Get-WinDNSRecords -Prettify -IncludeDetails | Format-Table


    -------------------------- EXAMPLE 2 --------------------------

    PS C:\>$Output = Get-WinDNSRecords -Prettify -IncludeDetails -Verbose

    $Output.Count
    $Output | Sort-Object -Property Count -Descending | Select-Object -First 30 | Format-Table


REMARKS
    To see the examples, type: "get-help Get-WinDNSRecords -examples".
    For more information, type: "get-help Get-WinDNSRecords -detailed".
    For technical information, type: "get-help Get-WinDNSRecords -full".

#>
