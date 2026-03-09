#Requires -Version 5.1
<#
.SYNOPSIS
    AD DNS Report - Enterprise DNS Documentation Script
    Created by: Stephen McKee - Server Administrator 2 - IGT - Everi

.DESCRIPTION
    Comprehensively documents ALL DNS configuration across every Domain Controller
    in the Active Directory environment, including:

      - AD Domain & DC Inventory
      - DNS Server Settings & Configuration (per server)
      - DNS Server Recursion & Cache Settings
      - Root Hints
      - DNS Forwarders & Conditional Forwarders
      - All DNS Zones (Primary, Secondary, Stub, Forward Lookup, Reverse Lookup)
      - Zone Aging & Scavenging Configuration
      - Zone Transfer Settings
      - DNS Resource Records (A, AAAA, CNAME, MX, PTR, SOA, SRV, NS, TXT, and more)
      - DNS Record Summary by Type
      - AD-Integrated DNS Partitions
      - DNS Server Event Log Summary

    Exports to:
      - Multiple CSV files (one per category)
      - Multi-tab XLSX workbook
      - Enterprise-grade dark-themed searchable HTML report

.NOTES
    Requirements:
      - PowerShell 5.1+
      - RSAT: DNS Server Tools
      - RSAT: Active Directory Domain Services Tools
      - ImportExcel module (auto-installed from PSGallery if missing)
      - Run as Administrator or Domain Admin
#>

#region --- INITIALIZATION ---
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"
$SkipXLSX = $false

$ScriptTitle   = "AD DNS Report"
$ScriptAuthor  = "Stephen McKee - Server Administrator 2 - IGT - Everi"
$RunDate       = Get-Date
$DateTimeStamp = $RunDate.ToString("yyyy-MM-dd_HH-mm-ss")
$DateDisplay   = $RunDate.ToString("MMMM dd, yyyy hh:mm:ss tt")

$DesktopPath  = [Environment]::GetFolderPath("Desktop")
$OutputFolder = Join-Path $DesktopPath "AD_DNS_Report_$DateTimeStamp"
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  $ScriptTitle" -ForegroundColor Yellow
Write-Host "  $ScriptAuthor" -ForegroundColor Yellow
Write-Host "  Started : $DateDisplay" -ForegroundColor Gray
Write-Host "  Output  : $OutputFolder" -ForegroundColor Gray
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Auto-install ImportExcel
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[*] ImportExcel not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[+] ImportExcel installed." -ForegroundColor Green
    } catch {
        Write-Warning "ImportExcel install failed. XLSX will be skipped. Error: $_"
        $SkipXLSX = $true
    }
}
Import-Module ImportExcel -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# Safe HTML encoder - falls back to manual replace if System.Web not available
function Safe-HtmlEncode {
    param([string]$str)
    try {
        return [System.Web.HttpUtility]::HtmlEncode($str)
    } catch {
        return $str -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'","&#39;"
    }
}
#endregion

#region --- HELPER FUNCTIONS ---
function Write-Section { param([string]$m) Write-Host ""; Write-Host "[>>] $m" -ForegroundColor Cyan }
function Write-OK      { param([string]$m) Write-Host "  [+] $m" -ForegroundColor Green }
function Write-Warn    { param([string]$m) Write-Host "  [!] $m" -ForegroundColor Yellow }
function Write-Fail    { param([string]$m) Write-Host "  [-] $m" -ForegroundColor Red }
function Write-Info    { param([string]$m) Write-Host "  [i] $m" -ForegroundColor Gray }

function Safe-Export-CSV {
    param([object[]]$Data, [string]$Path, [string]$Label)
    if ($Data -and $Data.Count -gt 0) {
        $Data | Export-Csv -Path $Path -NoTypeInformation -Force
        Write-OK "$Label -> CSV ($($Data.Count) rows)"
    } else {
        Write-Warn "$Label - No data"
    }
}

function ConvertTo-HtmlTable {
    param([object[]]$Data, [string]$TableId, [string]$Caption)
    if (-not $Data -or $Data.Count -eq 0) {
        return "<p class='no-data'>No data available.</p>"
    }
    $headers = $Data[0].PSObject.Properties.Name
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<div class='table-wrapper'>")
    [void]$sb.Append("<input type='text' class='search-box' placeholder='Search $Caption...' onkeyup=""filterTable(this,'$TableId')"" />")
    [void]$sb.Append("<div class='table-scroll'><table id='$TableId' class='data-table'>")
    [void]$sb.Append("<thead><tr>")
    $idx = 0
    foreach ($h in $headers) {
        [void]$sb.Append("<th onclick=""sortTable('$TableId',$idx)"">$h <span class='sort-icon'>&#8645;</span></th>")
        $idx++
    }
    [void]$sb.Append("</tr></thead><tbody>")
    foreach ($row in $Data) {
        [void]$sb.Append("<tr>")
        foreach ($h in $headers) {
            $val = $row.$h; if ($null -eq $val) { $val = "" }
            $vs = Safe-HtmlEncode($val.ToString())
            $cls = ""
            if ($h -eq "ZoneType")     { if ($val -eq "Primary") { $cls=" class='z-primary'" } elseif ($val -eq "Secondary") { $cls=" class='z-secondary'" } elseif ($val -eq "Stub") { $cls=" class='z-stub'" } elseif ($val -eq "Forwarder") { $cls=" class='z-forwarder'" } }
            if ($h -eq "RecordType")   { $cls=" class='rt-$($val.ToLower() -replace '[^a-z]','')'" }
            if ($h -eq "IsDsIntegrated" -or $h -eq "IsADIntegrated") { if ($val -eq "True") { $cls=" class='bool-yes'" } else { $cls=" class='bool-no'" } }
            if ($h -eq "DynamicUpdate"){ if ($val -eq "None") { $cls=" class='bool-no'" } elseif ($val -like "*Secure*") { $cls=" class='bool-yes'" } }
            if ($h -eq "ZoneState" -or $h -eq "Status") { if ($val -eq "Running") { $cls=" class='bool-yes'" } else { $cls=" class='bool-no'" } }
            [void]$sb.Append("<td$cls>$vs</td>")
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</tbody></table></div></div>")
    return $sb.ToString()
}
#endregion

#region --- AD DOMAIN & DC INVENTORY ---
Write-Section "Collecting Active Directory Domain & DC Information"

$DomainInfo        = $null
$ForestInfo        = $null
$DomainControllers = @()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $DomainInfo        = Get-ADDomain  -ErrorAction Stop
    $ForestInfo        = Get-ADForest  -ErrorAction Stop
    $DomainControllers = Get-ADDomainController -Filter * -ErrorAction Stop | Sort-Object HostName
    Write-OK "Domain: $($DomainInfo.DNSRoot) | Forest: $($ForestInfo.Name) | DCs: $($DomainControllers.Count)"
} catch {
    Write-Fail "AD query failed: $_"
    $DomainControllers = @([PSCustomObject]@{
        HostName = $env:COMPUTERNAME; Site = "Unknown"; IPv4Address = $env:COMPUTERNAME
        IsGlobalCatalog = $false; OperationMasterRoles = @(); IsReadOnly = $false
        OperatingSystem = "Unknown"; OperatingSystemVersion = "Unknown"
    })
}

$ADSummaryData = if ($DomainInfo) {
    [PSCustomObject]@{
        DomainName             = $DomainInfo.DNSRoot
        NetBIOSName            = $DomainInfo.NetBIOSName
        ForestName             = $ForestInfo.Name
        ForestMode             = $ForestInfo.ForestMode
        DomainMode             = $DomainInfo.DomainMode
        PDCEmulator            = $DomainInfo.PDCEmulator
        RIDMaster              = $DomainInfo.RIDMaster
        InfrastructureMaster   = $DomainInfo.InfrastructureMaster
        SchemaMaster           = $ForestInfo.SchemaMaster
        DomainNamingMaster     = $ForestInfo.DomainNamingMaster
        TotalDomainControllers = $DomainControllers.Count
        DNSServers             = ($DomainControllers.HostName -join "; ")
        ReportGeneratedBy      = $env:USERNAME
        ReportDate             = $DateDisplay
    }
} else { @() }

$DCDetailData = $DomainControllers | ForEach-Object {
    [PSCustomObject]@{
        Hostname               = $_.HostName
        Site                   = $_.Site
        IPv4Address            = $_.IPv4Address
        IsGlobalCatalog        = $_.IsGlobalCatalog
        IsReadOnly             = $_.IsReadOnly
        OperationMasterRoles   = ($_.OperationMasterRoles -join "; ")
        OperatingSystem        = $_.OperatingSystem
        OSVersion              = $_.OperatingSystemVersion
    }
}

Safe-Export-CSV -Data @($ADSummaryData) -Path "$OutputFolder\01_AD_Domain_Summary.csv"  -Label "AD Domain Summary"
Safe-Export-CSV -Data $DCDetailData     -Path "$OutputFolder\02_Domain_Controllers.csv" -Label "Domain Controllers"
#endregion

#region --- DNS SERVER SETTINGS ---
Write-Section "Collecting DNS Server Settings & Configuration"

$AllDNSServerSettings   = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDNSServerCache      = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllRootHints           = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDNSForwarders       = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllCondForwarders      = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDNSZones            = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllZoneAging           = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllZoneTransfer        = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDNSRecords          = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDNSPartitions       = [System.Collections.Generic.List[PSCustomObject]]::new()

$DNSServers = $DomainControllers | Select-Object -ExpandProperty HostName

foreach ($dnsServer in $DNSServers) {
    Write-Info "Processing DNS Server: $dnsServer"

    #-- Server-level Settings
    try {
        $cfg = Get-DnsServer -ComputerName $dnsServer -ErrorAction Stop
        $s   = $cfg.ServerSetting

        $AllDNSServerSettings.Add([PSCustomObject]@{
            Server                        = $dnsServer
            ListeningIPAddresses          = ($s.ListeningIPAddress -join "; ")
            AllIPAddress                  = $s.AllIPAddress
            DisableRecursion              = $s.DisableRecursion
            RecursionTimeout              = $s.RecursionTimeout
            RecursionRetry                = $s.RecursionRetry
            MaxCacheTTL                   = $s.MaxCacheTTL
            MaxNegativeCacheTTL           = $s.MaxNegativeCacheTTL
            ForwardingTimeout             = $s.ForwardingTimeout
            IsSlave                       = $s.IsSlave
            EnableDnsSec                  = $s.EnableDnsSec
            DnssecKeymasterZone           = $s.DnssecKeymasterZone
            SendPort                      = $s.SendPort
            BindSecondaries               = $s.BindSecondaries
            BootMethod                    = $s.BootMethod
            LocalNetPriority              = $s.LocalNetPriority
            LocalNetPriorityMask          = $s.LocalNetPriorityMask
            EnableEventLog                = $s.EnableEventLog
            LogLevel                      = $s.LogLevel
            EventLogLevel                 = $s.EventLogLevel
            AutoConfigFileZones           = $s.AutoConfigFileZones
            AutoCreateDelegation          = $s.AutoCreateDelegation
            AddressAnswerLimit            = $s.AddressAnswerLimit
            DefaultAgingState             = $s.DefaultAgingState
            DefaultNoRefreshInterval      = $s.DefaultNoRefreshInterval
            DefaultRefreshInterval        = $s.DefaultRefreshInterval
            DeleteOutsideGlue             = $s.DeleteOutsideGlue
            EnableDirectoryPartitions     = $s.EnableDirectoryPartitions
            EnableDuplicateQuerySuppression = $s.EnableDuplicateQuerySuppression
            EnableIPv6                    = $s.EnableIPv6
            EnableOnlineSigning           = $s.EnableOnlineSigning
            IgnoreServerLevelPolicies     = $s.IgnoreServerLevelPolicies
            IgnoreAllPolicies             = $s.IgnoreAllPolicies
            LameDelegationTTL             = $s.LameDelegationTTL
            MaxResourceRecordsInNonSecureUpdate = $s.MaxResourceRecordsInNonSecureUpdate
            NameCheckFlag                 = $s.NameCheckFlag
            NoUpdateDelegations           = $s.NoUpdateDelegations
            OpenAclOnProxyUpdates         = $s.OpenAclOnProxyUpdates
            QuietRecvFaultInterval        = $s.QuietRecvFaultInterval
            QuietRecvLogInterval          = $s.QuietRecvLogInterval
            RoundRobin                    = $s.RoundRobin
            RpcProtocol                   = $s.RpcProtocol
            SecureResponses               = $s.SecureResponses
            SelfTest                      = $s.SelfTest
            SocketPoolExcludedPortRanges  = ($s.SocketPoolExcludedPortRanges -join "; ")
            SocketPoolSize                = $s.SocketPoolSize
            StrictFileParsing             = $s.StrictFileParsing
            SyncDsZoneSerial              = $s.SyncDsZoneSerial
            UpdateOptions                 = $s.UpdateOptions
            Version                       = $s.Version
            WriteAuthorityNS              = $s.WriteAuthorityNS
            XfrConnectTimeout             = $s.XfrConnectTimeout
        })
        Write-OK "  Server settings collected"
    } catch { Write-Warn "  Server settings failed on $dnsServer : $_" }

    #-- Cache Settings
    try {
        $cache = Get-DnsServerCache -ComputerName $dnsServer -ErrorAction Stop
        $AllDNSServerCache.Add([PSCustomObject]@{
            Server                   = $dnsServer
            EnablePollutionProtection = $cache.EnablePollutionProtection
            LockingPercent           = $cache.LockingPercent
            MaxKBSize                = $cache.MaxKBSize
            MaxNegativeTtl           = $cache.MaxNegativeTtl
            MaxTtl                   = $cache.MaxTtl
            StoreEmptyAuthenticationResponse = $cache.StoreEmptyAuthenticationResponse
        })
        Write-OK "  Cache settings collected"
    } catch { Write-Warn "  Cache settings failed on $dnsServer : $_" }

    #-- Root Hints
    try {
        $rootHints = Get-DnsServerRootHint -ComputerName $dnsServer -ErrorAction Stop
        foreach ($rh in $rootHints.NameServer) {
            $AllRootHints.Add([PSCustomObject]@{
                Server     = $dnsServer
                NameServer = $rh.RecordData.NameServer
                IPAddress  = ($rootHints.IPAddress | Where-Object { $_.HostName -eq $rh.HostName } | ForEach-Object { $_.RecordData.IPv4Address.IPAddressToString } | Select-Object -First 1)
            })
        }
        Write-OK "  Root hints collected ($($rootHints.NameServer.Count))"
    } catch { Write-Warn "  Root hints failed on $dnsServer : $_" }

    #-- Forwarders
    try {
        $fwd = Get-DnsServerForwarder -ComputerName $dnsServer -ErrorAction Stop
        foreach ($ip in $fwd.IPAddress) {
            $AllDNSForwarders.Add([PSCustomObject]@{
                Server           = $dnsServer
                ForwarderIP      = $ip.ToString()
                UseRootHint      = $fwd.UseRootHint
                Timeout_sec      = $fwd.Timeout
                EnableReordering = $fwd.EnableReordering
            })
        }
        Write-OK "  Forwarders collected ($($fwd.IPAddress.Count))"
    } catch { Write-Warn "  Forwarders failed on $dnsServer : $_" }

    #-- Conditional Forwarders
    try {
        $condFwd = Get-DnsServerZone -ComputerName $dnsServer -ErrorAction Stop | Where-Object { $_.ZoneType -eq "Forwarder" }
        foreach ($cf in $condFwd) {
            $AllCondForwarders.Add([PSCustomObject]@{
                Server              = $dnsServer
                ZoneName            = $cf.ZoneName
                MasterServers       = ($cf.MasterServers -join "; ")
                ForwarderTimeout    = $cf.ForwarderTimeout
                IsDsIntegrated      = $cf.IsDsIntegrated
                ReplicationScope    = $cf.ReplicationScope
                DirectoryPartition  = $cf.DirectoryPartitionName
                UseRecursion        = $cf.UseRecursion
            })
        }
        Write-OK "  Conditional forwarders collected ($($condFwd.Count))"
    } catch { Write-Warn "  Conditional forwarders failed on $dnsServer : $_" }

    #-- Directory Partitions
    try {
        $parts = Get-DnsServerDirectoryPartition -ComputerName $dnsServer -ErrorAction Stop
        foreach ($p in $parts) {
            $AllDNSPartitions.Add([PSCustomObject]@{
                Server                  = $dnsServer
                DirectoryPartitionName  = $p.DirectoryPartitionName
                Flags                   = $p.Flags
                State                   = $p.State
                ZoneCount               = $p.ZoneCount
            })
        }
        Write-OK "  Directory partitions collected ($($parts.Count))"
    } catch { Write-Warn "  Directory partitions failed on $dnsServer : $_" }

    #-- DNS Zones
    try {
        $zones = Get-DnsServerZone -ComputerName $dnsServer -ErrorAction Stop | Where-Object { $_.ZoneType -ne "Forwarder" }
        Write-OK "  DNS Zones: $($zones.Count)"

        foreach ($zone in $zones) {
            $AllDNSZones.Add([PSCustomObject]@{
                Server                  = $dnsServer
                ZoneName                = $zone.ZoneName
                ZoneType                = $zone.ZoneType
                IsDsIntegrated          = $zone.IsDsIntegrated
                IsReverseLookupZone     = $zone.IsReverseLookupZone
                IsAutoCreated           = $zone.IsAutoCreated
                IsPaused                = $zone.IsPaused
                IsShutdown              = $zone.IsShutdown
                IsReadOnly              = $zone.IsReadOnly
                IsSigned                = $zone.IsSigned
                IsWinsEnabled           = $zone.IsWinsEnabled
                DynamicUpdate           = $zone.DynamicUpdate
                ReplicationScope        = $zone.ReplicationScope
                DirectoryPartitionName  = $zone.DirectoryPartitionName
                ZoneFile                = $zone.ZoneFile
                MasterServers           = ($zone.MasterServers -join "; ")
                NotifyServers           = ($zone.NotifyServers -join "; ")
                SecureSecondaries       = $zone.SecureSecondaries
                SecondaryServers        = ($zone.SecondaryServers -join "; ")
                UseWins                 = $zone.UseWins
                UseNbstat               = $zone.UseNbstat
            })

            #-- Zone Aging/Scavenging
            try {
                $aging = Get-DnsServerZoneAging -ComputerName $dnsServer -ZoneName $zone.ZoneName -ErrorAction Stop
                $AllZoneAging.Add([PSCustomObject]@{
                    Server             = $dnsServer
                    ZoneName           = $zone.ZoneName
                    AgingEnabled       = $aging.AgingEnabled
                    NoRefreshInterval  = $aging.NoRefreshInterval
                    RefreshInterval    = $aging.RefreshInterval
                    ScavengeServers    = ($aging.ScavengeServers -join "; ")
                    AvailForScavengeTime = $aging.AvailForScavengeTime
                })
            } catch {}

            #-- Zone Transfer
            try {
                $xfr = Get-DnsServerZoneTransfer -ComputerName $dnsServer -Name $zone.ZoneName -ErrorAction Stop
                $AllZoneTransfer.Add([PSCustomObject]@{
                    Server              = $dnsServer
                    ZoneName            = $zone.ZoneName
                    SecureSecondaries   = $xfr.SecureSecondaries
                    SecondaryServers    = ($xfr.SecondaryServers -join "; ")
                    NotifyServers       = ($xfr.NotifyServers -join "; ")
                    Notify              = $xfr.Notify
                })
            } catch {}

            #-- DNS Records (only collect from primary zones on first DC to avoid massive duplication)
            if ($zone.ZoneType -in @("Primary","Stub") -and $dnsServer -eq $DNSServers[0]) {
                try {
                    $records = Get-DnsServerResourceRecord -ComputerName $dnsServer -ZoneName $zone.ZoneName -ErrorAction Stop
                    foreach ($rec in $records) {
                        $recData = ""
                        try {
                            switch ($rec.RecordType) {
                                "A"     { $recData = $rec.RecordData.IPv4Address.IPAddressToString }
                                "AAAA"  { $recData = $rec.RecordData.IPv6Address.IPAddressToString }
                                "CNAME" { $recData = $rec.RecordData.HostNameAlias }
                                "MX"    { $recData = "$($rec.RecordData.MailExchange) [Pref=$($rec.RecordData.Preference)]" }
                                "NS"    { $recData = $rec.RecordData.NameServer }
                                "PTR"   { $recData = $rec.RecordData.PtrDomainName }
                                "SOA"   { $recData = "PrimaryNS=$($rec.RecordData.PrimaryServer); Email=$($rec.RecordData.ResponsiblePerson); Serial=$($rec.RecordData.SerialNumber); Refresh=$($rec.RecordData.RefreshInterval); Retry=$($rec.RecordData.RetryDelay); Expire=$($rec.RecordData.ExpireLimit); MinTTL=$($rec.RecordData.MinimumTimeToLive)" }
                                "SRV"   { $recData = "$($rec.RecordData.DomainName):$($rec.RecordData.Port) [Pri=$($rec.RecordData.Priority) Wt=$($rec.RecordData.Weight)]" }
                                "TXT"   { $recData = ($rec.RecordData.DescriptiveText -join " ") }
                                "WINS"  { $recData = "WINS Servers: $($rec.RecordData.WinsServers -join ', ')" }
                                "WINSR" { $recData = "ResultDomain: $($rec.RecordData.ResultDomain)" }
                                default {
                                    $recData = try { ($rec.RecordData | Out-String -Width 300).Trim().Split("`n")[0].Trim() } catch { "N/A" }
                                }
                            }
                        } catch { $recData = "ParseError" }

                        $fqdn = if ($rec.HostName -eq "@") { $zone.ZoneName } elseif ($rec.HostName -like "*.") { $rec.HostName.TrimEnd('.') } else { "$($rec.HostName).$($zone.ZoneName)" }

                        $AllDNSRecords.Add([PSCustomObject]@{
                            Zone            = $zone.ZoneName
                            ZoneType        = $zone.ZoneType
                            IsReverseZone   = $zone.IsReverseLookupZone
                            RecordName      = $rec.HostName
                            FQDN            = $fqdn
                            RecordType      = $rec.RecordType
                            TTL             = $rec.TimeToLive.ToString()
                            Data            = $recData
                            Timestamp       = if ($rec.Timestamp) { $rec.Timestamp.ToString() } else { "Static" }
                            IsAgeingEnabled = if ($rec.Timestamp) { "Yes" } else { "No" }
                            Server          = $dnsServer
                        })
                    }
                    Write-Info "    Zone '$($zone.ZoneName)' - $($records.Count) records"
                } catch { Write-Warn "  Records failed for zone $($zone.ZoneName) : $_" }
            }
        }
    } catch { Write-Fail "  Zones failed on $dnsServer : $_" }
}

#-- DNS Record Type Summary
$RecordTypeSummary = $AllDNSRecords | Group-Object RecordType | Sort-Object Count -Descending | ForEach-Object {
    [PSCustomObject]@{
        RecordType = $_.Name
        Count      = $_.Count
        Zones      = (($_.Group | Select-Object -ExpandProperty Zone -Unique | Sort-Object) -join "; ")
    }
}

#-- Zone Summary
$ZoneSummary = $AllDNSZones | Select-Object ZoneName, ZoneType, IsDsIntegrated, IsReverseLookupZone,
    DynamicUpdate, ReplicationScope, IsPaused, IsShutdown -Unique | Sort-Object ZoneName

Write-Section "Exporting CSV Files"
Safe-Export-CSV -Data $AllDNSServerSettings.ToArray() -Path "$OutputFolder\03_DNS_Server_Settings.csv"    -Label "DNS Server Settings"
Safe-Export-CSV -Data $AllDNSServerCache.ToArray()    -Path "$OutputFolder\04_DNS_Cache_Settings.csv"    -Label "DNS Cache Settings"
Safe-Export-CSV -Data $AllRootHints.ToArray()         -Path "$OutputFolder\05_DNS_Root_Hints.csv"        -Label "DNS Root Hints"
Safe-Export-CSV -Data $AllDNSForwarders.ToArray()     -Path "$OutputFolder\06_DNS_Forwarders.csv"        -Label "DNS Forwarders"
Safe-Export-CSV -Data $AllCondForwarders.ToArray()    -Path "$OutputFolder\07_DNS_Conditional_Fwds.csv"  -Label "Conditional Forwarders"
Safe-Export-CSV -Data $AllDNSPartitions.ToArray()     -Path "$OutputFolder\08_DNS_Partitions.csv"        -Label "DNS AD Partitions"
Safe-Export-CSV -Data $AllDNSZones.ToArray()          -Path "$OutputFolder\09_DNS_Zones.csv"             -Label "DNS Zones"
Safe-Export-CSV -Data $ZoneSummary                    -Path "$OutputFolder\10_DNS_Zone_Summary.csv"      -Label "Zone Summary"
Safe-Export-CSV -Data $AllZoneAging.ToArray()         -Path "$OutputFolder\11_DNS_Zone_Aging.csv"        -Label "Zone Aging/Scavenging"
Safe-Export-CSV -Data $AllZoneTransfer.ToArray()      -Path "$OutputFolder\12_DNS_Zone_Transfer.csv"     -Label "Zone Transfer Settings"
Safe-Export-CSV -Data $AllDNSRecords.ToArray()        -Path "$OutputFolder\13_DNS_Records_All.csv"       -Label "All DNS Records"
Safe-Export-CSV -Data $RecordTypeSummary              -Path "$OutputFolder\14_DNS_Record_Type_Summary.csv" -Label "Record Type Summary"
#endregion

#region --- XLSX EXPORT ---
if (-not $SkipXLSX) {
    Write-Section "Building Excel Workbook (.xlsx)"
    $XLSXPath = "$OutputFolder\AD_DNS_Report_$DateTimeStamp.xlsx"
    $xlP = @{ Path = $XLSXPath; AutoSize = $true; FreezeTopRow = $true; BoldTopRow = $true; TableStyle = "Medium9" }

    function Add-Sheet {
        param($Data, [string]$Sheet)
        if ($Data -and @($Data).Count -gt 0) {
            @($Data) | Export-Excel @xlP -WorksheetName $Sheet -Append
            Write-OK "Sheet '$Sheet' ($(@($Data).Count) rows)"
        } else { Write-Warn "Sheet '$Sheet' - no data" }
    }

    Add-Sheet -Data @($ADSummaryData)             -Sheet "AD_Summary"
    Add-Sheet -Data $DCDetailData                 -Sheet "Domain_Controllers"
    Add-Sheet -Data $AllDNSServerSettings.ToArray()-Sheet "DNS_Server_Settings"
    Add-Sheet -Data $AllDNSServerCache.ToArray()  -Sheet "DNS_Cache_Settings"
    Add-Sheet -Data $AllRootHints.ToArray()       -Sheet "DNS_Root_Hints"
    Add-Sheet -Data $AllDNSForwarders.ToArray()   -Sheet "DNS_Forwarders"
    Add-Sheet -Data $AllCondForwarders.ToArray()  -Sheet "DNS_Conditional_Fwds"
    Add-Sheet -Data $AllDNSPartitions.ToArray()   -Sheet "DNS_AD_Partitions"
    Add-Sheet -Data $AllDNSZones.ToArray()        -Sheet "DNS_Zones"
    Add-Sheet -Data $ZoneSummary                  -Sheet "DNS_Zone_Summary"
    Add-Sheet -Data $AllZoneAging.ToArray()       -Sheet "DNS_Zone_Aging"
    Add-Sheet -Data $AllZoneTransfer.ToArray()    -Sheet "DNS_Zone_Transfer"
    Add-Sheet -Data $AllDNSRecords.ToArray()      -Sheet "DNS_Records_All"
    Add-Sheet -Data $RecordTypeSummary            -Sheet "DNS_Record_Type_Summary"
    Write-OK "XLSX saved: $XLSXPath"
}
#endregion

#region --- HTML REPORT ---
Write-Section "Building Enterprise HTML Report"

# Turn off strict mode for HTML generation - null variables in here-string cause terminating errors
Set-StrictMode -Off

# Ensure all table variables have safe defaults if something went wrong above
if (-not $tADSummary)      { $tADSummary      = "<p class='no-data'>No data.</p>" }
if (-not $tDCs)            { $tDCs            = "<p class='no-data'>No data.</p>" }
if (-not $tSrvSettings)    { $tSrvSettings    = "<p class='no-data'>No data.</p>" }
if (-not $tSrvCache)       { $tSrvCache       = "<p class='no-data'>No data.</p>" }
if (-not $tRootHints)      { $tRootHints      = "<p class='no-data'>No data.</p>" }
if (-not $tForwarders)     { $tForwarders     = "<p class='no-data'>No data.</p>" }
if (-not $tCondFwd)        { $tCondFwd        = "<p class='no-data'>No data.</p>" }
if (-not $tPartitions)     { $tPartitions     = "<p class='no-data'>No data.</p>" }
if (-not $tZones)          { $tZones          = "<p class='no-data'>No data.</p>" }
if (-not $tZoneSummary)    { $tZoneSummary    = "<p class='no-data'>No data.</p>" }
if (-not $tZoneAging)      { $tZoneAging      = "<p class='no-data'>No data.</p>" }
if (-not $tZoneXfr)        { $tZoneXfr        = "<p class='no-data'>No data.</p>" }
if (-not $tRecords)        { $tRecords        = "<p class='no-data'>No data.</p>" }
if (-not $tRecTypeSummary) { $tRecTypeSummary = "<p class='no-data'>No data.</p>" }
if (-not $chartLabels)     { $chartLabels     = "" }
if (-not $chartCounts)     { $chartCounts     = "" }

# Derived stats
$totalDCs          = $DomainControllers.Count
$totalZones        = ($AllDNSZones | Select-Object ZoneName -Unique).Count
$primaryZones      = ($AllDNSZones | Where-Object { $_.ZoneType -eq "Primary" } | Select-Object ZoneName -Unique).Count
$secondaryZones    = ($AllDNSZones | Where-Object { $_.ZoneType -eq "Secondary" } | Select-Object ZoneName -Unique).Count
$reverseZones      = ($AllDNSZones | Where-Object { $_.IsReverseLookupZone -eq $true } | Select-Object ZoneName -Unique).Count
$adIntegrated      = ($AllDNSZones | Where-Object { $_.IsDsIntegrated -eq $true } | Select-Object ZoneName -Unique).Count
$totalRecords      = $AllDNSRecords.Count
$totalForwarders   = $AllDNSForwarders.Count
$totalCondFwd      = $AllCondForwarders.Count
$totalRecordTypes  = ($AllDNSRecords | Select-Object RecordType -Unique).Count
$agingEnabled      = ($AllZoneAging | Where-Object { $_.AgingEnabled -eq $true }).Count

# Build HTML tables - each wrapped in try/catch so a single failure can't abort the report
function Safe-HtmlTableBuild {
    param($Data, [string]$TableId, [string]$Caption)
    try {
        return ConvertTo-HtmlTable -Data $Data -TableId $TableId -Caption $Caption
    } catch {
        Write-Warn "Table build failed for $Caption : $_"
        return "<p class='no-data'>Table generation failed for $Caption.</p>"
    }
}

$tADSummary      = Safe-HtmlTableBuild -Data @($ADSummaryData)               -TableId "t_adsum"    -Caption "AD Summary"
$tDCs            = Safe-HtmlTableBuild -Data $DCDetailData                   -TableId "t_dcs"      -Caption "Domain Controllers"
$tSrvSettings    = Safe-HtmlTableBuild -Data $AllDNSServerSettings.ToArray() -TableId "t_srvsett"  -Caption "Server Settings"
$tSrvCache       = Safe-HtmlTableBuild -Data $AllDNSServerCache.ToArray()    -TableId "t_srvcache" -Caption "Cache Settings"
$tRootHints      = Safe-HtmlTableBuild -Data $AllRootHints.ToArray()         -TableId "t_roothint" -Caption "Root Hints"
$tForwarders     = Safe-HtmlTableBuild -Data $AllDNSForwarders.ToArray()     -TableId "t_fwds"     -Caption "Forwarders"
$tCondFwd        = Safe-HtmlTableBuild -Data $AllCondForwarders.ToArray()    -TableId "t_condfwd"  -Caption "Conditional Forwarders"
$tPartitions     = Safe-HtmlTableBuild -Data $AllDNSPartitions.ToArray()     -TableId "t_parts"    -Caption "AD Partitions"
$tZones          = Safe-HtmlTableBuild -Data $AllDNSZones.ToArray()          -TableId "t_zones"    -Caption "DNS Zones"
$tZoneSummary    = Safe-HtmlTableBuild -Data $ZoneSummary                    -TableId "t_zsum"     -Caption "Zone Summary"
$tZoneAging      = Safe-HtmlTableBuild -Data $AllZoneAging.ToArray()         -TableId "t_aging"    -Caption "Zone Aging"
$tZoneXfr        = Safe-HtmlTableBuild -Data $AllZoneTransfer.ToArray()      -TableId "t_xfr"      -Caption "Zone Transfer"
$tRecords        = Safe-HtmlTableBuild -Data $AllDNSRecords.ToArray()        -TableId "t_recs"     -Caption "DNS Records"
$tRecTypeSummary = Safe-HtmlTableBuild -Data $RecordTypeSummary              -TableId "t_recsum"   -Caption "Record Type Summary"

# Record type chart data (PowerShell variables - these must stay unescaped)
$chartLabels = ($RecordTypeSummary | ForEach-Object { "'$($_.RecordType)'" }) -join ","
$chartCounts = ($RecordTypeSummary | ForEach-Object { $_.Count }) -join ","

# Resolve PowerShell variables that go into the HTML before building the here-string
$htmlDomainName   = if ($DomainInfo)  { $DomainInfo.DNSRoot  } else { "N/A" }
$htmlForestName   = if ($ForestInfo)  { $ForestInfo.Name     } else { "N/A" }
$htmlRunByUser    = $env:USERNAME

$HTMLPath = "$OutputFolder\AD_DNS_Report_$DateTimeStamp.html"

# NOTE: All JavaScript variables (e.g. var f, var i, var s) that start with $ are
#       escaped with a backtick (`$) so PowerShell does not try to expand them.
#       Pure PowerShell variables like $totalDCs, $tDCs, $DateDisplay etc. are left
#       unescaped so they ARE expanded as intended.

$HTML = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>AD DNS Report - Stephen McKee - IGT Everi</title>
<style>
:root {
  --bg:#0d1117; --surface:#161b22; --surface2:#1c2128; --border:#30363d;
  --accent:#2563eb; --text:#e6edf3; --muted:#7d8590;
  --green:#3fb950; --yellow:#d29922; --red:#f85149;
  --purple:#a371f7; --orange:#db6d28; --cyan:#39d3f2; --teal:#56d364;
}
*{box-sizing:border-box;margin:0;padding:0;}
html{scroll-behavior:smooth;}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;}

/* HEADER */
.top-header{
  background:linear-gradient(135deg,#0a0e14 0%,#161b22 60%,#0a0e14 100%);
  border-bottom:2px solid var(--accent);padding:22px 32px 18px;
  position:sticky;top:0;z-index:200;
  display:flex;justify-content:space-between;align-items:flex-end;
}
.title-block h1{font-size:21px;font-weight:700;color:#fff;letter-spacing:.5px;}
.title-block h1 em{color:var(--cyan);font-style:normal;}
.title-block .sub{font-size:11px;color:var(--muted);margin-top:3px;}
.meta-block{text-align:right;font-size:11px;color:var(--muted);line-height:1.8;}
.meta-block strong{color:var(--cyan);}

/* SIDENAV */
.sidenav{
  position:fixed;top:0;left:0;width:215px;height:100vh;
  background:var(--surface);border-right:1px solid var(--border);
  overflow-y:auto;padding-top:85px;z-index:100;
}
.nav-group{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:1.2px;padding:14px 16px 5px;}
.sidenav a{
  display:block;padding:6px 16px;color:var(--muted);text-decoration:none;
  font-size:12px;border-left:3px solid transparent;transition:all .15s;
}
.sidenav a:hover,.sidenav a.active{color:var(--text);background:var(--surface2);border-left-color:var(--accent);}

/* MAIN */
.main{margin-left:215px;padding:22px 28px 40px;}

/* GLOBAL SEARCH */
.global-bar{
  background:var(--surface);border:1px solid var(--border);border-radius:8px;
  padding:12px 18px;margin-bottom:20px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;
}
.global-bar input{
  background:var(--surface2);border:1px solid var(--border);color:var(--text);
  padding:7px 14px;border-radius:6px;font-size:13px;width:340px;outline:none;
  transition:border-color .2s;
}
.global-bar input:focus{border-color:var(--accent);}
.global-bar label{color:var(--muted);font-size:12px;}
.match-count{font-size:12px;color:var(--cyan);}

/* CARDS */
.cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(148px,1fr));gap:12px;margin-bottom:24px;}
.card{
  background:var(--surface);border:1px solid var(--border);border-radius:10px;
  padding:15px;position:relative;overflow:hidden;
}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.card.c-blue::before{background:var(--accent);}
.card.c-green::before{background:var(--green);}
.card.c-yellow::before{background:var(--yellow);}
.card.c-red::before{background:var(--red);}
.card.c-purple::before{background:var(--purple);}
.card.c-cyan::before{background:var(--cyan);}
.card.c-orange::before{background:var(--orange);}
.card.c-teal::before{background:var(--teal);}
.card-val{font-size:30px;font-weight:700;color:#fff;line-height:1.1;}
.card-lbl{font-size:11px;color:var(--muted);margin-top:4px;}

/* CHART CONTAINER */
.chart-container{
  background:var(--surface);border:1px solid var(--border);border-radius:10px;
  padding:18px;margin-bottom:22px;
}
.chart-container h3{font-size:13px;color:var(--muted);margin-bottom:12px;font-weight:500;}
.chart-inner{display:flex;align-items:flex-end;gap:6px;height:140px;padding-bottom:20px;position:relative;}
.bar-wrap{display:flex;flex-direction:column;align-items:center;flex:1;min-width:36px;height:100%;justify-content:flex-end;}
.bar{width:100%;border-radius:3px 3px 0 0;min-height:4px;transition:height .3s;}
.bar-lbl{font-size:9px;color:var(--muted);margin-top:4px;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;width:100%;}
.bar-count{font-size:10px;color:var(--text);margin-bottom:2px;font-weight:600;}

/* SECTION */
.section{margin-bottom:18px;border:1px solid var(--border);border-radius:10px;overflow:hidden;}
.section-header{
  display:flex;align-items:center;gap:10px;
  padding:12px 16px;
  background:var(--surface2);
  cursor:pointer;user-select:none;
  transition:background .15s;
  list-style:none;
}
.section-header:hover{background:#1e242c;}
.section-header h2{font-size:15px;font-weight:600;color:#fff;flex:1;}
.badge{background:var(--accent);color:#fff;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;}
.badge.red{background:var(--red);}
.badge.green{background:var(--green);}
.badge.yellow{background:var(--yellow);}
.chevron{width:18px;height:18px;flex-shrink:0;color:var(--muted);transition:transform .25s ease;}
.section.collapsed .chevron{transform:rotate(-90deg);}
.section-body{
  padding:16px;
  overflow:hidden;
  max-height:20000px;
  transition:max-height .35s ease,padding .3s ease,opacity .25s ease;
  opacity:1;
}
.section.collapsed .section-body{max-height:0;padding-top:0;padding-bottom:0;opacity:0;pointer-events:none;}
.collapse-controls{display:flex;gap:8px;align-items:center;}
.ctrl-btn{
  background:var(--surface2);border:1px solid var(--border);color:var(--muted);
  padding:5px 12px;border-radius:5px;font-size:11px;cursor:pointer;transition:all .15s;white-space:nowrap;
}
.ctrl-btn:hover{background:var(--surface);color:var(--text);border-color:var(--accent);}

/* TABLE */
.table-wrapper{position:relative;}
.search-box{
  background:var(--surface2);border:1px solid var(--border);color:var(--text);
  padding:6px 12px;border-radius:5px;font-size:12px;margin-bottom:8px;min-width:260px;outline:none;
}
.search-box:focus{border-color:var(--accent);}
.table-scroll{overflow-x:auto;border-radius:8px;border:1px solid var(--border);}
.data-table{width:100%;border-collapse:collapse;white-space:nowrap;}
.data-table thead{background:var(--surface2);position:sticky;top:0;z-index:10;}
.data-table th{
  padding:8px 12px;text-align:left;font-size:11px;font-weight:600;
  color:var(--muted);text-transform:uppercase;letter-spacing:.5px;
  border-bottom:1px solid var(--border);cursor:pointer;user-select:none;
}
.data-table th:hover{color:var(--text);}
.sort-icon{opacity:.35;font-size:10px;}
.data-table td{
  padding:7px 12px;border-bottom:1px solid #21262d;font-size:12px;
  max-width:300px;overflow:hidden;text-overflow:ellipsis;
}
.data-table tbody tr:hover{background:var(--surface2);}
.data-table tbody tr:last-child td{border-bottom:none;}
.data-table tr.hidden{display:none;}

td.z-primary  {color:var(--cyan);font-weight:600;}
td.z-secondary{color:var(--yellow);}
td.z-stub     {color:var(--orange);}
td.z-forwarder{color:var(--purple);}
td.bool-yes   {color:var(--green);font-weight:600;}
td.bool-no    {color:var(--muted);}
td.rt-a       {color:var(--cyan);}
td.rt-aaaa    {color:var(--purple);}
td.rt-cname   {color:var(--yellow);}
td.rt-mx      {color:var(--orange);}
td.rt-ptr     {color:var(--teal);}
td.rt-soa     {color:var(--red);}
td.rt-srv     {color:var(--accent);}
td.rt-txt     {color:var(--muted);}
td.rt-ns      {color:var(--green);}

.no-data{color:var(--muted);font-style:italic;padding:12px 0;}

/* Footer */
.footer{
  margin-left:215px;padding:14px 28px;border-top:1px solid var(--border);
  color:var(--muted);font-size:11px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;
}

::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:var(--bg);}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}

@media print{
  .sidenav,.global-bar,.search-box,.top-header{display:none!important;}
  .main,.footer{margin-left:0!important;}
}
</style>
</head>
<body>

<!-- HEADER -->
<header class="top-header">
  <div class="title-block">
    <h1><em>AD DNS</em> Report</h1>
    <div class="sub">Created by $ScriptAuthor</div>
  </div>
  <div class="meta-block">
    <div>Generated: <strong>$DateDisplay</strong></div>
    <div>Run By: <strong>$htmlRunByUser</strong></div>
    <div>Domain: <strong>$htmlDomainName</strong></div>
    <div>Forest: <strong>$htmlForestName</strong></div>
  </div>
</header>

<!-- SIDENAV -->
<nav class="sidenav">
  <div class="nav-group">Overview</div>
  <a href="#sec-summary">Summary &amp; Stats</a>
  <a href="#sec-ad">AD Domain Info</a>
  <a href="#sec-dcs">Domain Controllers</a>
  <div class="nav-group">DNS Servers</div>
  <a href="#sec-srv-settings">Server Settings</a>
  <a href="#sec-cache">Cache Settings</a>
  <a href="#sec-roothints">Root Hints</a>
  <div class="nav-group">Forwarding</div>
  <a href="#sec-forwarders">Forwarders</a>
  <a href="#sec-condfwd">Conditional Forwarders</a>
  <div class="nav-group">AD Integration</div>
  <a href="#sec-partitions">AD DNS Partitions</a>
  <div class="nav-group">Zones</div>
  <a href="#sec-zone-summary">Zone Summary</a>
  <a href="#sec-zones">All Zones Detail</a>
  <a href="#sec-aging">Aging &amp; Scavenging</a>
  <a href="#sec-xfr">Zone Transfers</a>
  <div class="nav-group">Records</div>
  <a href="#sec-rec-type">Record Type Summary</a>
  <a href="#sec-records">All DNS Records</a>
</nav>

<main class="main">

  <!-- GLOBAL SEARCH -->
  <div class="global-bar">
    <label>&#128269; Global Search:</label>
    <input type="text" id="globalSearch" placeholder="Search all tables..." oninput="globalFilter()"/>
    <span class="match-count" id="matchCount"></span>
    <div class="collapse-controls" style="margin-left:auto;">
      <button class="ctrl-btn" onclick="expandAll()">&#9660; Expand All</button>
      <button class="ctrl-btn" onclick="collapseAll()">&#9658; Collapse All</button>
    </div>
  </div>

  <!-- SUMMARY SECTION -->
  <div class="section" id="sec-summary">
    <div class="section-header" onclick="toggleSection('sec-summary')">
      <h2>&#128202; DNS Environment Summary</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">
    <div class="cards-grid">
      <div class="card c-blue"><div class="card-val">$totalDCs</div><div class="card-lbl">DNS Servers (DCs)</div></div>
      <div class="card c-cyan"><div class="card-val">$totalZones</div><div class="card-lbl">Total DNS Zones</div></div>
      <div class="card c-green"><div class="card-val">$primaryZones</div><div class="card-lbl">Primary Zones</div></div>
      <div class="card c-yellow"><div class="card-val">$secondaryZones</div><div class="card-lbl">Secondary Zones</div></div>
      <div class="card c-teal"><div class="card-val">$reverseZones</div><div class="card-lbl">Reverse Lookup Zones</div></div>
      <div class="card c-purple"><div class="card-val">$adIntegrated</div><div class="card-lbl">AD-Integrated Zones</div></div>
      <div class="card c-orange"><div class="card-val">$totalForwarders</div><div class="card-lbl">DNS Forwarders</div></div>
      <div class="card c-blue"><div class="card-val">$totalCondFwd</div><div class="card-lbl">Conditional Forwarders</div></div>
      <div class="card c-green"><div class="card-val">$totalRecords</div><div class="card-lbl">Total DNS Records</div></div>
      <div class="card c-cyan"><div class="card-val">$totalRecordTypes</div><div class="card-lbl">Record Types Present</div></div>
      <div class="card c-teal"><div class="card-val">$agingEnabled</div><div class="card-lbl">Zones w/ Aging ON</div></div>
    </div>
    <div class="chart-container">
      <h3>DNS Record Distribution by Type</h3>
      <div class="chart-inner" id="barChart"></div>
    </div>
    </div>
  </div>

  <!-- AD SUMMARY -->
  <div class="section" id="sec-ad">
    <div class="section-header" onclick="toggleSection('sec-ad')">
      <h2>Active Directory Domain Summary</h2><span class="badge">AD</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tADSummary</div>
  </div>

  <!-- DOMAIN CONTROLLERS -->
  <div class="section" id="sec-dcs">
    <div class="section-header" onclick="toggleSection('sec-dcs')">
      <h2>Domain Controllers / DNS Servers</h2><span class="badge">$totalDCs</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tDCs</div>
  </div>

  <!-- DNS SERVER SETTINGS -->
  <div class="section" id="sec-srv-settings">
    <div class="section-header" onclick="toggleSection('sec-srv-settings')">
      <h2>DNS Server Settings &amp; Options</h2><span class="badge">Per Server</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tSrvSettings</div>
  </div>

  <!-- CACHE SETTINGS -->
  <div class="section" id="sec-cache">
    <div class="section-header" onclick="toggleSection('sec-cache')">
      <h2>DNS Server Cache Settings</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tSrvCache</div>
  </div>

  <!-- ROOT HINTS -->
  <div class="section" id="sec-roothints">
    <div class="section-header" onclick="toggleSection('sec-roothints')">
      <h2>Root Hints</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tRootHints</div>
  </div>

  <!-- FORWARDERS -->
  <div class="section" id="sec-forwarders">
    <div class="section-header" onclick="toggleSection('sec-forwarders')">
      <h2>DNS Forwarders</h2><span class="badge">$totalForwarders</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tForwarders</div>
  </div>

  <!-- CONDITIONAL FORWARDERS -->
  <div class="section" id="sec-condfwd">
    <div class="section-header" onclick="toggleSection('sec-condfwd')">
      <h2>Conditional Forwarders</h2><span class="badge">$totalCondFwd</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tCondFwd</div>
  </div>

  <!-- AD DNS PARTITIONS -->
  <div class="section" id="sec-partitions">
    <div class="section-header" onclick="toggleSection('sec-partitions')">
      <h2>Active Directory DNS Partitions</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tPartitions</div>
  </div>

  <!-- ZONE SUMMARY -->
  <div class="section" id="sec-zone-summary">
    <div class="section-header" onclick="toggleSection('sec-zone-summary')">
      <h2>DNS Zone Summary</h2><span class="badge">$totalZones Zones</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tZoneSummary</div>
  </div>

  <!-- ALL ZONES DETAIL -->
  <div class="section" id="sec-zones">
    <div class="section-header" onclick="toggleSection('sec-zones')">
      <h2>All DNS Zones (Full Detail)</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tZones</div>
  </div>

  <!-- ZONE AGING -->
  <div class="section" id="sec-aging">
    <div class="section-header" onclick="toggleSection('sec-aging')">
      <h2>Zone Aging &amp; Scavenging</h2><span class="badge green">$agingEnabled Enabled</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tZoneAging</div>
  </div>

  <!-- ZONE TRANSFER -->
  <div class="section" id="sec-xfr">
    <div class="section-header" onclick="toggleSection('sec-xfr')">
      <h2>Zone Transfer Settings</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tZoneXfr</div>
  </div>

  <!-- RECORD TYPE SUMMARY -->
  <div class="section" id="sec-rec-type">
    <div class="section-header" onclick="toggleSection('sec-rec-type')">
      <h2>DNS Record Type Summary</h2><span class="badge">$totalRecordTypes Types</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tRecTypeSummary</div>
  </div>

  <!-- ALL DNS RECORDS -->
  <div class="section" id="sec-records">
    <div class="section-header" onclick="toggleSection('sec-records')">
      <h2>All DNS Resource Records</h2><span class="badge">$totalRecords Records</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tRecords</div>
  </div>

</main>

<footer class="footer">
  <span>$ScriptTitle &mdash; $ScriptAuthor</span>
  <span>Generated: $DateDisplay &nbsp;|&nbsp; By: $htmlRunByUser</span>
</footer>

<script>
// ---- Collapsible Sections ----
function toggleSection(id) {
  var sec = document.getElementById(id);
  var wasCollapsed = sec.classList.contains('collapsed');
  sec.classList.toggle('collapsed');
  try {
    var states = JSON.parse(sessionStorage.getItem('secStates') || '{}');
    states[id] = !wasCollapsed;
    sessionStorage.setItem('secStates', JSON.stringify(states));
  } catch(e) {}
}

function collapseAll() {
  document.querySelectorAll('.section').forEach(function(s) { s.classList.add('collapsed'); });
  try { sessionStorage.removeItem('secStates'); } catch(e) {}
}

function expandAll() {
  document.querySelectorAll('.section').forEach(function(s) { s.classList.remove('collapsed'); });
  try { sessionStorage.removeItem('secStates'); } catch(e) {}
}

(function restoreStates() {
  try {
    var states = JSON.parse(sessionStorage.getItem('secStates') || '{}');
    Object.keys(states).forEach(function(id) {
      if (states[id]) {
        var el = document.getElementById(id);
        if (el) el.classList.add('collapsed');
      }
    });
  } catch(e) {}
})();

document.querySelectorAll('.sidenav a').forEach(function(a) {
  a.addEventListener('click', function(e) {
    var href = a.getAttribute('href');
    if (href && href.startsWith('#')) {
      var target = document.getElementById(href.slice(1));
      if (target) target.classList.remove('collapsed');
    }
  });
});

// ---- Per-table search ----
function filterTable(input, tid) {
  var f = input.value.toUpperCase();
  var rows = document.getElementById(tid).getElementsByTagName('tr');
  for (var i = 1; i < rows.length; i++) {
    var found = false;
    var cells = rows[i].getElementsByTagName('td');
    for (var j = 0; j < cells.length; j++) {
      if (cells[j].textContent.toUpperCase().indexOf(f) > -1) { found = true; break; }
    }
    rows[i].classList.toggle('hidden', !found);
  }
}

// ---- Global search ----
function globalFilter() {
  var f = document.getElementById('globalSearch').value.toUpperCase();
  var tables = document.querySelectorAll('.data-table');
  var hits = 0;
  tables.forEach(function(t) {
    var rows = t.getElementsByTagName('tr');
    var tableHits = 0;
    for (var i = 1; i < rows.length; i++) {
      var show = !f || rows[i].textContent.toUpperCase().indexOf(f) > -1;
      rows[i].classList.toggle('hidden', !show);
      if (show && f) { hits++; tableHits++; }
    }
    if (f && tableHits > 0) {
      var sec = t.closest('.section');
      if (sec) sec.classList.remove('collapsed');
    }
    var wrap = t.closest('.table-wrapper');
    if (wrap && f) { var sb = wrap.querySelector('.search-box'); if (sb) sb.value = ''; }
  });
  var el = document.getElementById('matchCount');
  el.textContent = f ? (hits + ' match' + (hits !== 1 ? 'es' : '')) : '';
}

// ---- Sort ----
function sortTable(tid, col) {
  var t = document.getElementById(tid);
  var rows = Array.from(t.tBodies[0].rows);
  var asc = t.dataset.sc == col && t.dataset.sd == 'asc' ? false : true;
  t.dataset.sc = col; t.dataset.sd = asc ? 'asc' : 'desc';
  rows.sort(function(a, b) {
    var av = a.cells[col] ? a.cells[col].textContent.trim() : '';
    var bv = b.cells[col] ? b.cells[col].textContent.trim() : '';
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
  });
  rows.forEach(function(r) { t.tBodies[0].appendChild(r); });
}

// ---- Nav scroll highlight ----
window.addEventListener('scroll', function() {
  var secs = document.querySelectorAll('.section');
  var links = document.querySelectorAll('.sidenav a');
  var cur = '';
  secs.forEach(function(s) { if (s.getBoundingClientRect().top <= 130) cur = s.id; });
  links.forEach(function(a) {
    a.classList.toggle('active', a.getAttribute('href') === '#' + cur);
  });
});

// ---- Bar chart ----
(function() {
  var labels = [$chartLabels];
  var counts = [$chartCounts];
  if (!labels.length) return;
  var max = Math.max.apply(null, counts);
  var colors = ['#39d3f2','#3fb950','#d29922','#f85149','#a371f7','#db6d28','#56d364','#2563eb','#ff7b72','#79c0ff','#ffa657','#d2a8ff'];
  var container = document.getElementById('barChart');
  labels.forEach(function(lbl, i) {
    var pct = max > 0 ? (counts[i] / max * 100) : 0;
    var wrap = document.createElement('div'); wrap.className = 'bar-wrap';
    var cntEl = document.createElement('div'); cntEl.className = 'bar-count'; cntEl.textContent = counts[i];
    var bar = document.createElement('div'); bar.className = 'bar';
    bar.style.height = pct + '%';
    bar.style.background = colors[i % colors.length];
    bar.title = lbl + ': ' + counts[i];
    var lblEl = document.createElement('div'); lblEl.className = 'bar-lbl'; lblEl.textContent = lbl;
    wrap.appendChild(cntEl); wrap.appendChild(bar); wrap.appendChild(lblEl);
    container.appendChild(wrap);
  });
})();
</script>
</body>
</html>
"@

try {
    [System.IO.File]::WriteAllText($HTMLPath, $HTML, [System.Text.Encoding]::UTF8)
    if (Test-Path $HTMLPath) {
        $htmlSize = (Get-Item $HTMLPath).Length
        Write-OK "HTML report saved: $HTMLPath ($htmlSize bytes)"
        Write-Info "Opening HTML report in default browser..."
        Start-Process $HTMLPath
    } else {
        Write-Fail "HTML file not found after write attempt: $HTMLPath"
    }
} catch {
    Write-Fail "HTML WriteAllText FAILED: $_"
    Write-Fail "Attempting fallback write via Set-Content..."
    try {
        $HTML | Set-Content -Path $HTMLPath -Encoding UTF8 -Force
        Write-OK "HTML fallback write succeeded: $HTMLPath"
        Write-Info "Opening HTML report in default browser..."
        Start-Process $HTMLPath
    } catch {
        Write-Fail "HTML fallback also failed: $_"
    }
}
#endregion

#region --- FINAL SUMMARY ---
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  REPORT COMPLETE" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Output Folder : $OutputFolder" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Files Generated:" -ForegroundColor White
Get-ChildItem -Path $OutputFolder | ForEach-Object {
    Write-Host ("    {0,-60} {1,8}" -f $_.Name, ("{0:N0} KB" -f ($_.Length/1KB))) -ForegroundColor Gray
}
Write-Host ""
Write-Host "  DNS Summary:" -ForegroundColor White
Write-Host "    DNS Servers (DCs)      : $totalDCs"
Write-Host "    Total DNS Zones        : $totalZones ($primaryZones Primary / $secondaryZones Secondary)"
Write-Host "    Reverse Lookup Zones   : $reverseZones"
Write-Host "    AD-Integrated Zones    : $adIntegrated"
Write-Host "    Zones w/ Aging ON      : $agingEnabled"
Write-Host "    Conditional Forwarders : $totalCondFwd"
Write-Host "    DNS Forwarders         : $totalForwarders"
Write-Host "    Total DNS Records      : $totalRecords"
Write-Host "    Record Types Present   : $totalRecordTypes"
Write-Host ""
Write-Host "  Opening output folder..." -ForegroundColor Gray
Start-Process explorer.exe $OutputFolder
Write-Host "================================================================" -ForegroundColor Cyan
#endregion
