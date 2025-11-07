<#
.SYNOPSIS
  Build AD DNS replication map (CSV + Visio). One Visio page per AD Site + Overview page.
  Shows AD-integrated zones and draws replication links labeled with zone, scope and site names.
.REQUIREMENTS
  - Run elevated on admin workstation
  - Modules: ActiveDirectory, DnsServer (RSAT)
  - Microsoft Visio (desktop) installed
  - Account: DNS Admin or Domain Admin
.NOTES
  - Script queries current forest. For multiple forests, run in each forest or extend to remote credentials.
#>

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module DnsServer -ErrorAction Stop

# Output files (adjust paths as needed)
$timestamp = (Get-Date -Format "yyyyMMdd_HHmmss")
$csvPath   = "C:\AD_DNS_ReplicationData_$timestamp.csv"
$vsdxPath  = "C:\AD_DNS_ReplicationMap_$timestamp.vsdx"

Write-Host "Collecting domain controllers and DNS zone info..." -ForegroundColor Cyan

# Get DCs (all sites)
$dcList = Get-ADDomainController -Filter * | Sort-Object Site

$allData = @()

foreach ($dc in $dcList) {
    $server = $dc.HostName
    Write-Host "Querying DNS on $server ..." -ForegroundColor Yellow
    try {
        $siteName = (Get-ADReplicationSite -Server $server -ErrorAction SilentlyContinue).Name
    } catch {
        $siteName = "Unknown"
    }

    # Attempt to get AD-integrated zones on this server
    try {
        $zones = Get-DnsServerZone -ComputerName $server -ErrorAction Stop |
                 Where-Object { $_.IsDsIntegrated -eq $true }
    } catch {
        Write-Warning "Get-DnsServerZone failed on {$server}: $_"
        $zones = @()
    }

    foreach ($z in $zones) {
        $allData += [PSCustomObject]@{
            ServerName      = $server
            IPv4            = $dc.IPv4Address
            SiteName        = $siteName
            ZoneName        = $z.ZoneName
            ReplicationScope= $z.ReplicationScope
            ZoneType        = if ($z.ZoneType) { $z.ZoneType } else { "Unknown" }
        }
    }
}

# Export CSV (raw data)
$allData | Sort-Object SiteName, ServerName, ZoneName | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Raw CSV exported to $csvPath" -ForegroundColor Green

# ----------------------
# Build Visio document
# ----------------------
Write-Host "Launching Visio and creating pages..." -ForegroundColor Cyan
$visio = New-Object -ComObject Visio.Application
$visio.Visible = $true
$doc = $visio.Documents.Add("")   # new document

# Prepare pages: one per site + Overview page
$sites = $allData.SiteName | Sort-Object -Unique
if (-not $sites) {
    Write-Warning "No AD-integrated zones found or no reachable DNS servers. Exiting."
    exit 1
}

# Create Overview page first
$overviewPage = $doc.Pages.Add()
$overviewPage.Name = "Overview - Sites & Replication"

# Create one page per site
$sitePages = @{}
foreach ($s in $sites) {
    $p = $doc.Pages.Add()
    $p.Name = "Site - $s"
    $sitePages[$s] = $p
}

# Helper: draw servers on a page and keep shape refs
$serverShapes = @{}   # indexed by ServerName
foreach ($s in $sites) {
    $p = $sitePages[$s]
    $pageCenterX = 4
    $pageTopY = 10
    # get servers in this site
    $servers = ($allData | Where-Object { $_.SiteName -eq $s } | Select-Object -ExpandProperty ServerName -Unique)
    $row = 0
    foreach ($srv in $servers) {
        # create oval shape
        $ox = $pageCenterX + (($row % 3) * 2.2) - 2
        $oy = $pageTopY - ( [math]::Floor($row / 3) * 1.8 )
        $shape = $p.DrawOval($ox, $oy, $ox + 1.8, $oy - 0.8)
        $shape.Text = $srv
        $shape.CellsU("FillForegnd").FormulaU = "RGB(200,200,255)"
        $serverShapes[$srv] = $shape
        $row++
    }
    # small title
    $title = $p.DrawRectangle(0.5, 10.5, 6, 9.6)
    $title.Text = "Site: $s"
    $title.CellsU("LineColor").FormulaU = "RGB(0,0,255)"
}

# Place summary server shapes on Overview page too (compact grid)
$overviewP = $overviewPage
$ovX = 1; $ovY = 10; $col = 0
$overviewServerShapes = @{}
$overviewServers = ($allData.ServerName | Sort-Object -Unique)
foreach ($srv in $overviewServers) {
    $sx = $ovX + ($col * 2.0)
    $sy = $ovY
    $sshape = $overviewP.DrawOval($sx, $sy, $sx + 1.6, $sy - 0.7)
    $sshape.Text = $srv
    $sshape.CellsU("FillForegnd").FormulaU = "RGB(200,200,255)"
    $overviewServerShapes[$srv] = $sshape
    $col++
    if ($col -gt 6) { $col = 0; $ovY -= 1.2 }
}

# Draw per-site items (zones) on each site page near its servers
foreach ($s in $sites) {
    $p = $sitePages[$s]
    $servers = ($allData | Where-Object { $_.SiteName -eq $s } | Select-Object -ExpandProperty ServerName -Unique)
    foreach ($srv in $servers) {
        $srvZones = $allData | Where-Object { $_.ServerName -eq $srv } | Sort-Object ZoneName -Unique
        # find shape on that page
        $sShape = $serverShapes[$srv]
        if (-not $sShape) { continue }
        $offsetX = 1.6; $offsetY = 0
        foreach ($z in $srvZones) {
            $zx = $sShape.CellsU("PinX").ResultIU + $offsetX
            $zy = $sShape.CellsU("PinY").ResultIU + $offsetY
            $zshape = $p.DrawRectangle($zx, $zy, $zx + 2.2, $zy - 0.6)
            $zshape.Text = ($z.ZoneName -replace "`n"," ") + "`n[" + $z.ReplicationScope + "]"
            # color by replication scope
            switch ($z.ReplicationScope) {
                "ForestDnsZones" { $zshape.CellsU("FillForegnd").FormulaU = "RGB(180,220,255)" }  # light blue
                "DomainDnsZones" { $zshape.CellsU("FillForegnd").FormulaU = "RGB(200,255,200)" }  # light green
                default { $zshape.CellsU("FillForegnd").FormulaU = "RGB(255,230,180)" }           # orange-ish
            }
            # connector line from server to zone item
            $p.DrawLine(
                $sShape.CellsU("PinX").ResultIU, $sShape.CellsU("PinY").ResultIU,
                $zshape.CellsU("PinX").ResultIU, $zshape.CellsU("PinY").ResultIU
            ) | Out-Null

            $offsetY -= 0.9
        }
    }
}

# ----------------------------
# Draw replication links on Overview page
# ----------------------------
$overview = $overviewP
Write-Host "Drawing replication links on Overview..." -ForegroundColor Cyan

# For each zone, determine all servers that host it (across sites)
$zones = $allData.ZoneName | Sort-Object -Unique
foreach ($zone in $zones) {
    $members = $allData | Where-Object { $_.ZoneName -eq $zone } | Sort-Object SiteName, ServerName
    if ($members.Count -le 1) { continue }

    # Choose color based on replication scope (use first member's scope)
    $scope = ($members | Select-Object -First 1).ReplicationScope
    switch ($scope) {
        "ForestDnsZones" { $lineColor = "RGB(100,150,255)"; $scopeLabel="Forest" }
        "DomainDnsZones" { $lineColor = "RGB(150,255,150)"; $scopeLabel="Domain" }
        default { $lineColor = "RGB(255,200,100)"; $scopeLabel="Custom" }
    }

    # connect each pair and label with zone + scope + site names
    for ($i=0; $i -lt $members.Count; $i++) {
        for ($j = $i+1; $j -lt $members.Count; $j++) {
            $m1 = $members[$i]
            $m2 = $members[$j]
            $s1 = $m1.ServerName; $s2 = $m2.ServerName
            if ($overviewServerShapes[$s1] -and $overviewServerShapes[$s2]) {
                $x1 = $overviewServerShapes[$s1].CellsU("PinX").ResultIU
                $y1 = $overviewServerShapes[$s1].CellsU("PinY").ResultIU
                $x2 = $overviewServerShapes[$s2].CellsU("PinX").ResultIU
                $y2 = $overviewServerShapes[$s2].CellsU("PinY").ResultIU
                $ln = $overview.DrawLine($x1, $y1, $x2, $y2)
                # set color
                $ln.CellsU("LineColor").FormulaU = $lineColor
                # label: zone (scope) - SiteA <> SiteB
                $label = "$($zone) ($scopeLabel) - $($m1.SiteName) <> $($m2.SiteName)"
                $ln.Text = $label
            }
        }
    }
}

# Save the document
try {
    $doc.SaveAs($vsdxPath)
    Write-Host "Visio file saved to $vsdxPath" -ForegroundColor Green
} catch {
    Write-Warning "Failed to save Visio file: $_"
}

# Cleanup COM objects (keep Visio open if you want to tweak)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($overviewP) | Out-Null
foreach ($p in $sitePages.Values) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($p) | Out-Null }
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($visio) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host "Done. CSV: $csvPath  |  Visio: $vsdxPath" -ForegroundColor Cyan
