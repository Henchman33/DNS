$ZoneName = 'someinternalzone.com'
$DNSRecords = Get-DnsServerResourceRecord -ComputerName (Get-ADRootDSE).dnshostname -ZoneName $ZoneName -RRType A
$UniqueDNSRecords = $DNSRecords.hostname | select -Unique
diff $DNSRecords.hostname $UniqueDNSRecords



$DNSRecords = Get-DnsServerResourceRecord -ComputerName (Get-ADRootDSE).dnshostname -ZoneName $ZoneName -RRType A
$UniqueDNSRecords = $DNSRecords.RecordData.IPv4Address.IPAddressToString | select -Unique
diff $DNSRecords.RecordData.IPv4Address.IPAddressToString $UniqueDNSRecords
