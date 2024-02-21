# Automatic services
$auto = "MSExchangeADTopology",
"MSExchangeAntispamUpdate",
"MSExchangeDagMgmt",
"MSExchangeDiagnostics",
"MSExchangeEdgeSync",
"MSExchangeFrontEndTransport",
"MSExchangeHM",
"MSExchangeImap4",
"MSExchangeIMAP4BE",
"MSExchangeIS",
"MSExchangeMailboxAssistants",
"MSExchangeMailboxReplication",
"MSExchangeDelivery",
"MSExchangeSubmission",
"MSExchangeRepl",
"MSExchangeRPC",
"MSExchangeFastSearch",
"HostControllerService",
"MSExchangeServiceHost",
"MSExchangeThrottling",
"MSExchangeTransport",
"MSExchangeTransportLogSearch",
"MSExchangeUM",
"MSExchangeUMCR",
"FMS",
"IISADMIN",
"RemoteRegistry",
"SearchExchangeTracing",
"Winmgmt",
"W3SVC"

# Manual services
$man = "MSExchangePop3",
"MSExchangePOP3BE",
"wsbexchange",
"AppIDSvc",
"pla"

# Enable Services
foreach ($service in $auto) {
   Set-Service -Name $service -StartupType Automatic
   Write-Host "Enabling "$service
}
foreach ($service2 in $man) {
   Set-Service -Name $service2 -StartupType Manual
   Write-Host "Enabling "$service2
}

# Start Services
foreach ($service in $auto) {
   Start-Service -Name $service
   Write-Host "Starting "$service 
}