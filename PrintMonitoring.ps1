# Custom Print Monitor

# Global Variables
$htmlFile = "C:\inetpub\wwwroot\index.html"
$serverName = "servername"
$checkFileFolder = "c:\inetpub\wwwroot\PrintMonitoring\"

$uptime = New-TimeSpan -Start ([Management.ManagementDateTimeConverter]::ToDateTime((gwmi Win32_OperatingSytem).LastBootupTime)) -End (Get-Date)
$uptimeString = "$($uptime.Days) Days, $($uptime.Hours) Hours, $($uptime.Minutes) Mins"

function customStatus ($thisStatus) {
    switch ($thisStatus) {
        "0"   { $customText = "Ready" }
        "1"   { $customText = "Other Error" }
        "2"   { $customText = "Paused" }
        "3"   { $customText = "Low Paper" }
        "4"   { $customText = "Ready" }
        "5"   { $customText = "Ready" }
        "6"   { $customText = "Ready" }
        "7"   { $customText = "Ready" }
        "8"   { $customText = "Ready" }
        "9"   { $customText = "Ready" }
        "10"  { $customText = "Ready" }
        "11"  { $customText = "Ready" }
        "12"  { $customText = "Ready" }
        "13"  { $customText = "Ready" }
        "14"  { $customText = "Ready" }
        "15"  { $customText = "Ready" }
        default {$customText = "Error (D)"}
    }
    return $customText
}

#region processing

# main array
$prnStatus = gwmi Win32_Printer -ComputerName $serverName | % { $prName = $_.Name; $prtName = $_.PortName; $srv = $_.SystemName; $comment = $_.Comment; $status = $_.DetectedErrorState; $jobCount = $_.JobCountSinceLastReset; $location = $_.Location; gwmi Win32_TcpIpPrinterPort -ComputerName $serverName |
   where { $_.Name -eq $prtName } |
   select @{name="Name";expression={$prName}}, 
        @{name="Server";expression={$srv}},
        @{name="Status";expression={$status}},
        @{name="Jobs";expression={$jobCount}}, 
        @{name="Location";expression={$location}}, 
        hostaddress,
        @{name="Comment";expression={$comment}},
        @{name="LastErrorTime";expression={$null}},
        @{name="Uptime";expression={$uptimeString}}
    }

# filter specific status results into arrays
$onlinePrinters = $prnStatus | where {$_.Status -eq "0"}
$issuePrinters = $prnStatus | where {$_.Status -ne "0"}
$errorStatusPrinters = $prnStatus | where { ($_.Status -ne "0") -and ($_.Status -ne "3") -and ($_.Status -ne "5") }
$warningStatusPrinters = $prnStatus | where { ($_.Status -ne "0") -and ($_.Status -ne "3") -or ($_.Status -ne "5") }
# modify status table with last error time

$errorStatusPrinters | foreach {
    $latestError = Get-PrintJob -ComputerName $serverName -PrinterName $_.Name | select Id, ComputerName, PrinterName, UserName. DocumentName, SubmittedTime, JobStatus -First 1
    $status = customStatus $_.Status
    $firstErrorTime = (Get-Date $latestError.SubmittedTime)
    $threshold = (Get-Date).AddMinutes(-4)
    # See if error is past given threshold
    if ($threshold -gt $firstErrorTime) {
        $checkFile = $checkFileFolder + $($latestError | select -ExpandProperty PrinterName) + ".txt"
        # check to see if checkfile exists, if not create it
        if (!(Test-Path -Path $checkFile)) {
            $latestError.SubmittedTime.ToString() | Out-File $checkFile -Force
        }
        Get-Content $checkFile | Write-Warning
    }
    # Modify original table to include updated LastErrorTime data
    $prnStatus | where Name -EQ $_.Name | foreach {$_.LastErrorTime = $firstErrorTime}
}

# determine the last error for all printers
$currentErrorFiles = gci $checkFileFolder | select -exp BaseName

# determine all printers with error files existing
$printersWithErrorFiles = $prnStatus | where {$currentErrorFiles -contains $_.Name}

# update table with all printers that have errors listed
$printersWithErrorFiles | foreach {
    $checkFile = $checkFileFolder + $_.Name + ".txt"
    $firstErrorTime = Get-Content $checkFile
    $uptime = New-TimeSpan -Start $firstErrorTime -end (Get-Date)
    $UptimeString = "$($uptime.Days) Days, $($uptime.Hours) Hours, $($uptime.Minutes) Mins"
    $prnStatus | where Name -EQ $_.Name | foreach {$_.LastErrorTime = $firstErrorTime}
    $prnStatus | where Name -EQ $_.Name | foreach {$_.Uptime = $UptimeString}
}
# convert status table into one that links the Name column to the printer's IP address, convert the status into human-readable format, then convert to HTML table
$table = $prnStatus | select @{name="Name";Expression={"<a href='http://{0}/'><{1}</a>" -f $_.HostAddress,$_.Name}}, @{name="Status";Expression={customStatus $_.Status}}, Jobs, Location, Comment, LastErrorTime, Uptime | ConvertTo-Html -Fragment

#endregion

#region HTML generation
$title = "$serverName Printer Status"
$reportDescription = ""
$head = @"
<meta http-equiv="refresh" content="30">
<Title>$title | Online: $($onlinePrinters.Count) | Warning: $(($warningStatusPrinters | select -exp Name).Count) | Error: $(($errorStatusPrinters | select -exp Name).Count) </title>
<style>
body { background-color: #white; font-family: Segoe UI, Sans-Serif; font-size: 11pt; }
td, th, table { border: 1px solid grey; border-collapse:collapse; }
h1, h2, h3, h4, h5, h6 { font-family: Segoe UI, Segoe UI Light, Sans-Serif; font-weight: lighter; }
h1 { font-size: 26pt; }
h4 { font-size: 14pt; }
th { color: #383838; background-color: #lightgrey; text-align: left; }
table, tr, td, th { padding: 2px; margin: 0px; }
table { width: 95%; margin-left: 5px; margin-bottom: 10px; }
tr:nth-child(odd) { background-color: #e8f2f3; }
.green { color: green }
.orange { color: orange }
.red { color: red }
</style>
<h1>$title</h1>
"@

#region CSS class conditional formatting
$ptrnReady            = "<td>Ready</td>"
$ptrnReadyRepl        = "<td class=green>Ready</td>"
$ptrnOtherErr         = "<td>Other Error</td>"
$ptrnOtherErrRepl     = "<td class=red>Other Error</td>"
$ptrnPaused           = "<td>Paused</td>"
$ptrnPausedRepl       = "<td class=red>Paused</td>"
$ptrnLowPaper         = "<td>Low Paper</td>"
$ptrnLowPaperRepl     = "<td class=orange>Low Paper</td>"
$ptrnNoPaper          = "<td>Low Paper</td>"
$ptrnNoPaperRepl      = "<td class=red>Low Paper</td>"
$ptrnLowToner         = "<td>Low Toner</td>"
$ptrnLowTonerRepl     = "<td class=orange>Low Toner</td>"
$ptrnNoToner          = "<td>No Toner</td>"
$ptrnNoTonerRepl      = "<td class=red>No Toner</td>"
$ptrnDoorOpen         = "<td>Door Open</td>"
$ptrnDoorOpenRepl     = "<td class=red>Door Open</td>"
$ptrnPaperJam         = "<td>Paper Jam</td>"
$ptrnPaperJamRepl     = "<td class=red>Paper Jam</td>"
$ptrnOffline          = "<td>Offline</td>"
$ptrnOfflineRepl      = "<td class=red>Offline</td>"
$ptrnOutBinFull      = "<td>Output Bin Full</td>"
$ptrnOutBinFullRepl   = "<td class=red>Output Bin Full</td>"
$ptrnPaperProblem     = "<td>Paper Problem</td>"
$ptrnPaperProblemRepl = "<td class=red>Paper Problem</td>"
$ptrnCantPrtPage      = "<td>cannot Print Page</td>"
$ptrnCantPrtPageRepl  = "<td class=red>cannot Print Page</td>"
$ptrnUserIntReq       = "<td>User Intervention Required</td>"
$ptrnUserIntReqRepl   = "<td class=red>User Intervention Required</td>"
$ptrnUnknown          = "<td>Server Unknown</td>"
$ptrnUnknownRepl      = "<td class=red>Server Unknown</td>"

$table = $table -replace $ptrnReady,$ptrnReadyRepl
$table = $table -replace $ptrnOtherErr,$ptrnOtherErrRepl
$table = $table -replace $ptrnPaused,$ptrnPausedRepl
$table = $table -replace $ptrnLowPaper,$ptrnLowPaperRepl
$table = $table -replace $ptrnNoPaper,$ptrnNoPaperRepl
$table = $table -replace $ptrnLowToner,$ptrnLowTonerRepl
$table = $table -replace $ptrnNoToner,$ptrnNoTonerRepl
$table = $table -replace $ptrnDoorOpen,$ptrnDoorOpenRepl
$table = $table -replace $ptrnPaperJam,$ptrnPaperJamRepl
$table = $table -replace $ptrnOffline,$ptrnOfflineRepl
$table = $table -replace $ptrnOutBinFull,$ptrnOutBinFullRepl
$table = $table -replace $ptrnPaperProblem,$ptrnPaperProblemRepl
$table = $table -replace $ptrnCantPrtPage,$ptrnCantPrtPageRepl
$table = $table -replace $ptrnUserIntReq,$ptrnUserIntReqRepl
$table = $table -replace $ptrnUnknown,$ptrnUnknownRepl

#endregion

#endregion HTML generation

$output = ConvertTo-Html -Head $head -Body $table -PostContent "Total printers: <b>$($table.count)</b> | Online: <b class=green>$($onlinePrinters.Count)</b> | Warning: <b class=orange>$(($warningStatusPrinters | select -exp Name).Count)</b> | ErrorL <b class=red>$(($errorStatusPrinters | select -exp Name).Count)</b> | Script runtime: $(Get-Date)"

# fix HTML generation formatting issues and export to file
$output -replace '&gt;','>' -replace '&lt;','<' -replace '&#39;',"'" | Out-File $htmlFile -Encoding ascii -Force
