# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

$DueDays = -30
$Now = Get-Date

function CreateAuditLog {
    param ([PARAMETER(Mandatory=$TRUE,ValueFromPipeline=$FALSE)]
    [string]$Mailbox,
    [PARAMETER(Mandatory=$TRUE,ValueFromPipeline=$FALSE)]
    [string]$StartDate,
    [PARAMETER(Mandatory=$TRUE,ValueFromPipeline=$FALSE)]
    [string]$EndDate,
    [PARAMETER(Mandatory=$FALSE,ValueFromPipeline=$FALSE)]
    [string]$Subject,
    [PARAMETER(Mandatory=$False,ValueFromPipeline=$FALSE)]
    [switch]$IncludeFolderBind,
    [PARAMETER(Mandatory=$False,ValueFromPipeline=$FALSE)]
    [switch]$ReturnObject)
    BEGIN {
      [string[]]$LogParameters = @('Operation', 'LogonUserDisplayName', 'LastAccessed', 'DestFolderPathName', 'FolderPathName', 'ClientInfoString', 'ClientIPAddress', 'ClientMachineName', 'ClientProcessName', 'ClientVersion', 'LogonType', 'MailboxResolvedOwnerName', 'OperationResult')
    }
    END {
        if ($ReturnObject) {
            return $SearchResults
        }
        elseif ($SearchResults.count -gt 0) {
            #$Date = get-date -Format yyMMdd_HHmmss
            #$OutFileName = "AuditLogResults$Date.csv"
            #write-host
            #write-host -ForegroundColor green "Posting results to file: $OutfileName"
            #$SearchResults | export-csv $OutFileName -notypeinformation -encoding UTF8
        }
    }
    PROCESS
    {
        write-host -ForegroundColor green 'Searching Mailbox Audit Logs...'
        $SearchResults = @(search-mailboxAuditLog $Mailbox -StartDate $StartDate -EndDate $EndDate -LogonTypes Owner, Admin, Delegate -ShowDetails -resultsize 50000)
        write-host ForegroundColor green $SearchREsults.Count Total entries Found
        if (-not $IncludeFolderBind) {
            write-host ForegroundColor green 'Removing FolderBind operations.'
            $SearchResults = @($SearchResults | Where-Object {$.Operation -notlike 'FolderBind'})
            write-host ForegroundColor green Filtered to $SearchREsults.Count Entries
        }


        $SearchResults = @($SearchResults | Select-Object ($LogParameters + @{Name='Subject';e={if (($_.SourceItems.Count -eq 0) -or ($null -eq $_.SourceItems.Count)){$_.ItemSubject} else {($_.SourceItems[0].SourceItemSubject).TrimStart(" ")}}},
        @{Name='CrossMailboxOp';e={if (@("SendAs","Create","Update") -contains $_.Operation) {"N/A"} else {$_.CrossMailboxOperation}}}))

        $LogParameters = @('Subject') + $LogParameters + @('CrossMailboxOp')
        If ($Subject -ne '' -and $null -ne $Subject) {
            write-host -ForegroundColor green 'Searching for Subject: $Subject'
        $SearchResults = @($SearchResults | Where-Object {$.Subject -match $Subject -or $_.Subject -eq $Subject})
        write-host ForegroundColor green 'Filtered to $($SearchREsults.Count) Entries'
        }
        $SearchResults = @($SearchResults | Select-Object $LogParameters)
    }

}


# The tenant name (orgname.onmicrosoft.com) set in the Function App configuration
$tenant = $env:Tenant 

try {
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    # Execute the EXO commands you want here
    $audit = $env:auditList
    Write-Host "AuditLog:"
    Write-Host $audit -ForegroundColor Yellow

    $StartDate = $Now.AddDays($DueDays).TotalSeconds -le 0

    $mailboxes = $audit.Split(",")
    $data = ""
    foreach($mailbox in $mailboxes) {
        $data += CreateAuditLog -Mailbox $mailbox -StartDate $StartDate -EndDate $Now -ReturnObject -IncludeFolderBind
    }

    # Convert the data to CSV format and store it in a MemoryStream
    $memoryStream = New-Object System.IO.MemoryStream
    $streamWriter = New-Object System.IO.StreamWriter($memoryStream)
    $csv = $data | ConvertTo-Csv -NoTypeInformation
    $csv | ForEach-Object { $streamWriter.WriteLine($_) }
    $streamWriter.Flush()
    $memoryStream.Position = 0


    


}
catch {    
    # Implement error handling here    
    throw $_
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false
    Get-PSSession | Remove-PSSession
}


# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"
