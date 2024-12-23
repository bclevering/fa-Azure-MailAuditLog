# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}


# The tenant name (orgname.onmicrosoft.com) set in the Function App configuration
$tenant = $env:Tenant 

try {
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    # Execute the EXO commands you want here
    $audit = $env:auditList
    Write-Host "AuditLog:"
    Write-Host $audit -ForegroundColor Yellow
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
