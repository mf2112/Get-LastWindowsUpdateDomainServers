# This script will get the last Windows Update and the time it was installed from all domain servers.

# Get all active domain servers
$staledate = (Get-Date).AddDays(-90)
$computers = Get-ADComputer -Filter {(OperatingSystem -Like "*Server*") -and (Enabled -eq $True) -and (LastLogonDate -ge $staledate) -and (Modified -ge $staledate) -and (PasswordLastSet -ge $staledate) -and (whenChanged -ge $staledate)} | Select name -expandproperty name | Where {(Resolve-DNSName $_.name -ea 0) -and (Test-Connection -ComputerName $_.Name -Count 1 -ea 0)} | Sort

# Initialize export array
$ToExport = @()

# Loop through all the computers to check them
ForEach ($computer in $computers) {
    # Try getting the hotfix info and catch the error if one fails
    Try {
    $last = get-hotfix -computername $computer | Select hotfixid,installedon | Sort -Descending installedon | Select -First 1
    }
    Catch {
    Write-Warning "Can't connect to $computer"
    }

# Set the output variables
$date = ($last.installedon).ToString()
$pdate = $date.Split(" ")

# Add the entry to the export array    
$entry = $computer + ";" + $last.hotfixid + ";" + $pdate[0]
$ToExport += $entry

# End of computers loop
}

# Write the final output
$ToExport | Add-Content -Path "C:\Temp\lastwindowsUpdates.csv"
