
# This makes all non terminating errors stop
$ErrorActionPreference = 'Stop'

# Array to store all old locations in order to be processed in this case to be deleted from the firewall.
$deprecatedPathLocationArray = @()
# Variable to store the path string with the current location of the rule.
$currentPathLocation = "C:\users\wishnoo\appdata\local\microsoft\teams\current\teams.exe"

$action_value = "Block"

# Loop through the Array to delete existing old path rules
if ($deprecatedPathLocationArray) {
    Write-Host "Inside loop"
    foreach ($path in $deprecatedPathLocationArray) {
        $applicationObject_ProgramItem = Get-NetFirewallApplicationFilter -Program $path
        Remove-NetFirewallRule -InputObject $applicationObject_ProgramItem
    }
}



# Remove firewall rules with appropriate action and current Path

# Get-NetFirewallRule allows us to return instances of firewall rules that matches the action "Allow".
# Get-NetFirewallApplicationFilter accepts AssociatedNetFirewallRule <CimInstance> through the pipeline (Byvalue) and it returns objects filter objects based on the input.
try {
    if (Get-NetFirewallRule -Action $action_value) {
        $applicationObject_Allow = Get-NetFirewallRule -Action $action_value | Get-NetFirewallApplicationFilter
    
        # Return the filter objects for the program path that team is currently in.
        $applicationObject_Program = $applicationobject_Allow | Where-Object {$_.Program -eq $currentPathLocation}
    
        $applicationObject_Program | Remove-NetFirewallRule
    }
    else {
        Write-Host "No firewall rules"
    }
}
catch {
    Write-Host "Exception.Message [$($_.Exception.Message)]"
}

