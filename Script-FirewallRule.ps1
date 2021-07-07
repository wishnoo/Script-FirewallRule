
# This makes all non terminating errors stop
$ErrorActionPreference = 'Stop'

# Array to store all old locations in order to be processed in this case to be deleted from the firewall.
$deprecatedPathLocationArray = @("AppData\Local\Microsoft\Teams\Current\Teams.exe")
# Variable to store the path string with the current location of the rule.
$currentPathLocation = "AppData\Local\Microsoft\Teams\Current\Teams.exe"

# Action value depicts the action filter we can choose e.g.: "Allow","Block" 
$action_value = "Block"

$users = Get-ChildItem (Join-Path -Path $env:SystemDrive -ChildPath 'Users') -Exclude 'Public', 'ADMINI~*'

if ($null -ne $users) {
    try {
        # Loop through the Array to delete existing old path rules
        if ($deprecatedPathLocationArray) {
            Write-Host "Inside loop"
            foreach ($path in $deprecatedPathLocationArray) {
                Write-Host "Path in array : $($path)"
                # Iterate through each user folder and remove rules with respect to the user regardless of the action
                foreach ($user in $users) {
                    # Combine the user folder bath with path provided in the deprecated location array
                    $fullProgramPath = Join-Path -Path $user.FullName -ChildPath $path
                    Write-Host "This is the program path: $fullProgramPath"
                    $applicationObject_ProgramItem = Get-NetFirewallApplicationFilter -Program $fullProgramPath
                    $applicationObject_ProgramItem | Remove-NetFirewallRule
                }
            }
        }
    }
    catch {
        Write-Host "Exception.Message [$($_.Exception.Message)]"
    }
    
    
    
    
    # Remove firewall rules with appropriate action and current Path
    
    # Get-NetFirewallRule allows us to return instances of firewall rules that matches the action "Allow".
    # Get-NetFirewallApplicationFilter accepts AssociatedNetFirewallRule <CimInstance> through the pipeline (Byvalue) and it returns objects filter objects based on the input.
    try {
        if (Get-NetFirewallRule -Action $action_value) {
            # Iterate through each user folder and remove rules with respect to the user and the action
            foreach ($user in $users) {
                # Combine the user folder bath with path provided in the current location variable
                $fullProgramPath = Join-Path -Path $user.FullName -ChildPath $currentPathLocation
                Write-Host "This is the program path: $fullProgramPath"

                $applicationObject_Allow = Get-NetFirewallRule -Action $action_value | Get-NetFirewallApplicationFilter
        
                # Return the filter objects for the program path that team is currently in.
                $applicationObject_Program = $applicationobject_Allow | Where-Object {$_.Program -eq $fullProgramPath}
            
                $applicationObject_Program | Remove-NetFirewallRule
            }

        }
        else {
            Write-Host "No firewall rules"
        }
    }
    catch {
        Write-Host "Exception.Message [$($_.Exception.Message)]"
    }
}