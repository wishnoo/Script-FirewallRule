
# This makes all non terminating errors stop
$ErrorActionPreference = 'Stop'

# Array to store all old locations in order to be processed in this case to be deleted from the firewall.
$deprecatedPathLocationArray = @("AppData\Local\Microsoft\Teams\Current\Teams.exe")
# Variable to store the path string with the current location of the rule.
$currentPathLocation = ""

# Action value depicts the action filter we can choose e.g.: "Allow","Block" 
$actionToDelete = "Block"

#Action value required to add rules in the firewall
$actionToAdd = "Allow"

$users = Get-ChildItem (Join-Path -Path $env:SystemDrive -ChildPath 'Users') -Exclude 'Public', 'ADMINI~*'

if ($null -ne $users) {
    try {
        # Loop through the Array to delete existing old path rules
        Write-Host "`n START: Section to delete old path rules `n"
        if ($deprecatedPathLocationArray) {
            foreach ($path in $deprecatedPathLocationArray) {
                Write-Host "Path in array : $($path)"
                # Iterate through each user folder and remove rules with respect to the user regardless of the action
                foreach ($user in $users) {
                    # Combine the user folder bath with path provided in the deprecated location array
                    $fullProgramPath = Join-Path -Path $user.FullName -ChildPath $path
                    $applicationObject_ProgramItem = Get-NetFirewallApplicationFilter -Program $fullProgramPath
                    $applicationObject_ProgramItem | Remove-NetFirewallRule
                    Write-Host "Rule with path '$($fullProgramPath)' deleted"
                }
            }
        }
        else {
            Write-Host "No Old Path Location"
        }
    }
    catch {
        Write-Host "`n START: Catch Section to delete old path rules `n "
        Write-Host "Exception.Message [$($_.Exception.Message)]"
        Write-Host "`n END: Catch Section to delete old path rules `n"
    }
    finally{
        Clear-Variable fullProgramPath
        Write-Host "`n END: Section to delete old path rules `n"
    }
    
    
    
    
    # Remove firewall rules with appropriate action and current Path
    
    # Get-NetFirewallRule allows us to return instances of firewall rules that matches the action "Allow".
    # Get-NetFirewallApplicationFilter accepts AssociatedNetFirewallRule <CimInstance> through the pipeline (Byvalue) and it returns objects filter objects based on the input.
    try {
        Write-Host "`n START: Section to delete current path rules with specific Action `n"
        if (Get-NetFirewallRule -Action $actionToDelete) {
            # Iterate through each user folder and remove rules with respect to the user and the action
            foreach ($user in $users) {
                # Combine the user folder bath with path provided in the current location variable
                $fullProgramPath = Join-Path -Path $user.FullName -ChildPath $currentPathLocation

                $applicationObject_Action = Get-NetFirewallRule -Action $actionToDelete | Get-NetFirewallApplicationFilter
        
                # Return the filter objects for the program path that team is currently in.
                $applicationObject_Program = $applicationobject_Action | Where-Object {$_.Program -eq $fullProgramPath}
            
                $applicationObject_Program | Remove-NetFirewallRule
                Write-Host "Rule with path '$($fullProgramPath)' deleted"
            }
        }
        else {
            Write-Host "No firewall rules with give action to delete"
        }
    }
    catch {
        Write-Host "`n START: Catch section to delete current path rules with specific Action `n"
        Write-Host "Exception.Message [$($_.Exception.Message)]"
        Write-Host "`n END: Catch section to delete current path rules with specific Action `n"
    }
    finally{
        Clear-Variable fullProgramPath
        Write-Host "`n END: Section to delete current path rules with specific Action `n"
    }


    # Add new rules to the firewall based on the current location path.
    try {
        Write-Host "`n BEGIN: Section to Add new rules based on the current path `n"
        if ($currentPathLocation) {
            foreach($user in $users){
                $fullProgramPath = Join-Path -Path $user.FullName -ChildPath $currentPathLocation
                Write-Host "Full program path: $fullProgramPath"
                if (Test-Path $fullProgramPath) {
                    if (-not (Get-NetFirewallApplicationFilter -Program $fullProgramPath -ErrorAction SilentlyContinue)) {
                        $ruleName = "Microsoft Teams for user $($user.Name)"
                        "UDP", "TCP" | ForEach-Object { New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Profile Domain -Program $fullProgramPath -Action $actionToAdd -Protocol $_ } | Out-Null
                        Write-Host "Rule for $($ruleName) added"
                    }
                }
                else {
                    Write-Host "Path does not exist"
                }
            }       
        }
    }
    catch {
        Write-Host "`n BEGIN: Catch Section to Add new rules based on the current path `n"
        Write-Host "Exception.Message [$($_.Exception.Message)]"
        Write-Host "`n END: Catch Section to Add new rules based on the current path `n"
    }
    finally{
        Clear-Variable ruleName
        Clear-Variable fullProgramPath
        Write-Host "`n END: Section to Add new rules based on the current path `n"

    }
    
}