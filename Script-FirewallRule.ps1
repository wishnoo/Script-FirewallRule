## Overview:
## This script is used to add rules to the firewall. The reason this is needed is because Microsoft is running Teams from the AppData folder of each user when it starts
## a call.It also has dynamic sections that handles possible deprecated paths and blocked rules incorrectly created because of the user clicking cancel. 


## This makes all non terminating errors stop
$ErrorActionPreference = 'Stop'

## This array is used to store Teams old file locations. For example, Teams is currently in Appdata if that is changed we will
## not know what to delete. So, we store the old locations here in order to delete them. 
$deprecatedPathLocationArray = @("AppData\Local\Microsoft\Teams\Current\Teams.exe")

## Variable to store the path string with the current location of the rule.
$currentPathLocation = ""

## Action value depicts the action filter we can choose e.g.: "Allow","Block" 
$actionToDelete = "Block"

## Action value required to add rules in the firewall
$actionToAdd = "Allow"

$users = Get-ChildItem (Join-Path -Path $env:SystemDrive -ChildPath 'Users') -Exclude 'Public', 'ADMINI~*'

if ($null -ne $users) {


## Section: Delete old paths for the program TEAMS
## How it works:
## Iterate through the progam field in the firewall in order to find deprecated paths.
## Iterate through the users folder to create proper user paths.
## Join current path with the users path
## Use Get-NetFirewallApplicationFilter to find the rules with the created path.
## Delete all the rules which we found

    try {
        ## Loop through the Array to delete existing old path rules
        Write-Host "`n START: Section to delete old path rules `n"
        if ($deprecatedPathLocationArray) {
            $counter = 0
            foreach ($path in $deprecatedPathLocationArray) {
                Write-Host "OUTPUT: Path in array : $($path)"
                ## Iterate through each user folder and remove rules with respect to the user regardless of the action
                foreach ($user in $users) {
                    ## Combine the user folder path with path provided in the deprecated location array
                    $fullProgramPath = Join-Path -Path $user.FullName -ChildPath $path
                    ## Use the ErrorAction Parameter to avoid catching the error and continue the search when no rules are found.
                    ## The ErrorVariable parameter is used to store the error in a custom name field which could be later 
                    $applicationObject_ProgramItem = Get-NetFirewallApplicationFilter -Program $fullProgramPath -ErrorAction SilentlyContinue -ErrorVariable oldPathSearchError
                    If($applicationObject_ProgramItem){
                        $applicationObject_ProgramItem | Remove-NetFirewallRule
                        Write-Host "OUTPUT: Rule with path '$($fullProgramPath)' deleted"
                        $counter = 1
                    }
                }
            }
            if ($counter -eq 0) {
                Write-Host "OUTPUT: No Rule Deleted."
            }
             
        }
       
        else {
            Write-Host "OUTPUT: No Old Location Path to delete."
        }

    }
    catch {
        Write-Host "`n START: Catch Section to delete old path rules `n "
        Write-Host "Exception.Message [$($_.Exception.Message)]"
        Write-Host "Invocation Command [$($_.InvocationInfo.MyCommand.Name)]"
        Write-Host "FullyQualifiedErrorId [$($_.FullyQualifiedErrorId)]"
        Write-Host "`n END: Catch Section to delete old path rules `n"
    }
    finally{
        Write-Host "`n END: Section to delete old path rules `n"
    }
    
## Section: Delete the block rules
## This is needed if a user clicks on the cancel button on the TEAMS popup becasue TEAMS creates block rules when this action is taken
## and we need to remove these block rules. 
## How it works:
## Check the firewall rules to find the block rules.
## Iterate through the progam field in the firewall in order to find deprecated paths.
## Iterate through the users folder to create proper user paths.
## Get all the firewall block rules with Get-NetFirewallRule and pass it as an argument to Get-NetFirewallApplicationFilter to create an object of all the rules.
## Use the where object to find if there are any of our paths in the block rules.
## If you find any, delete them. 
## NOTES:
    # -Get-NetFirewallRule allows us to return instances of firewall rules that matches the action "Allow".
    # -Get-NetFirewallApplicationFilter accepts AssociatedNetFirewallRule <CimInstance> through the pipeline (Byvalue) and it returns objects filter objects based on the input.

    try {
        Write-Host "`n START: Section to delete current path rules with specific Action `n"
        ## Use the ErrorAction Parameter to avoid catching the error and continue the search when no rules are found.
        ## The ErrorVariable parameter is used to store the error in a custom name field which could be later 
        $rulesWithAction = Get-NetFirewallRule -Action $actionToDelete
        # $rulesWithAction = Get-NetFirewallRule -Action $actionToDelete -ErrorAction SilentlyContinue -ErrorVariable currentPathBlockSearchError
        if ($rulesWithAction) {
            ## Input object of Get-NetFirewallRule to Get-NetFirewallApplicationFilter to obtain properties based on Get-NetFirewallApplicationFilter
            $applicationObject_Action = $rulesWithAction | Get-NetFirewallApplicationFilter
            ## Iterate through each user folder and remove rules with respect to the user and the action
            foreach ($user in $users) {
                ## Combine the user folder path with path provided in the current location variable
                $fullProgramPath = Join-Path -Path $user.FullName -ChildPath $currentPathLocation
        
                ## Return the filter objects for the program path that team is currently in.
                $applicationObject_Program = $applicationobject_Action | Where-Object {$_.Program -eq $fullProgramPath}

                if ($applicationObject_Program) {
                    $applicationObject_Program | Remove-NetFirewallRule
                    Write-Host "OUTPUT: Rule with path '$($fullProgramPath)' deleted"        
                }
            }
        }
        else {
            Write-Host "OUTPUT: No firewall rules with given action to delete"
        }
    }
    catch {
        Write-Host "`n START: Catch section to delete current path rules with specific Action `n"
        Write-Host "Exception.Message [$($_.Exception.Message)]"
        Write-Host "Invocation Command [$($_.InvocationInfo.MyCommand.Name)]"
        Write-Host "FullyQualifiedErrorId [$($_.FullyQualifiedErrorId)]"

        Write-Host "`n END: Catch section to delete current path rules with specific Action `n"
    }
    finally{
        Write-Host "`n END: Section to delete current path rules with specific Action `n"
    }

## Section: Add the appropriate rules. 
## How it works:
## Check if the current path locations exist
## Iterate through the progam field in the firewall in order to find deprecated paths.
## Iterate through the users folder to create proper user paths.
## Test if the path exists and then test if the paths exists. If yes, do not create the rule.
## Else create the rule(s) 

    try {
        Write-Host "`n BEGIN: Section to Add new rules based on the current path `n"
        if ($currentPathLocation) {
            $counter = 0
            foreach($user in $users){
                $fullProgramPath = Join-Path -Path $user.FullName -ChildPath $currentPathLocation
                if (Test-Path $fullProgramPath) {
                    if (-not (Get-NetFirewallApplicationFilter -Program $fullProgramPath -ErrorAction SilentlyContinue)) {
                        $ruleName = "Microsoft Teams for user $($user.Name)"
                        "UDP", "TCP" | ForEach-Object { New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Profile Domain -Program $fullProgramPath -Action $actionToAdd -Protocol $_ } | Out-Null
                        Write-Host "OUTPUT: Rule for $($ruleName) for the program path $($fullProgramPath) added"
                        $counter = 1
                    }
                }
            }
            if ($counter -eq 0) {
                Write-Host "OUTPUT: No Firewall Rule Added."
            }
        }
        else {
            Write-Host "OUTPUT: Path is empty"
        }
    }
    catch {
        Write-Host "`n BEGIN: Catch Section to Add new rules based on the current path `n"
        Write-Host "Exception.Message [$($_.Exception.Message)]"
        Write-Host "Invocation Command [$($_.InvocationInfo.MyCommand.Name)]"
        Write-Host "FullyQualifiedErrorId [$($_.FullyQualifiedErrorId)]"
        Write-Host "`n END: Catch Section to Add new rules based on the current path `n"
    }
    finally{
        
        Write-Host "`n END: Section to Add new rules based on the current path `n"

    }
    
}