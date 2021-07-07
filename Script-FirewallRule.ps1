
# Get-NetFirewallRule allows us to return instances of firewall rules that matches the action "Allow".
# Get-NetFirewallApplicationFilter accepts AssociatedNetFirewallRule <CimInstance> through the pipeline (Byvalue) and it returns objects filter objects based on the input.
$AllowApplicationobject = Get-NetFirewallRule -Action "Allow" | Get-NetFirewallApplicationFilter

# Return the filter objects for the program path that team is currently in.
$AllowApplicationobject | Where-Object {$_.Program -eq "C:\users\wishnoo\appdata\local\microsoft\teams\current\teams.exe"} | Format-List *

