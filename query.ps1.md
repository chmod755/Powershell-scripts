```powershell
if($AZConn.Account -eq $null){
	if ($global:Debug_Flag -eq 1) { Write-Host "AZConn : Not connected so connecting!" }
    $UPN = whoami /upn
    $global:AZConn = Connect-AzureAD -AccountId $UPN
    # use this if you need a prompt for a different domain
    # $global:AZConn = Connect-AzureAD
}

# test if our session is connected to MgGraph and if not then we connect
if (((Get-MgContext).Account) -eq $null){
	if ($global:Debug_Flag -eq 1) { Write-Host "MgGraph : Not connected so connecting!" }
    Connect-MgGraph
}

if (($global:Tconnreturn).Tenant -eq $NULL) {
   $UPN = whoami /upn
   $global:Tconnreturn = Connect-MicrosoftTeams -AccountId $UPN
}

# $UPN = whoami /upn
# Connect-AzureAD -AccountId $UPN
if ($global:Debug_Flag -eq 1) { Write-Host "=======================================" }
# escaping quotes - use double quoted string and put a backtick in front of double quotes and $ to escape them 
' ' 
" Get-AzureADReports tim.heslop[@unilever.com]"
" Get-AzureADReports mohamed.massoum | Select-String -Pattern `"^Emp`""
" Get-AzureADReports mohamed.massoum | Select-String -Pattern `"^Third`""
" Get-AzureADReports mohamed.massoum | Select-String -Pattern `"^Non`""
' '
"  Get-AzureADUserDirectReport -ObjectID  steve.mccrystal@unilever.com | select DisplayName, UserPrincipalName, JobTitle"
' '
" Get-ULUser-Managers tim.heslop[@unilever.com]"
' ' 
" Get-ULUser-Details tim.heslop[@unilever.com]"
" Process-ULUser-Details-Report list.txt"
" (this function processes an input file full of email addresses and outputs a .csv in the same place)"
' ' 
" Get-UserPresence-byEmail tim.heslop[@unilever.com]"
' Get-UserPresence-byDisplayName ''GB-RPS Gatehouse'''
' foreach ($phone in (Get-Content "C:\dev\phonemon\RCH_SIP.txt")) {Get-UserPresence-byEmail $phone}' 
' '
' for debugging info set this flag - $Debug_Flag = 1 (not much output)'
' '
" Get-Teams-Members 'UiM IT EU Experience IT' "
' ' 
" `$reports = Get-AzureADReports-simple mihail.staicu"
" foreach (`$name in `$reports) { (`$name + `" , `" + (Get-ULUser-Details `$name).JobTitle) }"
" foreach (`$name in `$reports) { if (`$name -notlike `"*(C)`") { (`$name + `" , `" + (Get-ULUser-Details `$name).JobTitle) } }"
" foreach (`$name in `$reports) { "
" `$name + `" , `" + (Get-MgUser -Filter `"UserPrincipalName eq '`$name'`" -CountVariable CountVar -ConsistencyLevel Eventual).JobTitle "
" } "
">>>>>>>> cd 'C:\Users\Tim.Heslop\OneDrive - Unilever\2023\CoPilot'"

Function Get-ULUser-Details {

Param ($emailaddress)

if($AZConn.Account -eq $null){
    $global:AZConn = Connect-AzureAD
}

# this to add @unilever.com on if no @ exists in the string
if (-not ($emailaddress -Match "@") ) {
    $emailaddress = $emailaddress + "@unilever.com"
}

Get-AzureADUser -ObjectID $emailaddress |Select-Object Mail, GivenName, Surname, DisplayName, City, Country, State, JobTitle, Department, PhysicalDeliveryOfficeName, UserPrincipalName, ObjectId

$extended = Get-AzureADUserExtension -ObjectID $emailaddress

$WL = $extended.extension_58ae4ee4020d4690b95b326c6084b737_extensionAttribute6
$CostCentre = $extended.extension_58ae4ee4020d4690b95b326c6084b737_extensionAttribute13
$acct_type =  $extended.extension_58ae4ee4020d4690b95b326c6084b737_employeeType

$mgr= Get-AzureADUserManager -ObjectID $emailaddress 
$mgr_name = $mgr.UserPrincipalName


"Work Level :  $WL"
"Cost Centre : $CostCentre"
"Manager :     $mgr_name"
"Type :        $acct_type"

}

function  Get-AzureADReports {
	
	# Microsoft Graph version
	# ((Get-MgUserDirectReport -UserId tim.heslop@unilever.com).AdditionalProperties).userPrincipalName
    
    param (

    [Parameter()]

    [string]$StartingName = 'joann.mccann@unilever.com',
    [string]$manager = ''

    )
    if (-not ($StartingName -Match "@") ) {
        $StartingName = $StartingName + "@unilever.com"
    }
    # this call queries AzureAD and returns an array of people reporting to the Starting Name
    $daMembers = Get-AzureADUserDirectReport -ObjectID  $StartingName

    # now look up the account type - not all accounts have a type coded
    try {
        # get the type
        $account_type = ((Get-AzureADUserExtension -ObjectID $StartingName).get_item("extension_58ae4ee4020d4690b95b326c6084b737_employeeType"))
    }
    catch {
        # Write-Warning $Error[0]
        # hard code if not coded
        $account_type ="Not Coded"
    }

     # this is the output line
     $account_type + "," + $manager + $StartingName


    if ($daMembers.count -eq 0 ) {
        # do nothing - this is the return
    } else {
        # create a hierarchy string
        $manager = $manager + $StartingName + ","
        # cycle through the reports and recursively call the function
        $daMembers | ForEach-Object {
                  Get-AzureADReports $_.UserPrincipalName $manager
        }
    }

}

Function  Get-AzureADReports-simple {
    
    param (

    [Parameter()]

    [string]$StartingName = 'joann.mccann@unilever.com',
    [string]$manager = ''

    )
    if (-not ($StartingName -Match "@") ) {
        $StartingName = $StartingName + "@unilever.com"
    }
    # this call queries AzureAD and returns an array of people reporting to the Starting Name
    $daMembers = Get-AzureADUserDirectReport -ObjectID  $StartingName

    # now look up the account type - not all accounts have a type coded
    try {
        # get the type
        $account_type = ((Get-AzureADUserExtension -ObjectID $StartingName).get_item("extension_58ae4ee4020d4690b95b326c6084b737_employeeType"))
    }
    catch {
        # Write-Warning $Error[0]
        # hard code if not coded
        $account_type ="Not Coded"
    }

    # this is the output line
    # $account_type + "," + $manager + $StartingName
	if ($account_type -eq "Employee") { 
		$StartingName
	} else
	{
		$StartingName + " (C)"
	}
	
    if ($daMembers.count -eq 0 ) {
        # do nothing - this is the return
    } else {
        # create a hierarchy string
        $manager = $manager + $StartingName + ","
        # cycle through the reports and recursively call the function
        $daMembers | ForEach-Object {
                  Get-AzureADReports-simple $_.UserPrincipalName $manager
        }
    }

}


Function Get-ULUser-Details-Report {

Param ($emailaddress)

if($AZConn.Account -eq $null){
    $global:AZConn = Connect-AzureAD
}

# this is a data cleanup cludge
# use regex to clean up email addresses with comments in brackets
# if the left bracket is there
if ($emailaddress -match '\(') {
	#replace all the contents of the brackets with nothing, includes the brackets
	# https://regex101.com/
	$emailaddress = $emailaddress -replace '\(.*\)',''
}

# this to add @unilever.com on if no @ exists in the string
if (-not ($emailaddress -Match "@") ) {
    $emailaddress = $emailaddress + "@unilever.com"
}

if ($global:Debug_Flag -eq 1) { Write-Host "$emailaddress : Processing" }

try {
	# this omits lines with #N/A in them
	# we avoid the system call
		if ($emailaddress -ne "#N/A") {
		$result = Get-AzureADUser -ObjectID $emailaddress `
		    	 |Select-Object Mail, GivenName, Surname, DisplayName, City, Country, State, `
				  JobTitle, Department, PhysicalDeliveryOfficeName, UserPrincipalName

		$extended = Get-AzureADUserExtension -ObjectID $emailaddress

		$WL = $extended.extension_58ae4ee4020d4690b95b326c6084b737_extensionAttribute6
		$CostCentre = $extended.extension_58ae4ee4020d4690b95b326c6084b737_extensionAttribute13
		$acct_type =  $extended.extension_58ae4ee4020d4690b95b326c6084b737_employeeType

		$mgr= Get-AzureADUserManager -ObjectID $emailaddress 
		$mgr_name = $mgr.UserPrincipalName
	}	
}
catch
{
	# 
	if ($global:Debug_Flag -eq 1) { Write-Host "$emailaddress : Thie email address does not exist" }
	# Write-Host $_
	$acct_type = "Not found in AzureAD"
}

$Out_object = [PSCustomObject]@{
	Email = $emailaddress
	City = $result.City
	Country = $result.Country
	State = $result.state
	JobTitle = $result.JobTitle
	Department = $result.Department
	Office = $result.PhysicalDeliveryOfficeName
	Work_Level = $WL
	Cost_Centre = $CostCentre
	Manager = $mgr_name
	Type = $acct_type
}

return $Out_object
}

Function Process-ULUser-Details-Report {

Param ($inputfile)

if ($inputfile -eq $NULL) {
	Write-Host "Input file not specified"
	return
}

if (-not(Test-Path -Path $inputfile -PathType Leaf)) {
	Write-Host "Input file $inputfile does not exist"
	return
}

$outdir = (Get-Item $inputfile).DirectoryName
$outfile = (Get-Item $inputfile).Basename
$outfile = $outdir + '\'+ $outfile + '.csv'
$outfile

if (Test-Path -Path $outfile -PathType Leaf) {
	Write-Host "Output file $outfile exists, any key to continue ^C to stop"
	$zzz = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
	Write-Host "Processing ..."
}

$list = Get-Content -Path $inputfile

# force the type of the output to be an array of objects
$out_array = @()

# cycle over the list
foreach ($emailaddress in $list) {
	# the function returns an object so we can just call it and concatenate the response
	# replaced for big jobs with line by line output
	# $out_array += Get-ULUser-Details-Report ($emailaddress)
	# 
	# this version sets the output variable
	$out_array = Get-ULUser-Details-Report ($emailaddress)
	# and in this version we write inside the loop
	if ($global:Debug_Flag -eq 1) { Write-Host $out_array }
	$out_array | Export-CSV  $outfile -NoTypeInformation -Append
}

# this is the line that goes with the export all the data at the end version
# $out_array | Export-CSV  $outfile -NoTypeInformation 

}

Function Get-UserPresence-byEmail {

Param ($emailaddress)

if (((Get-MgContext).Account) -eq $null){
    Connect-MgGraph
}

# this to add @unilever.com on if no @ exists in the string
if (-not ($emailaddress -Match "@") ) {
    $emailaddress = $emailaddress + "@unilever.com"
}

$UserDets = Get-MgUser -UserId $emailaddress
$UserDisplayName = "'" + $UserDets.DisplayName + "'"

$Presence = (Get-MgUserPresence -UserId ($UserDets.Id)).Availability
"$UserDisplayName, $emailaddress, $Presence"
}

Function Get-UserPresence-byDisplayName {

Param ($displayname)

if (((Get-MgContext).Account) -eq $null){
    Connect-MgGraph
}

$displayname2 = $displayname -replace "'","''"

$UserDets = Get-MgUser -Count usercount -ConsistencyLevel eventual -Filter "DisplayName eq '$displayname2'"
$UserId = $UserDets.Id
$UserDisplayName = "'" + $UserDets.DisplayName + "'"

if ($usercount -eq 0) {
	"$displayname,NOT FOUND"
} else {
	#$Presence = (Get-MgUserPresence -UserId ((Get-MgUser -Filter "DisplayName eq '$displayname'").Id)).Availability
	$Presence = (Get-MgUserPresence -UserId $UserId).Availability
	#"$displayname,$Presence"
	"$UserDisplayName, $emailaddress, $Presence"
}

}

Function Get-Teams-Members {

Param ($DisplayName)

if ($DisplayName -eq $NULL) {
	Write-Host "No Team Name specified"
	return
}

	if (($global:Tconnreturn).Tenant -eq $NULL) {
	   $UPN = whoami /upn
	   $global:Tconnreturn = Connect-MicrosoftTeams -AccountId $UPN
	}

	$TeamID = 	(Get-Team -User $UPN | where DisplayName -eq $Displayname).GroupId

	$users = Get-TeamUser -GroupId $TeamId

	$email_list = ""

	foreach ($user in $users) {

		$email_list = $email_list + $user.User + " ; "

	}

	$email_list

}

Function Get-ULUser-Managers {

Param ($emailaddress)
	
if($AZConn.Account -eq $null){
	$global:AZConn = Connect-AzureAD
}

# this to add @unilever.com on if no @ exists in the string
if (-not ($emailaddress -Match "@") ) {
	$emailaddress = $emailaddress + "@unilever.com"
}

# check user exists
try {
$usr = Get-AzureAdUser -ObjectId $emailaddress
}
catch {
 "$emailaddress not found"
 return
}

$mgr= Get-AzureADUserManager -ObjectID $emailaddress 

#$emailaddress + "," + $mgr.UserPrincipalName + "," + $mgr.JobTitle
#$mgr.UserPrincipalName + ", " + $mgr.JobTitle

if ($mgr.UserPrincipalName -like 'Hein.Schumacher@unilever.com' ) {
		# return from recursion	
		"Management hierarchy"
		"===================="
		$mgr.UserPrincipalName + "," + $mgr.JobTitle
	} else
	{
		# recursion call
		Get-ULUser-Managers ($mgr.UserPrincipalName)
		$mgr.UserPrincipalName + "," + $mgr.JobTitle
} # end of if 

} # end of function

# read all child keys (*) from all four locations and do not emit
# errors if one of these keys does not exist:
# $keys = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
#                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
#                    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
#                    'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction Ignore
# $keys | Where-Object DisplayName `
#	  | Select-Object -Property DisplayName, DisplayVersion, UninstallString, InstallDate `
# 	  | Sort-Object -Property DisplayName

```

