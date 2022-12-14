# #######################################################################################################
# Script to be used by DMTools Support team to query groups used by EnterWorks application
# and all their member users.
# 
# Before starting check and update accordingly the variables below 
# $CSVFile ----> set the name and location to save output file
# $groups ----> file with all groups to be queried by the script, needs to be updated according to report from EW system access
# 
# For questions or enhancements please contact Jonathan Bolwerk or email our team's PDL at DMTools.Support.UG
# #######################################################################################################

Import-Module -Name ActiveDirectory

# file with groups to be queried by this script
# $groups = get-content -path C:\xom\SPTTemp\xxxxxxxx.xx""
$groups = get-content -path "F:\DMTools_Serv_admin\DMTools_Serv_admin\groups.txt"

# day/month/year
$DateTime = Get-Date -f "ddMMyyyy"

# File name
$CSVFile = "F:\emit_enterworks_solutions\Licensing information\EWGroups_" + $DateTime + ".csv"

# Create empty array for CSV data
$CSVOutput = @()

# Create variables for progress bar for groups being queried
$i = 0
$tot = $groups.count

foreach($group in $groups)
{    
    # Start progress bar
    $i++
    $status = "{0:N0}" -f ($i / $tot * 100)
    Write-Progress -Activity "Exporting AD Groups" -status "Processing Group $i of $tot : $status% Completed" -PercentComplete ($i / $tot * 100)

    #query each user info inside each group inside $groups 
    $users = (get-adgroup $group -server xom.com:3268 -properties *).members

    # Set up progress bar for users inside group
    $ui = 0
    $utot = $users.count

    foreach($u in $users)
    {

        $ui++
        $ustatus = "{0:N0}" -f ($ui / $utot * 100)
        Write-Progress -Activity "Exporting AD users from Group: $group" -status "Processing user $ui of $utot : $status% Completed" -PercentComplete ($ui / $utot * 100)
    
        $user = get-aduser $u -server xom.com:3268 -properties *

        # Create hash table and include values
        $HashTab = $null
        $HashTab = [ordered]@{
            "EW AD Group"   = $group
            # "Domain"        = $user.Domain
            "Lan ID"        = $user.samaccountname
            "UPN"           = $user.UserPrincipalName
            "Full Name"     = $user.DisplayName
            "Email Address" = $user.mail
            "Last Logon Date" = $user.LastLogonDate
            # query just DN's for user group membership as they appear on memberof prop of the user, much faster
                # "MemberOf"      = $user.memberof
            # To query user group membership names only and include them separated by ";"
                # "MemberOf"      = ($user.memberof | foreach-object { (Get-ADObject $_ -server xom.com:3268).Name }) -join ";
            # "
        }
            # Include hash table to CSV array
            $CSVOutput += New-Object PSObject -Property $HashTab
    }


}

# Export to file
$CSVOutput | Sort-Object Name | Export-Csv -Encoding UTF8 -Path $CSVFile -NoTypeInformation #-Delimiter ";"

