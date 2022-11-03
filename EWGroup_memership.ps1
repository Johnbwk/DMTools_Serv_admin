# 
# 
# 
# 
# 

# day/month/year
$DateTime =Get-Date -f "ddMMyyyy"

# File name
$CSVFile = "C:\xom\spttemp\EWGroups_" + $DateTime + ".csv"

# Create empty array for CSV data
$CSVOutput = @()

# distinguishedName can be used as searchbase, you can use one DN or multiple DNs
# Or use the root domain like DC=exoip,DC=local if looking for domain information
#$DNs = @(
#    "DC=exoip,DC=local"
#)
#$file = get-content -path C:\xom\SPTTemp\xxxxxxxx.xx""
$groups = "DMTools.Support.ug"

# Create progress bar for groups being queried
$i = 0
$tot = $groups.count

foreach($group in $groups)
{
    
    # Set up progress bar
    $i++
    # $status = "{0:N0}" -f ($i / $tot * 100)
    # Write-Progress -Activity "Exporting AD Groups" -status "Processing Group $i of $tot : $status% Completed" -PercentComplete ($i / $tot * 100)

    #query each user info inside each group inside $groups 
    $users = (get-adgroup $group -server na.xom.com:3268 -properties *).members
    foreach($u in $users)
    {
        $status = "{0:N0}" -f ($i / $tot * 100)
        Write-Progress -Activity "Exporting AD Groups" -status "Processing Group $i of $tot : $status% Completed" -PercentComplete ($i / $tot * 100)
    
        $user = get-aduser $u -server na.xom.com:3268 -properties *

        # Create hash table and include values
        $HashTab = $null
        $HashTab = [ordered]@{
            "Name"          = $user.Name
            "Email"         = $user.mail
            "Principal"     = $user.UserPrincipalName
            "DisplayName"   = $user.DisplayName
            "LastLogonDate" = $user.LastLogonDate
            # query just DN's for group membership as they appear on memberof prop of the user, much faster
            # "MemberOf"      = $user.memberof
            # To query group membership names only and include them separated by ";"
            "MemberOf"      = ($user.memberof | % { (Get-ADObject $_ -server na.xom.com).Name }) -join ";
            "
        }
            # Include hash table to CSV array
            $CSVOutput += New-Object PSObject -Property $HashTab
    }


}

    # # Create hash table and include values
    # $HashTab = $null
    # $HashTab = [ordered]@{
    #     "Name"     = $user.Name
    #     "Category" = $user.GroupCategory
    #     "Scope"    = $user.GroupScope
    #     "MemberOf"  = $user.memberof
    #     "Users"    = $ADGroup.member
    # }

    # # Include hash table to CSV array
    # $CSVOutput += New-Object PSObject -Property $HashTab

# Export to file
$CSVOutput | Sort-Object Name | Export-Csv -Encoding UTF8 -Path $CSVFile -NoTypeInformation #-Delimiter ";"

