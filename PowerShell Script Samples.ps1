# PowerShell Script Samples
# Written By Adam Elamrani 
# 01/10/2023

Import-Module ActiveDirectory
Install-Module MSOnline
Install-Module ExchangeOnlineManagement


# Output the 10 largest files stored outside of Desktop and Documents 
$files = Get-ChildItem -Exclude Desktop, Documents
$files | Get-ChildItem -Recurse | Sort-Object -Descending -Property Length | 
Select -First 10 Name, Fullname, @{name="GB"; expression={[Math]::Round($_.length / 1GB, 2)}}


# Output all files edited in the last 30 days
$workingDirectory = Get-Location
$fileName = "Files_To_Be_Deleted.csv"
$files | Get-ChildItem -Recurse | 
Where-Object {$_.LastWriteTime -ge ((Get-Date).AddDays(-30))} | Sort-Object LastWriteTime -DESC |
Export-CSV -Path ".\$fileName" -NoTypeInformation 


# Generate excel file of all files at risk sorted by last write time and display the 10 most recent
Write-Warning "Files you have worked on recently are currently at risk for deletion. 
Please review the full list in the excel file found here: $workingDirectory\$fileName"
$files | Get-ChildItem -Recurse | Sort-Object LastWriteTime -DESC | Select -First 10 |
FT Name, Fullname, LastWriteTime


# Get CPU information, utiliztion, and top 10 unique consuming processes
$cpuProperties = @("SystemName", "Name", "Status", "LoadPercentage", "CurrentClockSpeed", 
"MaxClockSpeed", "NumberOfCores", "NumberOfLogicalProcessors", "ThreadCount")
$cpu = Get-CimInstance -ClassName Win32_Processor 
$cpu | FT -Property $cpuProperties 
Get-Process | Sort-Object cpu -descending | Get-Unique | Select -first 10 | FT ProcessName,Id,cpu,Description


# Get RAM information and utilization 
$os = Get-Ciminstance Win32_OperatingSystem
$FreePhysicalMemory = [Math]::Round(($os.FreePhysicalMemory)/1MB, 2)
Write-OutPut "`nFreePhysicalMemory: $FreePhysicalMemory mb`n"
$TotalVisibleMemorySize = [Math]::Round(($os.TotalVisibleMemorySize)/1MB, 2)
Write-OutPut "TotalVisibleMemorySize: $TotalVisibleMemorySize mb`n"


# Get all disk volumes and utilization information for the C drive 
$volume = Get-Volume
$volume | Sort-Object -Property @{Expression = 'SizeRemaining'; Ascending = $TRUE} 
$disk = Get-CimInstance -ClassName Win32_LogicalDisk
$disk | Format-Table DeviceId, MediaType, @{name="Size"; expression={[Math]::Round(($_.Size/1GB), 2)}},
@{name="FreeSpace"; expression={[Math]::Round(($_.FreeSpace/1GB), 2)}}


# Get a list of all stopped services on the machine 
$services = Get-Service 
$services | Where-Object {$_.Status -eq "Stopped"} | FT


# Get all users with an active Visio license, sorted by last active date 
Connect-MsolService
$visioUsers = Get-MsolUser | Where-Object {($_.licenses).AccountSkuId -match "VISIOCLIENT"}
$visioUsers | Sort-Object -Property LastDirSyncTime | FT UserPrincipalName, DisplayName, LastDirSyncTime


# Output a list of all AD users who have not logged in in the past 90 days
$inactivityDate = (Get-Date).AddDays(-90)
$adUsers = Get-ADUser -filter * -Properties * 
$inactiveUsers = $adUsers | Where-Object {$_.LastLogonDate -le $inactivityDate}


# Loop though the list of inactive users and export an excel file with the name of their managers
$managers = @()
$outFileName = "Managers_of_Inactive_Users.csv"
Foreach ($user in $inactiveUsers){
    $managers += $_.Manager
}
$Managers | Export-CSV -Path ".\$outFileName" -NoTypeInformation


# Reset the password of all AD users with expired passwords 
# *elevated privileges required*
$tempPassword = ConvertTo-SecureString "p@ssw0rd" -AsPlainText -Force 
$expiredADUsers = Get-ADUser -filter * -Properties * | Where-Object {$_.PasswordExpired -eq "True"}
Foreach ($user in $adUsers){
    Set-ADAccountPassword -Identity $_.ObjectGUID -Reset -NewPassword $tempPassword -ChangePasswordAtLogon: $True
}


# Read a list of new user emails via text file and verifies valid email format via regex  
# Then adds the users to a Distribution List based on their department
# *elevated privileges required*
Connect-ExchangeOnline
$emailFilePath = Read-Host "`nPlease enter the location of the text file containing users emails"
$usersArray = Get-Content $emailFilePath
$userErrors = @()

Foreach ($user in $usersArray) {
    Try {
        $usersArray = Get-Content $emailFilePath
        If (($user -Match "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$") -and 
            ($adUser = Get-ADUser -Filter {EmailAddress -eq $user} -Properties *)) {
            Add-DistributionGroupMember -Identity $adUser.Department -Member $adUser.EmailAddress
        }
        Else {
            Write-Warning "Invalid email detected. $user not processed."
            $userErrors += $user
        }
        
    }
    catch {
          "Error occured processing $user"
          $userErrors += $user
    }
    Finally {
        $totalProcessedUsers = ($usersArray.Length - $userErrors.Length)
         
    }
}

Write-Output "`nSuccessfully processed $totalProcessedUsers users."
Write-Output "`nPlease review the following user email addresses for errors: "

Foreach ($user in $userErrors) {
    Write-Output "$user" 
}
