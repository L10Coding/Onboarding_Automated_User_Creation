# Onboarding_Automated_User_Creation

# Module Description
Built a PowerShell script that reads new hire data from a CSV file, automatically creates Microsoft 365 accounts, assigns department-based group memberships, adds every user to an All Employees base group, and exports a full onboarding audit report. Designed to replace manual onboarding that takes 30 minutes per user with an automated process that runs in seconds.
**Step 1 — Connected to Microsoft Graph with required scopes**
- Connected using three scopes
- User.ReadWrite.All — permission to create and modify user accounts
- Group.ReadWrite.All — permission to add members to groups
- Directory.ReadWrite.All — permission to write directory objects
```powershell
Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All" -NoWelcome
Write-Host "Connected to Graph" -ForegroundColor Green
```
**Step 2 — Created department groups for the tenant**
- Created seven security groups — one per department plus an All Employees base group
- All Employees group serves as the base group every new hire gets added to regardless of department
- Used for company-wide Conditional Access policies and communications
```powershell
New-MgGroup -DisplayName "Engineering Team" -SecurityEnabled:$true ...
New-MgGroup -DisplayName "Sales Team" -SecurityEnabled:$true ...
New-MgGroup -DisplayName "Finance Team" -SecurityEnabled:$true ...
New-MgGroup -DisplayName "IT Team" -SecurityEnabled:$true ...
New-MgGroup -DisplayName "HR Team" -SecurityEnabled:$true ...
New-MgGroup -DisplayName "Management Team" -SecurityEnabled:$true ...
New-MgGroup -DisplayName "All Employees" -SecurityEnabled:$true ...
```

<img width="588" height="168" alt="Screenshot 2026-08-03 at 21 47 48" src="https://github.com/user-attachments/assets/d8f532f3-a0fe-4ad0-9140-8901642b06cd" />

**Step 3 — Created the new hires CSV file**
- CSV acts as the input from HR — simulates how HR would submit new hire requests
- Each row represents one new hire with all properties needed to create their account
- Script reads this file and processes each row automatically
```PowerShell
FirstName,LastName,Department,JobTitle,ManagerUPN,StartDate,LicenseType,Location
John,Smith,Engineering,Software Engineer,...,2026-08-01,E3,Tampa
Maria,Garcia,HR,HR Coordinator,...,2026-08-01,E3,Tampa
```

<img width="892" height="154" alt="Screenshot 2026-08-03 at 21 50 11" src="https://github.com/user-attachments/assets/44175b20-75fa-4bfd-8bc1-c3dcff24a97e" />

**Step 4 — Read the CSV file**     
- Used Import-Csv to read the CSV file
- Each row automatically becomes a PowerShell object with named properties
- Same as building a custom object manually but reading from a file instead
```powershell
$newHires = Import-Csv -Path "$HOME/Desktop/newhires.csv"
Write-Host "Total new hires to process: $($newHires.Count)" -ForegroundColor Green
```
**Step 5 — Built department to group mapping**
- Created a hashtable that maps department names to group names
- Used in the loop to automatically find the right group for each hire
- $departmentGroups["Engineering"] returns "Engineering Team"
- Makes the script dynamic — adding a new department only requires adding one line here
```powershell
$departmentGroups = @{
    "Engineering" = "Engineering Team"
    "Sales"       = "Sales Team"
    "Finance"     = "Finance Team"
    "IT"          = "IT Team"
    "HR"          = "HR Team"
}
```
**Step 6 — Built the main processing loop**
- Looped through every row in the CSV one at a time
- Built the display name by combining first and last name
- Built the UPN using .ToLower() to ensure lowercase format
- Generated a temporary password using the start date with dashes removed using .Replace('-','')
- Example: start date 2026-08-01 becomes password Welcome@20260801!
```powershell
Write-Host "`nStarting onboarding process..." -ForegroundColor Cyan

foreach ($hire in $newHires) {

    Write-Host "`nProcessing: $($hire.FirstName) $($hire.LastName)..." -ForegroundColor Cyan

    try {

        $displayName       = "$($hire.FirstName) $($hire.LastName)"
        $userPrincipalName = "$($hire.FirstName.ToLower()).$($hire.LastName.ToLower())@arbofelipe2003outlook.onmicrosoft.com"
        $mailNickname      = "$($hire.FirstName.ToLower()).$($hire.LastName.ToLower())"
        $tempPassword      = "Welcome@$($hire.StartDate.Replace('-',''))!"
```
**Step 7 — Built password profile and created the account**
- Built password profile as a separate hashtable to avoid syntax issues
- ForceChangePasswordNextSignIn = $true forces user to change password on first login — security best practice
- UsageLocation "US" required before any license can be assigned — must be a two letter country code
- New-MgUser creates the actual account in Entra ID with all properties set
```powershell
$passwordProfile = @{
    Password                      = $tempPassword
    ForceChangePasswordNextSignIn = $true
}
$newUser = New-MgUser -DisplayName $displayName -UserPrincipalName $userPrincipalName -MailNickname $mailNickname -GivenName $hire.FirstName -Surname $hire.LastName -Department $hire.Department -JobTitle $hire.JobTitle -UsageLocation "US" -AccountEnabled:$true -PasswordProfile $passwordProfile
```
**Step 8 — Added user to department group**
- Looked up the correct group name from the hashtable using the hire's department
- Used Get-MgGroup -Filter to find the group by name
- Added the new user to the group using New-MgGroupMember
- If group not found prints a warning and continues — doesn't crash the whole script
```powershell
  if ($group) {
            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $newUser.Id
            Write-Host "  [+] Added to group: $groupName" -ForegroundColor Green
        } else {
            Write-Host "  [!] Group not found: $groupName" -ForegroundColor Yellow
        }
```
**Step 9 — Added user to All Employees group**
- Every new hire gets added to All Employees regardless of department
- Separate from department group — runs after department assignment
- Used for company-wide access policies and communications
```powershell
    $allEmployeesGroup = Get-MgGroup -Filter "DisplayName eq 'All Employees'"
        if ($allEmployeesGroup) {
            New-MgGroupMember -GroupId $allEmployeesGroup.Id -DirectoryObjectId $newUser.Id
            Write-Host "  [+] Added to All Employees group" -ForegroundColor Green
        }
```
**Step 10 — Added result to tracking list**
- Added a custom object to $results for every successfully processed hire
- Includes the temp password in the report so it can be shared with the new hire
- Status set to "Success" for completed accounts
```powershell
        $results.Add([PSCustomObject]@{
            DisplayName       = $displayName
            UserPrincipalName = $userPrincipalName
            Department        = $hire.Department
            JobTitle          = $hire.JobTitle
            GroupAssigned     = $groupName
            TempPassword      = $tempPassword
            StartDate         = $hire.StartDate
            Status            = "Success"
            Error             = ""
        })
```
**Step 11 — Handled errors gracefully with try/catch**
- Wrapped entire account creation in try/catch
- If any step fails for one user the error is caught, printed, and logged
- Script continues to the next hire instead of stopping entirely
- $_.Exception.Message gives the human readable error description
-Failed hires still get added to results with Status "Failed" and the error reason
```powershell

    } catch {
        Write-Host "  [!] Failed: $($_.Exception.Message)" -ForegroundColor Red

        $results.Add([PSCustomObject]@{
            DisplayName       = "$($hire.FirstName) $($hire.LastName)"
            UserPrincipalName = ""
            Department        = $hire.Department
            JobTitle          = $hire.JobTitle
            GroupAssigned     = ""
            TempPassword      = ""
            StartDate         = $hire.StartDate
            Status            = "Failed"
            Error             = $_.Exception.Message
        })
    }
}
```
**Step 12 — Exported onboarding report to CSV**
- Exported full results list to CSV on Desktop
- Report includes every hire — both successful and failed
- Failed hires show the error reason so they can be manually remediated
- Report serves as audit evidence of who was onboarded, when, and what access they received
```powershell
$results | Export-Csv -Path "$HOME/Desktop/Onboarding_Report.csv" -NoTypeInformation
```
**Step 13 — Printed summary**
- Counted successful and failed using Where-Object
- Printed clean summary box at the end
- Green for successful, red for failed
```powershell
$success = ($results | Where-Object Status -eq "Success").Count
$failed  = ($results | Where-Object Status -eq "Failed").Count

Write-Host "  ONBOARDING SUMMARY" -ForegroundColor White
Write-Host "  Total processed : $($results.Count)" -ForegroundColor White
Write-Host "  Successful      : $success" -ForegroundColor Green
Write-Host "  Failed          : $failed" -ForegroundColor Red
```
**PowerShell printed results summary**

<img width="582" height="547" alt="Screenshot 2026-08-03 at 22 05 50" src="https://github.com/user-attachments/assets/0a744da9-f5fc-4914-9f6b-74e290086875" />

**CSV file results**

<img width="1042" height="154" alt="Screenshot 2026-08-03 at 22 06 48" src="https://github.com/user-attachments/assets/f0db7477-442c-4598-8ec3-8e04e9ead71d" />

**Entra Admin Center Results**

<img width="1420" height="570" alt="Screenshot 2026-08-03 at 22 09 45" src="https://github.com/user-attachments/assets/2082c2e8-d008-418f-ba14-4cd084a77fe7" />

<img width="1426" height="428" alt="Screenshot 2026-08-03 at 22 10 08" src="https://github.com/user-attachments/assets/58241ed2-e71b-4c62-81cb-97f0e0ca991b" />

<img width="1419" height="476" alt="Screenshot 2026-08-03 at 22 11 02" src="https://github.com/user-attachments/assets/866348dd-c2d1-45a7-bacb-5a9b6b32c421" />

<img width="1417" height="432" alt="Screenshot 2026-08-03 at 22 10 43" src="https://github.com/user-attachments/assets/ebed387e-8ec8-411c-8a15-a7243eeea231" />

<img width="1419" height="449" alt="Screenshot 2026-08-03 at 22 11 25" src="https://github.com/user-attachments/assets/5dc455b9-3cb9-44b9-afd8-96d93d4e39e2" />

<img width="1424" height="423" alt="Screenshot 2026-08-03 at 22 11 50" src="https://github.com/user-attachments/assets/c6396b3d-eadc-4e33-8509-40b19d5c4992" />
