<#
.SYNOPSIS

Script for comparing users in a company list to users in a distribution list.

DESCRIPTION

This PowerShell script take a Company Name (Company) and an optional Department (Dept) name along with the name of a Distribution list (Distro) and compares the active Directory information for the two entries and displays a list of differences if any are found. Thus, allowing users to ensure Distribution lists are properly maintained when new employees are added to the system. If multiple distro checks are done against the same Company/Dept, the previously cached Company/Dept data will be used to increase the efficiency of the process. If a different Company or Dept is used, the User list will be refreshed automatically.

A copy of the Company user list, the users in the Distribution list as well as a list of the differences between the two lists are Automatically saved to the user�s Desktop in a Sub-Folder called Contacts

In Addition, the Get-Departments function allows the user to retrieve a list of departments belonging to the selected company where needed.

USAGE:

Check-Distro [-Company (Required)] [-Distro (Required)] [-Dept] or Get-Departments [-Company (Required)]'

.PARAMETER Company The Company name to pull the list of Employee names and Email addresses from in the Active Directory. This item is Required.

.PARAMETER Distro Name of the Distribution List used to verify the Members of the Company is Contains.

.PARAMETER Dept Filters the list of Employees in the Company to a specific department. If not specified, all users from all departments will be used.

.PARAMETER Company For the Get-Departments function, this required command specifies which Company in the Active Directory to get the list of Departments from.

.EXAMPLE

PS C:> Get-Departments -Company 'Mercury' This command would scan the Active Directory listing for the Company 'Mercury' and return a list of all the departments listed in the AD for Mercury

PS C:>Get-Departments -Company 'a la mode' This command would scan the Active Directory listing for the Company 'a la mode' and return a list of all the departments listed in the AD for a la mode

PS C:> Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept' This command will query the Active Directory listing for all users with the Company name of 'a la mode' and will check if the selected users exist in the Distrobution list 'vsg-dl-alm-supportdept'; If any users from 'a la mode' are missing, they will be displayed on screen.

PS C:> Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept' -Dept 'Customer Support' This command will query the Active Directory listing for all users with the Company name of 'a la mode' and will filter the list to only users listed with a Department of 'Customer Support'; Next it will check if the selected users exist in the Distrobution list 'vsg-dl-alm-supportdept'; If any users from 'a la mode' are missing, they will be displayed on screen.

PS C:> Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept' -Dept 'Customer Support' This function will perform the same process as the example above; however, it will additionally prompt the user to add the missing contacts to the Distrobution list. (Currently this function is disabled)

PS C:> Check-Distro -Company 'a la mode' -Distro 'CVS-DL-OKCCampus' This command will query the Active Directory listing for all users with the Company name of 'a la mode' and will check if the selected users exist in the Distrobution list 'CVS-DL-OKCCampus'; If any users from 'a la mode' are missing, they will be displayed on screen.

PS C:> Check-Distro -Company 'a la mode' -Distro 'CVS-DL-ALM-Everyone' This command will query the Active Directory listing for all users with the Company name of 'a la mode' and will check if the selected users exist in the Distrobution list 'CVS-DL-ALM-Everyone'; If any users from 'a la mode' are missing, they will be displayed on screen.

PS C:> Check-Distro -Company 'TSG' -Distro 'CVS-DL-ALM-Everyone' -Dept 'ALM*' This command will query the Active Directory listing for all users with the Company name of 'TSG' and a department of ALM* (so anything that starts with ALM like 'ALM Testing' will match and check if the selected users exist in the Distrobution list 'CVS-DL-ALM-Everyone'; If any users from 'a la mode' are missing, they will be displayed on screen.

.INPUTS

For Check-Distro : 
Company - This is the name of the Company filter to get the list of users from. (I.E. 'a la mode', 'Mercury' or 'TSG') 
Distro - The name of the Distribution list to compare against (This is a required field). (I.E. 'Distribution-dl-alm-supportdept', 'CVS-DL-OKCCampus' or 'CVS-DL-ALM-Everyone') 
Dept - This Option field applies a secondary filter to the company list narrowing it down from the whole company to a specific department within the company. (I.E. 'Customer Support', 'Mercury Sales' or 'Marketing-Nzd')

For Get-Departments: Company - This required field is used to search the Active Directory listing and returns a list of unique departments for the selected company

.OUTPUTS

For Check-Distro : A list of users missing from the Distro (if any are found) are disabled on the screen in Red A csv list of the Company users (and Department is used) is saved to the user�s desktop in a folder called 'Contacts' A CSV list of the Users from the Distrobution List is saved to the user�s desktop in a folder called 'Contacts' A CSV list of the difference in users, if any, is saved to the user�s desktop in a folder called 'Contacts'

For Get-Departments: A list of the departments is displayed on screen

.NOTES



	This PowerShell script and its functions were written by Jason Krise and utilizes information found online in thefollowing references:
	https://shellgeek.com/PowerShell-get-list-of-users-in-ad-group/
	https://stackoverflow.com/questions/59216952/get-aduser-not-recognized
	https://shellgeek.com/PowerShell-export-active-directory-group-members/
	https://shellgeek.com/set-adgroup-modify-active-directory-group-attributes-in-PowerShell/
	https://dotnet-helpers.com/PowerShell/compare-two-files-list-differences/
	https://stackoverflow.com/questions/30543430/using-PowerShell-get-values-from-sql-table
	https://www.microsoft.com/en-us/download/details.aspx?id=35588
	https://docs.microsoft.com/en-us/PowerShell/module/microsoft.PowerShell.core/about/about_comment_based_help?view=PowerShell-7.1
	https://adamtheautomator.com/PowerShell-comment/
#>
Set-ExecutionPolicy -ExecutionPolicy bypass

Clear-Host <# Clear Screen #>

$global:PreviousCompany = $null #Used to Cache the Company user list since it may be checked against multiple distros; Allows for faster loading of the data
$global:Company = $null #Creates the variable to store the company info in for the initial call too. It is stored as a global since a couple of different functions may use it
$global:rptpath = Get-Location #Loads the current path, used for ensuring the csv files that are written out go to the same location the script is loaded from
$global:arrUsers = @() #Cache of the list of user names pulled from the company list, cache is used to speed up the AD lookup process when multiple distros are checked against the same company list

<# Checking for required ActiveDirectory modules #>
Write-Host 'Checking For ActiveDirectory Module' -BackgroundColor DarkBlue -ForegroundColor White
Write-Host ''

<# If the AD modules are not installed, which by default they are not on a worksatation, enable and install the components#>
if (Get-Module -ListAvailable -Name "ActiveDirectory") { Write-Host 'Active Directory Modules Detected' }
else {
	Write-Host '	Installing Active Directory Modules`n`rThis can take several minutes`n`rPlease Wait. . .'

	Set-ItemProperty "REGISTRY::HKLM\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" UseWUserver -Value 0
	Get-Service wuauserv | Restart-Service
	Get-WindowsCapability -Online -Name RSAT* | Add-WindowsCapability -Online
	Set-ItemProperty "REGISTRY::HKLM\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" UseWUserver -Value 1

	Get-WindowsCapability -Online | Where-Object { $_.Name -like "*ActiveDirectory.DS-LDS*" } | Add-WindowsCapability -Online
	Import-Module activedirectory
	Install-Module -Name ExchangeOnlineManagement
}

Write-Host '	Checking the ActiveDirectory module'

<# Preloading the ActiveDirectory Modules #>
if ((Get-Module -Name "ActiveDirectory")) { '' }
else {
	Import-Module activedirectory
	Write-Host ''
}

Write-Host 'All Set - The required modules are installed and loaded' -BackgroundColor DarkBlue -ForegroundColor White
Write-Host ''

<# ActiveDirectory functionality is loaded and ready for use #>

# Function to check a list of users tied to a company (with a department filter available) against a specified Distro in the current Active Directory
function Check-Distro {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$Company,
		[Parameter(Mandatory = $true)]
		[string]$Distro
	)

	Clear-Host #clear the screen, used more for debugging purposes, but is also useful when there is already a lot of other data in the PowerShell window

	<#  New Multi-company Check Begin    #>
	<# Beging Company check clean up, standradizes the input to removing the trailing semi-colon ";" if it exist #>
	if ($Company.Chars($Company.Length - 1) -eq ';')
	{
		$Company = ($Company.TrimEnd(';'))
	}
	<# End standardization of Company input   #>

	# checks to see if the company list passed to the function was the same as the last one. If so, it will load the previous list of users from the cache instead of reloading from the AD
	# This cache is only stored for as long as the company string remains constant, as long as the current PowerShell instance is loaded, and as long as the PowerShell script (not the function) is not reloaded
	# The Distro is ALWAYS reloaded from the AD and is never cached	
	if ($global:PreviousCompany -eq $Company) {
		<# Write-Host "Contacts from $($Company) were loaded Previously" #>
		Write-Host "The Same Company list was specified, reloading contacts from the Cache" -BackgroundColor DarkBlue -ForegroundColor White

	} else {

		$global:PreviousCompany = $Company #Set Cache reference to the standardized Company string
		[System.Collections.ArrayList]$global:arrUsers = @() #Clears the current cache if set

		#Separate each individual company so we can load the specified users Contacts (Companies are separated by ; whereas the  department filter for a company uses the : designation
		$Check = $Company.Split(";")
		$item = ''

		#For each company found, look for the department filter
		foreach ($item in $Check) {
			$item = $item.Trim() #used to trim leading and trailing white space to account for different ways a user could type the input string
			$Newitem = '' #inistiate the NewItem variable, done here so it is always blank before it is called so old data cannot accidentally get re-used
			$Newitem = $item.Split(":") #Splits the compant and department via the : in the string

			#If the company string contains more than just a company and a department, throw a warning about incorrect formatting and abort (I.E. 'TSG:alamode:Mercury' would be wrong, but 'TSG:alamode; TSG:Mercury' would be acceptable)
			if ($newitem.count -gt 2) {
				Write-Host ""
				Write-Warning "`n`rThe Entry '$($Item)' is invalid, containing to many departments filters per individual Company. `n`rPlease correct your entry using the single company name with optional single department name separated by a colon : (I.E. 'Company name:department') and try again.`n`rTo use multiple departments, please list the company multiple times and separate the companies using a semi-colon ; (I.E. 'Company1:dept1; Company1:dept2; Company2:DeptA; Company2:DeptR) "
				Write-Host ''
				return
			}

			# if a company was specified, but no department filter was included, set the department filter to * (or all) for the search and set a descriptive text to all
			if ($newitem.count -ne 2) { $Newitem += "*" }
			if (($Newitem[1] -eq '*') -or ($Newitem[1] -eq '')) { $DeptFilter = "in 'All' Departments" } else { $DeptFilter = "in the '$($Newitem[1])' Department" }
			$DeptFilter = $DeptFilter -replace "\*","-ALL"
			Write-Host "Loading Users from '$($Newitem[0])' $($DeptFilter):" -BackgroundColor DarkBlue -ForegroundColor White
			[System.Collections.ArrayList]$global:arrUsers += Get-ADUser -Filter "Company -like '$($Newitem[0])' -and Department -like '$($Newitem[1])'" | Get-ADUser -Property DisplayName | Where-Object { $_.Enabled -eq "true" } | Select-Object DisplayName,UserPrincipalName
		}

		Write-Host '' #Give us an blank line
		# $Continue = Read-Host "Press enter to continue" # Used for Debugging
		# return #Used for Debugging
		<# NEW COPMPANY CHECK END #>
	}

	Write-Host "Checking Distro $($Distro) for missing users: `r`n" -BackgroundColor DarkBlue -ForegroundColor White
	Write-Host "Please wait . . ."
	<# Load contact info from Global contacts list where company is listed as referenced in $Company #>

	<# Load list of contcts in the Distro $Distro to comapre against #>
	Clear-Variable arrDi* -Scope Global
	[System.Collections.ArrayList]$arrDistro = Get-ADGroupMember -Identity $Distro -Recursive | Get-ADUser -Property DisplayName | Select-Object DisplayName,UserPrincipalName

	# If The distro is not found or returns as null, Display an error on screen about it
	if ([string]::IsNullOrEmpty($arrDistro)) {
		Write-Host "`n`r`n`r" #Give us an blank line
		Write-Host "The Distribution Name `'$($Distro)`' was not found, `n`rPlease check the spelling of the Distribution list and try again." -ForegroundColor Red
		Write-Host "`n`r`n`r" #Give us an blank line
		return #Returns to prompt and do not process the rest of the function
	}

	<#Run Comparisson #>
	$list = '' #Initiate variable list and set it to null to prevent previous data from causing erroneous results
	$Missing = @() # Initiate The missing variable and set it as an empty array 
	$itemCheck = ''
	foreach ($itemCheck in $arrUsers) {
		if (!($arrDistro -match $itemCheck)) { $Missing += $itemCheck } #If user from company is not found in the Distro,add them to the missing list
	}
	#Format the returned list of missing users to make is easier to read on the screen
	$list = $missing -replace "@{","`n`r"
	$list = $list -replace "}","`n`r"
	$list = $list -replace "DisplayName=","Name: "
	$list = $list -replace "UserPrincipalName=","Email: "


	# is the list of users is empty, then the Company contacts are all in the Distro that was selected.
	if ([string]::IsNullOrEmpty($list)) {
		Write-Host "The $($Company) and $($Distro) lists are in sync" -ForegroundColor Green
	}
	else {
		#Otherwise, display the list of missing users
		Write-Host "The following $($Company) user(s) are missing from the Distro $($Distro):`r`n"
		Write-Host $list -ForegroundColor Red
	}

	<# Save list of users / contacts missing from the Distro #>
	$CompanyFilter = $Company -replace "\*","(ALL)" # cleans up the names to remove invalid file name characters
	$CompanyFilter = $CompanyFilter -replace "\:","-" # cleans up the names to remove invalid file name characters
	$arrUsers | Export-Csv -Path $rptpath"\Get-ADUser $($CompanyFilter).csv" -NoTypeInformation
	$arrDistro | Export-Csv -Path $rptpath"\Get-ADGroupMember $($Distro).csv" -NoTypeInformation
	$Missing | Export-Csv -Path $rptpath"\Users missing from $($Distro).csv" -NoTypeInformation
	Write-Host "A copy of the data gathered was saved in '$($rptpath)' as .csv reports " -ForegroundColor Green
	Write-Host ''
	<# $Continue = Read-Host "Press enter to continue"  # Used for Debugging #>
	''
}

#This function returns the list of departments for the specified company
function Get-Departments {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$Company
	)
	Write-Host "`n`rDepartment Listing for $($Company):" -BackgroundColor DarkBlue -ForegroundColor White
	$Dept = Get-ADUser -Filter "Company -like '$($Company)'" -Property department | Select-Object department | Sort-Object department -Unique
	$Depts = $Dept -replace "@{department=",""
	$Depts = $Depts -replace "}","`n`r"

	# Filter company name provided to use Filename safe characters	
	if (($Company -eq '*') -or ($Company -eq '')) { $CompanyFilter = 'All' } else { $CompanyFilter = $Company }
	$CompanyFilter = $CompanyFilter -replace "\*","-ALL"
	$Dept | Export-Csv -Path $rptpath"\Get-Departments - $($CompanyFilter).csv" -NoTypeInformation

	Write-Host $Depts -ForegroundColor Cyan
	''
}

#After loading functions, display quick help for using the function
[console]::ForegroundColor = "Green"; Get-Help $rptpath"\Check-Distro.ps1" -Detailed;

<# Use Import-Module '.\Check-Distro.ps1' to import the functions into PowerShell #>
<# 
Test Examples 
Get-Departments -Company 'Mercury'
Get-Departments -Company 'a la mode'
Get-Departments -Company 'TSG'
Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept'
Check-Distro -Company 'a la mode' -Distro 'CVS-DL-OKCCampus'
Check-Distro -Company 'a la mode' -Distro 'CVS-DL-ALM-Everyone'
Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept' -Dept 'Customer Support'
Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept' -Dept 'Customer Support' -UpdateAD 'True'
Check-Distro -Company 'Mercury' -Distro 'CVS-DL-ALM-MERC-Everyone'
## https://github.com/DTW-DanWard/PowerShell-Beautifier
## Edit-DTWBeautifyScript "Check-Distro.ps1" -IndentType Tabs
#>
# SIG # Begin signature block
# MIIjPgYJKoZIhvcNAQcCoIIjLzCCIysCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+k4RZ5sx8B9uLCg/Be0n7mBv
# XKGggh3WMIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTEwggQZ
# oAMCAQICEAqhJdbWMht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTE2MDEwNzEyMDAwMFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnF
# OVQoV7YjSsQOB0UzURB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQA
# OPcuHjvuzKb2Mln+X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhis
# EeTwmQNtO4V8CdPuXciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQj
# MF287DxgaqwvB8z98OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+f
# MRTWrdXyZMt7HgXQhBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW
# /5MCAwEAAaOCAc4wggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/
# AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDBQBgNVHSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYc
# aHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggEBAHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafD
# DiBCLK938ysfDCFaKrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6
# HHssIeLWWywUNUMEaLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4
# H9YLFKWA1xJHcLN11ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHK
# eZR+WfyMD+NvtQEmtmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIo
# xhhWz0E0tmZdtnR79VYzIi8iNrJLokqV2PWmjlIwggWQMIIDeKADAgECAhAFmxtX
# no4hMuI5B72nd3VcMA0GCSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0xMzA4MDExMjAwMDBa
# Fw0zODAxMTUxMjAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjQjBAMA8G
# A1UdEwEB/wQFMAMBAf8wDgYDVR0PAQH/BAQDAgGGMB0GA1UdDgQWBBTs1+OC0nFd
# ZEzfLmc/57qYrhwPTzANBgkqhkiG9w0BAQwFAAOCAgEAu2HZfalsvhfEkRvDoaIA
# jeNkaA9Wz3eucPn9mkqZucl4XAwMX+TmFClWCzZJXURj4K2clhhmGyMNPXnpbWvW
# VPjSPMFDQK4dUPVS/JA7u5iZaWvHwaeoaKQn3J35J64whbn2Z006Po9ZOSJTROvI
# XQPK7VB6fWIhCoDIc2bRoAVgX+iltKevqPdtNZx8WorWojiZ83iL9E3SIAveBO6M
# m0eBcg3AFDLvMFkuruBx8lbkapdvklBtlo1oepqyNhR6BvIkuQkRUNcIsbiJeoQj
# YUIp5aPNoiBB19GcZNnqJqGLFNdMGbJQQXE9P01wI4YMStyB0swylIQNCAmXHE/A
# 7msgdDDS4Dk0EIUhFQEI6FUy3nFJ2SgXUE3mvk3RdazQyvtBuEOlqtPDBURPLDab
# 4vriRbgjU2wGb2dVf0a1TD9uKFp5JtKkqGKX0h7i7UqLvBv9R0oN32dmfrJbQdA7
# 5PQ79ARj6e/CVABRoIoqyc54zNXqhwQYs86vSYiv85KZtrPmYQ/ShQDnUBrkG5Wd
# GaG5nLGbsQAe79APT0JsyQq87kP6OnGlyE0mpTX9iV28hWIdMtKgK1TtmlfB2/oQ
# zxm3i0objwG2J5VT6LaJbVu8aNQj6ItRolb58KaAoNYes7wPD1N1KarqE3fk3oyB
# Ia0HEEcRrYc9B9F1vM/zZn4wggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67Z
# MA0GCSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5
# NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8G
# A1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBT
# SEEzODQgMjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDV
# tC9C0CiteLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0
# F9ce2vnS1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lv
# y0daE6ZMswEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrM
# xe6TSXBCMo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vku
# nKoAFdE3/hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNH
# R7OhD26jq22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/r
# JvmM1bL5OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i
# 4F4z8ujo7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyD
# KK05huzUtw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgR
# QRNYmtwmKwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhF
# MJlP/2NPTLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYD
# VR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIw
# HwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGG
# MBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBD
# BgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRUcnVzdGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEE
# ATANBgkqhkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm
# 95RysQDKr2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+
# N3HLIvdaqpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE
# 5BtfQ/g+lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4Imh
# vTnhOE7abrs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KV
# ssIhdXNSy0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwf
# iThV9zeKiwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSc
# hh/jwVYbKyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3x
# GFYHKi8QxAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJs
# QfmCXBVmzGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pK
# HJ6l/aCnHwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52Mb
# OoZWeE4wggdTMIIFO6ADAgECAhAK+FP5bpXTvx4sHxE0Euw+MA0GCSqGSIb3DQEB
# CwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8G
# A1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBT
# SEEzODQgMjAyMSBDQTEwHhcNMjEwODA0MDAwMDAwWhcNMjMwMTAzMjM1OTU5WjCB
# rDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCkNhbGlmb3JuaWExDzANBgNVBAcTBkly
# dmluZTEYMBYGA1UEChMPQ29yZUxvZ2ljLCBJbmMuMRQwEgYDVQQLEwtUcmVzdGxl
# T0lEQzEYMBYGA1UEAxMPQ29yZUxvZ2ljLCBJbmMuMS0wKwYJKoZIhvcNAQkBFh50
# cmVzdGxlYWRtaW4ucmVzQGNvcmVsb2dpYy5jb20wggGiMA0GCSqGSIb3DQEBAQUA
# A4IBjwAwggGKAoIBgQDDPJ5rkYnfhulWJR8nqiWqkWLPSg/wj+OE/wt50PMrFNeZ
# dZtr8nr/K9PjJpv63e2y14T5x7R6hJqUOq/pFrOUe4LSPhYRhHCw1oai2lO/NfxA
# oQs7JDc5hE/bRE37PPPbTOwoVm18z2AxkXD+4lADgGuhgVF+nbQWPfS6zEXSz3/I
# HrQunZkxcQ/o2Ygib1SSGB43ktWpIe1WAaSWt5wOO3YhyUo/0sREd72NIdX0hEVF
# bsUKrdb1Yh5ETUHu7xhtp3g4CEh6zDR93wNn/Is+x+UPlXmedCsF+k1L0rvbWzED
# 9LHGKWd7CH4VSqpruDtoheFPlYgIvK5Ua8nmBcW3RFG+SpVVy8uP+45X+mm6wFCR
# dYX5E6VvQM1KrtY8l00JKQk909D4FwFio+sSDEvEh12w0YWMLyaUSRMTESZa61zT
# xEaZEvj2/wW0ii3UI9hO0Glzk848O97Cvc/8Cs1YE8xPFioB3oXuSTAApftXSogp
# BtKgMDKVAqUWWd+S5W0CAwEAAaOCAjEwggItMB8GA1UdIwQYMBaAFGg34Ou2O/hf
# EYb7/mF7CIhl9E5CMB0GA1UdDgQWBBS6O0aWUks/ZpumgY2MZyozGqmm8DApBgNV
# HREEIjAggR50cmVzdGxlYWRtaW4ucmVzQGNvcmVsb2dpYy5jb20wDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMIG1BgNVHR8Ega0wgaowU6BRoE+G
# TWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENvZGVT
# aWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3JsMFOgUaBPhk1odHRwOi8vY3Js
# NC5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRDb2RlU2lnbmluZ1JTQTQw
# OTZTSEEzODQyMDIxQ0ExLmNybDA+BgNVHSAENzA1MDMGBmeBDAEEATApMCcGCCsG
# AQUFBwIBFhtodHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwgZQGCCsGAQUFBwEB
# BIGHMIGEMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXAYI
# KwYBBQUHMAKGUGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRy
# dXN0ZWRHNENvZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAGaThQMF/SbszB9sXbcHc0RsYRUk
# XwMmsxqdcsrEnh889JSQROEgQCVCXjoj0g1zgcmroh1A05ukd5CpaSq5ySu/830c
# NeFQxMssa8RW6i9Mm1cL7sq0EhPa6HlqLzaZg8+2tqjx4dlPKVkgOGevUgMERhDe
# K9PquxjRsskf4f9TY2x+qssCRTm5ayP85ualfKj7Anw2SeacF+9UQOWOCmVtfqcU
# rVRaTQOLnANddzmIF0MWSgC8xhEgmgMU5tyzGLZHmstZiaZMu9MdebhmBlislgFp
# Ze2oNQOGj3LNLn+5P3HDdexjxIf2KW47CNxWrVSmUFF7Pesq3bq+ugcIWqP8dQ5A
# XhN6/RERTXT12EIUBj14Jak4oru8cL98ZfX98UY7uazikgdAFithWd1yu6769sbJ
# SWA8KWONWPPX2gF+u/U3z45hvLztgzaEFegeK0vPJr5TaPr/GYCnqwO+pex7Ajpq
# TNryWNVFjzxkwY/cCo3tKAIvsnaePfVeNt1EALN/i7diJ89TlnsQFCd+l0NIS9Mj
# usDCRgJ6RvrU6bPiktWEkQHWocBOs8wPOaaUM6jY6ofBKgqgt1cHiBePn3VvryCf
# BjzV/64Jmu4iDnKB3N68tNLWsuoxVH3NoWsbQ3HTboSDuL6bsWTkUYEuD/Ls7dcH
# Kk9iKGgOBQC3BGWjMYIE0jCCBM4CAQEwfTBpMQswCQYDVQQGEwJVUzEXMBUGA1UE
# ChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQg
# Q29kZSBTaWduaW5nIFJTQTQwOTYgU0hBMzg0IDIwMjEgQ0ExAhAK+FP5bpXTvx4s
# HxE0Euw+MAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkG
# CSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEE
# AYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSTuuEDKJk6FZiRv0HqrhYROn3YqDANBgkq
# hkiG9w0BAQEFAASCAYAn5yRaWQXXfE8y9vFB9azk2P2Df3ImF8BLbCaVzJKTB8CZ
# jYQdln60oS/UiSzLfv9IaMaS7D0f1qR+8dIa14pwEWggE9GD1cUwHRwJFFc+l8kV
# mC3hsZCpgeQrL+2KbYBsPUAaSyY+XSjCB2lXU5A5fOmDj348Dc+xpKojRKgS+Nhl
# ueuLyQfZ8kXsREg+m7TYjmwwQRbe5oYrn9/bwfqkfj6FnmjWZWHY4GvcAmZeNvnT
# LvJzOVJ+4z8A5rZ+fhOkm+5LPzL3u75ZzFLPw9PpsZqEJdhd4vuqkP1LomMgy+JX
# 3Vuh4WFOoV17adyW/P9jpsq8stsY0zED/NPs6AZAJE/x1L480mb886VQ9GxMBsXF
# /DwpTlSMGziY1Vh2kOqsCrhwaDXePsU1Qz495bJ3uhr5k5vU1oyy8gpU97oZqVaV
# I5SOWfvmVD7q0gNBbX7OnhDxXsXUujcQndJFPWfsXN7jUF02YkJOOz9MEOfUZ0fX
# KxyorpjuNKh9A/xM7vyhggIwMIICLAYJKoZIhvcNAQkGMYICHTCCAhkCAQEwgYYw
# cjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVk
# IElEIFRpbWVzdGFtcGluZyBDQQIQDUJK4L46iP9gQCHOFADw3TANBglghkgBZQME
# AgEFAKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
# DTIxMTExMjE1MDc1N1owLwYJKoZIhvcNAQkEMSIEIJmZAIubCqXVRagg5y1I4kIe
# 6J7q+Pas/Ji0uQvRZ66oMA0GCSqGSIb3DQEBAQUABIIBAJ/U4ah3WZcd0z2gOS1K
# /5pYO7xxJ/2EU35NdWnbdW/yY2wNxDfZEFBBd80QvPaQLC50Xv9gxHRpHrjf3LKq
# Ydvo/k/1EICH8gFbYtvu3m3sjcc+z6vv2ToP1DLyky8BHaZ6qUOg9zz8ONdsraY7
# pQNru7Ur3NBlg81dvHfocC+qPtD7q0mkcQ629C0v3X2RMvUgnIBp30glZgIPsnjn
# 0Hd/P924Lo0lTwGFmfDo/S525S27ETANQd14SjrnwfN4GdhgtQHt46MaPskJUSQH
# TKxM22ECbtUezhPeY/lnlmmjkXrZRxJw5dW9znRvYdOSpmjXAlVM3mnaZcZMUXHj
# l6c=
# SIG # End signature block
