CompanyCompanyCompanycomparecomparecomparecomparecomparecompareCompanyCompanyCompanyCompanyCompanyCompanyCompanyCompany.SYNOPSIS

Script for comparing users in a company list to users in a distribution list.

DESCRIPTION

This PowerShell script take a Company Name (Company) and an optional Department (Dept) name along with the name of a Distribution list (Distro) and compares the active Directory information for the two entries and displays a list of differences if any are found. Thus, allowing users to ensure Distribution lists are properly maintained when new employees are added to the system. If multiple distro checks are done against the same Company/Dept, the previously cached Company/Dept data will be used to increase the efficiency of the process. If a different Company or Dept is used, the User list will be refreshed automatically.

A copy of the Company user list, the users in the Distribution list as well as a list of the differences between the two lists are Automatically saved to the user’s Desktop in a Sub-Folder called Contacts

In Addition, the Get-Departments function allows the user to retrieve a list of departments belonging to the selected company where needed.

USAGE:

Check-Distro [[-Company (Required)] <String>] [-Distro (Required)] <String> [[-Dept] <String>] [[-UpdateAD] <String>]
or
Get-Departments [[-Company (Required)] <String>]'

.PARAMETER Company
The Company name to pull the list of Employee names and Email addresses from in the Active Directory. This item is Required.

.PARAMETER Distro
Name of the Distribution List used to verify the Members of the Company is Contains.

.PARAMETER Dept
Filters the list of Employees in the Company to a specific department. If not specified, all users from all departments will be used.

.PARAMETER UpdateAD
If specified, this option prompts the users to add the missing contacts to the Distribution list. (This function is deprecated)

.PARAMETER Company
For the Get-Departments function, this required command specifies which Company in the Active Directory to get the list of Departments from.

.EXAMPLE

PS C:\> Get-Departments -Company 'Mercury'
This command would scan the Active Directory listing for the Company 'Mercury' and return a list of all the departments listed in the AD for Mercury

PS C:\>Get-Departments -Company 'a la mode'
This command would scan the Active Directory listing for the Company 'a la mode' and return a list of all the departments listed in the AD for a la mode

PS C:\> Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept'
This command will query the Active Directory listing for all users with the Company name of 'a la mode' and will check if the selected users exist in the Distrobution list 'vsg-dl-alm-supportdept'; If any users from 'a la mode' are missing, they will be displayed on screen.

PS C:\> Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept' -Dept 'Customer Support'
This command will query the Active Directory listing for all users with the Company name of 'a la mode' and will filter the list to only users listed with a Department of 'Customer Support'; Next it will check if the selected users exist in the Distrobution list 'vsg-dl-alm-supportdept'; If any users from 'a la mode' are missing, they will be displayed on screen.

PS C:\> Check-Distro -Company 'a la mode' -Distro 'vsg-dl-alm-supportdept' -Dept 'Customer Support' -UpdateAD 'True'
This function will perform the same process as the example above; however, it will additionally prompt the user to add the missing contacts to the Distrobution list. (Currently this function is disabled)


PS C:\> Check-Distro -Company 'a la mode' -Distro 'CVS-DL-OKCCampus'
This command will query the Active Directory listing for all users with the Company name of 'a la mode' and will check if the selected users exist in the Distrobution list 'CVS-DL-OKCCampus'; If any users from 'a la mode' are missing, they will be displayed on screen.

PS C:\> Check-Distro -Company 'a la mode' -Distro 'CVS-DL-ALM-Everyone'
This command will query the Active Directory listing for all users with the Company name of 'a la mode' and will check if the selected users exist in the Distrobution list 'CVS-DL-ALM-Everyone'; If any users from 'a la mode' are missing, they will be displayed on screen.

PS C:\> Check-Distro -Company 'TSG' -Distro 'CVS-DL-ALM-Everyone' -Dept 'ALM*'
This command will query the Active Directory listing for all users with the Company name of 'TSG' and a department of ALM* (so anything that starts with ALM like 'ALM Testing' will match and check if the selected users exist in the Distrobution list 'CVS-DL-ALM-Everyone'; If any users from 'a la mode' are missing, they will be displayed on screen.


.INPUTS

For Check-Distro :
	Company - This is the name of the Company filter to get the list of users from. (I.E. 'a la mode', 'Mercury' or 'TSG')
	Distro - The name of the Distribution list to compare against (This is a required field). (I.E. 'Distribution-dl-alm-supportdept',  'CVS-DL-OKCCampus' or  'CVS-DL-ALM-Everyone')
	Dept - This Option field applies a secondary filter to the company list narrowing it down from the whole company to a specific department within the company. (I.E. 'Customer Support', 'Mercury Sales' or 'Marketing-Nzd')
	UpdateAD - This optional feature enables the prompt to update the Distrobution List with any missing users. Presently while the prompt can be enabled, the function itself is currently not available.

For Get-Departments:
	Company - This required field is used to search the Active Directory listing and returns a list of unique departments for the selected company

.OUTPUTS

For Check-Distro :
	A list of users missing from the Distro (if any are found) are disabled on the screen in Red
	A csv list of the Company users (and Department is used) is saved to the user’s desktop in a folder called 'Contacts'
	A CSV list of the Users from the Distrobution List is saved to the user’s desktop in a folder called 'Contacts'
	A CSV list of the difference in users, if any, is saved to the user’s desktop in a folder called 'Contacts'
	
For Get-Departments:
	A list of the departments is displayed on screen
	
.NOTES

	This PowerShell script and its functions were written by Jason Krise and utilizes information found online in the following references:
	https://shellgeek.com/powershell-get-list-of-users-in-ad-group/
	https://stackoverflow.com/questions/59216952/get-aduser-not-recognized
	https://shellgeek.com/powershell-export-active-directory-group-members/
	https://shellgeek.com/set-adgroup-modify-active-directory-group-attributes-in-powershell/
	https://dotnet-helpers.com/powershell/compare-two-files-list-differences/
	https://stackoverflow.com/questions/30543430/using-powershell-get-values-from-sql-table
	https://www.microsoft.com/en-us/download/details.aspx?id=35588
	https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-7.1
	https://adamtheautomator.com/powershell-comment/
