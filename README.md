# Export-AzureADGroupAssignmentFrequencies_To-HTML.ps1
### SYNOPSIS
Groups typically control permissions and sepperation of duties.

Information produced by this program can help with understanding the state of permission management.

To find out title and department demographics of each Azure AD Group with the help of frequency analysis.

### DESCRIPTION
Analyze Azure Active Directory Groups by frequency of title and department

### EXAMPLE
```PS C:\> Export-AzureADGroupAssignmentFrequencies_To-HTML.ps1```

## Input
Template.html
- Template for HTML layout 

styles.css
- Styles to display HTML layout properly

Users
- Contains list of all users

Groups
- Contains list of all groups and their members

## Output
~\AzureADGroupAssignmentFrequenyAnalysisReports\AzureADGroupAssignmentFrequenyAnalysisReport_$Date
- Groups.html
- Users.html
- JobTitles.html
- Departments.html
- styles.css

# Objects

## #Main-Objects

These objects are queried at runtime and are correlated together to make #sub-objects. All information presented in #Layout or correlated and queried through #Logic is the result of a query ran on one of these #main-objects or #sub-objects.

- #Groups 
	- Contains list of all groups and their members
	- Contains Department Frequencies for each group
	- Contains JobTitle Frequencies for each group

- #Users
	- Contains list of all users and the groups they're a member of

- #JobTitles
	- Contains frequency analysis of JobTitle assignments and relating group frequency analysis for each
		- Group occuring frequency calculated to total users assigned for each JobTitle

- #Departments
	- Contains frequency analysis of Department assignments and relating group frequency analysis for each
		- Group occuring frequency calculated to total users assigned for each Department

## #Sub-Objects

These objects return the content to be displayed in table format.

- #Groups 
	- #GroupAssignmentFrequencies
	- #DuplicateGroupNameFrequencies
	- #AmbiguousGroupNameAssignments
	- #DepartmentFrequenciesInEachGroup**
	- #JobTitleFrequenciesInEachGroup**
	- #GroupsWithoutMembers
	- #GroupsWithoutUsers_ButWithMembers

- #Users 
	- #EnabledUserFrequencies
	- #DirSyncEnabledFrequencies
	- #CountryFrequencies
	- #StateFrequencies
	- #OfficeFrequencies
	- #EnabledUsersWithoutGroupAssignments
	- #DisabledUsersWithGroupAssignments

- #JobTitles
	- #JobTitleFrequencies
	- #JobTitlesWithoutGroupAssignments
	- #GroupAssignmentFrequenciesInEachJobTitle**

- #Departments
	- #DepartmentFrequencies
	- #DepartmentsWithoutGroupAssignments
	- #GroupAssignmentFrequenciesInEachDepartment**


# Stats

### #Totals

- #Groups
	- #TotalGroups
	- #TotalUniqueGroupNames
	- #TotalDuplicateGroupNames
	- #TotalMailEnabledGroups
	- #TotalSecurityEnabledGroups
	- #TotalMailAndSecurityEnabledGroups
	- #TotalGroupsWithADescription
	- #TotalGroupsWithoutADescription
	- #TotalGroupsWithDirSyncEnabled
	- #TotalGroupsWithoutDirSyncEnabled
	- #TotalGroupsWithMembers
	- #TotalGroupsWithoutMembers
	- #TotalGroupsWithUsers
	- #TotalGroupsWithoutUsers
	- #TotalGroupsWithoutUsers_ButWithMembers
	- #TotalGroupsWhereLastDirSyncTimeGT180Days

- #Users
	- #TotalUsers
	- #TotalEnabledUsers
	- #TotalDisabledUsers
	- #TotalUsersWithGroupAssignments
	- #TotalUsersWithoutGroupAssignments
	- #TotalDisabledUsersWithGroupAssignments
	- #TotalEnabledUsersWithoutGroupAssignments
	- #TotalUsersWhereLastDirSyncTimeGT180DaysAndAccountEnabled

- #JobTitles
	- #TotalUniqueJobTitles
	- #TotalJobTitlesWithGroupAssignments
	- #TotalJobTitlesWithoutGroupAssignments

- #Departments
	- #TotalUniqueDepartments
	- #TotalDepartmentsWithGroupAssignments
	- #TotalDepartmentsWithoutGroupAssignments


---

# #Layout

### #Groups - Info-Div_Count: 4
- Title: Groups
- MainHeaderString: $FolderName
- SubTitle: Groups
- SubInfo: 
	- #TotalGroups
	- #TotalUniqueGroupNames
	- #TotalDuplicateGroupNames
	- #TotalMailEnabledGroups
	- #TotalSecurityEnabledGroups
	- #TotalMailAndSecurityEnabledGroups
	- #TotalGroupsWithADescription
	- #TotalGroupsWithoutADescription
	- #TotalGroupsWithDirSyncEnabled
	- #TotalGroupsWithoutDirSyncEnabled
	- #TotalGroupsWithMembers
	- #TotalGroupsWithoutMembers
	- #TotalGroupsWithUsers
	- #TotalGroupsWithoutUsers
	- #TotalGroupsWithoutUsers_ButWithMembers
	- #TotalGroupsWhereLastDirSyncTimeGT180Days
- **Info-Div_1**
	- SubTitle: Group Assignments Overview
	- SubInfo: 
		- #TotalGroups
		- #TotalUniqueGroupNames
		- #TotalDuplicateGroupNames
		- #TotalGroupsWithMembers
		- #TotalGroupsWithoutMembers
		- #TotalGroupsWithUsers
		- #TotalGroupsWithoutUsers
		- #TotalGroupsWithoutUsers_ButWithMembers
	- *Content:*
		- #GroupAssignmentFrequencies
			- Title: GroupAssignmentFrequencies
			- Info: #TotalGroupAssignmentFrequencies
		- #DuplicateGroupNameFrequencies
			- Title: DuplicateGroupNameFrequencies
			- Info: #TotalDuplicateGroupNames
		- #AmbiguousGroupNameAssignments
			- Title: AmbiguousGroupNameAssignments
			- Info: #TotalAmbiguousGroupNameAssignments
- **Info-Div_2**
	- SubTitle: DepartmentFrequenciesInEachGroup
	- SubInfo:
		- #TotalUsers
		- #TotalUsersWithGroupAssignments
		- #TotalUsersWithoutGroupAssignments
		- #TotalDisabledUsersWithGroupAssignments
		- #TotalEnabledUsersWithoutGroupAssignments
		- #TotalGroupsWithUsers
		- #TotalUniqueDepartments
		- #TotalDepartmentsWithGroupAssignments
		- #TotalDepartmentsWithoutGroupAssignments
	- *Content:*
		- #DepartmentFrequenciesInEachGroup**
			- Title: **#of Departments in: $Group**
			- Info: **Dynamic_FrequencyListItem**
- **Info-Div_3**
	- SubTitle: JobTitleFrequenciesInEachGroup
	- SubInfo:
		- #TotalUsers
		- #TotalUsersWithGroupAssignments
		- #TotalUsersWithoutGroupAssignments
		- #TotalDisabledUsersWithGroupAssignments
		- #TotalEnabledUsersWithoutGroupAssignments
		- #TotalGroupsWithUsers
		- #TotalUniqueJobTitles
		- #TotalJobTitlesWithGroupAssignments
		- #TotalJobTitlesWithoutGroupAssignments
	- *Content:*
		- #JobTitleFrequenciesInEachGroup**
			- Title: **#of JobTitles in: $Group**
			- Info: **Dynamic_FrequencyListItem**
- **Info-Div_4**
	- SubTitle: Possible Stale Groups
	- SubInfo:
		- #TotalGroupsWithMembers
		- #TotalGroupsWithoutMembers
		- #TotalGroupsWithUsers
		- #TotalGroupsWithoutUsers
		- #TotalGroupsWithoutUsers_ButWithMembers
	- *Content:*
		- #GroupsWithoutMembers
			- Title: GroupsWithoutMembers
			- Info: #TotalGroupsWithoutMembers
		- #GroupsWithoutUsers_ButWithMembers
			- Title: GroupsWithoutUsers_ButWithMembers
			- Info: #TotalGroupsWithoutUsers_ButWithMembers

### #Users - Info-Div_Count: 2
- Title: Users
- MainHeaderString: $FolderName
- SubTitle: Users
- SubInfo:
	- #TotalUsers
	- #TotalEnabledUsers
	- #TotalDisabledUsers
	- #TotalUsersWithGroupAssignments
	- #TotalUsersWithoutGroupAssignments
	- #TotalDisabledUsersWithGroupAssignments
	- #TotalEnabledUsersWithoutGroupAssignments
	- #TotalUsersWhereLastDirSyncTimeGT180DaysAndAccountEnabled
- **Info-Div_1**
	- SubTitle: General Info
	- SubInfo:
		- #TotalEnabledUsers
		- #TotalDisabledUsers
	- *Content:*
		- #EnabledUserFrequencies
			- Title: EnabledUserFrequencies
			- Info: NULL
		- #DirSyncEnabledFrequencies
			- Title: DirSyncEnabledFrequencies
			- Info: NULL
		- #CountryFrequencies
			- Title: CountryFrequencies
			- Info: TotalCountryFrequencies
		- #StateFrequencies
			- Title: StateFrequencies
			- Info: TotalStateFrequencies
		- #OfficeFrequencies
			- Title: OfficeFrequencies
			- Info: TotalOfficeFrequencies
- **Info-Div_2**
	- SubTitle: Possibly Needs Attention
	- SubInfo:
		- #TotalUsersWithGroupAssignments
		- #TotalUsersWithoutGroupAssignments
		- #TotalDisabledUsersWithGroupAssignments
		- #TotalEnabledUsersWithoutGroupAssignments
	- *Content:*
		- #EnabledUsersWithoutGroupAssignments
			- Title: EnabledUsersWithoutGroupAssignments
			- Info: #TotalEnabledUsersWithoutGroupAssignments
		- #DisabledUsersWithGroupAssignments
			- Title: DisabledUsersWithGroupAssignments
			- Info: #TotalDisabledUsersWithGroupAssignments

### #JobTitles - Info-Div_Count: 2
- Title: JobTitles
- MainHeaderString: $FolderName
- SubTitle: JobTitles
- SubInfo:
	- #TotalUsers
	- #TotalEnabledUsers
	- #TotalDisabledUsers
	- #TotalUniqueJobTitles
	- #TotalJobTitlesWithGroupAssignments
	- #TotalJobTitlesWithoutGroupAssignments
- **Info-Div_1**
	- SubTitle: JobTitles Overview
	- SubInfo:
		- #TotalUniqueJobTitles
		- #TotalJobTitlesWithoutGroupAssignments
	- *Content:*
		- #JobTitleFrequencies
			- Title: JobTitleFrequencies
			- Info: #TotalUniqueJobTitles
		- #JobTitlesWithoutGroupAssignments
			- Title: JobTitlesWithoutGroupAssignments
			- Info: #TotalJobTitlesWithoutGroupAssignments
- **Info-Div_2**
	- SubTitle: #GroupAssignmentFrequenciesInEachJobTitle
	- SubInfo: #TotalJobTitlesWithGroupAssignments
	- *Content:*
		- #GroupAssignmentFrequenciesInEachJobTitle**
			- Title: **#of Groups in: $JobTitle**
			- Info: **Dynamic_FrequencyListItem**

### #Departments - Info-Div_Count: 2
- Title: Departments
- MainHeaderString: $FolderName
- SubTitle: Departments
- SubInfo:
	- #TotalUsers
	- #TotalEnabledUsers
	- #TotalDisabledUsers
	- #TotalUniqueDepartments
	- #TotalDepartmentsWithGroupAssignments
	- #TotalDepartmentsWithoutGroupAssignments
- **Info-Div_1**
	- SubTitle: Departments Overview
	- SubInfo:
		- #TotalUniqueDepartments
		- #TotalDepartmentsWithoutGroupAssignments
	- *Content:*
		- #DepartmentFrequencies
			- Title: DepartmentFrequencies
			- Info: #TotalUniqueDepartments
		- #DepartmentsWithoutGroupAssignments
			- Title: DepartmentsWithoutGroupAssignments
			- Info: #TotalDepartmentsWithoutGroupAssignments
- **Info-Div_2**
	- SubTitle: #GroupAssignmentFrequenciesInEachDepartment
	- SubInfo: #TotalDepartmentsWithGroupAssignments
	- *Content:*
		- #GroupAssignmentFrequenciesInEachDepartment**
			- Title: **#of Groups in: $Department**
			- Info: **Dynamic_FrequencyListItem**

---

# #Logic 

1. Query #Main-Objects
2. Correlate #Main-Objects into #Sub-Objects
3. Query #Sub-Objects
4. Query #Totals
5. Display #Layout

---
